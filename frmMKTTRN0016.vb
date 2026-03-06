Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.IO

Friend Class frmMKTTRN0016
	Inherits System.Windows.Forms.Form
	'===================================================================================
	' (c) MothersonSumi INfotech & Design Ltd. All rights reserverd.
	' File Name         :   FRMMKTTRN0009.frm
	' Function          :   Used to add sale deatails
	' Created By        :   Nisha & Kapil
	' Created On        :   15 May, 2001
	' Revision History  :   Nisha Rai
	'21/09/2001 MARKED CHECKED BY BCs changed on version 3
	'03/10/2001 MARKED CHECKED BY BCs  for jobwork invoice changed on version 7
	'09/10/2001  changed on version 8 for schedule Status
	'09/01/2002 changed fof Smiel Chennei to add CVD_PER,SVD_Per,Insurance
	'25/01/2002 changed for decimal 4 places on Chacked Out Form No = 4019
	'28/01/2002 changed for decimal 4 places on Chacked Out Form No = 4033
	'in ChangeCellTypeStaticText()
	'15/01/2002 CHANGED FOR DOCUMENT NO. ON FORM NO. 4066
	'19/04/2002 changed on  for Tariff & TDA Changes
	'22/04/2002 changed for box quantity from ITem_Mst
	'23/04/2002 BOM structure changed
	'30/04/2002 to change on the basis of currency then decimal places
	'30/04/2002 for replacing Mod Function
	'08/05/2002 SCRAP invoice Changes
	'29/05/2002 schedule check
	'03/06/2002 for changes in refresh form to set list index to -1
	'04/06/2002 for enabling all text feilds in Rejection invoice
	'11/06/2002 for from box size changes in Quantity Check variable type int to double
	'13/06/2002 CHANGE IN BOMCHECK FUNCTION
	'17/06/2002 change label in Grim From Drawing No to Cust Part No & to Show Packing
	'in Percantage
	'21/06/2002 Changes for Machino & Change in Quantity Check in case of Item BAl Check
	'25/06/2002 Changed For SMIEL .a>)Consideration of ACTIVE FLAG in DTL table Instead of
	'                               HDR Table.  -   Nitin Sood
	'                              b>)Item Code and Drwg No. have a Relation of 1 to Many. - Nitin Sood
	'                              c>)Amendment Number Field Added on Form.
	'27/06/2002 Chenged for SMIEL  a>)DatePicker Added to Form,Date of New Challan Can be From Last Made
	'Challan (Build Flag) till Current Date
	'                                   -   Nitin Sood
	'01/07/2002 Change for More then one Selected RGPs Nos
	'05/07/02 changed for rounding off data
	'CHANGED ON 15/07/2002 FOR EXPORT OPTION ADDING AND CALCULATION SAME AS NORMAL INVOICE
	'23/07/2002 changed to add Grin Linking in Rejection Invoice
	'07/08/2002 changed for Jobwork invoice to check Customer supplied from Vendor Bom
	'Changed by nisha on 26/08/2002
	'changes done by nisha to check SO Qty in Challan Entry
	'Per Value changes on 22/10/2002 by nisha 04/12/2002
	'CHANGES DONE BY NISHA ON 7/03/2003 FOR ADDING TOOL COST IN CASE OF SAMPLE INVOICE
	'CHANGES DONE BY NISHA ON 13/03/2003
	'1.FOR FINAL MERGING & FOR FROM BOX & TO bOX UPDATION WHILE EDITING INVOICE
	'2.For Grin Cancellation flag
	'3.SAMPLE INVOICE TOOL COST COLUMN
	'4.CUNSUMABLES & MISC. SALE IN CASE OF NORMAL RAW MATERIAL INVOICE
	'CHANGES DONE BY NISHA ON 03/04/2003 - 14/04/2003 -16/04/2003
	'CHANGES DONE BY NISHA 16/05/2003 TO ALLOW 11 ITEMS.
	'changes DOnE BY NISHA  to add SRVDI & SRV LOCATION on 03/06/2003 05/06/200312/06/2003
	'changes done by nisha on 26/06/2003 02/07/2003
	'changes done by nisha on 09/10/2003 TO Disable Rate and Others
	'changes done by nisha on 07/11/2003
	'1 update TotalExcise Value in Sales_dtl
	'Changes Added by nisha on 16/02/2004
	'1.to add tool Amortisation
	'Changes Added by nisha on 20/02/2004
	'1.to add eNagare System
	'---------------
	'01/07/2004 By Arshad
	'Convert function is being used in query for ref no help with format yyyy/mm/dd ie. 111
	'---------------
	'08/07/2004 Changes Done By Nisha
	'For ECESS Calculations
	'Changes by Rajani Kant in tool Amortisation for BOM on 19/08/2004
	'Changes done By Sourabh on 03 Sep 2004 For Handling Schedule using FIFO Method(DSTracking-10623)
	'Changes Done by Nisha for Tool Amortization to Link With Tool Mst on 06/10/2004-2
	'Changes Done by Nisha for removing the check of Finished item in amor_dtl on 20/10/2004
	'Changes Done by Nisha for removing the check of tool code in Finished Items in amor_dtl on 24/11/2004
	'Changes Done by Nisha for total Roundoff proble Debit Credit Does Not Match on 07 Apr 05
	'Changed by Ravjeet. Round off of Difference Value to be done on the basis of decimals of the Invoice Value, this is to prevent DEBIT / CREDIT Mismatch message 27 Apr 2005
	'===================================================================================
	' Changed by Sandeep Chadha On 03-May-2005
	' Description : Add New Tax State Development Tax.
	'===================================================================================
	'Changed by Nisha Round off Replaced by Mid, this is to prevent DEBIT / CREDIT Mismatch message 20 May 2005
	'Changed by Nisha Round off Replaced by Mid, this is to prevent DEBIT / CREDIT Mismatch message 14 july 2005
	'==================================================================================================
	' Changes done by ashutosh on 24-08-2005 , issue id- 14999
	' save Cust_ref and amendment number to Sales_dtl table.
	'==================================================================================================
	' Changes done by Nisha on 26-10-2005
	' by default EC2 should come in ECESS
	'==================================================================================================
	'Revised By   : Ashutosh ,issue Id:16205
	'History      : On 14-11-2005, Changes done for check the schedule and dispatches one extra working day of invoice date.
	'==================================================================================================
	'Revised By   : Ashutosh ,issue Id:17355
	'History      : On 05-04-2006 , Tarriff code validation.
	'==================================================================================================
	'Revision  By       : Ashutosh , Issue Id :17610
	'Revision On        : 26-04-2006
	'History            : Save Bin Quantity in invoice.
	'                   : Save Stock Location in Invoice.
	'                   : Validate UOM for sales Quantity & Bin Quantity  from measurement master.
	'                   : Validate Currency  code from sales order while saving the invoice.
	'                   : Refresh from box & to box entries while cahnging bin qty & Invoice Qty.
	'                   : Calculate accessible value on MRP , in case of M type SO.
	'                   : Check for cust supplied material in SO before saving this in invoice.
	'                   : Check for wrong stock updation.
	'-----------------------------------------------------------------------------------
	'Revised By      : Davinder Singh
	'Issue ID        : 19575
	'Revision Date   : 27 Feb 2007
	'History         : New Tax (SEcess) added
	'-----------------------------------------------------------------------------------
	'Revised By      : Manoj Kr. Vaish
	'Issue ID        : 19992
	'Revision Date   : 28 June 2007
	'History         : Display Credit Term from Cust_Ord_Dtl and save into saleschallan_dtl
	'                  During Invoice Posting, fetch credit term from saleschallan_dtl for saving in ar_docmaster
	'***************************************************************************************
	'Revised By      : Manoj Kr. Vaish
	'Issue ID        : 21551
	'Revision Date   : 20-Nov-2007
	'History         : Add New Tax VAT with Sale Tax help
    '***********************************************************************************
    'Revised By      : Manoj Kr. Vaish
    'Issue ID        : eMpro-20080528-19322
    'Revision Date   : 28 May 2008
    'History         : If DS Tracking is Allowed then Schdule should not be reversed while deleting
    '                  the invoice.It causes the despatch quantity in negative value.
    '***********************************************************************************
    'Revised By      : Manoj Kr. Vaish
    'Issue ID        : eMpro-20090223-27780
    'Revision Date   : 25 Feb 2009
    'History         : Navigation problem while making the Invoice(focus was going on Quantity)
    '***********************************************************************************
    'Revised By      : Manoj Kr. Vaish
    'Issue ID        : eMpro-20090529-31889
    'Revision Date   : 29 May 2009
    'History         : Wrong calculation of Tool Amortization and
    '                : Total invoice round off value is not saving correctly in Export Invoice.
    '                : Add new Tax Additonal VAT
    '***********************************************************************************
    'Revised By      : Manoj Vaish
    'Revision On     : 01 Jun 2009
    'Issue ID        : eMpro-20090610-32326
    'History         : Addition of new additional CST tax
    '***********************************************************************************
    'Revised By      : Manoj Vaish
    'Revision On     : 21 Jul 2009
    'Issue ID        : eMpro-20090720-33879
    'History         : Show the additonal VAT from the selected Sales Order.
    '***********************************************************************************
    'Revised By      : SIDDHARTH RANJAN
    'Revision On     : 11 NOV 2009
    'Issue ID        : eMpro-20091113-38843
    'History         : ADD NEW INVOICE TYPE & SUB INVOICE TYPE ("CSM INVOICE")
    '**************************************************************************************
    ' Revised By                 -   Roshan Singh
    ' Revision Date              -   10 JUN 2011
    ' Description                -   FOR MULTIUNIT FUNCTIONALITY
    '***********************************************************************************
    'Revised by Vinod Singh on 18 Sep 2012  for Multi Unit Migration - SMIEL
    'Issue Id : 10277476
    '*******************************************************************************************
    ' Revised By                 -   Saurav Kumar
    ' Revision Date              -   14 Dec 2012
    ' Description                -   error for Raw material Invoicing, less parameters in INSERT
    ' Issue id                   -   10319722
    '*******************************************************************************************
    ' Revised By                 -   Prashant Rajpal
    ' Revision Date              -   07 june 2013
    ' Description                -   Sales Voucher wrongly generated in USD -exchnage rate not reset 
    ' Issue id                   -   10398657 
    '*******************************************************************************************
    ' Revised By                 -   Prashant Rajpal
    ' Revision Date              -   03 -sep 2013
    ' Description                -   Sales Voucher wrongly generated in USD -exchnage rate not reset 
    ' Issue id                   -   10354980 
    '*******************************************************************************************
    'REVISED BY      : PRASHANT RAJPAL
    'ISSUE ID        : 10465802
    'DESCRIPTION     : CHANGES FOR SMIIEL : TRANSFER INVOICE FOR MSSL NOW INSTEAD OF NORMAL INVOICE 
    'REVISION DATE   : 18 Dec 2013 - 19 Dec 2013
    '***********************************************************************************************
    'REVISED BY      : PRASHANT RAJPAL
    'ISSUE ID        : 10569249 
    'DESCRIPTION     : CHANGES FOR SMIIEL : BY DEFAULT TRANSFER INVOICE WITH FINISHED GOODS APPEARED FOR SMIEL UNIT
    'REVISION DATE   : 08 APR 2014'
    '***********************************************************************************************
    'REVISED BY     :  PRASHANT RAJPAL
    'REVISED DATE   :  13-JAN-2015
    'ISSUE ID       :  10736222
    'PURPOSE        :  TO INTEGRATE CT2 AR3 FUNCTIONALITY 

    'REVISED BY     -  ASHISH SHARMA    
    'REVISED ON     -  06 JUL 2017
    'PURPOSE        -  101188073 — GST CHANGES
    '************************************************************************************************
    'REVISED BY     -  ASHISH SHARMA    
    'REVISED ON     -  05 OCT 2018
    'PURPOSE        -  101631219 - SMIIEL Auto Invoicing Functionality
    '************************************************************************************************

    Dim mintIndex As Short 'Declared To Hold The Form Count
	Dim mdblPrevQty() As Object 'to store prev quantity in edit mode
	Dim mdblToolCost() As Object 'to insert tool cost item wise
	Public mstrItemCode As String 'To Get The Value Of Item Code
	Dim mstrInvoiceType As String 'To Get The Value Of Invoice Type
	Dim mstrInvoiceSubType As String 'To Get The Value Of Invoice Sub Type
	Dim mstrAmmendmentNo As String 'To Get The Value Of Ammendment No.
	Dim mstrInvType As String 'To Get Value Of Inv Type From SalesChallan_Dtl
	Dim mstrInvSubType As String 'To Get Value Of Inv SubType From SalesChallan_Dtl
	Dim mstrUpdDispatchSql As String 'To Make Update Query For Dispatch_Qty From Daily/Monthly Mkt Schedule
	Dim mstrAmmNo As String
	Dim mExchageRate As Double
	Dim mCurrencyCode As String
	Dim strupdateamordtlbom As String
	Dim mstrRefNo As String
	Dim strupSalechallan As String
	Dim strupSaleDtl As String
	Dim strInvType As String
	Dim strInvSubType As String
	Dim strBomItem As String 'For Latest Item To Explore
	Dim strBomMstRaw As String
	Dim blnFIFO As Boolean
	Dim blnEOU_FLAG As Boolean
	Dim inti As Short
	Dim arrItem() As String
	Dim arrQty() As Double
	Dim arrReqQty() As Double
	Dim mstrRGP As String
    Dim mstrLocation As String
    Dim blnmsg As Boolean = False
    Dim objRpt As ReportDocument
    Dim frmReportViewer As New eMProCrystalReportViewer
    '101188073
    Private Const IS_HSN_SAC As Byte = 23
    Private Const HSN_SAC_CODE As Byte = 24
    Private Const CGST_TYPE As Byte = 25
    Private Const SGST_TYPE As Byte = 26
    Private Const IGST_TYPE As Byte = 27
    Private Const UTGST_TYPE As Byte = 28
    Private Const COMP_CESS_TYPE As Byte = 29
    '101188073
    Private Enum enumExciseType
        RETURN_EXCISE = 1
        RETURN_CVD = 2
        RETURN_SAD = 3
        RETURN_ALLExcise = 4
    End Enum
    Dim objInvoicePrint As New prj_InvoicePrinting.clsInvoicePrinting(gstrDateFormat)
    Dim intNoCopies As Short
	Dim strTensArray(15) As String 'Used in converting number into words
    Dim strHundredsArrey(5) As String 'Used in converting number into words
	Dim strToWords As String 'Used in converting number into words
	Dim mIntDecimalPlace As Short
	Dim updatestockflag, updatePOflag As Boolean
	Dim saleschallan As String
	Dim strStockLocation As String
	Dim mAmortization As Double
	Dim mStrCustMst As String
	Dim mblnEOUUnit As Boolean
	Dim mAssessableValue As Double
	Dim mOpeeningBalance As Double
	Dim strsaledetails As String
	Dim strupdateGrinhdr As String
	Dim strupdateitbalmst As String
    Dim strSelectItmbalmst As String
    Dim strupdatecustodtdtl As String
	Dim strUpdateAmorDtl As String
	Dim salesconf As String
	Dim msubTotal, mInvNo, mExDuty, mBasicAmt, mOtherAmt As Double
	Dim mFrAmt, mGrTotal, mStAmt, mCustmtrl As Double
	Dim mDoc_No As Short
	Dim mAccount_Code, mInvType, mSubCat, mlocation As String
	Dim mstrAnnex As String
	Dim mCust_Ref, mAmendment_No As String
    Dim arrCustAnnex() As Object
	Dim ref57f4 As String 'used in BomCheck() insertupdateAnnex()
	Dim dblFinishedQty As Double 'To get Qty of Finished Item from Spread
	Dim strCustCode As String 'used in BomCheck() insertupdateAnnex()
	Dim strItemCode As String 'used in BomCheck() insertupdateAnnex()
    Dim blnFIFOFlag As Boolean
	Dim rsBomMst As ClsResultSetDB
	Dim mstrMasterString As String 'To store master string for passing to Dr Cr COM
	Dim mstrDetailString As String 'To store detail string for passing to Dr Cr COM
	Dim mstrPurposeCode As String 'To store the Purpose Code which will be used for the fetching of GL and SL
	Dim mblnAddCustomerMaterial As Boolean 'To decide whether to add customer material in basic or not
	Dim mblnSameSeries As Boolean 'To store the flag whether the selected invoice will have same series as others
	Dim mstrReportFilename As String 'To store the report filename
	Dim mblnInsuranceFlag As Boolean 'To store insurance flag
	Dim mblnpostinfin As Boolean
	Dim mIntNoCopies As Short
	Dim mblnExciseRoundOFFFlag As Boolean
	Dim mSaleConfNo As Double
	Dim mstrExcisePriorityUpdationString As String
    Dim blnTotalInvoiceAmount As Boolean
	Dim intTotalInvoiceAmountRoundOffDecimal As Short
	Dim ldblTotalInvoiceValueRoundOff As Double
    Dim mstrCreditTermId As String

	Private Sub chkDTRemoval_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkDTRemoval.CheckStateChanged
		If chkDTRemoval.CheckState = System.Windows.Forms.CheckState.Checked Then
			dtpRemoval.Enabled = True
			dtpRemovalTime.Enabled = True
		Else
			dtpRemoval.Enabled = False
			dtpRemovalTime.Enabled = False
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
        On Error GoTo ErrHandler
		Select Case KeyAscii
			Case System.Windows.Forms.Keys.Return
                If dtpDateDesc.Enabled Then dtpDateDesc.Focus()
		End Select
		GoTo EventExitSub
ErrHandler: 'The Error Handling Code Starts here
		Call gobjError.RaiseError(Err.Number, err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
		GoTo EventExitSub
EventExitSub: 
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	Private Sub CmbInvSubType_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmbInvSubType.Leave
        On Error GoTo ErrHandler
        Dim strsql As String
        Dim blntcscheck As Boolean = False

		If UCase(CmbInvType.Text) = "NORMAL INVOICE" Then
			If UCase(CmbInvSubType.Text) = "SCRAP" Then
				txtRefNo.Enabled = False : txtRefNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdempHelpRefNo.Enabled = False : txtRefNo.Text = ""
				txtAmendNo.Enabled = False : txtAmendNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : txtAmendNo.Text = ""
                ctlPerValue.Enabled = True
                ctlPerValue.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
			Else
                ctlPerValue.Enabled = False
                ctlPerValue.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                txtRefNo.Enabled = True : txtRefNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : cmdempHelpRefNo.Enabled = True
                txtAmendNo.Enabled = True : txtAmendNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            End If
        End If
        If Len(Trim(txtCustCode.Text)) > 0 Then
            strSQl = "select dbo.UDF_IRN_TCSREQUIRED( '" & gstrUNITID & "','" & txtCustCode.Text.Trim & "')"
            If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSql)) = True Then
                blntcscheck = True
            Else
                blntcscheck = False
            End If
            If blntcscheck = True Then
                Call checktcsvalue(CmbInvType.Text, CmbInvSubType.Text)
            Else
                If (UCase(Trim(CmbInvType.Text) = "NORMAL INVOICE") And (UCase(Trim(CmbInvSubType.Text)) = "SCRAP")) Then
                    If gblnGSTUnit = False Then
                        TxtTCSTaxcode.Enabled = True : TxtTCSTaxcode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : CmdHelpTCSTax.Enabled = True
                    End If
                Else
                    TxtTCSTaxcode.Enabled = False : TxtTCSTaxcode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : CmdHelpTCSTax.Enabled = False : TxtTCSTaxcode.Text = ""
                End If
            End If
        End If

        SpChEntry.maxRows = 0
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub CmbInvType_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmbInvType.SelectedIndexChanged
        On Error GoTo ErrHandler
        'Procedure Call To Select InvoiceSubTypeDescription From Sale Conf Acc. To Invoice Type
        Call SelectInvoiceSubTypeFromSaleConf((CmbInvType.Text))
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub cmbInvType_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles cmbInvType.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        If CmbInvSubType.Enabled Then CmbInvSubType.Focus()
                End Select
        End Select
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        GoTo EventExitSub
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub CmbInvType_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmbInvType.Leave
        On Error GoTo ErrHandler
        Dim strsql As String
        Dim blntcscheck As Boolean = False

        Select Case CmdGrpChEnt.mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                blnFIFO = False
                Select Case UCase(CmbInvType.Text)
                    Case "NORMAL INVOICE", "EXPORT INVOICE", "TRANSFER INVOICE"
                        txtRefNo.Enabled = True : txtRefNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : cmdempHelpRefNo.Enabled = True
                        txtAmendNo.Enabled = True : txtAmendNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        ctlPerValue.Enabled = False
                        ctlPerValue.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        ctlInsurance.Enabled = True
                        ctlInsurance.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        '101188073
                        TaxesEnableDisable(txtSaleTaxType)
                        TaxesHelpEnableDisable(CmdSaleTaxType)
                        TaxesLabelEnableDisable(lblSaltax_Per)
                        TaxesEnableDisable(txtSurchargeTaxType)
                        TaxesHelpEnableDisable(cmdSurchargeTaxCode)
                        TaxesLabelEnableDisable(lblSurcharge_Per)
                        '101188073
                        txtSurchargeTaxType.Text = "" : txtRefNo.Text = ""
                        ctlInsurance.Text = "" : txtSaleTaxType.Text = ""
                        lblCurrencyDes.Text = ""
                        If UCase(CmbInvType.Text) = "EXPORT INVOICE" Then
                            lblCurrency.Visible = True : lblCurrencyDes.Visible = True
                            lblExchangeRateLable.Visible = True : lblExchangeRateValue.Visible = True
                        Else
                            lblCurrency.Visible = False : lblCurrencyDes.Visible = False
                            lblExchangeRateLable.Visible = False : lblExchangeRateValue.Visible = False
                        End If
                        With SpChEntry
                            .Col = 15 : .Col2 = 15 : .BlockMode = True : .ColHidden = True : .BlockMode = False
                        End With
                        '101188073
                        TaxesEnableDisable(txtSDTType)
                        TaxesHelpEnableDisable(cmdSDTax_Help)
                        TaxesLabelEnableDisable(lblSDTax_Per)
                        '101188073
                    Case "JOBWORK INVOICE"
                        txtRefNo.Enabled = True : txtRefNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : cmdempHelpRefNo.Enabled = True
                        txtAmendNo.Enabled = True : txtAmendNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        ctlPerValue.Enabled = False
                        ctlPerValue.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        ctlInsurance.Enabled = False
                        ctlInsurance.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        '101188073
                        TaxesEnableDisable(txtSaleTaxType)
                        TaxesHelpEnableDisable(CmdSaleTaxType, True)
                        TaxesLabelEnableDisable(lblSaltax_Per)
                        TaxesEnableDisable(txtSurchargeTaxType)
                        TaxesHelpEnableDisable(cmdSurchargeTaxCode)
                        TaxesLabelEnableDisable(lblSurcharge_Per)
                        '101188073
                        lblCurrency.Visible = False : lblCurrencyDes.Visible = False
                        lblExchangeRateLable.Visible = False : lblExchangeRateValue.Visible = False
                        With SpChEntry
                            .Col = 15 : .Col2 = 15 : .BlockMode = True : .ColHidden = True : .BlockMode = False
                        End With
                        txtSurchargeTaxType.Text = "" : txtRefNo.Text = ""
                        ctlInsurance.Text = "" : txtSaleTaxType.Text = ""
                        lblCurrencyDes.Text = ""
                        '101188073
                        TaxesEnableDisable(txtSDTType, True)
                        TaxesHelpEnableDisable(cmdSDTax_Help, True)
                        TaxesLabelEnableDisable(lblSDTax_Per, True)
                        '101188073
                    Case "SAMPLE INVOICE", "CSM INVOICE"
                        txtRefNo.Enabled = False : txtRefNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdempHelpRefNo.Enabled = False
                        txtAmendNo.Enabled = False : txtAmendNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        ctlPerValue.Enabled = True
                        ctlPerValue.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        lblCurrency.Visible = False : lblCurrencyDes.Visible = False
                        lblExchangeRateLable.Visible = False : lblExchangeRateValue.Visible = False
                        txtSurchargeTaxType.Text = "" : txtRefNo.Text = ""
                        ctlInsurance.Text = "" : txtSaleTaxType.Text = ""
                        lblCurrencyDes.Text = ""
                        If UCase(CmbInvType.Text) = "TRANSFER INVOICE" Then
                            ctlInsurance.Enabled = True
                            ctlInsurance.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                            '101188073
                            TaxesEnableDisable(txtSaleTaxType, True)
                            TaxesHelpEnableDisable(CmdSaleTaxType, True)
                            TaxesLabelEnableDisable(lblSaltax_Per, True)
                            TaxesEnableDisable(txtSurchargeTaxType, True)
                            TaxesHelpEnableDisable(cmdSurchargeTaxCode, True)
                            TaxesLabelEnableDisable(lblSurcharge_Per, True)
                            '101188073
                            With SpChEntry
                                .Col = 15 : .Col2 = 15 : .BlockMode = True : .ColHidden = True : .BlockMode = False
                            End With
                            '101188073
                            TaxesEnableDisable(txtSDTType, True)
                            TaxesHelpEnableDisable(cmdSDTax_Help, True)
                            TaxesLabelEnableDisable(lblSDTax_Per, True)
                            '101188073
                        Else
                            ctlInsurance.Enabled = False
                            ctlInsurance.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                            '101188073
                            TaxesEnableDisable(txtSaleTaxType)
                            TaxesHelpEnableDisable(CmdSaleTaxType)
                            TaxesLabelEnableDisable(lblSaltax_Per)
                            TaxesEnableDisable(txtSurchargeTaxType)
                            TaxesHelpEnableDisable(cmdSurchargeTaxCode)
                            TaxesLabelEnableDisable(lblSurcharge_Per)
                            '101188073
                            With SpChEntry
                                .Col = 15 : .Col2 = 15 : .BlockMode = True : .ColHidden = False : .BlockMode = False
                                .Col = 15 : .Col2 = 15 : .BlockMode = True : .Lock = False : .BlockMode = False
                            End With
                            '101188073
                            TaxesEnableDisable(txtSDTType)
                            TaxesHelpEnableDisable(cmdSDTax_Help)
                            TaxesLabelEnableDisable(lblSDTax_Per)
                            '101188073
                        End If
                    Case "REJECTION"
                        ctlPerValue.Enabled = True
                        ctlPerValue.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        txtSurchargeTaxType.Text = "" : txtRefNo.Text = "" : txtAmendNo.Text = ""
                        ctlInsurance.Text = "" : txtSaleTaxType.Text = ""
                        lblCurrencyDes.Text = ""
                        ctlInsurance.Enabled = True
                        ctlInsurance.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        '101188073
                        TaxesEnableDisable(txtSaleTaxType)
                        TaxesHelpEnableDisable(CmdSaleTaxType)
                        TaxesLabelEnableDisable(lblSaltax_Per)
                        TaxesEnableDisable(txtSurchargeTaxType)
                        TaxesHelpEnableDisable(cmdSurchargeTaxCode)
                        TaxesLabelEnableDisable(lblSurcharge_Per)
                        '101188073
                        With SpChEntry
                            .Col = 15 : .Col2 = 15 : .BlockMode = True : .ColHidden = True : .BlockMode = False
                        End With
                        '101188073
                        TaxesEnableDisable(txtSDTType)
                        TaxesHelpEnableDisable(cmdSDTax_Help)
                        TaxesLabelEnableDisable(lblSDTax_Per)
                        '101188073
                End Select
        End Select
        SpChEntry.MaxRows = 0
        lblCreditTerm.Text = ""
        lblCreditTermDesc.Text = ""
        If Len(Trim(txtCustCode.Text)) > 0 Then
            strSQl = "select dbo.UDF_IRN_TCSREQUIRED( '" & gstrUNITID & "','" & txtCustCode.Text.Trim & "')"
            If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSql)) = True Then
                blntcscheck = True
            Else
                blntcscheck = False
            End If
            If blntcscheck = True Then
                Call checktcsvalue(CmbInvType.Text, CmbInvSubType.Text)
            Else
                If (UCase(Trim(CmbInvType.Text) = "NORMAL INVOICE") And (UCase(Trim(CmbInvSubType.Text)) = "SCRAP")) Then
                    If gblnGSTUnit = False Then
                        TxtTCSTaxcode.Enabled = True : TxtTCSTaxcode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : CmdHelpTCSTax.Enabled = True
                    End If
                Else
                    TxtTCSTaxcode.Enabled = False : TxtTCSTaxcode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : CmdHelpTCSTax.Enabled = False : TxtTCSTaxcode.Text = ""
                End If
            End If
        End If

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
                        If txtVehNo.Enabled Then txtVehNo.Focus()
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
    Private Sub CmdChallanNo_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdChallanNo.Click
        On Error GoTo ErrHandler
        Dim strHelpString As String
        Dim strChallanNo() As String
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                If Trim(txtLocationCode.Text) = "" Then
                    Call ConfirmWindow(10239, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO, 100)
                    If txtLocationCode.Enabled Then txtLocationCode.Focus()
                    Exit Sub
                End If
                If blnEOU_FLAG = True Then
                    strHelpString = "Select Doc_No,Invoice_Date from SalesChallan_Dtl where Location_code ='" & txtLocationCode.Text & "' and invoice_type <> 'EXP' and UNIT_CODE = '" & gstrUNITID & "'"
                Else
                    strHelpString = "Select Doc_No,Invoice_Date from SalesChallan_Dtl where Location_code ='" & txtLocationCode.Text & "' and UNIT_CODE = '" & gstrUNITID & "'"
                End If
        End Select
        strChallanNo = Me.ctlEMPHelpInvoiceEntry.ShowList(gstrCONNECTIONSERVER,gstrDSNName, gstrDatabaseName, strHelpString, "Challan/Invoice No")
        If UBound(strChallanNo) <= 0 Then Exit Sub
        If strChallanNo(0) = "0" Then
            MsgBox("No Challan/Invoice Available To Display", MsgBoxStyle.Information, "eMPro") : txtChallanNo.Text = "" : txtChallanNo.Focus() : Exit Sub
        Else
            txtChallanNo.Text = strChallanNo(0)
        End If
        txtChallanNo.Focus()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub CmdCustCodeHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdCustCodeHelp.Click
        Dim strCustMst As String
        Dim strCustHelp As String
        Dim rsCustMst As ClsResultSetDB
        Dim strCustomer() As String
        Dim blnNTRF_INV_GROUPCOMP As Boolean
        Dim strsql As String
        Dim blntcscheck As Boolean = False

        blnNTRF_INV_GROUPCOMP = Find_Value("SELECT ISNULL(TRF_INV_GROUPCOMP ,0) AS TRF_INV_GROUPCOMP  FROM SALES_PARAMETER WHERE UNIT_CODE = '" & gstrUNITID & "'")

        On Error GoTo ErrHandler
        If Len(Trim(txtLocationCode.Text)) = 0 Then
            Call ConfirmWindow(10116, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
            If txtLocationCode.Enabled Then txtLocationCode.Focus()
            Exit Sub
        End If
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                If UCase(Trim(mstrInvoiceType)) = "INV" Or UCase(Trim(mstrInvoiceType)) = "SMP" Or UCase(Trim(mstrInvoiceType)) = "TRF" Or UCase(Trim(mstrInvoiceType)) = "JOB" Or UCase(Trim(mstrInvoiceType)) = "EXP" Or UCase(Trim(mstrInvoiceType)) = "CSM" Or UCase(Trim(mstrInvoiceType)) = "ITD" Then
                    If blnNTRF_INV_GROUPCOMP = True And UCase(Trim(mstrInvoiceType)) = "TRF" Then
                        strCustHelp = "Select Customer_code,Cust_Name from Customer_Mst where UNIT_CODE = '" & gstrUNITID & "' and Group_Customer=1 and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
                    ElseIf blnNTRF_INV_GROUPCOMP = True And UCase(Trim(mstrInvoiceType)) = "ITD" Then
                        strCustHelp = "Select Customer_code,Cust_Name from Customer_Mst where UNIT_CODE = '" & gstrUNITID & "' and Group_Customer_InterDivision =1  and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
                    ElseIf blnNTRF_INV_GROUPCOMP = True And UCase(Trim(mstrInvoiceType)) = "INV" Then
                        strCustHelp = "Select Customer_code,Cust_Name from Customer_Mst where UNIT_CODE = '" & gstrUNITID & "' and Group_Customer=0 and Group_Customer_InterDivision =0 and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
                    Else
                        strCustHelp = "Select Customer_code,Cust_Name from Customer_Mst where UNIT_CODE = '" & gstrUNITID & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
                    End If
                Else
                    strCustHelp = "Select Vendor_code,Vendor_Name from Vendor_Mst where UNIT_CODE = '" & gstrUNITID & "'"
                End If
        End Select
        strCustomer = Me.ctlEMPHelpInvoiceEntry.ShowList(gstrCONNECTIONSERVER,gstrDSNName, gstrDatabaseName, strCustHelp, "Customer List")
        If UBound(strCustomer) <= 0 Then Exit Sub
        If strCustomer(0) = "0" Then
            If UCase(Trim(mstrInvoiceType)) = "INV" Or UCase(Trim(mstrInvoiceType)) = "SMP" Or UCase(Trim(mstrInvoiceType)) = "TRF" Or UCase(Trim(mstrInvoiceType)) = "JOB" Or UCase(Trim(mstrInvoiceType)) = "EXP" Or UCase(Trim(mstrInvoiceType)) = "CSM" Or UCase(Trim(mstrInvoiceType)) = "ITD" Then
                MsgBox("No Customer Available to Display") : txtCustCode.Text = "" : txtCustCode.Focus() : Exit Sub
            Else
                MsgBox("No Vendor Available to Display") : txtCustCode.Text = "" : txtCustCode.Focus() : Exit Sub
            End If
        Else
            txtCustCode.Text = strCustomer(0)
            lblCustCodeDes.Text = strCustomer(1)
        End If
        If Len(Trim(txtCustCode.Text)) > 0 Then
            rsCustMst = New ClsResultSetDB
            strCustMst = "SELECT Bill_Address1 + ', '  +  Bill_Address2 + ', ' + Bill_City + ' - ' + Bill_Pin as  invoiceAddress from Customer_Mst where Customer_code ='" & txtCustCode.Text & "'  and UNIT_CODE = '" & gstrUNITID & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
            rsCustMst.GetResult(strCustMst)
            If rsCustMst.GetNoRows > 0 Then

                lblAddressDes.Text = rsCustMst.GetValue("InvoiceAddress")
            End If

            rsCustMst.ResultSetClose()
        End If

        Call txtCustCode_Validating(txtCustCode, New System.ComponentModel.CancelEventArgs(False))
        If Len(Trim(txtCustCode.Text)) > 0 Then
            strSQl = "select dbo.UDF_IRN_TCSREQUIRED( '" & gstrUNITID & "','" & txtCustCode.Text.Trim & "')"
            If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSql)) = True Then
                blntcscheck = True
            Else
                blntcscheck = False
            End If
            If blntcscheck = True Then
                Call checktcsvalue(CmbInvType.Text, CmbInvSubType.Text)
            Else
                If (UCase(Trim(CmbInvType.Text) = "NORMAL INVOICE") And (UCase(Trim(CmbInvSubType.Text)) = "SCRAP")) Then
                    If gblnGSTUnit = False Then
                        TxtTCSTaxcode.Enabled = True : TxtTCSTaxcode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : cmdHelpTCSTax.Enabled = True
                    End If
                Else
                    TxtTCSTaxcode.Enabled = False : TxtTCSTaxcode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdHelpTCSTax.Enabled = False : TxtTCSTaxcode.Text = ""
                End If
            End If
        End If

        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Public Sub checktcsvalue(ByVal pInvType As String, ByVal pInvSubType As String)
        Dim rsTCSReq As ClsResultSetDB
        Try
            rsTCSReq = New ClsResultSetDB
            rsTCSReq.GetResult("Select isnull(REQD_TCS,0) as REQD_TCS , TCSTXRT_TYPE from saleConf Where UNIT_CODE='" + gstrUNITID + "' AND description ='" & Trim(pInvType) & "' and Sub_Type_Description='" & Trim(pInvSubType) & "' and  (fin_start_date <= getdate() and fin_end_date >= getdate())")
            If rsTCSReq.GetValue("REQD_TCS") = True Then
                txtTCSTaxCode.Enabled = True : txtTCSTaxCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : cmdHelpTCSTax.Enabled = True : txtTCSTaxCode.Text = rsTCSReq.GetValue("TCSTXRT_TYPE").ToString
                If CheckExistanceOfFieldData((txtTCSTaxCode.Text), "TxRt_Rate_No", "Gen_TaxRate", " (Tx_TaxeID='TCS')  and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))") Then
                    'lblTCSTaxPerDes.Text = CStr(GetTaxRate((txtTCSTaxCode.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " (Tx_TaxeID='TCS')"))
                End If
            Else
                If (UCase(Trim(pInvType) = "NORMAL INVOICE") And (UCase(Trim(pInvSubType)) = "SCRAP")) Then
                    If gblnGSTUnit = False Then
                        txtTCSTaxCode.Enabled = True : txtTCSTaxCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : cmdHelpTCSTax.Enabled = True
                    End If
                Else
                    TxtTCSTaxcode.Enabled = False : TxtTCSTaxcode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : CmdHelpTCSTax.Enabled = False : TxtTCSTaxcode.Text = ""
                End If

            End If
            rsTCSReq.ResultSetClose()

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub cmdECESSCode_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdECESSCode.Click
        On Error GoTo ErrHandler
        '101188073
        If gblnGSTUnit Then Exit Sub
        '101188073
        Dim strHelp As String
        Dim strSSTaxHelp() As String
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                '------------------Satvir Handa------------------------
                strHelp = "Select TxRt_Rate_No,TxRt_Percentage from Gen_TaxRate where tx_TaxeID ='ECS' and UNIT_CODE = '" & gstrUNITID & "' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date)) "
                '------------------Satvir Handa------------------------
                strSSTaxHelp = Me.ctlEMPHelpInvoiceEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strHelp, "E CESS Tax Help")
                If UBound(strSSTaxHelp) <= 0 Then Exit Sub
                If strSSTaxHelp(0) = "0" Then
                    Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK) : txtECESS.Text = "" : txtECESS.Focus() : Exit Sub
                Else
                    txtECESS.Text = strSSTaxHelp(0)
                    lblECESS_Per.Text = strSSTaxHelp(1)
                End If
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub cmdhelpSRVDI_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdhelpSRVDI.Click
        On Error GoTo ErrHandler
        Dim strHelp As String
        Dim strMktNagare() As String
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        Dim intmaxitems As Short
        Dim VarDelete As Object
        Dim varItemCode As Object
        Dim varDrgNo As Object
        Dim StrHelpSql As String
        Dim strdate As String
        With SpChEntry
            'To Check Non Deleted items No in Grid
            intmaxitems = 0 : intMaxLoop = .MaxRows
            For intLoopCounter = 1 To intMaxLoop
                VarDelete = Nothing
                Call .GetText(14, intLoopCounter, VarDelete)
                
                If UCase(Trim(VarDelete)) <> "D" Then
                    intmaxitems = intmaxitems + 1
                End If
            Next
            'To Fetch Item Code and Drawing No from Current Non-Deleted Row in Spread
            intMaxLoop = .MaxRows
            For intLoopCounter = 1 To intMaxLoop
                VarDelete = Nothing
                Call .GetText(14, intLoopCounter, VarDelete)
                If UCase(Trim(VarDelete)) <> "D" Then
                    varItemCode = Nothing
                    Call .GetText(1, intLoopCounter, varItemCode)
                    varDrgNo = Nothing
                    Call .GetText(2, intLoopCounter, varDrgNo)
                    Exit For
                End If
            Next
        End With
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                strdate = VB6.Format(dtpDateDesc.Value, gstrDateFormat)
                StrHelpSql = "select KanbanNo,UNLOC,USLOC,Sch_time,sch_date from MKT_Enagaredtl m"
                StrHelpSql = StrHelpSql & " where m.UNIT_CODE = '" & gstrUNITID & "' and  m.Account_code = '" & Trim(txtCustCode.Text) & "'"
                StrHelpSql = StrHelpSql & " and m.Item_code = '" & Trim(varItemCode) & "'"
                StrHelpSql = StrHelpSql & " and m.Cust_drgno = '" & varDrgNo & "'"
                StrHelpSql = StrHelpSql & " and m.quantity > ( select isnull(sum(b.sales_quantity),0) from salesChallan_dtl a inner join sales_dtl b on a.Unit_code = b.Unit_code and  a.location_code = b.location_code and a.doc_no=b.doc_no where m.kanbanNo  = a.srvdino and a.unit_code = '" & gstrUNITID & "' )"
                StrHelpSql = StrHelpSql & " order by sch_date desc, Sch_time asc"
                strMktNagare = Me.ctlEMPHelpInvoiceEntry.ShowList(gstrCONNECTIONSERVER,gstrDSNName, gstrDatabaseName, StrHelpSql, "eNagare Details")
                If UBound(strMktNagare) <= 0 Then Exit Sub
                If strMktNagare(0) = "0" Then
                    MsgBox("No Record Available to Display", MsgBoxStyle.Information, "eMPro") : txtSRVDI.Text = "" : txtSRVDI.Focus() : Exit Sub
                Else
                    txtSRVDI.Text = strMktNagare(0)
                    txtSRVLoc.Text = strMktNagare(1)
                    txtUsLoc.Text = strMktNagare(2)
                    txtSchTime.Text = strMktNagare(3)
                End If
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub CmdLocCodeHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdLocCodeHelp.Click
        Dim strLocationCode() As String
        On Error GoTo ErrHandler
        strLocationCode = Me.ctlEMPHelpInvoiceEntry.ShowList(gstrCONNECTIONSERVER,gstrDSNName, gstrDatabaseName, "Select distinct s.Location_Code,l.Description from Location_Mst l,SaleConf s where l.unit_code = s.Unit_code and l.unit_code = '" & gstrUNITID & "' and  s.Location_code = l.Location_code and s.Location_code like'" & txtLocationCode.Text & "%'", "Accounting Locations")
        If UBound(strLocationCode) <= 0 Then Exit Sub
        If strLocationCode(0) = "0" Then
            MsgBox("No Accounting Location Available to Display.") : txtLocationCode.Text = "" : txtLocationCode.Focus() : Exit Sub
        Else
            txtLocationCode.Text = strLocationCode(0)
            lblLocCodeDes.Text = strLocationCode(1)
        End If
        'Procedure Call To Select The Location Code Description
        Call SelectDescriptionForField("Description", "Location_Code", "Location_Mst", lblLocCodeDes, (txtLocationCode.Text))
        Call txtLocationCode_Validating(txtLocationCode, New System.ComponentModel.CancelEventArgs(False))
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub cmdRGPCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdRGPCancel.Click
        On Error GoTo ErrHandler
        mstrRGP = ""
        lblRGPDes.Text = ""
        fraRGPs.Visible = False
        txtCustCode.Focus()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub cmdRGPOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdRGPOK.Click
        On Error GoTo ErrHandler
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        mstrRGP = ""
        intMaxLoop = lvwRGPs.Items.Count
        For intLoopCounter = 0 To intMaxLoop - 1
            If lvwRGPs.Items.Item(intLoopCounter).Checked = True Then
                If Len(Trim(mstrRGP)) > 0 Then
                    mstrRGP = Trim(mstrRGP) & "§" & lvwRGPs.Items.Item(intLoopCounter).Text
                Else
                    mstrRGP = lvwRGPs.Items.Item(intLoopCounter).Text
                End If
            End If
        Next
        lblRGPDes.Text = Replace(mstrRGP, "§", ", ", 1)
        If Len(Trim(mstrRGP)) > 0 Then
            fraRGPs.Visible = False
        Else
            MsgBox("Select atleast one RGP from List", MsgBoxStyle.Information, "eMPro")
            lvwRGPs.Focus()
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub CmdSaleTaxType_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdSaleTaxType.Click
        Dim strHelp As String
        Dim strSTaxHelp() As String
        On Error GoTo ErrHandler
        '101188073
        If gblnGSTUnit Then Exit Sub
        '101188073
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                strHelp = "Select TxRT_Rate_No,TxRt_Percentage from Gen_taxRate where UNIT_CODE = '" & gstrUNITID & "' and Tx_TaxeID ='CST' OR Tx_TaxeID ='LST' or Tx_TaxeID='VAT'"
                strSTaxHelp = Me.ctlEMPHelpInvoiceEntry.ShowList(gstrCONNECTIONSERVER,gstrDSNName, gstrDatabaseName, strHelp, "S.Tax Help")
                If UBound(strSTaxHelp) <= 0 Then Exit Sub
                If strSTaxHelp(0) = "0" Then
                    Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK) : txtSaleTaxType.Text = "" : txtSaleTaxType.Focus() : Exit Sub
                Else
                    txtSaleTaxType.Text = strSTaxHelp(0)
                    lblSaltax_Per.Text = strSTaxHelp(1)
                End If
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub cmdSDTax_Help_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSDTax_Help.Click
        On Error GoTo ErrHandler
        '101188073
        If gblnGSTUnit Then Exit Sub
        '101188073
        Dim strHelp As String
        strHelp = ShowList(1, (txtSaleTaxType.MaxLength), "", "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND Tx_TaxeID='SDT'")
        If strHelp = "-1" Then 'If No Record Exists In The Table
            Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
            Exit Sub
        Else
            If Len(Trim(strHelp)) <> 0 Then
                txtSDTType.Text = Trim(strHelp)
                txtSDTType_Validating(txtSDTType, New System.ComponentModel.CancelEventArgs(False))
            End If
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub cmdSECESSCode_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSECESSCode.Click
        On Error GoTo ErrHandler
        '101188073
        If gblnGSTUnit Then Exit Sub
        '101188073
        Dim strHelp As String
        Dim strSSTaxHelp() As String
        Select Case CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                '------------------Satvir Handa------------------------
                strHelp = "Select TxRt_Rate_No,TxRt_Percentage from Gen_TaxRate where tx_TaxeID ='ECSSH' and UNIT_CODE = '" & gstrUNITID & "' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date)) "
                '------------------Satvir Handa------------------------
                strSSTaxHelp = Me.ctlEMPHelpInvoiceEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strHelp, "SE CESS Tax Help")
                If UBound(strSSTaxHelp) <= 0 Then Exit Sub
                If strSSTaxHelp(0) = "0" Then
                    Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK) : txtECESS.Text = "" : txtECESS.Focus() : Exit Sub
                Else
                    txtSECESS.Text = strSSTaxHelp(0)
                    lblSECESS_Per.Text = strSSTaxHelp(1)
                End If
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub cmdSurchargeTaxCode_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSurchargeTaxCode.Click
        On Error GoTo ErrHandler
        '101188073
        If gblnGSTUnit Then Exit Sub
        '101188073
        Dim strHelp As String
        Dim strSSTaxHelp() As String
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                strHelp = "Select TxRt_Rate_No,TxRt_Percentage from Gen_TaxRate where tx_TaxeID ='SST' and UNIT_CODE = '" & gstrUNITID & "'"
                strSSTaxHelp = Me.ctlEMPHelpInvoiceEntry.ShowList(gstrCONNECTIONSERVER,gstrDSNName, gstrDatabaseName, strHelp, "S.Sales Tax Help")
                If UBound(strSSTaxHelp) <= 0 Then Exit Sub
                If strSSTaxHelp(0) = "0" Then
                    Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK) : txtSurchargeTaxType.Text = "" : txtSurchargeTaxType.Focus() : Exit Sub
                Else
                    txtSurchargeTaxType.Text = strSSTaxHelp(0)
                    lblSurcharge_Per.Text = strSSTaxHelp(1)
                End If
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0016_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo Err_Handler
        If KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0 Then
            Call ctlFormHeader1_Click(ctlFormHeader1, New System.EventArgs())
        End If
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0016_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Escape
                'If user press the ESC Key ,the Form will be in View Mode
                If Me.CmdGrpChEnt.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                    If ConfirmWindow(10053, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                        Call Me.CmdGrpChEnt.Revert()
                        lblSDTax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        lblSDTax_Per.Text = "0.00"
                        Call EnableControls(False, Me, True)
                        'In View Mode Disable Combo Of Invoice Type and Inv. Sub type
                        With SpChEntry
                            .Col = 15 : .Col2 = 15 : .BlockMode = True : .ColHidden = True : .BlockMode = False
                        End With
                        CmbInvType.Visible = False : CmbInvSubType.Visible = False
                        lblInvSubType.Visible = False : lblInvType.Visible = False
                        txtLocationCode.Enabled = True : txtLocationCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : lblLocCodeDes.Text = ""
                        txtChallanNo.Enabled = True : txtChallanNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        CmdLocCodeHelp.Enabled = True : CmdChallanNo.Enabled = True : Me.SpChEntry.Enabled = False
                        Me.CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                        Me.CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                        Me.CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                        TaxesLabelEnableDisable(lblSaltax_Per, True)
                        TaxesLabelEnableDisable(lblSurcharge_Per, True)
                        TaxesLabelEnableDisable(lblECESS_Per, True)
                        TaxesLabelEnableDisable(lblSECESS_Per, True)
                        If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                            If blnEOU_FLAG = False Then
                                CmbInvType.SelectedIndex = 2 : CmbInvSubType.SelectedIndex = 2
                            Else
                                CmbInvType.SelectedIndex = 1 : CmbInvSubType.SelectedIndex = 2
                            End If
                        End If
                        gblnCancelUnload = False
                        gblnFormAddEdit = False
                        With Me.SpChEntry
                            .MaxRows = 1 : .set_RowHeight(1, 300)
                            .Row = 1 : .Row2 = 1 : .Col = 1 : .Col2 = .MaxCols : .BlockMode = True : .Text = "" : .BlockMode = False
                        End With
                        'Get Server Date
                        lblDateDes.Text = VB6.Format(GetServerDate(), gstrDateFormat)
                        dtpDateDesc.Visible = False
                        If Len(Trim(mstrLocation)) > 0 Then
                            txtLocationCode.Text = mstrLocation
                        End If
                        txtLocationCode.Focus()
                        chkDTRemoval.Enabled = True
                        chkDTRemoval.CheckState = System.Windows.Forms.CheckState.Unchecked
                        dtpRemoval.Enabled = False
                        dtpRemovalTime.Enabled = False
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
    Private Sub frmMKTTRN0016_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo ErrHandler
        Dim rsParameterData As ClsResultSetDB
        Dim strParamQuery As String
        'Add Form Name To Window List
        mintIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader1.Tag)
        'Fill Lebels From Resource File
        Call FillLabelFromResFile(Me)
        Call FitToClient(Me, FraChEnt, ctlFormHeader1, CmdGrpChEnt, 650)
        'Set Help Pictures At Command Button
        CmdLocCodeHelp.Image = My.Resources.ico111.ToBitmap
        CmdChallanNo.Image = My.Resources.ico111.ToBitmap
        CmdCustCodeHelp.Image = My.Resources.ico111.ToBitmap
        CmdSaleTaxType.Image = My.Resources.ico111.ToBitmap
        cmdempHelpRefNo.Image = My.Resources.ico111.ToBitmap
        ctlPerValue.Text = 1
        'set the date format
        dtpRemoval.Format = DateTimePickerFormat.Custom
        dtpDateDesc.Format = DateTimePickerFormat.Custom
        dtpRemoval.CustomFormat = gstrDateFormat
        dtpDateDesc.CustomFormat = gstrDateFormat
        'Check If Company is 100% EOU then CVD SVD fields are SHOWN
        gobjDB = New ClsResultSetDB
        If gobjDB.GetResult("Select EOU_FLAG From Company_Mst where UNIT_CODE = '" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic) Then
            If gobjDB.GetNoRows > 0 Then
                blnEOU_FLAG = gobjDB.GetValue("EOU_FLAG")
            End If
        End If
        'Initially Disable All Controls
        Call EnableControls(False, Me, True)
        TaxesLabelEnableDisable(lblSaltax_Per, True)
        TaxesLabelEnableDisable(lblSurcharge_Per, True)
        'Get Server Date
        lblDateDes.Text = VB6.Format(GetServerDate(), gstrDateFormat)
        'Date is Also Added in DatePicker,and Its Visible Property is set to False
        With dtpDateDesc
            .Value = GetServerDate()
            .Visible = False
        End With
        'Add Transport Type To Combo
        Call AddTransPortTypeToCombo()
        txtLocationCode.Enabled = True : txtLocationCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        txtChallanNo.Enabled = True : txtChallanNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        CmdLocCodeHelp.Enabled = True : CmdChallanNo.Enabled = True
        Me.SpChEntry.Enabled = False
        Me.CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
        Me.CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
        Me.CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
        'Set Column Headers
        With Me.SpChEntry
            .set_ColWidth(0, 300)
            '101188073
            .MaxCols = COMP_CESS_TYPE
            '101188073
            .Row = 0 : .Col = 1 : .Text = "Internal Part No." : .set_ColWidth(1, 1700)
            .Row = 0 : .Col = 2 : .Text = "Cust.Part No." : .set_ColWidth(2, 1700)
            .Row = 0 : .Col = 3 : .Text = "Rate" : .set_ColWidth(2, 1500)
            .Row = 0 : .Col = 4 : .Text = "Cust Material" : .set_ColWidth(4, 1800)
            .Row = 0 : .Col = 5 : .Text = "Quantity" : .set_ColWidth(5, 1500)
            .Row = 0 : .Col = 6 : .Text = "Packing(%)" : .set_ColWidth(6, 1500)
            .Row = 0 : .Col = 7 : .Text = "EXC(%)"
            .Row = 0 : .Col = 8 : .Text = "CVD(%)"
            .Row = 0 : .Col = 9 : .Text = "SAD(%)"
            If Not blnEOU_FLAG Then
                .Col = 8 : .Col2 = 8
                .ColHidden = True
                .BlockMode = False
                .Col = 9 : .Col2 = 9
                .BlockMode = True
                .ColHidden = True
                .BlockMode = False
            End If
            .Row = 0 : .Col = 10 : .Text = "Others"
            .Row = 0 : .Col = 11 : .Text = "From Box" : .set_ColWidth(11, 1500)
            .Row = 0 : .Col = 12 : .Text = "To Box"
            .Row = 0 : .Col = 13 : .Text = "Cumulative Boxes" : .set_ColWidth(13, 1600)
            .Col = 13 : .Col2 = 13
            : .Lock = True : .BlockMode = False
            .Row = 0 : .Col = 14 : .Text = "Delete"
            .Col = 14 : .Col2 = 14
            .ColHidden = True : .BlockMode = False
            .Row = 0 : .Col = 15 : .Text = "Tool Cost"
            .Col = 15 : .Col2 = 15
            .ColHidden = True : .BlockMode = False
            .Row = 0 : .Col = 16 : .Text = "Rate"
            .Row = 0 : .Col = 17 : .Text = "Cust Supp Mat" : .set_ColWidth(17, 1500)
            .Row = 0 : .Col = 18 : .Text = "Others"
            .Row = 0 : .Col = 19 : .Text = "tool cost"
            .Row = 0 : .Col = 20 : .Text = "Tariff Code"
            .Row = 0 : .Col = 21 : .Text = "Edit Flag"
            .Row = 0 : .Col = 22 : .Text = "Bin Quantity" : .set_ColWidth(22, 1500)
            .Col = 16 : .Col2 = 21 : .BlockMode = True : .ColHidden = True : .BlockMode = False
            '101188073
            .Row = 0 : .Col = IS_HSN_SAC : .Text = "HSN/SAC" : .set_ColWidth(IS_HSN_SAC, 1500)
            .Row = 0 : .Col = HSN_SAC_CODE : .Text = "HSN/SAC CODE" : .set_ColWidth(HSN_SAC_CODE, 1700)
            .Row = 0 : .Col = CGST_TYPE : .Text = "CGST TYPE" : .set_ColWidth(CGST_TYPE, 1500)
            .Row = 0 : .Col = SGST_TYPE : .Text = "SGST TYPE" : .set_ColWidth(SGST_TYPE, 1500)
            .Row = 0 : .Col = IGST_TYPE : .Text = "IGST TYPE" : .set_ColWidth(IGST_TYPE, 1500)
            .Row = 0 : .Col = UTGST_TYPE : .Text = "UTGST TYPE" : .set_ColWidth(UTGST_TYPE, 1500)
            .Row = 0 : .Col = COMP_CESS_TYPE : .Text = "COMP. CESS TYPE" : .set_ColWidth(COMP_CESS_TYPE, 1500)
            If gblnGSTUnit Then
                .BlockMode = True : .Col = IS_HSN_SAC : .Col2 = COMP_CESS_TYPE : .ColHidden = False : .BlockMode = False
            Else
                .BlockMode = True : .Col = IS_HSN_SAC : .Col2 = COMP_CESS_TYPE : .ColHidden = True : .BlockMode = False
            End If
            '101188073
        End With
        'Function Call To Add Invoice Types In The Inv. Type Combo Box
        Call SelectInvoiceTypeFromSaleConf()
        'In View Mode Disable Combo Of Invoice Type and Inv. Sub type
        CmbInvType.Visible = False : CmbInvSubType.Visible = False
        lblInvSubType.Visible = False : lblInvType.Visible = False
        'Add Row
        Call addRowAtEnterKeyPress(1)
        fraRGPs.Visible = False
        lblRGPDes.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        lblCurrency.Visible = False
        lblCurrencyDes.Visible = False
        lblExchangeRateLable.Visible = False
        lblExchangeRateValue.Visible = False
        chkDTRemoval.Enabled = True
        chkDTRemoval.CheckState = System.Windows.Forms.CheckState.Unchecked
        dtpRemoval.Enabled = False
        dtpRemovalTime.Enabled = False
        dtpRemoval.Value = GetServerDate()
        dtpRemovalTime.Value = GetServerDate()
        TaxesLabelEnableDisable(lblSECESS_Per, True)
        strParamQuery = "SELECT decimal_place FROM currency_mst where currency_code='" & lblCurrencyDes.Text & "' and UNIT_CODE = '" & gstrUNITID & "'"
        rsParameterData = New ClsResultSetDB
        rsParameterData.GetResult(strParamQuery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsParameterData.GetNoRows > 0 Then
            mIntDecimalPlace = rsParameterData.GetValue("decimal_place")
        End If
        rsParameterData.ResultSetClose()
        rsParameterData = Nothing
        chkFOC.Visible = False
        If Not Directory.Exists(gstrLocalCDrive + "EmproInv") Then
            Directory.CreateDirectory(gstrLocalCDrive + "EmproInv")
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0016_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo ErrHandler
        mdifrmMain.CheckFormName = mintIndex
        txtLocationCode.Focus()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0016_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Me.Tag) = False
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0016_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        On Error GoTo ErrHandler
        Dim enmValue As eMPowerFunctions.ConfirmWindowReturnEnum
        If UnloadMode >= 0 And UnloadMode <= 5 Then
            If Me.CmdGrpChEnt.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                enmValue = ConfirmWindow(10055, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNOCANCEL, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION)
                If enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_NO Or enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                    If enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                        'Save data before saving
                        Call CmdGrpChEnt_ButtonClick(CmdGrpChEnt, New UCActXCtl.UCbtndgrp.ButtonClickEventArgs(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE))
                    ElseIf enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_NO Then
                        gblnCancelUnload = False
                        gblnFormAddEdit = False
                    End If
                Else
                    'Set Global VAriable
                    gblnCancelUnload = True
                    gblnFormAddEdit = True
                    Me.CmdGrpChEnt.Focus()
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
    Private Sub frmMKTTRN0016_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error GoTo ErrHandler
        Me.Dispose() 'Assign form to nothing
        frmModules.NodeFontBold(Me.Tag) = False
        mdifrmMain.RemoveFormNameFromWindowList = mintIndex
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub addRowAtEnterKeyPress(ByRef pintRows As Short)
        On Error GoTo ErrHandler
        Dim intRowHeight As Short
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        Dim intNoRows As Short
        Dim VarDelete As Object
        Dim intDecimalPlaces As Short
        With Me.SpChEntry
            .CursorStyle = FPSpreadADO.CursorStyleConstants.CursorStyleArrow
            If .MaxRows > 0 Then
                If ValidRowData(.ActiveRow) = True Then
                    '*****To Chaeck Nomber of rows Already Added
                    intMaxLoop = SpChEntry.MaxRows
                    intNoRows = 0
                    For intLoopCounter = 1 To intMaxLoop
                        VarDelete = Nothing
                        Call .GetText(14, intLoopCounter, VarDelete)
                        If UCase(VarDelete) <> "D" Then
                            intNoRows = intNoRows + 1
                        End If
                    Next
                    If intNoRows < 11 Then
                        For intRowHeight = 1 To pintRows
                            .MaxRows = .MaxRows + 1
                            If Len(Trim(mCurrencyCode)) > 0 Then
                                intDecimalPlaces = ToGetDecimalPlaces(mCurrencyCode)
                                SetMaxLengthInSpread(intDecimalPlaces)
                            Else
                                SetMaxLengthInSpread(0)
                            End If
                            .Row = .MaxRows
                            .set_RowHeight(.Row, 300)
                            If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                                Call .SetText(14, .MaxRows, "A")
                            End If
                        Next intRowHeight
                        If .MaxRows > 4 Then .ScrollBars = FPSpreadADO.ScrollBarsConstants.ScrollBarsBoth
                    Else
                        If CmdGrpChEnt.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                            MsgBox("Can Not Enter Items More then 11.", MsgBoxStyle.Information, "eMPro")
                        End If
                    End If
                    Exit Sub
                End If
            Else
                For intRowHeight = 1 To pintRows
                    .MaxRows = .MaxRows + 1
                    If Len(Trim(mCurrencyCode)) > 0 Then
                        intDecimalPlaces = ToGetDecimalPlaces(mCurrencyCode)
                        SetMaxLengthInSpread(intDecimalPlaces)
                    Else
                        SetMaxLengthInSpread(0)
                    End If
                    .Row = .MaxRows
                    .set_RowHeight(.Row, 300)
                    If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                        Call .SetText(14, .MaxRows, "A")
                    End If
                Next intRowHeight
            End If
        End With
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Function SelectInvoiceTypeFromSaleConf() As Object
        On Error GoTo ErrHandler
        Dim strSaleConfSql As String
        Dim rsSaleConf As ClsResultSetDB
        Dim intRecCount As Short
        Dim intLoopCounter As Short
        If blnEOU_FLAG = False Then
            strSaleConfSql = "Select Distinct(Description) from SaleConf where UNIT_CODE = '" & gstrUNITID & "' and Invoice_Type Not in('STX') and (fin_start_date <= getdate() and fin_end_date >= getdate())"
        Else
            strSaleConfSql = "Select Distinct(Description) from SaleConf where UNIT_CODE = '" & gstrUNITID & "' and Invoice_Type Not in('EXP','STX') and (fin_start_date <= getdate() and fin_end_date >= getdate())"
        End If
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
        On Error GoTo ErrHandler
        Dim strSaleConfSql As String
        Dim rsSaleConf As ClsResultSetDB
        Dim intRecCount As Short
        Dim intLoopCounter As Short
        If pstrInvType = "CSM INVOICE" Then
            chkFOC.Checked = True
        Else
            chkFOC.Checked = False
        End If
        strSaleConfSql = "Select Distinct(Sub_Type_Description) from SaleConf where Description='" & Trim(pstrInvType) & "' and UNIT_CODE = '" & gstrUNITID & "' and (fin_start_date <= getdate() and fin_end_date >= getdate())"
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
        '****************************************************
        'Created By     -  Nisha
        'Description    -  To Select The Field Description In The Description Labels
        'Arguments      -  pstrFieldName1 - Field Name1,pstrFieldName2 - Field Name2,pstrTableName - Table Name
        '               -  pContName - Name Of The Control where Caption Is To Be Set
        '               -  pstrControlText - Field Text
        '****************************************************
        On Error GoTo ErrHandler
        Dim strDesSql As String 'Declared to make Select Query
        Dim rsDescription As ClsResultSetDB
        If pstrFieldName2 = "Customer_Code" Then
            strDesSql = "Select " & Trim(pstrFieldName1) & " from " & Trim(pstrTableName) & " where " & Trim(pstrFieldName2) & "='" & Trim(pstrControlText) & "' and UNIT_CODE = '" & gstrUNITID & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
        Else
            strDesSql = "Select " & Trim(pstrFieldName1) & " from " & Trim(pstrTableName) & " where " & Trim(pstrFieldName2) & "='" & Trim(pstrControlText) & "' and UNIT_CODE = '" & gstrUNITID & "'"
        End If

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
    Private Sub txtamendno_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAmendNo.TextChanged
        If Me.CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            SpChEntry.MaxRows = 0 : mstrItemCode = "" : txtSaleTaxType.Text = "" : txtSurchargeTaxType.Text = "" : lblCurrencyDes.Text = ""
            lblCreditTerm.Text = ""
            lblCreditTermDesc.Text = ""
        End If
    End Sub
    Private Sub txtAmendNo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAmendNo.Enter
        On Error GoTo ErrHandler
        With txtAmendNo
            .SelectionStart = 0
            .SelectionLength = Len(txtAmendNo.Text)
        End With
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub txtAmendNo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtAmendNo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        If Len(txtAmendNo.Text) > 0 Then
                            Call txtAmendNo_Validating(txtAmendNo, New System.ComponentModel.CancelEventArgs(False))
                        Else
                            If Not (CmbInvType.Text = "JOBWORK INVOICE") Then
                                txtCarrServices.Focus()
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
    Private Sub txtAmendNo_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtAmendNo.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = 112 Then
            If cmdempHelpRefNo.Enabled Then Call cmdempHelpRefNo_Click(cmdempHelpRefNo, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtAmendNo_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtAmendNo.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        'Only if Some Ref No. is Added
        If Trim(txtRefNo.Text) <> "" Then
            'if Some Amend No is Entered
            If Trim(txtAmendNo.Text) <> "" Then
                If SelectDataFromTable("Amendment_No", "Cust_ORD_HDR", " Where Account_Code = '" & Trim(txtCustCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "' And Cust_Ref = '" & Trim(txtRefNo.Text) & "' And Active_Flag = 'A'  AND  Amendment_No <> '' AND  Amendment_No = '" & Trim(txtAmendNo.Text) & "'") <> "" Then
                    'Verified,Set focus to Another Control
                    Call displayDeatilsfromCustOrdHdr()
                    If Not (CmbInvType.Text = "JOBWORK INVOICE") Then
                        txtCarrServices.Focus()
                    End If
                Else
                    MsgBox("Entered Amendment Number for Ref No." & txtRefNo.Text & vbCr & " does not Exist or is Not Active.", MsgBoxStyle.Information, "eMPro")
                    Cancel = False
                    txtAmendNo.Text = ""
                    txtAmendNo.Focus()
                End If
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
        On Error GoTo ErrHandler
        Me.txtChallanNo.SelectionStart = 0
        Me.txtChallanNo.SelectionLength = Len(Me.txtChallanNo.Text)
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtChallanNo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtChallanNo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
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
        'Allow only Numbers
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
        Dim strInvoiceType As String
        Dim strCondition As String
        Dim rsSalesChallan As ClsResultSetDB
        Dim rsChallanEntry As ClsResultSetDB
        On Error GoTo ErrHandler
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                If Len(txtChallanNo.Text) > 0 Then
                    'Check Existance Of Doc No In The SalesChallan_Dtl
                    If blnEOU_FLAG = True Then
                        strCondition = "Invoice_type <> 'EXP'"
                    Else
                        strCondition = ""
                    End If
                    If CheckExistanceOfFieldData((txtChallanNo.Text), "Doc_No", "SalesChallan_Dtl", strCondition) Then
                        'If Challan No. Exists
                        'Get Data From Challan_Dtl,Cust_Ord_Dtl,Sales_Dtl
                        If Len(txtLocationCode.Text) > 0 Then
                            If GetDataInViewMode() Then 'if record found
                                rsChallanEntry = New ClsResultSetDB
                                rsChallanEntry.GetResult("Select a.Description,a.Sub_Type_Description from SaleConf a,SalesChallan_Dtl b where a.UNIT_CODE = b.UNIT_CODE AND a.UNIT_CODE = '" & gstrUNITID & "' AND Doc_No = " & txtChallanNo.Text & " and a.Invoice_Type = b.Invoice_type and a.Sub_type = b.Sub_Category and a.Location_code = b.Location_code and (fin_start_date <= getdate() and fin_end_date >= getdate())")
                                strInvoiceType = UCase(rsChallanEntry.GetValue("Description"))
                                rsChallanEntry.ResultSetClose()
                                If UCase(strInvoiceType) <> "SAMPLE INVOICE" And UCase(strInvoiceType) <> "CSM INVOICE" Then
                                    With SpChEntry
                                        .Col = 15 : .Col2 = 15 : .BlockMode = True : .ColHidden = True : .BlockMode = False
                                    End With
                                Else
                                    With SpChEntry
                                        .Col = 15 : .Col2 = 15 : .BlockMode = True : .ColHidden = False : .BlockMode = False
                                        .Col = 15 : .Col2 = 15 : .BlockMode = True : .Lock = False : .BlockMode = False
                                    End With
                                End If
                                With SpChEntry
                                    .Enabled = True
                                    .Col = 0 : .Col2 = .MaxCols : .Row = 1 : .Row2 = .MaxRows : .BlockMode = True : .Lock = True : .BlockMode = False
                                End With
                                rsSalesChallan = New ClsResultSetDB
                                rsSalesChallan.GetResult("Select Bill_Flag from SalesChallan_Dtl where Location_Code = '" & txtLocationCode.Text & "' and Doc_No = " & txtChallanNo.Text & " and UNIT_CODE = '" & gstrUNITID & "'")
                                If rsSalesChallan.GetNoRows > 0 Then
                                    If rsSalesChallan.GetValue("Bill_Flag") = True Then
                                        CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                                        CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                                    Else
                                        CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = True
                                        CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = True
                                    End If
                                End If
                                rsSalesChallan.ResultSetClose()
                            Else 'if no record found then display message
                                Call ConfirmWindow(10414, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                            End If
                            CmdGrpChEnt.Focus()
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
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtCustCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCustCode.TextChanged
        On Error GoTo ErrHandler
        If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            lblCustCodeDes.Text = ""
            txtRefNo.Text = ""
            SpChEntry.MaxRows = 0
            mstrItemCode = ""
            lblAddressDes.Text = ""
            fraRGPs.Visible = False
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
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        If Len(txtCustCode.Text) > 0 Then
                            Call txtCustCode_Validating(txtCustCode, New System.ComponentModel.CancelEventArgs(False))
                        Else
                            'ISSUE ID :10465802
                            'If (UCase(CmbInvType.Text) = "NORMAL INVOICE") Or (UCase(CmbInvType.Text) = "JOBWORK INVOICE") Or (UCase(CmbInvType.Text) = "EXPORT INVOICE") Then
                            If (UCase(CmbInvType.Text) = "NORMAL INVOICE") Or (UCase(CmbInvType.Text) = "JOBWORK INVOICE") Or (UCase(CmbInvType.Text) = "EXPORT INVOICE") Or (UCase(CmbInvType.Text) = "TRANSFER INVOICE") Then
                                If (CmbInvSubType.Text <> "SCRAP") Then
                                    txtRefNo.Focus()
                                Else
                                    txtCarrServices.Focus()
                                End If
                            Else
                                txtCarrServices.Focus()
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
    Private Sub txtcustcode_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCustCode.KeyUp
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
        Dim rsCustMst As ClsResultSetDB
        Dim strCustMst As String
        Dim blnNTRF_INV_GROUPCOMP As Boolean = False
        Dim strcondGroupcompany As String

        On Error GoTo ErrHandler
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                blnNTRF_INV_GROUPCOMP = Find_Value("SELECT ISNULL(TRF_INV_GROUPCOMP ,0) AS TRF_INV_GROUPCOMP  FROM SALES_PARAMETER WHERE UNIT_CODE = '" & gstrUNITID & "'")

                If blnNTRF_INV_GROUPCOMP = True And UCase(Trim(mstrInvoiceType)) = "TRF" Then
                    strcondGroupcompany = " and group_customer =1 "
                ElseIf blnNTRF_INV_GROUPCOMP = True And UCase(Trim(mstrInvoiceType)) = "ITD" Then
                    strcondGroupcompany = " and Group_Customer_InterDivision =1 "
                ElseIf blnNTRF_INV_GROUPCOMP = True And UCase(Trim(mstrInvoiceType)) = "INV" Then
                    strcondGroupcompany = " and group_customer =0 "
                Else
                    strcondGroupcompany = ""
                End If


                If Len(txtCustCode.Text) > 0 Then
                    If UCase(Trim(mstrInvoiceType)) = "INV" Or UCase(Trim(mstrInvoiceType)) = "SMP" Or UCase(Trim(mstrInvoiceType)) = "TRF" Or UCase(Trim(mstrInvoiceType)) = "JOB" Or UCase(Trim(mstrInvoiceType)) = "EXP" Or UCase(Trim(mstrInvoiceType)) = "CSM" Or UCase(Trim(mstrInvoiceType)) = "ITD" Then
                        If CheckExistanceOfFieldData((txtCustCode.Text), "Customer_Code", "Customer_Mst", "((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))" & strcondGroupcompany & "") Then
                            Call SelectDescriptionForField("Cust_Name", "Customer_Code", "Customer_Mst", lblCustCodeDes, Trim(txtCustCode.Text))
                            'ISSUE ID :10465802
                            'If (UCase(CmbInvType.Text) = "NORMAL INVOICE") Or (UCase(CmbInvType.Text) = "JOBWORK INVOICE") Or (UCase(CmbInvType.Text) = "EXPORT INVOICE") Then
                            If (UCase(CmbInvType.Text) = "NORMAL INVOICE") Or (UCase(CmbInvType.Text) = "JOBWORK INVOICE") Or (UCase(CmbInvType.Text) = "EXPORT INVOICE") Or (UCase(CmbInvType.Text) = "TRANSFER INVOICE") Then
                                If UCase(CmbInvSubType.Text) <> "SCRAP" Then
                                    txtRefNo.Focus()
                                Else
                                    If SpChEntry.MaxRows = 0 Then
                                        Call addRowAtEnterKeyPress(1)
                                        Call ChangeCellTypeStaticText()
                                        System.Windows.Forms.Application.DoEvents()
                                    End If
                                    txtCarrServices.Focus()
                                End If
                            Else
                                If SpChEntry.MaxRows = 0 Then
                                    Call addRowAtEnterKeyPress(1)
                                    Call ChangeCellTypeStaticText()
                                    System.Windows.Forms.Application.DoEvents()
                                End If
                                txtCarrServices.Focus()
                            End If
                        Else
                            Cancel = True
                            Call ConfirmWindow(10417, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                            txtCustCode.Text = ""
                            txtCustCode.Focus()
                        End If
                        '***To Display invoice Address of Customer
                        If Len(Trim(txtCustCode.Text)) > 0 Then
                            rsCustMst = New ClsResultSetDB
                            strCustMst = "SELECT Bill_Address1 + ', '  +  Bill_Address2 + ', ' + Bill_City + ' - ' + Bill_Pin as  invoiceAddress from Customer_Mst where Customer_code ='" & txtCustCode.Text & "' and UNIT_CODE = '" & gstrUNITID & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
                            rsCustMst.GetResult(strCustMst)
                            If rsCustMst.GetNoRows > 0 Then

                                lblAddressDes.Text = rsCustMst.GetValue("InvoiceAddress")
                            End If
                            rsCustMst.ResultSetClose()
                        End If
                    Else
                        If CheckExistanceOfFieldData((txtCustCode.Text), "Vendor_Code", "Vendor_Mst") Then
                            Call SelectDescriptionForField("Vendor_name", "Vendor_Code", "Vendor_Mst", lblCustCodeDes, Trim(txtCustCode.Text))
                            If txtRefNo.Enabled Then
                                If SpChEntry.MaxRows = 0 Then
                                    Call addRowAtEnterKeyPress(1)
                                    Call ChangeCellTypeStaticText()
                                End If
                                txtRefNo.Focus()
                            Else
                                If SpChEntry.MaxRows = 0 Then
                                    Call addRowAtEnterKeyPress(1)
                                    Call ChangeCellTypeStaticText()
                                    System.Windows.Forms.Application.DoEvents()
                                End If
                                txtCarrServices.Focus()
                            End If
                        Else
                            Cancel = True
                            Call ConfirmWindow(10417, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                            txtCustCode.Text = ""
                            txtCustCode.Focus()
                        End If
                    End If
                    '***for rgpListadd
                    If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                        If UCase(CmbInvType.Text) = "JOBWORK INVOICE" Then
                            If MsgBox("Would Like to Follow FIFO Method For JobWork Material Process.", MsgBoxStyle.YesNo, "eMPro") = 7 Then
                                blnFIFO = False
                                mstrRGP = ""
                                If AddDataTolstRGPs() = True Then
                                    fraRGPs.Visible = True
                                Else
                                    MsgBox("No RGP's in last 180 days for this Customer.", MsgBoxStyle.Information, "eMPro")
                                    Cancel = True
                                    txtRefNo.Text = ""
                                    txtRefNo.Focus()
                                End If
                            Else
                                blnFIFO = True
                                mstrRGP = ""
                            End If
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
    Private Sub txtEcess_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtECESS.TextChanged
        On Error GoTo ErrHandler
        If Trim(txtECESS.Text) = "" Then
            lblECESS_Per.Text = "0.00"
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub txtEcess_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtECESS.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F1 Then
            Call cmdECESSCode_Click(cmdECESSCode, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtECESS_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtECESS.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        txtRemarks.Focus()
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
    Private Sub txtECESS_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtECESS.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        '101188073
        If gblnGSTUnit Then Exit Sub
        '101188073
        If Trim(txtECESS.Text) <> "" Then
            '------------------Satvir Handa------------------------
            If CheckExistanceOfFieldData((txtECESS.Text), "TxRt_Rate_No", "Gen_TaxRate", " Tx_TaxeID='ECS' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))") Then
                '------------------Satvir Handa------------------------
                lblECESS_Per.Text = CStr(GetTaxRate((txtECESS.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='ECS'"))
                If SpChEntry.Enabled Then
                    With Me.SpChEntry
                        .Row = 1 : .Col = 1 : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                    End With
                End If
            Else
                Call ConfirmWindow(10248, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                Cancel = True
                txtECESS.Text = ""
                If txtECESS.Enabled Then txtECESS.Focus()
            End If
        Else
            If SpChEntry.Enabled Then
                With Me.SpChEntry
                    .Row = 1 : .Col = 5 : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                End With
            End If
        End If
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtLocationCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLocationCode.TextChanged
        On Error GoTo ErrHandler
        If Len(Trim(txtLocationCode.Text)) = 0 Then
            Select Case Me.CmdGrpChEnt.Mode
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                    Call RefreshForm("LOCATION")
            End Select
        End If
        txtCustCode.Text = ""
        lblCustCodeDes.Text = ""
        lblLocCodeDes.Text = ""
        txtRefNo.Text = ""
        SpChEntry.MaxRows = 0
        mstrItemCode = ""
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtLocationCode_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLocationCode.Enter
        On Error GoTo ErrHandler
        Me.txtLocationCode.SelectionStart = 0
        Me.txtLocationCode.SelectionLength = Len(Me.txtLocationCode.Text)
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtLocationCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtLocationCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
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
        On Error GoTo ErrHandler
        If KeyCode = 112 Then
            If CmdLocCodeHelp.Enabled Then Call CmdLocCodeHelp_Click(CmdLocCodeHelp, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub SelectInvTypeSubTypeFromSaleConf(ByRef pstrInvType As String, ByRef pstrInvSubtype As String)
        On Error GoTo ErrHandler
        Dim strSaleConfSql As String
        Dim rsSaleConf As ClsResultSetDB
        strSaleConfSql = "Select Invoice_Type,Sub_Type from SaleConf where  UNIT_CODE = '" & gstrUNITID & "' and  Description='" & Trim(pstrInvType) & "' and (fin_start_date <= getdate() and fin_end_date >= getdate())"
        strSaleConfSql = strSaleConfSql & " and Sub_Type_Description='" & Trim(pstrInvSubtype) & "'"
        rsSaleConf = New ClsResultSetDB
        rsSaleConf.GetResult(strSaleConfSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsSaleConf.GetNoRows > 0 Then
            mstrInvoiceType = rsSaleConf.GetValue("Invoice_Type")
            mstrInvoiceSubType = rsSaleConf.GetValue("Sub_Type")
        Else
            mstrInvoiceType = ""
            mstrInvoiceSubType = ""
        End If
        rsSaleConf.ResultSetClose()
        rsSaleConf = Nothing
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtLocationCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtLocationCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                If Len(txtLocationCode.Text) > 0 Then
                    If CheckExistanceOfFieldData((txtLocationCode.Text), "Location_Code", "saleconf") Then
                        If Len(Trim(mstrLocation)) = 0 Then
                            mstrLocation = txtLocationCode.Text
                        End If
                        Call SelectDescriptionForField("Description", "Location_Code", "Location_Mst", lblLocCodeDes, (txtLocationCode.Text))
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
                    If CheckExistanceOfFieldData((txtLocationCode.Text), "Location_Code", "SaleConf") Then
                        Call SelectDescriptionForField("Description", "Location_Code", "Location_Mst", lblLocCodeDes, (txtLocationCode.Text))
                        If Len(Trim(mstrLocation)) = 0 Then
                            mstrLocation = txtLocationCode.Text
                        End If
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
        On Error GoTo ErrHandler
        If Me.CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            SpChEntry.MaxRows = 0 : mstrItemCode = "" : If txtRefNo.Enabled = True Then txtRefNo.Focus() Else txtCarrServices.Focus()
            txtAmendNo.Text = ""
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub txtRefNo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtRefNo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        If Len(txtCustCode.Text) > 0 Then
                            Call txtRefNo_Validating(txtRefNo, New System.ComponentModel.CancelEventArgs(False))
                        Else
                            If (CmbInvType.Text = "JOBWORK INVOICE") Then
                                txtCarrServices.Focus()
                            Else
                                txtCarrServices.Focus()
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
        On Error GoTo ErrHandler
        If KeyCode = 112 Then
            If cmdempHelpRefNo.Enabled Then Call cmdempHelpRefNo_Click(cmdempHelpRefNo, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtRefNo_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtRefNo.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        If Len(txtLocationCode.Text) > 0 Then
            If Len(txtRefNo.Text) > 0 Then
                If SelectDataFromCustOrd_Dtl((txtCustCode.Text), (CmbInvType.Text)) Then
                    If CmbInvType.Text <> "REJECTION" Then
                        Call displayDeatilsfromCustOrdHdr()
                        txtAmendNo.Focus()
                        If SpChEntry.MaxRows = 0 Then
                            Call addRowAtEnterKeyPress(1)
                            Call SetMaxLengthInSpread(0)
                            Call ChangeCellTypeStaticText()
                        End If
                    Else
                        If Len(Trim(txtRefNo.Text)) > 0 Then
                            If SpChEntry.MaxRows = 0 Then
                                Call addRowAtEnterKeyPress(1)
                                Call SetMaxLengthInSpread(0)
                                Call ChangeCellTypeStaticText()
                            End If
                        End If
                    End If
                Else
                    If CmbInvType.Text <> "REJECTION" Then
                        Call ConfirmWindow(10436, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                    Else
                        MsgBox("GRIN No Entered by you is inValid,Press F1 for Help.", MsgBoxStyle.Information, "eMPro")
                    End If
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
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        With Me.SpChEntry
                            If SpChEntry.MaxRows > 0 Then
                                If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                                    With SpChEntry
                                        .Row = 1 : .Col = 5 : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                                    End With
                                    System.Windows.Forms.Application.DoEvents()
                                Else
                                    .Row = 1 : .Col = 1 : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                                End If
                            Else
                                If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                                    Call addRowAtEnterKeyPress(1)
                                    Call ChangeCellTypeStaticText()
                                    With SpChEntry
                                        .Row = 1 : .Col = 1 : .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                    End With
                                    System.Windows.Forms.Application.DoEvents()
                                Else
                                    CmdGrpChEnt.Focus()
                                End If
                            End If
                        End With
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
    Private Sub txtSaleTaxType_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSaleTaxType.TextChanged
        On Error GoTo ErrHandler
        If Len(txtSaleTaxType.Text) = 0 Then
            lblSaltax_Per.Text = "0.00"
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub txtSaleTaxType_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtSaleTaxType.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        If Len(txtSaleTaxType.Text) > 0 Then
                            Call txtSaleTaxType_Validating(txtSaleTaxType, New System.ComponentModel.CancelEventArgs(False))
                        Else
                            If txtAddVAT.Enabled Then txtAddVAT.Focus()
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
        On Error GoTo ErrHandler
        '101188073
        If gblnGSTUnit Then Exit Sub
        '101188073
        If Len(txtSaleTaxType.Text) > 0 Then
            If CheckExistanceOfFieldData((txtSaleTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", " (Tx_TaxeID='CST' OR Tx_TaxeID='LST' OR Tx_TaxeID='VAT')") Then
                lblSaltax_Per.Text = CStr(GetTaxRate((txtSaleTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " (Tx_TaxeID='CST' OR Tx_TaxeID='LST' OR Tx_TaxeID='VAT')"))
                If txtAddVAT.Enabled Then txtAddVAT.Focus()
            Else
                Call ConfirmWindow(10248, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                Cancel = True
                txtSaleTaxType.Text = ""
                If txtSaleTaxType.Enabled Then txtSaleTaxType.Focus()
            End If
        End If
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtSchTime_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSchTime.TextChanged
        If Len(Trim(txtSchTime.Text)) = 0 Then
            txtSRVDI.Text = "" : txtSRVLoc.Text = "" : txtUsLoc.Text = ""
        End If
    End Sub
    Private Sub txtSchTime_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtSchTime.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        CmdGrpChEnt.Focus()
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
    Private Sub txtSDTType_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtSDTType.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F1 Then
            Call cmdSDTax_Help_Click(cmdSDTax_Help, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtSDTType_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtSDTType.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        '101188073
        If gblnGSTUnit Then Exit Sub
        '101188073
        Dim strSQL As String
        Dim rsTmp As New ClsResultSetDB
        If Len(Trim(txtSDTType.Text)) <> 0 Then
            strSQL = "Select TxRT_Percentage from Gen_TaxRate where Tx_TaxeId='SDT' and TxRt_Rate_No='" & Trim(txtSDTType.Text) & "' and UNIT_CODE = '" & gstrUNITID & "'"
            rsTmp.GetResult(strSQL)
            If Not rsTmp.EOFRecord Then
                
                lblSDTax_Per.Text = rsTmp.GetValue("TxRT_Percentage")
            Else
                MsgBox("Invalid Tax Type", MsgBoxStyle.Information, "eMPro")
                txtSDTType.Text = ""
                lblSDTax_Per.Text = "0.00"
            End If
            rsTmp.ResultSetClose()
        End If
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtSECESS_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSECESS.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        If Len(txtSECESS.Text) > 0 Then
                            Call txtSECESS_Validating(txtSECESS, New System.ComponentModel.CancelEventArgs(False))
                        Else
                            If txtRemarks.Enabled Then txtRemarks.Focus()
                        End If
                End Select
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        e.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            e.Handled = True
        End If
    End Sub
    Private Sub txtSECESS_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSECESS.TextChanged
        If Len(Trim(txtSECESS.Text)) = 0 Then lblSECESS_Per.Text = CStr(0)
    End Sub
    Private Sub txtSECESS_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtSECESS.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then cmdSECESSCode.PerformClick()
    End Sub
    Private Sub txtSRVDI_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSRVDI.TextChanged
        If Len(Trim(txtSRVDI.Text)) = 0 Then
            txtSchTime.Text = "" : txtSRVLoc.Text = "" : txtUsLoc.Text = ""
        End If
    End Sub
    Private Sub txtSRVDI_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtSRVDI.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        If txtSRVLoc.Enabled = True Then
                            txtSRVLoc.Focus()
                        Else
                            txtSRVLoc.Focus()
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
    Private Sub txtSRVDI_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtSRVDI.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.F1 Then
            Call cmdhelpSRVDI_Click(cmdhelpSRVDI, New System.EventArgs())
        End If
    End Sub
    Private Sub txtSRVLoc_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSRVLoc.TextChanged
        If Len(Trim(txtSRVLoc.Text)) = 0 Then
            txtSRVDI.Text = "" : txtSchTime.Text = "" : txtUsLoc.Text = ""
        End If
    End Sub
    Private Sub txtSRVLoc_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtSRVLoc.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        txtUsLoc.Focus()
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
    Private Sub txtSurchargeTaxType_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSurchargeTaxType.TextChanged
        On Error GoTo ErrHandler
        If Trim(txtSurchargeTaxType.Text) = "" Then
            lblSurcharge_Per.Text = "0.00"
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub txtSurchargeTaxType_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtSurchargeTaxType.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F1 Then
            Call cmdSurchargeTaxCode_Click(cmdSurchargeTaxCode, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtSurchargeTaxType_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtSurchargeTaxType.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        If ctlPerValue.Enabled = True Then
                            ctlPerValue.Focus()
                        Else
                            txtECESS.Focus()
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
    Private Sub txtSurchargeTaxType_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtSurchargeTaxType.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        '101188073
        If gblnGSTUnit Then Exit Sub
        '101188073
        If Trim(txtSurchargeTaxType.Text) <> "" Then
            If CheckExistanceOfFieldData((txtSurchargeTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", " Tx_TaxeID='SST'") Then
                lblSurcharge_Per.Text = CStr(GetTaxRate((txtSurchargeTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='SST'"))
                txtECESS.Focus()
            Else
                Call ConfirmWindow(10248, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                Cancel = True
                txtSurchargeTaxType.Text = ""
                If txtSurchargeTaxType.Enabled Then txtSurchargeTaxType.Focus()
            End If
        Else
            If txtECESS.Enabled Then txtECESS.Focus()
        End If
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtUsLoc_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUsLoc.TextChanged
        If Len(Trim(txtUsLoc.Text)) = 0 Then
            txtSRVDI.Text = "" : txtSRVLoc.Text = "" : txtSchTime.Text = ""
        End If
    End Sub
    Private Sub txtUsLoc_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtUsLoc.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        txtSchTime.Focus()
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
    Private Sub txtVehNo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtVehNo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        If ctlInsurance.Enabled Then
                            ctlInsurance.Focus()
                        Else
                            txtFreight.Focus()
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
    Private Function CheckExistanceOfFieldData(ByRef pstrFieldText As String, ByRef pstrColumnName As String, ByRef pstrTableName As String, Optional ByRef pstrCondition As String = "") As Boolean
        '****************************************************
        'Created By     -  Nisha
        'Description    -  To Check Validity Of Field Data Whethet it Exists In The
        '                  Database Or Not
        'Arguments      -  pstrFieldText - Field Text,pstrColumnName - Column Name
        '               -  pstrTableName - Table Name,pstrCondition - Optional Parameter For Condition
        '****************************************************
        On Error GoTo ErrHandler
        CheckExistanceOfFieldData = False
        Dim strTableSql As String 'Declared To Make Select Query
        Dim rsExistData As ClsResultSetDB
        If Len(Trim(pstrCondition)) > 0 Then
            strTableSql = "select " & Trim(pstrColumnName) & " from " & Trim(pstrTableName) & " where " & Trim(pstrColumnName) & "='" & Trim(pstrFieldText) & "' and UNIT_CODE = '" & gstrUNITID & "' and " & pstrCondition
        Else
            strTableSql = "select " & Trim(pstrColumnName) & " from " & Trim(pstrTableName) & " where " & Trim(pstrColumnName) & "='" & Trim(pstrFieldText) & "' and UNIT_CODE = '" & gstrUNITID & "'"
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
        Call gobjError.RaiseError(Err.Number, err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Function GetDataInViewMode() As Boolean
        '****************************************************
        'Created By     -  Nisha
        'Modified By    -  Nitin Sood
        'Modifcation    -  Amendment Number Displayed
        'Description    -  To display data in view mode from SalasChallan_Dtl,Sales_Dtl acc.to
        'LocationCode & Challan_No.
        '****************************************************
        On Error GoTo ErrHandler
        GetDataInViewMode = False
        Dim strGetData As String
        Dim rsGetData As ClsResultSetDB
        Dim rsCustMst As ClsResultSetDB
        Dim rsItemMst As ClsResultSetDB
        Dim strSalesChallanDtl As String
        Dim strRGPNOs As String
        Dim strCustMst As String
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        strSalesChallanDtl = "SELECT Transport_type,Vehicle_No,Account_Code,Cust_ref,Amendment_No,SalesTax_Type,"
        strSalesChallanDtl = strSalesChallanDtl & "Insurance,Invoice_Date,"
        strSalesChallanDtl = strSalesChallanDtl & "Invoice_Type,Sub_Category,Cust_Name,Carriage_Name,Frieght_Amount, "
        strSalesChallanDtl = strSalesChallanDtl & "Surcharge_salesTaxType,Amendment_No,ref_doc_no,Currency_Code,Exchange_Rate,PerValue, SDTAX_Type, SDTAX_Per,SDTAX_Amount,"
        strSalesChallanDtl = strSalesChallanDtl & "Remarks,FIFO_FLAG,SRVDINO,SRVLocation,USLOC,Schtime,ECESS_Type,"
        strSalesChallanDtl = strSalesChallanDtl & "AddVAT_Type,SECESS_Type,SECESS_Per,"
        strSalesChallanDtl = strSalesChallanDtl & "ECESS_Per,payment_terms From Saleschallan_dtl WHERE Location_Code ='"
        strSalesChallanDtl = strSalesChallanDtl & Trim(txtLocationCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "' and Doc_No = " & Val(txtChallanNo.Text)
        rsGetData = New ClsResultSetDB
        rsGetData.GetResult(strSalesChallanDtl, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsGetData.GetNoRows > 0 Then
            GetDataInViewMode = True
            txtVehNo.Text = rsGetData.GetValue("Vehicle_No")
            If Trim(rsGetData.GetValue("Transport_type")) = "R" Then
                CmbTransType.SelectedIndex = 0
            ElseIf Trim(rsGetData.GetValue("Transport_type")) = "L" Then
                CmbTransType.SelectedIndex = 1
            ElseIf Trim(rsGetData.GetValue("Transport_type")) = "S" Then
                CmbTransType.SelectedIndex = 2
            ElseIf Trim(rsGetData.GetValue("Transport_type")) = "A" Then
                CmbTransType.SelectedIndex = 3
            ElseIf Trim(rsGetData.GetValue("Transport_type")) = "H" Then
                CmbTransType.SelectedIndex = 4
            ElseIf Trim(rsGetData.GetValue("Transport_type")) = "C" Then
                CmbTransType.SelectedIndex = 5
            End If
            txtCustCode.Text = rsGetData.GetValue("Account_Code")
            txtRefNo.Text = rsGetData.GetValue("Cust_ref")
            txtAmendNo.Text = rsGetData.GetValue("Amendment_No")
            txtCarrServices.Text = rsGetData.GetValue("Carriage_Name")
            ctlInsurance.Text = rsGetData.GetValue("Insurance")
            txtFreight.Text = rsGetData.GetValue("Frieght_Amount")
            txtSaleTaxType.Text = rsGetData.GetValue("SalesTax_Type")
            Call txtSaleTaxType_Validating(txtSaleTaxType, New System.ComponentModel.CancelEventArgs(False))
            txtSurchargeTaxType.Text = rsGetData.GetValue("Surcharge_salesTaxType")
            Call txtSurchargeTaxType_Validating(txtSurchargeTaxType, New System.ComponentModel.CancelEventArgs(False))
            txtECESS.Text = rsGetData.GetValue("ECESS_Type")
            Call txtECESS_Validating(txtECESS, New System.ComponentModel.CancelEventArgs(False))
            strRGPNOs = rsGetData.GetValue("ref_doc_no")
            strRGPNOs = Replace(strRGPNOs, "§", ", ", 1)
            lblRGPDes.Text = strRGPNOs
            lblCustCodeDes.Text = rsGetData.GetValue("Cust_Name")
            mstrAmmendmentNo = rsGetData.GetValue("Amendment_No")
            lblDateDes.Text = VB6.Format(rsGetData.GetValue("Invoice_Date"), gstrDateFormat)
            mstrInvType = rsGetData.GetValue("Invoice_Type")
            mstrInvSubType = rsGetData.GetValue("Sub_Category")
            ctlPerValue.Text = rsGetData.GetValue("PerValue")
            mCurrencyCode = rsGetData.GetValue("Currency_code")
            txtSRVDI.Text = rsGetData.GetValue("SRVDINO")
            txtSRVLoc.Text = rsGetData.GetValue("SRVLocation")
            txtUsLoc.Text = rsGetData.GetValue("USLoc")
            txtSchTime.Text = rsGetData.GetValue("SchTime")
            mstrInvoiceType = rsGetData.GetValue("Invoice_Type")
            mstrInvoiceSubType = rsGetData.GetValue("Sub_Category")
            lblCurrencyDes.Text = rsGetData.GetValue("currency_code")
            If UCase(rsGetData.GetValue("Invoice_Type")) = "JOB" Then
                blnFIFO = rsGetData.GetValue("FIFO_FLAG")
            End If
            If UCase(rsGetData.GetValue("Invoice_Type")) <> "JOB" And UCase(rsGetData.GetValue("Invoice_Type")) <> "TRF" Then
                txtSDTType.Text = Trim(rsGetData.GetValue("SDTAX_Type"))
                lblSDTax_Per.Text = Trim(rsGetData.GetValue("SDTAX_Per"))
            End If
            If UCase(mstrInvType) = "EXP" Then
                lblCurrency.Visible = True : lblCurrencyDes.Visible = True
                lblCurrencyDes.Text = rsGetData.GetValue("Currency_code")
                lblExchangeRateLable.Visible = True : lblExchangeRateValue.Visible = True
            Else
                lblCurrencyDes.Text = ""
                lblCurrency.Visible = False : lblCurrencyDes.Visible = False
                lblExchangeRateLable.Visible = False : lblExchangeRateValue.Visible = False
            End If
            txtRemarks.Text = rsGetData.GetValue("Remarks")
            lblCreditTerm.Text = IIf(IsDBNull(rsGetData.GetValue("payment_terms")), "", rsGetData.GetValue("payment_terms"))
            If Len(Trim(lblCreditTerm.Text)) > 0 Then
                Call SelectDescriptionForField("crTrm_desc", "crtrm_termID", "Gen_CreditTrmMaster", lblCreditTermDesc, Trim(lblCreditTerm.Text))
            Else
                lblCreditTermDesc.Text = ""
            End If
            txtAddVAT.Text = rsGetData.GetValue("AddVAT_Type")
            Call txtAddVAT_Validating(txtAddVAT, New System.ComponentModel.CancelEventArgs(False))
            txtSECESS.Text = rsGetData.GetValue("SECESS_Type")
            Call txtSECESS_Validating(txtSECESS, New System.ComponentModel.CancelEventArgs(False))
        Else
            GetDataInViewMode = False
        End If
        rsGetData.ResultSetClose()
        '***To Display invoice Address of Customer
        If Len(Trim(txtCustCode.Text)) > 0 Then
            rsCustMst = New ClsResultSetDB
            strCustMst = "SELECT Bill_Address1 + ', '  +  Bill_Address2 + ', ' + Bill_City + ' - ' + Bill_Pin as  invoiceAddress from Customer_Mst where Customer_code ='" & txtCustCode.Text & "' and UNIT_CODE = '" & gstrUNITID & "' "
            rsCustMst.GetResult(strCustMst)
            If rsCustMst.GetNoRows > 0 Then
                lblAddressDes.Text = rsCustMst.GetValue("InvoiceAddress")
            End If
            rsCustMst.ResultSetClose()
        End If
        rsGetData = New ClsResultSetDB
        rsGetData.GetResult("Select * from Sales_dtl where location_code ='" & txtLocationCode.Text & "' and UNIT_CODE = '" & gstrUNITID & "' and Doc_no = " & txtChallanNo.Text)
        If rsGetData.GetNoRows > 0 Then
            intMaxLoop = rsGetData.GetNoRows
            rsGetData.MoveFirst()
            With SpChEntry
                SpChEntry.maxRows = 0
                For intLoopCounter = 1 To intMaxLoop
                    Call addRowAtEnterKeyPress(1)
                    Call .SetText(1, intLoopCounter, rsGetData.GetValue("Item_code"))
                    Call .SetText(2, intLoopCounter, rsGetData.GetValue("Cust_Item_code"))
                    Call .SetText(3, intLoopCounter, (rsGetData.GetValue("Rate") * CDbl(ctlPerValue.Text)))
                    Call .SetText(16, intLoopCounter, rsGetData.GetValue("Rate"))
                    Call .SetText(4, intLoopCounter, (Val(rsGetData.GetValue("Cust_mtrl")) * CDbl(ctlPerValue.Text)))
                    Call .SetText(17, intLoopCounter, rsGetData.GetValue("Cust_mtrl"))
                    Call .SetText(5, intLoopCounter, rsGetData.GetValue("Sales_Quantity"))
                    Call .SetText(6, intLoopCounter, rsGetData.GetValue("Packing_Type"))
                    Call .SetText(7, intLoopCounter, rsGetData.GetValue("Excise_Type"))
                    Call .SetText(8, intLoopCounter, rsGetData.GetValue("CVD_Type"))
                    Call .SetText(9, intLoopCounter, rsGetData.GetValue("SAD_Type"))
                    Call .SetText(10, intLoopCounter, (Val(rsGetData.GetValue("Others")) * CDbl(ctlPerValue.Text)))
                    Call .SetText(18, intLoopCounter, (Val(rsGetData.GetValue("Others")) * CDbl(ctlPerValue.Text)))
                    Call .SetText(11, intLoopCounter, rsGetData.GetValue("From_Box"))
                    Call .SetText(12, intLoopCounter, rsGetData.GetValue("To_Box"))
                    Call .SetText(13, intLoopCounter, (Val(rsGetData.GetValue("To_Box")) - Val(rsGetData.GetValue("From_Box"))) + 1)
                    Call .SetText(15, intLoopCounter, (Val(rsGetData.GetValue("tool_cost")) * CDbl(ctlPerValue.Text)))
                    Call .SetText(19, intLoopCounter, rsGetData.GetValue("tool_cost"))
                    Call .SetText(22, intLoopCounter, rsGetData.GetValue("BinQuantity"))
                    '101188073
                    If gblnGSTUnit Then
                        Call .SetText(IS_HSN_SAC, intLoopCounter, rsGetData.GetValue("ISHSNORSAC"))
                        Call .SetText(HSN_SAC_CODE, intLoopCounter, rsGetData.GetValue("HSNSACCODE"))
                        Call .SetText(CGST_TYPE, intLoopCounter, rsGetData.GetValue("CGSTTXRT_TYPE"))
                        Call .SetText(SGST_TYPE, intLoopCounter, rsGetData.GetValue("SGSTTXRT_TYPE"))
                        Call .SetText(IGST_TYPE, intLoopCounter, rsGetData.GetValue("IGSTTXRT_TYPE"))
                        Call .SetText(UTGST_TYPE, intLoopCounter, rsGetData.GetValue("UTGSTTXRT_TYPE"))
                        Call .SetText(COMP_CESS_TYPE, intLoopCounter, rsGetData.GetValue("COMPENSATION_CESS_TYPE"))
                    End If
                    '101188073
                    rsItemMst = New ClsResultSetDB
                    rsItemMst.GetResult("Select Tariff_Code from Item_Mst where Item_code = '" & rsGetData.GetValue("Item_code") & "' and UNIT_CODE = '" & gstrUnitId & "'")
                    Call .SetText(20, intLoopCounter, rsItemMst.GetValue("Tariff_Code"))
                    rsItemMst.ResultSetClose()
                    rsGetData.MoveNext()
                Next
            End With
        End If
        rsGetData.ResultSetClose()
        rsGetData = Nothing
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Function ValidatebeforeSave(ByRef pstrMode As String) As Boolean
        On Error GoTo ErrHandler
        Dim lstrControls As String
        Dim lNo As Integer
        Dim lctrFocus As System.Windows.Forms.Control
        Dim blnCheckAddVAT As Boolean
        Dim strSQL As String = ""

        ValidatebeforeSave = True
        lNo = 1
        lstrControls = ResolveResString(10059)
        blnCheckAddVAT = Find_Value("SELECT ISNULL(CHECKADDITONALVAT,0) AS CHECKADDITONALVAT FROM SALES_PARAMETER WHERE UNIT_CODE = '" & gstrUNITID & "'")
        BlankRowCheckandDelete()
        Select Case UCase(Trim(pstrMode))
            Case "ADD"
                '10736222
                strSQL = "DELETE FROM TMP_CT2_INVOICE_KNOCKOFF where UNIT_CODE='" + gstrUNITID + "' and IP_ADDRESS='" & gstrIpaddressWinSck & "'"
                SqlConnectionclass.ExecuteNonQuery(strSQL)
                '10736222

                If (Len(Me.txtLocationCode.Text) = 0) Then
                    lstrControls = lstrControls & vbCrLf & lNo & ". Location Code."
                    lNo = lNo + 1
                    If lctrFocus Is Nothing Then
                        lctrFocus = Me.txtLocationCode
                    End If
                    ValidatebeforeSave = False
                End If
                '101188073
                If Not gblnGSTUnit Then
                    If Val(lblECESS_Per.Text) > 0 Then
                        If Len(Trim(txtSECESS.Text)) = 0 Then
                            lstrControls = lstrControls & vbCrLf & lNo & ". Secondary ECESS"
                            lNo = lNo + 1
                            If lctrFocus Is Nothing Then
                                lctrFocus = Me.txtSECESS
                            End If
                            ValidatebeforeSave = False
                        End If
                    End If
                    If blnCheckAddVAT = True Then
                        If Me.txtSaleTaxType.Text.Length() > 0 Then
                            If Me.txtAddVAT.Text.Length = 0 Then
                                lstrControls = lstrControls & vbCrLf & lNo & ". Additional VAT"
                                lNo = lNo + 1
                                If lctrFocus Is Nothing Then
                                    lctrFocus = Me.txtAddVAT
                                End If
                                ValidatebeforeSave = False
                            End If
                        End If
                    End If
                End If
                '101188073
                If (Len(Me.txtCustCode.Text) = 0) Then
                    lstrControls = lstrControls & vbCrLf & lNo & ". Customer Code."
                    lNo = lNo + 1
                    If lctrFocus Is Nothing Then
                        lctrFocus = Me.txtCustCode
                    End If
                    ValidatebeforeSave = False
                End If
                'Check If Date is Appropriate
                If Not DateIsAppropriate() Then
                    lstrControls = lstrControls & vbCrLf & lNo & ". Date specified either Falls Before the LAST Invoice Date or is Greater than Todays Date."
                    lNo = lNo + 1
                    If lctrFocus Is Nothing Then
                        lctrFocus = Me.txtCustCode
                    End If
                    ValidatebeforeSave = False
                End If
                'ISSUE ID :10465802
                If (UCase(Trim(CmbInvType.Text)) = "NORMAL INVOICE") Or (UCase(Trim(CmbInvType.Text)) = "JOBWORK INVOICE") Or (UCase(Trim(CmbInvType.Text)) = "EXPORT INVOICE") Or (UCase(Trim(CmbInvType.Text)) = "TRANSFER INVOICE") Then
                    If (Trim(CmbInvSubType.Text) <> "SCRAP") Then
                        If (Len(Me.txtRefNo.Text) = 0) Then
                            lstrControls = lstrControls & vbCrLf & lNo & ". Reference No.."
                            lNo = lNo + 1
                            If lctrFocus Is Nothing Then
                                lctrFocus = Me.cmdempHelpRefNo
                            End If
                            ValidatebeforeSave = False
                        End If
                    End If
                    If blnFIFO = False Then
                        If (Trim(CmbInvType.Text) = "JOBWORK INVOICE") Then
                            If (Len(mstrRGP) = 0) Then
                                lstrControls = lstrControls & vbCrLf & lNo & ". RGP No.."
                                lNo = lNo + 1
                                If lctrFocus Is Nothing Then
                                    lctrFocus = Me.cmdempHelpRefNo
                                End If
                                ValidatebeforeSave = False
                            End If
                        End If
                    End If
                End If
                If SpChEntry.MaxRows = 0 Then
                    lstrControls = lstrControls & vbCrLf & lNo & ". Add Atleast One Item."
                    lNo = lNo + 1
                    If lctrFocus Is Nothing Then
                        lctrFocus = SpChEntry
                    End If
                    ValidatebeforeSave = False
                End If
                If (Len(Me.txtFreight.Text) = 0) Then
                    txtFreight.Text = "0.00"
                End If
                If (Len(Me.ctlInsurance.Text) = 0) Then
                    ctlInsurance.Text = "0.00"
                End If
                If (Len(lblCurrencyDes.Text) = 0) Then
                    lblCurrencyDes.Text = gstrCURRENCYCODE
                End If
                If gblnGSTUnit = True And TxtTCSTaxcode.Enabled = True Then
                    If TxtTCSTaxcode.Text.Trim.ToString.Length = 0 Then
                        lstrControls = lstrControls & vbCrLf & lNo & ". TCS Tax."
                        lNo = lNo + 1
                        If lctrFocus Is Nothing Then
                            lctrFocus = Me.TxtTCSTaxcode
                        End If
                        ValidatebeforeSave = False
                    End If
                End If

            Case "EDIT"
                '10736222
                strSQL = "DELETE FROM TMP_CT2_INVOICE_KNOCKOFF where UNIT_CODE='" + gstrUnitId + "' and IP_ADDRESS='" & gstrIpaddressWinSck & "'"
                SqlConnectionclass.ExecuteNonQuery(strSQL)
                '10736222
                '101188073
                If Not gblnGSTUnit Then
                    If Val(lblECESS_Per.Text) > 0 Then
                        If Len(Trim(txtSECESS.Text)) = 0 Then
                            lstrControls = lstrControls & vbCrLf & lNo & ". Secondary ECESS"
                            lNo = lNo + 1
                            If lctrFocus Is Nothing Then
                                lctrFocus = Me.txtSECESS
                            End If
                            ValidatebeforeSave = False
                        End If
                    End If
                    If blnCheckAddVAT = True Then
                        If Me.txtSaleTaxType.Text.Length() > 0 Then
                            If Me.txtAddVAT.Text.Length = 0 Then
                                lstrControls = lstrControls & vbCrLf & lNo & ". Additional VAT"
                                lNo = lNo + 1
                                If lctrFocus Is Nothing Then
                                    lctrFocus = Me.txtAddVAT
                                End If
                                ValidatebeforeSave = False
                            End If
                        End If
                    End If
                End If
                '101188073
                If (Len(Me.txtFreight.Text) = 0) Then
                    txtFreight.Text = "0.00"
                End If
                If gblnGSTUnit = True And TxtTCSTaxcode.Enabled = True Then
                    If TxtTCSTaxcode.Text.Trim.ToString.Length = 0 Then
                        lstrControls = lstrControls & vbCrLf & lNo & ". TCS Tax."
                        lNo = lNo + 1
                        If lctrFocus Is Nothing Then
                            lctrFocus = Me.TxtTCSTaxcode
                        End If
                        ValidatebeforeSave = False
                    End If
                End If

        End Select
        If Not ValidatebeforeSave Then
            MsgBox(lstrControls, MsgBoxStyle.Information, ResolveResString(10059))
            lctrFocus.Focus()
        End If
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        gblnCancelUnload = True 'Initialize the Variable
        Exit Function
    End Function
    Private Sub ChangeCellTypeStaticText()
        On Error GoTo ErrHandler
        Dim intRow As Short
        Dim intcol As Short
        Dim rsChallanEntry As ClsResultSetDB
        Dim strInvoiceType As String
        Dim strInvoiceSubType As String
        Dim VarDelete As Object
        With Me.SpChEntry
            Select Case Me.CmdGrpChEnt.mode
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                    'ISSUE ID :10465802
                    If (UCase(Trim(CmbInvType.Text)) = "NORMAL INVOICE") Or (UCase(Trim(CmbInvType.Text)) = "EXPORT INVOICE") Or (UCase(Trim(CmbInvType.Text)) = "TRANSFER INVOICE") Then
                        If UCase(Trim(CmbInvSubType.Text)) <> "SCRAP" Then
                            For intRow = 1 To .MaxRows
                                .Row = intRow
                                For intcol = 1 To .MaxCols
                                    .Col = intcol
                                    If intcol = 5 Or intcol = 22 Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                    ElseIf intcol = 15 Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                    ElseIf intcol = 11 Or intcol = 12 Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeInteger
                                    ElseIf intcol = 8 Or intcol = 9 Or intcol = 14 Or intcol = 2 Or intcol = 1 Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                                    Else
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                                        System.Windows.Forms.Application.DoEvents()
                                    End If
                                Next intcol
                            Next intRow
                        Else
                            For intRow = 1 To .MaxRows
                                .Row = intRow
                                For intcol = 1 To .MaxCols
                                    .Col = intcol
                                    If intcol = 5 Or intcol = 15 Or intcol = 22 Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                    ElseIf intcol = 11 Or intcol = 12 Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeInteger
                                    ElseIf intcol = 3 Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                    ElseIf intcol = 14 Or intcol = 8 Or intcol = 9 Or intcol = 7 Or intcol = 1 Or intcol = 2 Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                                    Else
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                                    End If
                                Next intcol
                            Next intRow
                        End If
                    Else
                        For intRow = 1 To .MaxRows
                            .Row = intRow
                            For intcol = 1 To .MaxCols
                                .Col = intcol
                                If intcol = 5 Or intcol = 15 Or intcol = 22 Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                ElseIf intcol = 11 Or intcol = 12 Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeInteger
                                ElseIf intcol = 3 Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                ElseIf intcol = 14 Or intcol = 8 Or intcol = 9 Or intcol = 7 Or intcol = 1 Or intcol = 2 Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                                Else
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                                End If
                            Next intcol
                        Next intRow
                    End If
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                    rsChallanEntry = New ClsResultSetDB
                    rsChallanEntry.GetResult("Select a.Description,a.Sub_Type_Description from SaleConf a,SalesChallan_Dtl b where a.UNIT_CODE = b.UNIT_CODE AND a.UNIT_CODE = '" & gstrUNITID & "' AND Doc_No = " & txtChallanNo.Text & " and a.Invoice_Type = b.Invoice_type and a.Sub_type = b.Sub_Category and a.Location_code = b.Location_code and (fin_start_date <= getdate() and fin_end_date >= getdate())")
                    strInvoiceType = UCase(rsChallanEntry.GetValue("Description"))
                    strInvoiceSubType = UCase(rsChallanEntry.GetValue("sub_type_Description"))
                    rsChallanEntry.ResultSetClose()
                    If (UCase(strInvoiceType) = "NORMAL INVOICE") Or (UCase(strInvoiceType) = "EXPORT INVOICE") Or (UCase(strInvoiceType) = "TRANSFER INVOICE") Then
                        If (UCase(strInvoiceSubType) <> "SCRAP") Then
                            For intRow = 1 To .MaxRows
                                VarDelete = Nothing
                                Call SpChEntry.GetText(14, intRow, VarDelete)
                                .Row = intRow
                                For intcol = 1 To .MaxCols
                                    .Col = intcol
                                    If intcol = 5 Or intcol = 15 Or intcol = 22 Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                    ElseIf intcol = 11 Or intcol = 12 Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeInteger
                                    ElseIf intcol = 14 Or intcol = 8 Or intcol = 9 Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                                    ElseIf intcol = 1 Or intcol = 2 Then
                                        If UCase(Trim(VarDelete)) = "A" Then
                                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                                        Else
                                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                                        End If
                                    Else
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                                    End If
                                Next intcol
                            Next intRow
                        Else
                            For intRow = 1 To .MaxRows
                                VarDelete = Nothing
                                Call SpChEntry.GetText(14, intRow, VarDelete)
                                .Row = intRow
                                For intcol = 1 To .MaxCols
                                    .Col = intcol
                                    If intcol = 5 Or intcol = 15 Or intcol = 22 Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                    ElseIf intcol = 11 Or intcol = 12 Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeInteger
                                    ElseIf intcol = 3 Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                    ElseIf intcol = 14 Or intcol = 8 Or intcol = 9 Or intcol = 7 Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                                    ElseIf intcol = 1 Or intcol = 2 Then
                                        If UCase(Trim(VarDelete)) = "A" Then
                                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                                        Else
                                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                                        End If
                                    Else
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                                    End If
                                Next intcol
                            Next intRow
                        End If
                    Else
                        For intRow = 1 To .MaxRows
                            VarDelete = Nothing
                            Call SpChEntry.GetText(14, intRow, VarDelete)
                            .Row = intRow
                            For intcol = 1 To .MaxCols
                                .Col = intcol
                                If intcol = 5 Or intcol = 15 Or intcol = 22 Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                ElseIf intcol = 11 Or intcol = 12 Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeInteger
                                ElseIf intcol = 3 Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                ElseIf intcol = 14 Or intcol = 8 Or intcol = 9 Or intcol = 7 Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                                ElseIf intcol = 1 Or intcol = 2 Then
                                    If UCase(Trim(VarDelete)) = "A" Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                                    Else
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                                    End If
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
        Call gobjError.RaiseError(Err.Number, err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Function QuantityCheck() As Boolean
        On Error GoTo ErrHandler
        QuantityCheck = False
        Dim strScheduleSql As String
        Dim strScheduleSql1 As String
        Dim rsMktSchedule As ClsResultSetDB
        Dim rsSaleConf As ClsResultSetDB
        Dim rsSalesParameter As ClsResultSetDB
        Dim rsbom As ClsResultSetDB
        Dim strQuantity As String
        Dim ldblNetDispatchQty As Double
        Dim intRwCount As Short 'To Count No. Of Rows
        Dim intLoopCount As Short
        Dim varItemQty As Object 'To Get Quantity Acc. To Drawing No and Item Code
        Dim varItemCode As Object
        Dim varDrgNo As Object
        Dim VarDelete As Object
        Dim varToolCost As Object
        Dim strItembal As String
        Dim PresQty As Object
        Dim intcol As Short
        Dim strInvoiceType As String
        Dim strInvoiceSubType As String
        Dim irowcount As Short
        Dim intRwCount1 As Short
        Dim intFromBox As Double
        Dim varItemQty1 As Object
        Dim blnDSTracking As Boolean
        Dim strToolCode As String
        'To Check Proper value in Quantity,From/To Box
        mstrUpdDispatchSql = ""
        For intRwCount = 1 To SpChEntry.MaxRows
            VarDelete = Nothing
            Call SpChEntry.GetText(14, intRwCount, VarDelete)
            '****Delete Flag Check
            If UCase(VarDelete) <> "D" Then
                For intcol = 1 To SpChEntry.MaxCols
                    SpChEntry.Col = intcol
                    If (SpChEntry.Col = 5) Or (SpChEntry.Col = 22) Or (SpChEntry.Col = 3) Or (SpChEntry.Col = 12) Or (SpChEntry.Col = 11) Then ''Column Changed By Tapan
                        SpChEntry.Row = intRwCount
                        If (Val(Trim(SpChEntry.Text)) = 0) Then
                            QuantityCheck = True
                            Call ConfirmWindow(10419, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                            SpChEntry.Row = intRwCount : SpChEntry.Col = intcol : SpChEntry.Action = 0 : SpChEntry.Focus()
                            Exit Function
                        End If
                        If (SpChEntry.Col = 12) Then
                            SpChEntry.Row = intRwCount : SpChEntry.Col = 11 : intFromBox = Val(Trim(SpChEntry.Text))
                            SpChEntry.Row = intRwCount : SpChEntry.Col = 12
                            'To Check Valid Quantity of From/To Box
                            If Val(Trim(SpChEntry.Text)) < intFromBox Then
                                QuantityCheck = True
                                Call ConfirmWindow(10235, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                                SpChEntry.Row = intRwCount : SpChEntry.Col = 12 : SpChEntry.Action = 0 : SpChEntry.Focus()
                                Exit Function
                            End If
                        End If
                    End If
                Next intcol
            End If
        Next intRwCount
        'Validation for Schedule Quantity Start Here
        If ValidateScheduleQuantity() = False Then QuantityCheck = True : Exit Function
        'Validation for Schedule Quantity End Here
        'To check Current Balance from Itembal_Mst
        'If Quantity Entered Is Greater Then Cur_Bal In The ItemBal_Mst
        'Then Restrict User To Change The Entered Quantity
        'To Get Item Code From Spread
        Dim strItCode As String 'To Make Item Code String
        For intRwCount = 1 To Me.SpChEntry.MaxRows
            VarDelete = Nothing
            Call Me.SpChEntry.GetText(14, intRwCount, VarDelete)
            
            If UCase(VarDelete) <> "D" Then
                varItemCode = Nothing
                Call Me.SpChEntry.GetText(1, intRwCount, varItemCode)
                strItCode = strItCode & "'" & Trim(varItemCode) & "',"
            End If
        Next intRwCount
        strItCode = Mid(strItCode, 1, Len(strItCode) - 1)
        Select Case Me.CmdGrpChEnt.mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT, UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                rsSaleConf = New ClsResultSetDB
                rsSaleConf.GetResult(" Select Invoice_type,Sub_Category from SalesChallan_Dtl Where Doc_No=" & txtChallanNo.Text & " and UNIT_CODE = '" & gstrUNITID & "'")
                mstrInvoiceType = rsSaleConf.GetValue("Invoice_Type")
                mstrInvSubType = rsSaleConf.GetValue("Sub_Category")
                rsSaleConf.ResultSetClose()
                rsSaleConf = New ClsResultSetDB
                rsSaleConf.GetResult("select Stock_Location From saleconf where invoice_type ='" & Trim(mstrInvoiceType) & "' and sub_type ='" & Trim(mstrInvSubType) & "' and UNIT_CODE = '" & gstrUNITID & "'  AND Location_Code ='" & Trim(txtLocationCode.Text) & "' and (fin_start_date <= getdate() and fin_end_date >= getdate())")
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                rsSaleConf = New ClsResultSetDB
                rsSaleConf.GetResult("select Stock_Location From saleconf where Description ='" & Trim(CmbInvType.Text) & "' and UNIT_CODE = '" & gstrUNITID & "' and  sub_type_Description ='" & Trim(CmbInvSubType.Text) & "' AND Location_Code ='" & Trim(txtLocationCode.Text) & "' and (fin_start_date <= getdate() and fin_end_date >= getdate())")
        End Select
        If Len(Trim(rsSaleConf.GetValue("Stock_Location"))) = 0 Then
            MsgBox("Please Define Stock Location in Sales Conf First", MsgBoxStyle.OKOnly, "eMPro")
            QuantityCheck = True
            Exit Function
        End If
        Dim varItemCodeinVeiw As Object
        For intRwCount = 1 To Me.SpChEntry.MaxRows
            varItemCodeinVeiw = Nothing
            Call SpChEntry.GetText(1, intRwCount, varItemCodeinVeiw)
            VarDelete = Nothing
            Call SpChEntry.GetText(14, intRwCount, VarDelete)
            If UCase(VarDelete) <> "D" Then
                strItembal = "Select Cur_Bal From ItemBal_Mst where Location_Code ='" & Trim(rsSaleConf.GetValue("Stock_Location")) & "' and item_Code ='" & varItemCodeinVeiw & "' and UNIT_CODE = '" & gstrUNITID & "'"
                rsMktSchedule = New ClsResultSetDB
                rsMktSchedule.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                strQuantity = CStr(Val(rsMktSchedule.GetValue("Cur_Bal")))
                rsMktSchedule.ResultSetClose()
                varItemQty = Nothing
                Call Me.SpChEntry.GetText(5, intRwCount, varItemQty)
                If Val(varItemQty) > Val(strQuantity) Then
                    QuantityCheck = True
                    If CDbl(strQuantity) = 0 Then
                        MsgBox("No Balance Available for Item (" & varItemCodeinVeiw & ")", MsgBoxStyle.OkOnly, "eMPro")
                    Else
                        MsgBox("Quantity should not be Greater then Current Balance  at location  " & rsSaleConf.GetValue("Stock_Location") & " " & strQuantity, MsgBoxStyle.OkOnly, "eMPro")
                    End If
                    With Me.SpChEntry
                        .Row = intRwCount : .Col = 5 : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                    End With
                    Exit Function
                Else
                    QuantityCheck = False
                End If
            End If
        Next intRwCount
        rsSaleConf.ResultSetClose()
        rsSaleConf = Nothing
        'To check if tool Amortization Check is required
        'then in Invoice if Tool Amortization is there or not
        'to check if this qty is available in Tool Amortization details
        rsSalesParameter = New ClsResultSetDB
        rsSalesParameter.GetResult("Select CheckToolAmortisation from Sales_Parameter WHERE UNIT_CODE = '" & gstrUNITID & "'")
        If rsSalesParameter.GetNoRows > 0 Then
            rsSalesParameter.MoveFirst()
            If Len(Trim(rsSalesParameter.GetValue("CheckToolAmortisation"))) = 0 Then
                MsgBox("First define Check Tool Amortisation in Sales Parameter", MsgBoxStyle.Information, "eMPro")
                QuantityCheck = True
                rsSalesParameter.ResultSetClose()
                Exit Function
            End If
            If rsSalesParameter.GetValue("CheckToolAmortisation") = True Then
                For intRwCount = 1 To Me.SpChEntry.MaxRows
                    varItemCodeinVeiw = Nothing
                    Call SpChEntry.GetText(1, intRwCount, varItemCodeinVeiw)
                    varDrgNo = Nothing
                    Call SpChEntry.GetText(2, intRwCount, varDrgNo)
                    varToolCost = Nothing
                    Call SpChEntry.GetText(15, intRwCount, varToolCost)
                    VarDelete = Nothing
                    Call SpChEntry.GetText(14, intRwCount, VarDelete)
                    If UCase(VarDelete) <> "D" Then
                        With mP_Connection
                            .Execute("DELETE FROM  tmpBOM WHERE  UNIT_CODE = '" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            .Execute("BOMExplosion '" & Trim(varItemCodeinVeiw) & "','" & Trim(varItemCodeinVeiw) & "',1,0,0,0,'" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        End With
                        rsbom = New ClsResultSetDB
                        rsbom.GetResult("select * from tmpBOM WHERE UNIT_CODE = '" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        If rsbom.GetNoRows > 0 Then
                            irowcount = rsbom.GetNoRows
                            rsbom.MoveFirst()
                            For intRwCount1 = 1 To irowcount
                                strItembal = "select BalanceQty = isnull(a.proj_qty,0) - isnull(a.ClosingValueSMIEL,0),a.Tool_c from Amor_dtl a, tool_mst b "
                                strItembal = strItembal & " where a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND account_code = '" & Trim(txtCustCode.Text) & "'"
                                strItembal = strItembal & " and Item_code = '" & rsbom.GetValue("item_code") & "' and a.Tool_c = b.Tool_c and a.Item_code = b.Product_No order by a.tool_c"
                                rsMktSchedule = New ClsResultSetDB
                                rsMktSchedule.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                                If rsMktSchedule.GetNoRows > 0 Then
                                    rsMktSchedule.MoveFirst()
                                    strQuantity = CStr(Val(rsMktSchedule.GetValue("BalanceQty")))
                                    strToolCode = rsMktSchedule.GetValue("Tool_c")
                                    rsMktSchedule.ResultSetClose()
                                    varItemQty = Nothing
                                    Call Me.SpChEntry.GetText(5, intRwCount, varItemQty)
                                    varItemQty1 = (varItemQty * Val(rsbom.GetValue("grossweight")))
                                    strItembal = "select BalanceQty = sum(isnull(UsedProjQty,0)) from Amor_dtl "
                                    strItembal = strItembal & " where "
                                    strItembal = strItembal & " Item_code = '" & rsbom.GetValue("item_code") & "' and tool_c = '" & strToolCode & "'"
                                    strItembal = strItembal & " and account_code = '" & Trim(txtCustCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "'"
                                    rsMktSchedule = New ClsResultSetDB
                                    rsMktSchedule.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                                    rsMktSchedule.MoveFirst()
                                    strQuantity = CStr(CDbl(strQuantity) - Val(rsMktSchedule.GetValue("BalanceQty")))
                                    rsMktSchedule.ResultSetClose()
                                    If Val(varItemQty1) > Val(strQuantity) Then
                                        QuantityCheck = True
                                        If CDbl(strQuantity) = 0 Then
                                            MsgBox("No Balance Available for Item (" & rsbom.GetValue("item_code") & ") and customer Part Code (" & varDrgNo & ") For Amortisation Calculations. ", MsgBoxStyle.OkOnly, "eMPro")
                                        Else
                                            MsgBox("Quantity should not be Greater then available Balance Quantity for Amortisarion of this Item (" & rsbom.GetValue("item_code") & ")" & strQuantity, MsgBoxStyle.OkOnly, "eMPro")
                                            With Me.SpChEntry
                                                .Row = intRwCount : .Col = 5 : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                                            End With
                                            rsbom.ResultSetClose()
                                            Exit Function
                                        End If
                                        Exit Function
                                    Else
                                        QuantityCheck = False
                                    End If
                                Else
                                    rsMktSchedule.ResultSetClose()
                                End If
                                rsbom.MoveNext()
                            Next
                        Else
                            rsbom.ResultSetClose()
                        End If
                        'Here I Check The Finished Item
                        strItembal = "select BalanceQty = isnull(a.proj_qty,0) - isnull(a.ClosingValueSMIEL,0),a.Tool_c from Amor_dtl a,Tool_Mst b"
                        strItembal = strItembal & " where account_code = '" & Trim(txtCustCode.Text) & "' AND a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "'"
                        strItembal = strItembal & " and Item_code = '" & varItemCodeinVeiw & "' and a.Tool_c = b.tool_c and a.item_code = b.Product_No order by a.tool_c"
                        rsMktSchedule = New ClsResultSetDB
                        rsMktSchedule.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        If rsMktSchedule.GetNoRows > 0 Then
                            rsMktSchedule.MoveFirst()
                            strQuantity = CStr(Val(rsMktSchedule.GetValue("BalanceQty")))
                            strToolCode = rsMktSchedule.GetValue("Tool_c")
                            rsMktSchedule.ResultSetClose()
                            varItemQty = Nothing
                            Call Me.SpChEntry.GetText(5, intRwCount, varItemQty)
                            strItembal = "select BalanceQty = sum(isnull(UsedProjQty,0)) from Amor_dtl "
                            strItembal = strItembal & " where "
                            strItembal = strItembal & " Item_code = '" & rsbom.GetValue("item_code") & "' and UNIT_CODE = '" & gstrUNITID & "' and tool_c = '" & strToolCode & "'"
                            strItembal = strItembal & " and account_code = '" & Trim(txtCustCode.Text) & "'"
                            rsMktSchedule = New ClsResultSetDB
                            rsMktSchedule.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                            rsMktSchedule.MoveFirst()
                            strQuantity = CStr(CDbl(strQuantity) - Val(rsMktSchedule.GetValue("BalanceQty")))
                            rsMktSchedule.ResultSetClose()
                            If Val(varItemQty) > Val(strQuantity) Then
                                QuantityCheck = True
                                If CDbl(strQuantity) = 0 Then
                                    MsgBox("No Balance Available for Item (" & varItemCodeinVeiw & ") and customer Part Code (" & varDrgNo & ") For Amortisation Calculations. ", MsgBoxStyle.OkOnly, "eMPro")
                                Else
                                    MsgBox("Quantity should not be Greater then available Balance Quantity for Amortisarion " & strQuantity, MsgBoxStyle.OkOnly, "eMPro")
                                End If
                                With Me.SpChEntry
                                    .Row = intRwCount : .Col = 5 : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                                End With
                                Exit Function
                            Else
                                QuantityCheck = False
                            End If
                        Else
                            rsMktSchedule.ResultSetClose()
                            strItembal = "select BalanceQty = isnull(a.proj_qty,0) - isnull(a.ClosingValueSMIEL,0) from Amor_dtl a"
                            strItembal = strItembal & " where account_code = '" & Trim(txtCustCode.Text) & "' and a.UNIT_CODE = '" & gstrUNITID & "'"
                            strItembal = strItembal & " and Item_code = '" & varItemCodeinVeiw & "'"
                            rsMktSchedule = New ClsResultSetDB
                            rsMktSchedule.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                            If rsMktSchedule.GetNoRows > 0 Then
                                rsMktSchedule.MoveFirst()
                                strQuantity = CStr(Val(rsMktSchedule.GetValue("BalanceQty")))
                                rsMktSchedule.ResultSetClose()
                                varItemQty = Nothing
                                Call Me.SpChEntry.GetText(5, intRwCount, varItemQty)
                                strItembal = "select BalanceQty = sum(isnull(UsedProjQty,0)) from Amor_dtl "
                                strItembal = strItembal & " where "
                                strItembal = strItembal & " Item_code = '" & varItemCodeinVeiw & "' and UNIT_CODE = '" & gstrUNITID & "'"
                                strItembal = strItembal & " and account_code = '" & Trim(txtCustCode.Text) & "'"
                                rsMktSchedule = New ClsResultSetDB
                                rsMktSchedule.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                                rsMktSchedule.MoveFirst()
                                strQuantity = CStr(CDbl(strQuantity) - Val(rsMktSchedule.GetValue("BalanceQty")))
                                rsMktSchedule.ResultSetClose()
                                If Val(varItemQty) > Val(strQuantity) Then
                                    QuantityCheck = True
                                    If CDbl(strQuantity) = 0 Then
                                        MsgBox("No Balance Available for Item (" & varItemCodeinVeiw & ") and customer Part Code (" & varDrgNo & ") For Amortisation Calculations. ", MsgBoxStyle.OkOnly, "eMPro")
                                    Else
                                        MsgBox("Quantity should not be Greater then available Balance Quantity for Amortisarion " & strQuantity, MsgBoxStyle.OkOnly, "eMPro")
                                    End If
                                    With Me.SpChEntry
                                        .Row = intRwCount : .Col = 5 : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                                    End With
                                    Exit Function
                                Else
                                    QuantityCheck = False
                                End If
                            Else
                                rsMktSchedule.ResultSetClose()
                            End If
                        End If
                    End If
                Next intRwCount
            End If
        Else
            rsSalesParameter.ResultSetClose()
        End If
        'to check quantity available in CustAnnex_dtl
        'in case of JobWork Order
        If UCase(Trim(mstrInvoiceType)) = "JOB" And GetBOMCheckFlagValue("BomCheck_Flag") Then
            If BomCheck() = False Then
                QuantityCheck = True
                Exit Function
            Else
                QuantityCheck = False
            End If
        End If
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Sub RefreshForm(ByRef pstrType As String)
        On Error GoTo ErrHandler
        Select Case UCase(pstrType)
            Case "LOCATION"
                txtLocationCode.Text = "" : lblLocCodeDes.Text = "" : lblRGPDes.Text = ""
                txtChallanNo.Text = "" : txtCustCode.Text = "" : lblCustCodeDes.Text = "" : lblAddressDes.Text = ""
                txtCarrServices.Text = "" : txtVehNo.Text = ""
                txtFreight.Text = "" : txtSaleTaxType.Text = "" : lblSaltax_Per.Text = "0.00"
                txtSurchargeTaxType.Text = "" : lblSurcharge_Per.Text = "0.00"
                ctlInsurance.Text = "" : lblCurrencyDes.Text = "" : txtRefNo.Text = ""
                txtECESS.Text = "" : lblECESS_Per.Text = "0.00"
                CmbInvType.SelectedIndex = -1 : CmbInvSubType.SelectedIndex = -1
                ctlPerValue.Text = 1
                Me.CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                Me.CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                mCurrencyCode = ""
            Case "CHALLAN"
                txtChallanNo.Text = "" : txtCustCode.Text = "" : lblCustCodeDes.Text = "" : lblAddressDes.Text = ""
                txtCarrServices.Text = "" : txtVehNo.Text = ""
                txtFreight.Text = "" : txtSaleTaxType.Text = "" : lblSaltax_Per.Text = "0.00"
                txtSurchargeTaxType.Text = "" : lblSurcharge_Per.Text = "0.00"
                ctlInsurance.Text = "" : lblRGPDes.Text = ""
                CmbInvType.SelectedIndex = -1 : CmbInvSubType.SelectedIndex = -1 : lblCurrencyDes.Text = "" : txtRefNo.Text = ""
                ctlPerValue.Text = 1
                Me.CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                Me.CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                txtECESS.Text = "" : lblECESS_Per.Text = "0.00"
                mCurrencyCode = ""
        End Select
        With Me.SpChEntry
            .maxRows = 1
            .Row = 1 : .Row2 = 1 : .Col = 1 : .Col2 = .MaxCols : .BlockMode = True : .Text = "" : .BlockMode = False
        End With
        lblCreditTerm.Text = ""
        lblCreditTermDesc.Text = ""
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub AddTransPortTypeToCombo()
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
        Call gobjError.RaiseError(Err.Number, err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub SelectChallanNoFromSalesChallanDtl()
        On Error GoTo ErrHandler
        Dim strChallanNo As String
        Dim rsChallanNo As New ClsResultSetDB
        strChallanNo = "SELECT (CURRENT_NO + 1)CURRENT_NO FROM DOCUMENTTYPE_MST WHERE DOC_TYPE = 9999 and UNIT_CODE = '" & gstrUNITID & "' AND GETDATE() BETWEEN FIN_START_DATE AND FIN_END_DATE"
        rsChallanNo.GetResult(strChallanNo, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsChallanNo.GetNoRows > 0 Then
            strChallanNo = rsChallanNo.GetValue("CURRENT_NO").ToString
            While Len(strChallanNo) < 6
                strChallanNo = "0" + strChallanNo
            End While
            strChallanNo = "99" + strChallanNo
            txtChallanNo.Text = strChallanNo
            strChallanNo = "UPDATE DOCUMENTTYPE_MST SET CURRENT_NO = CURRENT_NO + 1 WHERE DOC_TYPE = 9999 and UNIT_CODE = '" & gstrUNITID & "' AND GETDATE() BETWEEN FIN_START_DATE AND FIN_END_DATE"
            mP_Connection.Execute(strChallanNo, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        Else
            MsgBox("Temporary Invoice No. Series Not Define. Invoice Entry Can Not Be Saved.", MsgBoxStyle.Information, ResolveResString(100))
            txtChallanNo.Text = ""
        End If
        rsChallanNo.ResultSetClose()
        rsChallanNo = Nothing
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        rsChallanNo = Nothing
    End Sub
    Private Function SetMaxLengthInSpread(ByRef pintDecimalSize As Short) As Object
        On Error GoTo ErrHandler
        Dim intRow As Short
        Dim strMin As String
        Dim strMax As String
        Dim intLoopCounter As Short
        Dim intDecimal As Short
        If pintDecimalSize < 2 Then
            pintDecimalSize = 2
        End If
        strMin = "0." : strMax = "99999999999999."
        For intLoopCounter = 1 To intDecimal
            strMin = strMin & "0"
            strMax = strMax & "9"
        Next
        With Me.SpChEntry
            For intRow = 1 To .maxRows
                .Row = intRow
                .Col = 1 : .TypeMaxEditLen = 16
                .Col = 2 : .TypeMaxEditLen = 30
                .Col = 1 : .Col2 = 2 : .Row = 1 : .Row2 = .maxRows : .BlockMode = True : .ColsFrozen = 2 : .BlockMode = False
                .Row = intRow
                .Col = 3 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = pintDecimalSize : .TypeFloatMin = CDbl(strMin) : .TypeFloatMax = CDbl(strMax)
                .Col = 4 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = pintDecimalSize : .TypeFloatMin = CDbl(strMin) : .TypeFloatMax = CDbl(strMax)
                .Col = 5 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = 2 : .TypeFloatMin = CDbl("0.00") : .TypeFloatMax = CDbl("99999999999999.99")
                .Col = 6 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                .CtlEditMode = False
                If CmbInvType.Text = "NORMAL INVOICE" Or CmbInvType.Text = "JOBWORK INVOICE" Or CmbInvType.Text = "EXPORT INVOICE" Or CmbInvType.Text = "TRANSFER INVOICE" Then
                    If UCase(Trim(CmbInvSubType.Text)) <> "SCRAP" Then
                        .Col = 7
                        .CtlEditMode = False
                    Else
                        .Col = 7
                        .CtlEditMode = True
                    End If
                Else
                    .Col = 7 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                    .CtlEditMode = True
                End If
                .Col = 10 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = 2 : .TypeFloatMin = CDbl("0.00") : .TypeFloatMax = CDbl("99999999999999.99")
                .Col = 11 : .Col2 = 11 : .Row = intRow : .Row2 = intRow : .BlockMode = True : .CellType = FPSpreadADO.CellTypeConstants.CellTypeInteger : .TypeIntegerMax = 9999 : .BlockMode = False
                .Col = 12 : .Col2 = 12 : .Row = intRow : .Row2 = intRow : .BlockMode = True : .CellType = FPSpreadADO.CellTypeConstants.CellTypeInteger : .TypeIntegerMax = 9999 : .BlockMode = False
                .Col = 13 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                .Col = 14 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                .Col = 15 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = pintDecimalSize : .TypeFloatMin = CDbl(strMin) : .TypeFloatMax = CDbl(strMax)
            Next intRow
        End With
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Public Function DeleteRecords() As Boolean
        On Error GoTo ErrHandler
        DeleteRecords = False
        strupSalechallan = ""
        strupSaleDtl = ""
        strupSalechallan = "Delete SalesChallan_Dtl where UNIT_CODE = '" & gstrUNITID & "' and Doc_No =" & Trim(txtChallanNo.Text)
        strupSalechallan = strupSalechallan & " and Location_Code ='" & Trim(txtLocationCode.Text) & "'"
        strupSaleDtl = "Delete Sales_Dtl where UNIT_CODE = '" & gstrUNITID & "' and Doc_No =" & Trim(txtChallanNo.Text)
        strupSaleDtl = strupSaleDtl & " and Location_Code ='" & Trim(txtLocationCode.Text) & "'"
        DeleteRecords = True
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Public Function CheckMeasurmentUnit(ByRef strItem As Object, ByRef strQuantity As Object, ByRef intRow As Short, ByRef blnQtyStatus As Boolean) As Boolean
        Dim strMeasure As String
        Dim rsMeasure As ClsResultSetDB
        On Error GoTo ErrHandler
        strMeasure = "select a.Decimal_allowed_flag from Measure_Mst a,Item_Mst b"
        strMeasure = strMeasure & " where a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND b.cons_Measure_Code=a.Measure_Code and b.Item_Code = '" & strItem & "'"
        rsMeasure = New ClsResultSetDB
        rsMeasure.GetResult(strMeasure, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        If rsMeasure.GetValue("Decimal_allowed_flag") = False Then
            rsMeasure.ResultSetClose()
            If System.Math.Round(strQuantity, 0) - Val(strQuantity) <> 0 Then
                If blnQtyStatus = True Then
                    MsgBox("Quantity of item -- " & strItem & "  is not defined in Decimal / Fraction .", MsgBoxStyle.Information, "eMpro")
                    CheckMeasurmentUnit = False
                    Call SpChEntry.SetText(5, intRow, strQuantity)
                    SpChEntry.Col = 5
                    SpChEntry.Row = intRow
                    SpChEntry.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                    SpChEntry.Focus()
                    Exit Function
                Else
                    MsgBox("Bin Quantity of item -- " & strItem & "  is not defined in Decimal / Fraction .", MsgBoxStyle.Information, "eMpro")
                    CheckMeasurmentUnit = False
                    Call SpChEntry.SetText(22, intRow, strQuantity)
                    SpChEntry.Col = 22
                    SpChEntry.Row = intRow
                    SpChEntry.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                    SpChEntry.Focus()
                    Exit Function
                End If
            Else
                CheckMeasurmentUnit = True
            End If
        Else
            rsMeasure.ResultSetClose()
            CheckMeasurmentUnit = True
        End If
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Public Function ParentQty(ByRef pstrItemCode As String, ByRef pstrfinished As Object) As Double
        On Error GoTo ErrHandler
        Dim strParentQty As String
        Dim rsParentQty As ClsResultSetDB
        strParentQty = "select sum(required_qty + waste_Qty) as TotalQty from Bom_Mst where finished_Product_code ='"
        strParentQty = strParentQty & pstrfinished & "' and UNIT_CODE = '" & gstrUNITID & "' and rawMaterial_Code ='" & pstrItemCode & "'"
        rsParentQty = New ClsResultSetDB
        rsParentQty.GetResult(strParentQty, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        ParentQty = rsParentQty.GetValue("TotalQty")
        rsParentQty.ResultSetClose()
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Function StockLocationSalesConf(ByRef pstrInvType As String, ByRef pstrInvSubtype As String, ByRef pstrFeild As String) As String
        Dim rsSalesConf As ClsResultSetDB
        Dim StockLocation As String
        On Error GoTo ErrHandler
        rsSalesConf = New ClsResultSetDB
        Select Case pstrFeild
            Case "DESCRIPTION"
                rsSalesConf.GetResult("Select Stock_Location from SaleConf Where Description ='" & Trim(pstrInvType) & "' and UNIT_CODE = '" & gstrUNITID & "' and Sub_type_Description ='" & Trim(pstrInvSubtype) & "' AND Location_Code='" & Trim(txtLocationCode.Text) & "' and (fin_start_date <= getdate() and fin_end_date >= getdate())")
            Case "TYPE"
                rsSalesConf.GetResult("Select Stock_Location from SaleConf Where Invoice_type ='" & Trim(pstrInvType) & "' and UNIT_CODE = '" & gstrUNITID & "' and Sub_type ='" & Trim(pstrInvSubtype) & "' AND Location_Code='" & Trim(txtLocationCode.Text) & "' and (fin_start_date <= getdate() and fin_end_date >= getdate())")
        End Select
        If rsSalesConf.GetNoRows > 0 Then
            StockLocation = rsSalesConf.GetValue("Stock_Location")
        End If
        rsSalesConf.ResultSetClose()
        StockLocationSalesConf = StockLocation
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Public Function ExploreBom(ByRef pstrItemCode As String, ByRef pstrFinishedQty As Object, ByRef pstrSPCurrentRow As Object, ByRef pstrFinishedProduct As String) As Boolean
        Dim strBomMstRaw As String
        Dim rsBomMstRaw As ClsResultSetDB
        Dim rsCustAnnexDtl As ClsResultSetDB
        Dim intBomMaxRaw As Short
        Dim intCurrentRaw As Short
        Dim dblTotalReqQty As Double
        Dim strCustAnnexDtl As String
        Dim strRGPQuote As String
        Dim rsVandorBom As ClsResultSetDB
        Dim rsItemMst As ClsResultSetDB
        On Error GoTo ErrHandler
        strBomMstRaw = "Select RawMaterial_Code,Required_qty + Waste_qty "
        strBomMstRaw = strBomMstRaw & " As TotalReqQty,Process_Type from Bom_Mst where "
        strBomMstRaw = strBomMstRaw & " item_Code ='" & strBomItem
        strBomMstRaw = strBomMstRaw & "' and UNIT_CODE = '" & gstrUNITID & "' and finished_product_code ='"
        strBomMstRaw = strBomMstRaw & pstrItemCode & "'"
        rsBomMstRaw = New ClsResultSetDB
        rsBomMstRaw.GetResult(strBomMstRaw, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        Dim intArrCount As Short
        Dim blnArrItemFound As Boolean
        If rsBomMstRaw.GetNoRows > 0 Then ' If Item Found in Bom Mst
            intBomMaxRaw = rsBomMstRaw.GetNoRows
            rsBomMstRaw.MoveFirst()
            For intCurrentRaw = 1 To intBomMaxRaw
                strBomItem = rsBomMstRaw.GetValue("RawMaterial_code")
                dblTotalReqQty = rsBomMstRaw.GetValue("TotalReqQty")
                'String for CustAnnex_dtl
                strCustAnnexDtl = "Select Item_Code,Balance_qty = sum(Balance_qty) from CustAnnex_hdr where Customer_code ='"
                strCustAnnexDtl = strCustAnnexDtl & Trim(txtCustCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "' "
                If blnFIFO = False Then
                    strRGPQuote = Replace(mstrRGP, "§", "','", 1)
                    strRGPQuote = "'" & strRGPQuote & "'"
                    strCustAnnexDtl = strCustAnnexDtl & " and ref57f4_no in ("
                    strCustAnnexDtl = strCustAnnexDtl & Trim(strRGPQuote) & ") "
                End If
                strCustAnnexDtl = strCustAnnexDtl & " and getdate() <= "
                strCustAnnexDtl = strCustAnnexDtl & " DateAdd(d, 180, ref57f4_date)"
                strCustAnnexDtl = strCustAnnexDtl & " and Item_code ='" & strBomItem & "' group by Item_code"
                rsCustAnnexDtl = New ClsResultSetDB
                rsCustAnnexDtl.GetResult(strCustAnnexDtl, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                If rsCustAnnexDtl.GetNoRows >= 1 Then 'if item Found in CustAnnex then replace that item from Parant string
                    rsVandorBom = New ClsResultSetDB
                    rsVandorBom.GetResult("Select RawMaterial_Code from Vendor_bom where Finish_Product_code = '" & pstrFinishedProduct & "' and UNIT_CODE = '" & gstrUNITID & "' and RawMaterial_code = '" & strBomItem & "' and Vendor_code = '" & txtCustCode.Text & "'")
                    If rsVandorBom.GetNoRows > 0 Then
                        rsCustAnnexDtl.MoveFirst()
                        inti = inti + 1
                        ReDim Preserve arrItem(inti)
                        ReDim Preserve arrQty(inti)
                        ReDim Preserve arrReqQty(inti)
                        blnArrItemFound = False
                        For intArrCount = 0 To UBound(arrItem) - 1 'to check if ITem Already there in ArrItem Array
                            If UCase(Trim(arrItem(intArrCount))) = UCase(Trim(rsCustAnnexDtl.GetValue("Item_code"))) Then
                                ' if found then sum up Requird Quantity in array arrReqQty and assign value true to blnArrITemFound
                                blnArrItemFound = True
                                arrReqQty(intArrCount) = arrReqQty(intArrCount) + (dblTotalReqQty * pstrFinishedQty)
                                If arrQty(intArrCount) < arrReqQty(intArrCount) Then ' to Check with Quantity supplieded in Cust Annex
                                    MsgBox("Customer Supplied Materail for Item " & arrItem(inti) & " is " & arrQty(inti) & " .", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, "eMPro")
                                    SpChEntry.Row = pstrSPCurrentRow
                                    SpChEntry.Col = 5
                                    SpChEntry.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                    ExploreBom = False
                                    Exit Function
                                Else
                                    ExploreBom = True
                                    Exit For
                                End If
                            Else
                                blnArrItemFound = False
                            End If
                        Next
                        If blnArrItemFound = False Then ' if item not found
                            inti = inti + 1
                            ReDim Preserve arrItem(inti)
                            ReDim Preserve arrQty(inti)
                            ReDim Preserve arrReqQty(inti)
                            arrItem(inti) = rsCustAnnexDtl.GetValue("Item_code")
                            arrQty(inti) = rsCustAnnexDtl.GetValue("Balance_qty")
                            arrReqQty(inti) = dblTotalReqQty * pstrFinishedQty
                            If arrQty(inti) < arrReqQty(inti) Then
                                MsgBox("Customer Supplied Materail for Item " & arrItem(inti) & " is " & arrQty(inti) & " .", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, "eMPro")
                                SpChEntry.Row = pstrSPCurrentRow
                                SpChEntry.Col = 5
                                SpChEntry.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                ExploreBom = False
                                Exit Function
                            Else
                                ExploreBom = True
                            End If
                        End If
                        rsCustAnnexDtl.ResultSetClose()
                    End If
                    rsVandorBom.ResultSetClose()
                Else
                    rsCustAnnexDtl.ResultSetClose()
                    rsVandorBom = New ClsResultSetDB
                    rsVandorBom.GetResult("Select RawMaterial_Code from Vendor_bom where Finish_Product_code = '" & pstrFinishedProduct & "' and UNIT_CODE = '" & gstrUNITID & "' and RawMaterial_code = '" & strBomItem & "' and vendor_code = '" & txtCustCode.Text & "'")
                    If rsVandorBom.GetNoRows > 0 Then
                        MsgBox("Item " & strBomItem & " is not supplied.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, "eMPro")
                        SpChEntry.Row = pstrSPCurrentRow
                        SpChEntry.Col = 5
                        SpChEntry.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                        ExploreBom = False
                        rsVandorBom.ResultSetClose()
                        Exit Function
                    Else 'if not of Process type I then again Explore
                        rsItemMst = New ClsResultSetDB
                        rsItemMst.GetResult("Select Item_Main_grp from Item_Mst Where  UNIT_CODE = '" & gstrUNITID & "' and Item_code = '" & strBomItem & "'")
                        If (UCase(rsItemMst.GetValue("Item_Main_grp")) = "R") Or (UCase(rsItemMst.GetValue("Item_Main_grp")) = "C") Then
                            ExploreBom = True
                        Else
                            pstrFinishedQty = pstrFinishedQty * dblTotalReqQty
                            Call ExploreBom(strBomItem, pstrFinishedQty, pstrSPCurrentRow, pstrFinishedProduct)
                        End If
                        rsItemMst.ResultSetClose()
                    End If
                    rsVandorBom.ResultSetClose()
                End If
                rsBomMstRaw.MoveNext()
            Next
        Else
            MsgBox("No BOM Defind for Item (" & pstrItemCode & ") defined in challan", MsgBoxStyle.Information, "eMPro")
            ExploreBom = False
            rsBomMstRaw.ResultSetClose()
            Exit Function
        End If
        rsBomMstRaw.ResultSetClose()
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Public Function BomCheck() As Boolean
        Dim intSpreadRow As Short
        Dim intSpCurrentRow As Short
        Dim intCurrentItem As Short
        Dim VarFinishedItem As Object
        Dim VarFinishedQty As Object
        Dim VarDelete As Object
        Dim strBomMst As String
        Dim strCustAnnexDtl As String
        Dim strRgpsWithQuots As String
        Dim intBomMaxItem As Short
        Dim rsBomMst As ClsResultSetDB
        Dim rsCustAnnexDtl As ClsResultSetDB
        Dim rsVandorBom As ClsResultSetDB
        Dim rsItemMst As ClsResultSetDB
        Dim dblTotalReqQty As Double
        On Error GoTo ErrHandler
        BomCheck = False
        intSpreadRow = SpChEntry.maxRows
        inti = 0
        Dim intArrCount As Short
        Dim blnItemFoundinArray As Boolean ' to be used to check if item already exist in Array arrItem where we are storing all item we found in Cust annex
        If SpChEntry.maxRows >= 1 Then
            For intSpCurrentRow = 1 To intSpreadRow
                With SpChEntry
                    VarFinishedItem = Nothing
                    Call .GetText(1, intSpCurrentRow, VarFinishedItem)
                    VarFinishedQty = Nothing
                    Call .GetText(5, intSpCurrentRow, VarFinishedQty)
                    VarDelete = Nothing
                    Call .GetText(14, intSpCurrentRow, VarDelete)
                End With
                If UCase(Trim(VarDelete)) <> "D" Then
                    strBomMst = "Select RawMaterial_Code,Process_type,Required_qty + Waste_qty "
                    strBomMst = strBomMst & " As TotalReqQty"
                    strBomMst = strBomMst & " from Bom_Mst where  UNIT_CODE = '" & gstrUNITID & "' and  Finished_Product_code ='"
                    strBomMst = strBomMst & VarFinishedItem & "' Order By Bom_Level"
                    rsBomMst = New ClsResultSetDB
                    rsBomMst.GetResult(strBomMst, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                    intBomMaxItem = rsBomMst.GetNoRows
                    rsBomMst.MoveFirst()
                    If intBomMaxItem > 0 Then ' Item Found in Bom_Mst
                        rsVandorBom = New ClsResultSetDB
                        rsVandorBom.GetResult("Select RawMaterial_Code from Vendor_bom where  UNIT_CODE = '" & gstrUNITID & "' and Finish_Product_code = '" & VarFinishedItem & "' and Vendor_code = '" & txtCustCode.Text & "'")
                        If rsVandorBom.GetNoRows > 0 Then
                            rsVandorBom.ResultSetClose()
                            'Loop for Parent Items of Items at First lavel
                            For intCurrentItem = 1 To intBomMaxItem
                                strBomItem = ""
                                strBomItem = rsBomMst.GetValue("RawMaterial_Code")
                                'String for CustAnnex_dtl
                                strCustAnnexDtl = "Select Item_Code,Balance_qty = sum(Balance_qty) from CustAnnex_hdr where  UNIT_CODE = '" & gstrUNITID & "' and Customer_code ='"
                                strCustAnnexDtl = strCustAnnexDtl & Trim(txtCustCode.Text) & "'"
                                If blnFIFO = False Then
                                    strRgpsWithQuots = Replace(mstrRGP, "§", "','", 1)
                                    strRgpsWithQuots = "'" & strRgpsWithQuots & "'"
                                    strCustAnnexDtl = strCustAnnexDtl & " and ref57f4_no in ("
                                    strCustAnnexDtl = strCustAnnexDtl & Trim(strRgpsWithQuots) & ") "
                                End If
                                strCustAnnexDtl = strCustAnnexDtl & " and getdate() <= "
                                strCustAnnexDtl = strCustAnnexDtl & " DateAdd(d, 180, ref57f4_date)"
                                strCustAnnexDtl = strCustAnnexDtl & " and Item_code ='" & strBomItem & "' group By Item_code"
                                rsCustAnnexDtl = New ClsResultSetDB
                                rsCustAnnexDtl.GetResult(strCustAnnexDtl)
                                If rsCustAnnexDtl.GetNoRows >= 1 Then 'if item Found in Cust Annex
                                    rsVandorBom = New ClsResultSetDB
                                    rsVandorBom.GetResult("Select RawMaterial_Code from Vendor_bom where  UNIT_CODE = '" & gstrUNITID & "' and Finish_Product_code = '" & VarFinishedItem & "'and RawMaterial_code = '" & strBomItem & "' and Vendor_code = '" & txtCustCode.Text & "'")
                                    If rsVandorBom.GetNoRows > 0 Then
                                        rsCustAnnexDtl.MoveFirst()
                                        ReDim Preserve arrItem(inti)
                                        ReDim Preserve arrQty(inti)
                                        ReDim Preserve arrReqQty(inti)
                                        dblTotalReqQty = ParentQty(strBomItem, VarFinishedItem)
                                        If inti > 0 Then
                                            blnItemFoundinArray = False
                                            For intArrCount = 0 To UBound(arrItem) - 1
                                                'if item already exist in array then to sumup required Quantity
                                                If UCase(Trim(arrItem(intArrCount))) = UCase(rsCustAnnexDtl.GetValue("Item_code")) Then
                                                    ' if item already exist in arritem then will sum up its requied Quantity in arrreqQty() and mark blnFoundinarray as true will be used later
                                                    blnItemFoundinArray = True
                                                    arrReqQty(intArrCount) = arrReqQty(intArrCount) + (dblTotalReqQty * VarFinishedQty)
                                                    If arrQty(intArrCount) < arrReqQty(intArrCount) Then 'in case if sum up is less then Quantity suplied in cust annex
                                                        MsgBox("Customer Supplied Materail for Item " & arrItem(inti) & "is" & arrQty(inti) & ".", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, "eMPro")
                                                        SpChEntry.Row = intSpCurrentRow
                                                        SpChEntry.Col = 5
                                                        SpChEntry.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                                        BomCheck = False
                                                        rsCustAnnexDtl.ResultSetClose()
                                                        rsBomMst.ResultSetClose()
                                                        Exit Function
                                                    End If
                                                End If
                                            Next
                                            If blnItemFoundinArray = False Then
                                                'in case item not found in arrItem with help of blnItemFoundinarray = false then will add new value to Arrays
                                                arrItem(inti) = rsCustAnnexDtl.GetValue("Item_code")
                                                arrQty(inti) = rsCustAnnexDtl.GetValue("Balance_qty")
                                                arrReqQty(inti) = dblTotalReqQty * VarFinishedQty
                                                If arrQty(inti) < arrReqQty(inti) Then 'again  check for Quantity requird as compare to supplied in CustAnnex
                                                    MsgBox("Customer Supplied Material for Item " & arrItem(inti) & "is" & arrQty(inti) & ".", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, "eMPro")
                                                    SpChEntry.Row = intSpCurrentRow
                                                    SpChEntry.Col = 5
                                                    SpChEntry.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                                    rsBomMst.ResultSetClose()
                                                    BomCheck = False
                                                    Exit Function
                                                End If
                                            End If
                                        Else ' if inti=0 then to add values
                                            arrItem(inti) = rsCustAnnexDtl.GetValue("Item_code")
                                            arrQty(inti) = rsCustAnnexDtl.GetValue("Balance_qty")
                                            arrReqQty(inti) = dblTotalReqQty * VarFinishedQty
                                            If arrQty(inti) < arrReqQty(inti) Then 'Again Same Check
                                                MsgBox("Customer Supplied Material for Item " & arrItem(inti) & "is" & arrQty(inti) & ".", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, "eMPro")
                                                SpChEntry.Row = intSpCurrentRow
                                                SpChEntry.Col = 5
                                                SpChEntry.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                                rsBomMst.ResultSetClose()
                                                BomCheck = False
                                                Exit Function
                                            End If
                                        End If
                                    End If
                                    rsVandorBom.ResultSetClose()
                                Else ' if Item Not Found in Cust Annex
                                    rsVandorBom = New ClsResultSetDB
                                    rsVandorBom.GetResult("Select RawMaterial_Code from Vendor_bom where  UNIT_CODE = '" & gstrUNITID & "' and Finish_Product_code = '" & VarFinishedItem & "'and RawMaterial_code = '" & strBomItem & "' and Vendor_code = '" & txtCustCode.Text & "'")
                                    If rsVandorBom.GetNoRows > 0 Then
                                        MsgBox("Item " & strBomItem & " is not supplied by Customer.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, "eMPro")
                                        SpChEntry.Row = intSpCurrentRow
                                        SpChEntry.Col = 5
                                        SpChEntry.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                        BomCheck = False
                                        rsVandorBom.ResultSetClose()
                                        Exit Function
                                    Else
                                        rsItemMst = New ClsResultSetDB
                                        rsItemMst.GetResult("Select Item_Main_grp from Item_Mst Where  UNIT_CODE = '" & gstrUNITID & "' and Item_code = '" & strBomItem & "'")
                                        If (UCase(rsItemMst.GetValue("Item_Main_grp")) = "R") Or (UCase(rsItemMst.GetValue("Item_Main_grp")) = "C") Then
                                            BomCheck = True
                                        Else
                                            VarFinishedQty = VarFinishedQty * rsBomMst.GetValue("TotalReqQty")
                                            If ExploreBom(strBomItem, VarFinishedQty, intSpCurrentRow, CStr(VarFinishedItem)) = False Then
                                                BomCheck = False
                                                Exit Function
                                            End If
                                        End If
                                        rsItemMst.ResultSetClose()
                                    End If
                                    rsVandorBom.ResultSetClose()
                                End If
                                rsCustAnnexDtl.ResultSetClose()
                                rsBomMst.MoveNext()
                                inti = inti + 1
                            Next
                        Else
                            rsVandorBom.ResultSetClose()
                            MsgBox("No Customer BOM Defind for Item (" & VarFinishedItem & ") defined in challan", MsgBoxStyle.Information, "eMPro")
                            BomCheck = False
                            Exit Function
                        End If
                    Else ' if no Item Found from Grid
                        MsgBox("No BOM Defind for Item (" & VarFinishedItem & ") defined in challan", MsgBoxStyle.Information, "eMPro")
                        rsBomMst.ResultSetClose()
                        BomCheck = False
                        Exit Function
                    End If
                    rsBomMst.ResultSetClose()
                End If
            Next
        End If
        BomCheck = True
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Public Function ToGetDecimalPlaces(ByRef pstrCurrency As String) As Short
        On Error GoTo ErrHandler
        Dim rscurrency As ClsResultSetDB
        rscurrency = New ClsResultSetDB
        rscurrency.GetResult("Select Decimal_Place from Currency_Mst where Currency_code ='" & pstrCurrency & "' and UNIT_CODE = '" & gstrUNITID & "'")
        ToGetDecimalPlaces = Val(rscurrency.GetValue("Decimal_Place"))
        rscurrency.ResultSetClose()
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Function
    End Function
    Private Function SelectDataFromCustOrd_Dtl(ByRef pstrCustCode As String, ByRef pstrInvType As String) As Boolean
        On Error GoTo ErrHandler
        Dim strSelectSql As String 'Declared To Make Select Query
        Dim strSelectPerValue As String
        Dim rsCustOrdDtl As ClsResultSetDB
        Dim rsCustOrdHdr As ClsResultSetDB
        SelectDataFromCustOrd_Dtl = False
        If UCase(pstrInvType) = "JOBWORK INVOICE" Then
            strSelectSql = "Select b.Item_Code,b.Cust_DrgNo,b.Amendment_No,b.Cust_Ref from Cust_Ord_hdr a,Cust_Ord_Dtl b"
            strSelectSql = strSelectSql & " where a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND  b.Account_Code='" & Trim(pstrCustCode) & "' and b.Active_flag ='A' and "
            strSelectSql = strSelectSql & " a.Account_Code = b.Account_Code and a.Cust_ref =b.Cust_ref and "
            'Active Flag to Be Checked Item Wise and Not for Sales Order
            strSelectSql = strSelectSql & " a.Amendment_No = b.amendment_No AND a.Authorized_Flag = 1 and a.PO_type in ('J') "
            strSelectSql = strSelectSql & " and a.Valid_date >='" & getDateForDB(GetServerDate()) & "' and effect_Date <='" & getDateForDB(GetServerDate()) & "'"
            strSelectSql = strSelectSql & " AND b.Cust_Ref = '" & Trim(txtRefNo.Text) & "'"
            strSelectSql = strSelectSql & " order by b.Cust_Ref,b.Amendment_No,b.Cust_DrgNo,b.Item_Code "
            '*******To Fatch Per Value
            strSelectPerValue = "Select distinct PerValue from Cust_Ord_hdr "
            strSelectPerValue = strSelectPerValue & " where Account_Code='" & Trim(pstrCustCode) & "' and UNIT_CODE = '" & gstrUNITID & "'  and Active_flag ='A' and "
            strSelectPerValue = strSelectPerValue & " Authorized_Flag = 1 and PO_type in ('J') "
            strSelectPerValue = strSelectPerValue & " and Valid_date >='" & getDateForDB(GetServerDate()) & "' and effect_Date <='" & getDateForDB(GetServerDate()) & "'"
            strSelectPerValue = strSelectPerValue & " AND Cust_Ref = '" & Trim(txtRefNo.Text) & "'"
        ElseIf UCase(pstrInvType) = "EXPORT INVOICE" Then
            strSelectSql = "Select b.Item_Code,b.Cust_DrgNo,b.Amendment_No,b.Cust_Ref from Cust_Ord_hdr a,Cust_Ord_Dtl b"
            strSelectSql = strSelectSql & " where a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND  b.Account_Code='" & Trim(pstrCustCode) & "' and b.Active_flag ='A' and "
            strSelectSql = strSelectSql & " a.Account_Code = b.Account_Code and a.Cust_ref =b.Cust_ref and "
            'Active Flag to Be Checked Item Wise and Not for Sales Order
            strSelectSql = strSelectSql & " a.Amendment_No = b.amendment_No AND a.Authorized_Flag = 1 and a.PO_type in ('E') "
            strSelectSql = strSelectSql & " and a.Valid_date >='" & getDateForDB(GetServerDate()) & "' and effect_date <='" & getDateForDB(GetServerDate()) & "'"
            strSelectSql = strSelectSql & " AND b.Cust_Ref = '" & Trim(txtRefNo.Text) & "'"
            strSelectSql = strSelectSql & " order by b.Cust_Ref,b.Amendment_No,b.Cust_DrgNo,b.Item_Code "
            '*******To Fatch Per Value
            strSelectPerValue = "Select distinct PerValue from Cust_Ord_hdr "
            strSelectPerValue = strSelectPerValue & " where  Account_Code='" & Trim(pstrCustCode) & "' and Active_flag ='A' and "
            strSelectPerValue = strSelectPerValue & " Authorized_Flag = 1 and PO_type in ('E') "
            strSelectPerValue = strSelectPerValue & " and Valid_date >='" & getDateForDB(GetServerDate()) & "' and effect_Date <='" & getDateForDB(GetServerDate()) & "'"
            strSelectPerValue = strSelectPerValue & " AND Cust_Ref = '" & Trim(txtRefNo.Text) & "' and UNIT_CODE = '" & gstrUNITID & "' "
        ElseIf UCase(pstrInvType) = "REJECTION" Then
            strSelectSql = "select a.Doc_No,a.Item_code,a.Rejected_Quantity from grn_Dtl a,grn_hdr b Where a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND "
            strSelectSql = strSelectSql & "a.Doc_type = b.Doc_type And a.Doc_No = b.Doc_No and "
            strSelectSql = strSelectSql & "a.From_Location = b.From_Location and a.From_Location ='01R1'"
            strSelectSql = strSelectSql & "and a.Rejected_quantity > 0  and b.Vendor_code = '" & pstrCustCode & "' AND A.Doc_No = " & txtRefNo.Text & " and isnull(b.GRN_Cancelled,0) = 0 order by a.Doc_No"
        Else
            strSelectSql = "Select b.Item_Code,b.Cust_DrgNo,b.Amendment_No,b.Cust_Ref from Cust_Ord_hdr a,Cust_Ord_Dtl b"
            strSelectSql = strSelectSql & " where a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND  b.Account_Code='" & Trim(pstrCustCode) & "' and b.Active_flag ='A' and "
            strSelectSql = strSelectSql & " a.Account_Code = b.Account_Code and a.Cust_ref =b.Cust_ref and "
            'Active Flag to Be Checked Item Wise and Not for Sales Order
            strSelectSql = strSelectSql & " a.Amendment_No = b.amendment_No  AND a.Authorized_Flag = 1 and a.PO_type in ('O','S','M') "
            strSelectSql = strSelectSql & " and a.Valid_date >='" & getDateForDB(GetServerDate()) & "' and effect_Date <= '" & getDateForDB(GetServerDate()) & "'"
            strSelectSql = strSelectSql & " AND b.Cust_Ref = '" & Trim(txtRefNo.Text) & "'"
            strSelectSql = strSelectSql & " order by b.Cust_Ref,b.Amendment_No,b.Cust_DrgNo,b.Item_Code "
            '*******To Fatch Per Value
            strSelectPerValue = "Select distinct PerValue from Cust_Ord_hdr "
            strSelectPerValue = strSelectPerValue & " where Account_Code='" & Trim(pstrCustCode) & "' and Active_flag ='A' and "
            strSelectPerValue = strSelectPerValue & " Authorized_Flag = 1 and PO_type in ('O','S','M') "
            strSelectPerValue = strSelectPerValue & " and Valid_date >='" & getDateForDB(GetServerDate()) & "' and effect_Date <='" & getDateForDB(GetServerDate()) & "'"
            strSelectPerValue = strSelectPerValue & " AND Cust_Ref = '" & Trim(txtRefNo.Text) & "' and UNIT_CODE = '" & gstrUNITID & "'"
        End If
        rsCustOrdDtl = New ClsResultSetDB
        rsCustOrdDtl.GetResult(strSelectSql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        If rsCustOrdDtl.GetNoRows > 0 Then '          'if record found
            If UCase(pstrInvType) <> "REJECTION" Then
                rsCustOrdHdr = New ClsResultSetDB
                rsCustOrdHdr.GetResult(strSelectPerValue)
                If rsCustOrdHdr.GetNoRows = 1 Then
                    rsCustOrdHdr.MoveFirst()
                    ctlPerValue.Text = rsCustOrdHdr.GetValue("PerValue")
                End If
                rsCustOrdHdr.ResultSetClose()
            End If
            SelectDataFromCustOrd_Dtl = True 'Return TRUE
            rsCustOrdDtl.ResultSetClose()
            
            rsCustOrdDtl = Nothing
        End If
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Function OriginalRefNoOVER(ByVal strRefNumber As String) As Boolean
        On Error GoTo ErrHandler
        '1st Check if Any Blank Amendment no for Ref No. Exists
        If SelectDataFromTable("Active_Flag", "Cust_ORD_HDR", " Where Account_Code ='" & Trim(txtCustCode.Text) & "' AND Cust_Ref = '" & txtRefNo.Text & "' AND Amendment_No = '' and UNIT_CODE = '" & gstrUNITID & "'") = "O" Then
            OriginalRefNoOVER = True
        Else
            OriginalRefNoOVER = False
        End If
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Function SelectDataFromTable(ByRef mstrFieldName As String, ByRef mstrTableName As String, ByRef mstrCondition As String) As String
        Dim StrSQLQuery As String
        Dim GetDataFromTable As ClsResultSetDB
        On Error GoTo ErrHandler
        StrSQLQuery = "Select " & mstrFieldName & " From " & mstrTableName & mstrCondition
        GetDataFromTable = New ClsResultSetDB
        If GetDataFromTable.GetResult(StrSQLQuery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic) Then
            If GetDataFromTable.GetNoRows > 0 Then
                
                SelectDataFromTable = GetDataFromTable.GetValue(mstrFieldName)
            Else
                SelectDataFromTable = ""
            End If
        Else
            SelectDataFromTable = ""
        End If
        GetDataFromTable.ResultSetClose()
        Exit Function
ErrHandler:
        gobjError.RaiseError(Err.Number, err.Source, Err.Description, mP_Connection)
    End Function
    Private Function DateIsAppropriate() As Boolean
        On Error GoTo ErrHandler
        Dim MaxInvoiceDate As String 'Get Max Date of Last Invoice made
        Dim CurrentDate As Date
        MaxInvoiceDate = Find_Value("select convert(varchar(11),max(Invoice_date),106) as invoice_date from saleschallan_dtl where Bill_Flag=1 and UNIT_CODE = '" & gstrUNITID & "'")
        CurrentDate = GetServerDate()
        If Len(MaxInvoiceDate) = 0 Then
            MaxInvoiceDate = getDateForDB(CurrentDate)
        End If
        If (dtpDateDesc.Value <= CurrentDate) And (dtpDateDesc.Value >= CDate(MaxInvoiceDate)) Then
            'Date Being Entered Falls in Limitations
            DateIsAppropriate = True
        Else
            'Date Being Entered Does not Falls in Limitations
            DateIsAppropriate = False
        End If
        Exit Function
ErrHandler:
        gobjError.RaiseError(Err.Number, err.Source, Err.Description, mP_Connection)
    End Function
    Public Function AddDataTolstRGPs() As Boolean
        Dim rsCustAnnex As ClsResultSetDB
        Dim intLoopCounter As Short
        Dim intMaxCounter As Short
        On Error GoTo ErrHandler
        With lvwRGPs
            .Gridlines = True : .Items.Clear() : .Columns.Clear()
            Call .Columns.Insert(0, "", "RGP No(s)", CInt(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(lvwRGPs.Width) / 2)))
            Call .Columns.Insert(1, "", "RGP Date", CInt(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(lvwRGPs.Width) / 2 - 700)))
            rsCustAnnex = New ClsResultSetDB
            rsCustAnnex.GetResult("select distinct ref57f4_No,ref57f4_date from custannex_HDR where  UNIT_CODE = '" & gstrUNITID & "' and customer_Code='" & Trim(txtCustCode.Text) & "' and getdate() < dateadd(d,180,ref57f4_Date) order by ref57f4_Date")
            If rsCustAnnex.GetNoRows > 0 Then
                AddDataTolstRGPs = True
                intMaxCounter = rsCustAnnex.GetNoRows
                rsCustAnnex.MoveFirst()
                For intLoopCounter = 0 To intMaxCounter - 1
                    Call .Items.Insert(intLoopCounter, rsCustAnnex.GetValue("ref57f4_No"))
                    If .Items.Item(intLoopCounter).SubItems.Count > 1 Then
                        .Items.Item(intLoopCounter).SubItems(1).Text = VB6.Format(rsCustAnnex.GetValue("ref57f4_Date"), gstrDateFormat)
                    Else
                        .Items.Item(intLoopCounter).SubItems.Insert(1, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, VB6.Format(rsCustAnnex.GetValue("ref57f4_Date"), gstrDateFormat)))
                    End If
                    rsCustAnnex.MoveNext()
                Next
            Else
                AddDataTolstRGPs = False
            End If
            rsCustAnnex.ResultSetClose()
        End With
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Public Function CheckExchangeRate() As Boolean
        On Error GoTo ErrHandler
        If Trim(lblExchangeRateValue.Text) = "" Then
            MsgBox("Please Define Exchange Rate For this Month in Exchange Master", MsgBoxStyle.Information, "eMPro")
            CheckExchangeRate = False
        Else
            mExchageRate = Val(Trim(lblExchangeRateValue.Text))
            CheckExchangeRate = True
        End If
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Public Function ItemQtyCaseRejGrin() As Boolean
        Dim rsGrnDtl As ClsResultSetDB
        Dim strSQL As String
        Dim varItemCode As Object
        Dim varItemQty As Object
        Dim VarDelete As Object
        Dim dblRejQty As Double
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        On Error GoTo ErrHandler
        intMaxLoop = SpChEntry.maxRows
        ItemQtyCaseRejGrin = False
        For intLoopCounter = 1 To intMaxLoop
            varItemCode = Nothing
            Call SpChEntry.GetText(1, intLoopCounter, varItemCode)
            VarDelete = Nothing
            Call SpChEntry.GetText(14, intLoopCounter, VarDelete)
            varItemQty = Nothing
            Call SpChEntry.GetText(5, intLoopCounter, varItemQty)
            If VarDelete <> "D" Then
                strSQL = "select a.Doc_No,a.Item_code, MaxAllowedQty = ((isnull(a.Rejected_Quantity,0) + isnull(a.Excess_PO_Quantity,0)) - (isnull(a.Despatch_Quantity,0) + isnull(a.Inspected_Quantity,0) + isnull(a.RGP_Quantity,0)))from grn_Dtl a,grn_hdr b Where a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND "
                strSQL = strSQL & " a.Doc_type = b.Doc_type And a.Doc_No = b.Doc_No and "
                strSQL = strSQL & "a.From_Location = b.From_Location and a.From_Location ='01R1'"
                strSQL = strSQL & "and a.Rejected_quantity > 0 and b.Vendor_code = '" & txtCustCode.Text
                strSQL = strSQL & "' and a.Doc_No = " & CDbl(txtRefNo.Text) & " and a.Item_code = '" & varItemCode & "' and isnull(b.GRN_Cancelled,0) = 0"
                rsGrnDtl = New ClsResultSetDB
                rsGrnDtl.GetResult(strSQL)
                If rsGrnDtl.GetNoRows > 0 Then
                    dblRejQty = rsGrnDtl.GetValue("MaxAllowedQty")
                    rsGrnDtl.ResultSetClose()
                    If varItemQty > dblRejQty Then
                        MsgBox("Quantity Allowed For This Item is " & dblRejQty & ", cannot Enter More then This.")
                        SpChEntry.Row = intLoopCounter : SpChEntry.Col = 5 : SpChEntry.Action = FPSpreadADO.ActionConstants.ActionActiveCell : SpChEntry.Focus()
                        ItemQtyCaseRejGrin = False
                        Exit Function
                    Else
                        ItemQtyCaseRejGrin = True
                    End If
                Else
                    rsGrnDtl.ResultSetClose()
                    MsgBox("No Item -" & varItemCode & " available in GRIN No - " & txtRefNo.Text & " Having Rejected Quantity >0 ")
                    ItemQtyCaseRejGrin = False
                    Exit Function
                End If
            Else
                ItemQtyCaseRejGrin = True
            End If
        Next
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Public Function ScheduleCheckEditMode() As Boolean
        On Error GoTo ErrHandler
        Dim strScheduleSql As String
        Dim strScheduleSql1 As String
        Dim varDrgNo As Object
        Dim strQuantity As Double
        Dim rsMktSchedule As New ClsResultSetDB
        Dim rsMktSchedule1 As New ClsResultSetDB
        Dim varItemQty As Object
        Dim VarDelete As Object
        Dim PresQty As Object
        Dim intRwCount As Short
        Dim varItemCode As Object
        Dim intLoopCount As Short
        Dim strMakeDate As String
        If ((UCase(mstrInvType) = "INV") And (UCase(mstrInvSubType) = "F") Or (UCase(mstrInvSubType) = "T")) Or (UCase(Trim(CmbInvType.Text)) = "JOBWORK INVOICE") Or (UCase(mstrInvType) = "EXP") Then
            'Check From DailyMktSchedule
            strScheduleSql = "Select Quantity=Schedule_Quantity-isnull(Despatch_Qty,0),Cust_DrgNo,Item_Code from DailyMktSchedule where  UNIT_CODE = '" & gstrUNITID & "' and  Account_Code='" & Trim(txtCustCode.Text) & "' and "
            strScheduleSql = strScheduleSql & " datepart(yyyy,Trans_Date)='" & Year(CDate(getDateForDB(Trim(lblDateDes.Text)))) & "'"
            strScheduleSql = strScheduleSql & " and datepart(mm,Trans_Date)='" & Month(CDate(getDateForDB(Trim(lblDateDes.Text)))) & "'"
            strScheduleSql = strScheduleSql & " and datepart(dd,Trans_Date)='" & VB.Day(CDate(getDateForDB(Trim(lblDateDes.Text)))) & "'"
            strScheduleSql = strScheduleSql & " and Cust_DrgNo in(" & Trim(mstrItemCode) & ") and Status =1 and Schedule_Flag =1"
            rsMktSchedule = New ClsResultSetDB
            rsMktSchedule.GetResult(strScheduleSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If rsMktSchedule.GetNoRows > 0 Then 'If Record Found
                rsMktSchedule.MoveFirst()
                For intRwCount = 1 To Me.SpChEntry.maxRows
                    'Select Quantity From The Spread
                    varItemQty = Nothing
                    Call Me.SpChEntry.GetText(5, intRwCount, varItemQty)
                    VarDelete = Nothing
                    Call Me.SpChEntry.GetText(14, intRwCount, VarDelete)
                    strQuantity = rsMktSchedule.GetValue("Quantity")
                    'If Quantity Entered Is Greater Then Schedule Quantity
                    If UCase(VarDelete) <> "D" Then
                        If (Val(varItemQty) - Val(mdblPrevQty(intLoopCount))) > Val(CStr(strQuantity)) Then
                            ScheduleCheckEditMode = False
                            MsgBox("Quantity should not be Greater then Schedule Quantity " & strQuantity)
                            With Me.SpChEntry
                                .Row = intRwCount : .Col = 5 : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                            End With
                            Exit Function
                        Else
                            ScheduleCheckEditMode = True
                            ' Make Update Query For Dispatch
                            mstrUpdDispatchSql = ""
                            For intLoopCount = 1 To SpChEntry.MaxRows
                                varDrgNo = Nothing
                                Call Me.SpChEntry.GetText(2, intLoopCount, varDrgNo)
                                varItemCode = Nothing
                                Call Me.SpChEntry.GetText(1, intLoopCount, varItemCode)
                                PresQty = Nothing
                                Call Me.SpChEntry.GetText(5, intLoopCount, PresQty)
                                strScheduleSql = "select Despatch_qty  = isnull(Despatch_Qty,0) - (" & Val(mdblPrevQty(intLoopCount - 1)) - Val(PresQty) & "),SChedule_Quantity from DailyMktSchedule "
                                strScheduleSql = strScheduleSql & " Where  UNIT_CODE = '" & gstrUNITID & "' and  Account_Code='" & Trim(txtCustCode.Text) & "' and "
                                strScheduleSql = strScheduleSql & " datepart(yyyy,Trans_Date)='" & Year(CDate(getDateForDB(Trim(lblDateDes.Text)))) & "'"
                                strScheduleSql = strScheduleSql & " and datepart(mm,Trans_Date)='" & Month(CDate(getDateForDB(Trim(lblDateDes.Text)))) & "'"
                                strScheduleSql = strScheduleSql & " and datepart(dd,Trans_Date)='" & VB.Day(CDate(getDateForDB(Trim(lblDateDes.Text)))) & "'"
                                strScheduleSql = strScheduleSql & " and Cust_DrgNo ='" & Trim(varDrgNo) & "'and Item_code = '" & varItemCode & "' and Status =1 and Schedule_flag =1" & vbCrLf
                                rsMktSchedule1.GetResult(strScheduleSql)
                                mstrUpdDispatchSql = Trim(mstrUpdDispatchSql) & "Update DailyMktSchedule set Despatch_qty ="
                                mstrUpdDispatchSql = mstrUpdDispatchSql & "isnull(Despatch_Qty,0) - (" & Val(mdblPrevQty(intLoopCount - 1)) - Val(PresQty) & ")"
                                If Val(rsMktSchedule1.GetValue("Despatch_Qty")) = Val(rsMktSchedule1.GetValue("Schedule_Quantity")) Then
                                    mstrUpdDispatchSql = mstrUpdDispatchSql & ", Schedule_Flag = 0 "
                                End If
                                mstrUpdDispatchSql = mstrUpdDispatchSql & " Where  UNIT_CODE = '" & gstrUNITID & "' and  Account_Code='" & Trim(txtCustCode.Text) & "' and "
                                mstrUpdDispatchSql = mstrUpdDispatchSql & " datepart(yyyy,Trans_Date)='" & Year(CDate(getDateForDB(Trim(lblDateDes.Text)))) & "'"
                                mstrUpdDispatchSql = mstrUpdDispatchSql & " and datepart(mm,Trans_Date)='" & Month(CDate(getDateForDB(Trim(lblDateDes.Text)))) & "'"
                                mstrUpdDispatchSql = mstrUpdDispatchSql & " and datepart(dd,Trans_Date)='" & VB.Day(CDate(getDateForDB(Trim(lblDateDes.Text)))) & "'"
                                mstrUpdDispatchSql = mstrUpdDispatchSql & " and Cust_DrgNo ='" & Trim(varDrgNo) & "'and Item_code = '" & varItemCode & "' and Status =1 and Schedule_flag =1" & vbCrLf
                            Next
                        End If
                    End If
                    rsMktSchedule.MoveNext()
                Next intRwCount
                'If Record Not Found In DailyMktSchedule Then Check From
                'MonthlyMktSchedule
            ElseIf rsMktSchedule.GetNoRows = 0 Then
                If Val(CStr(Month(CDate(getDateForDB(lblDateDes.Text))))) < 10 Then
                    strMakeDate = Year(CDate(getDateForDB(lblDateDes.Text))) & "0" & Month(CDate(getDateForDB(lblDateDes.Text)))
                Else
                    strMakeDate = Year(CDate(getDateForDB(lblDateDes.Text))) & Month(CDate(getDateForDB(lblDateDes.Text)))
                End If
                strScheduleSql = "Select Quantity=Schedule_Qty-isnull(Despatch_Qty,0) from MonthlyMktSchedule where  UNIT_CODE = '" & gstrUNITID & "' and  Account_Code='" & Trim(txtCustCode.Text) & "' and "
                strScheduleSql = strScheduleSql & " Year_Month=" & Val(Trim(strMakeDate)) & ""
                strScheduleSql = strScheduleSql & " and Cust_DrgNo in(" & Trim(mstrItemCode) & ") and status =1 and Schedule_flag =1"
                rsMktSchedule.ResultSetClose()
                rsMktSchedule = New ClsResultSetDB
                rsMktSchedule.GetResult(strScheduleSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                If rsMktSchedule.GetNoRows > 0 Then
                    rsMktSchedule.MoveFirst()
                    For intRwCount = 1 To Me.SpChEntry.maxRows
                        Select Case CmdGrpChEnt.mode
                            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                                strQuantity = rsMktSchedule.GetValue("Quantity")
                                varItemQty = Nothing
                                Call Me.SpChEntry.GetText(5, intRwCount, varItemQty)
                            Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                                varItemQty = Nothing
                                Call Me.SpChEntry.GetText(5, intRwCount, varItemQty)
                                strQuantity = Val(rsMktSchedule.GetValue("Quantity")) + Val(varItemQty)
                        End Select
                        VarDelete = Nothing
                        Call Me.SpChEntry.GetText(11, intRwCount, VarDelete)
                        If UCase(VarDelete) <> "D" Then
                            If Val(varItemQty) > Val(CStr(strQuantity)) Then
                                ScheduleCheckEditMode = False
                                MsgBox("Quantity should not be Greater then Schedule Quantity " & strQuantity, MsgBoxStyle.Information, "eMPro")
                                With Me.SpChEntry
                                    .Row = intRwCount : .Col = 5 : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                                End With
                                Exit Function
                            Else
                                ScheduleCheckEditMode = False
                                mstrUpdDispatchSql = ""
                                For intLoopCount = 1 To SpChEntry.MaxRows
                                    varDrgNo = Nothing
                                    Call Me.SpChEntry.GetText(2, intLoopCount, varDrgNo)
                                    varItemCode = Nothing
                                    Call Me.SpChEntry.GetText(1, intLoopCount, varItemCode)
                                    PresQty = Nothing
                                    Call Me.SpChEntry.GetText(5, intLoopCount, PresQty)
                                    '**** To Check schedule Quantity
                                    strScheduleSql = "Select Despatch_qty = "
                                    strScheduleSql = strScheduleSql & "isnull(Despatch_Qty,0) - (" & Val(mdblPrevQty(intLoopCount - 1)) - Val(PresQty) & "),Schedule_Qty"
                                    strScheduleSql = strScheduleSql & " From MonthlyMktSchedule "
                                    strScheduleSql = strScheduleSql & " Where  UNIT_CODE = '" & gstrUNITID & "' and  Account_Code='" & Trim(txtCustCode.Text) & "' and "
                                    strScheduleSql = strScheduleSql & " Year_Month=" & Val(Trim(strMakeDate)) & ""
                                    strScheduleSql = strScheduleSql & " and Cust_DrgNo ='" & Trim(varDrgNo) & "'and Item_code = '" & varItemCode & "' and status =1 and Schedule_flag =1" & vbCrLf
                                    rsMktSchedule1.GetResult(strScheduleSql)
                                    mstrUpdDispatchSql = Trim(mstrUpdDispatchSql) & "Update MonthlyMktSchedule set Despatch_qty ="
                                    mstrUpdDispatchSql = mstrUpdDispatchSql & "isnull(Despatch_Qty,0) - (" & Val(mdblPrevQty(intLoopCount - 1)) - Val(PresQty) & ")"
                                    If rsMktSchedule1.GetValue("Despatch_Qty") = rsMktSchedule1.GetValue("Schedule_Qty") Then
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & ", Schedule_Flag = 0 "
                                    End If
                                    mstrUpdDispatchSql = mstrUpdDispatchSql & " Where  UNIT_CODE = '" & gstrUNITID & "' and  Account_Code='" & Trim(txtCustCode.Text) & "' and "
                                    mstrUpdDispatchSql = mstrUpdDispatchSql & " Year_Month=" & Val(Trim(strMakeDate)) & ""
                                    mstrUpdDispatchSql = mstrUpdDispatchSql & " and Cust_DrgNo ='" & Trim(varDrgNo) & "'and Item_code = '" & varItemCode & "' and status =1 and Schedule_flag =1" & vbCrLf
                                Next
                            End If
                        End If
                        rsMktSchedule.MoveNext()
                    Next intRwCount
                Else
                    MsgBox("No Schedule Defined For Selected Items,Define Schedule First")
                    ScheduleCheckEditMode = False
                    SpChEntry.Focus()
                    Exit Function
                End If
                rsMktSchedule.ResultSetClose()
            Else
                MsgBox("No Schedule Defined For Selected Items,Define Schedule First")
                ScheduleCheckEditMode = False
                SpChEntry.Focus()
                Exit Function
            End If
        End If
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Function
    End Function
    Private Function GetTaxRate(ByRef pstrFieldText As String, ByRef pstrColumnName As String, ByRef pstrTableName As String, ByRef pstrFieldName_WhichValueRequire As String, Optional ByRef pstrCondition As String = "") As Double
        On Error GoTo ErrHandler
        GetTaxRate = 0
        Dim strTableSql As String 'Declared To Make Select Query
        Dim rsExistData As ClsResultSetDB
        If Len(Trim(pstrCondition)) > 0 Then
            strTableSql = "select " & Trim(pstrFieldName_WhichValueRequire) & " from " & Trim(pstrTableName) & " where  UNIT_CODE = '" & gstrUNITID & "' and  " & Trim(pstrColumnName) & "='" & Trim(pstrFieldText) & "' and " & pstrCondition
        Else
            strTableSql = "select " & Trim(pstrFieldName_WhichValueRequire) & " from " & Trim(pstrTableName) & " where   UNIT_CODE = '" & gstrUNITID & "' and " & Trim(pstrColumnName) & "='" & Trim(pstrFieldText) & "'"
        End If
        rsExistData = New ClsResultSetDB
        rsExistData.GetResult(strTableSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsExistData.GetNoRows > 0 Then
            GetTaxRate = rsExistData.GetValue(Trim(pstrFieldName_WhichValueRequire))
        Else
            GetTaxRate = 0
        End If
        rsExistData.ResultSetClose()
        rsExistData = Nothing
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Function GetExchangeRate(ByVal pstrCurrencyCode As String, ByVal pstrDate As String, ByVal IsCustomer As Boolean) As Double
        On Error GoTo ErrHandler
        GetExchangeRate = 1.0#
        Dim strTableSql As String 'Declared To Make Select Query
        Dim rsExistData As ClsResultSetDB
        If IsCustomer = True Then
            strTableSql = "SELECT CExch_MultiFactor FROM Gen_CurExchMaster WHERE   UNIT_CODE = '" & gstrUNITID & "' and  CExch_CurrencyTo='" & Trim(pstrCurrencyCode) & "' AND CExch_InOut=1 AND  '" & Trim(pstrDate) & "' BETWEEN CExch_DateFrom AND CExch_DateTo"
        Else
            strTableSql = "SELECT CExch_MultiFactor FROM Gen_CurExchMaster WHERE   UNIT_CODE = '" & gstrUNITID & "' and  CExch_CurrencyTo='" & Trim(pstrCurrencyCode) & "' AND CExch_InOut=0 AND '" & Trim(pstrDate) & "' BETWEEN CExch_DateFrom AND CExch_DateTo"
        End If
        rsExistData = New ClsResultSetDB
        rsExistData.GetResult(strTableSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsExistData.GetNoRows > 0 Then
            GetExchangeRate = rsExistData.GetValue("CExch_MultiFactor")
        Else
            GetExchangeRate = 1.0#
        End If
        rsExistData.ResultSetClose()
        rsExistData = Nothing
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function

    Private Function SaveData(ByVal Button As String) As Boolean

        Dim ldblTotalBasicValue As Double
        Dim ldblTotalAccessibleValue As Double
        Dim lintLoopCounter As Short
        Dim ldblTempAccessibleVal As Double
        Dim ldblTotalExciseValue As Double
        Dim ldblTotalSaleTaxAmount As Double
        Dim ldblTotalSurchargeTaxAmount As Double
        Dim ldblNetInsurenceValue As Double
        Dim ldblTotalInvoiceValue As Double
        Dim ldblTotalOthersValues As Double
        Dim rsParameterData As ClsResultSetDB
        Dim strParamQuery As String
        Dim ldblTotalCVDValue As Double
        Dim ldblTotalSADValue As Double
        Dim ldblAllExciseValue As Double
        ''-----------Variable For Saving Data---------
        Dim strSalesChallan As String
        Dim updateSalesChallan As String
        Dim strSalesDtl As String
        Dim strSalesDtlDelete As String
        Dim rsCustItemMst As ClsResultSetDB
        Dim rsSaleConf As ClsResultSetDB
        Dim rsItemMst As ClsResultSetDB
        Dim rsPacking_Tax As ClsResultSetDB
        Dim lintItemQuantity As Double
        Dim lstrItemDrgno As String
        Dim lstrItemCode As String
        Dim ldblItemRate As Double
        Dim ldblItemCustMtrl As Double
        Dim ldblItemPacking As Double
        Dim strPackingCode As String
        Dim ldblItemOthers As Double
        Dim ldblItemFromBox As Double
        Dim ldblItemToBox As Double
        Dim lstrItemDelete As String
        Dim lintItemPresQty As Double
        Dim lstrItemExciseCode As String
        Dim lstrItemCVDCode As String
        Dim lstrItemSADCode As String
        Dim ldblItemToolCost As Double
        Dim TempAccessibleVal As Double
        Dim ldblTotalCustMatrlValue As Double
        Dim blnISInsExcisable As Boolean
        Dim blnISECESSRoundoff As Boolean
        Dim blnEOUFlag As Boolean
        Dim blnISExciseRoundOff As Boolean
        Dim blnISSalesTaxRoundOff As Boolean
        Dim blnISSurChargeTaxRoundOff As Boolean
        Dim blnAddCustMatrl As Boolean
        Dim blnISBasicRoundOff As Boolean
        Dim ldblExciseValueForSaleTax As Double
        Dim ldblTotalECESSAmount As Double
        Dim ldblTotalSECESSAmount As Double
        Dim blnDSWiseTracking As Boolean
        Dim strSDTType As String
        Dim dblSDT_Per As Double
        Dim dblSDT_Amt As Double
        Dim blnIsSDTRoundoff As Boolean
        Dim intSDTNoofDecimal As Short
        Dim dblBinQty As Double
        Dim strQry As String
        Dim rsCustOrdDtl As ClsResultSetDB
        Dim dblCustMtrl_SO As Double
        Dim intSTaxNoOfDecimal As Short
        Dim intEcessNoOfDecimal As Short
        Dim blnPackingRoundoff As Boolean
        Dim intPackingRoundoff_Decimal As Short
        Dim dblTotalPacking_Amount As Double
        Dim dblTempPacking_Amount As Double
        Dim dblItemPacking_Amount As Double
        Dim dblAddVATamount As Double
        Dim dblExcise_Amount As Double
        Dim strSqlct2qry As String
        Dim strsql As String
        Dim blnIsCt2 As Boolean = False

        On Error GoTo ErrHandler

        ldblTotalBasicValue = 0
        ldblTotalAccessibleValue = 0
        ldblTotalExciseValue = 0
        ldblAllExciseValue = 0
        ldblTotalCVDValue = 0
        ldblTotalSADValue = 0
        ldblTotalSaleTaxAmount = 0
        ldblTotalSurchargeTaxAmount = 0
        ldblTotalInvoiceValue = 0
        ldblTotalOthersValues = 0
        ldblTotalCustMatrlValue = 0
        ldblExciseValueForSaleTax = 0
        ldblTotalECESSAmount = 0
        dblSDT_Amt = 0
        dblCustMtrl_SO = 0
        intSTaxNoOfDecimal = 0
        intEcessNoOfDecimal = 0
        dblTotalPacking_Amount = 0
        SaveData = True
        strParamQuery = "SELECT InsExc_Excise,CustSupp_Inc,EOU_Flag,SalesTax_Roundoff,Basic_roundoff,Excise_Roundoff,SST_Roundoff,ECESS_Roundoff,DSWiseTracking,TotalInvoiceAmount_RoundOff,TotalInvoiceAmountRoundOff_Decimal, SDTRoundOff, SDTRoundOff_Decimal,salesTax_Roundoff_decimal,ECESSRoundoff_decimal,Packing_Roundoff,PackingRoundoff_Decimal FROM Sales_Parameter WHERE UNIT_CODE = '" & gstrUNITID & "'"
        rsParameterData = New ClsResultSetDB
        rsParameterData.GetResult(strParamQuery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsParameterData.GetNoRows > 0 Then
            blnISInsExcisable = rsParameterData.GetValue("InsExc_Excise")
            blnEOUFlag = rsParameterData.GetValue("EOU_Flag")
            blnISExciseRoundOff = rsParameterData.GetValue("Excise_Roundoff")
            blnISBasicRoundOff = rsParameterData.GetValue("Basic_Roundoff")
            blnISSalesTaxRoundOff = rsParameterData.GetValue("SalesTax_Roundoff")
            blnISSurChargeTaxRoundOff = rsParameterData.GetValue("SST_Roundoff")
            blnAddCustMatrl = rsParameterData.GetValue("CustSupp_Inc")
            blnISECESSRoundoff = rsParameterData.GetValue("ECESS_Roundoff")
            blnDSWiseTracking = IIf(IsDBNull(rsParameterData.GetValue("DSWiseTracking")), False, IIf(rsParameterData.GetValue("DSWiseTracking") = False, False, True))
            blnTotalInvoiceAmount = rsParameterData.GetValue("TotalInvoiceAmount_RoundOff")
            If rsParameterData.GetValue("TotalInvoiceAmountRoundOff_Decimal").ToString = "" Then
                intTotalInvoiceAmountRoundOffDecimal = 0
            Else
                intTotalInvoiceAmountRoundOffDecimal = rsParameterData.GetValue("TotalInvoiceAmountRoundOff_Decimal")
            End If
            blnIsSDTRoundoff = rsParameterData.GetValue("SDTRoundOff")
            intSDTNoofDecimal = rsParameterData.GetValue("SDTRoundOff_Decimal")
            If rsParameterData.GetValue("SDTRoundOff_Decimal").ToString = "" Then
                intSDTNoofDecimal = 0
            Else
                intSDTNoofDecimal = rsParameterData.GetValue("SDTRoundOff_Decimal")
            End If
            intSTaxNoOfDecimal = rsParameterData.GetValue("salesTax_Roundoff_decimal")
            If rsParameterData.GetValue("ECESSRoundoff_decimal").ToString = "" Then
                intEcessNoOfDecimal = 0
            Else
                intEcessNoOfDecimal = rsParameterData.GetValue("ECESSRoundoff_decimal")
            End If
            blnPackingRoundoff = rsParameterData.GetValue("Packing_Roundoff")
            If rsParameterData.GetValue("PackingRoundoff_Decimal").ToString = "" Then
                intPackingRoundoff_Decimal = 0
            Else
                intPackingRoundoff_Decimal = rsParameterData.GetValue("PackingRoundoff_Decimal")
            End If
        Else
            MsgBox("No data define in Sales_Parameter Table", MsgBoxStyle.Information, "eMPro")
            SaveData = False
            rsParameterData.ResultSetClose()
            rsParameterData = Nothing
            Exit Function
        End If
        rsParameterData.ResultSetClose()
        rsParameterData = Nothing
        strParamQuery = "SELECT decimal_place FROM currency_mst where currency_code='" & lblCurrencyDes.Text & "' and UNIT_CODE = '" & gstrUNITID & "'"
        rsParameterData = New ClsResultSetDB
        rsParameterData.GetResult(strParamQuery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsParameterData.GetNoRows > 0 Then
            mIntDecimalPlace = rsParameterData.GetValue("decimal_place")
        End If
        rsParameterData.ResultSetClose()
        rsParameterData = Nothing
        ldblNetInsurenceValue = System.Math.Round(Val(ctlInsurance.Text)) / Val(CStr(SpChEntry.MaxRows))
        For lintLoopCounter = 1 To SpChEntry.MaxRows
            dblTempPacking_Amount = CalculatePackingValue(lintLoopCounter, blnPackingRoundoff)
            dblTotalPacking_Amount = dblTotalPacking_Amount + CalculatePackingValue(lintLoopCounter, blnPackingRoundoff)
            ldblTotalBasicValue = ldblTotalBasicValue + CalculateBasicValue(lintLoopCounter, blnISBasicRoundOff)
            ldblTempAccessibleVal = CalculateAccessibleValue(lintLoopCounter, ldblNetInsurenceValue, blnISInsExcisable)
            If blnISExciseRoundOff Then
                ldblTotalExciseValue = System.Math.Round(CalculateExciseValue(lintLoopCounter, ldblTempAccessibleVal + dblTempPacking_Amount, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff))
                ldblTotalCVDValue = System.Math.Round(CalculateExciseValue(lintLoopCounter, ldblTempAccessibleVal + dblTempPacking_Amount, enumExciseType.RETURN_CVD, blnEOUFlag, blnISExciseRoundOff))
                ldblTotalSADValue = System.Math.Round(CalculateExciseValue(lintLoopCounter, ldblTempAccessibleVal + dblTempPacking_Amount, enumExciseType.RETURN_SAD, blnEOUFlag, blnISExciseRoundOff))
                ldblAllExciseValue = System.Math.Round(CalculateExciseValue(lintLoopCounter, ldblTempAccessibleVal + dblTempPacking_Amount, enumExciseType.RETURN_ALLExcise, blnEOUFlag, blnISExciseRoundOff))
            Else
                ldblTotalExciseValue = CalculateExciseValue(lintLoopCounter, ldblTempAccessibleVal + dblTempPacking_Amount, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff)
                ldblTotalCVDValue = System.Math.Round(CalculateExciseValue(lintLoopCounter, ldblTempAccessibleVal + dblTempPacking_Amount, enumExciseType.RETURN_CVD, blnEOUFlag, blnISExciseRoundOff))
                ldblTotalSADValue = System.Math.Round(CalculateExciseValue(lintLoopCounter, ldblTempAccessibleVal + dblTempPacking_Amount, enumExciseType.RETURN_SAD, blnEOUFlag, blnISExciseRoundOff))
                ldblAllExciseValue = System.Math.Round(CalculateExciseValue(lintLoopCounter, ldblTempAccessibleVal + dblTempPacking_Amount, enumExciseType.RETURN_ALLExcise, blnEOUFlag, blnISExciseRoundOff))
            End If
            ldblTotalAccessibleValue = ldblTotalAccessibleValue + ldblTempAccessibleVal
            SpChEntry.Row = lintLoopCounter : SpChEntry.Col = 5
            lintItemQuantity = Val(SpChEntry.Text)
            SpChEntry.Row = lintLoopCounter : SpChEntry.Col = 10
            ldblTotalOthersValues = ldblTotalOthersValues + ((Val(SpChEntry.Text) / CDbl(ctlPerValue.Text)) * lintItemQuantity)
            SpChEntry.Row = lintLoopCounter : SpChEntry.Col = 4
            ldblTotalCustMatrlValue = ldblTotalCustMatrlValue + ((Val(SpChEntry.Text) / CDbl(ctlPerValue.Text)) * lintItemQuantity)
            If blnEOU_FLAG Then
                If blnISExciseRoundOff Then
                    ldblExciseValueForSaleTax = ldblExciseValueForSaleTax + System.Math.Round((ldblTotalExciseValue + ldblTotalCVDValue + ldblTotalSADValue) / 2)
                Else
                    ldblExciseValueForSaleTax = ldblExciseValueForSaleTax + (ldblTotalExciseValue + ldblTotalCVDValue + ldblTotalSADValue) / 2
                End If
            Else
                If blnISExciseRoundOff Then
                    ldblExciseValueForSaleTax = ldblExciseValueForSaleTax + System.Math.Round(ldblTotalExciseValue)
                Else
                    ldblExciseValueForSaleTax = ldblExciseValueForSaleTax + ldblTotalExciseValue
                End If
            End If
        Next
        If blnISECESSRoundoff Then
            ldblTotalECESSAmount = System.Math.Round(CalculateECESSValue(ldblExciseValueForSaleTax))
            ldblTotalSECESSAmount = System.Math.Round(CalculateSECESSValue(ldblExciseValueForSaleTax))
        Else
            ldblTotalSECESSAmount = System.Math.Round(CalculateSECESSValue(ldblExciseValueForSaleTax), intEcessNoOfDecimal)
            ldblTotalECESSAmount = System.Math.Round(CalculateECESSValue(ldblExciseValueForSaleTax), intEcessNoOfDecimal)
        End If
        If blnISSalesTaxRoundOff Then
            ldblTotalSaleTaxAmount = System.Math.Round(CalculateSalesTaxValue(ldblTotalBasicValue, ldblExciseValueForSaleTax + ldblTotalECESSAmount + ldblTotalSECESSAmount + dblTotalPacking_Amount))
            dblAddVATamount = System.Math.Round(CalculateAddionalSalesTaxValue(ldblTotalBasicValue, ldblExciseValueForSaleTax + ldblTotalECESSAmount + ldblTotalSECESSAmount + dblTotalPacking_Amount))
        Else
            ldblTotalSaleTaxAmount = System.Math.Round(CalculateSalesTaxValue(ldblTotalBasicValue, ldblExciseValueForSaleTax + ldblTotalECESSAmount + ldblTotalSECESSAmount + dblTotalPacking_Amount), intSTaxNoOfDecimal)
            dblAddVATamount = System.Math.Round(CalculateAddionalSalesTaxValue(ldblTotalBasicValue, ldblExciseValueForSaleTax + ldblTotalECESSAmount + ldblTotalSECESSAmount + dblTotalPacking_Amount), intSTaxNoOfDecimal)
        End If
        If mstrInvoiceType <> "TRF" And mstrInvoiceType <> "JOB" Then
            If Len(Trim(txtSDTType.Text)) <> 0 Then
                If blnIsSDTRoundoff Then
                    dblSDT_Amt = System.Math.Round(CalculateStateDevelopmentTaxValue(ldblTotalBasicValue, ldblExciseValueForSaleTax + ldblTotalECESSAmount + ldblTotalSECESSAmount + dblTotalPacking_Amount))
                Else
                    dblSDT_Amt = System.Math.Round(CalculateStateDevelopmentTaxValue(ldblTotalBasicValue, ldblExciseValueForSaleTax + ldblTotalECESSAmount + ldblTotalSECESSAmount + dblTotalPacking_Amount), intSDTNoofDecimal)
                End If
                strSDTType = Trim(txtSDTType.Text)
                dblSDT_Per = CDbl(lblSDTax_Per.Text)
            Else
                strSDTType = ""
                dblSDT_Per = CDbl("0.00")
            End If
        End If
        If blnISSurChargeTaxRoundOff Then
            ldblTotalSurchargeTaxAmount = System.Math.Round(CalculateSurchargeTaxValue(ldblTotalSaleTaxAmount))
        Else
            ldblTotalSurchargeTaxAmount = CalculateSurchargeTaxValue(ldblTotalSaleTaxAmount)
        End If
        If blnAddCustMatrl Then
            ldblTotalInvoiceValue = ldblTotalBasicValue + dblTotalPacking_Amount + ldblExciseValueForSaleTax + dblSDT_Amt + ldblTotalSaleTaxAmount + ldblTotalSurchargeTaxAmount + ldblTotalECESSAmount + ldblTotalSECESSAmount + System.Math.Round(Val(txtFreight.Text)) + System.Math.Round(ldblTotalOthersValues) + System.Math.Round(Val(ctlInsurance.Text)) + System.Math.Round(ldblTotalCustMatrlValue) + dblAddVATamount
        Else
            ldblTotalInvoiceValue = ldblTotalBasicValue + dblTotalPacking_Amount + ldblExciseValueForSaleTax + dblSDT_Amt + ldblTotalSaleTaxAmount + ldblTotalSurchargeTaxAmount + ldblTotalECESSAmount + ldblTotalSECESSAmount + System.Math.Round(Val(txtFreight.Text)) + System.Math.Round(ldblTotalOthersValues) + System.Math.Round(Val(ctlInsurance.Text)) + dblAddVATamount
        End If
        If blnTotalInvoiceAmount Then
            ldblTotalInvoiceValueRoundOff = ldblTotalInvoiceValue - System.Math.Round(ldblTotalInvoiceValue)
            ldblTotalInvoiceValue = System.Math.Round(ldblTotalInvoiceValue)
        Else
            ldblTotalInvoiceValueRoundOff = ldblTotalInvoiceValue - System.Math.Round(ldblTotalInvoiceValue, intTotalInvoiceAmountRoundOffDecimal)
            ldblTotalInvoiceValue = System.Math.Round(ldblTotalInvoiceValue, intTotalInvoiceAmountRoundOffDecimal)
        End If
        Dim strStock_Loc As String
        strStock_Loc = StockLocationSalesConf((Me.CmbInvType.Text), (Me.CmbInvSubType).Text, "DESCRIPTION")
        Select Case Button
            Case "ADD"
                rsSaleConf = New ClsResultSetDB
                rsSaleConf.GetResult("Select Invoice_Type,Sub_Type from SaleConf where  UNIT_CODE = '" & gstrUNITID & "'  and  Description ='" & Trim(CmbInvType.Text) & "' and Sub_type_Description ='" & Trim(CmbInvSubType.Text) & "' and (fin_start_date <= getdate() and fin_end_date >= getdate())")
                strSalesChallan = ""
                If UCase(CmbInvType.Text) <> "JOBWORK INVOICE" Then
                    mstrRGP = ""
                End If
                If (UCase(CmbInvType.Text) = "CSM INVOICE") And chkFOC.CheckState Then
                    ldblTotalInvoiceValue = ldblTotalInvoiceValue - ldblTotalBasicValue
                End If
                strSalesChallan = "INSERT INTO SalesChallan_dtl (UNIT_CODE,Location_Code,Packing_amount,Doc_No,Suffix,Transport_Type,Vehicle_No,From_Location,"
                strSalesChallan = strSalesChallan & "From_Station,To_Station,Account_Code,"
                strSalesChallan = strSalesChallan & "Cust_Ref,Amendment_No,Bill_Flag,Form3,Carriage_Name,"
                strSalesChallan = strSalesChallan & "Year,Insurance,invoice_Type,Ref_Doc_No,"
                strSalesChallan = strSalesChallan & "Cust_Name ,Sales_Tax_Amount , Surcharge_Sales_Tax_Amount,"
                strSalesChallan = strSalesChallan & "Frieght_Amount,Sub_Category,SalesTax_Type,SalesTax_FormNo,"
                strSalesChallan = strSalesChallan & "SalesTax_FormValue,Annex_no,invoice_Date,Currency_code,Ent_dt,"
                strSalesChallan = strSalesChallan & "Ent_UserId,Upd_dt,Upd_UserId,Exchange_Rate,total_amount,"
                strSalesChallan = strSalesChallan & "Surcharge_salesTaxType,SalesTax_Per,Surcharge_SalesTax_Per,PerValue,"
                strSalesChallan = strSalesChallan & "Remarks,SRVDINO,SRVLocation,ECESS_Type,ECESS_Per,ECESS_Amount,SECESS_Type,SECESS_Per,SECESS_Amount"
                If UCase(CmbInvType.Text) = "JOBWORK INVOICE" Then
                    strSalesChallan = strSalesChallan & ",FIFO_Flag "
                End If
                strSalesChallan = strSalesChallan & ",USLOC,SchTime,TotalInvoiceAmtRoundOff_diff"
                strSalesChallan = strSalesChallan & ",Payment_Terms"
                If UCase(CmbInvType.Text) <> "JOBWORK INVOICE" And UCase(CmbInvType.Text) <> "TRANSFER INVOICE" Then
                    strSalesChallan = strSalesChallan & ", SDTax_Type, SDTax_Per, SDTax_Amount "
                End If
                strSalesChallan = strSalesChallan & ",ADDVAT_Type,ADDVAT_Per,ADDVAT_Amount"
                strSalesChallan = strSalesChallan & ",FOC_Invoice)"
                strSalesChallan = strSalesChallan & " Values ('" & gstrUNITID & "','" & Trim(txtLocationCode.Text)
                strSalesChallan = strSalesChallan & "', " & dblTotalPacking_Amount & ","
                strSalesChallan = strSalesChallan & "'" & Trim(txtChallanNo.Text) & "',''"
                strSalesChallan = strSalesChallan & ",'" & Mid(Trim(CmbTransType.Text), 1, 1) & "', '" & Trim(txtVehNo.Text) & "','" & Trim(strStock_Loc) & "','"
                strSalesChallan = strSalesChallan & "','','" & Trim(txtCustCode.Text)
                strSalesChallan = strSalesChallan & "','" & Trim(txtRefNo.Text) & "','" & Trim(txtAmendNo.Text) & "','0'"
                strSalesChallan = strSalesChallan & ",'','" & Trim(txtCarrServices.Text)
                strSalesChallan = strSalesChallan & "','" & Trim(CStr(Year(dtpDateDesc.Value))) & "',"
                strSalesChallan = strSalesChallan & System.Math.Round(Val(ctlInsurance.Text)) & ",'" & Trim(rsSaleConf.GetValue("Invoice_type")) & "','"
                strSalesChallan = strSalesChallan & Trim(mstrRGP) & "','"
                strSalesChallan = strSalesChallan & Trim(lblCustCodeDes.Text) & "',"
                strSalesChallan = strSalesChallan & Val(CStr(ldblTotalSaleTaxAmount)) & "," & Val(CStr(ldblTotalSurchargeTaxAmount)) & "," & System.Math.Round(Val(txtFreight.Text)) & ",'" & Trim(rsSaleConf.GetValue("Sub_Type")) & "','"
                strSalesChallan = strSalesChallan & Trim(txtSaleTaxType.Text) & "','"
                strSalesChallan = strSalesChallan & "0',0,'0','"
                strSalesChallan = strSalesChallan & VB6.Format(dtpDateDesc.Value, "dd/mmm/yyyy") & "','" & lblCurrencyDes.Text & "',getdate(),'" & mP_User & "',  getdate() ,'" & mP_User & "','"
                strSalesChallan = strSalesChallan & lblExchangeRateValue.Text & "'," & ldblTotalInvoiceValue & ",'"
                strSalesChallan = strSalesChallan & Trim(txtSurchargeTaxType.Text) & "'," & Val(lblSaltax_Per.Text) & ","
                strSalesChallan = strSalesChallan & Val(lblSurcharge_Per.Text) & "," & ctlPerValue.Text & ",'" & txtRemarks.Text & "','"
                strSalesChallan = strSalesChallan & Trim(txtSRVDI.Text) & "','" & Trim(txtSRVLoc.Text) & "','"
                strSalesChallan = strSalesChallan & Trim(txtECESS.Text) & "'," & Val(lblECESS_Per.Text) & "," & Val(CStr(ldblTotalECESSAmount)) & ",'"
                strSalesChallan = strSalesChallan & Trim(txtSECESS.Text) & "'," & Val(lblSECESS_Per.Text) & "," & Val(CStr(ldblTotalSECESSAmount))
                If UCase(CmbInvType.Text) = "JOBWORK INVOICE" Then
                    If blnFIFO = True Then
                        strSalesChallan = strSalesChallan & ",1"
                    Else
                        strSalesChallan = strSalesChallan & ",0"
                    End If
                End If
                strSalesChallan = strSalesChallan & ",'" & txtUsLoc.Text & "','" & txtSchTime.Text & "'"
                strSalesChallan = strSalesChallan & "," & ldblTotalInvoiceValueRoundOff
                strSalesChallan = strSalesChallan & ",'" & Trim(lblCreditTerm.Text) & "'"
                If UCase(CmbInvType.Text) <> "JOBWORK INVOICE" And UCase(CmbInvType.Text) <> "TRANSFER INVOICE" Then
                    strSalesChallan = strSalesChallan & ", '" & Trim(strSDTType) & "', " & dblSDT_Per & ", " & dblSDT_Amt
                End If
                strSalesChallan = strSalesChallan & ", '" & Trim(txtAddVAT.Text) & "', " & Val(lblAddVAT.Text) & ", " & dblAddVATamount
                strSalesChallan = strSalesChallan & "," & IIf(chkFOC.Checked, 1, 0) & " )"
                rsSaleConf.ResultSetClose()
                rsSaleConf = Nothing
                strSalesDtl = ""
                With SpChEntry
                    For lintLoopCounter = 1 To .MaxRows
                        .Row = lintLoopCounter
                        .Col = 1
                        lstrItemCode = Trim(.Text)
                        .Col = 2
                        lstrItemDrgno = Trim(.Text)
                        .Col = 3
                        ldblItemRate = Val(.Text) / CDbl(ctlPerValue.Text)
                        .Col = 4
                        ldblItemCustMtrl = Val(.Text) / CDbl(ctlPerValue.Text)
                        .Col = 5
                        lintItemQuantity = Val(.Text)
                        .Col = 6
                        strPackingCode = Trim(.Text)
                        rsPacking_Tax = New ClsResultSetDB
                        rsPacking_Tax.GetResult("Select Txrt_Rate_no,TxRt_Percentage from Gen_TaxRate where Tx_TaxeID = 'PKT' and Txrt_Rate_no = '" & Trim(strPackingCode) & "' and UNIT_CODE = '" & gstrUNITID & "'")
                        If rsPacking_Tax.GetNoRows > 0 Then
                            ldblItemPacking = rsPacking_Tax.GetValue("TxRt_Percentage")
                        End If
                        rsPacking_Tax.ResultSetClose()
                        .Col = 7
                        lstrItemExciseCode = Trim(.Text)
                        .Col = 8
                        lstrItemCVDCode = Trim(.Text)
                        .Col = 9
                        lstrItemSADCode = Trim(.Text)
                        .Col = 10
                        ldblItemOthers = Val(.Text) / CDbl(ctlPerValue.Text) * lintItemQuantity
                        .Col = 11
                        ldblItemFromBox = Val(.Text)
                        .Col = 12
                        ldblItemToBox = Val(.Text)
                        .Col = 14
                        lstrItemDelete = Trim(.Text)
                        .Col = 22
                        dblBinQty = Val(.Text)
                        If dblBinQty <= 0 Then
                            MsgBox("Bin Quantity can't be zero.", MsgBoxStyle.Information, "eMpro")
                            SaveData = False
                            Exit Function
                        End If
                        .Col = 15
                        ldblItemToolCost = Val(.Text) / CDbl(ctlPerValue.Text)
                        If Val(CStr(ldblItemCustMtrl)) > 0 Then
                            strQry = ""
                            strQry = "Select Cust_Mtrl from Cust_ord_dtl WHERE "
                            strQry = strQry & "Account_Code ='" & txtCustCode.Text & "'and Cust_ref ='"
                            strQry = strQry & txtRefNo.Text & "' and Amendment_No = '" & Trim(txtAmendNo.Text) & "'and "
                            strQry = strQry & " Active_flag ='A' "
                            strQry = strQry & " and Cust_DrgNo = '" & Trim(lstrItemDrgno) & "'"
                            strQry = strQry & " and Item_Code ='" & Trim(lstrItemCode) & "' and UNIT_CODE = '" & gstrUNITID & "'"
                            rsCustOrdDtl = New ClsResultSetDB
                            rsCustOrdDtl.GetResult(strQry)
                            If rsCustOrdDtl.GetNoRows > 0 Then
                                dblCustMtrl_SO = rsCustOrdDtl.GetValue("Cust_Mtrl")
                            End If
                            If Val(CStr(dblCustMtrl_SO)) = 0 Then
                                If Val(CStr(ldblItemCustMtrl)) > 0 Then
                                    ldblItemCustMtrl = 0
                                End If
                            End If
                            rsCustOrdDtl.ResultSetClose()
                            rsCustOrdDtl = Nothing
                        End If
                        rsCustItemMst = New ClsResultSetDB
                        rsItemMst = New ClsResultSetDB
                        rsItemMst.GetResult("SELECT Description FROM Item_Mst WHERE Item_Code ='" & Trim(lstrItemCode) & "' and UNIT_CODE = '" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                        rsCustItemMst.GetResult("SELECT Drg_desc FROM CustItem_Mst WHERE Account_code ='" & Trim(txtCustCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "'  and Cust_DrgNo='" & lstrItemDrgno & "'and Item_code ='" & lstrItemCode & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        If UCase(Trim(lstrItemDelete)) <> "D" Then
                            strSalesDtl = Trim(strSalesDtl) & "INSERT INTO sales_Dtl(UNIT_CODE,Cust_Ref,Packing_Type,BinQuantity,Amendment_No,Location_Code,Doc_No,Suffix,Item_Code,Sales_Quantity,"
                            strSalesDtl = strSalesDtl & "From_Box,To_Box,Rate,Sales_Tax,Excise_Tax,Packing,ItemPacking_Amount,Others,Cust_Mtrl,"
                            strSalesDtl = strSalesDtl & "Year,Cust_Item_Code,Cust_Item_Desc,Tool_Cost,Measure_Code,Excise_type,SalesTax_type,CVD_type,SAD_type,Basic_Amount,Accessible_amount,CVD_Amount,SVD_amount,"
                            strSalesDtl = strSalesDtl & "Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,Excise_per,CVD_per,SVD_per,CustMtrl_Amount,ToolCost_Amount,TotalExciseAmount) values ('" & gstrUNITID & "','" & Trim(txtRefNo.Text) & "','" & Trim(strPackingCode) & "' ," & dblBinQty & ",'" & Trim(txtAmendNo.Text) & "','" & Trim(txtLocationCode.Text) & "','"
                            strSalesDtl = strSalesDtl & Trim(txtChallanNo.Text) & "','','" & Trim(lstrItemCode) & "','" & Val(CStr(lintItemQuantity)) & "','"
                            strSalesDtl = strSalesDtl & Val(CStr(ldblItemFromBox)) & "','" & Val(CStr(ldblItemToBox)) & "'," & Val(CStr(ldblItemRate)) & "," & Trim(lblSaltax_Per.Text) & ","
                            TempAccessibleVal = CalculateAccessibleValue(lintLoopCounter, ldblNetInsurenceValue, blnISInsExcisable)
                            dblItemPacking_Amount = CalculatePackingValue(lintLoopCounter, blnPackingRoundoff)
                            If blnISExciseRoundOff Then
                                '10736222
                                dblExcise_Amount = System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal + dblItemPacking_Amount, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff))
                                strSalesDtl = strSalesDtl & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal + dblItemPacking_Amount, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff))
                            Else
                                '10736222
                                dblExcise_Amount = CalculateExciseValue(lintLoopCounter, TempAccessibleVal + dblItemPacking_Amount, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff)
                                strSalesDtl = strSalesDtl & CalculateExciseValue(lintLoopCounter, TempAccessibleVal + dblItemPacking_Amount, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff)
                            End If
                            strSalesDtl = strSalesDtl & "," & Val(CStr(ldblItemPacking)) & "," & Val(CStr(dblItemPacking_Amount)) & "," & Val(CStr(ldblItemOthers)) & "," & Val(CStr(ldblItemCustMtrl)) & ",'"
                            strSalesDtl = strSalesDtl & Trim(CStr(Year(dtpDateDesc.Value))) & "','" & Trim(lstrItemDrgno) & "','" & IIf((Len(Trim(rsCustItemMst.GetValue("Drg_Desc"))) <= 0 Or Trim(CStr(rsCustItemMst.GetValue("Drg_Desc") = "Unknown"))), Trim(rsItemMst.GetValue("Description")), Trim(rsCustItemMst.GetValue("Drg_Desc"))) & "',"
                            'If UCase(CmbInvType.Text) = "NORMAL INVOICE" Or UCase(CmbInvType.Text) = "EXPORT INVOICE" Then
                            If UCase(CmbInvType.Text) = "NORMAL INVOICE" Or UCase(CmbInvType.Text) = "EXPORT INVOICE" Or UCase(CmbInvType.Text) = "TRANSFER INVOICE" Then
                                If UCase(CmbInvSubType.Text) <> "SCRAP" Then
                                    strSalesDtl = strSalesDtl & ldblItemToolCost & ",'','"
                                Else
                                    strSalesDtl = strSalesDtl & ldblItemToolCost & ",'','"
                                End If
                            Else
                                strSalesDtl = strSalesDtl & ldblItemToolCost & ",'','"
                            End If
                            strSalesDtl = strSalesDtl & lstrItemExciseCode & "','" & Trim(txtSaleTaxType.Text) & "','" & lstrItemCVDCode & "','" & lstrItemSADCode & "',"
                            strSalesDtl = strSalesDtl & CalculateBasicValue(lintLoopCounter, blnISBasicRoundOff) & ","
                            strSalesDtl = strSalesDtl & TempAccessibleVal & ","
                            If blnISExciseRoundOff Then
                                strSalesDtl = strSalesDtl & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_CVD, blnEOUFlag, blnISExciseRoundOff))
                                strSalesDtl = strSalesDtl & "," & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_SAD, blnEOUFlag, blnISExciseRoundOff))
                            Else
                                strSalesDtl = strSalesDtl & (CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_CVD, blnEOUFlag, blnISExciseRoundOff))
                                strSalesDtl = strSalesDtl & "," & (CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_SAD, blnEOUFlag, blnISExciseRoundOff))
                            End If
                            strSalesDtl = strSalesDtl & ",GetDate(),'"
                            strSalesDtl = strSalesDtl & Trim(mP_User) & "', GetDate(),'" & Trim(mP_User) & "'," & GetTaxRate(lstrItemExciseCode, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='EXC' ") & "," & GetTaxRate(lstrItemCVDCode, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='CVD'") & "," & GetTaxRate(lstrItemSADCode, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='SAD'") & "," & System.Math.Round(Val(CStr(lintItemQuantity * ldblItemCustMtrl)), 2) & "," & System.Math.Round(Val(CStr(lintItemQuantity * ldblItemToolCost)), 2) & ","
                            If blnISExciseRoundOff Then
                                strSalesDtl = strSalesDtl & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal + dblItemPacking_Amount, enumExciseType.RETURN_ALLExcise, blnEOUFlag, blnISExciseRoundOff)) & ")"
                            Else
                                strSalesDtl = strSalesDtl & (CalculateExciseValue(lintLoopCounter, TempAccessibleVal + dblItemPacking_Amount, enumExciseType.RETURN_ALLExcise, blnEOUFlag, blnISExciseRoundOff)) & " )"
                            End If

                            '10736222
                            strSql = "select dbo.UDF_ISCT2INVOICE( '" & gstrUNITID & "','" & txtCustCode.Text.Trim & "','" & CmbInvType.Text.Trim & "','" & CmbInvSubType.Text.Trim & "','" & txtRefNo.Text.Trim & "')"
                            If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strsql)) = True Then
                                blnIsCt2 = True
                                strSqlct2qry = "INSERT INTO TMP_CT2_INVOICE_KNOCKOFF ([UNIT_CODE],[CUST_CODE],[SONO],[AMENDMENT_NO],[TMP_INVOICE_NO],[ITEM_CODE],[CUST_DRG_NO],[CURRENCY_CODE],[QTY],[RATE],[TOOL_COST],[EXCISE_TAX],[EXCISE_AMOUNT],[ECESS_TYPE],[SECESS_TYPE],[IP_ADDRESS]) "
                                strSqlct2qry = strSqlct2qry + " Values('" & gstrUNITID & "','" & txtCustCode.Text.Trim & "','" & txtRefNo.Text.Trim & "','" & txtAmendNo.Text.Trim & "','" & Me.txtChallanNo.Text.Trim & "',"
                                strSqlct2qry = strSqlct2qry + "'" & lstrItemCode.Trim & "','" & lstrItemDrgno.Trim & "','" & lblCurrencyDes.Text.Trim & "'," & Val(CStr(lintItemQuantity)) & "," & Val(CStr(ldblItemRate)) & "," & Val(ldblItemToolCost) & ",'" & lstrItemExciseCode.Trim & "'," & dblExcise_Amount & ",'" & txtECESS.Text.Trim & "','" & txtSECESS.Text.Trim & "','" & gstrIpaddressWinSck & "' ) "
                                SqlConnectionclass.ExecuteNonQuery(strSqlct2qry)
                            End If
                            '10736222

                        End If
                        rsItemMst.ResultSetClose()
                        rsCustItemMst.ResultSetClose()
                    Next
                End With
            Case "EDIT"
                strSalesChallan = ""
                strSalesChallan = "UPDATE SalesChallan_Dtl SET Insurance = " & System.Math.Round(Val(ctlInsurance.Text))
                If blnISSalesTaxRoundOff Then
                    strSalesChallan = strSalesChallan & ",Sales_Tax_Amount =" & System.Math.Round(Val(CStr(ldblTotalSaleTaxAmount)))
                    strSalesChallan = strSalesChallan & ",ADDVAT_Type='" & txtAddVAT.Text.Trim() & "',AddVat_Per=" & Val(lblAddVAT.Text) & ",ADDVAT_Amount =" & System.Math.Round(Val(CStr(dblAddVATamount)))
                Else
                    strSalesChallan = strSalesChallan & ",Sales_Tax_Amount =" & Val(CStr(ldblTotalSaleTaxAmount))
                    strSalesChallan = strSalesChallan & ",ADDVAT_Type='" & txtAddVAT.Text.Trim() & "',AddVat_Per=" & Val(lblAddVAT.Text) & ",ADDVAT_Amount =" & Val(CStr(dblAddVATamount))
                End If
                If blnISECESSRoundoff Then
                    strSalesChallan = strSalesChallan & ",ECESS_Amount =" & System.Math.Round(Val(CStr(ldblTotalECESSAmount)))
                    strSalesChallan = strSalesChallan & ",SECESS_Amount =" & System.Math.Round(Val(CStr(ldblTotalSECESSAmount)))
                Else
                    strSalesChallan = strSalesChallan & ",ECESS_Amount =" & Val(CStr(ldblTotalECESSAmount))
                    strSalesChallan = strSalesChallan & ",SECESS_Amount =" & Val(CStr(ldblTotalSECESSAmount))
                End If
                If blnISSurChargeTaxRoundOff Then
                    strSalesChallan = strSalesChallan & ",Surcharge_Sales_Tax_Amount =" & System.Math.Round(Val(CStr(ldblTotalSurchargeTaxAmount)))
                Else
                    strSalesChallan = strSalesChallan & ",Surcharge_Sales_Tax_Amount =" & Val(CStr(ldblTotalSurchargeTaxAmount))
                End If
                If UCase(mstrInvType) <> "JOB" And UCase(mstrInvType) <> "TRF" Then
                    strSalesChallan = strSalesChallan & ", SDTax_Type='" & strSDTType & "', SDTax_Per=" & dblSDT_Per & ", SDTax_Amount = " & dblSDT_Amt
                End If
                strSalesChallan = strSalesChallan & ",Frieght_Amount=" & System.Math.Round(Val(txtFreight.Text))
                strSalesChallan = strSalesChallan & ",SalesTax_Type='" & Trim(txtSaleTaxType.Text) & "'"
                strSalesChallan = strSalesChallan & ",total_amount=" & ldblTotalInvoiceValue
                strSalesChallan = strSalesChallan & ",Packing_amount=" & dblTotalPacking_Amount
                strSalesChallan = strSalesChallan & ",Surcharge_salesTaxType='" & Trim(txtSurchargeTaxType.Text) & "'"
                strSalesChallan = strSalesChallan & ",SalesTax_Per=" & Val(lblSaltax_Per.Text)
                strSalesChallan = strSalesChallan & ",Surcharge_SalesTax_Per=" & Val(lblSurcharge_Per.Text)
                strSalesChallan = strSalesChallan & ",PerValue=" & ctlPerValue.Text & ",Remarks = '" & txtRemarks.Text & "' "
                strSalesChallan = strSalesChallan & ",SRVDINO = '" & Trim(txtSRVDI.Text) & "',"
                strSalesChallan = strSalesChallan & " SRVLocation = '" & Trim(txtSRVLoc.Text)
                strSalesChallan = strSalesChallan & "',USLOC = '" & Trim(txtUsLoc.Text) & "',"
                strSalesChallan = strSalesChallan & " schTime = '" & Trim(txtSchTime.Text) & "'"
                strSalesChallan = strSalesChallan & ",ECESS_Type = '" & Trim(txtECESS.Text) & "',"
                strSalesChallan = strSalesChallan & "ECESS_Per = " & Val(lblECESS_Per.Text) & ","
                strSalesChallan = strSalesChallan & "PAYMENT_TERMS = '" & Trim(lblCreditTerm.Text) & "',"
                strSalesChallan = strSalesChallan & "SECESS_Type = '" & Trim(txtSECESS.Text) & "',"
                strSalesChallan = strSalesChallan & "SECESS_Per = " & Val(lblSECESS_Per.Text) & ","
                strSalesChallan = strSalesChallan & "TotalInvoiceAmtRoundOff_diff = " & ldblTotalInvoiceValueRoundOff
                strSalesChallan = strSalesChallan & " WHERE  UNIT_CODE = '" & gstrUNITID & "' and Location_Code ='" & Trim(txtLocationCode.Text) & "'"
                strSalesChallan = strSalesChallan & " and Doc_No ='" & Val(txtChallanNo.Text) & "'"
                strSalesDtl = ""
                strSalesDtlDelete = ""
                With SpChEntry
                    For lintLoopCounter = 1 To .MaxRows
                        .Row = lintLoopCounter
                        .Col = 1
                        lstrItemCode = Trim(.Text)
                        .Col = 3
                        ldblItemRate = Val(.Text) / CDbl(ctlPerValue.Text)
                        .Row = lintLoopCounter
                        .Col = 5
                        lintItemQuantity = Val(.Text)
                        .Col = 6
                        strPackingCode = Trim(.Text)
                        rsPacking_Tax = New ClsResultSetDB
                        rsPacking_Tax.GetResult("Select Txrt_Rate_no,TxRt_Percentage from Gen_TaxRate where  UNIT_CODE = '" & gstrUNITID & "' and Tx_TaxeID = 'PKT' and Txrt_Rate_no = '" & Trim(strPackingCode) & "'")
                        If rsPacking_Tax.GetNoRows > 0 Then
                            ldblItemPacking = rsPacking_Tax.GetValue("TxRt_Percentage")
                        End If
                        rsPacking_Tax.ResultSetClose()
                        .Col = 2
                        lstrItemDrgno = Trim(.Text)
                        .Col = 14
                        lstrItemDelete = Trim(.Text)
                        .Col = 15
                        ldblItemToolCost = Val(.Text) / CDbl(ctlPerValue.Text)
                        .Col = 7
                        lstrItemExciseCode = Trim(.Text)
                        .Col = 8
                        lstrItemCVDCode = Trim(.Text)
                        .Col = 9
                        lstrItemSADCode = Trim(.Text)
                        .Col = 15
                        ldblItemToolCost = Val(.Text)
                        .Col = 22
                        dblBinQty = Val(.Text)
                        If dblBinQty <= 0 Then
                            MsgBox("Bin Quantity can't be zero.", MsgBoxStyle.Information, "eMpro")
                            SaveData = False
                            Exit Function
                        End If
                        .Col = 11
                        ldblItemFromBox = Val(.Text)
                        .Col = 12
                        ldblItemToBox = Val(.Text)
                        If Val(CStr(ldblItemCustMtrl)) > 0 Then
                            strQry = ""
                            strQry = "Select Cust_Mtrl from Cust_ord_dtl WHERE "
                            strQry = strQry & "Account_Code ='" & txtCustCode.Text & "'and Cust_ref ='"
                            strQry = strQry & txtRefNo.Text & "' and Amendment_No = '" & Trim(txtAmendNo.Text) & "'and "
                            strQry = strQry & " Active_flag ='A' "
                            strQry = strQry & " and Cust_DrgNo = '" & Trim(lstrItemDrgno) & "'"
                            strQry = strQry & " and Item_Code ='" & Trim(lstrItemCode) & "' and UNIT_CODE = '" & gstrUNITID & "'"
                            rsCustOrdDtl = New ClsResultSetDB
                            rsCustOrdDtl.GetResult(strQry)
                            If rsCustOrdDtl.GetNoRows > 0 Then
                                dblCustMtrl_SO = rsCustOrdDtl.GetValue("Cust_Mtrl")
                            End If
                            If Val(CStr(dblCustMtrl_SO)) = 0 Then
                                If Val(CStr(ldblItemCustMtrl)) > 0 Then
                                    ldblItemCustMtrl = 0
                                End If
                            End If
                            rsCustOrdDtl.ResultSetClose()
                            rsCustOrdDtl = Nothing
                        End If

                        
                        If UCase(lstrItemDelete) <> "D" Then
                            If UCase(lstrItemDelete) <> "A" Then
                                strSalesDtl = Trim(strSalesDtl) & "UPDATE Sales_dtl SET Sales_Quantity ='" & Val(CStr(lintItemQuantity)) & "',BinQuantity=" & dblBinQty & " ,Sales_Tax =" & Trim(lblSaltax_Per.Text) & ","
                                strSalesDtl = Trim(strSalesDtl) & " TOOL_COST = " & ldblItemToolCost & ","
                                strSalesDtl = Trim(strSalesDtl) & "CustMtrl_Amount= " & Val(CStr(lintItemQuantity * ldblItemCustMtrl)) & ",ToolCost_Amount=" & Val(CStr(lintItemQuantity * ldblItemToolCost))
                                TempAccessibleVal = CalculateAccessibleValue(lintLoopCounter, ldblNetInsurenceValue, blnISInsExcisable)
                                dblItemPacking_Amount = CalculatePackingValue(lintLoopCounter, blnPackingRoundoff)
                                If blnISExciseRoundOff Then
                                    '10736222
                                    dblExcise_Amount = System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal + dblItemPacking_Amount, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff))
                                    strSalesDtl = Trim(strSalesDtl) & ",Excise_Tax=" & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal + dblItemPacking_Amount, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff))
                                Else
                                    dblExcise_Amount = CalculateExciseValue(lintLoopCounter, TempAccessibleVal + dblItemPacking_Amount, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff)
                                    strSalesDtl = Trim(strSalesDtl) & ",Excise_Tax=" & CalculateExciseValue(lintLoopCounter, TempAccessibleVal + dblItemPacking_Amount, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff)
                                End If
                                If blnISExciseRoundOff Then
                                    strSalesDtl = Trim(strSalesDtl) & ",TotalExciseAmount =" & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal + dblItemPacking_Amount, enumExciseType.RETURN_ALLExcise, blnEOUFlag, blnISExciseRoundOff))
                                Else
                                    strSalesDtl = Trim(strSalesDtl) & ",TotalExciseAmount =" & CalculateExciseValue(lintLoopCounter, TempAccessibleVal + dblItemPacking_Amount, enumExciseType.RETURN_ALLExcise, blnEOUFlag, blnISExciseRoundOff)
                                End If
                                strSalesDtl = Trim(strSalesDtl) & ",Excise_type='" & lstrItemExciseCode & "',SalesTax_type='" & Trim(txtSaleTaxType.Text) & "'"
                                strSalesDtl = Trim(strSalesDtl) & ", Packing=" & Val(CStr(ldblItemPacking)) & ",ItemPacking_Amount=" & Val(CStr(dblItemPacking_Amount)) & ""
                                strSalesDtl = Trim(strSalesDtl) & ",CVD_type='" & Trim(lstrItemCVDCode) & "',SAD_type='" & Trim(lstrItemSADCode) & "',Basic_Amount=" & CalculateBasicValue(lintLoopCounter, blnISBasicRoundOff)
                                strSalesDtl = Trim(strSalesDtl) & ",Accessible_amount=" & Val(CStr(TempAccessibleVal))
                                If blnISExciseRoundOff Then
                                    strSalesDtl = Trim(strSalesDtl) & ",CVD_Amount=" & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_CVD, blnEOUFlag, blnISExciseRoundOff)) & ",SVD_amount=" & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_SAD, blnEOUFlag, blnISExciseRoundOff))
                                Else
                                    strSalesDtl = Trim(strSalesDtl) & ",CVD_Amount=" & CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_CVD, blnEOUFlag, blnISExciseRoundOff) & ",SVD_amount=" & CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_SAD, blnEOUFlag, blnISExciseRoundOff)
                                End If
                                strSalesDtl = Trim(strSalesDtl) & ",Excise_per=" & GetTaxRate(lstrItemExciseCode, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='EXC'")
                                strSalesDtl = Trim(strSalesDtl) & ",CVD_per=" & GetTaxRate(lstrItemCVDCode, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='CVD'")
                                strSalesDtl = Trim(strSalesDtl) & ",SVD_per=" & GetTaxRate(lstrItemSADCode, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='SAD'")
                                strSalesDtl = Trim(strSalesDtl) & ",Rate=" & ldblItemRate & ",FROM_BOX = " & ldblItemFromBox & ",To_box = " & ldblItemToBox
                                strSalesDtl = Trim(strSalesDtl) & " WHERE  UNIT_CODE = '" & gstrUNITID & "' and Location_Code ='" & Trim(txtLocationCode.Text) & "'"
                                strSalesDtl = Trim(strSalesDtl) & " and Doc_No =" & Val(txtChallanNo.Text) & " and Cust_Item_Code='"
                                strSalesDtl = Trim(strSalesDtl) & Trim(lstrItemDrgno) & "'" & vbCrLf
                            ElseIf UCase(lstrItemDelete) = "A" Then
                                strSalesDtl = strSalesDtl & vbCrLf & InsertinSalesDtlinEditMode(lintLoopCounter)
                            End If
                        Else
                            strSalesDtlDelete = Trim(strSalesDtlDelete) & "DELETE Sales_dtl "
                            strSalesDtlDelete = Trim(strSalesDtlDelete) & " WHERE  UNIT_CODE = '" & gstrUNITID & "'  and Location_Code ='" & Trim(txtLocationCode.Text) & "'"
                            strSalesDtlDelete = Trim(strSalesDtlDelete) & " and Doc_No =" & Val(txtChallanNo.Text) & " and Cust_Item_Code='"
                            strSalesDtlDelete = Trim(strSalesDtlDelete) & Trim(lstrItemDrgno) & "'" & vbCrLf
                        End If
                        '10736222
                        strsql = "select dbo.UDF_ISCT2INVOICE( '" & gstrUNITID & "','" & txtCustCode.Text.Trim & "','" & CmbInvType.Text.Trim & "','" & CmbInvSubType.Text.Trim & "','" & txtRefNo.Text.Trim & "')"
                        If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strsql)) = True Then
                            blnIsCt2 = True
                            strSqlct2qry = "INSERT INTO TMP_CT2_INVOICE_KNOCKOFF ([UNIT_CODE],[CUST_CODE],[SONO],[AMENDMENT_NO],[TMP_INVOICE_NO],[ITEM_CODE],[CUST_DRG_NO],[CURRENCY_CODE],[QTY],[RATE],[TOOL_COST],[EXCISE_TAX],[EXCISE_AMOUNT],[ECESS_TYPE],[SECESS_TYPE],[IP_ADDRESS]) "
                            strSqlct2qry = strSqlct2qry + " Values('" & gstrUNITID & "','" & txtCustCode.Text.Trim & "','" & txtRefNo.Text.Trim & "','" & txtAmendNo.Text.Trim & "','" & Me.txtChallanNo.Text.Trim & "',"
                            strSqlct2qry = strSqlct2qry + "'" & lstrItemCode.Trim & "','" & lstrItemDrgno.Trim & "','" & lblCurrencyDes.Text.Trim & "'," & Val(CStr(lintItemQuantity)) & "," & Val(CStr(ldblItemRate)) & "," & Val(ldblItemToolCost) & ",'" & lstrItemExciseCode.Trim & "'," & dblExcise_Amount & ",'" & txtECESS.Text.Trim & "','" & txtSECESS.Text.Trim & "','" & gstrIpaddressWinSck & "' ) "
                            SqlConnectionclass.ExecuteNonQuery(strSqlct2qry)
                        End If
                        '10736222

                    Next
                End With
        End Select

        If blnIsCt2 = True Then
            '10736222
            Dim objValidateCmd As New ADODB.Command

            With objValidateCmd
                .ActiveConnection = mP_Connection
                .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                .CommandText = "USP_VALIDATE_CT2_INVOICE_KNOCKOFF"
                .CommandTimeout = 0
                .Parameters.Append(.CreateParameter("@MODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, IIf(CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, "A", "E")))
                .Parameters.Append(.CreateParameter("@UNIT_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, gstrUNITID))
                .Parameters.Append(.CreateParameter("@INVOICE_NO", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, , txtChallanNo.Text.Trim))
                .Parameters.Append(.CreateParameter("@IP_ADDRESS", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, gstrIpaddressWinSck))
                .Parameters.Append(.CreateParameter("@ERRMSG", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInputOutput, 8000, ""))
                .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End With

            If objValidateCmd.Parameters(objValidateCmd.Parameters.Count - 1).Value.ToString().Trim.Length <> 0 Then
                MsgBox("Unable To Save CT2 Invoice Knock Off Details." & vbCr & objValidateCmd.Parameters(objValidateCmd.Parameters.Count - 1).Value.ToString(), MsgBoxStyle.Information, ResolveResString(100))
                objValidateCmd = Nothing
                SaveData = False
                Exit Function
            End If
            objValidateCmd = Nothing
            '10736222
        End If

        With mP_Connection
            ResetDatabaseConnection()
            .BeginTrans()
            .Execute(strSalesChallan, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            If Len(Trim(strupSalechallan)) > 0 Then
                .Execute(strupSalechallan, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End If
            .Execute(strSalesDtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            If Len(Trim(mstrUpdDispatchSql)) > 0 Then
                .Execute(mstrUpdDispatchSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End If
            If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                If Len(Trim(strSalesDtlDelete)) > 0 Then
                    .Execute(strSalesDtlDelete, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                End If
            End If
            If blnIsCt2 = True Then
                '10736222
                Dim objCmd As New ADODB.Command

                With objCmd
                    .ActiveConnection = mP_Connection
                    .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                    .CommandText = "USP_SAVE_CT2_INVOICE_KNOCKOFFDTL"
                    .CommandTimeout = 0
                    .Parameters.Append(.CreateParameter("@MODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, IIf(CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, "A", "E")))
                    .Parameters.Append(.CreateParameter("@UNIT_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, gstrUNITID))
                    .Parameters.Append(.CreateParameter("@INVOICE_NO", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, , txtChallanNo.Text.Trim))
                    .Parameters.Append(.CreateParameter("@IP_ADDRESS", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, gstrIpaddressWinSck))
                    .Parameters.Append(.CreateParameter("@ERRMSG", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInputOutput, 8000, ""))
                    .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                End With

                If objCmd.Parameters(objCmd.Parameters.Count - 1).Value.ToString().Trim.Length <> 0 Then
                    MsgBox("Unable To Save CT2 Invoice Knock Off Details.", MsgBoxStyle.Information, ResolveResString(100))
                    objCmd = Nothing
                    mP_Connection.RollbackTrans()
                    SaveData = False
                    Exit Function
                End If
                objCmd = Nothing
                '10736222
            End If

            .CommitTrans()
        End With
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        mP_Connection.RollbackTrans()
        SaveData = False
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function CalculateStateDevelopmentTaxValue(ByVal pdblTotalBasicValue As Double, ByVal pdblTotalExciseValue As Double) As Double
        On Error GoTo ErrHandler
        CalculateStateDevelopmentTaxValue = ((pdblTotalBasicValue + pdblTotalExciseValue) * Val(lblSDTax_Per.Text)) / 100
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function CalculateBasicValue(ByVal pintRowNo As Short, ByVal blnRoundoff As Boolean) As Double
        Dim ldblPkg_Per As Double
        Dim ldblRate As Double
        Dim lintQty As Double
        On Error GoTo ErrHandler
        With SpChEntry
            .Row = pintRowNo
            .Col = 3
            ldblRate = Val(.Text) / Val(ctlPerValue.Text)
            .Col = 6
            ldblPkg_Per = Val(.Text)
            .Col = 5
            lintQty = Val(.Text)
            If blnRoundoff = True Then
                CalculateBasicValue = System.Math.Round(ldblRate * lintQty, 0)
            Else
                CalculateBasicValue = System.Math.Round(ldblRate * lintQty, 2)
            End If
        End With
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function CalculatePackingValue(ByVal pintRowNo As Short, ByVal blnRoundoff As Boolean) As Double
        Dim strPkg_Type As String
        Dim ldblPkg_Per As Double
        Dim ldblRate As Double
        Dim lintQty As Double
        Dim rsTaxRate As ClsResultSetDB
        Dim intPackingRoundoff_Decimal As Short
        On Error GoTo ErrHandler
        With SpChEntry
            .Row = pintRowNo
            .Col = 3
            ldblRate = Val(.Text) / Val(ctlPerValue.Text)
            .Col = 6
            strPkg_Type = Trim(.Text)
            .Col = 5
            lintQty = Val(.Text)
            intPackingRoundoff_Decimal = Val(Find_Value("select PackingRoundoff_Decimal from sales_parameter WHERE  UNIT_CODE = '" & gstrUNITID & "'"))
            rsTaxRate = New ClsResultSetDB
            rsTaxRate.GetResult("Select Txrt_Rate_no,TxRt_Percentage from Gen_TaxRate where Tx_TaxeID = 'PKT' and Txrt_Rate_no = '" & Trim(strPkg_Type) & "' and UNIT_CODE = '" & gstrUNITID & "'")
            If rsTaxRate.GetNoRows > 0 Then
                ldblPkg_Per = rsTaxRate.GetValue("TxRt_Percentage")
            Else
                ldblPkg_Per = 0
            End If
            rsTaxRate.ResultSetClose()
            If blnRoundoff = True Then
                CalculatePackingValue = System.Math.Round((ldblRate * lintQty) * ldblPkg_Per / 100, 0)
            Else
                CalculatePackingValue = System.Math.Round((ldblRate * lintQty) * ldblPkg_Per / 100, intPackingRoundoff_Decimal)
            End If
        End With
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function CalculateAccessibleValue(ByVal pintRowNo As Short, ByVal pdblInsurenceValue As Double, ByVal pblnISInsAdd As Boolean) As Double
        Dim ldblRate As Double
        Dim ldblCustMat As Double
        Dim ldblToolCost As Double
        Dim ldblPkg_Per As Double
        Dim lintQty As Double
        Dim dblMRPValue As Double
        Dim RSAccessibleVal As ClsResultSetDB
        Dim strSQL As String
        On Error GoTo ErrHandler
        With SpChEntry
            .Row = pintRowNo
            .Col = 3
            ldblRate = Val(.Text) / CDbl(ctlPerValue.Text)
            .Col = 6
            ldblPkg_Per = Val(.Text)
            .Col = 5
            lintQty = Val(.Text)
            .Col = 4
            ldblCustMat = Val(.Text) / CDbl(ctlPerValue.Text)
            .Col = 15
            ldblToolCost = Val(.Text) / CDbl(ctlPerValue.Text)
            If CheckSOType(pintRowNo) = "M" Then
                RSAccessibleVal = New ClsResultSetDB
                strSQL = "Select isnull(AccessibleRateforMRP,0) as MRPValue from Cust_Ord_Dtl where  UNIT_CODE = '" & gstrUNITID & "' and  Account_code = '" & Trim(txtCustCode.Text) & "' and Cust_Ref='" & Trim(txtRefNo.Text) & "'"
                strSQL = strSQL & " and Amendment_No='" & Trim(txtAmendNo.Text) & "'"
                .Col = 1
                strSQL = strSQL & " and Item_Code = '" & .Text & "'"
                .Col = 2
                strSQL = strSQL & " and Cust_Drgno = '" & .Text & "' "
                RSAccessibleVal.GetResult(strSQL, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                If RSAccessibleVal.GetNoRows > 0 Then
                    dblMRPValue = RSAccessibleVal.GetValue("MRPValue")
                    CalculateAccessibleValue = System.Math.Round(dblMRPValue * lintQty, 2)
                End If
                RSAccessibleVal.ResultSetClose()
                RSAccessibleVal = Nothing
            Else
                If pblnISInsAdd = True Then
                    CalculateAccessibleValue = System.Math.Round((ldblRate + ldblCustMat + ldblToolCost + pdblInsurenceValue) * lintQty, 2)
                Else
                    CalculateAccessibleValue = System.Math.Round((ldblRate + ldblCustMat + ldblToolCost) * lintQty, 2)
                End If
            End If
        End With
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function CalculateExciseValue(ByVal pintRowNo As Short, ByVal pdblAccessibleValue As Double, ByVal penumTaxType As enumExciseType, ByRef pblnEOU_FLAG As Boolean, ByRef blnExciseFlag As Boolean) As Double
        Dim strTableSql As String 'Declared To Make Select Query
        Dim rsGetTaxRate As ClsResultSetDB
        Dim ldblTaxRate As Double
        Dim ldblCVDRate As Double
        Dim ldblSADRate As Double
        Dim ldblTempTotalExcise As Double
        Dim ldblTempTotalCVD As Double
        Dim ldblTempTotalSAD As Double
        Dim ldblTempAllExcise As Double
        On Error GoTo ErrHandler
        ldblTempTotalExcise = 0
        ldblTempTotalCVD = 0
        ldblTempTotalSAD = 0
        ldblTempAllExcise = 0
        '101188073
        If gblnGSTUnit Then
            CalculateExciseValue = 0
        End If
        '101188073
        With SpChEntry
            .Row = pintRowNo
            .Col = 7
            rsGetTaxRate = New ClsResultSetDB
            strTableSql = "SELECT TxRt_Percentage FROM Gen_TaxRate WHERE TxRt_Rate_No='" & Trim(.Text) & "' and UNIT_CODE = '" & gstrUNITID & "'"
            rsGetTaxRate.GetResult(strTableSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If rsGetTaxRate.GetNoRows > 0 Then
                ldblTaxRate = rsGetTaxRate.GetValue("TxRt_Percentage")
            Else
                ldblTaxRate = 0
            End If
            rsGetTaxRate.ResultSetClose()
            If pblnEOU_FLAG Then
                .Col = 8
                strTableSql = "SELECT TxRt_Percentage FROM Gen_TaxRate WHERE TxRt_Rate_No='" & Trim(.Text) & "' and UNIT_CODE = '" & gstrUNITID & "'"
                rsGetTaxRate = New ClsResultSetDB
                rsGetTaxRate.GetResult(strTableSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                If rsGetTaxRate.GetNoRows > 0 Then
                    ldblCVDRate = rsGetTaxRate.GetValue("TxRt_Percentage")
                Else
                    ldblCVDRate = 0
                End If
                rsGetTaxRate.ResultSetClose()
                .Col = 9
                strTableSql = "SELECT TxRt_Percentage FROM Gen_TaxRate WHERE TxRt_Rate_No='" & Trim(.Text) & "' and UNIT_CODE = '" & gstrUNITID & "'"
                rsGetTaxRate = New ClsResultSetDB
                rsGetTaxRate.GetResult(strTableSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                If rsGetTaxRate.GetNoRows > 0 Then
                    ldblSADRate = rsGetTaxRate.GetValue("TxRt_Percentage")
                Else
                    ldblSADRate = 0
                End If
                rsGetTaxRate.ResultSetClose()
                ldblTempTotalExcise = ((pdblAccessibleValue * ldblTaxRate) / 100)
                If blnExciseFlag = True Then
                    ldblTempTotalExcise = System.Math.Round(ldblTempTotalExcise, 0)
                    ldblTempAllExcise = ldblTempTotalExcise / 2
                End If
                ldblTempTotalCVD = (((ldblTempTotalExcise + pdblAccessibleValue) * ldblCVDRate) / 100)
                If blnExciseFlag = True Then
                    ldblTempTotalCVD = System.Math.Round(ldblTempTotalCVD, 0)
                    ldblTempAllExcise = ldblTempAllExcise + (ldblTempTotalCVD / 2)
                End If
                ldblTempTotalSAD = (((ldblTempTotalCVD + ldblTempTotalExcise + pdblAccessibleValue) * ldblSADRate) / 100)
                If blnExciseFlag = True Then
                    ldblTempTotalSAD = System.Math.Round(ldblTempTotalSAD, 0)
                    ldblTempAllExcise = ldblTempAllExcise + (ldblTempTotalSAD / 2)
                End If
                If blnExciseFlag = True Then
                    ldblTempAllExcise = System.Math.Round(ldblTempAllExcise, 0)
                End If
                If penumTaxType = enumExciseType.RETURN_EXCISE Then
                    CalculateExciseValue = (ldblTempTotalExcise)
                ElseIf penumTaxType = enumExciseType.RETURN_CVD Then
                    CalculateExciseValue = ldblTempTotalCVD
                ElseIf penumTaxType = enumExciseType.RETURN_SAD Then
                    CalculateExciseValue = ldblTempTotalSAD
                Else
                    CalculateExciseValue = ldblTempAllExcise
                End If
            Else
                If penumTaxType = enumExciseType.RETURN_EXCISE Then
                    CalculateExciseValue = ((pdblAccessibleValue * ldblTaxRate) / 100)
                ElseIf penumTaxType = enumExciseType.RETURN_CVD Then
                    CalculateExciseValue = 0
                ElseIf penumTaxType = enumExciseType.RETURN_SAD Then
                    CalculateExciseValue = 0
                Else
                    CalculateExciseValue = ((pdblAccessibleValue * ldblTaxRate) / 100)
                End If
            End If
        End With
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function CalculateSalesTaxValue(ByVal pdblTotalBasicValue As Double, ByVal pdblTotalExciseValue As Double) As Double
        On Error GoTo ErrHandler
        CalculateSalesTaxValue = ((pdblTotalBasicValue + pdblTotalExciseValue) * Val(lblSaltax_Per.Text)) / 100
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function CalculateAddionalSalesTaxValue(ByVal pdblTotalBasicValue As Double, ByVal pdblTotalExciseValue As Double) As Double
        On Error GoTo ErrHandler
        CalculateAddionalSalesTaxValue = ((pdblTotalBasicValue + pdblTotalExciseValue) * Val(lblAddVAT.Text)) / 100
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function CalculateSurchargeTaxValue(ByVal pdblTotalCSTValue As Double) As Double
        On Error GoTo ErrHandler
        CalculateSurchargeTaxValue = (pdblTotalCSTValue * Val(lblSurcharge_Per.Text) / 100)
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function PrepareQueryForShowingExcise(ByVal pblnTarrifCodeReq As Boolean, ByRef pstrItemCode As String) As String
        Dim strSQL As String
        Dim lclsGetTariffCode As ClsResultSetDB
        On Error GoTo ErrHandler
        PrepareQueryForShowingExcise = ""
        If pblnTarrifCodeReq = True Then
            strSQL = "SELECT Tariff_code FROM Item_Mst WHERE Item_Code ='" & pstrItemCode & "' and UNIT_CODE = '" & gstrUNITID & "'"
            lclsGetTariffCode = New ClsResultSetDB
            Call lclsGetTariffCode.GetResult(strSQL, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If lclsGetTariffCode.GetNoRows > 0 Then
                lclsGetTariffCode.ResultSetClose()
                strSQL = "SELECT Excise_duty FROM Tax_Tariff_Mst WHERE Tariff_SubHead='" & lclsGetTariffCode.GetValue("Tariff_code") & "' and UNIT_CODE = '" & gstrUNITID & "'"
                lclsGetTariffCode = New ClsResultSetDB
                Call lclsGetTariffCode.GetResult(strSQL, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                If lclsGetTariffCode.GetNoRows > 0 Then
                    PrepareQueryForShowingExcise = " AND TxRt_Rate_No='" & lclsGetTariffCode.GetValue("Excise_duty") & "'"
                Else
                    lclsGetTariffCode.ResultSetClose()
                End If
            Else
                lclsGetTariffCode.ResultSetClose()
            End If
        Else
            PrepareQueryForShowingExcise = ""
        End If
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Function
    End Function
    Private Function GetBOMCheckFlagValue(ByVal pstrFieldName As String) As Boolean
        Dim strSQL As String
        Dim rsObj As New ADODB.Recordset
        On Error GoTo ErrHandler
        strSQL = ""
        strSQL = "SELECT " & pstrFieldName & " FROM Sales_Parameter  WHERE UNIT_CODE = '" & gstrUNITID & "'"
        If rsObj.State = 1 Then rsObj.Close()
        rsObj.Open(strSQL, mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
        If rsObj.EOF Or rsObj.BOF Then
            MsgBox("No Data define in Sales_Parameter Table", MsgBoxStyle.Information, "eMPro")
            GetBOMCheckFlagValue = False
        Else
            If rsObj.Fields(pstrFieldName).Value Then
                GetBOMCheckFlagValue = True
            Else
                GetBOMCheckFlagValue = False
            End If
        End If
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        GetBOMCheckFlagValue = False
    End Function
    Private Function GetTotalDispatchQuantityFromDailySchedule(ByVal pstrAccountCode As String, ByVal pstrCustomerDrawNo As String, ByVal pstrItemCode As String, ByVal pstrDate As String, ByVal pstrMode As String, ByVal pdblPrevQty As Double) As Double
        Dim strScheduleSql As String
        Dim objRsForSchedule As New ADODB.Recordset
        Dim ldblTotalDispatchQuantity As Double
        Dim ldblTotalScheduleQuantity As Double
        Dim lintLoopCounter As Short
        Dim rsDSTracking As New ClsResultSetDB
        Dim blnDSTracking As Boolean
        Dim blnCalanderDateTrac As Boolean
        Dim blnFutureDateDS As Boolean
        Dim intnoofFutureWorking As Integer

        On Error GoTo ErrHandler
        Call rsDSTracking.GetResult("Select DSWiseTracking,CalendarDateTrac,FUTUREDATE_DS_KNOCKOFF ,NOOFWORKINGDAYS_DS_KNOCKOFF From Sales_parameter  WHERE UNIT_CODE = '" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsDSTracking.RowCount > 0 Then
            blnDSTracking = IIf(IsDBNull(rsDSTracking.GetValue("DSWiseTracking")), False, IIf(rsDSTracking.GetValue("DSwisetracking") = False, False, True))
            blnCalanderDateTrac = rsDSTracking.GetValue("CalendarDateTrac")
            blnFutureDateDS = rsDSTracking.GetValue("FUTUREDATE_DS_KNOCKOFF")
            intnoofFutureWorking = rsDSTracking.GetValue("NOOFWORKINGDAYS_DS_KNOCKOFF")
        End If
        rsDSTracking.ResultSetClose()
        ldblTotalDispatchQuantity = 0
        ldblTotalScheduleQuantity = 0
        mP_Connection.Execute("SET DATEFORMAT 'mdy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        Dim rsCalendarDate As New ADODB.Recordset
        Dim strCalDate As String
        Dim strQuery As String
        If blnCalanderDateTrac Then
            If blnFutureDateDS Then
                strQuery = "select MAX(dt) dt from (select top " & intnoofFutureWorking & "  dt from calendar_mst where  UNIT_CODE = '" & gstrUNITID & "'  and dt > '" & VB6.Format(GetServerDate, "dd/mmm/yyyy") & "' and work_flg<>1 order by dt )a "
                If rsCalendarDate.State = ADODB.ObjectStateEnum.adStateOpen Then rsCalendarDate.Close()
                rsCalendarDate.Open(strQuery, mP_Connection, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockBatchOptimistic)
                If rsCalendarDate.EOF Or rsCalendarDate.BOF Or IsDBNull(rsCalendarDate.Fields("dt").Value) Then
                    MsgBox("Date in Calendar Master is not defined")
                    GetTotalDispatchQuantityFromDailySchedule = -1
                    mP_Connection.Execute("SET DATEFORMAT 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    rsCalendarDate.Close()
                    Exit Function
                Else
                    rsCalendarDate.MoveFirst()
                    strCalDate = VB6.Format(rsCalendarDate.Fields("DT").Value, "dd/mmm/yyyy")
                End If
                rsCalendarDate.Close()
            Else
                strQuery = "select dt from calendar_mst where  UNIT_CODE = '" & gstrUNITID & "'  and dt > '" & VB6.Format(pstrDate, "dd/mmm/yyyy") & "' and work_flg<>1 order by dt"
                If rsCalendarDate.State = ADODB.ObjectStateEnum.adStateOpen Then rsCalendarDate.Close()
                rsCalendarDate.Open(strQuery, mP_Connection, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockBatchOptimistic)
                If rsCalendarDate.EOF Or rsCalendarDate.BOF Or IsDBNull(rsCalendarDate.Fields("dt").Value) Then
                    MsgBox("Date in Calendar Master is not defined")
                    GetTotalDispatchQuantityFromDailySchedule = -1
                    mP_Connection.Execute("SET DATEFORMAT 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    rsCalendarDate.Close()
                    Exit Function
                Else
                    rsCalendarDate.MoveFirst()
                    strCalDate = VB6.Format(rsCalendarDate.Fields("DT").Value, "dd/mmm/yyyy")
                End If
                rsCalendarDate.Close()
            End If
        End If
        If pstrMode = "ADD" Then
            If blnDSTracking = True Then
                If blnCalanderDateTrac Then
                    strScheduleSql = "Select Schedule_Quantity=Sum(Isnull(Schedule_Quantity,0)),Despatch_Qty=Sum(isnull(Despatch_Qty,0)) from DailyMktSchedule where   UNIT_CODE = '" & gstrUNITID & "'  and Account_Code='" & pstrAccountCode & "'"
                    strScheduleSql = strScheduleSql & " and Trans_Date <='" & VB6.Format(strCalDate, "dd/mmm/yyyy") & "'"
                    strScheduleSql = strScheduleSql & " and Cust_DrgNo = '" & pstrCustomerDrawNo & "' AND ITEM_CODE  = '" & pstrItemCode & "' and Status =1  and Schedule_Quantity > isnull(Despatch_qty,0)"
                Else
                    strScheduleSql = "Select Schedule_Quantity=Sum(Isnull(Schedule_Quantity,0)),Despatch_Qty=Sum(isnull(Despatch_Qty,0)) from DailyMktSchedule where  UNIT_CODE = '" & gstrUNITID & "' and Account_Code='" & pstrAccountCode & "'"
                    strScheduleSql = strScheduleSql & " and Trans_Date <='" & VB6.Format(pstrDate, "dd/mmm/yyyy") & "'"
                    strScheduleSql = strScheduleSql & " and Cust_DrgNo = '" & pstrCustomerDrawNo & "' AND ITEM_CODE  = '" & pstrItemCode & "' and Status =1  and Schedule_Quantity > isnull(Despatch_qty,0)"
                End If
            Else
                If blnCalanderDateTrac Then
                    'In case of next day of the same month.
                    If Month(CDate(pstrDate)) = Month(CDate(strCalDate)) Then
                        strScheduleSql = "Select Schedule_Quantity=Sum(isnull(Schedule_Quantity,0)),Despatch_Qty=Sum(Isnull(Despatch_Qty,0)) from DailyMktSchedule where  UNIT_CODE = '" & gstrUNITID & "' and  Account_Code='" & pstrAccountCode & "' and "
                        strScheduleSql = strScheduleSql & " datepart(yyyy,Trans_Date)='" & Year(CDate(pstrDate)) & "'"
                        strScheduleSql = strScheduleSql & " and datepart(mm,Trans_Date)='" & Month(CDate(pstrDate)) & "'"
                        strScheduleSql = strScheduleSql & " and Trans_Date <='" & VB6.Format(strCalDate, "dd/mmm/yyyy") & "'"
                        strScheduleSql = strScheduleSql & " and Cust_DrgNo = '" & pstrCustomerDrawNo & "' AND ITEM_CODE  = '" & pstrItemCode & "' and Status =1 "
                        ' In case of last date of month (Except 31 december)
                    ElseIf (Month(CDate(pstrDate)) + 1) = Month(CDate(strCalDate)) Then
                        strScheduleSql = "Select Schedule_Quantity=Sum(isnull(Schedule_Quantity,0)),Despatch_Qty=Sum(Isnull(Despatch_Qty,0)) from DailyMktSchedule where  UNIT_CODE = '" & gstrUNITID & "' and Account_Code='" & pstrAccountCode & "' and "
                        strScheduleSql = strScheduleSql & " datepart(yyyy,Trans_Date)='" & Year(CDate(pstrDate)) & "'"
                        strScheduleSql = strScheduleSql & " and datepart(mm,Trans_Date) in ('" & Month(CDate(pstrDate)) & "','" & Month(CDate(strCalDate)) & "')"
                        strScheduleSql = strScheduleSql & " and Trans_Date <='" & VB6.Format(strCalDate, "dd/mmm/yyyy") & "'"
                        strScheduleSql = strScheduleSql & " and Cust_DrgNo = '" & pstrCustomerDrawNo & "' AND ITEM_CODE  = '" & pstrItemCode & "' and Status =1 "
                        ' In case of 31 december.
                    ElseIf (Year(CDate(pstrDate)) + 1) = Year(CDate(strCalDate)) Then
                        strScheduleSql = "Select Schedule_Quantity=Sum(isnull(Schedule_Quantity,0)),Despatch_Qty=Sum(Isnull(Despatch_Qty,0)) from DailyMktSchedule where  UNIT_CODE = '" & gstrUNITID & "' and Account_Code='" & pstrAccountCode & "' "
                        strScheduleSql = strScheduleSql & " and (( datepart(yyyy,Trans_Date) = '" & Year(CDate(pstrDate)) & "' and datepart(mm,Trans_Date)='" & Month(CDate(pstrDate)) & "')"
                        strScheduleSql = strScheduleSql & " or ( datepart(yyyy,Trans_Date) = '" & Year(CDate(strCalDate)) & "' and datepart(mm,Trans_Date)='" & Month(CDate(strCalDate)) & "'))"
                        strScheduleSql = strScheduleSql & " and Trans_Date <='" & VB6.Format(strCalDate, "dd/mmm/yyyy") & "'"
                        strScheduleSql = strScheduleSql & " and Cust_DrgNo = '" & pstrCustomerDrawNo & "' AND ITEM_CODE  = '" & pstrItemCode & "' and Status =1 "
                    End If
                Else
                    strScheduleSql = "Select Schedule_Quantity=Sum(isnull(Schedule_Quantity,0)),Despatch_Qty=Sum(Isnull(Despatch_Qty,0)) from DailyMktSchedule where  UNIT_CODE = '" & gstrUNITID & "' and  Account_Code='" & pstrAccountCode & "' and "
                    strScheduleSql = strScheduleSql & " datepart(yyyy,Trans_Date)='" & Year(CDate(pstrDate)) & "'"
                    strScheduleSql = strScheduleSql & " and datepart(mm,Trans_Date)='" & Month(CDate(pstrDate)) & "'"
                    strScheduleSql = strScheduleSql & " and Trans_Date <='" & VB6.Format(pstrDate, "dd/mmm/yyyy") & "'"
                    strScheduleSql = strScheduleSql & " and Cust_DrgNo = '" & pstrCustomerDrawNo & "' AND ITEM_CODE  = '" & pstrItemCode & "' and Status =1 " '''and Schedule_Flag =1   ( Now Not Consider)
                End If
            End If
            If objRsForSchedule.State = 1 Then objRsForSchedule.Close()
            objRsForSchedule.CursorLocation = ADODB.CursorLocationEnum.adUseClient
            objRsForSchedule.Open(strScheduleSql, mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
            If objRsForSchedule.EOF Or objRsForSchedule.BOF Or IsDBNull(objRsForSchedule.Fields("Schedule_Quantity").Value) Then
                GetTotalDispatchQuantityFromDailySchedule = -1
                mP_Connection.Execute("SET DATEFORMAT 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                objRsForSchedule.Close()
                Exit Function
            Else
                ldblTotalScheduleQuantity = Val(objRsForSchedule.Fields("Schedule_Quantity").Value)
                ldblTotalDispatchQuantity = Val(objRsForSchedule.Fields("Despatch_Qty").Value)
                GetTotalDispatchQuantityFromDailySchedule = Val(CStr(ldblTotalScheduleQuantity - ldblTotalDispatchQuantity))
                mP_Connection.Execute("SET DATEFORMAT 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                objRsForSchedule.Close()
                Exit Function
            End If
        Else
            If blnDSTracking = True Then
                If blnCalanderDateTrac Then
                    strScheduleSql = "Select Schedule_Quantity=Sum(Isnull(Schedule_Quantity,0)),Despatch_Qty=Sum(isnull(Despatch_Qty,0)) from DailyMktSchedule where  UNIT_CODE = '" & gstrUNITID & "' and Account_Code='" & pstrAccountCode & "'"
                    strScheduleSql = strScheduleSql & " and Trans_Date <='" & VB6.Format(strCalDate, "dd/mmm/yyyy") & "'"
                    strScheduleSql = strScheduleSql & " and Cust_DrgNo = '" & pstrCustomerDrawNo & "' AND ITEM_CODE  = '" & pstrItemCode & "' and Status =1  and Schedule_Quantity > isnull(Despatch_qty,0)"
                Else
                    strScheduleSql = "Select Schedule_Quantity=Sum(isnull(Schedule_Quantity,0)),Despatch_Qty=Sum(Isnull(Despatch_Qty,0)) from DailyMktSchedule where   UNIT_CODE = '" & gstrUNITID & "' and Account_Code='" & pstrAccountCode & "'"
                    strScheduleSql = strScheduleSql & " and Trans_Date <='" & VB6.Format(pstrDate, "dd/mmm/yyyy") & "'"
                    strScheduleSql = strScheduleSql & " and Cust_DrgNo = '" & pstrCustomerDrawNo & "' AND ITEM_CODE  = '" & pstrItemCode & "' and Status =1  and Schedule_Quantity > isnull(Despatch_qty,0)"
                End If
            Else
                If blnCalanderDateTrac Then
                    'In case of next day of the same month.
                    If Month(CDate(pstrDate)) = Month(CDate(strCalDate)) Then
                        strScheduleSql = "Select Schedule_Quantity=Sum(isnull(Schedule_Quantity,0)),Despatch_Qty=Sum(Isnull(Despatch_Qty,0)) from DailyMktSchedule where  UNIT_CODE = '" & gstrUNITID & "' and Account_Code='" & pstrAccountCode & "' and "
                        strScheduleSql = strScheduleSql & " datepart(yyyy,Trans_Date)='" & Year(CDate(pstrDate)) & "'"
                        strScheduleSql = strScheduleSql & " and datepart(mm,Trans_Date)='" & Month(CDate(pstrDate)) & "'"
                        strScheduleSql = strScheduleSql & " and Trans_Date <='" & VB6.Format(strCalDate, "dd/mmm/yyyy") & "'"
                        strScheduleSql = strScheduleSql & " and Cust_DrgNo = '" & pstrCustomerDrawNo & "' AND ITEM_CODE  = '" & pstrItemCode & "' and Status =1 "
                        ' In case of last date of month (Except 31 december)
                    ElseIf (Month(CDate(pstrDate)) + 1) = Month(CDate(strCalDate)) Then
                        strScheduleSql = "Select Schedule_Quantity=Sum(isnull(Schedule_Quantity,0)),Despatch_Qty=Sum(Isnull(Despatch_Qty,0)) from DailyMktSchedule where  UNIT_CODE = '" & gstrUNITID & "' and Account_Code='" & pstrAccountCode & "' and "
                        strScheduleSql = strScheduleSql & " datepart(yyyy,Trans_Date)='" & Year(CDate(pstrDate)) & "'"
                        strScheduleSql = strScheduleSql & " and datepart(mm,Trans_Date) in ('" & Month(CDate(pstrDate)) & "','" & Month(CDate(strCalDate)) & "')"
                        strScheduleSql = strScheduleSql & " and Trans_Date <='" & VB6.Format(strCalDate, "dd/mmm/yyyy") & "'"
                        strScheduleSql = strScheduleSql & " and Cust_DrgNo = '" & pstrCustomerDrawNo & "' AND ITEM_CODE  = '" & pstrItemCode & "' and Status =1 "
                        ' In case of 31 december.
                    ElseIf (Year(CDate(pstrDate)) + 1) = Year(CDate(strCalDate)) Then
                        strScheduleSql = "Select Schedule_Quantity=Sum(isnull(Schedule_Quantity,0)),Despatch_Qty=Sum(Isnull(Despatch_Qty,0)) from DailyMktSchedule where  UNIT_CODE = '" & gstrUNITID & "' and Account_Code='" & pstrAccountCode & "' "
                        strScheduleSql = strScheduleSql & " and (( datepart(yyyy,Trans_Date) = '" & Year(CDate(pstrDate)) & "' and datepart(mm,Trans_Date)='" & Month(CDate(pstrDate)) & "')"
                        strScheduleSql = strScheduleSql & " or ( datepart(yyyy,Trans_Date) = '" & Year(CDate(strCalDate)) & "' and datepart(mm,Trans_Date)='" & Month(CDate(strCalDate)) & "'))"
                        strScheduleSql = strScheduleSql & " and Trans_Date <='" & VB6.Format(strCalDate, "dd/mmm/yyyy") & "'"
                        strScheduleSql = strScheduleSql & " and Cust_DrgNo = '" & pstrCustomerDrawNo & "' AND ITEM_CODE  = '" & pstrItemCode & "' and Status =1 "
                    End If
                Else
                    strScheduleSql = "Select Schedule_Quantity=Sum(Isnull(Schedule_Quantity,0)),Despatch_Qty=Sum(Isnull(Despatch_Qty,0)) from DailyMktSchedule where  UNIT_CODE = '" & gstrUNITID & "' and Account_Code='" & pstrAccountCode & "' and "
                    strScheduleSql = strScheduleSql & " datepart(yyyy,Trans_Date)='" & Year(CDate(pstrDate)) & "'"
                    strScheduleSql = strScheduleSql & " and datepart(mm,Trans_Date)='" & Month(CDate(pstrDate)) & "'"
                    strScheduleSql = strScheduleSql & " and Trans_Date <='" & VB6.Format(pstrDate, "dd/mmm/yyyy") & "'"
                    strScheduleSql = strScheduleSql & " and Cust_DrgNo = '" & pstrCustomerDrawNo & "' AND ITEM_CODE  = '" & pstrItemCode & "' and Status =1 "
                End If
            End If
            If objRsForSchedule.State = 1 Then objRsForSchedule.Close()
            objRsForSchedule.CursorLocation = ADODB.CursorLocationEnum.adUseClient
            objRsForSchedule.Open(strScheduleSql, mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
            If objRsForSchedule.EOF Or objRsForSchedule.BOF Or IsDBNull(objRsForSchedule.Fields("Schedule_Quantity").Value) Then
                GetTotalDispatchQuantityFromDailySchedule = -1
                mP_Connection.Execute("SET DATEFORMAT 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                objRsForSchedule.Close()
                Exit Function
            Else
                ldblTotalScheduleQuantity = Val(objRsForSchedule.Fields("Schedule_Quantity").Value)
                ldblTotalDispatchQuantity = Val(objRsForSchedule.Fields("Despatch_Qty").Value)
                If blnDSTracking = False Then
                    GetTotalDispatchQuantityFromDailySchedule = Val(CStr(ldblTotalScheduleQuantity - ldblTotalDispatchQuantity)) + Val(CStr(pdblPrevQty))
                Else
                    GetTotalDispatchQuantityFromDailySchedule = Val(CStr(ldblTotalScheduleQuantity - ldblTotalDispatchQuantity))
                End If
                mP_Connection.Execute("SET DATEFORMAT 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                objRsForSchedule.Close()
                Exit Function
            End If
        End If
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        GetTotalDispatchQuantityFromDailySchedule = -1
        mP_Connection.Execute("SET DATEFORMAT 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function GetTotalDispatchQuantityFromMonthlySchedule(ByVal pstrAccountCode As String, ByVal pstrCustomerDrawNo As String, ByVal pstrItemCode As String, ByVal pstrDate As String, ByVal pstrMode As String, ByVal pdblPrevQty As Double) As Double
        Dim strScheduleSql As String
        Dim objRsForSchedule As ADODB.Recordset
        Dim ldblTotalDispatchQuantity As Double
        Dim ldblTotalScheduleQuantity As Double
        Dim lintLoopCounter As Short
        Dim strMakeDate As String
        Dim rsDSTracking As ClsResultSetDB
        Dim blnDSTracking As Boolean
        rsDSTracking = New ClsResultSetDB
        Call rsDSTracking.GetResult("Select DSWiseTracking From Sales_parameter WHERE UNIT_CODE = '" & gstrUNITID & "' ", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsDSTracking.RowCount > 0 Then blnDSTracking = IIf(IsDBNull(rsDSTracking.GetValue("DSWiseTracking")), False, IIf(rsDSTracking.GetValue("DSwisetracking") = False, False, True))
        rsDSTracking.ResultSetClose()
        On Error GoTo ErrHandler
        ldblTotalDispatchQuantity = 0
        ldblTotalScheduleQuantity = 0
        mP_Connection.Execute("SET DATEFORMAT 'mdy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        If Val(CStr(Month(CDate(pstrDate)))) < 10 Then
            strMakeDate = Year(CDate(pstrDate)) & "0" & Month(CDate(pstrDate))
        Else
            strMakeDate = Year(CDate(pstrDate)) & Month(CDate(pstrDate))
        End If
        If pstrMode = "ADD" Then
            objRsForSchedule = New ADODB.Recordset
            objRsForSchedule.CursorLocation = ADODB.CursorLocationEnum.adUseClient
            If blnDSTracking = True Then
                strScheduleSql = "Select Schedule_Qty=Sum(Isnull(Schedule_Qty,0)),Despatch_Qty=Sum(isnull(Despatch_Qty,0)) from MonthlyMktSchedule where  UNIT_CODE = '" & gstrUNITID & "' and Account_Code='" & pstrAccountCode & "'"
                strScheduleSql = strScheduleSql & " and Cust_DrgNo = '" & pstrCustomerDrawNo & "' AND Item_code = '" & pstrItemCode & "' and status =1  and Schedule_Qty > Despatch_qty and Year_Month <=" & Val(Trim(strMakeDate)) & ""
            Else
                strScheduleSql = "Select Schedule_Qty=Sum(Isnull(Schedule_Qty,0)),Despatch_Qty=Sum(isnull(Despatch_Qty,0))  from MonthlyMktSchedule where  UNIT_CODE = '" & gstrUNITID & "' and Account_Code='" & pstrAccountCode & "' and "
                strScheduleSql = strScheduleSql & " Year_Month=" & Val(Trim(strMakeDate)) & ""
                strScheduleSql = strScheduleSql & " and Cust_DrgNo = '" & pstrCustomerDrawNo & "' AND Item_code = '" & pstrItemCode & "' and status =1 "
            End If
            objRsForSchedule.Open(strScheduleSql, mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
            If objRsForSchedule.EOF Or objRsForSchedule.BOF Or IsDBNull(objRsForSchedule.Fields("Schedule_qty").Value) Then
                GetTotalDispatchQuantityFromMonthlySchedule = -1
                mP_Connection.Execute("SET DATEFORMAT 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                objRsForSchedule.Close()
                Exit Function
            Else
                ldblTotalScheduleQuantity = Val(objRsForSchedule.Fields("Schedule_Qty").Value)
                ldblTotalDispatchQuantity = Val(objRsForSchedule.Fields("Despatch_Qty").Value)
                GetTotalDispatchQuantityFromMonthlySchedule = Val(CStr(ldblTotalScheduleQuantity - ldblTotalDispatchQuantity))
                mP_Connection.Execute("SET DATEFORMAT 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                objRsForSchedule.Close()
                Exit Function
            End If
        Else
            If blnDSTracking = True Then
                strScheduleSql = "Select Schedule_Qty=Sum(Isnull(Schedule_Qty,0)),Despatch_Qty=Sum(isnull(Despatch_Qty,0)) from MonthlyMktSchedule where  UNIT_CODE = '" & gstrUNITID & "' and Account_Code='" & pstrAccountCode & "' "
                strScheduleSql = strScheduleSql & " and Cust_DrgNo = '" & pstrCustomerDrawNo & "' AND Item_code = '" & pstrItemCode & "' and status =1 and Schedule_Qty > Despatch_qty and Year_Month <=" & Val(Trim(strMakeDate)) & ""
            Else
                strScheduleSql = "Select Schedule_Qty=Sum(Isnull(Schedule_Qty,0)),Despatch_Qty=Sum(isnull(Despatch_Qty,0)) from MonthlyMktSchedule where  UNIT_CODE = '" & gstrUNITID & "' and Account_Code='" & pstrAccountCode & "' and "
                strScheduleSql = strScheduleSql & " Year_Month=" & Val(Trim(strMakeDate)) & ""
                strScheduleSql = strScheduleSql & " and Cust_DrgNo = '" & pstrCustomerDrawNo & "' AND Item_code = '" & pstrItemCode & "' and status =1 "
            End If
            If objRsForSchedule.State = 1 Then objRsForSchedule.Close()
            objRsForSchedule = New ADODB.Recordset
            objRsForSchedule.CursorLocation = ADODB.CursorLocationEnum.adUseClient
            objRsForSchedule.Open(strScheduleSql, mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
            If objRsForSchedule.EOF Or objRsForSchedule.BOF Or IsDBNull(objRsForSchedule.Fields("Schedule_qty").Value) Then
                GetTotalDispatchQuantityFromMonthlySchedule = -1
                mP_Connection.Execute("SET DATEFORMAT 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                objRsForSchedule.Close()
                Exit Function
            Else
                ldblTotalScheduleQuantity = Val(objRsForSchedule.Fields("Schedule_Qty").Value)
                ldblTotalDispatchQuantity = Val(objRsForSchedule.Fields("Despatch_Qty").Value)
                If blnDSTracking = False Then
                    GetTotalDispatchQuantityFromMonthlySchedule = Val(CStr(ldblTotalScheduleQuantity - ldblTotalDispatchQuantity)) + Val(CStr(pdblPrevQty))
                Else
                    GetTotalDispatchQuantityFromMonthlySchedule = Val(CStr(ldblTotalScheduleQuantity - ldblTotalDispatchQuantity))
                End If
                mP_Connection.Execute("SET DATEFORMAT 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                objRsForSchedule.Close()
                Exit Function
            End If
        End If
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        mP_Connection.Execute("SET DATEFORMAT 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        GetTotalDispatchQuantityFromMonthlySchedule = -1
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Public Function CheckcustorddtlQty(ByRef pstrMode As String, ByRef pstrItemCode As String, ByRef pstrDrgno As String, ByRef pdblQty As Double) As Boolean
        Dim rsCustOrdDtl As ClsResultSetDB
        Dim rssaledtl As ClsResultSetDB
        Dim dblSaleQuantity As Double
        Dim strCustOrdDtl As String
        On Error GoTo ErrHandler
        strCustOrdDtl = "Select openso,balance_Qty = order_qty - Despatch_qty from Cust_ord_dtl where "
        strCustOrdDtl = strCustOrdDtl & "Account_code ='" & txtCustCode.Text & "'" & " and Item_code ='"
        strCustOrdDtl = strCustOrdDtl & pstrItemCode & "' and cust_drgNo ='" & pstrDrgno
        strCustOrdDtl = strCustOrdDtl & "' and Authorized_flag = 1 and cust_ref = '" & txtRefNo.Text
        strCustOrdDtl = strCustOrdDtl & "' and Amendment_no = '" & txtAmendNo.Text & "' and UNIT_CODE = '" & gstrUNITID & "'"
        rsCustOrdDtl = New ClsResultSetDB
        rsCustOrdDtl.GetResult(strCustOrdDtl)
        If rsCustOrdDtl.GetValue("OpenSO") = True Then
            CheckcustorddtlQty = True
        Else
            Select Case pstrMode
                Case "ADD"
                    If rsCustOrdDtl.GetValue("Balance_Qty") < pdblQty Then
                        MsgBox("Balance Quantity available in SO for Customer Part code [ " & pstrDrgno & "] is " & rsCustOrdDtl.GetValue("Balance_Qty") & ".", MsgBoxStyle.Information, "eMPro")
                        CheckcustorddtlQty = False
                    Else
                        CheckcustorddtlQty = True
                    End If
                Case "EDIT"
                    rssaledtl = New ClsResultSetDB
                    rssaledtl.GetResult("Select Sales_Quantity from Sales_Dtl where doc_no = " & txtChallanNo.Text & " and item_code = '" & pstrItemCode & "' and cust_ITem_code = '" & pstrDrgno & "' and UNIT_CODE = '" & gstrUNITID & "'")
                    dblSaleQuantity = rssaledtl.GetValue("Sales_Quantity")
                    rssaledtl.ResultSetClose()
                    If (rsCustOrdDtl.GetValue("Balance_Qty")) < pdblQty Then
                        MsgBox("Balance Quantity available in SO for Customer Part code [ " & pstrDrgno & "] is " & rsCustOrdDtl.GetValue("Balance_Qty") & ".", MsgBoxStyle.Information, "eMPro")
                        CheckcustorddtlQty = False
                    Else
                        CheckcustorddtlQty = True
                    End If
            End Select
        End If
        rsCustOrdDtl.ResultSetClose()
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Public Function makeSelectSql(ByRef pstrCustno As String, ByRef pstrRefNo As String, ByRef pstrAmmNo As String, ByRef effectyrmon As String, ByRef Validyrmon As String, ByRef pstrstockLocation As String, ByRef strdate As String, ByRef pstrItemin As String, Optional ByRef pstrCondition As String = "") As String
        Dim strSelectSql As String
        Dim strNextWorkDay As String
        Dim RsobjSchedules As New ADODB.Recordset
        Dim blnCalendarDateTrac As Boolean
        Dim blnFutureDateDS As Boolean
        Dim intnoofFutureWorking As Integer
        Dim rsCalendarDate As New ADODB.Recordset
        Dim strCalDate As String
        Dim strquery As String
        Dim blnDSWiseTracking As Boolean

        On Error GoTo ErrHandler
        If RsobjSchedules.State = ADODB.ObjectStateEnum.adStateOpen Then RsobjSchedules.Close()
        RsobjSchedules.Open("SELECT DSWiseTracking,CalendarDateTrac,FUTUREDATE_DS_KNOCKOFF ,NOOFWORKINGDAYS_DS_KNOCKOFF  FROM sales_parameter WHERE  UNIT_CODE = '" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        blnFutureDateDS = RsobjSchedules.Fields("FUTUREDATE_DS_KNOCKOFF").Value
        intnoofFutureWorking = RsobjSchedules.Fields("NOOFWORKINGDAYS_DS_KNOCKOFF").Value.ToString
        blnDSWiseTracking = RsobjSchedules.Fields("NOOFWORKINGDAYS_DS_KNOCKOFF").Value.ToString
        If Not RsobjSchedules.EOF Then
            If IIf(RsobjSchedules.Fields(1).Value, 1, 0) = 1 Then
                blnCalendarDateTrac = True
            Else
                blnCalendarDateTrac = False
            End If
        End If
        If blnCalendarDateTrac Then
            If blnFutureDateDS Then
                strQuery = "select MAX(dt) dt from (select top " & intnoofFutureWorking & "  dt from calendar_mst where  UNIT_CODE = '" & gstrUNITID & "'  and dt > '" & VB6.Format(GetServerDate, "dd/mmm/yyyy") & "' and work_flg<>1 order by dt )a "
                If rsCalendarDate.State = ADODB.ObjectStateEnum.adStateOpen Then rsCalendarDate.Close()
                rsCalendarDate.Open(strQuery, mP_Connection, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockBatchOptimistic)
                If rsCalendarDate.EOF Or rsCalendarDate.BOF Or IsDBNull(rsCalendarDate.Fields("dt").Value) Then
                    MsgBox("Date in Calendar Master is not defined")
                    mP_Connection.Execute("SET DATEFORMAT 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    rsCalendarDate.Close()
                    Exit Function
                Else
                    rsCalendarDate.MoveFirst()
                    strCalDate = VB6.Format(rsCalendarDate.Fields("DT").Value, "dd/mmm/yyyy")
                    strNextWorkDay = strCalDate
                End If
                rsCalendarDate.Close()
            Else

                strNextWorkDay = GetNextWorkingDay(strdate)
                If strNextWorkDay = "-1" Then
                    makeSelectSql = ""
                    Exit Function
                End If
            End If
        End If

        strSelectSql = "Select b.Item_Code,c.Cust_DrgNo,c.Cust_Drg_Desc,d.Tariff_Code,a.unit_code "
        strSelectSql = strSelectSql & "from Cust_Ord_hdr a,MonthlyMktSchedule b,Cust_ord_dtl c,Item_Mst d where a.UNIT_CODE = '" & gstrUnitId & "' AND "
        strSelectSql = strSelectSql & " a.Cust_ref = c.Cust_ref and a.amendment_No = c.amendment_No and a.Account_code=c.account_code and a.unit_code = c.unit_code And c.Active_Flag ='A' "
        strSelectSql = strSelectSql & " and a.account_code=b.Account_code and a.unit_code = b.unit_code and c.Cust_drgNo=b.Cust_drgNo and c.unit_code = b.unit_code and b.ITem_code = d.Item_code and b.unit_code = d.unit_code and a.Account_Code='" & Trim(pstrCustno) & "'"
        strSelectSql = strSelectSql & " and a.Cust_Ref='" & Trim(pstrRefNo) & "' and a.Amendment_No='" & Trim(pstrAmmNo) & "' and b.status = 1 and b.Schedule_flag =1 and b.Year_Month =  " & Validyrmon
        strSelectSql = strSelectSql & " and b.Item_Code in(Select a.Item_code from Item_MSt a,Itembal_mst b where a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = '" & gstrUnitId & "' AND a.Item_Main_grp in (" & Trim(pstrItemin) & ") and a.Item_code = b.Item_code and b.Location_code ='" & pstrstockLocation & "' and b.Cur_bal >0 and a.hold_flag =0 and a.Status = 'A')"
        strSelectSql = strSelectSql & " UNION "
        strSelectSql = strSelectSql & "Select b.Item_Code,c.Cust_DrgNo,c.Cust_Drg_Desc,d.Tariff_Code,a.unit_code "
        strSelectSql = strSelectSql & "from Cust_Ord_hdr a,DailyMktSchedule b,Cust_ord_dtl c,ITem_Mst d  where a.UNIT_CODE = '" & gstrUnitId & "' AND "
        strSelectSql = strSelectSql & " a.Cust_ref = c.Cust_ref and a.amendment_No = c.amendment_No and a.Account_code=c.account_code and a.unit_code = c.unit_code"
        strSelectSql = strSelectSql & " and a.account_code=b.Account_code and a.unit_code = b.unit_code and c.Cust_drgNo=b.Cust_drgNo and c.unit_code = b.unit_code and b.ITem_code =d.ITem_code and b.unit_code = d.unit_code and b.status = 1 and b.Schedule_Flag = 1 And c.Active_Flag ='A' and a.Account_Code='" & Trim(pstrCustno) & "' "
        strSelectSql = strSelectSql & " and a.Cust_Ref='" & Trim(pstrRefNo) & "' and a.Amendment_No='" & Trim(pstrAmmNo) & "' "
        If blnCalendarDateTrac Then
            If blnDSWiseTracking = True Then
                strSelectSql = strSelectSql & "and b.Schedule_Quantity -b.despatch_qty >0 "
            Else
                If Month(CDate(strdate)) = Month(CDate(strNextWorkDay)) Then
                    strSelectSql = strSelectSql & " and  datepart(mm,b.trans_date) = '" & Month(CDate(strdate)) & "' And  b.Trans_Date <= '" & strNextWorkDay & "'  And DatePart(yyyy, b.Trans_Date) = '" & Year(CDate(strdate)) & "'"
                ElseIf (Month(CDate(strdate)) + 1) = Month(CDate(strNextWorkDay)) Then
                    strSelectSql = strSelectSql & " and  datepart(mm,b.trans_date) in ('" & Month(CDate(strdate)) & "','" & Month(CDate(strNextWorkDay)) & "') And  b.Trans_Date <= '" & strNextWorkDay & "'  And DatePart(yyyy, b.Trans_Date) = '" & Year(CDate(strdate)) & "'"
                ElseIf (Year(CDate(strdate)) + 1) = Year(CDate(strNextWorkDay)) Then
                    strSelectSql = strSelectSql & " and  (( datepart(yyyy,b.Trans_Date) = '" & Year(CDate(strdate)) & "' and datepart(mm,b.Trans_Date)='" & Month(CDate(strdate)) & "') or ( datepart(yyyy,b.Trans_Date) = '" & Year(CDate(strNextWorkDay)) & "' and datepart(mm,b.Trans_Date)='" & Month(CDate(strNextWorkDay)) & "')) And  b.Trans_Date <= '" & strNextWorkDay & "'"
                End If
            End If
        Else
            strSelectSql = strSelectSql & " and  datepart(mm,b.trans_date) = '" & Month(CDate(strdate)) & "' And  b.Trans_Date <= '" & strdate & "'  And DatePart(yyyy, b.Trans_Date) = '" & Year(CDate(strdate)) & "'"
        End If
            strSelectSql = strSelectSql & " and b.Item_Code in(Select a.Item_code from Item_MSt a,Itembal_mst b where a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND a.Item_Main_grp in (" & Trim(pstrItemin) & ") and a.Item_code = b.Item_code and b.Location_code ='" & pstrstockLocation & "' and b.Cur_bal >0 and a.hold_flag =0 and a.Status = 'A')"
            makeSelectSql = strSelectSql
            Exit Function
ErrHandler:  'The Error Handling Code Starts here
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
            Exit Function
    End Function
    Private Sub cmdempHelpRefNo_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdempHelpRefNo.Click
        On Error GoTo ErrHandler
        Dim strrefno() As String
        Dim strref As String
        Dim strString As String
        Dim strRefSql As String
        Dim intPlace As Short
        strString = txtRefNo.Text & "%"
        If UCase(CmbInvType.Text) <> "REJECTION" Then
            If UCase(CmbInvType.Text) = "JOBWORK INVOICE" Then
                strRefSql = "Select b.Cust_DrgNo,b.Cust_Ref,b.Amendment_No,b.Item_Code,a.PerValue from Cust_Ord_hdr a,Cust_Ord_Dtl b "
                strRefSql = strRefSql & " where a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND  b.Account_Code='" & Trim(txtCustCode.Text) & "' and b.Active_flag ='A' and "
                strRefSql = strRefSql & " a.Account_Code = b.Account_Code and a.Cust_ref =b.Cust_ref and "
                strRefSql = strRefSql & " a.Amendment_No = b.amendment_No AND a.Authorized_Flag = 1 and a.PO_type in ('J') "
                strRefSql = strRefSql & " and convert(varchar,a.Valid_date,111) >'" & VB6.Format(GetServerDate, "yyyy/mm/dd") & "' and convert(varchar,effect_Date,111) <='" & VB6.Format(GetServerDate, "yyyy/mm/dd") & "'"
                If txtRefNo.Text <> "" Then strref = " and b.cust_ref like '" & strString & "' " Else strref = ""
            ElseIf UCase(CmbInvType.Text) = "EXPORT INVOICE" Then
                strRefSql = "Select b.Cust_DrgNo,b.Cust_Ref,b.Amendment_No,b.Item_code,a.PerValue from Cust_Ord_hdr a,Cust_Ord_Dtl b"
                strRefSql = strRefSql & " where a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND  b.Account_Code='" & Trim(txtCustCode.Text) & "' and b.Active_flag ='A' and "
                strRefSql = strRefSql & " a.Account_Code = b.Account_Code and a.Cust_ref =b.Cust_ref and "
                strRefSql = strRefSql & " a.Amendment_No = b.amendment_No AND a.Authorized_Flag = 1 and a.PO_type in ('E') "
                strRefSql = strRefSql & " and convert(varchar,a.Valid_date,111) >'" & VB6.Format(GetServerDate, "yyyy/mm/dd") & "' and convert(varchar,effect_Date,111) <='" & VB6.Format(GetServerDate, "yyyy/mm/dd") & "'"
                If txtRefNo.Text <> "" Then strref = " and b.cust_ref like '" & strString & "' " Else strref = ""
            Else
                strRefSql = "Select b.Cust_DrgNo,b.Cust_Ref,b.Amendment_No,b.Item_Code,a.PerValue from Cust_Ord_hdr a,Cust_Ord_Dtl b"
                strRefSql = strRefSql & " where a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND b.Account_Code='" & Trim(txtCustCode.Text) & "' and b.Active_flag ='A' and "
                strRefSql = strRefSql & " a.Account_Code = b.Account_Code and a.Cust_ref =b.Cust_ref and "
                strRefSql = strRefSql & " a.Amendment_No = b.amendment_No  AND a.Authorized_Flag = 1 and a.PO_type in ('O','S','M') "
                strRefSql = strRefSql & " and convert(varchar,a.Valid_date,111) >'" & VB6.Format(GetServerDate, "yyyy/mm/dd") & "' and convert(varchar,effect_Date,111) <= '" & VB6.Format(GetServerDate, "yyyy/mm/dd") & "'"
                strRefSql = strRefSql & " order by b.Cust_Ref,b.Amendment_No,b.Cust_DrgNo,b.Item_Code "
                If txtRefNo.Text <> "" Then strref = " and b.cust_ref like '" & strString & "' " Else strref = ""
            End If
        Else
            strRefSql = "select a.Doc_No,a.Item_code,a.Rejected_Quantity from grn_Dtl a,grn_hdr b Where  a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND "
            strRefSql = strRefSql & "a.Doc_type = b.Doc_type And a.Doc_No = b.Doc_No and "
            strRefSql = strRefSql & "a.From_Location = b.From_Location and a.From_Location ='01R1'"
            strRefSql = strRefSql & "and a.Rejected_quantity > 0  and  b.Vendor_code = '" & txtCustCode.Text & "' and isnull(b.GRN_Cancelled,0) = 0"
            If txtRefNo.Text <> "" Then strref = " and a.Doc_No like " & strString & " " Else strref = ""
        End If
        If UCase(CmbInvType.Text) <> "REJECTION" Then
            strrefno = Me.ctlEMPHelpInvoiceEntry.ShowList(gstrCONNECTIONSERVER,gstrDSNName, gstrDatabaseName, strRefSql, "Refrence Details")
        Else
            strrefno = Me.ctlEMPHelpInvoiceEntry.ShowList(gstrCONNECTIONSERVER,gstrDSNName, gstrDatabaseName, strRefSql, "Refrence Details")
        End If
        If UBound(strrefno) <= 0 Then Exit Sub
        If strrefno(0) = "0" Then
            If UCase(CmbInvType.Text) <> "REJECTION" Then
                MsgBox("No Refrence available to Display", MsgBoxStyle.Information) : txtRefNo.Text = "" : txtRefNo.Focus() : Exit Sub
            Else
                MsgBox("No GRIN available to Display.") : txtRefNo.Text = "" : txtRefNo.Focus() : Exit Sub
            End If
        Else
            If UCase(CmbInvType.Text) <> "REJECTION" Then
                txtRefNo.Text = strrefno(1)
                txtAmendNo.Text = strrefno(2)
                intPlace = InStr(1, strrefno(4), " ")
                ctlPerValue.Text = Mid(strrefno(4), 1, intPlace)
            Else
                txtRefNo.Text = strrefno(0)
            End If
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Function DisplaydetailsfromCustOrdDtl(ByVal pintRow As Short, ByRef pstrLocationCode As String, Optional ByRef pstrItemCode As String = "", Optional ByRef pstrDrgno As String = "") As Boolean
        Dim strsaledtl As String
        Dim rsCustOrdDtl As ClsResultSetDB
        Dim rsStockLocation As ClsResultSetDB
        Dim strInvTypeDes As String
        Dim strInvSubTypeDes As String
        Dim strSqlBins As String
        Dim dblBins As Double
        Dim rsBinQty As ClsResultSetDB
        Dim strCustDrgNo As Object
        Dim varBinQtyTemp As Object
        On Error GoTo ErrHandler
        rsStockLocation = New ClsResultSetDB
        If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            strInvTypeDes = CmbInvType.Text
            strInvSubTypeDes = CmbInvSubType.Text
        ElseIf CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
            rsStockLocation.GetResult("Select Description,Sub_type_Description from SaleConf a,SalesChallan_dtl b where a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND  b.Invoice_type = a.Invoice_type and b.Sub_Category = a.Sub_type and a.Location_Code =b.Location_code and b.Location_code ='" & txtLocationCode.Text & "' and b.Doc_No = " & txtChallanNo.Text & " and (fin_start_date <= getdate() and fin_end_date >= getdate())")
            strInvTypeDes = rsStockLocation.GetValue("Description")
            strInvSubTypeDes = rsStockLocation.GetValue("Sub_type_Description")
            rsStockLocation.ResultSetClose()
        End If
        If (Len(Trim(pstrDrgno)) > 0) Or (Len(Trim(pstrItemCode)) > 0) Then
            If UCase(CStr(Trim(strInvTypeDes))) = "NORMAL INVOICE" Or UCase(CStr(Trim(strInvTypeDes))) = "JOBWORK INVOICE" Or UCase(CStr(Trim(strInvTypeDes))) = "EXPORT INVOICE" Or UCase(CStr(Trim(strInvTypeDes))) = "TRANSFER INVOICE" Then
                If UCase(Trim(strInvSubTypeDes)) <> "SCRAP" Then
                    strsaledtl = ""
                    strsaledtl = "Select Item_Code,Cust_DrgNo,Rate,Cust_Mtrl,Packing,Packing_Type,Others,tool_Cost,Excise_Duty "
                    '101188073
                    If gblnGSTUnit Then
                        strsaledtl = strsaledtl & ",ISNULL(HSNSACCODE,'') HSNSACCODE,ISNULL(ISHSNORSAC,'') ISHSNORSAC,CGSTTXRT_TYPE,SGSTTXRT_TYPE,UTGSTTXRT_TYPE,IGSTTXRT_TYPE,COMPENSATION_CESS "
                    End If
                    '101188073
                    strsaledtl = strsaledtl & "from Cust_ord_dtl WHERE  UNIT_CODE = '" & gstrUnitId & "' and "
                    strsaledtl = strsaledtl & "Account_Code ='" & txtCustCode.Text & "'and Cust_ref ='"
                    strsaledtl = strsaledtl & txtRefNo.Text & "' and Amendment_No = '" & Trim(txtAmendNo.Text) & "'and "
                    strsaledtl = strsaledtl & " Active_flag ='A' "
                    If Len(Trim(pstrDrgno)) > 0 Then
                        strsaledtl = strsaledtl & " and Cust_DrgNo = '" & pstrDrgno & "'"
                    End If
                    If Len(Trim(pstrItemCode)) > 0 Then
                        strsaledtl = strsaledtl & " and Item_Code ='" & pstrItemCode & "'"
                    End If
                    If Trim(UCase(strInvSubTypeDes)) = "FINISHED GOODS" Then
                        strsaledtl = strsaledtl & " and Item_Code in (select a.Item_code from Item_Mst a ,ItemBal_Mst b  where a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND a.Item_code = b.Item_code and a.ITem_Main_grp in('F','S') and b.Location_code = '" & pstrLocationCode & "')"
                    ElseIf Trim(UCase(strInvSubTypeDes)) = "COMPONENTS" Then
                        strsaledtl = strsaledtl & " and Item_Code in (select a.Item_code from Item_Mst a ,ItemBal_Mst b where a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND a.Item_code = b.Item_code and a.ITem_Main_grp in('C') and b.Location_code = '" & pstrLocationCode & "')"
                    ElseIf Trim(UCase(strInvSubTypeDes)) = "RAW MATERIAL" Then
                        strsaledtl = strsaledtl & " and Item_Code in (select a.Item_code from Item_Mst a ,ItemBal_Mst b  where a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND a.Item_code = b.Item_code and a.ITem_Main_grp in('R','S','B','M') and b.Location_code = '" & pstrLocationCode & "')"
                    ElseIf Trim(UCase(strInvSubTypeDes)) = "ASSETS" Then
                        strsaledtl = strsaledtl & " and Item_Code in (select a.Item_code from Item_Mst a ,ItemBal_Mst b  where a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND a.Item_code = b.Item_code and a.ITem_Main_grp in('P') and b.Location_code = '" & pstrLocationCode & "')"
                    ElseIf Trim(UCase(strInvSubTypeDes)) = "TRADING GOODS" Then
                        strsaledtl = strsaledtl & " and Item_Code in (select a.Item_code from Item_Mst a ,ItemBal_Mst b  where a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND a.Item_code = b.Item_code and a.ITem_Main_grp in('T','S') and b.Location_code = '" & pstrLocationCode & "')"
                    ElseIf Trim(UCase(strInvSubTypeDes)) = "TOOLS & DIES" Then
                        strsaledtl = strsaledtl & " and Item_Code in (select a.Item_code from Item_Mst a ,ItemBal_Mst b  where a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND a.Item_code = b.Item_code and a.ITem_Main_grp in('P','A') and b.Location_code = '" & pstrLocationCode & "')"
                    ElseIf Trim(UCase(strInvSubTypeDes)) = "EXPORTS" Then
                        strsaledtl = strsaledtl & " and Item_Code in (select a.Item_code from Item_Mst a ,ItemBal_Mst b  where a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND a.Item_code = b.Item_code and a.ITem_Main_grp in('F','S') and b.Location_code = '" & pstrLocationCode & "')"
                    End If
                Else
                    strsaledtl = ""
                    strsaledtl = "SELECT a.Item_Code,a.standard_Rate from Item_Mst a, itembal_Mst b where a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND a.ITem_code = b.Item_code and "
                    strsaledtl = strsaledtl & " Status = 'A' and Hold_flag <> 1 and Item_Code in ('" & pstrItemCode & "') and b.Location_code ='" & pstrLocationCode & "'"
                End If
            ElseIf UCase(Trim(strInvTypeDes)) = "TRANSFER INVOICE" And UCase(Trim(strInvSubTypeDes)) = "FINISHED GOODS" Then
                strsaledtl = ""
                strsaledtl = "SELECT Distinct a.Item_Code,c.Cust_drgNo,a.standard_Rate FROM Item_Mst a,Itembal_Mst b,CustItem_Mst c "
                strsaledtl = strsaledtl & " where a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = c.UNIT_CODE AND  a.UNIT_CODE = '" & gstrUNITID & "' AND a.Item_Code=b.Item_Code and a.Item_Code = c.ITem_Code"
                strsaledtl = strsaledtl & " and a.Status ='A' and a.Hold_Flag <> 1 and c.Account_code ='" & txtCustCode.Text & "'"
                strsaledtl = strsaledtl & " and a.Item_code  = '" & pstrItemCode & "' and b.location_code ='" & pstrLocationCode & "'"
            Else
            End If
            rsCustOrdDtl = New ClsResultSetDB
            rsCustOrdDtl.GetResult(strsaledtl)
            If rsCustOrdDtl.GetNoRows > 0 Then
                DisplaydetailsfromCustOrdDtl = True
                If rsCustOrdDtl.GetNoRows = 1 Then
                    rsCustOrdDtl.MoveFirst()
                    With SpChEntry
                        If Len(Trim(pstrItemCode)) = 0 Then
                            Call .SetText(1, pintRow, rsCustOrdDtl.GetValue("Item_code"))
                        End If
                        If Len(Trim(pstrDrgno)) = 0 Then
                            Call .SetText(2, pintRow, rsCustOrdDtl.GetValue("Cust_DrgNo"))
                            strCustDrgNo = rsCustOrdDtl.GetValue("Cust_DrgNo")
                        End If
                        strCustDrgNo = rsCustOrdDtl.GetValue("Cust_DrgNo")
                        rsBinQty = New ClsResultSetDB
                        strSqlBins = "Select isnull(BinQuantity,1) as BinQuantity from custitem_mst where cust_drgno= '" & strCustDrgNo & "' and UNIT_CODE = '" & gstrUNITID & "' and Account_code='" & Trim(Me.txtCustCode.Text) & "' "
                        rsBinQty.GetResult(strSqlBins)
                        If rsBinQty.GetNoRows > 0 Then
                            dblBins = rsBinQty.GetValue("BinQuantity")
                        Else
                            dblBins = 1
                        End If
                        rsBinQty.ResultSetClose()
                        varBinQtyTemp = Nothing
                        Call .GetText(22, pintRow, varBinQtyTemp)
                        If Val(varBinQtyTemp) > 0 Then
                            If Val(varBinQtyTemp) <> dblBins Then
                                Call .SetText(22, pintRow, Val(varBinQtyTemp))
                            Else
                                Call .SetText(22, pintRow, dblBins)
                            End If
                        Else
                            Call .SetText(22, pintRow, dblBins)
                        End If
                        Call .SetText(3, pintRow, (rsCustOrdDtl.GetValue("Rate") * CDbl(ctlPerValue.Text)))
                        Call .SetText(16, pintRow, rsCustOrdDtl.GetValue("Rate"))
                        Call .SetText(4, pintRow, (Val(rsCustOrdDtl.GetValue("Cust_mtrl")) * CDbl(ctlPerValue.Text)))
                        Call .SetText(17, pintRow, rsCustOrdDtl.GetValue("Cust_mtrl"))
                        Call .SetText(6, pintRow, rsCustOrdDtl.GetValue("Packing_Type"))
                        '101188073
                        If gblnGSTUnit Then
                            Call .SetText(7, pintRow, 0)
                        Else
                            Call .SetText(7, pintRow, rsCustOrdDtl.GetValue("Excise_duty"))
                        End If
                        '101188073
                        Call .SetText(10, pintRow, (Val(rsCustOrdDtl.GetValue("Others")) * CDbl(ctlPerValue.Text)))
                        Call .SetText(18, pintRow, rsCustOrdDtl.GetValue("Others"))
                        Call .SetText(15, pintRow, (Val(rsCustOrdDtl.GetValue("tool_cost")) * CDbl(ctlPerValue.Text)))
                        Call .SetText(19, pintRow, rsCustOrdDtl.GetValue("tool_cost"))
                        '101188073
                        If gblnGSTUnit Then
                            Call .SetText(IS_HSN_SAC, pintRow, rsCustOrdDtl.GetValue("ISHSNORSAC"))
                            Call .SetText(HSN_SAC_CODE, pintRow, rsCustOrdDtl.GetValue("HSNSACCODE"))
                            Call .SetText(CGST_TYPE, pintRow, rsCustOrdDtl.GetValue("CGSTTXRT_TYPE"))
                            Call .SetText(SGST_TYPE, pintRow, rsCustOrdDtl.GetValue("SGSTTXRT_TYPE"))
                            Call .SetText(IGST_TYPE, pintRow, rsCustOrdDtl.GetValue("IGSTTXRT_TYPE"))
                            Call .SetText(UTGST_TYPE, pintRow, rsCustOrdDtl.GetValue("UTGSTTXRT_TYPE"))
                            Call .SetText(COMP_CESS_TYPE, pintRow, rsCustOrdDtl.GetValue("COMPENSATION_CESS"))
                        End If
                        '101188073
                    End With
                End If
            Else
                DisplaydetailsfromCustOrdDtl = False
            End If
            rsCustOrdDtl.ResultSetClose()
        End If
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Function
    End Function
    Public Function ValidRowData(ByVal pintRow As Short) As Boolean
        Dim varItemCode As Object
        Dim varDrgNo As Object
        On Error GoTo ErrHandler
        With SpChEntry
            varItemCode = Nothing
            Call .GetText(1, pintRow, varItemCode)
            varDrgNo = Nothing
            Call .GetText(1, pintRow, varDrgNo)
            If Len(Trim(varItemCode)) = 0 Then
                .Row = pintRow : .Row2 = pintRow : .Col = 1 : .Col2 = 1 : .BlockMode = True
                .CtlEditMode = True : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .BlockMode = False
                MsgBox("Item Code Can Not Be Blank.", MsgBoxStyle.Information, "eMPro")
                .Row = pintRow : .Col = 1 : .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                ValidRowData = False
                Exit Function
            ElseIf Len(Trim(varDrgNo)) = 0 Then
                MsgBox("Customer Part Code Can Not Be Blank.", MsgBoxStyle.Information, "eMPro")
                .Row = pintRow : .Col = 2 : .Focus() : .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                ValidRowData = False
                Exit Function
            End If
            ValidRowData = True
        End With
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Function
    End Function
    Public Function MakeSelectStatementForITemMst(ByRef pstrItemMainGroup As String, ByRef pstrstockLocation As String) As String
        Dim strItembal As String
        On Error GoTo ErrHandler
        strItembal = "SELECT DISTINCT(A.ITEM_CODE),A.ITEM_CODE,A.DESCRIPTION,ISNULL(A.TARIFF_CODE,0) AS TARIFF_CODE,A.UNIT_CODE FROM ITEM_MST A,ITEMBAL_MST B"
        strItembal = strItembal & " where a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND a.Item_Code=b.Item_Code  "
        strItembal = strItembal & " and cur_bal >0 and a.Status ='A' and a.Hold_Flag <> 1"
        strItembal = strItembal & " and b.Location_Code = '" & pstrstockLocation & "'"
        strItembal = strItembal & " and a.Item_Main_Grp in (" & pstrItemMainGroup & ")"
        MakeSelectStatementForITemMst = strItembal
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Function
    End Function
    Public Function DisplayDetailsfromItemMst(ByVal pintRow As Short, ByRef pstrItemCode As String, ByRef pstrStocLocation As String, Optional ByRef pstrItemMainGrp As String = "", Optional ByRef pstrDrgno As String = "") As Boolean
        Dim strsaledtl As String
        Dim rsCustOrdDtl As ClsResultSetDB
        Dim rsStockLocation As ClsResultSetDB
        Dim strInvTypeDes As String
        Dim strInvSubTypeDes As String
        Dim strSqlBins As String
        Dim dblBins As Double
        Dim rsBinQty As ClsResultSetDB
        Dim strCustDrgNo As Object
        On Error GoTo ErrHandler
        If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            strInvTypeDes = CmbInvType.Text
            strInvSubTypeDes = CmbInvSubType.Text
        ElseIf CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
            rsStockLocation = New ClsResultSetDB
            rsStockLocation.GetResult("Select Description,Sub_type_Description,Stock_location from SaleConf a,SalesChallan_dtl b where a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND  b.Invoice_type = a.Invoice_type and b.Sub_Category = a.Sub_type and a.Location_Code =b.Location_code and b.Location_code ='" & txtLocationCode.Text & "' and b.Doc_No = " & txtChallanNo.Text & " and (fin_start_date <= getdate() and fin_end_date >= getdate())")
            strInvTypeDes = rsStockLocation.GetValue("Description")
            strInvSubTypeDes = rsStockLocation.GetValue("Sub_type_Description")
            rsStockLocation.ResultSetClose()
        End If
        If (Len(Trim(pstrDrgno)) > 0) Or (Len(Trim(pstrItemCode)) > 0) Then
            If UCase(Trim(strInvTypeDes)) = "TRANSFER INVOICE" And UCase(Trim(strInvSubTypeDes)) = "FINISHED GOODS" Then
                strsaledtl = ""
                strsaledtl = "SELECT Distinct a.Item_Code,c.Cust_drgNo,a.standard_Rate FROM Item_Mst a,Itembal_Mst b,CustItem_Mst c "
                strsaledtl = strsaledtl & " where a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = c.UNIT_CODE AND  a.UNIT_CODE = '" & gstrUNITID & "' AND a.Item_Code=b.Item_Code and a.Item_Code = c.ITem_Code"
                strsaledtl = strsaledtl & " and a.Status ='A' and a.Hold_Flag <> 1 and c.Account_code ='" & txtCustCode.Text & "'"
                strsaledtl = strsaledtl & " and a.Item_code in (select a.Item_code from Item_Mst a,itembal_mst b where a.unit_code = b.unit_code and a.unit_code = '" & gstrUNITID & "'  and a.Item_code = b.Item_code and Item_Main_grp in (" & pstrItemMainGrp & ")) and a.Item_Code  = '" & pstrItemCode & "' and b.Location_code  = '" & pstrStocLocation & "' "
            ElseIf UCase(Trim(strInvTypeDes)) = "REJECTION" Then
                If Len(Trim(txtRefNo.Text)) > 0 Then
                    strsaledtl = ""
                    strsaledtl = "SELECT a.Item_Code,Cust_DrgNo = a.Item_Code,standard_Rate = item_Rate from grn_Dtl a,ItemBal_Mst b where a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND "
                    strsaledtl = strsaledtl & " a.Item_Code ='" & pstrItemCode & "' and Doc_No =" & txtRefNo.Text & " and a.ITem_code = b.Item_code and b.Location_code ='" & pstrStocLocation & "'"
                Else
                    strsaledtl = ""
                    strsaledtl = "SELECT a.Item_Code,Cust_DrgNo = a.Item_Code,a.standard_Rate from Item_Mst a,Itembal_Mst b where a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND "
                    strsaledtl = strsaledtl & " Status = 'A' and Hold_flag <> 1 and a.Item_Code ='" & pstrItemCode & "' and a.Item_code = b.Item_code and b.Location_code ='" & pstrStocLocation & "'"
                End If
            ElseIf UCase(Trim(strInvTypeDes)) = "NORMAL INVOICE" And UCase(Trim(strInvSubTypeDes)) = "SCRAP" Then
                strsaledtl = ""
                strsaledtl = "SELECT a.Item_Code,Cust_DrgNo = a.Item_Code,a.standard_Rate from Item_Mst a,Itembal_Mst b where  a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND "
                strsaledtl = strsaledtl & " Status = 'A' and Hold_flag <> 1 and a.Item_Code ='" & pstrItemCode & "' and a.Item_code = b.Item_code and b.Location_code ='" & pstrStocLocation & "'"
            Else
                strsaledtl = ""
                strsaledtl = "SELECT Item_Code,Cust_DrgNo = Item_Code,standard_Rate from Item_Mst where UNIT_CODE = '" & gstrUNITID & "' AND "
                strsaledtl = strsaledtl & " Status = 'A' and Hold_flag <> 1 and Item_Code  = '" & pstrItemCode & "'"
                strsaledtl = strsaledtl & " and Item_code in (select a.Item_code from Item_Mst a,Itembal_Mst b where a.unit_code = b.unit_code and a.unit_code = '" & gstrUNITID & "' and  Item_Main_grp in (" & pstrItemMainGrp & ") and a.Item_code = b.Item_code and b.location_code ='" & pstrStocLocation & "')"
            End If
            rsCustOrdDtl = New ClsResultSetDB
            rsCustOrdDtl.GetResult(strsaledtl)
            If rsCustOrdDtl.GetNoRows > 0 Then
                DisplayDetailsfromItemMst = True
                If rsCustOrdDtl.GetNoRows = 1 Then
                    rsCustOrdDtl.MoveFirst()
                    With SpChEntry
                        If Len(Trim(pstrItemCode)) = 0 Then
                            Call .SetText(1, pintRow, rsCustOrdDtl.GetValue("Item_code"))
                        End If
                        If Len(Trim(pstrDrgno)) = 0 Then
                            Call .SetText(2, pintRow, rsCustOrdDtl.GetValue("Cust_DrgNo"))
                            strCustDrgNo = rsCustOrdDtl.GetValue("Cust_DrgNo")
                        End If
                        rsBinQty = New ClsResultSetDB
                        strSqlBins = "Select isnull(BinQuantity,1) as BinQuantity from custitem_mst where UNIT_CODE = '" & gstrUnitId & "' AND cust_drgno= '" & strCustDrgNo & "' and Account_code='" & Trim(Me.txtCustCode.Text) & "' "
                        rsBinQty.GetResult(strSqlBins)
                        If rsBinQty.GetNoRows > 0 Then
                            dblBins = rsBinQty.GetValue("BinQuantity")
                        Else
                            dblBins = 1
                        End If
                        Call .SetText(22, pintRow, dblBins)
                        rsBinQty.ResultSetClose()
                        Call .SetText(3, pintRow, (Val(rsCustOrdDtl.GetValue("Standard_Rate")) * CDbl(ctlPerValue.Text)))
                        Call .SetText(19, pintRow, Val(rsCustOrdDtl.GetValue("Standard_Rate")))
                        '101188073
                        If gblnGSTUnit Then
                            GetGSTTaxes(pintRow, rsCustOrdDtl.GetValue("Item_code"), strInvTypeDes)
                        End If
                        '101188073
                    End With
                End If
            Else
                DisplayDetailsfromItemMst = False
            End If
            rsCustOrdDtl.ResultSetClose()
        End If
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Function
    End Function
    Public Function MakeSelectSubQuery(ByRef pstrCustno As String, ByRef pstrRefNo As String, ByRef pstrAmmNo As String, ByRef pstrstockLocation As String, ByRef pstrItemin As String, Optional ByRef pstrItemNotin As String = "") As String
        Dim strSelectSql As String
        On Error GoTo ErrHandler
        ' Issue id   -   10319722
        strSelectSql = "Select c.Item_Code,c.Cust_DrgNo,c.Cust_Drg_Desc,d.Tariff_Code,c.Unit_code "
        strSelectSql = strSelectSql & "from Cust_Ord_hdr a,Cust_ord_dtl c,Item_Mst d where a.UNIT_CODE = '" & gstrUnitId & "' AND "
        strSelectSql = strSelectSql & " a.Cust_ref = c.Cust_ref and a.amendment_No = c.amendment_No and a.Account_code=c.account_code and a.UNIT_CODE = c.UNIT_CODE "
        strSelectSql = strSelectSql & " and  c.Item_code = d.Item_code and c.UNIT_CODE = d.UNIT_CODE and a.Account_Code='" & Trim(pstrCustno) & "' and a.Cust_Ref='" & Trim(pstrRefNo)
        strSelectSql = strSelectSql & "' and a.Amendment_No='" & Trim(pstrAmmNo) & "' And c.Active_Flag = 'A' "
        strSelectSql = strSelectSql & " and c.Item_Code in(Select a.Item_code from Item_MSt a,Itembal_mst b where a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND a.Item_Main_grp in (" & pstrItemin & ") and a.Item_code = b.Item_code and b.Location_code ='" & pstrstockLocation & "' and b.Cur_bal >0 and a.hold_flag =0 and a.Status = 'A'"
        strSelectSql = strSelectSql & ")"
        MakeSelectSubQuery = strSelectSql
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Function
    End Function
    Public Function AddDataFromGrinDtl(ByRef pstrVend As String, ByRef dblGrnNo As Double, ByRef pstrstockLocation As String, Optional ByRef intAlreadyItem As Short = 0) As String
        Dim strSQL As String
        Dim strItemCode As String
        Dim strItemNot As String
        Dim arrRejAcpt() As Object
        Dim intLoopCounter As Short
        Dim intArrLoopCount As Short
        Dim intMaxLoop As Short
        Dim intUbound As Short
        On Error GoTo ErrHandler
        strSQL = "select a.Item_code,a.Item_code,c.Description,c.Tariff_code from grn_dtl a,grn_hdr b,Item_Mst c where a.UNIT_CODE = b.UNIT_CODE AND a.UNIT_CODE = c.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND "
        strSQL = strSQL & "a.Doc_type = b.Doc_type and a.Doc_no = b.Doc_No "
        strSQL = strSQL & "and a.From_Location = b.From_Location "
        strSQL = strSQL & " and a.Item_Code = c.ITem_code and b.From_Location ='01R1'"
        strSQL = strSQL & " and c.Status = 'A' and Hold_Flag =0 "
        strSQL = strSQL & "and a.Rejected_quantity > 0 and b.Vendor_code = '" & pstrVend
        strSQL = strSQL & "' and a.Doc_No = " & dblGrnNo
        strSQL = strSQL & " and a.Item_code in (Select Item_Code from ItemBal_Mst Where unit_code='" & gstrUNITID & "' and Location_Code = '"
        strSQL = strSQL & pstrstockLocation & "' and Cur_bal > 0) and ((isnull(a.EXCESS_PO_Quantity,0) + isnull(a.Rejected_Quantity,0)) - isnull(a.Despatch_Quantity,0) - isnull(a.Inspected_Quantity,0) - isnull(a.RGP_Quantity,0)) > 0 and isnull(b.GRN_Cancelled,0) = 0"
        AddDataFromGrinDtl = strSQL
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Public Function CheckForTariffCode(ByRef pdblTariff As String) As Boolean
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        Dim VarDelete As Object
        Dim varItemCode As Object
        Dim varDrgNo As Object
        Dim varTariff As Object
        On Error GoTo ErrHandler
        intMaxLoop = SpChEntry.MaxRows
        '101188073
        If gblnGSTUnit Then
            CheckForTariffCode = True
            Exit Function
        End If
        '101188073
        If intMaxLoop > 1 Then
            With SpChEntry
                For intLoopCounter = 1 To intMaxLoop
                    VarDelete = Nothing
                    Call .GetText(14, intLoopCounter, VarDelete)
                    varItemCode = Nothing
                    Call .GetText(1, intLoopCounter, varItemCode)
                    varDrgNo = Nothing
                    Call .GetText(2, intLoopCounter, varDrgNo)
                    If VarDelete <> "D" Then
                        If Len(Trim(varItemCode)) > 0 Then
                            varTariff = Nothing
                            Call .GetText(20, intLoopCounter, varTariff)
                            If varTariff <> pdblTariff Then
                                MsgBox("Select Items of Same Tariff Code", MsgBoxStyle.Information, "eMPro")
                                CheckForTariffCode = False
                                Exit For
                            Else
                                CheckForTariffCode = True
                            End If
                        End If
                    End If
                Next
            End With
        Else
            CheckForTariffCode = True
        End If
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Function
    End Function
    Public Function ClearGridRow(ByRef pintRow As Short) As Object
        Dim intLoopCount As Short
        Dim intMaxLoop As Short
        On Error GoTo ErrHandler
        intMaxLoop = SpChEntry.MaxCols
        With SpChEntry
            For intLoopCount = 1 To intMaxLoop
                Call .SetText(intLoopCount, pintRow, "")
            Next
        End With
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Function
    End Function
    Public Function InsertinSalesDtlinEditMode(ByRef pintRow As Short) As String
        Dim ldblTotalBasicValue As Double
        Dim ldblTotalAccessibleValue As Double
        Dim ldblTempAccessibleVal As Double
        Dim ldblTotalExciseValue As Double
        Dim ldblTotalSaleTaxAmount As Double
        Dim ldblTotalSurchargeTaxAmount As Double
        Dim ldblNetInsurenceValue As Double
        Dim ldblTotalInvoiceValue As Double
        Dim ldblTotalOthersValues As Double
        Dim rsParameterData As ClsResultSetDB
        Dim strParamQuery As String
        Dim rsStockLocation As ClsResultSetDB
        Dim rsCustItemMst As ClsResultSetDB
        Dim rsItemMst As ClsResultSetDB
        Dim strSalesDtl As String
        Dim strInvDes As String
        Dim strInvSubTypeDes As String
        Dim lintItemQuantity As Double
        Dim lstrItemDrgno As String
        Dim lstrItemCode As String
        Dim ldblItemRate As Double
        Dim ldblItemCustMtrl As Double
        Dim ldblItemPacking As Double
        Dim ldblItemOthers As Double
        Dim ldblItemFromBox As Double
        Dim ldblItemToBox As Double
        Dim lstrItemDelete As String
        Dim lintItemPresQty As Double
        Dim lstrItemExciseCode As String
        Dim lstrItemCVDCode As String
        Dim lstrItemSADCode As String
        Dim ldblItemToolCost As Double
        Dim TempAccessibleVal As Double
        Dim ldblTotalCustMatrlValue As Double
        Dim blnISInsExcisable As Boolean
        Dim blnEOUFlag As Boolean
        Dim blnISExciseRoundOff As Boolean
        Dim blnISSalesTaxRoundOff As Boolean
        Dim blnISSurChargeTaxRoundOff As Boolean
        Dim blnAddCustMatrl As Boolean
        Dim blnISBasicRoundOff As Boolean
        On Error GoTo ErrHandler
        strParamQuery = "SELECT InsExc_Excise,CustSupp_Inc,EOU_Flag,Basic_Roundoff,SalesTax_Roundoff,Excise_Roundoff,SST_Roundoff,salesTax_Roundoff_decimal,ECESSRoundoff_decimal FROM Sales_Parameter where UNIT_CODE = '" & gstrUNITID & "'"
        rsParameterData = New ClsResultSetDB
        rsParameterData.GetResult(strParamQuery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        blnISBasicRoundOff = rsParameterData.GetValue("Basic_Roundoff")
        rsParameterData.ResultSetClose()
        rsStockLocation = New ClsResultSetDB
        rsStockLocation.GetResult("Select Description,Sub_type_Description from SaleConf a,SalesChallan_dtl b where a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND b.Invoice_type = a.Invoice_type and b.Sub_Category = a.Sub_type and a.Location_Code =b.Location_code and b.Location_code ='" & txtLocationCode.Text & "' and b.Doc_No = " & txtChallanNo.Text & " and (fin_start_date <= getdate() and fin_end_date >= getdate())")
        strInvDes = rsStockLocation.GetValue("Description")
        strInvSubTypeDes = rsStockLocation.GetValue("Sub_type_Description")
        rsStockLocation.ResultSetClose()
        strSalesDtl = ""
        With SpChEntry
            .Row = pintRow
            .Col = 1
            lstrItemCode = Trim(.Text)
            .Col = 2
            lstrItemDrgno = Trim(.Text)
            .Col = 3
            ldblItemRate = Val(.Text) / CDbl(ctlPerValue.Text)
            .Col = 4
            ldblItemCustMtrl = Val(.Text) / CDbl(ctlPerValue.Text)
            .Col = 5
            lintItemQuantity = Val(.Text)
            .Col = 6
            ldblItemPacking = Val(.Text)
            .Col = 7
            lstrItemExciseCode = Trim(.Text)
            .Col = 8
            lstrItemCVDCode = Trim(.Text)
            .Col = 9
            lstrItemSADCode = Trim(.Text)
            .Col = 10
            ldblItemOthers = Val(.Text) / CDbl(ctlPerValue.Text) * lintItemQuantity
            .Col = 11
            ldblItemFromBox = Val(.Text)
            .Col = 12
            ldblItemToBox = Val(.Text)
            .Col = 14
            lstrItemDelete = Trim(.Text)
            .Col = 15
            ldblItemToolCost = Val(.Text) / CDbl(ctlPerValue.Text)
            rsCustItemMst = New ClsResultSetDB
            rsItemMst = New ClsResultSetDB
            rsItemMst.GetResult("SELECT Description FROM Item_Mst WHERE Item_Code ='" & Trim(lstrItemCode) & "' and UNIT_CODE = '" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            rsCustItemMst.GetResult("SELECT Drg_desc FROM CustItem_Mst WHERE Account_code ='" & Trim(txtCustCode.Text) & "'and Cust_DrgNo='" & lstrItemDrgno & "'and Item_code ='" & lstrItemCode & "' and UNIT_CODE = '" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If UCase(Trim(lstrItemDelete)) <> "D" Then
                strSalesDtl = Trim(strSalesDtl) & "INSERT INTO sales_Dtl(UNIT_CODE,Cust_Ref,Amendment_No,Location_Code,Doc_No,Suffix,Item_Code,Sales_Quantity,"
                strSalesDtl = strSalesDtl & "From_Box,To_Box,Rate,Sales_Tax,Excise_Tax,Packing,Others,Cust_Mtrl,"
                strSalesDtl = strSalesDtl & "Year,Cust_Item_Code,Cust_Item_Desc,Tool_Cost,Measure_Code,Excise_type,SalesTax_type,CVD_type,SAD_type,Basic_Amount,Accessible_amount,CVD_Amount,SVD_amount,"
                strSalesDtl = strSalesDtl & "Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,Excise_per,CVD_per,SVD_per,CustMtrl_Amount,ToolCost_Amount) values ('" & gstrUNITID & "','" & Trim(txtRefNo.Text) & "','" & Trim(txtAmendNo.Text) & "','" & Trim(txtLocationCode.Text) & "','"
                strSalesDtl = strSalesDtl & Trim(txtChallanNo.Text) & "','','" & Trim(lstrItemCode) & "','" & Val(CStr(lintItemQuantity)) & "','"
                strSalesDtl = strSalesDtl & Val(CStr(ldblItemFromBox)) & "','" & Val(CStr(ldblItemToBox)) & "'," & Val(CStr(ldblItemRate)) & "," & Trim(lblSaltax_Per.Text) & ","
                TempAccessibleVal = CalculateAccessibleValue(pintRow, ldblNetInsurenceValue, blnISInsExcisable)
                If blnISExciseRoundOff Then
                    strSalesDtl = strSalesDtl & System.Math.Round(CalculateExciseValue(pintRow, TempAccessibleVal, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff))
                Else
                    strSalesDtl = strSalesDtl & CalculateExciseValue(pintRow, TempAccessibleVal, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff)
                End If
                strSalesDtl = strSalesDtl & "," & Val(CStr(ldblItemPacking)) & "," & Val(CStr(ldblItemOthers)) & "," & Val(CStr(ldblItemCustMtrl)) & ",'"
                strSalesDtl = strSalesDtl & Trim(CStr(Year(dtpDateDesc.Value))) & "','" & Trim(lstrItemDrgno) & "','" & IIf((Len(Trim(rsCustItemMst.GetValue("Drg_Desc"))) <= 0 Or Trim(CStr(rsCustItemMst.GetValue("Drg_Desc") = "Unknown"))), Trim(rsItemMst.GetValue("Description")), Trim(rsCustItemMst.GetValue("Drg_Desc"))) & "',"
                'If UCase(strInvDes) = "NORMAL INVOICE" Or UCase(strInvDes) = "EXPORT INVOICE" Then
                If UCase(strInvDes) = "NORMAL INVOICE" Or UCase(strInvDes) = "EXPORT INVOICE" Or UCase(strInvDes) = "TRANSFER INVOICE" Then
                    If UCase(CmbInvSubType.Text) <> "SCRAP" Then
                        strSalesDtl = strSalesDtl & ldblItemToolCost & ",'','"
                    End If
                Else
                    strSalesDtl = strSalesDtl & ldblItemToolCost & ",'','"
                End If
                strSalesDtl = strSalesDtl & lstrItemExciseCode & "','" & Trim(txtSaleTaxType.Text) & "','" & lstrItemCVDCode & "','" & lstrItemSADCode & "',"
                strSalesDtl = strSalesDtl & CalculateBasicValue(pintRow, blnISBasicRoundOff) & ","
                strSalesDtl = strSalesDtl & TempAccessibleVal & ","
                If blnISExciseRoundOff Then
                    strSalesDtl = strSalesDtl & System.Math.Round(CalculateExciseValue(pintRow, TempAccessibleVal, enumExciseType.RETURN_CVD, blnEOUFlag, blnISExciseRoundOff))
                    strSalesDtl = strSalesDtl & "," & System.Math.Round(CalculateExciseValue(pintRow, TempAccessibleVal, enumExciseType.RETURN_SAD, blnEOUFlag, blnISExciseRoundOff))
                Else
                    strSalesDtl = strSalesDtl & (CalculateExciseValue(pintRow, TempAccessibleVal, enumExciseType.RETURN_CVD, blnEOUFlag, blnISExciseRoundOff))
                    strSalesDtl = strSalesDtl & "," & (CalculateExciseValue(pintRow, TempAccessibleVal, enumExciseType.RETURN_SAD, blnEOUFlag, blnISExciseRoundOff))
                End If
                strSalesDtl = strSalesDtl & ",GetDate(),'"
                strSalesDtl = strSalesDtl & Trim(mP_User) & "', GetDate(),'" & Trim(mP_User) & "'," & GetTaxRate(lstrItemExciseCode, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='EXC'") & "," & GetTaxRate(lstrItemCVDCode, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='CVD'") & "," & GetTaxRate(lstrItemSADCode, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='SAD'") & "," & System.Math.Round(Val(CStr(lintItemQuantity * ldblItemCustMtrl)), 2) & "," & System.Math.Round(Val(CStr(lintItemQuantity * ldblItemToolCost)), 2) & ")" & vbCrLf
            End If
        End With
        rsItemMst.ResultSetClose()
        rsCustItemMst.ResultSetClose()
        InsertinSalesDtlinEditMode = strSalesDtl
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Function
    End Function
    Public Sub displayDeatilsfromCustOrdHdr()
        On Error GoTo ErrHandler
        Dim strCustOrdHdr As String
        Dim rsCustOrdHdr As ClsResultSetDB
        Dim strCurrency As String
        Dim intDecimalPlace As Short
        'To Get Data from Cusft_Ord_hdr
        Select Case CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                'If UCase(CStr((Trim(CmbInvType.Text)))) = "NORMAL INVOICE" Or UCase(CStr((Trim(CmbInvType.Text)))) = "JOBWORK INVOICE" Or UCase(CStr((Trim(CmbInvType.Text)))) = "EXPORT INVOICE" Then
                If UCase(CStr((Trim(CmbInvType.Text)))) = "NORMAL INVOICE" Or UCase(CStr((Trim(CmbInvType.Text)))) = "JOBWORK INVOICE" Or UCase(CStr((Trim(CmbInvType.Text)))) = "EXPORT INVOICE" Or UCase(CStr((Trim(CmbInvType.Text)))) = "TRANSFER INVOICE" Then
                    If CBool(UCase(CStr((Trim(CmbInvSubType.Text)) <> "SCRAP"))) Then
                        If Len(Trim(txtRefNo.Text)) Then
                            strCustOrdHdr = "Select max(Order_date),SalesTax_Type,AddVAT_Type,Currency_code,term_payment from Cust_ord_hdr"
                            strCustOrdHdr = strCustOrdHdr & " Where Account_Code='" & txtCustCode.Text & "' and Cust_Ref ='"
                            strCustOrdHdr = strCustOrdHdr & txtRefNo.Text & "'and Amendment_No ='" & txtAmendNo.Text & "'"
                            strCustOrdHdr = strCustOrdHdr & " and active_flag = 'A' and UNIT_CODE = '" & gstrUNITID & "' Group by salestax_type,AddVAT_Type,currency_code,term_payment"
                            rsCustOrdHdr = New ClsResultSetDB
                            rsCustOrdHdr.GetResult(strCustOrdHdr, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                            lblCreditTerm.Text = IIf(IsDBNull(rsCustOrdHdr.GetValue("term_payment")), "", rsCustOrdHdr.GetValue("term_payment"))
                            If Len(Trim(lblCreditTerm.Text)) > 0 Then
                                Call SelectDescriptionForField("crTrm_desc", "crtrm_termID", "Gen_CreditTrmMaster", lblCreditTermDesc, Trim(lblCreditTerm.Text))
                            Else
                                lblCreditTermDesc.Text = ""
                            End If
                            '101188073
                            If Not gblnGSTUnit Then
                                txtSaleTaxType.Text = rsCustOrdHdr.GetValue("SalesTax_Type")
                                Call txtSaleTaxType_Validating(txtSaleTaxType, New System.ComponentModel.CancelEventArgs(False))
                                txtAddVAT.Text = IIf(IsDBNull(rsCustOrdHdr.GetValue("AddVAT_Type")), "", rsCustOrdHdr.GetValue("AddVAT_Type"))
                                If txtAddVAT.Text.Length > 0 Then Call txtAddVAT_Validating(txtAddVAT, New System.ComponentModel.CancelEventArgs(False))
                                txtSurchargeTaxType.Text = ""
                                lblSurcharge_Per.Text = "0.00"
                            End If
                            '101188073
                            mCurrencyCode = rsCustOrdHdr.GetValue("Currency_code")
                            strCurrency = rsCustOrdHdr.GetValue("Currency_code")
                            If CBool(UCase(CStr((Trim(CmbInvType.Text)) = "EXPORT INVOICE"))) Then
                                lblCurrencyDes.Text = strCurrency
                            End If
                            intDecimalPlace = ToGetDecimalPlaces(mCurrencyCode)
                            If intDecimalPlace < 2 Then
                                intDecimalPlace = 2
                            End If
                            ctlInsurance.DecSize = intDecimalPlace : txtFreight.DecSize = intDecimalPlace
                            SetMaxLengthInSpread(intDecimalPlace)
                            Call ChangeCellTypeStaticText()
                            rsCustOrdHdr.ResultSetClose()
                            rsCustOrdHdr = Nothing
                        End If
                    End If
                Else
                    mCurrencyCode = ""
                    strCurrency = ""
                End If
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Public Function checkforDuplicateItemCodeandDrgNo(ByRef pstrItemCode As String, ByRef pstrDrgno As String, ByVal pintRow As Short) As Boolean
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        Dim varItemCode As Object
        Dim varDrgNo As Object
        Dim VarDelete As Object
        On Error GoTo ErrHandler
        intMaxLoop = pintRow - 1
        checkforDuplicateItemCodeandDrgNo = False
        For intLoopCounter = 1 To intMaxLoop
            With SpChEntry
                varItemCode = Nothing
                Call .GetText(1, intLoopCounter, varItemCode)
                varDrgNo = Nothing
                Call .GetText(2, intLoopCounter, varDrgNo)
                VarDelete = Nothing
                Call .GetText(14, intLoopCounter, VarDelete)
                If VarDelete <> "D" Then
                    If (Trim(varItemCode) = Trim(pstrItemCode)) And (Trim(varDrgNo) = Trim(pstrDrgno)) Then
                        checkforDuplicateItemCodeandDrgNo = True
                        MsgBox("Item you have entered already exist in grid", MsgBoxStyle.Information, "eMPro")
                        Exit For
                    End If
                End If
            End With
        Next
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Public Function Find_Value(ByRef strField As String) As String
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
    Private Sub cmdPrint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdPrint.Click
        On Error GoTo ErrHandler
        Dim intCount As Short
        Dim varTemp As Object
        Dim strFileName As String
        Kill(gstrLocalCDrive & "EmproInv\TypeToPrn.bat")
        If Len(objInvoicePrint.FileName) > 0 Then
            strFileName = objInvoicePrint.FileName
        End If
        If intNoCopies = 0 Then intNoCopies = 1
TypeFileNotFoundCreateRetry:
        For intCount = 1 To intNoCopies
            varTemp = Shell("cmd.exe /c " & gstrLocalCDrive & "EmproInv\TypeToPrn.bat " & strFileName, AppWinStyle.Hide)
            Sleep(5000)
            Call printBarCode(objInvoicePrint.BCFileName)
        Next
        Exit Sub
ErrHandler:
        If Err.Number = 53 Then
            FileOpen(1, gstrLocalCDrive & "EmproInv\TypeToPrn.bat", OpenMode.Append)
            PrintLine(1, "Type %1> prn")
            FileClose(1)
            GoTo TypeFileNotFoundCreateRetry
        End If
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClose.Click
        On Error GoTo ErrHandler
        FraInvoicePreview.Visible = False
        objInvoicePrint = Nothing
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub ReplaceJunkCharacters()
        On Error GoTo Errorhandler
        rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(15), "") 'Remove Uncompress Character
        rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(18), "") 'Remove Decompress Character
        rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(27) & "G", "") 'Remove Bold Character
        rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(27) & "H", "") 'Remove DeBold Character
        rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(12), "") 'Remove DeUnderline Character
        rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(27) & "-1", "") 'Remove Underline Character
        rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(27) & "-0", "") 'Remove DeUnderline Character
        rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(27) & "W1", "") 'Remove DoubleWidth Character
        rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(27) & "W0", "") 'Remove DeDoubleWidth Character
        rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(27) & "M", "") 'Remove Middle Character
        rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(27) & "P", "") 'Remove DeMiddle Character
        rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(27) & "E", "") 'Remove Elite Character
        rtbInvoicePreview.Text = Replace(rtbInvoicePreview.Text, Chr(27) & "F", "") 'Remove DeElite Character
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Sub printBarCode(ByVal pstrFileName As String)
        Dim varTemp As Object
        Dim strString As String
        strString = gstrLocalCDrive & "EmproInv\pdf-dot.bat BarCode.txt 4 2 2 1"
        varTemp = Shell("cmd.exe /c " & strString)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    '########################################################################################################
    '******************************************PRINTING******************************************************
    '########################################################################################################
    Sub PrintingInvoice()
        On Error GoTo ErrHandler
        If Val(txtChallanNo.Text) < 99000000 Then 'Check For Temporary Challan No.
            objInvoicePrint.ConnectionString = gstrCONNECTIONSTRING
            objInvoicePrint.Connection()
            objInvoicePrint.FileName = gstrLocalCDrive & "EmproInv\InvoicePrint.txt"
            objInvoicePrint.BCFileName = gstrLocalCDrive & "EmproInv\BarCode.txt"
            objInvoicePrint.CompanyName = gstrCOMPANY
            objInvoicePrint.Address1 = gstr_RGN_ADDRESS1
            objInvoicePrint.Address2 = gstr_RGN_ADDRESS2
            If chkDTRemoval.CheckState = System.Windows.Forms.CheckState.Checked Then
                objInvoicePrint.Print_Invoice(gstrUNITID, True, (txtLocationCode.Text), (txtChallanNo.Text), dtpRemoval.Text & " " & dtpRemovalTime.Value.Hour & ":" & dtpRemovalTime.Value.Minute)
            Else
                objInvoicePrint.Print_Invoice(gstrUNITID, True, (txtLocationCode.Text), (txtChallanNo.Text))
            End If
            rtbInvoicePreview.LoadFile(objInvoicePrint.FileName)
            rtbInvoicePreview.BackColor = System.Drawing.Color.White
            cmdPrint.Image = My.Resources.ico231.ToBitmap
            cmdClose.Image = My.Resources.ico217.ToBitmap
            FraInvoicePreview.Visible = True
            FraInvoicePreview.Enabled = True
            FraInvoicePreview.BringToFront()
            FraInvoicePreview.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Height) - 1050)
            FraInvoicePreview.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Width) - 400)
            FraInvoicePreview.Left = VB6.TwipsToPixelsX(100)
            FraInvoicePreview.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(ctlFormHeader1.Height) - 250)
            rtbInvoicePreview.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(FraInvoicePreview.Height) - 1000)
            rtbInvoicePreview.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(FraInvoicePreview.Width) - 200)
            rtbInvoicePreview.Left = VB6.TwipsToPixelsX(100)
            rtbInvoicePreview.Top = VB6.TwipsToPixelsY(900)
            rtbInvoicePreview.RightMargin = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(rtbInvoicePreview.Width) + 5000)
            shpInvoice.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(FraInvoicePreview.Width) - VB6.PixelsToTwipsX(shpInvoice.Width)) / 2)
            cmdPrint.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(shpInvoice.Left) + 100)
            cmdClose.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(cmdPrint.Left) + VB6.PixelsToTwipsX(cmdPrint.Width) + 100)
            cmdPrint.Enabled = True : cmdClose.Enabled = True
            FraInvoicePreview.Enabled = True : rtbInvoicePreview.Enabled = True : rtbInvoicePreview.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            ReplaceJunkCharacters()
            rtbInvoicePreview.Focus()
        Else
            'Printing unlocked Invoice
            Call PrintUnlockedInvoice()
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        'Function called, if error occurred
    End Sub
    Sub PrintUnlockedInvoice()
        Dim rsSalesConf As ClsResultSetDB
        Dim rssaledtl As ClsResultSetDB
        Dim rsItembal As ClsResultSetDB
        Dim rsSalesParameter As New ClsResultSetDB
        Dim strSalesconf As String
        Dim ItemCode As String
        Dim strDrgNo As String
        Dim strAccountCode As String
        Dim strCustRef As String
        Dim StrAmendmentNo As String
        Dim SALEDTL As String
        Dim intRow As Short
        Dim intLoopCount As Short
        Dim salesQuantity As Double
        Dim dblToolCost As Double
        Dim blnCheckToolCost As Boolean
        Dim strItembal As String
        Dim strtoolQuantity As String
        Dim strRetVal As String
        Dim objDrCr As New prj_DrCrNote.cls_DrCrNote(GetServerDate)
        Dim strInvoiceDate As String
        Dim varTmp As Object
        Dim varTmp1 As Object
        Dim intNoOfItem As Short
        Dim dblTmpItembal As Double
        Dim dblFinalItembal As Double
        On Error GoTo Err_Handler
        SALEDTL = "select * from Saleschallan_Dtl where UNIT_CODE = '" & gstrUNITID & "' AND Doc_No =" & txtChallanNo.Text & "  and Location_Code='" & Trim(txtLocationCode.Text) & "'"
        rssaledtl = New ClsResultSetDB
        rssaledtl.GetResult(SALEDTL, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        strAccountCode = rssaledtl.GetValue("Account_code")
        strCustRef = rssaledtl.GetValue("Cust_ref")
        StrAmendmentNo = rssaledtl.GetValue("Amendment_No")
        strInvoiceDate = getDateForDB(VB6.Format(rssaledtl.GetValue("Invoice_Date"), gstrDateFormat))
        rssaledtl.ResultSetClose()
        strSalesconf = "Select UpdatePO_Flag,UpdateStock_Flag,Stock_Location,OpenningBal,Preprinted_Flag,NoCopies from saleconf where  UNIT_CODE = '" & gstrUNITID & "' and "
        strSalesconf = strSalesconf & "Invoice_type = '" & mstrInvoiceType & "' and sub_type = '"
        strSalesconf = strSalesconf & mstrInvoiceSubType & "' and Location_Code='" & Trim(txtLocationCode.Text) & "' and datediff(dd,'" & strInvoiceDate & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & strInvoiceDate & "')<=0"
        rsSalesConf = New ClsResultSetDB
        rsSalesConf.GetResult(strSalesconf, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        updatePOflag = rsSalesConf.GetValue("UpdatePO_Flag")
        updatestockflag = rsSalesConf.GetValue("UpdateStock_Flag")
        strStockLocation = rsSalesConf.GetValue("Stock_Location")
        mOpeeningBalance = Val(rsSalesConf.GetValue("OpenningBal"))
        mIntNoCopies = rsSalesConf.GetValue("NoCopies")
        rsSalesConf.ResultSetClose()
        If Len(Trim(strStockLocation)) = 0 Then
            MsgBox("Please Define Stock Location in Sales Configuration. ")
            Exit Sub
        End If
        '***********To check if Tool Cost Deduction will be done or Not
        rsSalesParameter.GetResult("Select CheckToolAmortisation from Sales_Parameter WHERE UNIT_CODE = '" & gstrUNITID & "'")
        If rsSalesParameter.GetNoRows > 0 Then
            rsSalesParameter.MoveFirst()
            If Len(Trim(rsSalesParameter.GetValue("CheckToolAmortisation"))) = 0 Then
                MsgBox("First define Check Tool Amortisation in Sales Parameter", MsgBoxStyle.Information, "eMPro")
                rsSalesParameter.ResultSetClose()
                Exit Sub
            End If
            blnCheckToolCost = rsSalesParameter.GetValue("CheckToolAmortisation")
        Else
            MsgBox("No Data Defined in Sales Parameter", MsgBoxStyle.Information, "eMPro")
            rsSalesParameter.ResultSetClose()
            Exit Sub
        End If
        rsSalesParameter.ResultSetClose()
        SALEDTL = "Select Sales_Quantity,Item_code,Cust_Item_Code,toolcost_amount from sales_Dtl where Doc_No = " & txtChallanNo.Text & "  and UNIT_CODE = '" & gstrUNITID & "'  and Location_Code='" & Trim(txtLocationCode.Text) & "'"
        rssaledtl = New ClsResultSetDB
        rssaledtl.GetResult(SALEDTL, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        intRow = rssaledtl.GetNoRows
        rssaledtl.MoveFirst()
        '******Check for balance & despatch in Cust_ord_dtl
        For intLoopCount = 1 To intRow
            ItemCode = rssaledtl.GetValue("Item_code")
            salesQuantity = rssaledtl.GetValue("Sales_quantity")
            strDrgNo = rssaledtl.GetValue("Cust_Item_code")
            dblToolCost = rssaledtl.GetValue("ToolCost_amount")
            rsItembal = New ClsResultSetDB
            rsItembal.GetResult("Select Cur_bal from Itembal_Mst where Item_code = '" & ItemCode & "' and UNIT_CODE = '" & gstrUNITID & "' and Location_code ='" & strStockLocation & "'", ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            If rsItembal.GetNoRows > 0 Then
                If salesQuantity > rsItembal.GetValue("Cur_Bal") Then
                    MsgBox("Balance for item " & ItemCode & " at Location " & strStockLocation & " not available. ", MsgBoxStyle.Information, "eMPro")
                    rsItembal.ResultSetClose()
                    Exit Sub
                End If
            Else
                MsgBox("No Item in ItemMaster for Location " & strStockLocation & ".", MsgBoxStyle.OkOnly, "eMPro")
                rsItembal.ResultSetClose()
                Exit Sub
            End If
            rsItembal.ResultSetClose()
            If Len(Trim(strCustRef)) > 0 Then
                If UCase(mstrInvoiceType) <> "REJ" Then
                    rsItembal = New ClsResultSetDB
                    rsItembal.GetResult("Select balanceQty = order_qty - despatch_Qty,OpenSO from Cust_ord_dtl where  UNIT_CODE = '" & gstrUNITID & "' and account_code ='" & strAccountCode & "' and Cust_ref ='" & strCustRef & "' and Amendment_No = '" & StrAmendmentNo & "' and Item_code ='" & ItemCode & "' and Cust_drgNo ='" & strDrgNo & "' and Active_flag ='A'", ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                    If rsItembal.GetNoRows > 0 Then
                        If rsItembal.GetValue("OpenSO") = False Then
                            If salesQuantity > rsItembal.GetValue("BalanceQty") Then
                                MsgBox("Balance Quantity in SO for item " & ItemCode & " is " & rsItembal.GetValue("BalanceQty") & ".Check Quantity of Item in Challan.", MsgBoxStyle.Information, "eMPro")
                                rsItembal.ResultSetClose()
                                Exit Sub
                            End If
                        End If
                    Else
                        MsgBox("No Item (" & strItemCode & ") exist in SO - " & strCustRef & ".", MsgBoxStyle.Information, "eMPro")
                        rsItembal.ResultSetClose()
                        Exit Sub
                    End If
                    rsItembal.ResultSetClose()
                End If
            End If
            '************To Check for Tool Cost
            If blnCheckToolCost = True Then
                If dblToolCost > 0 Then
                    strItembal = "select BalanceQty = isnull(proj_qty,0) - isnull(UsedProjQty,0) from Amor_dtl "
                    strItembal = strItembal & " where UNIT_CODE = '" & gstrUNITID & "' AND account_code = '" & strAccountCode & "'"
                    strItembal = strItembal & " and Item_code = '" & ItemCode & "' and Cust_drgNo = '" & strDrgNo & "'"
                    rsItembal = New ClsResultSetDB
                    rsItembal.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                    If rsItembal.GetNoRows > 0 Then
                        strtoolQuantity = CStr(Val(rsItembal.GetValue("BalanceQty")))
                        If salesQuantity > Val(strtoolQuantity) Then
                            If CDbl(strtoolQuantity) = 0 Then
                                MsgBox("No Balance Available for Item (" & ItemCode & ") and customer Part Code (" & strDrgNo & ") For Amortisation Calculations. ", MsgBoxStyle.OkOnly, "eMPro")
                            Else
                                MsgBox("Quantity should not be Greater then available Balance Quantity for Amortisarion " & strtoolQuantity, MsgBoxStyle.OkOnly, "eMPro")
                            End If
                            Exit Sub
                        End If
                    Else
                        MsgBox("No Record Available in Tool Amortisation Master for Item (" & ItemCode & ") and customer Part Code (" & strDrgNo & ") For Amortisation Calculations. ", MsgBoxStyle.OkOnly, "eMPro")
                        rsItembal.ResultSetClose()
                        Exit Sub
                    End If
                    rsItembal.ResultSetClose()
                End If
            End If
            rssaledtl.MoveNext()
        Next
        rssaledtl.ResultSetClose()
        '****To Check in Rejection Invoice if Grin No Exist
        If UCase(mstrInvoiceType) = "REJ" Then
            If Len(Trim(strCustRef)) > 0 Then
                If CheckDataFromGrin(CDbl(Trim(strCustRef)), strAccountCode) = False Then
                    Exit Sub
                End If
            End If
        End If
        rssaledtl.ResultSetClose()
        rssaledtl = Nothing
        If Not (InvoiceGeneration() = True) Then
            Exit Sub
        End If
        If ConfirmWindow(10344, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
            If Len(Find_Value("select doc_no from SalesChallan_dtl where location_code='" & Trim(txtLocationCode.Text) & "' and doc_no='" & mInvNo & "' and UNIT_CODE = '" & gstrUNITID & "'")) > 0 Then
                MsgBox("Next Invoice number already generated." & vbCrLf & "Please skip current no either backward or forward" & vbCrLf & "in Sales Configuration Master Form.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "eMPro")
                Exit Sub
            End If
            ResetDatabaseConnection()
            mP_Connection.BeginTrans()
            mP_Connection.Execute("set Dateformat 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            mP_Connection.Execute(salesconf, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            If Len(Trim(mstrExcisePriorityUpdationString)) > 0 Then
                mP_Connection.Execute("update Saleschallan_dtl set Excise_type = '" & mstrExcisePriorityUpdationString & "' where  UNIT_CODE = '" & gstrUNITID & "' and Doc_no = " & txtChallanNo.Text, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End If
            mP_Connection.Execute(saleschallan, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            If updatePOflag = True Then
                mP_Connection.Execute(strupdatecustodtdtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End If
            If updatestockflag = True Then
                mP_Connection.Execute(strupdateitbalmst, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End If
            '***********To check if Tool Cost Deduction will be done or Not
            If blnCheckToolCost = True Then
                If Len(Trim(strUpdateAmorDtl)) > 0 Then
                    mP_Connection.Execute(strUpdateAmorDtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                End If
            End If
            If UCase(mstrInvoiceType) = "JOB" And GetBOMCheckFlagValue("BomCheck_Flag") Then
                mP_Connection.Execute("SET DATEFORMAT 'DMY'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                mP_Connection.Execute(mstrAnnex, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End If
            If UCase(mstrInvoiceType) = "REJ" Then
                If Len(Trim(mCust_Ref)) > 0 Then
                    mP_Connection.Execute(strupdateGrinhdr, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                End If
            End If
            'Accounts Posting is done here
            If mblnpostinfin = True Then
                If UCase(Trim(mstrInvoiceType)) <> "REJ" Then
                    strRetVal = objDrCr.SetARInvoiceDocument(gstrUNITID, mstrMasterString, mstrDetailString, prj_GLTransactions.cls_GLTransactions.udtOperationType.optInsert, gstrCONNECTIONSTRING)
                Else
                    If MsgBox("No Effects in Accounts.", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "eMPro") = MsgBoxResult.Yes Then
                        strRetVal = "Y"
                    Else
                        strRetVal = objDrCr.SetAPDocument(gstrUNITID, mstrMasterString, mstrDetailString, prj_GLTransactions.cls_GLTransactions.udtOperationType.optInsert, gstrCONNECTIONSTRING)
                    End If
                End If
                strRetVal = CheckString(strRetVal)
            Else
                strRetVal = "Y"
            End If
            If Not strRetVal = "Y" Then
                MsgBox(strRetVal, MsgBoxStyle.Information, "eMPro")
                mP_Connection.RollbackTrans()
                Exit Sub
            Else
                mP_Connection.CommitTrans()
                MsgBox("Invoice has been locked successfully with number " & mInvNo, MsgBoxStyle.Information, "eMPro")
                CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
            End If
            txtChallanNo.Text = CStr(mInvNo)
            txtChallanNo_Validating(txtChallanNo, New System.ComponentModel.CancelEventArgs(False))
        End If
        rtbInvoicePreview.LoadFile(objInvoicePrint.FileName)
        rtbInvoicePreview.BackColor = System.Drawing.Color.White
        cmdPrint.Image = My.Resources.ico231.ToBitmap
        cmdClose.Image = My.Resources.ico217.ToBitmap
        FraInvoicePreview.Visible = True
        FraInvoicePreview.Enabled = True
        FraInvoicePreview.BringToFront()
        FraInvoicePreview.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Height) - 1050)
        FraInvoicePreview.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Width) - 400)
        FraInvoicePreview.Left = VB6.TwipsToPixelsX(100)
        FraInvoicePreview.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(ctlFormHeader1.Height) - 250)
        rtbInvoicePreview.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(FraInvoicePreview.Height) - 1000)
        rtbInvoicePreview.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(FraInvoicePreview.Width) - 200)
        rtbInvoicePreview.Left = VB6.TwipsToPixelsX(100)
        rtbInvoicePreview.Top = VB6.TwipsToPixelsY(900)
        rtbInvoicePreview.RightMargin = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(rtbInvoicePreview.Width) + 5000)
        shpInvoice.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(FraInvoicePreview.Width) - VB6.PixelsToTwipsX(shpInvoice.Width)) / 2)
        cmdPrint.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(shpInvoice.Left) + 100)
        cmdClose.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(cmdPrint.Left) + VB6.PixelsToTwipsX(cmdPrint.Width) + 100)
        cmdPrint.Enabled = True : cmdClose.Enabled = True
        FraInvoicePreview.Enabled = True : rtbInvoicePreview.Enabled = True : rtbInvoicePreview.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        ReplaceJunkCharacters()
        rtbInvoicePreview.Focus()
        Exit Sub
Err_Handler:
        If Err.Number = 20545 Then
            Resume Next
        Else
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub
    Public Function InvoiceGeneration() As Boolean
        Dim gobjDB As New ClsResultSetDB
        Dim rsSalesConf As New ADODB.Recordset
        Dim Phone, Range, RegNo, EccNo, Address, Invoice_Rule As String
        Dim CST, PLA, Fax, EMail, UPST, Division As String
        Dim Commissionerate As String
        Dim strSQL As String
        Dim strCompMst, DeliveredAdd As String
        Dim strGRNDate As String
        Dim strVendorInvNo As String
        Dim strVendorInvDate As String
        Dim strCustRefForGrn As String
        Dim strSuffix As String
        gobjDB.GetResult("SELECT EOU_Flag, CustSupp_Inc,InsExc_Excise,postinfin,Excise_RoundOFF FROM sales_parameter WHERE UNIT_CODE = '" & gstrUNITID & "'")
        If gobjDB.GetValue("EOU_Flag") = True Then
            mStrCustMst = "Select Doc_No,Invoice_type from SalesChallan_Dtl where Invoice_Type <> 'EXP' and Location_Code='" & Trim(txtLocationCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "'"
            mblnEOUUnit = True
        Else
            mStrCustMst = "Select Doc_No,Invoice_type from SalesChallan_Dtl where Location_Code='" & Trim(txtLocationCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "'"
            mblnEOUUnit = False
        End If
        mblnAddCustomerMaterial = gobjDB.GetValue("CustSupp_Inc")
        mblnInsuranceFlag = gobjDB.GetValue("InsExc_Excise")
        mblnpostinfin = gobjDB.GetValue("postinfin")
        mblnExciseRoundOFFFlag = gobjDB.GetValue("Excise_RoundOFF")
        rsSalesConf.Open("SELECT * FROM SaleConf WHERE Invoice_Type='" & mstrInvoiceType & "' and UNIT_CODE = '" & gstrUNITID & "' AND  Sub_Type ='" & mstrInvoiceSubType & "' AND Location_Code='" & Trim(txtLocationCode.Text) & "' and datediff(dd,getdate(),fin_start_date)<=0  and datediff(dd,fin_end_date,getdate())<=0 ", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
        If Not rsSalesConf.EOF Then
            mstrPurposeCode = IIf(IsDBNull(rsSalesConf.Fields("inv_GLD_prpsCode").Value), "", rsSalesConf.Fields("inv_GLD_prpsCode").Value)
            mblnSameSeries = rsSalesConf.Fields("Single_Series").Value
            mstrReportFilename = IIf(IsDBNull(rsSalesConf.Fields("Report_filename").Value), "", rsSalesConf.Fields("Report_filename").Value)
            If mstrPurposeCode = "" Then
                MsgBox("Please select a Purpose Code in Sales Configuration", MsgBoxStyle.Information, "eMPro")
                mstrPurposeCode = ""
                Exit Function
            End If
        Else
            MsgBox("No record found in Sales Configuration for the selected Location, Invoice Type and Sub-Category", MsgBoxStyle.Information, "eMPro")
            mstrPurposeCode = ""
            Exit Function
        End If
        gobjDB.ResultSetClose()
        gobjDB = Nothing
        rsSalesConf.Close()
        rsSalesConf = Nothing
        InvoiceGeneration = False
        Call InitializeValues()
        Call ValuetoVariables()
        If mblnpostinfin = True Then
            If Not CreateStringForAccounts() Then
                InvoiceGeneration = False
                Exit Function
            End If
        End If
        Call updatesalesconfandsaleschallan()
        Call UpdateinSale_Dtl()
        If UCase(mstrInvoiceType) = "REJ" Then
            If Len(Trim(mCust_Ref)) > 0 Then
                Call UpdateGrnHdr(CDbl(mCust_Ref), mInvNo)
            End If
        End If
        objInvoicePrint.ConnectionString = gstrCONNECTIONSTRING
        objInvoicePrint.Connection()
        objInvoicePrint.FileName = gstrLocalCDrive & "EmproInv\InvoicePrint.txt"
        objInvoicePrint.BCFileName = gstrLocalCDrive & "EmproInv\BarCode.txt"
        objInvoicePrint.CompanyName = gstrCOMPANY
        objInvoicePrint.Address1 = gstr_RGN_ADDRESS1
        objInvoicePrint.Address2 = gstr_RGN_ADDRESS2
        If chkDTRemoval.CheckState = System.Windows.Forms.CheckState.Checked Then
            objInvoicePrint.Print_Invoice(gstrUNITID, True, (txtLocationCode.Text), (txtChallanNo.Text), dtpRemoval.Text & " " & dtpRemovalTime.Value.Hour & ":" & dtpRemovalTime.Value.Minute)
        Else
            objInvoicePrint.Print_Invoice(gstrUNITID, True, (txtLocationCode.Text), (txtChallanNo.Text))
        End If
        InvoiceGeneration = True
        Exit Function
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Public Sub InitializeValues()
        On Error GoTo ErrHandler
        mExDuty = 0 : mInvNo = 0 : mBasicAmt = 0 : msubTotal = 0 : mOtherAmt = 0 : mGrTotal = 0 : mStAmt = 0 : mFrAmt = 0
        mDoc_No = 0 : mCustmtrl = 0 : mAmortization = 0 : mstrAnnex = "" : strupdateGrinhdr = ""
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Public Sub ValuetoVariables()
        Dim strSQL As String
        Dim rsSalesChallan As ClsResultSetDB
        Dim strInvoiceDate As String
        On Error GoTo Err_Handler
        strSQL = "select INVOICE_DATE from Saleschallan_Dtl where Doc_No =" & txtChallanNo.Text & " and UNIT_CODE = '" & gstrUNITID & "'  and Location_Code='" & Trim(txtLocationCode.Text) & "'"
        rsSalesChallan = New ClsResultSetDB
        rsSalesChallan.GetResult(strSQL, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        strInvoiceDate = getDateForDB(VB6.Format(rsSalesChallan.GetValue("Invoice_Date"), gstrDateFormat))
        mInvNo = CDbl(GenerateInvoiceNo(mstrInvoiceType, mstrInvoiceSubType, strInvoiceDate))
        rsSalesChallan.ResultSetClose()
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Function CreateStringForAccounts() As Boolean
        Dim objRecordSet As New ADODB.Recordset
        Dim objTmpRecordset As New ADODB.Recordset
        Dim strRetVal As String
        Dim strInvoiceNo As String
        Dim strInvoiceDate As String
        Dim strCurrencyCode As String
        Dim dblInvoiceAmt As Double
        Dim dblTCStaxAmt As Double
        Dim dblExchangeRate As Double
        Dim dblBasicAmount As Double
        Dim dblBaseCurrencyAmount As Double
        Dim dblTaxAmt As Double
        Dim strTaxType As String
        Dim strCreditTermsID As String
        Dim strBasicDueDate As String
        Dim strPaymentDueDate As String
        Dim strExpectedDueDate As String
        Dim strCustomerGL As String
        Dim strCustomerSL As String
        Dim strTaxGL As String
        Dim strTaxSL As String
        Dim strItemGL As String
        Dim strItemSL As String
        Dim strGlGroupId As String
        Dim dblTaxRate As Double
        Dim varTmp As Object
        Dim dblCCShare As Double
        Dim iCtr As Short
        Dim strCustRef As String
        Dim blnExciseExumpted As Boolean
        Dim dblDiscountAmt As Double
        Dim arrstrExcPriority() As String
        Dim rsFULLExciseAmount As ClsResultSetDB
        Dim dblFullExciseAmount As Double
        Dim blnMsgBox As Boolean
        Dim blnFOC As Boolean = False
        Dim dblInvoiceAmtRoundOff_diff As Double
        mstrExcisePriorityUpdationString = ""
        blnMsgBox = False
        On Error GoTo ErrHandler
        objRecordSet.Open("SELECT * FROM  saleschallan_dtl WHERE Doc_No='" & Trim(txtChallanNo.Text) & "' and Location_Code='" & Trim(txtLocationCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
        If objRecordSet.EOF Then
            MsgBox("Invoice details not found", MsgBoxStyle.Information, "eMPro")
            CreateStringForAccounts = False
            If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                objRecordSet.Close()
                objRecordSet = Nothing
            End If
            Exit Function
        End If
        strInvoiceNo = CStr(mInvNo)
        strInvoiceDate = VB6.Format(objRecordSet.Fields("Invoice_Date").Value, "dd-MMM-yyyy")
        strCurrencyCode = Trim(IIf(IsDBNull(objRecordSet.Fields("Currency_Code").Value), "", objRecordSet.Fields("Currency_Code").Value))
        dblInvoiceAmt = IIf(IsDBNull(objRecordSet.Fields("total_amount").Value), 0, objRecordSet.Fields("total_amount").Value)
        dblExchangeRate = IIf(IsDBNull(objRecordSet.Fields("Exchange_Rate").Value), 1, objRecordSet.Fields("Exchange_Rate").Value)
        dblTCStaxAmt = IIf(IsDBNull(objRecordSet.Fields("TCSTaxAmount").Value), 1, objRecordSet.Fields("TCSTaxAmount").Value)
        strCustCode = Trim(objRecordSet.Fields("Account_Code").Value)
        strCustRef = Trim(IIf(IsDBNull(objRecordSet.Fields("cust_ref").Value), "", objRecordSet.Fields("cust_ref").Value))
        blnExciseExumpted = objRecordSet.Fields("ExciseExumpted").Value
        strCreditTermsID = Trim(IIf(IsDBNull(objRecordSet.Fields("payment_terms").Value), "", objRecordSet.Fields("payment_terms").Value))
        mstrCreditTermId = strCreditTermsID
        dblInvoiceAmtRoundOff_diff = IIf(IsDBNull(objRecordSet.Fields("TotalInvoiceAmtRoundOff_diff").Value), 0, objRecordSet.Fields("TotalInvoiceAmtRoundOff_diff").Value)
        Dim objCreditTerms As New prj_CreditTerm.clsCR_Term_Resolver
        If UCase(mstrInvoiceType) <> "SMP" Then 'if invoice type is not sample sales then
            'Retreiving the customer gl, sl and credit term id
            If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
            If UCase(Trim(mstrInvoiceType)) = "REJ" Then
                objTmpRecordset.Open("SELECT ISNULL(SUM(Basic_Amount),0) AS Basic_Amt FROM sales_dtl WHERE  UNIT_CODE = '" & gstrUNITID & "' and doc_no =" & txtChallanNo.Text, mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
                If Not objTmpRecordset.EOF Then
                    dblBasicAmount = objTmpRecordset.Fields("Basic_Amt").Value
                End If
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
                If (UCase(Trim(mstrInvoiceType)) = "REJ" And strCustRef <> "") Then 'In case of non line rejections Basic posting is not done
                    dblInvoiceAmt = dblInvoiceAmt - dblBasicAmount
                End If
                dblBasicAmount = 0
                objTmpRecordset.Open("SELECT GL_AccountID, Ven_slCode, CrTrm_Termid FROM Pur_VendorMaster where  UNIT_CODE = '" & gstrUNITID & "' and Prty_PartyID='" & strCustCode & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
            Else
                objTmpRecordset.Open("SELECT Cst_ArCode, Cst_slCode, Cst_CreditTerm FROM Sal_CustomerMaster where  UNIT_CODE = '" & gstrUNITID & "' and Prty_PartyID='" & strCustCode & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
            End If
            If objTmpRecordset.EOF Then
                If UCase(Trim(mstrInvoiceType)) = "REJ" Then
                    MsgBox("Vendor details not found", MsgBoxStyle.Information, "eMPro")
                Else
                    MsgBox("Customer details not found", MsgBoxStyle.Information, "eMPro")
                End If
                CreateStringForAccounts = False
                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSet.Close()
                    objRecordSet = Nothing
                End If
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objTmpRecordset.Close()
                    objTmpRecordset = Nothing
                End If
                Exit Function
            End If
            If UCase(Trim(mstrInvoiceType)) = "REJ" Then
                strCustomerGL = Trim(IIf(IsDBNull(objTmpRecordset.Fields("GL_AccountID").Value), "", objTmpRecordset.Fields("GL_AccountID").Value))
                strCustomerSL = Trim(IIf(IsDBNull(objTmpRecordset.Fields("Ven_slCode").Value), "", objTmpRecordset.Fields("Ven_slCode").Value))
                strCreditTermsID = Trim(IIf(IsDBNull(objTmpRecordset.Fields("CrTrm_Termid").Value), "", objTmpRecordset.Fields("CrTrm_Termid").Value))
            Else
                strCustomerGL = Trim(IIf(IsDBNull(objTmpRecordset.Fields("Cst_ArCode").Value), "", objTmpRecordset.Fields("Cst_ArCode").Value))
                strCustomerSL = Trim(IIf(IsDBNull(objTmpRecordset.Fields("Cst_slCode").Value), "", objTmpRecordset.Fields("Cst_slCode").Value))
                If strCreditTermsID = "" Then
                    strCreditTermsID = Trim(IIf(IsDBNull(objTmpRecordset.Fields("Cst_CreditTerm").Value), "", objTmpRecordset.Fields("Cst_CreditTerm").Value))
                    mstrCreditTermId = strCreditTermsID
                End If
            End If
            If strCreditTermsID = "" Then
                MsgBox("Credit Terms not found", MsgBoxStyle.Information, "eMPro")
                CreateStringForAccounts = False
                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSet.Close()
                    objRecordSet = Nothing
                End If
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objTmpRecordset.Close()
                    objTmpRecordset = Nothing
                End If
                Exit Function
            End If
            strRetVal = objCreditTerms.RetCR_Term_Dates("", "INV", strCreditTermsID, strInvoiceDate, gstrUNITID, "", "", gstrCONNECTIONSTRING)
            If CheckString(strRetVal) = "Y" Then
                strRetVal = Mid(strRetVal, 3)
                varTmp = Split(strRetVal, "»")
                strBasicDueDate = VB6.Format(varTmp(0), "dd-MMM-yyyy")
                strPaymentDueDate = VB6.Format(varTmp(1), "dd-MMM-yyyy")
                strExpectedDueDate = VB6.Format(varTmp(1), "dd-MMM-yyyy")
            Else
                MsgBox(CheckString(strRetVal), MsgBoxStyle.Information, "eMPro")
                CreateStringForAccounts = False
                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSet.Close()
                    objRecordSet = Nothing
                End If
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objTmpRecordset.Close()
                    objTmpRecordset = Nothing
                End If
                Exit Function
            End If
        Else 'if  the invoice type is sample sales then
            strRetVal = GetItemGLSL("", "Sample_Expences")
            If strRetVal = "N" Then
                CreateStringForAccounts = False
                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSet.Close()
                    objRecordSet = Nothing
                End If
                Exit Function
            End If
            varTmp = Split(strRetVal, "»")
            strCustomerGL = varTmp(0)
            strCustomerSL = varTmp(1)
        End If
        mstrMasterString = ""
        mstrDetailString = ""
        If UCase(Trim(mstrInvoiceType)) <> "REJ" Then
            mstrMasterString = "I»" & strInvoiceNo & "»Dr»»" & strInvoiceDate & "»»»»»SAL»I»" & strInvoiceNo & "»" & strInvoiceDate & "»"
            If UCase(mstrInvoiceType) <> "SMP" Then
                mstrMasterString = mstrMasterString & Trim(strCustCode) & "»" & gstrUNITID & "»" & strCurrencyCode & "»"
            Else
                mstrMasterString = mstrMasterString & "»" & gstrUNITID & "»" & strCurrencyCode & "»"
            End If
            mstrMasterString = mstrMasterString & System.Math.Round(dblInvoiceAmt, 0) & "»" & System.Math.Round(dblInvoiceAmt * dblExchangeRate, 0) & "»" & dblExchangeRate & "»" & strCreditTermsID & "»" & strBasicDueDate & "»" & strPaymentDueDate & "»" & strExpectedDueDate & "»" & strCustomerGL & "»" & strCustomerSL & "»" & mP_User & "»getdate()»»"
        Else
            mstrMasterString = "M»»" & VB6.Format(GetServerDate(), "dd-MMM-yyyy") & "»0»»" & gstrUNITID & "»" & Trim(strCustCode) & "»" & strInvoiceNo & "»" & strInvoiceDate & "»" & strBasicDueDate & "»" & strPaymentDueDate & "»" & strExpectedDueDate & "»" & strCurrencyCode & "»" & dblExchangeRate & "»" & System.Math.Round(dblInvoiceAmt) & "»0»»»Rej. Inv. " & strInvoiceNo & "»" & strCustomerGL & "»" & strCustomerSL & "»DR»" & strCustomerGL & "»" & strCustomerSL & "»»" & gstrCURRENCYCODE & "»" & mP_User & "»getdate()»0»AP»»»»0»»¦"
        End If
        iCtr = 1
        'CST/LST/SRT/VAT Posting
        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
        If Trim(IIf(IsDBNull(objRecordSet.Fields("SalesTax_Type").Value), "", objRecordSet.Fields("SalesTax_Type").Value)) <> "" Then
            objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE  UNIT_CODE = '" & gstrUNITID & "' and TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("SalesTax_Type").Value), "", objRecordSet.Fields("SalesTax_Type").Value)) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
            If Not objTmpRecordset.EOF Then
                strTaxType = Trim(UCase(objTmpRecordset.Fields("Tx_TaxeID").Value))
            Else
                MsgBox("Tax type not found", MsgBoxStyle.Information, "eMPro")
                CreateStringForAccounts = False
                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSet.Close()
                    objRecordSet = Nothing
                End If
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objTmpRecordset.Close()
                    objTmpRecordset = Nothing
                End If
                Exit Function
            End If
            If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
            If strTaxType = "LST" Or strTaxType = "CST" Or strTaxType = "SRT" Or strTaxType = "VAT" Then
                dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("Sales_Tax_Amount").Value), 0, objRecordSet.Fields("Sales_Tax_Amount").Value)
                dblBaseCurrencyAmount = dblTaxAmt
                dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("SalesTax_Per").Value), 0, objRecordSet.Fields("SalesTax_Per").Value)
                If dblBaseCurrencyAmount > 0 Then
                    'initializing the tax gl and sl here
                    strRetVal = GetTaxGlSl(strTaxType)
                    If strRetVal = "N" Then
                        MsgBox("GL for ARTAX is not defined for " & strTaxType, MsgBoxStyle.Information, "eMPro")
                        CreateStringForAccounts = False
                        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                            objRecordSet.Close()
                            objRecordSet = Nothing
                        End If
                        Exit Function
                    End If
                    varTmp = Split(strRetVal, "»")
                    strTaxGL = varTmp(0)
                    strTaxSL = varTmp(1)
                    If UCase(Trim(mstrInvoiceType)) <> "REJ" Then
                        mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»" & strTaxType & "»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                    Else
                        mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & dblTaxAmt & "»»CST/LST/VAT for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                    End If
                    iCtr = iCtr + 1
                End If
            End If
        End If
        '********Ecess Posting***********
        If UCase(CmbInvType.Text) = "CSM INVOICE" Then
            If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
            If Trim(IIf(IsDBNull(objRecordSet.Fields("ECESS_Type").Value), "", objRecordSet.Fields("ECESS_Type").Value)) <> "" Then
                objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE  UNIT_CODE = '" & gstrUNITID & "' and TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("ECESS_Type").Value), "", objRecordSet.Fields("ECESS_Type").Value)) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
                If Not objTmpRecordset.EOF Then
                    strTaxType = Trim(UCase(objTmpRecordset.Fields("Tx_TaxeID").Value))
                Else
                    MsgBox("Tax type not found", MsgBoxStyle.Information, "eMPro")
                    CreateStringForAccounts = False
                    If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                        objRecordSet.Close()
                        objRecordSet = Nothing
                    End If
                    If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then
                        objTmpRecordset.Close()
                        objTmpRecordset = Nothing
                    End If
                    Exit Function
                End If
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
                If strTaxType = "ECS" Then
                    dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("ECESS_Amount").Value), 0, objRecordSet.Fields("ECESS_Amount").Value)
                    dblBaseCurrencyAmount = dblTaxAmt
                    dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("ECESS_Per").Value), 0, objRecordSet.Fields("ECESS_Per").Value)
                    If dblBaseCurrencyAmount > 0 Then
                        'initializing the tax gl and sl here
                        strRetVal = GetTaxGlSl(strTaxType)
                        If strRetVal = "N" Then
                            MsgBox("GL for ARTAX is not defined for " & strTaxType, MsgBoxStyle.Information, "eMPro")
                            CreateStringForAccounts = False
                            If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                                objRecordSet.Close()
                                objRecordSet = Nothing
                            End If
                            Exit Function
                        End If
                        varTmp = Split(strRetVal, "»")
                        strTaxGL = varTmp(0)
                        strTaxSL = varTmp(1)
                        If UCase(Trim(mstrInvoiceType)) <> "REJ" Then
                            mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»" & strTaxType & "»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                        Else
                            mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & dblTaxAmt & "»»ECS for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                        End If
                        iCtr = iCtr + 1
                    End If
                End If
            End If
            '********SH-Ecess Posting***********
            If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
            If Trim(IIf(IsDBNull(objRecordSet.Fields("SECESS_Type").Value), "", objRecordSet.Fields("SECESS_Type").Value)) <> "" Then
                objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE  UNIT_CODE = '" & gstrUNITID & "' and TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("SECESS_Type").Value), "", objRecordSet.Fields("SECESS_Type").Value)) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
                If Not objTmpRecordset.EOF Then
                    strTaxType = Trim(UCase(objTmpRecordset.Fields("Tx_TaxeID").Value))
                Else
                    MsgBox("Tax type not found", MsgBoxStyle.Information, "eMPro")
                    CreateStringForAccounts = False
                    If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                        objRecordSet.Close()
                        objRecordSet = Nothing
                    End If
                    If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then
                        objTmpRecordset.Close()
                        objTmpRecordset = Nothing
                    End If
                    Exit Function
                End If
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
                If strTaxType = "ECSSH" Then
                    dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("SECESS_Amount").Value), 0, objRecordSet.Fields("SECESS_Amount").Value)
                    dblBaseCurrencyAmount = dblTaxAmt
                    dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("SECESS_Per").Value), 0, objRecordSet.Fields("SECESS_Per").Value)
                    If dblBaseCurrencyAmount > 0 Then
                        'initializing the tax gl and sl here
                        strRetVal = GetTaxGlSl(strTaxType)
                        If strRetVal = "N" Then
                            MsgBox("GL for ARTAX is not defined for " & strTaxType, MsgBoxStyle.Information, "eMPro")
                            CreateStringForAccounts = False
                            If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                                objRecordSet.Close()
                                objRecordSet = Nothing
                            End If
                            Exit Function
                        End If
                        varTmp = Split(strRetVal, "»")
                        strTaxGL = varTmp(0)
                        strTaxSL = varTmp(1)
                        If UCase(Trim(mstrInvoiceType)) <> "REJ" Then
                            mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»" & strTaxType & "»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                        Else
                            mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & dblTaxAmt & "»»ECSSH for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                        End If
                        iCtr = iCtr + 1
                    End If
                End If
            End If
        End If
        ''---- SST Posting
        dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("Surcharge_Sales_Tax_Amount").Value), 0, objRecordSet.Fields("Surcharge_Sales_Tax_Amount").Value)
        dblBaseCurrencyAmount = dblTaxAmt
        dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("Surcharge_SalesTax_Per").Value), 0, objRecordSet.Fields("Surcharge_SalesTax_Per").Value)
        If dblBaseCurrencyAmount > 0 Then
            'initializing the tax gl and sl here
            strRetVal = GetTaxGlSl("SST")
            If strRetVal = "N" Then
                MsgBox("GL for ARTAX is not defined for SST", MsgBoxStyle.Information, "eMPro")
                CreateStringForAccounts = False
                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSet.Close()
                    objRecordSet = Nothing
                End If
                Exit Function
            End If
            varTmp = Split(strRetVal, "»")
            strTaxGL = varTmp(0)
            strTaxSL = varTmp(1)
            If UCase(Trim(mstrInvoiceType)) <> "REJ" Then
                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»SST»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
            Else
                mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & dblTaxAmt & "»»Surcharge for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
            End If
            iCtr = iCtr + 1
        End If
        'Insurance Posting
        dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("Insurance").Value), 0, objRecordSet.Fields("Insurance").Value)
        dblBaseCurrencyAmount = dblTaxAmt
        If dblBaseCurrencyAmount > 0 Then
            'initializing the tax gl and sl here
            strRetVal = GetTaxGlSl("INS")
            If strRetVal = "N" Then
                MsgBox("GL for ARTAX is not defined for INS", MsgBoxStyle.Information, "eMPro")
                CreateStringForAccounts = False
                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSet.Close()
                    objRecordSet = Nothing
                End If
                Exit Function
            End If
            varTmp = Split(strRetVal, "»")
            strTaxGL = varTmp(0)
            strTaxSL = varTmp(1)
            If UCase(Trim(mstrInvoiceType)) <> "REJ" Then
                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»INS»0»" & "»»0»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
            Else
                mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & dblTaxAmt & "»»Insurance for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
            End If
            iCtr = iCtr + 1
        End If
        ''---- Freight Posting
        dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("Frieght_Amount").Value), 0, objRecordSet.Fields("Frieght_Amount").Value)
        dblBaseCurrencyAmount = dblTaxAmt
        If dblBaseCurrencyAmount > 0 Then
            'initializing the tax gl and sl here
            strRetVal = GetTaxGlSl("FRT")
            If strRetVal = "N" Then
                MsgBox("GL for ARTAX is not defined for FRT", MsgBoxStyle.Information, "eMPro")
                CreateStringForAccounts = False
                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSet.Close()
                    objRecordSet = Nothing
                End If
                Exit Function
            End If
            varTmp = Split(strRetVal, "»")
            strTaxGL = varTmp(0)
            strTaxSL = varTmp(1)
            If UCase(Trim(mstrInvoiceType)) <> "REJ" Then
                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»FRT»0»" & "»»0»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
            Else
                mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & dblTaxAmt & "»»Freight for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
            End If
            iCtr = iCtr + 1
        End If
        '******************Discount Posting
        dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("Discount_Amount").Value), 0, objRecordSet.Fields("Discount_Amount").Value)
        dblBaseCurrencyAmount = dblTaxAmt
        If dblBaseCurrencyAmount > 0 Then
            'initializing the tax gl and sl here
            strRetVal = GetItemGLSL("", "Discount_Interest")
            If strRetVal = "N" Then
                MsgBox("GL For Purpose Code Discount_Interest is not defined. ", MsgBoxStyle.Information, "eMPro")
                CreateStringForAccounts = False
                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSet.Close()
                    objRecordSet = Nothing
                End If
                Exit Function
            End If
            varTmp = Split(strRetVal, "»")
            strTaxGL = varTmp(0)
            strTaxSL = varTmp(1)
            If System.Math.Abs(dblBaseCurrencyAmount) > 0 Then
                If UCase(Trim(mstrInvoiceType)) <> "REJ" Then
                    mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»»TAX»0»" & "»»0»" & strTaxGL & "»" & strTaxSL & "»" & System.Math.Abs(dblBaseCurrencyAmount) & "»" & "Dr»»»»»»0»0»0»0»0" & "¦"
                Else
                    mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»DR»" & System.Math.Abs(dblBaseCurrencyAmount) & "»Discount amount for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                End If
            End If
            iCtr = iCtr + 1
        End If
        '******************TCS Tax Posting
        If (UCase(Trim(mstrInvoiceType)) = "INV") And (UCase(Trim(mstrInvoiceSubType)) = "L") Then
            dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("TCSTaxAmount").Value), 0, objRecordSet.Fields("TCSTaxAmount").Value)
            dblBaseCurrencyAmount = dblTaxAmt
            If dblBaseCurrencyAmount > 0 Then
                'initializing the tax gl and sl here
                strRetVal = GetTaxGlSl("TCS")
                If strRetVal = "N" Then
                    MsgBox("GL For Purpose Code TCS Tax is not defined. ", MsgBoxStyle.Information, "eMPro")
                    CreateStringForAccounts = False
                    If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                        objRecordSet.Close()
                        objRecordSet = Nothing
                    End If
                    Exit Function
                End If
                varTmp = Split(strRetVal, "»")
                strTaxGL = varTmp(0)
                strTaxSL = varTmp(1)
                If System.Math.Abs(dblBaseCurrencyAmount) > 0 Then
                    mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»»TCS»0»" & "»»0»" & strTaxGL & "»" & strTaxSL & "»" & System.Math.Abs(dblBaseCurrencyAmount) & "»" & "Cr»»»»»»0»0»0»0»0" & "¦"
                End If
                iCtr = iCtr + 1
            End If
        End If
        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close()
        objRecordSet.Open("SELECT sales_dtl.*, item_mst.GlGrp_code FROM sales_dtl, item_mst WHERE sales_dtl.UNIT_CODE = item_mst.UNIT_CODE and sales_dtl.UNIT_CODE = '" & gstrUNITID & "' AND sales_dtl.Doc_No='" & Trim(txtChallanNo.Text) & "' and sales_dtl.Item_Code=item_mst.Item_Code and sales_dtl.Location_Code='" & Trim(txtLocationCode.Text) & "'")
        If objRecordSet.EOF Then
            MsgBox("Item details not found.", MsgBoxStyle.Information, "eMPro")
            CreateStringForAccounts = False
            If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                objRecordSet.Close()
                objRecordSet = Nothing
            End If
            Exit Function
        End If
        While Not objRecordSet.EOF
            strGlGroupId = Trim(IIf(IsDBNull(objRecordSet.Fields("GlGrp_code").Value), "", objRecordSet.Fields("GlGrp_code").Value))
            blnFOC = CBool(Find_Value("select foc_invoice from salesChallan_dtl where  UNIT_CODE = '" & gstrUNITID & "' and Location_Code='" & Trim(txtLocationCode.Text) & "' and doc_no='" & Trim(txtChallanNo.Text) & "'"))
            'Basic Amount Posting
            If UCase(Trim(mstrInvoiceType)) = "CSM" And blnFOC Then
                'skip posting of basic if invoice is FOC CSM invoice
            ElseIf (UCase(Trim(mstrInvoiceType)) = "REJ" And strCustRef = "") Or UCase(Trim(mstrInvoiceType)) <> "REJ" Then 'In case of non line rejections Basic posting is not done
                dblBasicAmount = IIf(IsDBNull(objRecordSet.Fields("Basic_Amount").Value), 0, objRecordSet.Fields("Basic_Amount").Value)
                If mblnAddCustomerMaterial Then
                    dblBaseCurrencyAmount = dblBasicAmount + IIf(IsDBNull(objRecordSet.Fields("CustMtrl_Amount").Value), 0, objRecordSet.Fields("CustMtrl_Amount").Value)
                Else
                    dblBaseCurrencyAmount = dblBasicAmount
                End If
                If dblBaseCurrencyAmount > 0 Then
                    'initializing the item gl and sl************************
                    strRetVal = GetItemGLSL(strGlGroupId, mstrPurposeCode)
                    If strRetVal = "N" Then
                        CreateStringForAccounts = False
                        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                            objRecordSet.Close()
                            objRecordSet = Nothing
                        End If
                        Exit Function
                    End If
                    varTmp = Split(strRetVal, "»")
                    strItemGL = varTmp(0)
                    strItemSL = varTmp(1)
                    'initializing of item gl and sl ends here****************
                    'Posting the basic amount into cost centers, percentage wise
                    If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
                    objTmpRecordset.Open("SELECT * FROM invcc_dtl WHERE Invoice_Type='" & mstrInvoiceType & "' AND Sub_Type = '" & mstrInvoiceSubType & "' AND Location_Code ='" & Trim(txtLocationCode.Text) & "' AND ccM_cc_Percentage > 0 and UNIT_CODE = '" & gstrUNITID & "' ", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
                    If Not objTmpRecordset.EOF Then
                        While Not objTmpRecordset.EOF
                            dblCCShare = (dblBaseCurrencyAmount / 100) * objTmpRecordset.Fields("ccM_cc_Percentage").Value
                            If UCase(Trim(mstrInvoiceType)) <> "REJ" Then
                                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»ITM»SAL»" & iCtr & "»" & Trim(objRecordSet.Fields("item_code").Value) & "»" & strGlGroupId & "»0»" & strItemGL & "»" & strItemSL & "»" & dblCCShare & "»Cr»»" & Trim(objTmpRecordset.Fields("ccM_ccCode").Value) & "»»»»0»0»0»0»0¦"
                            Else
                                mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strItemGL & "»" & strItemSL & "»" & Trim(objTmpRecordset.Fields("ccM_ccCode").Value) & "»»»CR»" & dblCCShare & "»»Basic for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                            End If
                            objTmpRecordset.MoveNext()
                            iCtr = iCtr + 1
                        End While
                    Else
                        If UCase(Trim(mstrInvoiceType)) <> "REJ" Then
                            mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»ITM»SAL»" & iCtr & "»" & Trim(objRecordSet.Fields("item_code").Value) & "»" & strGlGroupId & "»0»" & strItemGL & "»" & strItemSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                        Else
                            mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strItemGL & "»" & strItemSL & "»»»»CR»" & dblBaseCurrencyAmount & "»»Basic for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                        End If
                        iCtr = iCtr + 1
                    End If
                End If
            End If
            'EXC Duty Posting
            'IF Condition added for Excise Exumption
            If blnExciseExumpted = False Then
                If mblnEOUUnit = False Then
                    dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("Excise_Tax").Value), 0, objRecordSet.Fields("Excise_Tax").Value)
                Else
                    dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("TotalExciseAmount").Value), 0, objRecordSet.Fields("TotalExciseAmount").Value)
                End If
                If mblnExciseRoundOFFFlag Then dblTaxAmt = System.Math.Round(dblTaxAmt, 0)
                dblBaseCurrencyAmount = dblTaxAmt
                If dblBaseCurrencyAmount > 0 Then
                    'initializing the tax gl and sl here
                    rsFULLExciseAmount = New ClsResultSetDB
                    rsFULLExciseAmount.GetResult("Select Sum(isnull(TotalExciseAmount,0)) as TotalExciseAmount from Sales_dtl where  UNIT_CODE = '" & gstrUNITID & "' and Doc_no =" & txtChallanNo.Text)
                    dblFullExciseAmount = rsFULLExciseAmount.GetValue("TotalExciseAmount")
                    rsFULLExciseAmount.ResultSetClose()
                    If CheckExcPriority() = 0 Then
                        If blnMsgBox = False Then
                            If MsgBox("No Excise Priority is Defined Would like to Post in ARTax ?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "eMPro") = MsgBoxResult.Yes Then
                                blnMsgBox = True
                            Else
                                CreateStringForAccounts = False
                                Exit Function
                            End If
                        End If
                        strRetVal = GetTaxGlSl("EXC")
                        If strRetVal = "N" Then
                            MsgBox("GL for ARTAX is not defined for EXC", MsgBoxStyle.Information, "eMPro")
                            CreateStringForAccounts = False
                            If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                                objRecordSet.Close()
                                objRecordSet = Nothing
                            End If
                            Exit Function
                        End If
                        varTmp = Split(strRetVal, "»")
                        strTaxGL = varTmp(0)
                        strTaxSL = varTmp(1)
                        mstrExcisePriorityUpdationString = ""
                    Else
                        arrstrExcPriority = ReturnGLSLAccExcPriority(1, dblFullExciseAmount)
                        If Len(Trim(arrstrExcPriority(0))) = 0 Then
                            arrstrExcPriority = ReturnGLSLAccExcPriority(2, dblFullExciseAmount)
                            If Len(Trim(arrstrExcPriority(1))) = 0 Then
                                arrstrExcPriority = ReturnGLSLAccExcPriority(3, dblFullExciseAmount)
                                If Len(Trim(arrstrExcPriority(1))) = 0 Then
                                    If blnMsgBox = False Then
                                        If MsgBox("Excise amount To be Posted is Greater then avalaible in All the Three Priorities Defined. would You like to Post in ARTax ?", MsgBoxStyle.YesNo, "eMPro") = MsgBoxResult.Yes Then
                                            blnMsgBox = True
                                        Else
                                            CreateStringForAccounts = False
                                            Exit Function
                                        End If
                                    End If
                                    strRetVal = GetTaxGlSl("EXC")
                                    If strRetVal = "N" Then
                                        MsgBox("GL for ARTAX is not defined for EXC", MsgBoxStyle.Information, "eMPro")
                                        CreateStringForAccounts = False
                                        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                                            objRecordSet.Close()
                                            objRecordSet = Nothing
                                        End If
                                        Exit Function
                                    End If
                                    varTmp = Split(strRetVal, "»")
                                    'To be Posted agianest ARtax 3
                                    strTaxGL = varTmp(0)
                                    strTaxSL = varTmp(1)
                                    mstrExcisePriorityUpdationString = ""
                                Else
                                    'To be Posted agianest Priority 3
                                    strTaxGL = arrstrExcPriority(0)
                                    strTaxSL = arrstrExcPriority(1)
                                    mstrExcisePriorityUpdationString = arrstrExcPriority(2)
                                End If
                            Else
                                'To be Posted agianest Priority 2
                                strTaxGL = arrstrExcPriority(0)
                                strTaxSL = arrstrExcPriority(1)
                                mstrExcisePriorityUpdationString = arrstrExcPriority(2)
                            End If
                        Else
                            'To be Posted agianest Priority 1
                            strTaxGL = arrstrExcPriority(0)
                            strTaxSL = arrstrExcPriority(1)
                            mstrExcisePriorityUpdationString = arrstrExcPriority(2)
                        End If
                    End If
                    If UCase(Trim(mstrInvoiceType)) <> "REJ" Then
                        mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»EXC»0»" & Trim(objRecordSet.Fields("item_code").Value) & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                    Else
                        mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & dblTaxAmt & "»»Excise for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                    End If
                    iCtr = iCtr + 1
                End If
            End If
            'Others Posting
            dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("Others").Value), 0, objRecordSet.Fields("Others").Value)
            dblBaseCurrencyAmount = dblTaxAmt
            '--------Handling of Freight charges being entered into Others field. Currently no ledger is defined for posting of Others
            'initialize the tax gl and sl here
            strRetVal = GetTaxGlSl("OTH")
            If strRetVal = "N" Then
                MsgBox("GL for ARTAX is not defined for OTHERS", MsgBoxStyle.Information, "eMPro")
                CreateStringForAccounts = False
                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSet.Close()
                    objRecordSet = Nothing
                End If
                Exit Function
            End If
            varTmp = Split(strRetVal, "»")
            strTaxGL = varTmp(0)
            strTaxSL = varTmp(1)
            'initialize the tax gl and sl here
            If dblBaseCurrencyAmount > 0 Then
                If UCase(Trim(mstrInvoiceType)) <> "REJ" Then
                    mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»OTH»0»" & Trim(objRecordSet.Fields("item_code").Value) & "»»0»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                Else
                    mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & dblTaxAmt & "»»Other Charges for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                End If
                iCtr = iCtr + 1
            End If
            objRecordSet.MoveNext()
        End While
        'Posting of rounded off amount
        strRetVal = GetItemGLSL("", "Rounded_Amt")
        If strRetVal = "N" Then
            CreateStringForAccounts = False
            If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                objRecordSet.Close()
                objRecordSet = Nothing
            End If
            Exit Function
        End If
        varTmp = Split(strRetVal, "»")
        strItemGL = varTmp(0)
        strItemSL = varTmp(1)
        If UCase(CmbInvType.Text) <> "CSM INVOICE" Then
            dblBaseCurrencyAmount = dblInvoiceAmt - System.Math.Round(dblInvoiceAmt, 0)
        Else
            dblBaseCurrencyAmount = dblInvoiceAmtRoundOff_diff
        End If
        dblBaseCurrencyAmount = System.Math.Round(dblBaseCurrencyAmount, 4)
        If System.Math.Abs(dblBaseCurrencyAmount) > 0 Then
            If UCase(Trim(mstrInvoiceType)) <> "REJ" Then
                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»»RND»0»" & "»»0»" & strItemGL & "»" & strItemSL & "»" & System.Math.Abs(dblBaseCurrencyAmount) & "»"
                If dblBaseCurrencyAmount < 0 Then
                    mstrDetailString = mstrDetailString & "Cr»»»»»»0»0»0»0»0" & "¦"
                Else
                    mstrDetailString = mstrDetailString & "Dr»»»»»»0»0»0»0»0" & "¦"
                End If
            Else
                If dblBaseCurrencyAmount < 0 Then
                    mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strItemGL & "»" & strItemSL & "»»»»CR»" & System.Math.Abs(dblBaseCurrencyAmount) & "»Rounding off amount for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                Else
                    mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strItemGL & "»" & strItemSL & "»»»»DR»" & System.Math.Abs(dblBaseCurrencyAmount) & "»Rounding off amount for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                End If
            End If
        End If
        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
            objRecordSet.Close()
            objRecordSet = Nothing
        End If
        CreateStringForAccounts = True
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        CreateStringForAccounts = False
        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
            objRecordSet.Close()
            objRecordSet = Nothing
        End If
    End Function
    Public Sub updatesalesconfandsaleschallan()
        Dim strSQL As String
        Dim rsSalesChallan As ClsResultSetDB
        Dim dblInvoiceAmt As Double
        Dim strInvoiceDate As String
        On Error GoTo Err_Handler
        strSQL = "select *  from Saleschallan_dtl where Doc_No = " & txtChallanNo.Text
        strSQL = strSQL & " and Invoice_type = '" & mstrInvoiceType & "'  and  sub_category =  '" & mstrInvoiceSubType & "' and Location_Code='" & Trim(txtLocationCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "'"
        rsSalesChallan = New ClsResultSetDB
        rsSalesChallan.GetResult(strSQL, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsSalesChallan.GetNoRows > 0 Then
            mAccount_Code = rsSalesChallan.GetValue("Account_Code")
            mCust_Ref = rsSalesChallan.GetValue("Cust_ref")
            mAmendment_No = rsSalesChallan.GetValue("Amendment_No")
            dblInvoiceAmt = rsSalesChallan.GetValue("total_amount")
            strInvoiceDate = getDateForDB(VB6.Format(rsSalesChallan.GetValue("Invoice_Date"), gstrDateFormat))
        End If
        rsSalesChallan.ResultSetClose()
        rsSalesChallan = Nothing
        If mblnEOUUnit = True Then
            If UCase(mstrInvoiceType) <> "EXP" Then
                If Not mblnSameSeries Then
                    salesconf = "update saleconf set current_No = " & mSaleConfNo & ", OpenningBal = openningBal - " & mAssessableValue & " where  UNIT_CODE = '" & gstrUNITID & "' and Invoice_type <> 'EXP' and Location_Code='" & Trim(txtLocationCode.Text) & "' and datediff(dd,'" & strInvoiceDate & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & strInvoiceDate & "')<=0"
                Else
                    salesconf = "update saleconf set current_No = " & mSaleConfNo & " where  UNIT_CODE = '" & gstrUNITID & "' and Single_Series = 1 and Invoice_type <> 'EXP' and Location_Code='" & Trim(txtLocationCode.Text) & "' and datediff(dd,'" & strInvoiceDate & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & strInvoiceDate & "')<=0" & vbCrLf
                    salesconf = salesconf & "update saleconf set OpenningBal = openningBal - " & mAssessableValue & " where  UNIT_CODE = '" & gstrUNITID & "' and Invoice_type <> 'EXP' and Location_Code='" & Trim(txtLocationCode.Text) & "' and datediff(dd,'" & strInvoiceDate & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & strInvoiceDate & "')<=0"
                End If
            Else
                salesconf = "update saleconf set current_No = " & mSaleConfNo & " where  UNIT_CODE = '" & gstrUNITID & "' and Invoice_type = 'EXP' and Location_Code='" & Trim(txtLocationCode.Text) & "' and datediff(dd,'" & strInvoiceDate & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & strInvoiceDate & "')<=0"
            End If
        Else
            If Not mblnSameSeries Then
                salesconf = "update saleconf set current_No = " & mSaleConfNo & " where  UNIT_CODE = '" & gstrUNITID & "' and Invoice_type = '" & mstrInvoiceType & "' and Location_Code='" & Trim(txtLocationCode.Text) & "' and datediff(dd,'" & strInvoiceDate & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & strInvoiceDate & "')<=0"
            Else
                salesconf = "update saleconf set current_No = " & mSaleConfNo & " where  UNIT_CODE = '" & gstrUNITID & "' and Single_Series = 1 and Location_Code='" & Trim(txtLocationCode.Text) & "' and datediff(dd,'" & strInvoiceDate & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & strInvoiceDate & "')<=0"
            End If
        End If
        Dim intInvoicePostingFlag As Short
        If mblnpostinfin = True Then
            intInvoicePostingFlag = 1
        Else
            intInvoicePostingFlag = 0
        End If
        saleschallan = "UPDATE SalesChallan_Dtl SET doc_no=" & mInvNo & ", Bill_Flag=1,print_flag = 1 , postingFlag=" & intInvoicePostingFlag & ",Payment_terms='" & mstrCreditTermId & "',Upd_dt=getdate(),Upd_Userid='" & mP_User & "' WHERE Doc_No=" & txtChallanNo.Text & " and UNIT_CODE = '" & gstrUNITID & "' and Location_Code='" & Trim(txtLocationCode.Text) & "' " & vbCrLf
        saleschallan = saleschallan & "UPDATE Sales_Dtl SET doc_no=" & mInvNo & " ,Upd_dt=getdate(),Upd_Userid='" & mP_User & "' WHERE  UNIT_CODE = '" & gstrUNITID & "' and Doc_No=" & txtChallanNo.Text & " and Location_Code='" & Trim(txtLocationCode.Text) & "'" & vbCrLf
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Sub UpdateinSale_Dtl()
        Dim rssaledtl As ClsResultSetDB
        Dim rsSaleConf As ClsResultSetDB
        Dim rsSalesChallan As ClsResultSetDB
        Dim strSQL As String
        Dim strInvoiceDate As String
        Dim strStockLocCode As String
        Dim rsSalesParameter As ClsResultSetDB
        Dim intRow, intLoopCount As Short
        Dim mItem_Code, mCust_Item_Code As String
        Dim mSales_Quantity As Double
        Dim mToolCost As Double
        Dim blnCheckToolCost As Boolean
        Dim strAccountCode As String
        Dim rsbom As ClsResultSetDB
        Dim irowcount As Short
        Dim intRwCount1 As Short
        strupdateitbalmst = ""
        strSelectItmbalmst = ""
        strupdatecustodtdtl = ""
        strUpdateAmorDtl = ""
        strupdateamordtlbom = ""
        On Error GoTo Err_Handler
        strSQL = "select * from Saleschallan_Dtl where  UNIT_CODE = '" & gstrUNITID & "' and Doc_No =" & txtChallanNo.Text & "  and Location_Code='" & Trim(txtLocationCode.Text) & "'"
        rsSalesChallan = New ClsResultSetDB
        rsSalesChallan.GetResult(strSQL, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        strInvoiceDate = getDateForDB(VB6.Format(rsSalesChallan.GetValue("Invoice_Date"), gstrDateFormat))
        strAccountCode = rsSalesChallan.GetValue("Account_code")
        rsSalesChallan.ResultSetClose()
        rsSaleConf = New ClsResultSetDB
        rsSaleConf.GetResult("Select Stock_Location from saleconf where  UNIT_CODE = '" & gstrUNITID & "' and Description = '" & CmbInvType.Text & "' and Sub_Type_Description ='" & Me.CmbInvSubType.Text & "' and Location_Code='" & Trim(txtLocationCode.Text) & "'and datediff(dd,'" & strInvoiceDate & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & strInvoiceDate & "')<=0")
        strStockLocCode = rsSaleConf.GetValue("Stock_Location")
        rsSaleConf.ResultSetClose()
        strSQL = "Select * from sales_Dtl where  UNIT_CODE = '" & gstrUNITID & "' and Doc_No = " & txtChallanNo.Text & " and Location_Code='" & Trim(txtLocationCode.Text) & "'"
        rsSalesParameter = New ClsResultSetDB
        rsSalesParameter.GetResult("Select CheckToolAmortisation from Sales_Parameter WHERE UNIT_CODE = '" & gstrUNITID & "'")
        If rsSalesParameter.GetNoRows > 0 Then
            rsSalesParameter.MoveFirst()
            If Len(Trim(rsSalesParameter.GetValue("CheckToolAmortisation"))) = 0 Then
                MsgBox("First define Check Tool Amortisation in Sales Parameter", MsgBoxStyle.Information, "eMPro")
                rsSalesParameter.ResultSetClose()
                Exit Sub
            End If
            blnCheckToolCost = rsSalesParameter.GetValue("CheckToolAmortisation")
        End If
        rsSalesParameter.ResultSetClose()
        rssaledtl = New ClsResultSetDB
        rssaledtl.GetResult(strSQL, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rssaledtl.GetNoRows > 0 Then
            intRow = rssaledtl.GetNoRows
            rssaledtl.MoveFirst()
            For intLoopCount = 1 To intRow
                If Not rssaledtl.EOFRecord Then
                    mItem_Code = rssaledtl.GetValue("Item_Code")
                    mCust_Item_Code = rssaledtl.GetValue("Cust_Item_Code")
                    mSales_Quantity = IIf(rssaledtl.GetValue("Sales_Quantity") = "", 0, rssaledtl.GetValue("Sales_Quantity"))
                    mToolCost = rssaledtl.GetValue("toolCost_amount")
                    strSelectItmbalmst = Trim(strSelectItmbalmst) & "Select cur_bal from Itembal_mst "
                    strSelectItmbalmst = strSelectItmbalmst & " where  UNIT_CODE = '" & gstrUNITID & "' and Location_code = '" & strStockLocation
                    strSelectItmbalmst = strSelectItmbalmst & "' and item_code = '" & mItem_Code & "'»"
                    strupdateitbalmst = Trim(strupdateitbalmst) & "Update Itembal_mst set cur_bal= cur_bal-"
                    strupdateitbalmst = strupdateitbalmst & mSales_Quantity & " where  UNIT_CODE = '" & gstrUNITID & "' and Location_code = '" & strStockLocation
                    strupdateitbalmst = strupdateitbalmst & "' and item_code = '" & mItem_Code & "'"
                    strupdatecustodtdtl = Trim(strupdatecustodtdtl) & "Update Cust_ord_dtl set Despatch_Qty = Despatch_Qty + "
                    strupdatecustodtdtl = strupdatecustodtdtl & mSales_Quantity & " where Account_code ='"
                    strupdatecustodtdtl = strupdatecustodtdtl & mAccount_Code & "'and Cust_DrgNo = '"
                    strupdatecustodtdtl = strupdatecustodtdtl & mCust_Item_Code & "' and Cust_ref = '" & mCust_Ref
                    strupdatecustodtdtl = strupdatecustodtdtl & "'and amendment_no = '" & mAmendment_No & "' and active_Flag ='A' and UNIT_CODE = '" & gstrUNITID & "'"
                    '***********To check if Tool Cost Deduction will be done or Not
                    If blnCheckToolCost = True Then
                        If mToolCost > 0 Then
                            strUpdateAmorDtl = Trim(strUpdateAmorDtl) & " Update Amor_dtl set usedProjQty = "
                            strUpdateAmorDtl = Trim(strUpdateAmorDtl) & " isnull(usedProjQty,0) + " & mSales_Quantity
                            strUpdateAmorDtl = Trim(strUpdateAmorDtl) & " where account_code = '" & strAccountCode
                            strUpdateAmorDtl = Trim(strUpdateAmorDtl) & "' and Item_code = '" & mItem_Code & "' and UNIT_CODE = '" & gstrUNITID & "'"
                            With mP_Connection
                                .Execute("DELETE FROM tmpBOM WHERE UNIT_CODE = '" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                .Execute("BOMExplosion '" & Trim(mItem_Code) & "','" & Trim(mItem_Code) & "',1,0,0,0,'" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            End With
                            rsbom = New ClsResultSetDB
                            rsbom.GetResult("select * from tmpBOM WHERE UNIT_CODE = '" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                            If rsbom.GetNoRows > 0 Then
                                irowcount = rsbom.GetNoRows
                                rsbom.MoveFirst()
                                For intRwCount1 = 1 To irowcount
                                    strupdateamordtlbom = Trim(strupdateamordtlbom) & " Update Amor_dtl set usedProjQty = "
                                    strupdateamordtlbom = Trim(strupdateamordtlbom) & " isnull(usedProjQty,0) + " & mSales_Quantity * Val(rsbom.GetValue("grossweight"))
                                    strupdateamordtlbom = Trim(strupdateamordtlbom) & " where account_code = '" & strAccountCode
                                    strupdateamordtlbom = Trim(strupdateamordtlbom) & "' and Item_code = '" & rsbom.GetValue("item_code") & "' and UNIT_CODE = '" & gstrUNITID & "'"
                                    rsbom.MoveNext()
                                Next
                            End If
                            rsbom.ResultSetClose()
                        End If
                    End If
                    rssaledtl.MoveNext()
                End If
            Next
        End If
        rssaledtl.ResultSetClose()
        rssaledtl = Nothing
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Function UpdateGrnHdr(ByRef pdblGrinNo As Double, ByRef pdblinvoiceNo As Double) As Object
        Dim rsSalesDtl As ClsResultSetDB
        Dim intMaxLoop As Short
        Dim strItemCode As String
        Dim dblqty As Double
        Dim intLoopCount As Short
        rsSalesDtl = New ClsResultSetDB
        rsSalesDtl.GetResult("select * from sales_dtl where Doc_No = " & txtChallanNo.Text & " and Location_Code='" & Trim(txtLocationCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "'")
        If rsSalesDtl.GetNoRows > 0 Then
            intMaxLoop = rsSalesDtl.GetNoRows
            rsSalesDtl.MoveFirst()
            strupdateGrinhdr = ""
            For intLoopCount = 1 To intMaxLoop
                strItemCode = rsSalesDtl.GetValue("ITem_code")
                dblqty = rsSalesDtl.GetValue("Sales_Quantity")
                If Len(Trim(strupdateGrinhdr)) = 0 Then
                    strupdateGrinhdr = "Update Grn_Dtl Set Despatch_Quantity = isnull(Despatch_Quantity,0) +" & dblqty
                    strupdateGrinhdr = strupdateGrinhdr & " Where ITem_Code = '" & strItemCode & "' and UNIT_CODE = '" & gstrUNITID & "' and Doc_No = " & pdblGrinNo
                Else
                    strupdateGrinhdr = strupdateGrinhdr & vbCrLf & "Update Grn_Dtl Set Despatch_Quantity = isnull(Despatch_Quantity,0) + " & dblqty
                    strupdateGrinhdr = strupdateGrinhdr & " Where ITem_Code = '" & strItemCode & "' and UNIT_CODE = '" & gstrUNITID & "' and Doc_No = " & pdblGrinNo
                End If
                rsSalesDtl.MoveNext()
            Next
        Else
            MsgBox("No Items Available in Invoice " & txtChallanNo.Text)
        End If
        rsSalesDtl.ResultSetClose()
    End Function
    Public Function GenerateInvoiceNo(ByVal pstrInvoiceType As String, ByRef pstrInvoiceSubType As String, ByVal pstrRequiredDate As String) As String
        On Error GoTo ErrHandler
        Dim clsInstEMPDBDbase As New EMPDataBase.EMPDB(gstrUNITID)
        Dim strCheckDOcNo As String 'Gets the Doc Number from Back End
        Dim strTempSeries As String 'Find the Numeric series in Doc No
        Dim strSuffix As String 'Generate a NEW Series
        Dim strZeroSuffix As String
        Dim strFin_Start_Date As String
        Dim strFin_End_Date As String
        Dim strSQL As String 'String SQL Query
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        If Len(Trim(pstrInvoiceType)) > 0 Then 'For Dated Docs
            strSQL = "Select Current_No,Suffix,Fin_start_date,Fin_end_Date From saleConf Where  UNIT_CODE = '" & gstrUNITID & "' and "
            strSQL = strSQL & "Invoice_Type ='" & pstrInvoiceType & "' and  sub_type='" & pstrInvoiceSubType & "' AND Location_Code ='" & Trim(txtLocationCode.Text) & "' and datediff(dd,'" & pstrRequiredDate & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & pstrRequiredDate & "')<=0"
            With clsInstEMPDBDbase.CConnection
                .OpenConnection(gstrDSNName, gstrDatabaseName)
                .ExecuteSQL("Set Dateformat 'dmy'")
            End With
            clsInstEMPDBDbase.CRecordset.OpenRecordset(strSQL, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
            If clsInstEMPDBDbase.CRecordset.Recordcount > 0 Then
                'Get Last Doc No Saved
                strCheckDOcNo = CStr(clsInstEMPDBDbase.CRecordset.GetFieldValue("Current_No", EMPDataBase.EMPDB.ADODataType.ADONumeric, EMPDataBase.EMPDB.ADOCustomFormat.CustomZeroDecimal))
                strSuffix = CStr(clsInstEMPDBDbase.CRecordset.GetFieldValue("suffix", EMPDataBase.EMPDB.ADODataType.ADONumeric, EMPDataBase.EMPDB.ADOCustomFormat.CustomZeroDecimal))
                strFin_Start_Date = CStr(clsInstEMPDBDbase.CRecordset.GetFieldValue("Fin_Start_Date", EMPDataBase.EMPDB.ADODataType.ADODate, EMPDataBase.EMPDB.ADOCustomFormat.CustomDate))
                strFin_End_Date = CStr(clsInstEMPDBDbase.CRecordset.GetFieldValue("Fin_End_Date", EMPDataBase.EMPDB.ADODataType.ADODate, EMPDataBase.EMPDB.ADOCustomFormat.CustomDate))
            Else
                'No Records Found
                Err.Raise(vbObjectError + 20008, "[GenerateDocNo]", "Incorrect Parameters Passed Invoice Number cannot be Generated.")
            End If
            clsInstEMPDBDbase.CRecordset.CloseRecordset() 'Close Recordset
        Else
            'ELSE Raise Error If Wanted Date Not Passed
            Err.Raise(vbObjectError + 20007, "[GenerateDocNo]", "Wanted Date Information not Passed")
        End If
        If Len(Trim(strCheckDOcNo)) > 0 Then 'That is the Document is Made for that Perio
            'Add 1 to it
            strTempSeries = CStr(CInt(strCheckDOcNo) + 1)
            mSaleConfNo = Val(strTempSeries)
            If Len(Trim(strTempSeries)) < 6 Then
                intMaxLoop = 6 - Len(Trim(strTempSeries))
                strZeroSuffix = ""
                For intLoopCounter = 1 To intMaxLoop
                    strZeroSuffix = Trim(strZeroSuffix) & "0"
                Next
            End If
            strTempSeries = strSuffix & strZeroSuffix & strTempSeries
            'UpDate Back New Number
            GenerateInvoiceNo = strTempSeries
        End If
        Exit Function
ErrHandler:
        'Logging the ERROR at Application's Path
        Dim clsErrorInst As New EMPDataBase.EMPDB(gstrUNITID)
        clsErrorInst.CError.RaiseError(20008, "[frmmkttrn0030]", "[GenerateInvoiceNo]", "", "No. Not Generated For DocType = " & pstrInvoiceType & " due to [ " & Err.Description & " ].", My.Application.Info.DirectoryPath, gstrDSNName, gstrDatabaseName)
    End Function
    Public Function CheckDataFromGrin(ByRef pdblDocNo As Double, ByRef pstrCustCode As String) As Boolean
        Dim rsGrnDtl As ClsResultSetDB
        Dim rsSalesDtl As ClsResultSetDB
        Dim strSQL As String
        Dim strItemCode As String
        Dim dblItemQty As Double
        Dim dblRejQty As Double
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        On Error GoTo ErrHandler
        rsSalesDtl = New ClsResultSetDB
        rsSalesDtl.GetResult("Select Item_Code,Sales_Quantity from Sales_dtl where unit_code = '" & gstrUNITID & "'  and doc_No =" & txtChallanNo.Text & " and Location_Code='" & Trim(txtLocationCode.Text) & "'")
        intMaxLoop = rsSalesDtl.GetNoRows : rsSalesDtl.MoveFirst()
        CheckDataFromGrin = False
        For intLoopCounter = 1 To intMaxLoop
            strItemCode = rsSalesDtl.GetValue("Item_code")
            dblItemQty = rsSalesDtl.GetValue("Sales_quantity")
            strSQL = "select a.Doc_No,a.Item_code,a.Rejected_Quantity, a.excess_po_quantity ,"
            strSQL = strSQL & "Despatch_Quantity = isnull(a.Despatch_Quantity,0),"
            strSQL = strSQL & " Inspected_Quantity = isnull(Inspected_Quantity,0),"
            strSQL = strSQL & "RGP_Quantity = isnull(RGP_Quantity,0) from grn_Dtl a,grn_hdr b Where a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND"
            strSQL = strSQL & "a.Doc_type = b.Doc_type And a.Doc_No = b.Doc_No and "
            strSQL = strSQL & "a.From_Location = b.From_Location and a.From_Location ='01R1'"
            strSQL = strSQL & "and a.Rejected_quantity > 0 and b.Vendor_code = '" & pstrCustCode
            strSQL = strSQL & "' and a.Doc_No = " & pdblDocNo & " and a.Item_code = '" & strItemCode & "'"
            rsGrnDtl = New ClsResultSetDB
            rsGrnDtl.GetResult(strSQL)
            dblRejQty = rsGrnDtl.GetValue("Rejected_Quantity") + rsGrnDtl.GetValue("excess_po_Quantity") - rsGrnDtl.GetValue("Despatch_Quantity") - rsGrnDtl.GetValue("Inspected_Quantity") - rsGrnDtl.GetValue("RGP_Quantity")
            If rsGrnDtl.GetNoRows > 0 Then
                If dblItemQty > (dblRejQty) Then
                    MsgBox("Max. Quantity Allowed For Item " & strItemCode & " is " & dblRejQty & ", Quantity Entered in Invoice is : " & dblItemQty)
                    CheckDataFromGrin = False
                    rsGrnDtl.ResultSetClose()
                    Exit Function
                Else
                    CheckDataFromGrin = True
                End If
            End If
            rsGrnDtl.ResultSetClose()
            rsSalesDtl.MoveNext()
        Next
        rsSalesDtl.ResultSetClose()
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Function GetItemGLSL(ByVal InventoryGlGroup As String, ByVal PurposeCode As String) As String
        Dim objRecordSet As New ADODB.Recordset
        Dim strGL As String
        Dim strSL As String
        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close()
        objRecordSet.Open("SELECT invGld_glcode, invGld_slcode FROM fin_InvGLGrpDtl WHERE invGld_prpsCode = '" & PurposeCode & "' AND invGld_invGLGrpId = '" & InventoryGlGroup & "' and UNIT_CODE = '" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
        If objRecordSet.EOF Then
            objRecordSet.Close()
            objRecordSet.Open("SELECT gbl_glCode, gbl_slCode FROM fin_globalGL WHERE gbl_prpsCode = '" & PurposeCode & "' and UNIT_CODE = '" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
            If objRecordSet.EOF Then
                GetItemGLSL = "N"
                MsgBox("GL and SL not defined for Purpose Code: " & PurposeCode, MsgBoxStyle.Information, "eMPro")
                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSet.Close()
                    objRecordSet = Nothing
                End If
                Exit Function
            Else
                strGL = Trim(IIf(IsDBNull(objRecordSet.Fields("gbl_glCode").Value), "", objRecordSet.Fields("gbl_glCode").Value))
                strSL = Trim(IIf(IsDBNull(objRecordSet.Fields("gbl_slCode").Value), "", objRecordSet.Fields("gbl_slCode").Value))
            End If
        Else
            strGL = Trim(IIf(IsDBNull(objRecordSet.Fields("invGld_glcode").Value), "", objRecordSet.Fields("invGld_glcode").Value))
            strSL = Trim(IIf(IsDBNull(objRecordSet.Fields("invGld_slcode").Value), "", objRecordSet.Fields("invGld_slcode").Value))
        End If
        If strGL = "" Then
            GetItemGLSL = "N"
            MsgBox("GL and SL not defined for Purpose Code:" & PurposeCode, MsgBoxStyle.Information, "eMPro")
            If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                objRecordSet.Close()
                objRecordSet = Nothing
            End If
            Exit Function
        End If
        GetItemGLSL = strGL & "»" & strSL
        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
            objRecordSet.Close()
            objRecordSet = Nothing
        End If
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        GetItemGLSL = "N"
        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
            objRecordSet.Close()
            objRecordSet = Nothing
        End If
    End Function
    Private Function GetTaxGlSl(ByVal TaxType As String) As String
        Dim objRecordSet As New ADODB.Recordset
        On Error GoTo ErrHandler
        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close()
        objRecordSet.Open("SELECT tx_glCode, tx_slCode FROM fin_TaxGlRel where tx_rowType = 'ARTAX' AND tx_taxId ='" & TaxType & "' and UNIT_CODE = '" & gstrUNITID & "' ", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
        If objRecordSet.EOF Then
            GetTaxGlSl = "N"
            If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                objRecordSet.Close()
                objRecordSet = Nothing
            End If
            Exit Function
        End If
        GetTaxGlSl = Trim(IIf(IsDBNull(objRecordSet.Fields("tx_glCode").Value), "", objRecordSet.Fields("tx_glCode").Value)) & "»" & Trim(IIf(IsDBNull(objRecordSet.Fields("tx_slCode").Value), "", objRecordSet.Fields("tx_slCode").Value))
        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
            objRecordSet.Close()
            objRecordSet = Nothing
        End If
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        GetTaxGlSl = "N"
        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
            objRecordSet.Close()
            objRecordSet = Nothing
        End If
    End Function
    Public Function CheckExcPriority() As Boolean
        Dim strSQL As String
        Dim strTaxGL As String
        Dim strTaxSL As String
        Dim rsTaxPriority As ClsResultSetDB
        rsTaxPriority = New ClsResultSetDB
        strSQL = "Select * from Tax_PriorityMst WHERE UNIT_CODE = '" & gstrUNITID & "'"
        rsTaxPriority.GetResult(strSQL)
        If rsTaxPriority.GetNoRows > 0 Then
            rsTaxPriority.MoveFirst()
            CheckExcPriority = True
            If Len(Trim(rsTaxPriority.GetValue("VarExPriority1"))) = 0 Then
                If Len(Trim(rsTaxPriority.GetValue("VarExPriority2"))) = 0 Then
                    If Len(Trim(rsTaxPriority.GetValue("VarExPriority3"))) = 0 Then
                        rsTaxPriority.ResultSetClose()
                        CheckExcPriority = False
                        Exit Function
                    End If
                End If
            End If
        Else
            CheckExcPriority = False
        End If
        rsTaxPriority.ResultSetClose()
    End Function
    Public Function ReturnGLSLAccExcPriority(ByRef pintPriority As Object, ByRef pdblamount As Double) As String()
        Dim strSQL As String
        Dim strBalance As String
        Dim strExcGL As String
        Dim strExcSL As String
        Dim StrData(2) As String
        Dim strExcType As String
        Dim rsExGLSLCode As ClsResultSetDB
        Dim rsCheckBalance As ClsResultSetDB
        rsExGLSLCode = New ClsResultSetDB
        strSQL = "Select VarExPriority1,VarExGL1,VarExSL1,VarExPriority2,VarExGL2,VarExSL2,VarExPriority3,VarExGL3,VarExSL3 from Tax_PriorityMst WHERE UNIT_CODE = '" & gstrUNITID & "'"
        rsExGLSLCode.GetResult(strSQL)
        rsExGLSLCode.MoveFirst()
        Select Case pintPriority
            Case 1
                strExcGL = Trim(rsExGLSLCode.GetValue("VarExGL1"))
                strExcSL = Trim(rsExGLSLCode.GetValue("VarExSL1"))
                strExcType = Trim(rsExGLSLCode.GetValue("VarExPriority1"))
                If Len(Trim(strExcGL)) > 0 Then
                    If Len(Trim(strExcSL)) > 0 Then
                        '********To check about in case data is found on first Priority
                        strBalance = "Select sum(isnull(br_amount,0)) as br_amount From fin_balRel where br_UntCodeID = '"
                        strBalance = strBalance & Trim(txtLocationCode.Text) & "' and br_slCode = '" & strExcSL & "' and br_UntCodeID = '" & gstrUNITID & "'"
                        strBalance = strBalance & " and br_glCode = '" & strExcGL & "'"
                        rsCheckBalance = New ClsResultSetDB
                        rsCheckBalance.GetResult(strBalance)
                        If rsCheckBalance.GetNoRows > 0 Then
                            rsCheckBalance.MoveFirst()
                            If Val(rsCheckBalance.GetValue("br_amount")) >= pdblamount Then
                                StrData(0) = strExcGL
                                StrData(1) = strExcSL
                                StrData(2) = strExcType
                                ReturnGLSLAccExcPriority = VB6.CopyArray(StrData)
                            Else
                                ReturnGLSLAccExcPriority = VB6.CopyArray(StrData)
                            End If
                        Else
                            ReturnGLSLAccExcPriority = VB6.CopyArray(StrData)
                        End If
                        rsCheckBalance.ResultSetClose()
                    Else
                        ReturnGLSLAccExcPriority = VB6.CopyArray(StrData)
                    End If
                Else
                    ReturnGLSLAccExcPriority = VB6.CopyArray(StrData)
                End If
            Case 2
                strExcGL = Trim(rsExGLSLCode.GetValue("VarExGL2"))
                strExcSL = Trim(rsExGLSLCode.GetValue("VarExSL2"))
                strExcType = Trim(rsExGLSLCode.GetValue("VarExPriority2"))
                If Len(Trim(strExcGL)) > 0 Then
                    If Len(Trim(strExcSL)) > 0 Then
                        '********To check about in case data is found on first Priority
                        strBalance = "Select sum(isnull(br_amount,0)) as br_amount From fin_balRel where br_UntCodeID = '"
                        strBalance = strBalance & Trim(txtLocationCode.Text) & "' and br_UntCodeID = '" & gstrUNITID & "' and br_slCode = '" & strExcSL & "'"
                        strBalance = strBalance & " and br_glCode = '" & strExcGL & "'"
                        rsCheckBalance = New ClsResultSetDB
                        rsCheckBalance.GetResult(strBalance)
                        If rsCheckBalance.GetNoRows > 0 Then
                            rsCheckBalance.MoveFirst()
                            If rsCheckBalance.GetValue("br_amount") >= pdblamount Then
                                StrData(0) = strExcGL
                                StrData(1) = strExcSL
                                StrData(2) = strExcType
                                ReturnGLSLAccExcPriority = VB6.CopyArray(StrData)
                            Else
                                ReturnGLSLAccExcPriority = VB6.CopyArray(StrData)
                            End If
                        Else
                            ReturnGLSLAccExcPriority = VB6.CopyArray(StrData)
                        End If
                        rsCheckBalance.ResultSetClose()
                    Else
                        ReturnGLSLAccExcPriority = VB6.CopyArray(StrData)
                    End If
                Else
                    ReturnGLSLAccExcPriority = VB6.CopyArray(StrData)
                End If
            Case 3
                strExcGL = Trim(rsExGLSLCode.GetValue("VarExGL3"))
                strExcSL = Trim(rsExGLSLCode.GetValue("VarExSL3"))
                strExcType = Trim(rsExGLSLCode.GetValue("VarExPriority3"))
                If Len(Trim(strExcGL)) > 0 Then
                    If Len(Trim(strExcSL)) > 0 Then
                        '********To check about in case data is found on first Priority
                        strBalance = "Select sum(isnull(br_amount,0)) as br_amount From fin_balRel where br_UntCodeID = '"
                        strBalance = strBalance & Trim(txtLocationCode.Text) & "' and br_slCode = '" & strExcSL & "'"
                        strBalance = strBalance & " and br_glCode = '" & strExcGL & "' and br_UntCodeID = '" & gstrUNITID & "'"
                        rsCheckBalance = New ClsResultSetDB
                        rsCheckBalance.GetResult(strBalance)
                        If rsCheckBalance.GetNoRows > 0 Then
                            rsCheckBalance.MoveFirst()
                            If Val(rsCheckBalance.GetValue("br_amount")) >= pdblamount Then
                                StrData(0) = strExcGL
                                StrData(1) = strExcSL
                                StrData(2) = strExcType
                                ReturnGLSLAccExcPriority = VB6.CopyArray(StrData)
                            Else
                                ReturnGLSLAccExcPriority = VB6.CopyArray(StrData)
                            End If
                        Else
                            ReturnGLSLAccExcPriority = VB6.CopyArray(StrData)
                        End If
                        rsCheckBalance.ResultSetClose()
                    Else
                        ReturnGLSLAccExcPriority = VB6.CopyArray(StrData)
                    End If
                Else
                    ReturnGLSLAccExcPriority = VB6.CopyArray(StrData)
                End If
        End Select
        rsExGLSLCode.ResultSetClose()
    End Function
    Public Function ToGetIteminAcustannex(ByRef pvarArray(,) As Object, ByRef pstrItemCode As Object, ByRef pintArrMaxCount As Short, ByRef pdblReqQuantity As Double) As Object
        Dim intLoopCounter As Short
        On Error GoTo ErrHandler
        For intLoopCounter = 0 To pintArrMaxCount - 1
            If UCase(Trim(pvarArray(0, intLoopCounter))) = UCase(Trim(pstrItemCode)) Then
                pvarArray(1, intLoopCounter) = pvarArray(1, intLoopCounter) + pdblReqQuantity
                pvarArray(2, intLoopCounter) = pvarArray(2, intLoopCounter) - pdblReqQuantity
                ToGetIteminAcustannex = True
            Else
                ToGetIteminAcustannex = False
            End If
        Next
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Public Function InsertUpdateAnnex(ByRef parrCustAnnex As Object, ByRef pstrFinishedItem As Object, ByRef intMaxCount As Short) As Object
        Dim intLoopCount As Short
        Dim intLoopcount1 As Short
        Dim intMaxLoop As Short
        Dim strRef57F4 As String
        Dim strannex As String
        Dim str57f4Date As String
        Dim rsCustAnnex As ClsResultSetDB
        Dim rsVandBom As ClsResultSetDB
        Dim dblbalanceqty As Double
        On Error GoTo ErrHandler
        For intLoopCount = 0 To intMaxCount
            rsVandBom = New ClsResultSetDB
            rsVandBom.GetResult("Select RawMaterial_Code from Vendor_bom where Finish_Product_code = '" & pstrFinishedItem & "' and Vendor_code = '" & strCustCode & "' and rawMaterial_code ='" & parrCustAnnex(0, intLoopCount) & "' and UNIT_CODE = '" & gstrUNITID & "'")
            If rsVandBom.GetNoRows > 0 Then
                strRef57F4 = Replace(ref57f4, "§", "','")
                strRef57F4 = "'" & strRef57F4 & "'"
                strannex = "Select Balance_qty,Ref57f4_No,ref57f4_Date from CustAnnex_HDR "
                strannex = strannex & " WHERE Item_code ='" & parrCustAnnex(0, intLoopCount) & "' and UNIT_CODE = '" & gstrUNITID & "'  and Customer_code ='"
                strannex = strannex & strCustCode & "'"
                If blnFIFOFlag = False Then
                    strannex = strannex & " and Ref57f4_No in (" & strRef57F4 & ") "
                End If
                strannex = strannex & " order by ref57f4_Date"
                rsCustAnnex = New ClsResultSetDB
                rsCustAnnex.GetResult(strannex)
                intMaxLoop = rsCustAnnex.GetNoRows
                rsCustAnnex.MoveFirst()
                For intLoopcount1 = 1 To intMaxLoop
                    If parrCustAnnex(1, intLoopCount) > 0 Then
                        strRef57F4 = rsCustAnnex.GetValue("Ref57f4_No")
                        dblbalanceqty = rsCustAnnex.GetValue("Balance_Qty")
                        str57f4Date = getDateForDB(VB6.Format(rsCustAnnex.GetValue("ref57f4_Date"), gstrDateFormat))
                        mstrAnnex = Trim(mstrAnnex) & " Update CustAnnex_HDR "
                        If dblbalanceqty < parrCustAnnex(1, intLoopCount) Then
                            mstrAnnex = Trim(mstrAnnex) & " Set Balance_Qty = 0 "
                        Else
                            mstrAnnex = Trim(mstrAnnex) & " Set Balance_Qty = Balance_Qty - " & parrCustAnnex(1, intLoopCount)
                        End If
                        mstrAnnex = mstrAnnex & " WHERE Item_code ='" & parrCustAnnex(0, intLoopCount) & "' and UNIT_CODE = '" & gstrUNITID & "'  and Customer_code ='"
                        mstrAnnex = mstrAnnex & strCustCode & "' and Ref57f4_No ='" & strRef57F4 & "' "
                        mstrAnnex = mstrAnnex & "Insert into CustAnnex_dtl (UNIT_CODE,Doc_Ty,"
                        mstrAnnex = mstrAnnex & "Invoice_No,Invoice_Date,ref57f4_Date,Ref57f4_No,"
                        mstrAnnex = mstrAnnex & "Item_Code,Quantity,"
                        mstrAnnex = mstrAnnex & "Customer_Code,"
                        mstrAnnex = mstrAnnex & "Location_Code,Product_Code,Ent_Userid,Ent_dt,"
                        mstrAnnex = mstrAnnex & "Upd_Userid,Upd_dt) values ('" & gstrUNITID & "','O'," & mInvNo & ",GetDate(),'" & str57f4Date & "','"
                        mstrAnnex = mstrAnnex & ref57f4 & "','" & parrCustAnnex(0, intLoopCount) & "'," & parrCustAnnex(1, intLoopCount) & ","
                        mstrAnnex = mstrAnnex & "'" & strCustCode & "',"
                        mstrAnnex = mstrAnnex & "'SMIL','" & pstrFinishedItem & "','" & mP_User & "',GETDATE(),'"
                        mstrAnnex = mstrAnnex & mP_User & "',GETDATE())"
                        If dblbalanceqty < parrCustAnnex(1, intLoopCount) Then
                            mP_Connection.Execute(" insert into tempCustAnnex (Ref57f4_No,Annex_No,ref57f4_date,Item_code,Quantity,Balance_qty,finishedItem,UNIT_CODE) values ('" & strRef57F4 & "',0,'" & str57f4Date & "','" & parrCustAnnex(0, intLoopCount) & "'," & dblbalanceqty & ",0,'" & pstrFinishedItem & "','" & gstrUNITID & "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            parrCustAnnex(1, intLoopCount) = parrCustAnnex(1, intLoopCount) - dblbalanceqty
                        Else
                            mP_Connection.Execute(" insert into tempCustAnnex (Ref57f4_No,Annex_No,ref57f4_date,Item_code,Quantity,Balance_qty,finishedItem,UNIT_CODE) values ('" & strRef57F4 & "',0,'" & str57f4Date & "','" & parrCustAnnex(0, intLoopCount) & "'," & parrCustAnnex(1, intLoopCount) & "," & dblbalanceqty - parrCustAnnex(1, intLoopCount) & ",'" & pstrFinishedItem & "','" & gstrUNITID & "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            parrCustAnnex(1, intLoopCount) = 0
                        End If
                        rsCustAnnex.MoveNext()
                    Else
                        Exit For
                    End If
                Next
            End If
            rsVandBom.ResultSetClose()
        Next
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    '########################################################################################################
    '******************************************PRINTING WITH CRYSTAL REPORT**********************************
    '########################################################################################################
    Sub PrintingInvoiceRPT()
        Dim rsSalesConf As ClsResultSetDB
        Dim rssaledtl As ClsResultSetDB
        Dim rsItembal As ClsResultSetDB
        Dim rsCompany As ClsResultSetDB
        Dim rsSalesChallan As ClsResultSetDB
        Dim rsSalesParameter As New ClsResultSetDB
        Dim strSalesconf As String
        Dim ItemCode As String
        Dim strDrgNo As String
        Dim strAccountCode As String
        Dim strCustRef As String
        Dim StrAmendmentNo As String
        Dim SALEDTL As String
        Dim intRow As Short
        Dim intLoopCount As Short
        Dim salesQuantity As Double
        Dim dblToolCost As Double
        Dim blnCheckToolCost As Boolean
        Dim strItembal As String
        Dim strtoolQuantity As String
        Dim strRetVal As String
        Dim objDrCr As New prj_DrCrNote.cls_DrCrNote(GetServerDate)
        Dim strInvoiceDate As String
        Dim strSelection As String
        Dim varTmp As Object
        Dim varTmp1 As Object
        Dim intNoOfItem As Short
        Dim dblTmpItembal As Double
        Dim dblFinalItembal As Double
        Dim rsItemBalance As ClsResultSetDB
        On Error GoTo Err_Handler
        rsCompany = New ClsResultSetDB
        SALEDTL = "select * from Saleschallan_Dtl where Doc_No =" & txtChallanNo.Text & "  and Location_Code='" & Trim(txtLocationCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "'"
        rsSalesChallan = New ClsResultSetDB
        rssaledtl = New ClsResultSetDB
        rssaledtl.GetResult(SALEDTL, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        strAccountCode = rssaledtl.GetValue("Account_code")
        strCustRef = rssaledtl.GetValue("Cust_ref")
        StrAmendmentNo = rssaledtl.GetValue("Amendment_No")
        strInvoiceDate = getDateForDB(VB6.Format(rssaledtl.GetValue("Invoice_Date"), gstrDateFormat))
        rssaledtl.ResultSetClose()
        strSalesconf = "Select UpdatePO_Flag,UpdateStock_Flag,Stock_Location,OpenningBal, report_filename, Single_Series ,Preprinted_Flag,NoCopies from saleconf where "
        strSalesconf = strSalesconf & "Invoice_type = '" & mstrInvoiceType & "' and UNIT_CODE = '" & gstrUNITID & "'  and sub_type = '"
        strSalesconf = strSalesconf & mstrInvoiceSubType & "' and Location_Code='" & Trim(txtLocationCode.Text) & "' and datediff(dd,'" & strInvoiceDate & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & strInvoiceDate & "')<=0"
        rsSalesConf = New ClsResultSetDB
        rsSalesConf.GetResult(strSalesconf, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        updatePOflag = rsSalesConf.GetValue("UpdatePO_Flag")
        updatestockflag = rsSalesConf.GetValue("UpdateStock_Flag")
        strStockLocation = rsSalesConf.GetValue("Stock_Location")
        mOpeeningBalance = Val(rsSalesConf.GetValue("OpenningBal"))
        mIntNoCopies = rsSalesConf.GetValue("NoCopies")
        mstrReportFilename = rsSalesConf.GetValue("Report_Filename")
        rsSalesConf.ResultSetClose()
        If Len(Trim(strStockLocation)) = 0 Then
            MsgBox("Please Define Stock Location in Sales Configuration. ")
            Exit Sub
        End If
        If Val(txtChallanNo.Text) > 99000000 Then
            '***********To check if Tool Cost Deduction will be done or Not
            rsSalesParameter.GetResult("Select CheckToolAmortisation from Sales_Parameter WHERE UNIT_CODE = '" & gstrUNITID & "'")
            If rsSalesParameter.GetNoRows > 0 Then
                rsSalesParameter.MoveFirst()
                If Len(Trim(rsSalesParameter.GetValue("CheckToolAmortisation"))) = 0 Then
                    MsgBox("First define Check Tool Amortisation in Sales Parameter", MsgBoxStyle.Information, "eMPro")
                    Exit Sub
                End If
                blnCheckToolCost = rsSalesParameter.GetValue("CheckToolAmortisation")
            Else
                MsgBox("No Data Defined in Sales Parameter", MsgBoxStyle.Information, "eMPro")
                Exit Sub
            End If
            SALEDTL = "Select Sales_Quantity,Item_code,Cust_Item_Code,toolcost_amount from sales_Dtl where Doc_No = " & txtChallanNo.Text & " and UNIT_CODE = '" & gstrUNITID & "'  and Location_Code='" & Trim(txtLocationCode.Text) & "'"
            rssaledtl = New ClsResultSetDB
            rssaledtl.GetResult(SALEDTL, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            intRow = rssaledtl.GetNoRows
            rssaledtl.MoveFirst()
            '******Check for balance & despatch in Cust_ord_dtl
            For intLoopCount = 1 To intRow
                ItemCode = rssaledtl.GetValue("Item_code")
                salesQuantity = rssaledtl.GetValue("Sales_quantity")
                strDrgNo = rssaledtl.GetValue("Cust_Item_code")
                dblToolCost = rssaledtl.GetValue("ToolCost_amount")
                rsItembal = New ClsResultSetDB
                rsItembal.GetResult("Select Cur_bal from Itembal_Mst where Item_code = '" & ItemCode & "' and UNIT_CODE = '" & gstrUNITID & "' and Location_code ='" & strStockLocation & "'", ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                If rsItembal.GetNoRows > 0 Then
                    If salesQuantity > rsItembal.GetValue("Cur_Bal") Then
                        MsgBox("Balance for item " & ItemCode & " at Location " & strStockLocation & " not available. ", MsgBoxStyle.Information, "eMPro")
                        rsItembal.ResultSetClose()
                        Exit Sub
                    End If
                Else
                    MsgBox("No Item in ItemMaster for Location " & strStockLocation & ".", MsgBoxStyle.OkOnly, "eMPro")
                    rsItembal.ResultSetClose()
                    Exit Sub
                End If
                rsItembal.ResultSetClose()
                If Len(Trim(strCustRef)) > 0 Then
                    If UCase(mstrInvoiceType) <> "REJ" Then
                        rsItembal = New ClsResultSetDB
                        rsItembal.GetResult("Select balanceQty = order_qty - despatch_Qty,OpenSO from Cust_ord_dtl where account_code ='" & strAccountCode & "' and Cust_ref ='" & strCustRef & "' and Amendment_No = '" & StrAmendmentNo & "' and Item_code ='" & ItemCode & "' and Cust_drgNo ='" & strDrgNo & "' and Active_flag ='A' and UNIT_CODE = '" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                        If rsItembal.GetNoRows > 0 Then
                            'for Open So Check
                            If rsItembal.GetValue("OpenSO") = False Then
                                If salesQuantity > rsItembal.GetValue("BalanceQty") Then
                                    MsgBox("Balance Quantity in SO for item " & ItemCode & " is " & rsItembal.GetValue("BalanceQty") & ".Check Quantity of Item in Challan.", MsgBoxStyle.Information, "eMPro")
                                    rsItembal.ResultSetClose()
                                    Exit Sub
                                End If
                            End If
                        Else
                            MsgBox("No Item (" & strItemCode & ") exist in SO - " & strCustRef & ".", MsgBoxStyle.Information, "eMPro")
                            rsItembal.ResultSetClose()
                            Exit Sub
                        End If
                        rsItembal.ResultSetClose()
                    End If
                End If
                '************To Check for Tool Cost
                If blnCheckToolCost = True Then
                    If dblToolCost > 0 Then
                        strItembal = "select BalanceQty = isnull(proj_qty,0) - isnull(UsedProjQty,0) from Amor_dtl "
                        strItembal = strItembal & " where account_code = '" & strAccountCode & "'"
                        strItembal = strItembal & " and Item_code = '" & ItemCode & "' and UNIT_CODE = '" & gstrUNITID & "'"
                        rsItembal = New ClsResultSetDB
                        rsItembal.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        If rsItembal.GetNoRows > 0 Then
                            strtoolQuantity = CStr(Val(rsItembal.GetValue("BalanceQty")))
                            If salesQuantity > Val(strtoolQuantity) Then
                                If CDbl(strtoolQuantity) = 0 Then
                                    MsgBox("No Balance Available for Item (" & ItemCode & ") and customer Part Code (" & strDrgNo & ") For Amortisation Calculations. ", MsgBoxStyle.OkOnly, "eMPro")
                                Else
                                    MsgBox("Quantity should not be Greater then available Balance Quantity for Amortisarion " & strtoolQuantity, MsgBoxStyle.OkOnly, "eMPro")
                                End If
                                Exit Sub
                            End If
                        Else
                            MsgBox("No Record Available in Tool Amortisation Master for Item (" & ItemCode & ") and customer Part Code (" & strDrgNo & ") For Amortisation Calculations. ", MsgBoxStyle.OkOnly, "eMPro")
                            rsItembal.ResultSetClose()
                            Exit Sub
                        End If
                        rsItembal.ResultSetClose()
                    End If
                End If
                rssaledtl.MoveNext()
            Next
            rssaledtl.ResultSetClose()
            '****To Check in Rejection Invoice if Grin No Exist
            If UCase(mstrInvoiceType) = "REJ" Then
                If Len(Trim(strCustRef)) > 0 Then
                    If CheckDataFromGrin(CDbl(Trim(strCustRef)), strAccountCode) = False Then
                        Exit Sub
                    End If
                End If
            End If
        End If
        If Not (InvoiceGenerationRPT() = True) Then
            Exit Sub
        End If
        If Val(txtChallanNo.Text) > 99000000 Then
            If ConfirmWindow(10344, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                If Len(Find_Value("select doc_no from SalesChallan_dtl where location_code='" & Trim(txtLocationCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "' and doc_no='" & mInvNo & "'")) > 0 Then
                    MsgBox("Next Invoice number already generated." & vbCrLf & "Please skip current no either backward or forward" & vbCrLf & "in Sales Configuration Master Form.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "eMPro")
                    Exit Sub
                End If
                ResetDatabaseConnection()
                mP_Connection.BeginTrans()
                mP_Connection.Execute("set Dateformat 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                mP_Connection.Execute(salesconf, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                If Len(Trim(mstrExcisePriorityUpdationString)) > 0 Then
                    mP_Connection.Execute("update Saleschallan_dtl set Excise_type = '" & mstrExcisePriorityUpdationString & "' where  UNIT_CODE = '" & gstrUNITID & "' and Doc_no = " & txtChallanNo.Text, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                End If
                mP_Connection.Execute(saleschallan, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                If updatePOflag = True Then
                    mP_Connection.Execute(strupdatecustodtdtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                End If
                If updatestockflag = True Then
                    mP_Connection.Execute(strupdateitbalmst, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                End If
                '***********To check if Tool Cost Deduction will be done or Not
                If blnCheckToolCost = True Then
                    If Len(Trim(strUpdateAmorDtl)) > 0 Then
                        mP_Connection.Execute(strUpdateAmorDtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        mP_Connection.Execute(strupdateamordtlbom, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End If
                End If
                If UCase(mstrInvoiceType) = "JOB" And GetBOMCheckFlagValue("BomCheck_Flag") Then
                    mP_Connection.Execute("SET DATEFORMAT 'DMY'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    mP_Connection.Execute(mstrAnnex, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                End If
                If UCase(mstrInvoiceType) = "REJ" Then
                    If Len(Trim(mCust_Ref)) > 0 Then
                        mP_Connection.Execute(strupdateGrinhdr, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End If
                End If
                'Accounts Posting is done here
                If mblnpostinfin = True Then
                    If UCase(Trim(mstrInvoiceType)) <> "REJ" Then
                        strRetVal = objDrCr.SetARInvoiceDocument(gstrUNITID, mstrMasterString, mstrDetailString, prj_GLTransactions.cls_GLTransactions.udtOperationType.optInsert, gstrCONNECTIONSTRING)
                    Else
                        'for Rejection Accounts Posting
                        If MsgBox("No Effects in Accounts.", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "eMPro") = MsgBoxResult.Yes Then
                            strRetVal = "Y"
                        Else
                            strRetVal = objDrCr.SetAPDocument(gstrUNITID, mstrMasterString, mstrDetailString, prj_GLTransactions.cls_GLTransactions.udtOperationType.optInsert, gstrCONNECTIONSTRING)
                        End If
                    End If
                    strRetVal = CheckString(strRetVal)
                Else
                    strRetVal = "Y"
                End If
                If Not strRetVal = "Y" Then
                    MsgBox(strRetVal, MsgBoxStyle.Information, "eMPro")
                    mP_Connection.RollbackTrans()
                    Exit Sub
                Else
                    mP_Connection.CommitTrans()
                    MsgBox("Invoice has been locked successfully with number " & mInvNo, MsgBoxStyle.Information, "eMPro")
                    CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                    CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                End If
                txtChallanNo.Text = CStr(mInvNo)
                txtChallanNo_Validating(txtChallanNo, New System.ComponentModel.CancelEventArgs(False))
                strSelection = "{SalesChallan_Dtl.Location_Code}='" & Trim(txtLocationCode.Text) & "' and {SalesChallan_Dtl.Doc_No} =" & Trim(txtChallanNo.Text) & ""
                strSelection = strSelection & "  and {SalesChallan_Dtl.Invoice_Type} = '" & Trim(mstrInvoiceType) & "' and {SalesChallan_Dtl.UNIT_CODE} = '" & gstrUNITID & "'  and {SalesChallan_Dtl.Sub_Category} = '" & Trim(mstrInvoiceSubType) & "'"
                objRpt.RecordSelectionFormula = strSelection
            End If
        End If
        Exit Sub
Err_Handler:
        If Err.Number = 20545 Then
            Resume Next
        Else
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub
    Public Function InvoiceGenerationRPT() As Boolean
        Dim rsCompMst As ClsResultSetDB
        Dim rsGrnHdr As ClsResultSetDB
        Dim rsSalesConf As ClsResultSetDB
        Dim rsSalesInvoiceDate As ClsResultSetDB
        Dim Phone, Range, RegNo, EccNo, Address, Invoice_Rule As String
        Dim CST, PLA, Fax, EMail, UPST, Division As String
        Dim Commissionerate As String
        Dim strSQL As String
        Dim strCompMst, DeliveredAdd As String
        Dim strGRNDate As String
        Dim strVendorInvNo As String
        Dim strVendorInvDate As String
        Dim strCustRefForGrn As String
        Dim strSuffix As String
        Dim gobjDB As ClsResultSetDB
        Dim rsSalesConf1 As New ADODB.Recordset
        gobjDB = New ClsResultSetDB
        gobjDB.GetResult("SELECT EOU_Flag, CustSupp_Inc,InsExc_Excise,postinfin,Excise_RoundOFF FROM sales_parameter WHERE UNIT_CODE = '" & gstrUNITID & "'")
        If gobjDB.GetValue("EOU_Flag") = True Then
            mStrCustMst = "Select Doc_No,Invoice_type from SalesChallan_Dtl where Invoice_Type <> 'EXP' and Location_Code='" & Trim(txtLocationCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "'"
            mblnEOUUnit = True
        Else
            mStrCustMst = "Select Doc_No,Invoice_type from SalesChallan_Dtl where Location_Code='" & Trim(txtLocationCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "'"
            mblnEOUUnit = False
        End If
        mblnAddCustomerMaterial = gobjDB.GetValue("CustSupp_Inc")
        mblnInsuranceFlag = gobjDB.GetValue("InsExc_Excise")
        mblnpostinfin = gobjDB.GetValue("postinfin")
        mblnExciseRoundOFFFlag = gobjDB.GetValue("Excise_RoundOFF")
        rsSalesConf1.Open("SELECT * FROM SaleConf WHERE  UNIT_CODE = '" & gstrUNITID & "' and Invoice_Type='" & mstrInvoiceType & "' AND Sub_Type ='" & mstrInvoiceSubType & "' AND Location_Code='" & Trim(txtLocationCode.Text) & "' and datediff(dd,getdate(),fin_start_date)<=0  and datediff(dd,fin_end_date,getdate())<=0 ", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
        If Not rsSalesConf1.EOF Then
            mstrPurposeCode = IIf(IsDBNull(rsSalesConf1.Fields("inv_GLD_prpsCode").Value), "", rsSalesConf1.Fields("inv_GLD_prpsCode").Value)
            mblnSameSeries = rsSalesConf1.Fields("Single_Series").Value
            mstrReportFilename = IIf(IsDBNull(rsSalesConf1.Fields("Report_filename").Value), "", rsSalesConf1.Fields("Report_filename").Value)
            If mstrPurposeCode = "" Then
                MsgBox("Please select a Purpose Code in Sales Configuration", MsgBoxStyle.Information, "eMPro")
                mstrPurposeCode = ""
                Exit Function
            End If
        Else
            MsgBox("No record found in Sales Configuration for the selected Location, Invoice Type and Sub-Category", MsgBoxStyle.Information, "eMPro")
            mstrPurposeCode = ""
            Exit Function
        End If
        gobjDB.ResultSetClose()
        rsSalesConf1.Close()
        rsSalesConf1 = Nothing
        On Error GoTo Err_Handler
        rsCompMst = New ClsResultSetDB
        strSQL = "{SalesChallan_Dtl.Location_Code}='" & Trim(txtLocationCode.Text) & "' and {SalesChallan_Dtl.Doc_No} =" & Trim(txtChallanNo.Text) & ""
        strSQL = strSQL & "  and {SalesChallan_Dtl.Invoice_Type} = '" & Trim(mstrInvoiceType) & "' and {SalesChallan_Dtl.UNIT_CODE} = '" & gstrUNITID & "'  and {SalesChallan_Dtl.Sub_Category} = '" & Trim(mstrInvoiceSubType) & "'"
        strCompMst = "Select * from Company_Mst WHERE UNIT_CODE = '" & gstrUNITID & "'"
        rsCompMst.GetResult(strCompMst)
        If rsCompMst.GetNoRows = 1 Then
            RegNo = rsCompMst.GetValue("Reg_NO")
            EccNo = rsCompMst.GetValue("Ecc_No")
            Range = rsCompMst.GetValue("Range_1")
            Phone = rsCompMst.GetValue("Phone")
            Fax = rsCompMst.GetValue("Fax")
            EMail = rsCompMst.GetValue("Email")
            PLA = rsCompMst.GetValue("PLA_No")
            UPST = rsCompMst.GetValue("LST_No")
            CST = rsCompMst.GetValue("CST_No")
            Division = rsCompMst.GetValue("Division")
            Commissionerate = rsCompMst.GetValue("Commissionerate")
            Invoice_Rule = rsCompMst.GetValue("Invoice_Rule")
        End If
        rsCompMst.ResultSetClose()
        If Val(txtChallanNo.Text) > 99000000 Then
            Call InitializeValues()
            Call ValuetoVariables()
            If mblnEOUUnit = True Then
                If mstrInvoiceType <> "EXP" Then
                    If mOpeeningBalance < mAssessableValue Then
                        MsgBox("Opening Balance is Less then Invoice Assessable Value", MsgBoxStyle.Information, "eMPro")
                        InvoiceGenerationRPT = False
                        Exit Function
                    End If
                End If
            End If
            If mblnpostinfin = True Then
                If Not CreateStringForAccounts() Then
                    InvoiceGenerationRPT = False
                    Exit Function
                End If
            End If
            Call updatesalesconfandsaleschallan()
            Call UpdateinSale_Dtl()
            If UCase(mstrInvoiceType) = "REJ" Then
                If Len(Trim(mCust_Ref)) > 0 Then
                    Call UpdateGrnHdr(CDbl(mCust_Ref), mInvNo)
                End If
            End If
        End If
        If UCase(mstrInvoiceType) = "JOB" And GetBOMCheckFlagValue("BomCheck_Flag") Then
            mP_Connection.Execute("DELETE FROM  tempCustAnnex WHERE  UNIT_CODE = '" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords) ' to delete all the records from table before inserting new one for selected invoice
            If BomCheck() = False Then
                InvoiceGenerationRPT = False
                Exit Function
            End If
        End If
        Address = gstr_WRK_ADDRESS1 & gstr_WRK_ADDRESS2
        '*******************To Calculate Value of Delivery Address in Case of Delivery Address requird
        '*******************To Calculate Value of consignee address on Parameter basis
        rsCompMst = New ClsResultSetDB
        rsCompMst.GetResult("Select ConsigneeDetails from Sales_parameter WHERE UNIT_CODE = '" & gstrUNITID & "'")
        If rsCompMst.GetValue("ConsigneeDetails") = False Then
            rsCompMst.ResultSetClose()
            rsCompMst = New ClsResultSetDB
            rsCompMst.GetResult("Select a.* from Customer_Mst a, saleschallan_dtl b where a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND a.Customer_code = b.Account_code and b.Doc_No = " & txtChallanNo.Text & " and b.Location_Code='" & Trim(txtLocationCode.Text) & "'")
            If rsCompMst.GetNoRows > 0 Then
                DeliveredAdd = Trim(rsCompMst.GetValue("Ship_address1"))
                If Len(Trim(DeliveredAdd)) Then
                    DeliveredAdd = Trim(DeliveredAdd) & "," & Trim(rsCompMst.GetValue("Ship_address2"))
                Else
                    DeliveredAdd = Trim(rsCompMst.GetValue("Ship_address2"))
                End If
            End If
            rsCompMst.ResultSetClose()
        Else
            rsCompMst.ResultSetClose()
            rsCompMst = New ClsResultSetDB
            rsCompMst.GetResult("Select ConsigneeAddress1,ConsigneeAddress2,ConsigneeAddress3 from Saleschallan_dtl where Doc_No = " & txtChallanNo.Text & " and Location_Code='" & Trim(txtLocationCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "'")
            If rsCompMst.GetNoRows > 0 Then
                DeliveredAdd = Trim(rsCompMst.GetValue("ConsigneeAddress1"))
                If Len(Trim(DeliveredAdd)) Then
                    DeliveredAdd = Trim(DeliveredAdd) & "," & Trim(rsCompMst.GetValue("ConsigneeAddress2"))
                Else
                    DeliveredAdd = Trim(rsCompMst.GetValue("ConsigneeAddress2"))
                End If
                If Len(Trim(DeliveredAdd)) Then
                    DeliveredAdd = Trim(DeliveredAdd) & "," & Trim(rsCompMst.GetValue("ConsigneeAddress3"))
                Else
                    DeliveredAdd = Trim(rsCompMst.GetValue("ConsigneeAddress3"))
                End If
            End If
            rsCompMst.ResultSetClose()
        End If
        objRpt = frmReportViewer.GetReportDocument()
        frmReportViewer.ShowPrintButton = True
        frmReportViewer.ShowTextSearchButton = True
        frmReportViewer.ShowZoomButton = True
        frmReportViewer.ReportHeader = Me.ctlFormHeader1.HeaderString()
        frmReportViewer.Zoom = 100
        objRpt.Load(My.Application.Info.DirectoryPath & "\Reports\" & mstrReportFilename & ".rpt")
        If UCase(mstrInvoiceType) <> "JOB" Then
            objRpt.DataDefinition.FormulaFields("Category").Text = "'" & mstrInvoiceType & "'"
        End If
        objRpt.DataDefinition.FormulaFields("Registration").Text = "'" & RegNo & "'"
        objRpt.DataDefinition.FormulaFields("ECC").Text = "'" & EccNo & "'"
        objRpt.DataDefinition.FormulaFields("Range").Text = "'" & Range & "'"
        objRpt.DataDefinition.FormulaFields("CompanyName").Text = "'" & gstrCOMPANY & "'"
        objRpt.DataDefinition.FormulaFields("CompanyAddress").Text = "'" & Address & "'"
        objRpt.DataDefinition.FormulaFields("Phone").Text = "'" & Phone & "'"
        objRpt.DataDefinition.FormulaFields("Fax").Text = "'" & Fax & "'"
        If UCase(mstrInvoiceType) <> "JOB" Then
            objRpt.DataDefinition.FormulaFields("EMail").Text = "'" & EMail & "'"
        End If
        objRpt.DataDefinition.FormulaFields("PLA").Text = "'" & PLA & "'"
        objRpt.DataDefinition.FormulaFields("UPST").Text = "'" & UPST & "'"
        objRpt.DataDefinition.FormulaFields("CST").Text = "'" & CST & "'"
        objRpt.DataDefinition.FormulaFields("Division").Text = "'" & Division & "'"
        objRpt.DataDefinition.FormulaFields("commissionerate").Text = "'" & Commissionerate & "'"
        objRpt.DataDefinition.FormulaFields("InvoiceRule").Text = "'" & Invoice_Rule & "'"
        objRpt.DataDefinition.FormulaFields("EOUFlag").Text = "'" & mblnEOUUnit & "'"
        If Val(txtChallanNo.Text) > 99000000 Then
            objRpt.DataDefinition.FormulaFields("DeliveredAt").Text = "' Delivered At '"
            objRpt.DataDefinition.FormulaFields("Address2").Text = "'" & DeliveredAdd & "'"
        Else
            objRpt.DataDefinition.FormulaFields("DeliveredAt").Text = "''"
            objRpt.DataDefinition.FormulaFields("Address2").Text = "''" 'to pass blanck Address in this case will overwrite this Formula written in Crystal Report for else case
        End If
        objRpt.DataDefinition.FormulaFields("InsuranceFlag").Text = "'" & mblnInsuranceFlag & "'"
        objRpt.DataDefinition.FormulaFields("StringYear").Text = "'" & Year(GetServerDate) & "'"
        Dim strInvoiceDate As String
        Dim dblExistingInvNo As Double
        Dim strSql1 As String
        If Val(txtChallanNo.Text) > 99000000 Then
            objRpt.DataDefinition.FormulaFields("InvoiceNo").Text = "'" & mSaleConfNo & "'"
        Else
            strSql1 = "select * from Saleschallan_Dtl where Doc_No =" & Me.txtChallanNo.Text & "  and Location_Code='" & Trim(txtLocationCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "' "
            rsSalesInvoiceDate = New ClsResultSetDB
            rsSalesInvoiceDate.GetResult(strSql1, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            strInvoiceDate = getDateForDB(VB6.Format(rsSalesInvoiceDate.GetValue("Invoice_Date"), gstrDateFormat))
            rsSalesInvoiceDate.ResultSetClose()
            rsSalesConf = New ClsResultSetDB
            rsSalesConf.GetResult("Select Suffix from SaleConf Where  UNIT_CODE = '" & gstrUNITID & "' and invoice_type ='" & mstrInvoiceType & "' AND Location_Code ='" & Trim(txtLocationCode.Text) & "' and datediff(dd,'" & strInvoiceDate & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & strInvoiceDate & "')<=0")
            strSuffix = rsSalesConf.GetValue("Suffix")
            rsSalesConf.ResultSetClose()
            If Len(Trim(strSuffix)) > 0 Then
                If Val(strSuffix) > 0 Then
                    dblExistingInvNo = Val(Mid(CStr(txtChallanNo.Text), Len(Trim(strSuffix)) + 1))
                Else
                    dblExistingInvNo = CDbl(txtChallanNo.Text)
                End If
            Else
                dblExistingInvNo = CDbl(txtChallanNo.Text)
            End If
            objRpt.DataDefinition.FormulaFields("InvoiceNo").Text = "'" & dblExistingInvNo & "'"
        End If
        If UCase(mstrInvoiceType) = "REJ" Then
            rsGrnHdr = New ClsResultSetDB
            strGRNDate = "" : strVendorInvDate = "" : strVendorInvNo = "" : strCustRefForGrn = ""
            rsGrnHdr.GetResult("Select Cust_ref from salesChallan_dtl where  UNIT_CODE = '" & gstrUNITID & "' and doc_No = " & txtChallanNo.Text)
            If rsGrnHdr.GetNoRows > 0 Then
                rsGrnHdr.MoveFirst()
                strCustRefForGrn = rsGrnHdr.GetValue("Cust_ref")
            End If
            rsGrnHdr.ResultSetClose()
            If Len(Trim(strCustRefForGrn)) > 0 Then
                rsGrnHdr = New ClsResultSetDB
                rsGrnHdr.GetResult("select grn_date,Invoice_no,Invoice_date from grn_hdr where  UNIT_CODE = '" & gstrUNITID & "' and From_Location ='01R1' and doc_No = " & strCustRefForGrn)
                If rsGrnHdr.GetNoRows > 0 Then
                    rsGrnHdr.MoveFirst()
                    strGRNDate = IIf(IsDBNull(rsGrnHdr.GetValue("grn_date")), "", VB6.Format(rsGrnHdr.GetValue("grn_date"), gstrDateFormat))
                    strVendorInvDate = IIf(IsDBNull(rsGrnHdr.GetValue("invoice_date")), "", VB6.Format(rsGrnHdr.GetValue("invoice_date")))
                    strVendorInvNo = rsGrnHdr.GetValue("Invoice_No")
                End If
            End If
            rsGrnHdr.ResultSetClose()
            objRpt.DataDefinition.FormulaFields("GrinDate").Text = "'" & strGRNDate & "'"
            objRpt.DataDefinition.FormulaFields("GrinInvoiceNo").Text = "'" & strVendorInvNo & "'"
            objRpt.DataDefinition.FormulaFields("GrinInvoiceDate").Text = "'" & strVendorInvDate & "'"
        End If
        If CBool(Find_Value("select TextPrinting from sales_parameter WHERE UNIT_CODE = '" & gstrUNITID & "'")) Then
        Else
            If mstrReportFilename = "" Then
                MsgBox("No Report filename selected for the invoice. Invoice cannot be printed", MsgBoxStyle.Information, "eMPro")
                Exit Function
            End If
        End If
        objRpt.RecordSelectionFormula = strSQL
        InvoiceGenerationRPT = True
        Exit Function
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function CalculateECESSValue(ByVal pdblTotalExciseValue As Double) As Double
        On Error GoTo ErrHandler
        CalculateECESSValue = ((pdblTotalExciseValue) * Val(lblECESS_Per.Text)) / 100
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function CalculateSECESSValue(ByVal pdblTotalExciseValue As Double) As Double
        On Error GoTo ErrHandler
        CalculateSECESSValue = ((pdblTotalExciseValue) * Val(lblSECESS_Per.Text)) / 100
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Public Function ValidateScheduleQuantity() As Boolean
        On Error GoTo ErrHandler
        Dim strInvoiceType As String
        Dim strInvoiceSubType As String
        Dim rsChallanEntry As ClsResultSetDB
        Dim rsMktDailySchedule As ClsResultSetDB
        Dim varItemCode As Object
        Dim varDrgNo As Object
        Dim varItemQty As Object
        Dim VarDelete As Object
        Dim intRwCount As Short
        Dim ldblNetDispatchQty As Double
        Dim blnDSTracking As Boolean
        Dim varBinQty As Object
        'Validation For Schedule Start From Here
        ValidateScheduleQuantity = True
        If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            strInvoiceType = UCase(Trim(CmbInvType.Text))
            strInvoiceSubType = UCase(Trim(CmbInvSubType.Text))
        ElseIf CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
            rsChallanEntry = New ClsResultSetDB
            rsChallanEntry.GetResult("Select a.Description,a.Sub_Type_Description from SaleConf a,SalesChallan_Dtl b where a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND Doc_No = " & txtChallanNo.Text & " and a.Invoice_Type = b.Invoice_type and a.Sub_type = b.Sub_Category and a.Location_code = b.Location_code and (fin_start_date <= getdate() and fin_end_date >= getdate())")
            strInvoiceType = UCase(rsChallanEntry.GetValue("Description"))
            strInvoiceSubType = UCase(rsChallanEntry.GetValue("sub_type_Description"))
            rsChallanEntry.ResultSetClose()
        End If
        Dim strMakeDate As String
        'If ((UCase(Trim(strInvoiceType)) = "NORMAL INVOICE") And (UCase(CStr((Trim(strInvoiceSubType)) = "FINISHED GOODS")) Or (UCase(Trim(strInvoiceSubType)) = "TRADING GOODS"))) Or (UCase(Trim(strInvoiceType)) = "JOBWORK INVOICE") Or (UCase(Trim(strInvoiceType)) = "EXPORT INVOICE") Then
        If ((UCase(Trim(strInvoiceType)) = "NORMAL INVOICE") And (UCase(CStr((Trim(strInvoiceSubType)) = "FINISHED GOODS")) Or (UCase(Trim(strInvoiceSubType)) = "TRADING GOODS"))) Or (UCase(Trim(strInvoiceType)) = "JOBWORK INVOICE") Or (UCase(Trim(strInvoiceType)) = "EXPORT INVOICE") Or (UCase(Trim(strInvoiceType)) = "TRANSFER INVOICE") Then
            rsChallanEntry = New ClsResultSetDB
            Call rsChallanEntry.GetResult("Select DSWiseTracking From Sales_parameter where UNIT_CODE = '" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If rsChallanEntry.RowCount > 0 Then blnDSTracking = IIf(IsDBNull(rsChallanEntry.GetValue("DSWiseTracking")), False, IIf(rsChallanEntry.GetValue("DSwisetracking") = False, False, True))
            rsChallanEntry.ResultSetClose()
            For intRwCount = 1 To SpChEntry.MaxRows
                varItemCode = Nothing
                Call SpChEntry.GetText(1, intRwCount, varItemCode)
                varDrgNo = Nothing
                Call SpChEntry.GetText(2, intRwCount, varDrgNo)
                varItemQty = Nothing
                Call SpChEntry.GetText(5, intRwCount, varItemQty)
                VarDelete = Nothing
                Call SpChEntry.GetText(14, intRwCount, VarDelete)
                varBinQty = Nothing
                Call SpChEntry.GetText(22, intRwCount, varBinQty)
                '****Delete Flag Check
                If UCase(VarDelete) <> "D" Then
                    If CheckMeasurmentUnit(varItemCode, varItemQty, intRwCount, True) = False Then
                        ValidateScheduleQuantity = False
                        Exit Function
                    End If
                End If
                If UCase(VarDelete) <> "D" Then
                    If CheckMeasurmentUnit(varItemCode, varBinQty, intRwCount, False) = False Then
                        ValidateScheduleQuantity = False
                        Exit Function
                    End If
                End If
                If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD And UCase(VarDelete) <> "D" Then
                    If CheckcustorddtlQty("ADD", CStr(varItemCode), CStr(varDrgNo), CDbl(varItemQty)) = True Then
                        ValidateScheduleQuantity = True
                    Else
                        ValidateScheduleQuantity = False
                        SpChEntry.Col = 5 : SpChEntry.Row = intRwCount : SpChEntry.Focus()
                        Exit Function
                    End If
                ElseIf CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT And UCase(VarDelete) <> "D" Then
                    If CheckcustorddtlQty("EDIT", CStr(varItemCode), CStr(varDrgNo), CDbl(varItemQty)) = True Then
                        ValidateScheduleQuantity = True
                    Else
                        ValidateScheduleQuantity = False
                        SpChEntry.Col = 5 : SpChEntry.Row = intRwCount : SpChEntry.Focus()
                        Exit Function
                    End If
                End If
                If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD And UCase(VarDelete) <> "D" Then
                    ldblNetDispatchQty = GetTotalDispatchQuantityFromDailySchedule(Trim(txtCustCode.Text), Trim(varDrgNo), Trim(varItemCode), getDateForDB(Trim(lblDateDes.Text)), "ADD", 0)
                ElseIf CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT And UCase(VarDelete) <> "D" Then
                    If Len(Trim(varDrgNo)) > 0 Then
                        If UCase(VarDelete) = "A" Then
                            ldblNetDispatchQty = GetTotalDispatchQuantityFromDailySchedule(Trim(txtCustCode.Text), Trim(varDrgNo), Trim(varItemCode), getDateForDB(Trim(lblDateDes.Text)), "EDIT", 0)
                        Else
                            ldblNetDispatchQty = GetTotalDispatchQuantityFromDailySchedule(Trim(txtCustCode.Text), Trim(varDrgNo), Trim(varItemCode), getDateForDB(Trim(lblDateDes.Text)), "EDIT", mdblPrevQty(intRwCount - 1))
                        End If
                    End If
                End If
                If ldblNetDispatchQty <> -1 And UCase(VarDelete) <> "D" Then
                    If Len(Trim(varDrgNo)) > 0 Then
                        If Val(varItemQty) > Val(CStr(ldblNetDispatchQty)) Then
                            ValidateScheduleQuantity = False
                            MsgBox("Quantity should not be Greater then Schedule Quantity " & CStr(ldblNetDispatchQty) & " For Item Code " & varItemCode, MsgBoxStyle.Information, "eMPro")
                            With Me.SpChEntry
                                .Row = intRwCount : .Col = 5 : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                            End With
                            Exit Function
                        Else
                            ValidateScheduleQuantity = True
                            'Updation or Insertion only if  DSWiseTracking Value is true in Sales Parameter
                            If blnDSTracking = False Then
                                If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD And Trim(UCase(VarDelete)) <> "D" Then
                                    rsMktDailySchedule = New ClsResultSetDB
                                    rsMktDailySchedule.GetResult("Select Schedule_quantity from DailyMktSchedule where Account_Code='" & Trim(txtCustCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "'  and  datepart(yyyy,Trans_Date)='" & Year(CDate(getDateForDB(Trim(lblDateDes.Text)))) & "' and datepart(mm,Trans_Date)='" & Month(CDate(getDateForDB(Trim(lblDateDes.Text)))) & "' and datepart(dd,Trans_Date)='" & VB.Day(CDate(getDateForDB(Trim(lblDateDes.Text)))) & "' and Cust_DrgNo ='" & Trim(varDrgNo) & "'and Item_code = '" & varItemCode & "' and Status =1 ")
                                    If rsMktDailySchedule.GetNoRows > 0 Then
                                        mstrUpdDispatchSql = Trim(mstrUpdDispatchSql) & "Update DailyMktSchedule set Despatch_qty ="
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & "isnull(Despatch_Qty,0) + (" & Val(varItemQty) & ")"
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & " Where Account_Code='" & Trim(txtCustCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "' and "
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & " datepart(yyyy,Trans_Date)='" & Year(CDate(getDateForDB(Trim(lblDateDes.Text)))) & "'"
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & " and datepart(mm,Trans_Date)='" & Month(CDate(getDateForDB(Trim(lblDateDes.Text)))) & "'"
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & " and datepart(dd,Trans_Date)='" & VB.Day(CDate(getDateForDB(Trim(lblDateDes.Text)))) & "'"
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & " and Cust_DrgNo ='" & Trim(varDrgNo) & "'and Item_code = '" & varItemCode & "' and Status =1 " & vbCrLf
                                    Else
                                        mstrUpdDispatchSql = Trim(mstrUpdDispatchSql) & " Insert into dailymktschedule "
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & "(Account_Code,Trans_date,cust_drgno,"
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & "Schedule_Flag,Item_Code,Schedule_Quantity,Despatch_qty,"
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & "Status,Ent_UserId,Upd_UserId,Ent_dt,Upd_dt,"
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & "RevisionNo,UNIT_CODE) values ('" & Trim(txtCustCode.Text) & "',"
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & "'" & getDateForDB(dtpDateDesc.Value) & "', '" & Trim(varDrgNo)
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & "',1,'" & varItemCode & "',0," & Val(varItemQty) & ",1,'" & mP_User & "',"
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & "'" & mP_User & "','" & getDateForDB(GetServerDate()) & "','" & getDateForDB(GetServerDate()) & "',0,'" & gstrUNITID & "')" & vbCrLf
                                    End If
                                    rsMktDailySchedule.ResultSetClose()
                                ElseIf CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                                    If Trim(UCase(VarDelete)) <> "D" Then
                                        mstrUpdDispatchSql = Trim(mstrUpdDispatchSql) & "Update DailyMktSchedule set Despatch_qty ="
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & "isnull(Despatch_Qty,0) + (" & Val(varItemQty) & ") - (" & mdblPrevQty(intRwCount - 1) & ")"
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & " Where Account_Code='" & Trim(txtCustCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "' and "
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & " datepart(yyyy,Trans_Date)='" & Year(CDate(getDateForDB(Trim(lblDateDes.Text)))) & "'"
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & " and datepart(mm,Trans_Date)='" & Month(CDate(getDateForDB(Trim(lblDateDes.Text)))) & "'"
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & " and datepart(dd,Trans_Date)='" & VB.Day(CDate(getDateForDB(Trim(lblDateDes.Text)))) & "'"
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & " and Cust_DrgNo ='" & Trim(varDrgNo) & "'and Item_code = '" & varItemCode & "' and Status =1 " & vbCrLf
                                    Else
                                        mstrUpdDispatchSql = Trim(mstrUpdDispatchSql) & "Update DailyMktSchedule set Despatch_qty ="
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & "isnull(Despatch_Qty,0)  - (" & mdblPrevQty(intRwCount - 1) & ")"
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & " Where Account_Code='" & Trim(txtCustCode.Text) & "' and "
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & " datepart(yyyy,Trans_Date)='" & Year(CDate(getDateForDB(Trim(lblDateDes.Text)))) & "'"
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & " and datepart(mm,Trans_Date)='" & Month(CDate(getDateForDB(Trim(lblDateDes.Text)))) & "'"
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & " and datepart(dd,Trans_Date)='" & VB.Day(CDate(getDateForDB(Trim(lblDateDes.Text)))) & "' and UNIT_CODE = '" & gstrUNITID & "' "
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & " and Cust_DrgNo ='" & Trim(varDrgNo) & "'and Item_code = '" & varItemCode & "' and Status =1 " & vbCrLf
                                    End If
                                End If
                            End If
                        End If
                    End If
                Else
                    If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD And UCase(VarDelete) <> "D" Then
                        ldblNetDispatchQty = GetTotalDispatchQuantityFromMonthlySchedule(Trim(txtCustCode.Text), Trim(varDrgNo), Trim(varItemCode), getDateForDB(Trim(lblDateDes.Text)), "ADD", 0)
                    ElseIf CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT And UCase(VarDelete) <> "D" Then
                        If Len(Trim(varDrgNo)) > 0 Then
                            If UCase(VarDelete) = "A" Then
                                ldblNetDispatchQty = GetTotalDispatchQuantityFromMonthlySchedule(Trim(txtCustCode.Text), Trim(varDrgNo), Trim(varItemCode), getDateForDB(Trim(lblDateDes.Text)), "EDIT", 0)
                            Else
                                ldblNetDispatchQty = GetTotalDispatchQuantityFromMonthlySchedule(Trim(txtCustCode.Text), Trim(varDrgNo), Trim(varItemCode), getDateForDB(Trim(lblDateDes.Text)), "EDIT", mdblPrevQty(intRwCount - 1))
                            End If
                        End If
                    End If
                    If ldblNetDispatchQty <> -1 And UCase(VarDelete) <> "D" Then
                        If Len(Trim(varDrgNo)) > 0 Then
                            If Val(varItemQty) > Val(CStr(ldblNetDispatchQty)) Then
                                ValidateScheduleQuantity = False
                                MsgBox("Quantity should not be Greater then Schedule Quantity " & CStr(ldblNetDispatchQty) & " For Item Code " & varItemCode, MsgBoxStyle.Information, "eMPro")
                                With Me.SpChEntry
                                    .Row = intRwCount : .Col = 5 : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                                End With
                                Exit Function
                            Else
                                ValidateScheduleQuantity = True
                                If Val(CStr(Month(CDate(getDateForDB(lblDateDes.Text))))) < 10 Then
                                    strMakeDate = Year(CDate(getDateForDB(lblDateDes.Text))) & "0" & Month(CDate(getDateForDB(lblDateDes.Text)))
                                Else
                                    strMakeDate = Year(CDate(getDateForDB(lblDateDes.Text))) & Month(CDate(getDateForDB(lblDateDes.Text)))
                                End If
                                'Updation or Insertion only if  DSWiseTracking Value is true in Sales Parameter
                                If blnDSTracking = False Then
                                    If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD And UCase(VarDelete) <> "D" Then
                                        mstrUpdDispatchSql = Trim(mstrUpdDispatchSql) & "Update MonthlyMktSchedule set Despatch_qty ="
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & "isnull(Despatch_Qty,0) + (" & Val(varItemQty) & ")"
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & " Where Account_Code='" & Trim(txtCustCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "'  and "
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & " Year_Month=" & Val(Trim(strMakeDate)) & ""
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & " and Cust_DrgNo ='" & Trim(varDrgNo) & "'and Item_code = '" & varItemCode & "' and status =1 " & vbCrLf
                                    ElseIf CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                                        If VarDelete = "A" Then
                                            mstrUpdDispatchSql = Trim(mstrUpdDispatchSql) & "Update MonthlyMktSchedule set Despatch_qty ="
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & "isnull(Despatch_Qty,0) + (" & Val(varItemQty) & ") "
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & " Where Account_Code='" & Trim(txtCustCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "' and "
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & " Year_Month=" & Val(Trim(strMakeDate)) & ""
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & " and Cust_DrgNo ='" & Trim(varDrgNo) & "'and Item_code = '" & varItemCode & "' and status =1 " & vbCrLf
                                        ElseIf VarDelete = "D" Then
                                            mstrUpdDispatchSql = Trim(mstrUpdDispatchSql) & "Update MonthlyMktSchedule set Despatch_qty ="
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & "isnull(Despatch_Qty,0)  - (" & mdblPrevQty(intRwCount - 1) & ") "
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & " Where Account_Code='" & Trim(txtCustCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "'  and "
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & " Year_Month=" & Val(Trim(strMakeDate)) & ""
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & " and Cust_DrgNo ='" & Trim(varDrgNo) & "'and Item_code = '" & varItemCode & "' and status =1 " & vbCrLf
                                        Else
                                            mstrUpdDispatchSql = Trim(mstrUpdDispatchSql) & "Update MonthlyMktSchedule set Despatch_qty ="
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & "isnull(Despatch_Qty,0) + (" & Val(varItemQty) & ") - (" & mdblPrevQty(intRwCount - 1) & ") "
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & " Where Account_Code='" & Trim(txtCustCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "' and "
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & " Year_Month=" & Val(Trim(strMakeDate)) & ""
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & " and Cust_DrgNo ='" & Trim(varDrgNo) & "'and Item_code = '" & varItemCode & "' and status =1 " & vbCrLf
                                        End If
                                    End If
                                End If 'End if For DSTracking Condition
                            End If
                        End If
                    Else
                        If VarDelete <> "D" Then
                            MsgBox("No Schedule Defined For " & varItemCode & " Item.", MsgBoxStyle.Information, "eMPro")
                            ValidateScheduleQuantity = False
                            SpChEntry.Focus()
                            Exit Function
                        End If
                    End If
                End If
            Next
            'Validation For Schedule End Here
            '*********************************************************
        Else 'To Check Decimal places for all type of invoices
            For intRwCount = 1 To SpChEntry.MaxRows
                varItemCode = Nothing
                Call SpChEntry.GetText(1, intRwCount, varItemCode)
                varItemQty = Nothing
                Call SpChEntry.GetText(5, intRwCount, varItemQty)
                VarDelete = Nothing
                Call SpChEntry.GetText(14, intRwCount, VarDelete)
                varBinQty = Nothing
                Call SpChEntry.GetText(22, intRwCount, varBinQty)
                '****Delete Flag Check
                If UCase(VarDelete) <> "D" Then
                    If CheckMeasurmentUnit(varItemCode, varItemQty, intRwCount, True) = False Then
                        ValidateScheduleQuantity = False
                        Exit Function
                    End If
                End If
                If UCase(VarDelete) <> "D" Then
                    If CheckMeasurmentUnit(varItemCode, varBinQty, intRwCount, False) = False Then
                        ValidateScheduleQuantity = False
                        Exit Function
                    End If
                End If
            Next
        End If
        ValidateScheduleQuantity = True
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function GetNextWorkingDay(ByVal pstrDate As String) As String
        Dim rsCalendarDate As New ADODB.Recordset
        Dim strCalDate As String
        Dim strQuery As String
        On Error GoTo ErrHandler
        strQuery = "select dt from calendar_mst where dt > '" & pstrDate & "' and work_flg<>1 and UNIT_CODE = '" & gstrUNITID & "' order by dt"
        If rsCalendarDate.State = ADODB.ObjectStateEnum.adStateOpen Then rsCalendarDate.Close()
        rsCalendarDate.Open(strQuery, mP_Connection, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockBatchOptimistic)
        If rsCalendarDate.EOF Or rsCalendarDate.BOF Or IsDBNull(rsCalendarDate.Fields("dt").Value) Then
            MsgBox("Date in Calendar Master not defined !", MsgBoxStyle.Information, "eMPro")
            GetNextWorkingDay = CStr(-1)
            rsCalendarDate.Close()
            Exit Function
        Else
            rsCalendarDate.MoveFirst()
            GetNextWorkingDay = getDateForDB(VB6.Format(rsCalendarDate.Fields("DT").Value, gstrDateFormat))
        End If
        rsCalendarDate.Close()
        Exit Function
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function ValidateTariffCode(ByVal strItem As String) As Boolean
        Dim rsTarriff As ClsResultSetDB
        Dim strSQL As String
        On Error GoTo ErrHandler
        strSQL = "Select Tariff_Code,item_code from item_mst where item_code= '" & Trim(strItem) & "' and UNIT_CODE = '" & gstrUNITID & "' "
        rsTarriff = New ClsResultSetDB
        rsTarriff.GetResult(strSQL, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsTarriff.GetNoRows > 0 Then
            If Trim(rsTarriff.GetValue("Tariff_code")) <> "" Then
                ValidateTariffCode = True
            Else
                ValidateTariffCode = False
            End If
        End If
        rsTarriff.ResultSetClose()
        Exit Function
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function ValidateTariff_CESS() As Boolean
        Dim intItem As Short
        Dim strItemList As String
        Dim blnExcisableItem As Boolean
        Dim strExciseTax As String
        Dim strECESSTax As String
        Dim rsECESSTax_Percentage As ClsResultSetDB
        Dim rsExcise_Percentage As ClsResultSetDB
        Dim rsPacking As ClsResultSetDB
        Dim strPackingTax As String
        Dim dblExcisePercentage As Double
        Dim dblTemp As Double
        Dim strItem As String
        Dim VarDelete As Object
        Dim dblPackingPercentage As Double
        Dim dblPackingTemp As Double
        On Error GoTo ErrHandler
        For intItem = 1 To SpChEntry.MaxRows
            VarDelete = Nothing
            Call SpChEntry.GetText(14, intItem, VarDelete)
            If UCase(Trim(VarDelete)) <> "D" Then
                SpChEntry.Col = 6 : SpChEntry.Row = intItem
                strPackingTax = Trim(SpChEntry.Text)
                SpChEntry.Col = 7 : SpChEntry.Row = intItem
                strExciseTax = Trim(SpChEntry.Text)
                SpChEntry.Col = 1 : SpChEntry.Row = intItem
                strItem = Trim(SpChEntry.Text)

                If gblnGSTUnit = False Then
                    If Trim(strExciseTax) = "" Then
                        If Len(Trim(strItem)) > 0 Then
                            MsgBox("Excise Tax Can't be blank for Item. Please enter Valid Excise Tax.", MsgBoxStyle.Information, "eMpro")
                            ValidateTariff_CESS = False
                            Exit Function
                        End If
                    End If
                End If
                rsPacking = New ClsResultSetDB
                rsPacking.GetResult("SELECT TxRt_Percentage FROM Gen_TaxRate WHERE TxRt_Rate_No ='" & Trim(strPackingTax) & "' AND Tx_TaxeID='PKT' and UNIT_CODE = '" & gstrUNITID & "' ", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                If rsPacking.GetNoRows > 0 Then
                    If intItem = 1 Then
                        dblPackingTemp = rsPacking.GetValue("TxRt_Percentage")
                    Else
                        dblPackingPercentage = rsPacking.GetValue("TxRt_Percentage")
                    End If
                End If
                If intItem > 1 Then
                    If dblPackingPercentage <> dblPackingTemp Then
                        MsgBox("Packing percentage should be same for all items.", MsgBoxStyle.Information, "eMpro")
                        ValidateTariff_CESS = False
                        Exit Function
                    End If
                End If
                rsPacking.ResultSetClose()
                rsExcise_Percentage = New ClsResultSetDB


                If gblnGSTUnit = False Then
                    rsExcise_Percentage.GetResult("SELECT TxRt_Percentage FROM Gen_TaxRate WHERE TxRt_Rate_No ='" & Trim(strExciseTax) & "' AND Tx_TaxeID='EXC'  and UNIT_CODE = '" & gstrUNITID & "'  ", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                    If rsExcise_Percentage.GetNoRows > 0 Then
                        If intItem = 1 Then
                            dblTemp = rsExcise_Percentage.GetValue("TxRt_Percentage")
                        Else
                            dblExcisePercentage = rsExcise_Percentage.GetValue("TxRt_Percentage")
                        End If
                    End If
                    If intItem > 1 Then
                        If dblExcisePercentage <> dblTemp Then
                            MsgBox("Excise percentage should be same for all items.", MsgBoxStyle.Information, "eMpro")
                            ValidateTariff_CESS = False
                            Exit Function
                        End If
                    End If
                    If rsExcise_Percentage.GetValue("TxRt_Percentage") <> 0 Then
                        blnExcisableItem = True
                        SpChEntry.Col = 1 : SpChEntry.Row = intItem
                        If ValidateTariffCode(Trim(SpChEntry.Text)) = False Then
                            If Len(strItemList) = 0 Then
                                strItemList = Trim(SpChEntry.Text)
                            Else
                                strItemList = strItemList & "," & Trim(SpChEntry.Text)
                            End If
                        End If
                    End If
                    rsExcise_Percentage.ResultSetClose()
                End If
            End If
        Next intItem
        If Len(strItemList) > 1 Then
            MsgBox("Tariff Code is required for Item(s)-- " & strItemList, MsgBoxStyle.Information, "eMpro")
            ValidateTariff_CESS = False
            Exit Function
        End If
        '''***** ECESS can't be zero for excisable items.


        If gblnGSTUnit = False Then
            strECESSTax = (Me.txtECESS.Text)
            If Trim(strECESSTax) = "" Then
                MsgBox("Ecess Can't be blank. Please enter Valid Ecess.", MsgBoxStyle.Information, "eMpro")
                ValidateTariff_CESS = False
                Exit Function
            End If
        End If

        If gblnGSTUnit = False Then
            rsECESSTax_Percentage = New ClsResultSetDB
            rsECESSTax_Percentage.GetResult("SELECT TxRt_Percentage FROM Gen_TaxRate WHERE TxRt_Rate_No ='" & Trim(strECESSTax) & "' AND Tx_TaxeID='ECS' and UNIT_CODE = '" & gstrUNITID & "' ", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If blnExcisableItem = True Then
                '------------------Satvir Handa------------------------
                'If rsECESSTax_Percentage.GetValue("TxRt_Percentage") = 0 Then
                '    MsgBox("Ecess can not be zero for Excisable Items.", MsgBoxStyle.Information, "eMpro")
                '    ValidateTariff_CESS = False
                '    Exit Function
                'End If
                '------------------Satvir Handa------------------------
            Else
                If rsECESSTax_Percentage.GetValue("TxRt_Percentage") <> 0 Then
                    MsgBox("Ecess can not be Charged for Non Excisable Items.", MsgBoxStyle.Information, "eMpro")
                    ValidateTariff_CESS = False
                    Me.txtECESS.Text = ""
                    Me.txtECESS.Focus()
                    Exit Function
                End If
            End If
            rsECESSTax_Percentage.ResultSetClose()
        End If



        ValidateTariff_CESS = True
        Exit Function
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function CheckSOType(ByVal Row As Short) As String
        On Error GoTo ErrHandler
        Dim RSchkSoType As ClsResultSetDB
        Dim strSQL As String
        RSchkSoType = New ClsResultSetDB
        strSQL = "select Po_Type from Cust_Ord_Hdr where Account_code = '" & Trim(txtCustCode.Text) & "' and Cust_Ref='" & Trim(txtRefNo.Text) & "'"
        strSQL = strSQL & " and Amendment_No='" & Trim(txtAmendNo.Text) & "' and UNIT_CODE = '" & gstrUNITID & "'"
        RSchkSoType.GetResult(strSQL, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        If RSchkSoType.GetNoRows > 0 Then
            CheckSOType = Trim(RSchkSoType.GetValue("PO_Type"))
        Else
            CheckSOType = ""
        End If
        RSchkSoType.ResultSetClose()
        RSchkSoType = Nothing
        Exit Function
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub CmdGrpChEnt_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.UCbtndgrp.ButtonClickEventArgs) Handles CmdGrpChEnt.ButtonClick
        On Error GoTo ErrHandler
        Dim strSalesChallan As String
        Dim updateSalesChallan As String
        Dim strSalesDtl As String
        Dim strSalesDtlDelete As String
        Dim strCurrency As String
        Dim Description As String
        Dim intLoopCount As Short
        Dim varQuantity As Object
        Dim varDrgNo As Object
        Dim varItemCode As Object
        Dim varRate As Object
        Dim varCustMtrl As Object
        Dim varPacking As Object
        Dim varOthers As Object
        Dim varFromBox As Object
        Dim VarToBox As Object
        Dim VarDelete As Object
        Dim PresQty As Object
        Dim rsSalesChallandtl As ClsResultSetDB
        Dim rsInvoiceType As ClsResultSetDB
        Dim rsECess As ClsResultSetDB
        Dim intLoopCounter As Short
        Dim strChallanNo As String
        Dim rsReportName As String
        Dim intDecimalPlaces As Short
        Dim intLoop As Short
        Dim strMakeDate As String
        Dim blnDSTracking As Boolean
        Dim strsql As String

        Select Case e.Button
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD
                Call EnableControls(True, Me, True)
                '101188073
                TaxesEnableDisable(txtSDTType)
                TaxesHelpEnableDisable(cmdSDTax_Help)
                TaxesLabelEnableDisable(lblSDTax_Per)
                TaxesEnableDisable(txtSaleTaxType)
                TaxesHelpEnableDisable(CmdSaleTaxType)
                TaxesLabelEnableDisable(lblSaltax_Per)
                TaxesEnableDisable(txtAddVAT)
                TaxesHelpEnableDisable(cmdAddVAT)
                TaxesLabelEnableDisable(lblAddVAT)
                TaxesEnableDisable(txtSurchargeTaxType)
                TaxesHelpEnableDisable(cmdSurchargeTaxCode)
                TaxesLabelEnableDisable(lblSurcharge_Per)
                TaxesEnableDisable(txtECESS)
                TaxesHelpEnableDisable(cmdECESSCode)
                TaxesLabelEnableDisable(lblECESS_Per)
                TaxesEnableDisable(txtSECESS)
                TaxesHelpEnableDisable(cmdSECESSCode)
                TaxesLabelEnableDisable(lblSECESS_Per)
                '101188073
                lblSaltax_Per.Text = "0.00"
                Call SelectChallanNoFromSalesChallanDtl()
                txtChallanNo.Enabled = False : txtChallanNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                CmdChallanNo.Enabled = False : txtChallanNo.Enabled = False
                txtRefNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdempHelpRefNo.Enabled = False
                lblLocCodeDes.Text = "" : lblCustCodeDes.Text = "" : mCurrencyCode = ""
                lblCurrencyDes.Text = gstrCURRENCYCODE
                lblExchangeRateValue.Text = CStr(1.0#)
                Me.SpChEntry.Enabled = True
                If blnEOU_FLAG = False Then
                    '10569249 
                    If (UCase(Trim(gstrUNITID)) = "SMT") Or (UCase(Trim(gstrUNITID)) = "SML") Or (UCase(Trim(gstrUNITID)) = "SMM") Then
                        For intLoopCounter = 0 To CmbInvType.Items.Count - 1 'Selecting transfer Invoice as default
                            If UCase(Trim(ObsoleteManagement.GetItemString(CmbInvType, intLoopCounter))) = "TRANSFER INVOICE" Then
                                Exit For
                            End If
                        Next
                    Else
                        For intLoopCounter = 0 To CmbInvType.Items.Count - 1 'Selecting Normal Invoice as default
                            If UCase(Trim(ObsoleteManagement.GetItemString(CmbInvType, intLoopCounter))) = "NORMAL INVOICE" Then
                                Exit For
                            End If
                        Next
                    End If
                    '10569249 
                    CmbInvType.SelectedIndex = intLoopCounter
                    For intLoopCounter = 0 To CmbInvSubType.Items.Count - 1 'Selecting Finished Goods as default
                        If UCase(Trim(ObsoleteManagement.GetItemString(CmbInvSubType, intLoopCounter))) = "FINISHED GOODS" Then
                            Exit For
                        End If
                    Next
                    CmbInvSubType.SelectedIndex = intLoopCounter
                    CmbTransType.SelectedIndex = 0
                Else
                    For intLoopCounter = 0 To CmbInvType.Items.Count - 1 'Selecting Normal Invoice as default
                        If UCase(Trim(ObsoleteManagement.GetItemString(CmbInvType, intLoopCounter))) = "EXPORT INVOICE" Then
                            Exit For
                        End If
                    Next
                    CmbInvType.SelectedIndex = intLoopCounter
                    CmbTransType.SelectedIndex = 0
                End If
                With Me.SpChEntry
                    .MaxRows = 1
                    .Row = 1 : .Row2 = 1 : .Col = 1 : .Col2 = 12 : .BlockMode = True : .Text = "" : .Lock = False : .BlockMode = False
                End With
                '10569249 
                If (UCase(Trim(gstrUNITID)) = "SMT") Or (UCase(Trim(gstrUNITID)) = "SML") Or (UCase(Trim(gstrUNITID)) = "SMM") Then
                    If UCase(CStr(Trim(CmbInvType.Text))) <> "TRANSFER INVOICE" Then
                        txtRefNo.Enabled = False
                        txtRefNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        cmdempHelpRefNo.Enabled = False
                    Else
                        txtRefNo.Enabled = True
                        txtRefNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        cmdempHelpRefNo.Enabled = True
                    End If
                Else

                    If UCase(CStr(Trim(CmbInvType.Text))) <> "NORMAL INVOICE" Then
                        txtRefNo.Enabled = False
                        txtRefNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        cmdempHelpRefNo.Enabled = False
                    Else
                        txtRefNo.Enabled = True
                        txtRefNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        cmdempHelpRefNo.Enabled = True
                    End If
                End If
                '10569249 
                'In Add Mode Enable Combo Of Invoice Type and Inv. Sub type
                CmbInvType.Visible = True : CmbInvSubType.Visible = True
                lblInvSubType.Visible = True : lblInvType.Visible = True
                'Get Server Date
                lblDateDes.Text = VB6.Format(GetServerDate(), gstrDateFormat)
                With dtpDateDesc
                    .Value = GetServerDate()
                    .Visible = True 'Show DatePicker
                End With
                'Set The Column Length in Spread
                Call SetMaxLengthInSpread(0)
                'Set Cell Type In Spread
                Call ChangeCellTypeStaticText()
                lblRGPDes.Text = ""
                txtLocationCode.Text = mstrLocation
                Call SelectDescriptionForField("Description", "Location_Code", "Location_Mst", lblLocCodeDes, (txtLocationCode.Text))
                If Len(gStrLocationCode) > 0 Then
                    txtLocationCode.Text = gStrLocationCode
                    txtLocationCode_Validating(txtLocationCode, New System.ComponentModel.CancelEventArgs(False))
                End If
                If Len(gStrCustomerCode) > 0 Then
                    txtCustCode.Text = gStrCustomerCode
                    txtCustCode_Validating(txtCustCode, New System.ComponentModel.CancelEventArgs(False))
                End If
                If Len(gStrVehicleNo) > 0 Then
                    txtVehNo.Text = gStrVehicleNo
                End If

                'rsECess = New ClsResultSetDB
                'rsECess.GetResult("Select TxRt_Rate_No,TxRt_Percentage from Gen_TaxRate where tx_TaxeID ='ECS' and TxRt_Percentage > 0 and UNIT_CODE = '" & gstrUNITID & "'")
                'If Not rsECess.EOFRecord Then
                '    rsECess.MoveFirst()
                '    txtECESS.Text = rsECess.GetValue("TxRt_Rate_No")
                '    lblECESS_Per.Text = rsECess.GetValue("TxRt_Percentage")
                'End If
                'rsECess.ResultSetClose()
                'rsECess = New ClsResultSetDB
                'rsECess.GetResult("Select TxRt_Rate_No,TxRt_Percentage from Gen_TaxRate where tx_TaxeID ='ECSSH' and TxRt_Percentage > 0  and UNIT_CODE = '" & gstrUNITID & "'")
                'If Not rsECess.EOFRecord Then
                '    rsECess.MoveFirst()
                '    txtSECESS.Text = rsECess.GetValue("TxRt_Rate_No")
                '    lblSECESS_Per.Text = rsECess.GetValue("TxRt_Percentage")
                'End If
                'rsECess.ResultSetClose()

                '------------------Satvir Handa------------------------
                If Not gblnGSTUnit Then
                    rsECess = New ClsResultSetDB
                    rsECess.GetResult("Select TxRt_Rate_No,TxRt_Percentage from Gen_TaxRate where tx_TaxeID ='ECS' and DEFAULT_FOR_INVOICE =1 And Unit_Code='" & gstrUnitId & "' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))")
                    If Not rsECess.EOFRecord Then
                        rsECess.MoveFirst()
                        txtECESS.Text = rsECess.GetValue("TxRt_Rate_No")
                        lblECESS_Per.Text = rsECess.GetValue("TxRt_Percentage")
                    End If
                    rsECess.ResultSetClose()

                    rsECess = New ClsResultSetDB
                    rsECess.GetResult("Select TxRt_Rate_No,TxRt_Percentage from Gen_TaxRate where tx_TaxeID ='ECSSH' and DEFAULT_FOR_INVOICE =1 And Unit_Code='" & gstrUnitId & "' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))")
                    If Not rsECess.EOFRecord Then
                        rsECess.MoveFirst()
                        txtSECESS.Text = rsECess.GetValue("TxRt_Rate_No")
                        lblSECESS_Per.Text = rsECess.GetValue("TxRt_Percentage")
                    End If
                    rsECess.ResultSetClose()
                End If
                '------------------Satvir Handa------------------------

                If txtRefNo.Enabled Then txtRefNo.Focus()
                txtLocationCode.Focus()
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT
                Call EnableControls(False, Me)
                rsSalesChallandtl = New ClsResultSetDB
                rsSalesChallandtl.GetResult("select Invoice_type from Saleschallan_dtl where  UNIT_CODE = '" & gstrUnitId & "' and doc_no = " & txtChallanNo.Text, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                If rsSalesChallandtl.GetValue("Invoice_type") <> "JOB" Then
                    ctlInsurance.Enabled = True
                    ctlInsurance.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                End If
                If rsSalesChallandtl.GetValue("Invoice_type") <> "JOB" And rsSalesChallandtl.GetValue("Invoice_type") <> "TRF" Then
                    '101188073
                    TaxesEnableDisable(txtSDTType)
                    TaxesHelpEnableDisable(cmdSDTax_Help)
                    TaxesLabelEnableDisable(lblSDTax_Per)
                    '101188073
                End If
                TaxesLabelEnableDisable(lblECESS_Per)
                If UCase(rsSalesChallandtl.GetValue("Invoice_type")) = "INV" Or UCase(CStr(rsSalesChallandtl.GetValue("Invoice_type") = "REJ")) Or UCase(CStr(rsSalesChallandtl.GetValue("Invoice_type") = "EXP")) Then
                    '101188073
                    TaxesEnableDisable(txtSaleTaxType)
                    TaxesHelpEnableDisable(CmdSaleTaxType)
                    TaxesLabelEnableDisable(lblSaltax_Per)
                    TaxesEnableDisable(txtAddVAT)
                    TaxesHelpEnableDisable(cmdAddVAT)
                    TaxesLabelEnableDisable(lblAddVAT)
                    '101188073
                    TaxesLabelEnableDisable(lblSurcharge_Per)
                    ctlInsurance.Enabled = True
                    ctlInsurance.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                End If
                txtFreight.Enabled = True
                txtFreight.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                TaxesEnableDisable(txtSurchargeTaxType)
                TaxesHelpEnableDisable(cmdSurchargeTaxCode)


                txtRemarks.Enabled = True : txtRemarks.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                txtSRVDI.Enabled = True : txtSRVDI.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : cmdhelpSRVDI.Enabled = True
                txtSRVLoc.Enabled = True : txtSRVLoc.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                txtUsLoc.Enabled = True : txtUsLoc.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                txtSchTime.Enabled = True : txtSchTime.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)

                TaxesEnableDisable(txtECESS)
                TaxesHelpEnableDisable(cmdECESSCode)
                TaxesEnableDisable(txtSECESS)
                TaxesHelpEnableDisable(cmdSECESSCode)
                TaxesLabelEnableDisable(lblSECESS_Per)
                SpChEntry.Enabled = True
                SpChEntry.Row = 1 : SpChEntry.Row2 = SpChEntry.MaxRows : SpChEntry.Col = 0 : SpChEntry.Col2 = 12
                SpChEntry.BlockMode = True : SpChEntry.Lock = False : SpChEntry.BlockMode = False
                SpChEntry.Row = 1 : SpChEntry.Row2 = SpChEntry.MaxRows : SpChEntry.Col = 0 : SpChEntry.Col2 = 22
                SpChEntry.BlockMode = True : SpChEntry.Lock = False : SpChEntry.BlockMode = False
                If rsSalesChallandtl.GetValue("Invoice_type") = "SMP" Or rsSalesChallandtl.GetValue("Invoice_type") = "CSM" Then
                    SpChEntry.Row = 1 : SpChEntry.Row2 = SpChEntry.MaxRows : SpChEntry.Col = 15 : SpChEntry.Col2 = 15
                    SpChEntry.BlockMode = True : SpChEntry.Lock = False : SpChEntry.BlockMode = False
                Else
                    SpChEntry.Row = 1 : SpChEntry.Row2 = SpChEntry.MaxRows : SpChEntry.Col = 15 : SpChEntry.Col2 = 15
                    SpChEntry.BlockMode = True : SpChEntry.Lock = True : SpChEntry.BlockMode = False
                End If
                rsSalesChallandtl.ResultSetClose()
                intDecimalPlaces = ToGetDecimalPlaces(mCurrencyCode)
                Call SetMaxLengthInSpread(intDecimalPlaces)
                Call ChangeCellTypeStaticText()
                ReDim mdblPrevQty(SpChEntry.MaxRows - 1) ' To get value of Quantity in Array for updation in despatch
                For intLoop = 1 To SpChEntry.MaxRows
                    Call SpChEntry.GetText(5, intLoop, mdblPrevQty(intLoop - 1))
                Next
                With SpChEntry
                    .Row = 1 : .Col = 5 : .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                End With
                System.Windows.Forms.Application.DoEvents()
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE
                'for checking the zero no of rows in case of
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        If Not ValidatebeforeSave("ADD") Then
                            gblnCancelUnload = True
                            gblnFormAddEdit = True
                            Exit Sub
                        End If
                        'Check Quantity Schedule
                        If QuantityCheck() Then
                            Exit Sub
                        End If
                        If ValidateTariff_CESS() = False Then Exit Sub
                        If UCase(CmbInvType.Text) = "REJECTION" Then
                            If Len(Trim(txtRefNo.Text)) > 0 Then
                                If ItemQtyCaseRejGrin() = False Then
                                    Exit Sub
                                End If
                            End If
                        End If
                        If UCase(CmbInvType.Text) = "EXPORT INVOICE" Then
                            If CheckExchangeRate() = False Then
                                Exit Sub
                            End If
                        End If
                        Call SelectChallanNoFromSalesChallanDtl()
                        gStrLocationCode = txtLocationCode.Text
                        gStrCustomerCode = txtCustCode.Text
                        gStrVehicleNo = txtVehNo.Text
                        '101188073
                        If gblnGSTUnit Then
                            'If Not ValidateGSTTaxes() Then Exit Sub
                            If Not SaveDataGST("ADD") Then Exit Sub
                        Else
                            If Not SaveData("ADD") Then Exit Sub
                        End If
                        '101188073
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        If Not ValidatebeforeSave("EDIT") Then
                            gblnCancelUnload = True
                            gblnFormAddEdit = True
                            Exit Sub
                        End If
                        'Check Quantity Schedule
                        If QuantityCheck() Then
                            Exit Sub
                        End If
                        If ValidateTariff_CESS() = False Then Exit Sub
                        rsInvoiceType = New ClsResultSetDB
                        rsInvoiceType.GetResult("select Invoice_type from Saleschallan_dtl where  UNIT_CODE = '" & gstrUnitId & "'  and doc_no = " & txtChallanNo.Text)
                        If UCase(rsInvoiceType.GetValue("Invoice_type")) = "REJ" Then
                            If Len(Trim(txtRefNo.Text)) > 0 Then
                                If ItemQtyCaseRejGrin() = False Then
                                    rsInvoiceType.ResultSetClose()
                                    Exit Sub
                                End If
                            End If
                        End If
                        If UCase(rsInvoiceType.GetValue("Invoice_type")) = "EXP" Then
                            If CheckExchangeRate() = False Then
                                rsInvoiceType.ResultSetClose()
                                Exit Sub
                            End If
                        End If
                        rsInvoiceType.ResultSetClose()
                        gStrLocationCode = txtLocationCode.Text
                        gStrCustomerCode = txtCustCode.Text
                        gStrVehicleNo = txtVehNo.Text
                        '101188073
                        If gblnGSTUnit Then
                            'If Not ValidateGSTTaxes() Then Exit Sub
                            If Not SaveDataGST("EDIT") Then Exit Sub
                        Else
                            If Not SaveData("EDIT") Then Exit Sub
                        End If
                        '101188073
                End Select
                Call ConfirmWindow(10049, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                Me.CmdGrpChEnt.Revert()
                Me.CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                gblnCancelUnload = False : gblnFormAddEdit = False
                Call EnableControls(False, Me)
                TaxesLabelEnableDisable(lblSDTax_Per, True)
                TaxesLabelEnableDisable(lblSaltax_Per, True)
                TaxesLabelEnableDisable(lblSurcharge_Per, True)
                TaxesLabelEnableDisable(lblECESS_Per, True)
                TaxesLabelEnableDisable(lblSECESS_Per, True)
                SpChEntry.Enabled = True
                SpChEntry.Row = 1 : SpChEntry.Row2 = SpChEntry.MaxRows : SpChEntry.Col = 0 : SpChEntry.Col2 = 12
                SpChEntry.BlockMode = True : SpChEntry.Lock = True : SpChEntry.BlockMode = False
                '****In View Mode Disable Combo Of Invoice Type and Inv. Sub type
                CmbInvType.Visible = False : CmbInvSubType.Visible = False
                lblInvSubType.Visible = False : lblInvType.Visible = False
                txtLocationCode.Enabled = True : txtLocationCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                txtChallanNo.Enabled = True : txtChallanNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                CmdLocCodeHelp.Enabled = True : CmdChallanNo.Enabled = True
                lblDateDes.Text = VB6.Format(dtpDateDesc.Value, gstrDateFormat)
                dtpDateDesc.Visible = False
                If txtLocationCode.Enabled Then
                    If Len(Trim(mstrLocation)) > 0 Then
                        txtLocationCode.Text = mstrLocation
                    End If
                    txtLocationCode.Focus()
                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL
                Call frmMKTTRN0016_KeyPress(Me, New System.Windows.Forms.KeyPressEventArgs(Chr(System.Windows.Forms.Keys.Escape)))
                chkDTRemoval.Enabled = True
                chkDTRemoval.CheckState = System.Windows.Forms.CheckState.Unchecked
                dtpRemoval.Enabled = False
                dtpRemovalTime.Enabled = False
                dtpRemoval.Value = GetServerDate()
                dtpRemovalTime.Value = GetServerDate()
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE
                If ConfirmWindow(10054, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                    mstrUpdDispatchSql = ""
                    blnDSTracking = Find_Value("select dswisetracking from sales_parameter WHERE UNIT_CODE = '" & gstrUnitId & "'")
                    If blnDSTracking = False Then
                        If Val(CStr(Month(CDate(getDateForDB(lblDateDes.Text))))) < 10 Then
                            strMakeDate = Year(CDate(getDateForDB(lblDateDes.Text))) & "0" & Month(CDate(getDateForDB(lblDateDes.Text)))
                        Else
                            strMakeDate = Year(CDate(getDateForDB(lblDateDes.Text))) & Month(CDate(getDateForDB(lblDateDes.Text)))
                        End If
                        For intLoopCount = 1 To SpChEntry.MaxRows
                            varDrgNo = Nothing
                            Call Me.SpChEntry.GetText(2, intLoopCount, varDrgNo)
                            varItemCode = Nothing
                            Call Me.SpChEntry.GetText(1, intLoopCount, varItemCode)
                            PresQty = Nothing
                            Call Me.SpChEntry.GetText(5, intLoopCount, PresQty)
                            mstrUpdDispatchSql = Trim(mstrUpdDispatchSql) & "Update DailyMktSchedule set Despatch_qty ="
                            mstrUpdDispatchSql = mstrUpdDispatchSql & "isnull(Despatch_Qty,0) -  " & Val(PresQty) & ",Schedule_flag =1 "
                            mstrUpdDispatchSql = mstrUpdDispatchSql & " Where unit_code = '" & gstrUnitId & "' and Account_Code='" & Trim(txtCustCode.Text) & "' and "
                            mstrUpdDispatchSql = mstrUpdDispatchSql & " datepart(yyyy,Trans_Date)='" & Year(CDate(getDateForDB(Trim(lblDateDes.Text)))) & "'"
                            mstrUpdDispatchSql = mstrUpdDispatchSql & " and datepart(mm,Trans_Date)='" & Month(CDate(getDateForDB(Trim(lblDateDes.Text)))) & "'"
                            mstrUpdDispatchSql = mstrUpdDispatchSql & " and datepart(dd,Trans_Date)='" & VB.Day(CDate(getDateForDB(Trim(lblDateDes.Text)))) & "'"
                            mstrUpdDispatchSql = mstrUpdDispatchSql & " and Cust_DrgNo ='" & Trim(varDrgNo) & "' and Item_code = '" & varItemCode & "' and Status =1" & vbCrLf
                            mstrUpdDispatchSql = Trim(mstrUpdDispatchSql) & " Update MonthlyMktSchedule set Despatch_qty ="
                            mstrUpdDispatchSql = mstrUpdDispatchSql & "isnull(Despatch_Qty,0)  - " & Val(PresQty) & ",Schedule_flag =1 "
                            mstrUpdDispatchSql = mstrUpdDispatchSql & " Where Account_Code='" & Trim(txtCustCode.Text) & "' and UNIT_CODE = '" & gstrUnitId & "' and "
                            mstrUpdDispatchSql = mstrUpdDispatchSql & " Year_Month=" & Val(Trim(strMakeDate)) & ""
                            mstrUpdDispatchSql = mstrUpdDispatchSql & " and Cust_DrgNo ='" & Trim(varDrgNo) & "' and Item_code = '" & varItemCode & "' and Status =1 " & vbCrLf
                        Next
                    End If
                    Call DeleteRecords()
                    ResetDatabaseConnection()
                    mP_Connection.BeginTrans()
                    '10736222
                    Dim objCmd As New ADODB.Command

                    With objCmd
                        .ActiveConnection = mP_Connection
                        .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                        .CommandText = "USP_SAVE_CT2_INVOICE_KNOCKOFFDTL"
                        .CommandTimeout = 0

                        .Parameters.Append(.CreateParameter("@MODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, "D"))
                        .Parameters.Append(.CreateParameter("@UNIT_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, gstrUnitId))
                        .Parameters.Append(.CreateParameter("@INVOICE_NO", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, , txtChallanNo.Text.Trim))
                        .Parameters.Append(.CreateParameter("@IP_ADDRESS", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, gstrIpaddressWinSck))
                        .Parameters.Append(.CreateParameter("@ERRMSG", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInputOutput, 8000, ""))
                        .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End With

                    If objCmd.Parameters(objCmd.Parameters.Count - 1).Value.ToString().Trim.Length <> 0 Then
                        MsgBox("Unable To delete CT2 Invoice Knock Off Details.", MsgBoxStyle.Information, ResolveResString(100))
                        objCmd = Nothing
                        mP_Connection.RollbackTrans()
                        Exit Sub
                    End If
                    objCmd = Nothing
                    '10736222


                    mP_Connection.Execute(strupSaleDtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    mP_Connection.Execute(strupSalechallan, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    If Len(mstrUpdDispatchSql) > 0 Then
                        mP_Connection.Execute(mstrUpdDispatchSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End If

                    '101631219
                    objCmd = New ADODB.Command

                    With objCmd
                        .ActiveConnection = mP_Connection
                        .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                        .CommandText = "USP_BSR_INVOICE_UPDATION"
                        .CommandTimeout = 0
                        .Parameters.Append(.CreateParameter("@UNIT_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUnitId))
                        .Parameters.Append(.CreateParameter("@TEMP_INV_NO", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, , txtChallanNo.Text.Trim))
                        .Parameters.Append(.CreateParameter("@INV_NO", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, , 0))
                        .Parameters.Append(.CreateParameter("@USER_ID", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 50, mP_User))
                        .Parameters.Append(.CreateParameter("@OPERATION_TYPE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 50, "DELETE"))
                        .Parameters.Append(.CreateParameter("@MESSAGE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInputOutput, 8000, ""))
                        .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End With

                    If objCmd.Parameters(objCmd.Parameters.Count - 1).Value.ToString().Trim.Length <> 0 Then
                        MsgBox(objCmd.Parameters(objCmd.Parameters.Count - 1).Value.ToString().Trim, MsgBoxStyle.Information, ResolveResString(100))
                        objCmd = Nothing
                        mP_Connection.RollbackTrans()
                        Exit Sub
                    End If
                    objCmd = Nothing
                    '101631219

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
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                If CBool(Find_Value("select TextPrinting from sales_parameter WHERE UNIT_CODE = '" & gstrUnitId & "'")) Then
                    Call PrintingInvoice()
                Else
                    Call PrintingInvoiceRPT()
                    frmReportViewer.ReportHeader = Me.ctlFormHeader1.HeaderString()
                    On Error Resume Next
                    frmReportViewer.Show()
                End If

                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default) 'Change the Mouse Pointer of the Screen
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub SpChEntry_Change(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ChangeEvent) Handles SpChEntry.Change
        Dim intRowCount As Short
        Dim intmaxrows As Short
        Dim varFromBox As Object
        Dim varItem As Object
        Dim VarToBox As Object
        Dim varQty As Object
        Dim boxqty As Object
        Dim varCumulativeBoxes As Object
        On Error GoTo ErrHandler
        With SpChEntry
            If e.col = 5 Or e.col = 22 Then
                With SpChEntry
                    intmaxrows = SpChEntry.MaxRows
                    For intRowCount = 1 To intmaxrows
                        varItem = Nothing
                        Call .GetText(1, intRowCount, varItem)
                        varQty = Nothing
                        Call .GetText(5, intRowCount, varQty)
                        boxqty = Nothing
                        Call .GetText(22, intRowCount, boxqty)
                        If Val(boxqty) > 0 Then
                            If Val(varQty) > 0 Then
                                If (varQty / boxqty) - Int(varQty / boxqty) > 0 Then
                                    If intRowCount = 1 Then
                                        Call .SetText(11, intRowCount, 1)
                                        Call .SetText(12, intRowCount, Int(varQty / boxqty) + 1)
                                        Call .SetText(13, intRowCount, Int(varQty / boxqty) + 1)
                                    Else
                                        VarToBox = Nothing
                                        Call .GetText(12, intRowCount - 1, VarToBox)
                                        varCumulativeBoxes = Nothing
                                        Call .GetText(13, intRowCount - 1, varCumulativeBoxes)
                                        Call .SetText(11, intRowCount, Val(VarToBox) + 1)
                                        Call .SetText(12, intRowCount, Val(VarToBox) + Int(varQty / boxqty) + 1)
                                        Call .SetText(13, intRowCount, Val(varCumulativeBoxes) + Int(varQty / boxqty) + 1)
                                    End If
                                Else
                                    If intRowCount = 1 Then
                                        Call .SetText(11, intRowCount, 1)
                                        Call .SetText(12, intRowCount, (Int(varQty / boxqty)))
                                        Call .SetText(13, intRowCount, Int(varQty / boxqty))
                                    Else
                                        VarToBox = Nothing
                                        Call .GetText(12, intRowCount - 1, VarToBox)
                                        varCumulativeBoxes = Nothing
                                        Call .GetText(13, intRowCount - 1, varCumulativeBoxes)
                                        Call .SetText(11, intRowCount, Val(VarToBox) + 1)
                                        Call .SetText(12, intRowCount, Val(VarToBox) + Int(varQty / boxqty))
                                        Call .SetText(13, intRowCount, Val(varCumulativeBoxes) + Int(varQty / boxqty))
                                    End If
                                End If
                            End If
                        End If
                    Next
                End With
            End If
            If (e.col = 11) Or (e.col = 12) Then
                intmaxrows = SpChEntry.MaxRows
                For intRowCount = 1 To intmaxrows
                    varFromBox = Nothing
                    Call .GetText(11, intRowCount, varFromBox)
                    VarToBox = Nothing
                    Call .GetText(12, intRowCount, VarToBox)
                    If intRowCount = 1 Then
                        If Len(Trim(varFromBox)) Then
                            If Len(Trim(VarToBox)) Then
                                Call .SetText(13, intRowCount, (Val(VarToBox) - Val(varFromBox)) + 1)
                            End If
                        End If
                    Else
                        varCumulativeBoxes = Nothing
                        Call .GetText(13, intRowCount - 1, varCumulativeBoxes)
                        If Len(Trim(varCumulativeBoxes)) Then
                            If Len(Trim(varFromBox)) Then
                                If Len(Trim(VarToBox)) Then
                                    Call .SetText(13, intRowCount, varCumulativeBoxes + ((Val(VarToBox) - Val(varFromBox)) + 1))
                                End If
                            End If
                        End If
                    End If
                Next
            End If
            If e.col = 3 Then
                With SpChEntry
                    Call .SetText(21, e.row, "C")
                End With
            End If
        End With
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub SpChEntry_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles SpChEntry.GotFocus
        If ctlPerValue.Enabled = True Then
            ctlPerValue.Enabled = False
            ctlPerValue.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        End If
    End Sub
    Private Sub SpChEntry_KeyDownEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyDownEvent) Handles SpChEntry.KeyDownEvent
        Dim strHelp As String
        Dim strdate As String
        Dim strCondition As String
        Dim strCreate As String
        Dim strItemCode As String
        Dim strStockLocation As String
        Dim strSelectSql As String
        Dim validMon As String
        Dim Validyrmon As String
        Dim effectMon As String
        Dim effectyrmon As String
        Dim rsCustOrdHdr As ClsResultSetDB
        Dim rsStockLocation As ClsResultSetDB
        Dim strInvDes As String
        Dim strInvSubTypeDes As String
        On Error GoTo ErrHandler
        If CmbInvType.Enabled Then
            strInvDes = CmbInvType.Text
            strInvSubTypeDes = CmbInvSubType.Text
        Else
            rsStockLocation = New ClsResultSetDB
            rsStockLocation.GetResult("Select Description,Sub_type_Description from SaleConf a,SalesChallan_dtl b where a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND b.Invoice_type = a.Invoice_type and b.Sub_Category = a.Sub_type and a.Location_Code =b.Location_code and b.Location_code ='" & txtLocationCode.Text & "' and b.Doc_No = " & txtChallanNo.Text & " And (fin_start_date <= getDate() And fin_end_date >= getDate())")
            strInvDes = rsStockLocation.GetValue("Description")
            strInvSubTypeDes = rsStockLocation.GetValue("Sub_type_Description")
            rsStockLocation.ResultSetClose()
        End If
        Dim strItemDrg() As String
        If CmdGrpChEnt.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            If e.keyCode = System.Windows.Forms.Keys.F1 And (SpChEntry.ActiveCol = 1 Or SpChEntry.ActiveCol = 2) Then
                '************To make Select SQL on Invoice Type Basis
                strdate = getDateForDB(VB6.Format(GetServerDate(), gstrDateFormat))
                strSelectSql = "Select effectMon=convert(char(2),month(effect_date)),effectYr=convert(char(4),Year(effect_date)),"
                strSelectSql = strSelectSql & " validMon=convert(char(2),month(Valid_date)),validYr=convert(char(4),year(Valid_date))"
                strSelectSql = strSelectSql & " from Cust_Ord_hdr where "
                strSelectSql = strSelectSql & " Account_Code='" & Trim(txtCustCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "' and Cust_Ref='" & Trim(txtRefNo.Text) & "'"
                strSelectSql = strSelectSql & " and Amendment_No='" & Trim(txtAmendNo.Text) & "' and Active_Flag = 'A'"
                rsCustOrdHdr = New ClsResultSetDB
                rsCustOrdHdr.GetResult(strSelectSql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                If rsCustOrdHdr.GetNoRows > 0 Then
                    validMon = CStr(Month(GetServerDate))
                    If CDbl(validMon) < 10 Then
                        validMon = "0" & validMon
                    End If
                    Validyrmon = Year(GetServerDate) & validMon
                    effectMon = rsCustOrdHdr.GetValue("EffectMon")
                    If CDbl(effectMon) < 10 Then
                        effectMon = "0" & effectMon
                    End If
                    effectyrmon = rsCustOrdHdr.GetValue("effectYr") & effectMon
                End If
                rsCustOrdHdr.ResultSetClose()
                If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                    strStockLocation = StockLocationSalesConf((CmbInvType.Text), (CmbInvSubType.Text), "DESCRIPTION")
                ElseIf CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                    rsStockLocation = New ClsResultSetDB
                    rsStockLocation.GetResult("Select Stock_Location from SaleConf a,SalesChallan_dtl b where a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND b.Invoice_type = a.Invoice_type and b.Sub_Category = a.Sub_type and a.Location_Code =b.Location_code and b.Location_code ='" & txtLocationCode.Text & "' and b.Doc_No = " & txtChallanNo.Text & " And (fin_start_date <= getDate() And fin_end_date >= getDate())")
                    strStockLocation = rsStockLocation.GetValue("Stock_Location")
                    rsStockLocation.ResultSetClose()
                End If
                If Len(Trim(strStockLocation)) <= 0 Then
                    MsgBox("Define Stock Location in SalesConf first.", MsgBoxStyle.Information, "eMPro")
                    With SpChEntry
                        .Col = .ActiveCol : .Row = .ActiveRow : SpChEntry.Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                    End With
                    Exit Sub
                End If
                Select Case UCase(strInvDes)
                    Case "NORMAL INVOICE", "EXPORT INVOICE"
                        Select Case UCase(strInvSubTypeDes)
                            Case "FINISHED GOODS"
                                strSelectSql = makeSelectSql((txtCustCode.Text), (txtRefNo.Text), (txtAmendNo.Text), effectyrmon, Validyrmon, strStockLocation, strdate, "'F','S'")
                            Case "COMPONENTS"
                                strSelectSql = MakeSelectSubQuery((txtCustCode.Text), (txtRefNo.Text), (txtAmendNo.Text), strStockLocation, "'C'")
                            Case "COMPONENTS"
                                strSelectSql = MakeSelectSubQuery((txtCustCode.Text), (txtRefNo.Text), (txtAmendNo.Text), strStockLocation, "'C'")
                            Case "RAW MATERIAL"
                                strSelectSql = MakeSelectSubQuery((txtCustCode.Text), (txtRefNo.Text), (txtAmendNo.Text), strStockLocation, "'R','S','B','M'")
                            Case "ASSETS"
                                strSelectSql = MakeSelectSubQuery((txtCustCode.Text), (txtRefNo.Text), (txtAmendNo.Text), strStockLocation, "'P'")
                            Case "TRADING GOODS"
                                strSelectSql = makeSelectSql((txtCustCode.Text), (txtRefNo.Text), (txtAmendNo.Text), effectyrmon, Validyrmon, strStockLocation, strdate, "'T','S'")
                            Case "TOOLS & DIES"
                                strSelectSql = MakeSelectSubQuery((txtCustCode.Text), (txtRefNo.Text), (txtAmendNo.Text), strStockLocation, "'P','A'")
                            Case "EXPORTS"
                                strSelectSql = makeSelectSql((txtCustCode.Text), (txtRefNo.Text), (txtAmendNo.Text), effectyrmon, Validyrmon, strStockLocation, strdate, "'F','S'")
                            Case "SCRAP"
                                strSelectSql = "SELECT Distinct(a.Item_Code),a.Item_Code,a.description, Tariff_code,a.Unit_code  FROM Item_Mst a,Itembal_Mst b"
                                strSelectSql = strSelectSql & " where a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND  a.Item_Code=b.Item_Code "
                                strSelectSql = strSelectSql & " and cur_bal >0 and a.Status ='A' and a.Hold_Flag <> 1"
                                strSelectSql = strSelectSql & " and b.Location_Code = '" & strStockLocation & "'"
                        End Select
                    Case "JOBWORK INVOICE"
                        strSelectSql = makeSelectSql((txtCustCode.Text), (txtRefNo.Text), (txtAmendNo.Text), effectyrmon, Validyrmon, strStockLocation, strdate, "'F'")
                    Case "TRANSFER INVOICE"
                        Select Case UCase(strInvSubTypeDes)
                            Case "FINISHED GOODS"
                                'strSelectSql = "SELECT Distinct a.Item_Code,c.Cust_drgNo,c.Drg_Desc, ISNULL(a.Tariff_code,0) AS Tariff_code,a.UNIT_CODE FROM Item_Mst a,Itembal_Mst b,CustItem_Mst c "
                                'strSelectSql = strSelectSql & " where a.UNIT_CODE = b.UNIT_CODE AND a.UNIT_CODE = c.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND  a.Item_Code=b.Item_Code and a.Item_Main_Grp IN( 'F','S') and a.Item_Code = c.ITem_Code"
                                'strSelectSql = strSelectSql & " and cur_bal >0 and a.Status ='A' and a.Hold_Flag <> 1 and c.Account_code ='" & txtCustCode.Text & "'"
                                'strSelectSql = strSelectSql & " and b.Location_Code = '" & strStockLocation & "'"
                                strSelectSql = makeSelectSql((txtCustCode.Text), (txtRefNo.Text), (txtAmendNo.Text), effectyrmon, Validyrmon, strStockLocation, strdate, "'F','S'")
                            Case "ASSETS"
                                strSelectSql = MakeSelectStatementForITemMst("'P'", strStockLocation)
                            Case "INPUTS"
                                strSelectSql = MakeSelectStatementForITemMst("'R','C','M','N','S','B','A'", strStockLocation)
                        End Select
                    Case "SAMPLE INVOICE"
                        Select Case UCase(strInvSubTypeDes)
                            Case "FINISHED GOODS"
                                strSelectSql = MakeSelectStatementForITemMst("'F','S'", strStockLocation)
                            Case "RAW MATERIAL"
                                strSelectSql = MakeSelectStatementForITemMst("'R'", strStockLocation)
                            Case "COMPONENTS"
                                strSelectSql = MakeSelectStatementForITemMst("'C'", strStockLocation)
                        End Select
                    Case "CSM INVOICE"
                        Select Case UCase(strInvSubTypeDes)
                            Case "CSM INVOICE"
                                strSelectSql = MakeSelectStatementForITemMst("'F'", strStockLocation)
                        End Select
                    Case "REJECTION"
                        If Len(Trim(txtRefNo.Text)) = 0 Then
                            strSelectSql = "SELECT Distinct(a.Item_Code),a.Item_Code,a.description,ISNULL(c.Tariff_code,0),a.UNIT_CODE FROM vend_item a ,Itembal_Mst b,Item_Mst c"
                            strSelectSql = strSelectSql & " where a.UNIT_CODE = b.UNIT_CODE AND a.UNIT_CODE = c.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND a.Item_Code=b.Item_Code and a.Item_code = c.Item_code and a.Account_code ='" & txtCustCode.Text & "' "
                            strSelectSql = strSelectSql & " and cur_bal >0 "
                            strSelectSql = strSelectSql & " and b.Location_Code = '" & strStockLocation & "'"
                        Else
                            strSelectSql = AddDataFromGrinDtl((txtCustCode.Text), CDbl(txtRefNo.Text), strStockLocation)
                        End If
                End Select
                mP_Connection.Execute("delete FROM invoiceItemHelp WHERE UNIT_CODE = '" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                System.Windows.Forms.Application.DoEvents()
                mP_Connection.Execute("insert into invoiceItemHelp (Item_code,Cust_DrgNo,Cust_drgDesc,Tariff_Code,UNIT_CODE) " & strSelectSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                System.Windows.Forms.Application.DoEvents()
                strItemDrg = Me.ctlEMPHelpInvoiceEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "SElect Cust_drgNo, Item_code, Cust_DrgDesc, isnull(Tariff_code,0) as Tariff_code from invoiceItemHelp where unit_code = '" & gstrUnitId & "' ", "Item Details")
                On Error Resume Next
                If strItemDrg Is Nothing Then Exit Sub
                If UBound(strItemDrg) = 0 Then
                    Exit Sub
                End If
                If Err.Number = 9 Then 'if subscript out of range error then exit sub
                    Exit Sub
                End If
                On Error GoTo ErrHandler
                If strItemDrg(0) = "0" Then
                    'MsgBox("No Item Available To Display", MsgBoxStyle.Information, "eMPro") : SpChEntry.Focus() : Exit Sub
                    MsgBox("No Items for selected Invoice in Sales Order.Please Check Following :" & vbCrLf & "1. Item in SO are Active and Not on Hold." & vbCrLf & "2. Check Balance of Items for location " & strStockLocation & "." & vbCrLf & "3. Check Marketing Schedule in Case of Finished\Trading Goods in SO.", MsgBoxStyle.Information, "empower") : SpChEntry.Focus() : Exit Sub
                Else
                    With SpChEntry
                        If checkforDuplicateItemCodeandDrgNo(strItemDrg(1), strItemDrg(0), .ActiveRow) = False Then
                            .Col = 1 : .Row = .ActiveRow
                            .Text = strItemDrg(1)
                            .Col = 2 : .Row = .ActiveRow
                            .Text = strItemDrg(0)
                            .Col = 20 : .Row = .ActiveRow
                            .Text = strItemDrg(3)
                        Else
                            .Col = .ActiveCol : .Row = .ActiveRow : .EditModePermanent = True : .EditModeReplace = True : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus() : Exit Sub
                        End If
                        If CheckForTariffCode(CStr(strItemDrg(3))) = True Then
                            Select Case UCase(strInvDes)
                                Case "NORMAL INVOICE", "EXPORT INVOICE", "JOBWORK INVOICE", "TRANSFER INVOICE"
                                    If UCase(strInvSubTypeDes) <> "SCRAP" Then
                                        Call DisplaydetailsfromCustOrdDtl(.ActiveRow, strStockLocation, strItemDrg(1), strItemDrg(0))
                                        .Col = 5 : .Row = .ActiveRow : .EditModePermanent = True : .EditModeReplace = True : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus() : Exit Sub
                                    Else
                                        Call DisplayDetailsfromItemMst(.ActiveRow, strItemDrg(1), strStockLocation)
                                        .Col = 5 : .EditModePermanent = True : .EditModeReplace = True : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus() : Exit Sub
                                    End If
                                Case "SAMPLE INVOICE", "CSM INVOICE"
                                    Select Case UCase(strInvSubTypeDes)
                                        Case "FINISHED GOODS"
                                            strStockLocation = StockLocationSalesConf((CmbInvType.Text), (CmbInvSubType.Text), "DESCRIPTION")
                                            'Call DisplayDetailsfromItemMst(.ActiveRow, strItemDrg(1), strStockLocation, "'F'")
                                            Call DisplayDetailsfromItemMst(.ActiveRow, strItemDrg(1), strStockLocation, "'F','S'")
                                            .Col = 5 : .EditModePermanent = True : .EditModeReplace = True : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus() : Exit Sub
                                        Case "ASSETS"
                                            strStockLocation = StockLocationSalesConf((CmbInvType.Text), (CmbInvSubType.Text), "DESCRIPTION")
                                            Call DisplayDetailsfromItemMst(.ActiveRow, strItemDrg(1), strStockLocation, "'P'")
                                            .Col = 5 : .EditModePermanent = True : .EditModeReplace = True : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus() : Exit Sub
                                        Case "INPUTS"
                                            strStockLocation = StockLocationSalesConf((CmbInvType.Text), (CmbInvSubType.Text), "DESCRIPTION")
                                            Call DisplayDetailsfromItemMst(.ActiveRow, strItemDrg(1), strStockLocation, "'R','C','M','N','S','B','A'")
                                            .Col = 5 : .EditModePermanent = True : .EditModeReplace = True : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus() : Exit Sub
                                        Case "RAW MATERIAL"
                                            strStockLocation = StockLocationSalesConf((CmbInvType.Text), (CmbInvSubType.Text), "DESCRIPTION")
                                            Call DisplayDetailsfromItemMst(.ActiveRow, strItemDrg(1), strStockLocation, "'R'")
                                            .Col = 5 : .EditModePermanent = True : .EditModeReplace = True : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus() : Exit Sub
                                        Case "COMPONENTS"
                                            strStockLocation = StockLocationSalesConf((CmbInvType.Text), (CmbInvSubType.Text), "DESCRIPTION")
                                            Call DisplayDetailsfromItemMst(.ActiveRow, strItemDrg(1), strStockLocation, "'C'")
                                            .Col = 5 : .EditModePermanent = True : .EditModeReplace = True : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus() : Exit Sub
                                        Case "CSM INVOICE"
                                            strStockLocation = StockLocationSalesConf((CmbInvType.Text), (CmbInvSubType.Text), "DESCRIPTION")
                                            Call DisplayDetailsfromItemMst(.ActiveRow, strItemDrg(1), strStockLocation, "'F'")
                                            .Col = 5 : .EditModePermanent = True : .EditModeReplace = True : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus() : Exit Sub
                                    End Select
                                Case "REJECTION"
                                    strStockLocation = StockLocationSalesConf((CmbInvType.Text), (CmbInvSubType.Text), "DESCRIPTION")
                                    Call DisplayDetailsfromItemMst(.ActiveRow, strItemDrg(1), strStockLocation)
                                    .Col = 5 : .EditModePermanent = True : .EditModeReplace = True : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus() : Exit Sub
                            End Select
                        Else
                            Exit Sub
                        End If
                    End With
                End If
            ElseIf e.keyCode = System.Windows.Forms.Keys.F1 And SpChEntry.ActiveCol = 8 Then
                '101188073
                If gblnGSTUnit Then Exit Sub
                '101188073
                With SpChEntry
                    .Row = .ActiveRow
                    .Col = .ActiveCol
                    If Len(Trim(.Text)) = 0 Then 'To check if There is No Text Then Show All Help
                        strHelp = ShowList(1, 6, "", "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND Tx_TaxeID='CVD'")
                        If strHelp = "-1" Then 'If No Record Exists In The Table
                            Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                            Exit Sub
                        Else
                            .Text = strHelp
                        End If
                    Else
                        'To Display All Possible Help Starting With Text in TextField
                        strHelp = ShowList(1, 6, Trim(.Text), "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND Tx_TaxeID='CVD'")
                        If strHelp = "-1" Then 'If No Record Exists In The Table
                            Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                            Exit Sub
                        Else
                            .Text = strHelp
                        End If
                    End If
                End With
            ElseIf e.keyCode = System.Windows.Forms.Keys.F1 And SpChEntry.ActiveCol = 9 Then
                '101188073
                If gblnGSTUnit Then Exit Sub
                '101188073
                With SpChEntry
                    .Row = .ActiveRow
                    .Col = .ActiveCol
                    If Len(Trim(.Text)) = 0 Then 'To check if There is No Text Then Show All Help
                        strHelp = ShowList(1, 6, "", "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND Tx_TaxeID='SAD'")
                        If strHelp = "-1" Then 'If No Record Exists In The Table
                            Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                            Exit Sub
                        Else
                            .Text = strHelp
                        End If
                    Else
                        'To Display All Possible Help Starting With Text in TextField
                        strHelp = ShowList(1, 6, Trim(.Text), "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND Tx_TaxeID='SAD'")
                        If strHelp = "-1" Then 'If No Record Exists In The Table
                            Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                            Exit Sub
                        Else
                            .Text = strHelp
                        End If
                    End If
                End With
            ElseIf e.keyCode = System.Windows.Forms.Keys.F1 And SpChEntry.ActiveCol = 7 Then
                '101188073
                If gblnGSTUnit Then Exit Sub
                '101188073
                With SpChEntry
                    .Row = .ActiveRow
                    .Col = 1
                    strItemCode = Trim(.Text)
                    .Col = .ActiveCol
                    strCondition = "AND Tx_TaxeID='EXC' " & PrepareQueryForShowingExcise(False, strItemCode)
                    If Len(Trim(.Text)) = 0 Then 'To check if There is No Text Then Show All Help
                        strHelp = ShowList(1, 6, "", "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", strCondition)
                        If strHelp = "-1" Then 'If No Record Exists In The Table
                            Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                            Exit Sub
                        Else
                            .Text = strHelp
                        End If
                    Else
                        'To Display All Possible Help Starting With Text in TextField
                        strHelp = ShowList(1, 6, "", "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", strCondition)
                        If strHelp = "-1" Then 'If No Record Exists In The Table
                            Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                            Exit Sub
                        Else
                            .Text = strHelp
                        End If
                    End If
                End With
            ElseIf SpChEntry.ActiveCol = 22 Then
                If e.keyCode = System.Windows.Forms.Keys.Return Then
                    If SpChEntry.ActiveRow = SpChEntry.MaxRows Then
                        Call addRowAtEnterKeyPress(1)
                        Call ChangeCellTypeStaticText()
                        SpChEntry.Col = 1 : SpChEntry.Row = SpChEntry.MaxRows
                        SpChEntry.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                    End If
                End If
            End If
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub SpChEntry_KeyPressEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyPressEvent) Handles SpChEntry.KeyPressEvent
        On Error GoTo ErrHandler
        Select Case e.keyAscii
            Case 39, 34, 96
                e.keyAscii = 0
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub SpChEntry_KeyUpEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyUpEvent) Handles SpChEntry.KeyUpEvent
        Dim intRow As Short
        Dim intDelete As Short
        Dim intLoopCount As Short
        Dim intMaxLoop As Short
        Dim VarDelete As Object
        Dim strInvoiceType As String
        Dim rsChallanEntry As ClsResultSetDB
        On Error GoTo ErrHandler
        If ((e.shift = 2) And (e.keyCode = System.Windows.Forms.Keys.D)) Then
            If CmdGrpChEnt.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                With SpChEntry
                    If .MaxRows > 1 Then
                        intRow = .ActiveRow : intMaxLoop = SpChEntry.MaxRows
                        For intLoopCount = 1 To intMaxLoop
                            If intLoopCount <> intRow Then
                                VarDelete = Nothing
                                Call .GetText(14, intLoopCount, VarDelete)
                                If UCase(VarDelete) = "D" Then
                                    intDelete = intDelete + 1
                                End If
                            End If
                        Next
                        If (intMaxLoop - intDelete) > 1 Then
                            Call .SetText(14, intRow, "D")
                            .Row = .ActiveRow : .Row2 = .ActiveRow : .BlockMode = True : .RowHidden = True : .MaxRows = .MaxRows - 1 : .BlockMode = False
                        End If
                    End If
                End With
            End If
        ElseIf e.keyCode = System.Windows.Forms.Keys.Return Then
            With SpChEntry
                If CmdGrpChEnt.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                    rsChallanEntry = New ClsResultSetDB
                    rsChallanEntry.GetResult("Select a.Description,a.Sub_Type_Description from SaleConf a,SalesChallan_Dtl b where a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND Doc_No = " & txtChallanNo.Text & " and a.Invoice_Type = b.Invoice_type and a.Sub_type = b.Sub_Category and a.Location_code = b.Location_code and (fin_start_date <= getdate() and fin_end_date >= getdate())")
                    strInvoiceType = UCase(rsChallanEntry.GetValue("Description"))
                    rsChallanEntry.ResultSetClose()
                Else
                    strInvoiceType = UCase(CmbInvType.Text)
                End If
                If UCase(strInvoiceType) = "SAMPLE INVOICE" Or UCase(strInvoiceType) = "CSM INVOICE" Then
                    If .ActiveCol = 15 Then
                        If .ActiveRow = .MaxRows Then
                            Call addRowAtEnterKeyPress(1)
                            Call ChangeCellTypeStaticText()
                            .Col = 1 : .Row = .MaxRows : .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                        Else
                            intMaxLoop = 1
                            For intLoopCount = .MaxRows To intMaxLoop Step -1
                                VarDelete = Nothing
                                Call .GetText(14, intLoopCount, VarDelete)
                                If VarDelete <> "D" Then
                                    If .ActiveRow = intLoopCount Then
                                        Call addRowAtEnterKeyPress(1)
                                        Call ChangeCellTypeStaticText()
                                        .Col = 1 : .Row = .MaxRows : .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                    Else
                                        Exit For
                                    End If
                                End If
                            Next
                        End If
                    End If
                Else
                End If
            End With
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub SpChEntry_LeaveCell(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_LeaveCellEvent) Handles SpChEntry.LeaveCell
        On Error GoTo ErrHandler
        Dim lstrReturnVal As String
        Dim strInvTypeDes As String
        Dim strInvSubTypeDes As String
        Dim strWhereClause As String
        Dim strStockLocation As String
        Dim strSQL As String
        Dim varItemCode As Object
        Dim varDrgNo As Object
        Dim varItemCodeDummy As Object
        Dim varDrgNoDummy As Object
        Dim VarDelete As Object
        Dim varChanged As Object
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        Dim rsStockLocation As ClsResultSetDB
        Dim rsGrnDtl As ClsResultSetDB
        Dim rsItemMst As ClsResultSetDB
        If e.newCol = -1 Then Exit Sub
        lstrReturnVal = ""
        If CmdGrpChEnt.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            With SpChEntry
                If (e.col = 1) Or (e.col = 2) Then
                    varItemCode = Nothing
                    Call .GetText(1, e.row, varItemCode)
                    varDrgNo = Nothing
                    Call .GetText(2, e.row, varDrgNo)
                    If e.col = 1 Then
                        If Len(Trim(varItemCode)) = 0 Then Exit Sub
                    Else
                        intMaxLoop = SpChEntry.MaxRows
                        If (Len(Trim(varItemCode)) > 0) And (Len(Trim(varDrgNo)) > 0) Then
                            For intLoopCounter = 1 To intMaxLoop
                                varItemCodeDummy = Nothing
                                Call .GetText(1, intLoopCounter, varItemCodeDummy)
                                varDrgNoDummy = Nothing
                                Call .GetText(2, intLoopCounter, varDrgNoDummy)
                                VarDelete = Nothing
                                Call .GetText(14, intLoopCounter, VarDelete)
                                If UCase(Trim(VarDelete)) <> "D" Then
                                    If UCase(Trim(varItemCode)) = UCase(Trim(varItemCodeDummy)) Then
                                        If UCase(Trim(varDrgNo)) = UCase(Trim(varDrgNoDummy)) Then
                                            If intLoopCounter <> e.row Then
                                                MsgBox("Duplicate Item Code and Customer Part Code.", MsgBoxStyle.Information, "eMPro")
                                                Call .SetText(1, e.row, "")
                                                Call .SetText(2, e.row, "")
                                                Call .SetText(20, e.row, "")
                                                .Col = e.col : .Row = e.row : .Focus() : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : Exit Sub
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    End If
                    If e.col = 2 Then
                        If Len(Trim(varDrgNo)) = 0 Then Exit Sub
                    Else
                        intMaxLoop = SpChEntry.MaxRows
                        If (Len(Trim(varItemCode)) > 0) And (Len(Trim(varDrgNo)) > 0) Then
                            For intLoopCounter = 1 To intMaxLoop
                                varItemCodeDummy = Nothing
                                Call .GetText(1, intLoopCounter, varItemCodeDummy)
                                varDrgNoDummy = Nothing
                                Call .GetText(2, intLoopCounter, varDrgNoDummy)
                                VarDelete = Nothing
                                Call .GetText(14, intLoopCounter, VarDelete)
                                If UCase(Trim(VarDelete)) <> "D" Then
                                    If UCase(Trim(varItemCode)) = UCase(Trim(varItemCodeDummy)) Then
                                        If UCase(Trim(varDrgNo)) = UCase(Trim(varDrgNoDummy)) Then
                                            If intLoopCounter <> e.row Then
                                                MsgBox("Duplicate Item Code and Customer Part Code.", MsgBoxStyle.Information, "eMPro")
                                                Call .SetText(1, e.row, "")
                                                Call .SetText(2, e.row, "")
                                                Call .SetText(20, e.row, "")
                                                .Col = e.col : .Row = e.row : .Focus() : .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                                Exit Sub
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    End If
                    If CmbInvType.Enabled Then
                        strInvTypeDes = CmbInvType.Text
                        strInvSubTypeDes = CmbInvSubType.Text
                    Else
                        rsStockLocation = New ClsResultSetDB
                        rsStockLocation.GetResult("Select Description,Sub_type_Description from SaleConf a,SalesChallan_dtl b where a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND b.Invoice_type = a.Invoice_type and b.Sub_Category = a.Sub_type and a.Location_Code =b.Location_code and b.Location_code ='" & txtLocationCode.Text & "' and b.Doc_No = " & txtChallanNo.Text & " and  (fin_start_date <= getdate() and fin_end_date >= getdate())")
                        strInvTypeDes = rsStockLocation.GetValue("Description")
                        strInvSubTypeDes = rsStockLocation.GetValue("Sub_type_Description")
                        rsStockLocation.ResultSetClose()
                    End If
                    Select Case UCase(strInvTypeDes)
                        Case "NORMAL INVOICE", "EXPORT INVOICE", "JOBWORK INVOICE", "TRANSFER INVOICE"
                            If strInvSubTypeDes <> "SCRAP" Then
                                varChanged = Nothing
                                Call .GetText(21, e.row, varChanged)
                                'Value of rate changed then dont refresh
                                If Trim(UCase(varChanged)) <> "C" Then
                                    strStockLocation = StockLocationSalesConf(strInvTypeDes, strInvSubTypeDes, "DESCRIPTION")
                                    If DisplaydetailsfromCustOrdDtl(CShort(e.row), strStockLocation, CStr(varItemCode), CStr(varDrgNo)) = False Then
                                        If e.col = 1 Then
                                            MsgBox("Invalid Item Code.", MsgBoxStyle.Information, "eMPro")
                                        Else
                                            MsgBox("Invalid Customer Part Code.", MsgBoxStyle.Information, "eMPro")
                                        End If
                                        .Col = e.col : .Row = e.row : .Focus() : .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                    Else
                                        If Len(Trim(varItemCode)) = 0 Then
                                            varItemCode = Nothing
                                            Call .GetText(1, e.row, varItemCode)
                                        End If
                                        rsItemMst = New ClsResultSetDB
                                        rsItemMst.GetResult("Select isnull(Tariff_Code,0) as Tariff_Code from Item_Mst where Item_code = '" & varItemCode & "' and UNIT_CODE = '" & gstrUNITID & "'")
                                        Call .SetText(20, e.row, rsItemMst.GetValue("Tariff_Code"))
                                        If CheckForTariffCode(CStr(CDbl(rsItemMst.GetValue("Tariff_Code")))) = False Then
                                            ClearGridRow(e.row)
                                        End If
                                        rsItemMst.ResultSetClose()
                                    End If
                                End If
                            Else
                                varChanged = Nothing
                                Call .GetText(21, e.row, varChanged)
                                'Value of rate changed then dont refresh
                                If Trim(UCase(varChanged)) <> "C" Then
                                    strStockLocation = StockLocationSalesConf((CmbInvType.Text), (CmbInvSubType.Text), "DESCRIPTION")
                                    If DisplayDetailsfromItemMst(CShort(e.row), CStr(varItemCode), strStockLocation) = False Then
                                        If e.col = 1 Then
                                            MsgBox("Invalid Item Code.", MsgBoxStyle.Information, "eMPro")
                                        Else
                                            MsgBox("Invalid Customer Part Code.", MsgBoxStyle.Information, "eMPro")
                                        End If
                                        .Col = e.col : .Row = e.row : .Focus() : .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                    End If
                                End If
                            End If
                        Case "SAMPLE INVOICE", "REJECTION", "CSM INVOICE"
                            If strInvTypeDes <> "REJECTION" Then
                                varChanged = Nothing
                                Call .GetText(21, e.row, varChanged)
                                'Value of rate changed then dont refresh
                                If Trim(UCase(varChanged)) <> "C" Then
                                    Select Case strInvSubTypeDes
                                        Case "FINISHED GOODS", "CSM INVOICE"
                                            strStockLocation = StockLocationSalesConf((CmbInvType.Text), (CmbInvSubType.Text), "DESCRIPTION")
                                            If DisplayDetailsfromItemMst(CShort(e.row), CStr(varItemCode), strStockLocation, "'F'") = False Then
                                                If e.col = 1 Then
                                                    MsgBox("Invalid Item Code.", MsgBoxStyle.Information, "eMPro")
                                                Else
                                                    MsgBox("Invalid Customer Part Code.", MsgBoxStyle.Information, "eMPro")
                                                End If
                                                .Col = e.col : .Row = e.row : .Focus() : .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                            Else
                                                rsItemMst = New ClsResultSetDB
                                                rsItemMst.GetResult("Select Tariff_Code from Item_Mst where Item_code = '" & varItemCode & "' and UNIT_CODE = '" & gstrUNITID & "'")
                                                Call .SetText(20, e.row, rsItemMst.GetValue("Tariff_Code"))
                                                If CheckForTariffCode(CStr(CDbl(rsItemMst.GetValue("Tariff_Code")))) = False Then
                                                    ClearGridRow(e.row)
                                                End If
                                                rsItemMst.ResultSetClose()
                                            End If
                                        Case "ASSETS"
                                            strStockLocation = StockLocationSalesConf((CmbInvType.Text), (CmbInvSubType.Text), "DESCRIPTION")
                                            If DisplayDetailsfromItemMst(CShort(e.row), CStr(varItemCode), strStockLocation, "'P'") = False Then
                                                If e.col = 1 Then
                                                    MsgBox("Invalid Item Code.", MsgBoxStyle.Information, "eMPro")
                                                Else
                                                    MsgBox("Invalid Customer Part Code.", MsgBoxStyle.Information, "eMPro")
                                                End If
                                                .Col = e.col : .Row = e.row : .Focus() : .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                            Else
                                                rsItemMst = New ClsResultSetDB
                                                rsItemMst.GetResult("Select Tariff_Code from Item_Mst where Item_code = '" & varItemCode & "' and UNIT_CODE = '" & gstrUNITID & "'")
                                                Call .SetText(20, e.row, rsItemMst.GetValue("Tariff_Code"))
                                                If CheckForTariffCode(CStr(CDbl(rsItemMst.GetValue("Tariff_Code")))) = False Then
                                                    ClearGridRow(e.row)
                                                End If
                                                rsItemMst.ResultSetClose()
                                            End If
                                        Case "INPUTS"
                                            strStockLocation = StockLocationSalesConf((CmbInvType.Text), (CmbInvSubType.Text), "DESCRIPTION")
                                            If DisplayDetailsfromItemMst(CShort(e.row), CStr(varItemCode), strStockLocation, "'R','C','M','N','S','B','A'") = False Then
                                                If e.col = 1 Then
                                                    MsgBox("Invalid Item Code.", MsgBoxStyle.Information, "eMPro")
                                                Else
                                                    MsgBox("Invalid Customer Part Code.", MsgBoxStyle.Information, "eMPro")
                                                End If
                                                .Col = e.col : .Row = e.row : .Focus() : .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                            Else
                                                rsItemMst = New ClsResultSetDB
                                                rsItemMst.GetResult("Select Tariff_Code from Item_Mst where Item_code = '" & varItemCode & "' and UNIT_CODE = '" & gstrUNITID & "'")
                                                Call .SetText(20, e.row, rsItemMst.GetValue("Tariff_Code"))
                                                If CheckForTariffCode(CStr(CDbl(rsItemMst.GetValue("Tariff_Code")))) = False Then
                                                    ClearGridRow(e.row)
                                                End If
                                                rsItemMst.ResultSetClose()
                                            End If
                                        Case "RAW MATERIAL"
                                            strStockLocation = StockLocationSalesConf((CmbInvType.Text), (CmbInvSubType.Text), "DESCRIPTION")
                                            If DisplayDetailsfromItemMst(CShort(e.row), CStr(varItemCode), strStockLocation, "'R'") = False Then
                                                If e.col = 1 Then
                                                    MsgBox("Invalid Item Code.", MsgBoxStyle.Information, "eMPro")
                                                Else
                                                    MsgBox("Invalid Customer Part Code.", MsgBoxStyle.Information, "eMPro")
                                                End If
                                                .Col = e.col : .Row = e.row : .Focus() : .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                            Else
                                                rsItemMst = New ClsResultSetDB
                                                rsItemMst.GetResult("Select Tariff_Code from Item_Mst where Item_code = '" & varItemCode & "' and UNIT_CODE = '" & gstrUNITID & "'")
                                                Call .SetText(20, e.row, rsItemMst.GetValue("Tariff_Code"))
                                                If CheckForTariffCode(CStr(CDbl(rsItemMst.GetValue("Tariff_Code")))) = False Then
                                                    ClearGridRow(e.row)
                                                End If
                                                rsItemMst.ResultSetClose()
                                            End If
                                        Case "COMPONENTS"
                                            strStockLocation = StockLocationSalesConf((CmbInvType.Text), (CmbInvSubType.Text), "DESCRIPTION")
                                            If DisplayDetailsfromItemMst(CShort(e.row), CStr(varItemCode), strStockLocation, "'C'") = False Then
                                                If e.col = 1 Then
                                                    MsgBox("Invalid Item Code.", MsgBoxStyle.Information, "eMPro")
                                                Else
                                                    MsgBox("Invalid Customer Part Code.", MsgBoxStyle.Information, "eMPro")
                                                End If
                                                .Col = e.col : .Row = e.row : .Focus() : .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                            Else
                                                rsItemMst = New ClsResultSetDB
                                                rsItemMst.GetResult("Select Tariff_Code from Item_Mst where Item_code = '" & varItemCode & "' and UNIT_CODE = '" & gstrUNITID & "'")
                                                Call .SetText(20, e.row, rsItemMst.GetValue("Tariff_Code"))
                                                If CheckForTariffCode(CStr(CDbl(rsItemMst.GetValue("Tariff_Code")))) = False Then
                                                    ClearGridRow(e.row)
                                                End If
                                                rsItemMst.ResultSetClose()
                                            End If
                                    End Select
                                End If
                            Else
                                If Len(Trim(txtRefNo.Text)) > 0 Then
                                    strStockLocation = StockLocationSalesConf(strInvTypeDes, strInvSubTypeDes, "DESCRIPTION")
                                    strSQL = "select a.Item_code,a.Item_code,c.Description, ISNULL(c.Tariff_code,0) from grn_dtl a,grn_hdr b,Item_Mst c where a.UNIT_CODE = b.UNIT_CODE AND a.UNIT_CODE = c.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND "
                                    strSQL = strSQL & "a.Doc_type = b.Doc_type and a.Doc_no = b.Doc_No "
                                    strSQL = strSQL & "and a.From_Location = b.From_Location "
                                    strSQL = strSQL & " and a.Item_Code = c.ITem_code and b.From_Location ='01R1'"
                                    strSQL = strSQL & " and c.Status = 'A' and Hold_Flag =0 "
                                    strSQL = strSQL & "and a.Rejected_quantity > 0 and b.Vendor_code = '" & txtCustCode.Text
                                    strSQL = strSQL & "' and a.Doc_No = " & CDbl(txtRefNo.Text)
                                    strSQL = strSQL & " and a.Item_code in (Select Item_Code from ItemBal_Mst Where Location_Code = '"
                                    strSQL = strSQL & strStockLocation & "' and UNIT_CODE = '" & gstrUNITID & "'  and Cur_bal > 0) and ((isnull(a.EXCESS_PO_Quantity,0) + isnull(a.Rejected_Quantity,0)) - isnull(a.Despatch_Quantity,0) - isnull(a.Inspected_Quantity,0) - isnull(a.RGP_Quantity,0)) > 0 and isnull(b.GRN_Cancelled,0) = 0"
                                    strSQL = strSQL & " and a.Item_Code ='" & varItemCode & "'"
                                    rsGrnDtl = New ClsResultSetDB
                                    rsGrnDtl.GetResult(strSQL)
                                    If rsGrnDtl.GetNoRows > 0 Then
                                        varChanged = Nothing
                                        Call .GetText(21, e.row, varChanged)
                                        'Value of rate changed then dont refresh
                                        If Trim(UCase(varChanged)) <> "C" Then
                                            strStockLocation = StockLocationSalesConf((CmbInvType.Text), (CmbInvSubType.Text), "DESCRIPTION")
                                            Call DisplayDetailsfromItemMst(CShort(e.row), CStr(varItemCode), strStockLocation)
                                        End If
                                    Else
                                        If e.col = 1 Then
                                            MsgBox("Invalid Item Code.", MsgBoxStyle.Information, "eMPro")
                                        Else
                                            MsgBox("Invalid Customer Part Code.", MsgBoxStyle.Information, "eMPro")
                                        End If
                                        .Col = e.col : .Row = e.row : .Focus() : .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                    End If
                                    rsGrnDtl.ResultSetClose()
                                Else
                                    varChanged = Nothing
                                    Call .GetText(21, e.row, varChanged)
                                    'Value of rate changed then dont refresh
                                    If Trim(UCase(varChanged)) <> "C" Then
                                        strStockLocation = StockLocationSalesConf((CmbInvType.Text), (CmbInvSubType.Text), "DESCRIPTION")
                                        If DisplayDetailsfromItemMst(CShort(e.row), CStr(varItemCode), strStockLocation) = False Then
                                            If e.col = 1 Then
                                                MsgBox("Invalid Item Code.", MsgBoxStyle.Information, "eMPro")
                                            Else
                                                MsgBox("Invalid Customer Part Code.", MsgBoxStyle.Information, "eMPro")
                                            End If
                                            .Col = e.col : .Row = e.row : .Focus() : .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                        End If
                                    End If
                                End If
                            End If
                    End Select
                ElseIf e.col = 7 Then
                    '101188073
                    If gblnGSTUnit Then Exit Sub
                    '101188073
                    .Col = 7
                    .Row = .ActiveRow
                    If Trim(.Text) <> "" Then
                        strWhereClause = " WHERE TxRt_Rate_No='" & Trim(.Text) & "' and UNIT_CODE = '" & gstrUnitId & "'  AND Tx_TaxeID='EXC'"
                        lstrReturnVal = SelectDataFromTable("TxRt_Rate_No", "Gen_TaxRate", strWhereClause)
                        If Len(lstrReturnVal) = 0 Then
                            .Text = ""
                            MsgBox("Invalid Tax Code", MsgBoxStyle.Information, "eMPro")
                            .Col = e.col : .Row = e.row : .Focus() : .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                        End If
                    End If
                    If Trim(.Text) <> "EX0" And Trim(.Text) <> "" Then
                        .Col = 1 : .Row = e.row
                        If ValidateTariffCode(Trim(.Text)) = False Then
                            MsgBox("Tariff Code is must for Item - " & (.Text), MsgBoxStyle.Information, "eMpro")
                            Exit Sub
                        End If
                    Else
                        If Trim(txtECESS.Text) <> "EC0" And Trim(txtECESS.Text) <> "" Then
                            MsgBox("ECESS can not be charged for this Invoice ", MsgBoxStyle.Information, "eMpro")
                            txtECESS.Text = ""
                            txtECESS.Focus()
                            Exit Sub
                        End If
                    End If
                ElseIf e.col = 8 Then
                    '101188073
                    If gblnGSTUnit Then Exit Sub
                    '101188073
                    .Col = 8
                    .Row = .ActiveRow
                    If Trim(.Text) <> "" Then
                        strWhereClause = " WHERE TxRt_Rate_No='" & Trim(.Text) & "' and UNIT_CODE = '" & gstrUnitId & "' AND Tx_TaxeID='CVD'"
                        lstrReturnVal = SelectDataFromTable("TxRt_Rate_No", "Gen_TaxRate", strWhereClause)
                        If Len(lstrReturnVal) = 0 Then
                            .Text = ""
                            MsgBox("Invalid Tax Code", MsgBoxStyle.Information, "eMPro")
                            .Col = e.col : .Row = e.row : .Focus() : .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                        End If
                    End If
                ElseIf e.col = 9 Then
                    '101188073
                    If gblnGSTUnit Then Exit Sub
                    '101188073
                    .Col = 9
                    .Row = .ActiveRow
                    If Trim(.Text) <> "" Then
                        strWhereClause = " WHERE TxRt_Rate_No='" & Trim(.Text) & "' and UNIT_CODE = '" & gstrUnitId & "' AND Tx_TaxeID='SAD'"
                        lstrReturnVal = SelectDataFromTable("TxRt_Rate_No", "Gen_TaxRate", strWhereClause)
                        If Len(lstrReturnVal) = 0 Then
                            .Text = ""
                            MsgBox("Invalid Tax Code", MsgBoxStyle.Information, "eMPro")
                            .Col = e.col : .Row = e.row : .Focus() : .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                        End If
                    End If
                End If
            End With
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub ctlFormHeader1_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlFormHeader1.Click
        On Error GoTo ErrHandler
        Call ShowHelp("HLPMKTTRN0005.HTM")
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub ctlInsurance_KeyPress(ByVal Sender As Object, ByVal e As UCActXCtl.UCctlFloat.KeyPressEventArgs) Handles ctlInsurance.KeyPress
        On Error GoTo ErrHandler
        Select Case e.KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        txtFreight.Focus()
                End Select
            Case 39, 34, 96
                e.KeyAscii = 0
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub ctlPerValue_KeyPress(ByVal Sender As Object, ByVal e As UCActXCtl.UCctlFloat.KeyPressEventArgs) Handles ctlPerValue.KeyPress
        If e.KeyAscii = System.Windows.Forms.Keys.Return Then
            txtECESS.Focus()
        End If
    End Sub
    Private Sub txtFreight_KeyPress(ByVal Sender As Object, ByVal e As UCActXCtl.UCctlFloat.KeyPressEventArgs) Handles txtFreight.KeyPress
        On Error GoTo ErrHandler
        Select Case e.KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        If (CmbInvType.Text = "SAMPLE INVOICE") Or (CmbInvType.Text = "TRANSFER INVOICE") Or (CmbInvType.Text = "JOBWORK INVOICE") Or (CmbInvType.Text = "CSM INVOICE") Then
                            If txtSurchargeTaxType.Enabled = True Then
                                txtSurchargeTaxType.Focus()
                            Else
                                txtECESS.Focus()
                            End If
                        Else
                            txtSaleTaxType.Focus()
                        End If
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        If txtSaleTaxType.Enabled Then
                            txtSaleTaxType.Focus()
                        Else
                            txtSurchargeTaxType.Focus()
                        End If
                End Select
            Case 39, 34, 96
                e.KeyAscii = 0
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub dtpDateDesc_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpDateDesc.KeyDown
        On Error GoTo Err_Handler
        If e.KeyCode = System.Windows.Forms.Keys.Return And e.Shift = 0 Then
            txtLocationCode.Focus()
        End If
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub dtpDateDesc_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpDateDesc.LostFocus
        On Error GoTo ErrHandler
        lblDateDes.Text = VB6.Format(dtpDateDesc.Value, gstrDateFormat)
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub ctlPerValue_TextChanged(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlPerValue.TextChanged
        Dim intLoopCounter As Short
        Dim varItemCode As Object
        Dim varDrgNo As Object
        Dim varRate As Object
        Dim varCustSupp As Object
        Dim varToolCost As Object
        Dim varOthers As Object
        On Error GoTo ErrHandler
        With SpChEntry
            If Len(Trim(ctlPerValue.Text)) = 0 Then ctlPerValue.Text = 1
            If Val(ctlPerValue.Text) > 1 Then
                .Row = 0 : .Col = 3
                .Text = "Rate (Per " & Val(ctlPerValue.Text) & ")" : .set_ColWidth(3, 1500)
                .Row = 0 : .Col = 4
                .Text = "Cust Supp Mat. (Per " & Val(ctlPerValue.Text) & ")" : .set_ColWidth(4, 1900)
                .Row = 0 : .Col = 15
                .Text = "Tool Cost (Per " & Val(ctlPerValue.Text) & ")" : .set_ColWidth(15, 1700)
                .Row = 0 : .Col = 10
                .Text = "Others (Per " & Val(ctlPerValue.Text) & ")" : .set_ColWidth(10, 1700)
            Else
                .Row = 0 : .Col = 3 : .Text = "Rate (Per Unit)" : .set_ColWidth(3, 1700)
                .Row = 0 : .Col = 4 : .Text = "Cust Supp Mat. (Per Unit)" : .set_ColWidth(4, 1900)
                .Row = 0 : .Col = 15 : .Text = "Tool Cost (Per Unit)" : .set_ColWidth(15, 1700)
                .Row = 0 : .Col = 10 : .Text = "Others (Per Unit)" : .set_ColWidth(10, 1700)
            End If
            For intLoopCounter = 1 To SpChEntry.MaxRows
                varDrgNo = Nothing
                Call .GetText(2, intLoopCounter, varDrgNo)
                varItemCode = Nothing
                Call .GetText(1, intLoopCounter, varItemCode)
                If (Len(Trim(CStr(varDrgNo))) > 0) And (Len(Trim(CStr(varItemCode))) > 0) Then
                    varRate = Nothing
                    Call .GetText(16, intLoopCounter, varRate)
                    varCustSupp = Nothing
                    Call .GetText(17, intLoopCounter, varCustSupp)
                    varToolCost = Nothing
                    Call .GetText(19, intLoopCounter, varToolCost)
                    varOthers = Nothing
                    Call .GetText(18, intLoopCounter, varOthers)
                    Call .SetText(3, intLoopCounter, varRate * CDbl(ctlPerValue.Text))
                    Call .SetText(4, intLoopCounter, Val(varCustSupp) * CDbl(ctlPerValue.Text))
                    Call .SetText(15, intLoopCounter, Val(varToolCost) * CDbl(ctlPerValue.Text))
                    Call .SetText(10, intLoopCounter, Val(varOthers) * CDbl(ctlPerValue.Text))
                End If
            Next
        End With
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub lblCurrencyDes_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblCurrencyDes.TextChanged
        On Error GoTo ErrHandler
        If Trim(lblCurrencyDes.Text) <> "" Then
            If Trim(lblCurrencyDes.Text) = Trim(gstrCURRENCYCODE) Then
                lblExchangeRateValue.Text = CStr(1.0#)
            Else
                If UCase(Trim(mstrInvoiceType)) = "INV" Or UCase(Trim(mstrInvoiceType)) = "SMP" Or UCase(Trim(mstrInvoiceType)) = "TRF" Or UCase(Trim(mstrInvoiceType)) = "JOB" Or UCase(Trim(mstrInvoiceType)) = "EXP" Or UCase(Trim(mstrInvoiceType)) = "CSM" Then
                    lblExchangeRateValue.Text = CStr(GetExchangeRate(lblCurrencyDes.Text, getDateForDB(VB6.Format(dtpDateDesc.Value, gstrDateFormat)), True))
                Else
                    lblExchangeRateValue.Text = CStr(GetExchangeRate(lblCurrencyDes.Text, getDateForDB(VB6.Format(dtpDateDesc.Value, gstrDateFormat)), False))
                End If
                If Val(Trim(lblExchangeRateValue.Text)) = 1 Then
                    MsgBox("Exchange Rate for " & Trim(lblCurrencyDes.Text) & " is not defined on " & VB6.Format(dtpDateDesc.Value, gstrDateFormat), MsgBoxStyle.Information, "eMPro")
                    lblExchangeRateValue.Text = ""
                End If
            End If
        Else
            lblExchangeRateValue.Text = ""
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub cmdAddVAT_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdAddVAT.Click
        Dim strHelp As String
        Dim strSTaxHelp() As String
        On Error GoTo ErrHandler
        '101188073
        If gblnGSTUnit Then Exit Sub
        '101188073
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                strHelp = "Select TxRT_Rate_No,TxRt_Percentage from Gen_taxRate where  UNIT_CODE = '" & gstrUNITID & "' and  Tx_TaxeID IN('ADVAT','ADCST')"
                strSTaxHelp = Me.ctlEMPHelpInvoiceEntry.ShowList(gstrCONNECTIONSERVER,gstrDSNName, gstrDatabaseName, strHelp, "Add. VAT/CST Tax Help")
                If UBound(strSTaxHelp) <= 0 Then Exit Sub
                If strSTaxHelp(0) = "0" Then
                    Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK) : txtAddVAT.Text = "" : txtAddVAT.Focus() : Exit Sub
                Else
                    txtAddVAT.Text = strSTaxHelp(0)
                    lblAddVAT.Text = strSTaxHelp(1)
                End If
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub txtAddVAT_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtAddVAT.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        If Len(txtAddVAT.Text) > 0 Then
                            Call txtAddVAT_Validating(txtAddVAT, New System.ComponentModel.CancelEventArgs(False))
                        Else
                            If txtSurchargeTaxType.Enabled Then txtSurchargeTaxType.Focus()
                        End If
                End Select
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        e.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            e.Handled = True
        End If
    End Sub
    Private Sub txtAddVAT_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtAddVAT.KeyUp
        Dim KeyCode As Short = e.KeyCode
        Dim Shift As Short = e.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = 112 Then
            If cmdAddVAT.Enabled Then Call cmdAddVAT_Click(cmdAddVAT, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtAddVAT_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAddVAT.TextChanged
        On Error GoTo ErrHandler
        If Len(txtAddVAT.Text) = 0 Then
            lblAddVAT.Text = "0.00"
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub txtAddVAT_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtAddVAT.Validating
        Dim Cancel As Boolean = e.Cancel
        On Error GoTo ErrHandler
        '101188073
        If gblnGSTUnit Then Exit Sub
        '101188073
        If Len(txtAddVAT.Text) > 0 Then
            If CheckExistanceOfFieldData((txtAddVAT.Text), "TxRt_Rate_No", "Gen_TaxRate", " (Tx_TaxeID='ADVAT' OR Tx_TaxeID='ADCST')") Then
                lblAddVAT.Text = CStr(GetTaxRate((txtAddVAT.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " (Tx_TaxeID='ADVAT' OR Tx_TaxeID='ADCST')"))
                If txtSurchargeTaxType.Enabled Then txtSurchargeTaxType.Focus()
            Else
                Call ConfirmWindow(10248, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                Cancel = True
                txtAddVAT.Text = ""
                If txtAddVAT.Enabled Then txtAddVAT.Focus()
            End If
        End If
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        e.Cancel = Cancel
    End Sub
    Private Sub txtSECESS_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtSECESS.Validating
        Dim Cancel As Boolean = e.Cancel
        On Error GoTo ErrHandler
        '101188073
        If gblnGSTUnit Then Exit Sub
        '101188073
        If Len(txtSECESS.Text) > 0 Then
            '------------------Satvir Handa------------------------
            If CheckExistanceOfFieldData((txtSECESS.Text), "TxRt_Rate_No", "Gen_TaxRate", " (Tx_TaxeID='ECSSH') and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))") Then
                '------------------Satvir Handa------------------------
                lblSECESS_Per.Text = CStr(GetTaxRate((txtSECESS.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " (Tx_TaxeID='ECSSH')"))
                If txtRemarks.Enabled Then txtRemarks.Focus()
            Else
                Call ConfirmWindow(10248, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                Cancel = True
                txtSECESS.Text = ""
                If txtSECESS.Enabled Then txtSECESS.Focus()
            End If
        End If
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        e.Cancel = Cancel
    End Sub
    Public Sub BlankRowCheckandDelete()
        On Error GoTo ErrHandler
        Dim Intcounter As Short
        Dim counter As Short
        Dim blnblankrowflag As Boolean
        With SpChEntry
            counter = .MaxRows
            For Intcounter = 1 To counter
                blnblankrowflag = False
                If Intcounter <= .MaxRows Then
                    blnblankrowflag = False
                    .Row = Intcounter
                    .Col = 1
                    If Len(Trim(.Text)) > 0 Then
                        blnblankrowflag = False
                    Else
                        blnblankrowflag = True
                    End If
                End If
                If blnblankrowflag = True Then
                    .Row = Intcounter
                    .Action = FPSpreadADO.ActionConstants.ActionDeleteRow
                    .MaxRows = .MaxRows - 1
                End If
            Next
        End With
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    '101188073
    Private Sub TaxesEnableDisable(ByRef txtTaxType As TextBox, Optional ByVal blnDisable As Boolean = False)
        If gblnGSTUnit Then
            txtTaxType.Enabled = False : txtTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        Else
            txtTaxType.Enabled = True : txtTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            If blnDisable Then
                txtTaxType.Enabled = False : txtTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            End If
        End If
    End Sub
    Private Sub TaxesHelpEnableDisable(ByRef btnHelp As Button, Optional ByRef blnDisable As Boolean = False)
        If gblnGSTUnit Then
            btnHelp.Enabled = False
        Else
            btnHelp.Enabled = True
            If blnDisable Then
                btnHelp.Enabled = False
            End If
        End If
    End Sub
    Private Sub TaxesLabelEnableDisable(ByRef lblTaxType As Label, Optional ByRef blnDisable As Boolean = False)
        If gblnGSTUnit Then
            lblTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        Else
            lblTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            If blnDisable Then
                lblTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            End If
        End If
    End Sub
    Private Sub GetGSTTaxes(ByVal rowIndex As Integer, ByVal strItemCode As String, ByVal strInvoiceType As String)
        Dim strSql As String = String.Empty
        Dim dt As New DataTable
        If Len(strItemCode) > 0 Then
            If strInvoiceType.ToUpper() = "REJECTION" Then
                strSql = "set dateformat 'dmy' SELECT * FROM DBO.UFN_GST_TAXES_REJECTIONINVOICE_DETAILS('" & gstrUnitId & "','" & txtCustCode.Text.Trim & "','" & strItemCode & "','" & GetServerDate() & "')"
            Else
                strSql = "set dateformat 'dmy' SELECT * FROM DBO.UFN_GST_ITEMWISETAXES('" & gstrUnitId & "','" & txtCustCode.Text.Trim & "','" & strItemCode & "','','')"
            End If
            dt = SqlConnectionclass.GetDataTable(strSql)
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                With SpChEntry
                    If strInvoiceType.ToUpper() = "REJECTION" Then
                        .SetText(IS_HSN_SAC, rowIndex, dt.Rows(0)("HSNSAC_CODE_TYPE"))
                        .SetText(HSN_SAC_CODE, rowIndex, dt.Rows(0)("ISHSNORSAC"))
                        .SetText(CGST_TYPE, rowIndex, dt.Rows(0)("CGST_TXRT_HEAD"))
                        .SetText(SGST_TYPE, rowIndex, dt.Rows(0)("SGST_TXRT_HEAD"))
                        .SetText(IGST_TYPE, rowIndex, dt.Rows(0)("IGST_TXRT_HEAD"))
                        .SetText(UTGST_TYPE, rowIndex, dt.Rows(0)("UGST_TXRT_HEAD"))
                        .SetText(COMP_CESS_TYPE, rowIndex, dt.Rows(0)("COMPENSATION_CESS"))
                    Else
                        .SetText(IS_HSN_SAC, rowIndex, dt.Rows(0)("ISHSNORSAC"))
                        .SetText(HSN_SAC_CODE, rowIndex, dt.Rows(0)("HSNSACCODE"))
                        .SetText(CGST_TYPE, rowIndex, dt.Rows(0)("CGST_TXRT_HEAD"))
                        .SetText(SGST_TYPE, rowIndex, dt.Rows(0)("SGST_TXRT_HEAD"))
                        .SetText(IGST_TYPE, rowIndex, dt.Rows(0)("IGST_TXRT_HEAD"))
                        .SetText(UTGST_TYPE, rowIndex, dt.Rows(0)("UGST_TXRT_HEAD"))
                        .SetText(COMP_CESS_TYPE, rowIndex, dt.Rows(0)("COMPENSATION_CESS"))
                    End If
                End With
                dt.Dispose()
            End If
        End If
    End Sub
    Private Function GetGSTTaxesPercentage(ByVal strCGST As String, ByVal strSGST As String, ByVal strIGST As String, ByVal strUTGST As String, ByVal strCompCess As String) As DataTable
        Dim strSql As String = String.Empty
        Dim dtGSTTaxes As New DataTable
        strSql = "SELECT CGST_PERCENT,SGST_PERCENT,IGST_PERCENT,UTGST_PERCENT,COMPENSATION_CESS_PERCENT FROM dbo.UDF_GST_TAX_RATE_PERCENT('" & gstrUnitId & "','" & Convert.ToString(strCGST) & "','" & Convert.ToString(strSGST) & "','" & Convert.ToString(strIGST) & "','" & Convert.ToString(strUTGST) & "','" & Convert.ToString(strCompCess) & "')"
        dtGSTTaxes = SqlConnectionclass.GetDataTable(strSql)
        Return dtGSTTaxes
    End Function
    Private Function CalculateGSTTaxes(ByVal dblTaxableValue As Double, ByVal dblTaxPercent As Double) As Double
        CalculateGSTTaxes = ((dblTaxableValue * dblTaxPercent) / 100)
    End Function
    Private Function SaveDataGST(ByVal Button As String) As Boolean
        Dim ldblTotalBasicValue As Double
        Dim ldblTotalAccessibleValue As Double
        Dim lintLoopCounter As Short
        Dim ldblTotalSaleTaxAmount As Double
        Dim ldblTotalSurchargeTaxAmount As Double
        Dim ldblNetInsurenceValue As Double
        Dim ldblTotalInvoiceValue As Double
        Dim ldblTotalOthersValues As Double
        Dim rsParameterData As ClsResultSetDB
        Dim strParamQuery As String
        ''-----------Variable For Saving Data---------
        Dim strSalesChallan As String
        Dim updateSalesChallan As String
        Dim strSalesDtl As String
        Dim strSalesDtlDelete As String
        Dim rsCustItemMst As ClsResultSetDB
        Dim rsSaleConf As ClsResultSetDB
        Dim rsItemMst As ClsResultSetDB
        Dim rsPacking_Tax As ClsResultSetDB
        Dim lintItemQuantity As Double
        Dim lstrItemDrgno As String
        Dim lstrItemCode As String
        Dim ldblItemRate As Double
        Dim ldblItemCustMtrl As Double
        Dim ldblItemPacking As Double
        Dim strPackingCode As String
        Dim ldblItemOthers As Double
        Dim ldblItemFromBox As Double
        Dim ldblItemToBox As Double
        Dim lstrItemDelete As String
        Dim lintItemPresQty As Double
        Dim lstrItemExciseCode As String
        Dim lstrItemCVDCode As String
        Dim lstrItemSADCode As String
        Dim ldblItemToolCost As Double
        Dim TempAccessibleVal As Double
        Dim ldblTotalCustMatrlValue As Double
        Dim blnISInsExcisable As Boolean
        Dim blnISECESSRoundoff As Boolean
        Dim blnEOUFlag As Boolean
        Dim blnISExciseRoundOff As Boolean
        Dim blnISSalesTaxRoundOff As Boolean
        Dim blnISSurChargeTaxRoundOff As Boolean
        Dim blnAddCustMatrl As Boolean
        Dim blnISBasicRoundOff As Boolean
        Dim ldblExciseValueForSaleTax As Double
        Dim ldblTotalECESSAmount As Double
        Dim ldblTotalSECESSAmount As Double
        Dim blnDSWiseTracking As Boolean
        Dim strSDTType As String
        Dim dblSDT_Per As Double
        Dim dblSDT_Amt As Double
        Dim blnIsSDTRoundoff As Boolean
        Dim intSDTNoofDecimal As Short
        Dim dblBinQty As Double
        Dim strQry As String
        Dim rsCustOrdDtl As ClsResultSetDB
        Dim dblCustMtrl_SO As Double
        Dim intSTaxNoOfDecimal As Short
        Dim intEcessNoOfDecimal As Short
        Dim blnPackingRoundoff As Boolean
        Dim intPackingRoundoff_Decimal As Short
        Dim dblTotalPacking_Amount As Double
        Dim dblItemPacking_Amount As Double
        Dim dblAddVATamount As Double
        Dim dblExcise_Amount As Double
        Dim strSqlct2qry As String
        Dim strsql As String
        Dim blnIsCt2 As Boolean = False

        On Error GoTo ErrHandler

        ldblTotalBasicValue = 0
        ldblTotalAccessibleValue = 0
        ldblTotalSaleTaxAmount = 0
        ldblTotalSurchargeTaxAmount = 0
        ldblTotalInvoiceValue = 0
        ldblTotalOthersValues = 0
        ldblTotalCustMatrlValue = 0
        ldblExciseValueForSaleTax = 0
        ldblTotalECESSAmount = 0
        ldblTotalSECESSAmount = 0
        dblSDT_Amt = 0
        dblCustMtrl_SO = 0
        intSTaxNoOfDecimal = 0
        intEcessNoOfDecimal = 0
        dblTotalPacking_Amount = 0
        Dim dblCGSTAMT As Double = 0
        Dim dblSGSTAMT As Double = 0
        Dim dblIGSTAMT As Double = 0
        Dim dblUTGSTAMT As Double = 0
        Dim dblCOMPCESSAMT As Double = 0
        Dim dblToolAmmortization As Double = 0
        Dim dblTaxableValue As Double = 0
        Dim CGSTType As String = String.Empty
        Dim SGSTType As String = String.Empty
        Dim IGSTType As String = String.Empty
        Dim UTGSTType As String = String.Empty
        Dim COMPCESSType As String = String.Empty
        Dim dtGSTTaxPercent As New DataTable
        Dim strHSNSACCode As String = String.Empty
        Dim strHSNSACType As String = String.Empty
        Dim dblCGSTPercentLine As Double = 0
        Dim dblSGSTPercentLine As Double = 0
        Dim dblIGSTPercentLine As Double = 0
        Dim dblUTGSTPercentLine As Double = 0
        Dim dblCCESSPercentLine As Double = 0
        Dim dblCGSTAmtLine As Double = 0
        Dim dblSGSTAmtLine As Double = 0
        Dim dblIGSTAmtLine As Double = 0
        Dim dblUTGSTAmtLine As Double = 0
        Dim dblCCESSAmtLine As Double = 0
        Dim dblItemTotalLine As Double = 0
        Dim blnGSTTAXroundoff As Boolean
        Dim intGSTTAXroundoff_decimal As Short

        SaveDataGST = True
        strParamQuery = "SELECT InsExc_Excise,CustSupp_Inc,EOU_Flag,SalesTax_Roundoff,Basic_roundoff,Excise_Roundoff,SST_Roundoff,ECESS_Roundoff,DSWiseTracking,TotalInvoiceAmount_RoundOff,TotalInvoiceAmountRoundOff_Decimal, SDTRoundOff, SDTRoundOff_Decimal,salesTax_Roundoff_decimal,ECESSRoundoff_decimal,Packing_Roundoff,PackingRoundoff_Decimal,GSTTAX_ROUNDOFF_DECIMAL,GSTTAX_ROUNDOFF FROM Sales_Parameter WHERE UNIT_CODE = '" & gstrUNITID & "'"
        rsParameterData = New ClsResultSetDB
        rsParameterData.GetResult(strParamQuery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsParameterData.GetNoRows > 0 Then
            blnISInsExcisable = rsParameterData.GetValue("InsExc_Excise")
            blnEOUFlag = rsParameterData.GetValue("EOU_Flag")
            blnISExciseRoundOff = rsParameterData.GetValue("Excise_Roundoff")
            blnISBasicRoundOff = rsParameterData.GetValue("Basic_Roundoff")
            blnISSalesTaxRoundOff = rsParameterData.GetValue("SalesTax_Roundoff")
            blnISSurChargeTaxRoundOff = rsParameterData.GetValue("SST_Roundoff")
            blnAddCustMatrl = rsParameterData.GetValue("CustSupp_Inc")
            blnISECESSRoundoff = rsParameterData.GetValue("ECESS_Roundoff")
            blnDSWiseTracking = IIf(IsDBNull(rsParameterData.GetValue("DSWiseTracking")), False, IIf(rsParameterData.GetValue("DSWiseTracking") = False, False, True))
            blnTotalInvoiceAmount = rsParameterData.GetValue("TotalInvoiceAmount_RoundOff")
            If rsParameterData.GetValue("TotalInvoiceAmountRoundOff_Decimal").ToString = "" Then
                intTotalInvoiceAmountRoundOffDecimal = 0
            Else
                intTotalInvoiceAmountRoundOffDecimal = rsParameterData.GetValue("TotalInvoiceAmountRoundOff_Decimal")
            End If
            blnIsSDTRoundoff = rsParameterData.GetValue("SDTRoundOff")
            intSDTNoofDecimal = rsParameterData.GetValue("SDTRoundOff_Decimal")
            If rsParameterData.GetValue("SDTRoundOff_Decimal").ToString = "" Then
                intSDTNoofDecimal = 0
            Else
                intSDTNoofDecimal = rsParameterData.GetValue("SDTRoundOff_Decimal")
            End If
            intSTaxNoOfDecimal = rsParameterData.GetValue("salesTax_Roundoff_decimal")
            If rsParameterData.GetValue("ECESSRoundoff_decimal").ToString = "" Then
                intEcessNoOfDecimal = 0
            Else
                intEcessNoOfDecimal = rsParameterData.GetValue("ECESSRoundoff_decimal")
            End If
            blnPackingRoundoff = rsParameterData.GetValue("Packing_Roundoff")
            If rsParameterData.GetValue("PackingRoundoff_Decimal").ToString = "" Then
                intPackingRoundoff_Decimal = 0
            Else
                intPackingRoundoff_Decimal = rsParameterData.GetValue("PackingRoundoff_Decimal")
            End If
            blnGSTTAXroundoff = rsParameterData.GetValue("GSTTAX_ROUNDOFF")
            intGSTTAXroundoff_decimal = rsParameterData.GetValue("GSTTAX_ROUNDOFF_DECIMAL")
        Else
            MsgBox("No data define in Sales_Parameter Table", MsgBoxStyle.Information, "eMPro")
            SaveDataGST = False
            rsParameterData.ResultSetClose()
            rsParameterData = Nothing
            Exit Function
        End If
        rsParameterData.ResultSetClose()
        rsParameterData = Nothing
        strParamQuery = "SELECT decimal_place FROM currency_mst where currency_code='" & lblCurrencyDes.Text & "' and UNIT_CODE = '" & gstrUnitId & "'"
        rsParameterData = New ClsResultSetDB
        rsParameterData.GetResult(strParamQuery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsParameterData.GetNoRows > 0 Then
            mIntDecimalPlace = rsParameterData.GetValue("decimal_place")
        End If
        rsParameterData.ResultSetClose()
        rsParameterData = Nothing
        ldblNetInsurenceValue = Math.Round(Val(ctlInsurance.Text)) / Val(CStr(SpChEntry.MaxRows))
        For lintLoopCounter = 1 To SpChEntry.MaxRows
            dblTotalPacking_Amount += CalculatePackingValue(lintLoopCounter, blnPackingRoundoff)
            ldblTotalBasicValue += CalculateBasicValue(lintLoopCounter, blnISBasicRoundOff)
            dblTaxableValue = CalculateAccessibleValue(lintLoopCounter, ldblNetInsurenceValue, blnISInsExcisable)
            ldblTotalAccessibleValue += dblTaxableValue
            SpChEntry.Row = lintLoopCounter : SpChEntry.Col = 5
            lintItemQuantity = Val(SpChEntry.Text)
            SpChEntry.Row = lintLoopCounter : SpChEntry.Col = 10
            ldblTotalOthersValues = ldblTotalOthersValues + ((Val(SpChEntry.Text) / CDbl(ctlPerValue.Text)) * lintItemQuantity)
            SpChEntry.Row = lintLoopCounter : SpChEntry.Col = 4
            ldblTotalCustMatrlValue = ldblTotalCustMatrlValue + ((Val(SpChEntry.Text) / CDbl(ctlPerValue.Text)) * lintItemQuantity)
            SpChEntry.Row = lintLoopCounter : SpChEntry.Col = 15
            dblToolAmmortization = dblToolAmmortization + Math.Round(Val(CStr(lintItemQuantity * (Val(SpChEntry.Text) / CDbl(ctlPerValue.Text)))), 2)
            SpChEntry.Row = lintLoopCounter : SpChEntry.Col = CGST_TYPE
            CGSTType = SpChEntry.Text
            SpChEntry.Row = lintLoopCounter : SpChEntry.Col = SGST_TYPE
            SGSTType = SpChEntry.Text
            SpChEntry.Row = lintLoopCounter : SpChEntry.Col = IGST_TYPE
            IGSTType = SpChEntry.Text
            SpChEntry.Row = lintLoopCounter : SpChEntry.Col = UTGST_TYPE
            UTGSTType = SpChEntry.Text
            SpChEntry.Row = lintLoopCounter : SpChEntry.Col = COMP_CESS_TYPE
            COMPCESSType = SpChEntry.Text
            dtGSTTaxPercent = GetGSTTaxesPercentage(CGSTType, SGSTType, IGSTType, UTGSTType, COMPCESSType)
            If dtGSTTaxPercent IsNot Nothing AndAlso dtGSTTaxPercent.Rows.Count > 0 Then
                If blnGSTTAXroundoff Then
                    dblCGSTAMT += CalculateGSTTaxes(dblTaxableValue, dtGSTTaxPercent.Rows(0)("CGST_PERCENT"))
                    dblSGSTAMT += CalculateGSTTaxes(dblTaxableValue, dtGSTTaxPercent.Rows(0)("SGST_PERCENT"))
                    dblIGSTAMT += CalculateGSTTaxes(dblTaxableValue, dtGSTTaxPercent.Rows(0)("IGST_PERCENT"))
                    dblUTGSTAMT += CalculateGSTTaxes(dblTaxableValue, dtGSTTaxPercent.Rows(0)("UTGST_PERCENT"))
                    dblCOMPCESSAMT += CalculateGSTTaxes(dblTaxableValue, dtGSTTaxPercent.Rows(0)("COMPENSATION_CESS_PERCENT"))
                Else
                    dblCGSTAMT += System.Math.Round(CalculateGSTTaxes(dblTaxableValue, dtGSTTaxPercent.Rows(0)("CGST_PERCENT")), intGSTTAXroundoff_decimal)
                    dblSGSTAMT += System.Math.Round(CalculateGSTTaxes(dblTaxableValue, dtGSTTaxPercent.Rows(0)("SGST_PERCENT")), intGSTTAXroundoff_decimal)
                    dblIGSTAMT += System.Math.Round(CalculateGSTTaxes(dblTaxableValue, dtGSTTaxPercent.Rows(0)("IGST_PERCENT")), intGSTTAXroundoff_decimal)
                    dblUTGSTAMT += System.Math.Round(CalculateGSTTaxes(dblTaxableValue, dtGSTTaxPercent.Rows(0)("UTGST_PERCENT")), intGSTTAXroundoff_decimal)
                    dblCOMPCESSAMT += System.Math.Round(CalculateGSTTaxes(dblTaxableValue, dtGSTTaxPercent.Rows(0)("COMPENSATION_CESS_PERCENT")), intGSTTAXroundoff_decimal)
                End If
                
            End If
        Next

        If Val(ldblTotalBasicValue) = 0 Then
            MsgBox("Total Basic Amt. Can not be 0.", MsgBoxStyle.Information, "eMPro")
            SaveDataGST = False
            Exit Function
        ElseIf Val(ldblTotalAccessibleValue) = 0 Then
            MsgBox("Total Assessable Amt. Can not be 0.", MsgBoxStyle.Information, "eMPro")
            SaveDataGST = False
            Exit Function
        End If
        If blnAddCustMatrl Then
            ldblTotalInvoiceValue = ldblTotalBasicValue + dblTotalPacking_Amount + System.Math.Round(Val(txtFreight.Text)) + System.Math.Round(ldblTotalOthersValues) + System.Math.Round(Val(ctlInsurance.Text)) + System.Math.Round(ldblTotalCustMatrlValue) + dblCGSTAMT + dblSGSTAMT + dblIGSTAMT + dblUTGSTAMT + dblCOMPCESSAMT
        Else
            ldblTotalInvoiceValue = ldblTotalBasicValue + dblTotalPacking_Amount + System.Math.Round(Val(txtFreight.Text)) + System.Math.Round(ldblTotalOthersValues) + System.Math.Round(Val(ctlInsurance.Text)) + dblCGSTAMT + dblSGSTAMT + dblIGSTAMT + dblUTGSTAMT + dblCOMPCESSAMT
        End If

        If blnTotalInvoiceAmount Then
            ldblTotalInvoiceValueRoundOff = ldblTotalInvoiceValue - System.Math.Round(ldblTotalInvoiceValue)
            ldblTotalInvoiceValue = System.Math.Round(ldblTotalInvoiceValue)
        Else
            ldblTotalInvoiceValueRoundOff = ldblTotalInvoiceValue - System.Math.Round(ldblTotalInvoiceValue, intTotalInvoiceAmountRoundOffDecimal)
            ldblTotalInvoiceValue = System.Math.Round(ldblTotalInvoiceValue, intTotalInvoiceAmountRoundOffDecimal)
        End If

        Dim strStock_Loc As String
        strStock_Loc = StockLocationSalesConf((Me.CmbInvType.Text), (Me.CmbInvSubType).Text, "DESCRIPTION")
        Select Case Button
            Case "ADD"
                rsSaleConf = New ClsResultSetDB
                rsSaleConf.GetResult("Select Invoice_Type,Sub_Type from SaleConf where  UNIT_CODE = '" & gstrUnitId & "'  and  Description ='" & Trim(CmbInvType.Text) & "' and Sub_type_Description ='" & Trim(CmbInvSubType.Text) & "' and (fin_start_date <= getdate() and fin_end_date >= getdate())")
                strSalesChallan = ""
                If UCase(CmbInvType.Text) <> "JOBWORK INVOICE" Then
                    mstrRGP = ""
                End If
                If (UCase(CmbInvType.Text) = "CSM INVOICE") And chkFOC.CheckState Then
                    ldblTotalInvoiceValue = ldblTotalInvoiceValue - ldblTotalBasicValue
                End If
                strSalesChallan = "INSERT INTO SalesChallan_dtl (UNIT_CODE,Location_Code,Packing_amount,Doc_No,Suffix,Transport_Type,Vehicle_No,From_Location,"
                strSalesChallan = strSalesChallan & "From_Station,To_Station,Account_Code,"
                strSalesChallan = strSalesChallan & "Cust_Ref,Amendment_No,Bill_Flag,Form3,Carriage_Name,"
                strSalesChallan = strSalesChallan & "Year,Insurance,invoice_Type,Ref_Doc_No,"
                strSalesChallan = strSalesChallan & "Cust_Name ,Sales_Tax_Amount , Surcharge_Sales_Tax_Amount,"
                strSalesChallan = strSalesChallan & "Frieght_Amount,Sub_Category,SalesTax_Type,SalesTax_FormNo,"
                strSalesChallan = strSalesChallan & "SalesTax_FormValue,Annex_no,invoice_Date,Currency_code,Ent_dt,"
                strSalesChallan = strSalesChallan & "Ent_UserId,Upd_dt,Upd_UserId,Exchange_Rate,total_amount,"
                strSalesChallan = strSalesChallan & "Surcharge_salesTaxType,SalesTax_Per,Surcharge_SalesTax_Per,PerValue,"
                strSalesChallan = strSalesChallan & "Remarks,SRVDINO,SRVLocation,ECESS_Type,ECESS_Per,ECESS_Amount,SECESS_Type,SECESS_Per,SECESS_Amount"
                If UCase(CmbInvType.Text) = "JOBWORK INVOICE" Then
                    strSalesChallan = strSalesChallan & ",FIFO_Flag "
                End If
                strSalesChallan = strSalesChallan & ",USLOC,SchTime,TotalInvoiceAmtRoundOff_diff"
                strSalesChallan = strSalesChallan & ",Payment_Terms"
                If UCase(CmbInvType.Text) <> "JOBWORK INVOICE" And UCase(CmbInvType.Text) <> "TRANSFER INVOICE" Then
                    strSalesChallan = strSalesChallan & ", SDTax_Type, SDTax_Per, SDTax_Amount "
                End If
                strSalesChallan = strSalesChallan & ",ADDVAT_Type,ADDVAT_Per,ADDVAT_Amount"
                strSalesChallan = strSalesChallan & ",FOC_Invoice,CGST_TOTAL_AMT,SGST_TOTAL_AMT,IGST_TOTAL_AMT,UTGST_TOTAL_AMT,CCESS_TOTAL_AMT)"
                strSalesChallan = strSalesChallan & " Values ('" & gstrUnitId & "','" & Trim(txtLocationCode.Text)
                strSalesChallan = strSalesChallan & "', " & dblTotalPacking_Amount & ","
                strSalesChallan = strSalesChallan & "'" & Trim(txtChallanNo.Text) & "',''"
                strSalesChallan = strSalesChallan & ",'" & Mid(Trim(CmbTransType.Text), 1, 1) & "', '" & Trim(txtVehNo.Text) & "','" & Trim(strStock_Loc) & "','"
                strSalesChallan = strSalesChallan & "','','" & Trim(txtCustCode.Text)
                strSalesChallan = strSalesChallan & "','" & Trim(txtRefNo.Text) & "','" & Trim(txtAmendNo.Text) & "','0'"
                strSalesChallan = strSalesChallan & ",'','" & Trim(txtCarrServices.Text)
                strSalesChallan = strSalesChallan & "','" & Trim(CStr(Year(dtpDateDesc.Value))) & "',"
                strSalesChallan = strSalesChallan & System.Math.Round(Val(ctlInsurance.Text)) & ",'" & Trim(rsSaleConf.GetValue("Invoice_type")) & "','"
                strSalesChallan = strSalesChallan & Trim(mstrRGP) & "','"
                strSalesChallan = strSalesChallan & Trim(lblCustCodeDes.Text) & "',"
                strSalesChallan = strSalesChallan & Val(CStr(ldblTotalSaleTaxAmount)) & "," & Val(CStr(ldblTotalSurchargeTaxAmount)) & "," & System.Math.Round(Val(txtFreight.Text)) & ",'" & Trim(rsSaleConf.GetValue("Sub_Type")) & "','"
                strSalesChallan = strSalesChallan & Trim(txtSaleTaxType.Text) & "','"
                strSalesChallan = strSalesChallan & "0',0,'0','"
                strSalesChallan = strSalesChallan & VB6.Format(dtpDateDesc.Value, "dd/mmm/yyyy") & "','" & lblCurrencyDes.Text & "',getdate(),'" & mP_User & "',  getdate() ,'" & mP_User & "','"
                strSalesChallan = strSalesChallan & lblExchangeRateValue.Text & "'," & ldblTotalInvoiceValue & ",'"
                strSalesChallan = strSalesChallan & Trim(txtSurchargeTaxType.Text) & "'," & Val(lblSaltax_Per.Text) & ","
                strSalesChallan = strSalesChallan & Val(lblSurcharge_Per.Text) & "," & ctlPerValue.Text & ",'" & txtRemarks.Text & "','"
                strSalesChallan = strSalesChallan & Trim(txtSRVDI.Text) & "','" & Trim(txtSRVLoc.Text) & "','"
                strSalesChallan = strSalesChallan & Trim(txtECESS.Text) & "'," & Val(lblECESS_Per.Text) & "," & Val(CStr(ldblTotalECESSAmount)) & ",'"
                strSalesChallan = strSalesChallan & Trim(txtSECESS.Text) & "'," & Val(lblSECESS_Per.Text) & "," & Val(CStr(ldblTotalSECESSAmount))
                If UCase(CmbInvType.Text) = "JOBWORK INVOICE" Then
                    If blnFIFO = True Then
                        strSalesChallan = strSalesChallan & ",1"
                    Else
                        strSalesChallan = strSalesChallan & ",0"
                    End If
                End If
                strSalesChallan = strSalesChallan & ",'" & txtUsLoc.Text & "','" & txtSchTime.Text & "'"
                strSalesChallan = strSalesChallan & "," & ldblTotalInvoiceValueRoundOff
                strSalesChallan = strSalesChallan & ",'" & Trim(lblCreditTerm.Text) & "'"
                If UCase(CmbInvType.Text) <> "JOBWORK INVOICE" And UCase(CmbInvType.Text) <> "TRANSFER INVOICE" Then
                    strSalesChallan = strSalesChallan & ", '" & Trim(strSDTType) & "', " & dblSDT_Per & ", " & dblSDT_Amt
                End If
                strSalesChallan = strSalesChallan & ", '" & Trim(txtAddVAT.Text) & "', " & Val(lblAddVAT.Text) & ", " & dblAddVATamount
                strSalesChallan = strSalesChallan & "," & IIf(chkFOC.Checked, 1, 0) & "," & dblCGSTAMT & "," & dblSGSTAMT & "," & dblIGSTAMT & "," & dblUTGSTAMT & "," & dblCOMPCESSAMT & " )"
                rsSaleConf.ResultSetClose()
                rsSaleConf = Nothing
                strSalesDtl = ""
                With SpChEntry
                    For lintLoopCounter = 1 To .MaxRows
                        .Row = lintLoopCounter
                        .Col = 1
                        lstrItemCode = Trim(.Text)
                        .Col = 2
                        lstrItemDrgno = Trim(.Text)
                        .Col = 3
                        ldblItemRate = Val(.Text) / CDbl(ctlPerValue.Text)
                        .Col = 4
                        ldblItemCustMtrl = Val(.Text) / CDbl(ctlPerValue.Text)
                        .Col = 5
                        lintItemQuantity = Val(.Text)
                        .Col = 6
                        strPackingCode = Trim(.Text)
                        rsPacking_Tax = New ClsResultSetDB
                        rsPacking_Tax.GetResult("Select Txrt_Rate_no,TxRt_Percentage from Gen_TaxRate where Tx_TaxeID = 'PKT' and Txrt_Rate_no = '" & Trim(strPackingCode) & "' and UNIT_CODE = '" & gstrUnitId & "'")
                        If rsPacking_Tax.GetNoRows > 0 Then
                            ldblItemPacking = rsPacking_Tax.GetValue("TxRt_Percentage")
                        End If
                        rsPacking_Tax.ResultSetClose()
                        .Col = 7
                        lstrItemExciseCode = Trim(.Text)
                        .Col = 8
                        lstrItemCVDCode = Trim(.Text)
                        .Col = 9
                        lstrItemSADCode = Trim(.Text)
                        .Col = 10
                        ldblItemOthers = Val(.Text) / CDbl(ctlPerValue.Text) * lintItemQuantity
                        .Col = 11
                        ldblItemFromBox = Val(.Text)
                        .Col = 12
                        ldblItemToBox = Val(.Text)
                        .Col = 14
                        lstrItemDelete = Trim(.Text)
                        .Col = 22
                        dblBinQty = Val(.Text)
                        If dblBinQty <= 0 Then
                            MsgBox("Bin Quantity can't be zero.", MsgBoxStyle.Information, "eMpro")
                            SaveDataGST = False
                            Exit Function
                        End If
                        .Col = 15
                        ldblItemToolCost = Val(.Text) / CDbl(ctlPerValue.Text)
                        If Val(CStr(ldblItemCustMtrl)) > 0 Then
                            strQry = ""
                            strQry = "Select Cust_Mtrl from Cust_ord_dtl WHERE "
                            strQry = strQry & "Account_Code ='" & txtCustCode.Text & "'and Cust_ref ='"
                            strQry = strQry & txtRefNo.Text & "' and Amendment_No = '" & Trim(txtAmendNo.Text) & "'and "
                            strQry = strQry & " Active_flag ='A' "
                            strQry = strQry & " and Cust_DrgNo = '" & Trim(lstrItemDrgno) & "'"
                            strQry = strQry & " and Item_Code ='" & Trim(lstrItemCode) & "' and UNIT_CODE = '" & gstrUnitId & "'"
                            rsCustOrdDtl = New ClsResultSetDB
                            rsCustOrdDtl.GetResult(strQry)
                            If rsCustOrdDtl.GetNoRows > 0 Then
                                dblCustMtrl_SO = rsCustOrdDtl.GetValue("Cust_Mtrl")
                            End If
                            If Val(CStr(dblCustMtrl_SO)) = 0 Then
                                If Val(CStr(ldblItemCustMtrl)) > 0 Then
                                    ldblItemCustMtrl = 0
                                End If
                            End If
                            rsCustOrdDtl.ResultSetClose()
                            rsCustOrdDtl = Nothing
                        End If
                        TempAccessibleVal = CalculateAccessibleValue(lintLoopCounter, ldblNetInsurenceValue, blnISInsExcisable)
                        .Col = HSN_SAC_CODE
                        strHSNSACCode = .Text
                        .Col = IS_HSN_SAC
                        strHSNSACType = .Text
                        .Col = CGST_TYPE
                        CGSTType = .Text
                        .Col = SGST_TYPE
                        SGSTType = .Text
                        .Col = IGST_TYPE
                        IGSTType = .Text
                        .Col = UTGST_TYPE
                        UTGSTType = .Text
                        .Col = COMP_CESS_TYPE
                        COMPCESSType = .Text
                        dblCGSTPercentLine = 0
                        dblCGSTAmtLine = 0
                        dblSGSTPercentLine = 0
                        dblSGSTAmtLine = 0
                        dblIGSTPercentLine = 0
                        dblIGSTAmtLine = 0
                        dblUTGSTPercentLine = 0
                        dblUTGSTAmtLine = 0
                        dblCCESSPercentLine = 0
                        dblCCESSAmtLine = 0

                        dtGSTTaxPercent = New DataTable()
                        dtGSTTaxPercent = GetGSTTaxesPercentage(CGSTType, SGSTType, IGSTType, UTGSTType, COMPCESSType)
                        If dtGSTTaxPercent IsNot Nothing AndAlso dtGSTTaxPercent.Rows.Count > 0 Then
                            dblCGSTPercentLine = dtGSTTaxPercent.Rows(0)("CGST_PERCENT")
                            dblSGSTPercentLine = dtGSTTaxPercent.Rows(0)("SGST_PERCENT")
                            dblIGSTPercentLine = dtGSTTaxPercent.Rows(0)("IGST_PERCENT")
                            dblUTGSTPercentLine = dtGSTTaxPercent.Rows(0)("UTGST_PERCENT")
                            dblCCESSPercentLine = dtGSTTaxPercent.Rows(0)("COMPENSATION_CESS_PERCENT")
                            If blnGSTTAXroundoff Then
                                dblCGSTAmtLine = CalculateGSTTaxes(TempAccessibleVal, dblCGSTPercentLine)
                                dblSGSTAmtLine = CalculateGSTTaxes(TempAccessibleVal, dblSGSTPercentLine)
                                dblIGSTAmtLine = CalculateGSTTaxes(TempAccessibleVal, dblIGSTPercentLine)
                                dblUTGSTAmtLine = CalculateGSTTaxes(TempAccessibleVal, dblUTGSTPercentLine)
                                dblCCESSAmtLine = CalculateGSTTaxes(TempAccessibleVal, dblCCESSPercentLine)
                            Else
                                dblCGSTAmtLine = System.Math.Round(CalculateGSTTaxes(TempAccessibleVal, dblCGSTPercentLine), intGSTTAXroundoff_decimal)
                                dblSGSTAmtLine = System.Math.Round(CalculateGSTTaxes(TempAccessibleVal, dblSGSTPercentLine), intGSTTAXroundoff_decimal)
                                dblUTGSTAmtLine = System.Math.Round(CalculateGSTTaxes(TempAccessibleVal, dblUTGSTPercentLine), intGSTTAXroundoff_decimal)
                                dblIGSTAmtLine = System.Math.Round(CalculateGSTTaxes(TempAccessibleVal, dblIGSTPercentLine), intGSTTAXroundoff_decimal)
                                dblCCESSAmtLine = System.Math.Round(CalculateGSTTaxes(TempAccessibleVal, dblCCESSPercentLine), intGSTTAXroundoff_decimal)
                            End If
                            dtGSTTaxPercent.Dispose()
                        End If
                        dblItemTotalLine = TempAccessibleVal + dblCGSTAmtLine + dblSGSTAmtLine + dblIGSTAmtLine + dblUTGSTAmtLine + dblCCESSAmtLine
                        rsCustItemMst = New ClsResultSetDB
                        rsItemMst = New ClsResultSetDB
                        rsItemMst.GetResult("SELECT Description FROM Item_Mst WHERE Item_Code ='" & Trim(lstrItemCode) & "' and UNIT_CODE = '" & gstrUnitId & "'", ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                        rsCustItemMst.GetResult("SELECT Drg_desc FROM CustItem_Mst WHERE Account_code ='" & Trim(txtCustCode.Text) & "' and UNIT_CODE = '" & gstrUnitId & "'  and Cust_DrgNo='" & lstrItemDrgno & "'and Item_code ='" & lstrItemCode & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        If UCase(Trim(lstrItemDelete)) <> "D" Then
                            strSalesDtl = Trim(strSalesDtl) & "INSERT INTO sales_Dtl(UNIT_CODE,Cust_Ref,Packing_Type,BinQuantity,Amendment_No,Location_Code,Doc_No,Suffix,Item_Code,Sales_Quantity,"
                            strSalesDtl = strSalesDtl & "From_Box,To_Box,Rate,Sales_Tax,Excise_Tax,Packing,ItemPacking_Amount,Others,Cust_Mtrl,"
                            strSalesDtl = strSalesDtl & "Year,Cust_Item_Code,Cust_Item_Desc,Tool_Cost,Measure_Code,Excise_type,SalesTax_type,CVD_type,SAD_type,Basic_Amount,Accessible_amount,CVD_Amount,SVD_amount,"
                            strSalesDtl = strSalesDtl & "Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,Excise_per,CVD_per,SVD_per,CustMtrl_Amount,ToolCost_Amount,TotalExciseAmount,"
                            strSalesDtl = strSalesDtl & "HSNSACCODE,ISHSNORSAC,CGSTTXRT_TYPE,CGST_PERCENT,CGST_AMT,SGSTTXRT_TYPE,SGST_PERCENT,"
                            strSalesDtl = strSalesDtl & "SGST_AMT,IGSTTXRT_TYPE,IGST_PERCENT,IGST_AMT,UTGSTTXRT_TYPE,UTGST_PERCENT,UTGST_AMT,"
                            strSalesDtl = strSalesDtl & "COMPENSATION_CESS_TYPE,COMPENSATION_CESS_PERCENT,COMPENSATION_CESS_AMT,Discount_perc,Discount_amt,ITEM_VALUE)"
                            strSalesDtl = strSalesDtl & "values ('" & gstrUnitId & "','" & Trim(txtRefNo.Text) & "','" & Trim(strPackingCode) & "' ," & dblBinQty & ",'" & Trim(txtAmendNo.Text) & "','" & Trim(txtLocationCode.Text) & "','"
                            strSalesDtl = strSalesDtl & Trim(txtChallanNo.Text) & "','','" & Trim(lstrItemCode) & "','" & Val(CStr(lintItemQuantity)) & "','"
                            strSalesDtl = strSalesDtl & Val(CStr(ldblItemFromBox)) & "','" & Val(CStr(ldblItemToBox)) & "'," & Val(CStr(ldblItemRate)) & "," & Trim(lblSaltax_Per.Text) & ","

                            dblItemPacking_Amount = CalculatePackingValue(lintLoopCounter, blnPackingRoundoff)
                            If blnISExciseRoundOff Then
                                '10736222
                                dblExcise_Amount = System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal + dblItemPacking_Amount, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff))
                                strSalesDtl = strSalesDtl & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal + dblItemPacking_Amount, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff))
                            Else
                                '10736222
                                dblExcise_Amount = CalculateExciseValue(lintLoopCounter, TempAccessibleVal + dblItemPacking_Amount, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff)
                                strSalesDtl = strSalesDtl & CalculateExciseValue(lintLoopCounter, TempAccessibleVal + dblItemPacking_Amount, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff)
                            End If
                            strSalesDtl = strSalesDtl & "," & Val(CStr(ldblItemPacking)) & "," & Val(CStr(dblItemPacking_Amount)) & "," & Val(CStr(ldblItemOthers)) & "," & Val(CStr(ldblItemCustMtrl)) & ",'"
                            strSalesDtl = strSalesDtl & Trim(CStr(Year(dtpDateDesc.Value))) & "','" & Trim(lstrItemDrgno) & "','" & IIf((Len(Trim(rsCustItemMst.GetValue("Drg_Desc"))) <= 0 Or Trim(CStr(rsCustItemMst.GetValue("Drg_Desc") = "Unknown"))), Trim(rsItemMst.GetValue("Description")), Trim(rsCustItemMst.GetValue("Drg_Desc"))) & "',"
                            'If UCase(CmbInvType.Text) = "NORMAL INVOICE" Or UCase(CmbInvType.Text) = "EXPORT INVOICE" Then
                            If UCase(CmbInvType.Text) = "NORMAL INVOICE" Or UCase(CmbInvType.Text) = "EXPORT INVOICE" Or UCase(CmbInvType.Text) = "TRANSFER INVOICE" Then
                                If UCase(CmbInvSubType.Text) <> "SCRAP" Then
                                    strSalesDtl = strSalesDtl & ldblItemToolCost & ",'','"
                                Else
                                    strSalesDtl = strSalesDtl & ldblItemToolCost & ",'','"
                                End If
                            Else
                                strSalesDtl = strSalesDtl & ldblItemToolCost & ",'','"
                            End If
                            strSalesDtl = strSalesDtl & lstrItemExciseCode & "','" & Trim(txtSaleTaxType.Text) & "','" & lstrItemCVDCode & "','" & lstrItemSADCode & "',"
                            strSalesDtl = strSalesDtl & CalculateBasicValue(lintLoopCounter, blnISBasicRoundOff) & ","
                            strSalesDtl = strSalesDtl & TempAccessibleVal & ","
                            If blnISExciseRoundOff Then
                                strSalesDtl = strSalesDtl & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_CVD, blnEOUFlag, blnISExciseRoundOff))
                                strSalesDtl = strSalesDtl & "," & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_SAD, blnEOUFlag, blnISExciseRoundOff))
                            Else
                                strSalesDtl = strSalesDtl & (CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_CVD, blnEOUFlag, blnISExciseRoundOff))
                                strSalesDtl = strSalesDtl & "," & (CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_SAD, blnEOUFlag, blnISExciseRoundOff))
                            End If
                            strSalesDtl = strSalesDtl & ",GetDate(),'"
                            strSalesDtl = strSalesDtl & Trim(mP_User) & "', GetDate(),'" & Trim(mP_User) & "'," & GetTaxRate(lstrItemExciseCode, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='EXC' ") & "," & GetTaxRate(lstrItemCVDCode, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='CVD'") & "," & GetTaxRate(lstrItemSADCode, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='SAD'") & "," & System.Math.Round(Val(CStr(lintItemQuantity * ldblItemCustMtrl)), 2) & "," & System.Math.Round(Val(CStr(lintItemQuantity * ldblItemToolCost)), 2) & ","
                            If blnISExciseRoundOff Then
                                strSalesDtl = strSalesDtl & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal + dblItemPacking_Amount, enumExciseType.RETURN_ALLExcise, blnEOUFlag, blnISExciseRoundOff)) & ",'" & strHSNSACCode & "','" & strHSNSACType & "','" & CGSTType & "'," & dblCGSTPercentLine & "," & dblCGSTAmtLine & ",'" & SGSTType & "'," & dblSGSTPercentLine & "," & dblSGSTAmtLine & ",'" & IGSTType & "'," & dblIGSTPercentLine & "," & dblIGSTAmtLine & ",'" & UTGSTType & "'," & dblUTGSTPercentLine & "," & dblUTGSTAmtLine & ",'" & COMPCESSType & "'," & dblCCESSPercentLine & "," & dblCCESSAmtLine & ",0,0," & dblItemTotalLine & ")"
                            Else
                                strSalesDtl = strSalesDtl & (CalculateExciseValue(lintLoopCounter, TempAccessibleVal + dblItemPacking_Amount, enumExciseType.RETURN_ALLExcise, blnEOUFlag, blnISExciseRoundOff)) & ",'" & strHSNSACCode & "','" & strHSNSACType & "','" & CGSTType & "'," & dblCGSTPercentLine & "," & dblCGSTAmtLine & ",'" & SGSTType & "'," & dblSGSTPercentLine & "," & dblSGSTAmtLine & ",'" & IGSTType & "'," & dblIGSTPercentLine & "," & dblIGSTAmtLine & ",'" & UTGSTType & "'," & dblUTGSTPercentLine & "," & dblUTGSTAmtLine & ",'" & COMPCESSType & "'," & dblCCESSPercentLine & "," & dblCCESSAmtLine & ",0,0," & dblItemTotalLine & " )"
                            End If

                            '10736222
                            strsql = "select dbo.UDF_ISCT2INVOICE( '" & gstrUnitId & "','" & txtCustCode.Text.Trim & "','" & CmbInvType.Text.Trim & "','" & CmbInvSubType.Text.Trim & "','" & txtRefNo.Text.Trim & "')"
                            If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strsql)) = True Then
                                blnIsCt2 = True
                                strSqlct2qry = "INSERT INTO TMP_CT2_INVOICE_KNOCKOFF ([UNIT_CODE],[CUST_CODE],[SONO],[AMENDMENT_NO],[TMP_INVOICE_NO],[ITEM_CODE],[CUST_DRG_NO],[CURRENCY_CODE],[QTY],[RATE],[TOOL_COST],[EXCISE_TAX],[EXCISE_AMOUNT],[ECESS_TYPE],[SECESS_TYPE],[IP_ADDRESS]) "
                                strSqlct2qry = strSqlct2qry + " Values('" & gstrUnitId & "','" & txtCustCode.Text.Trim & "','" & txtRefNo.Text.Trim & "','" & txtAmendNo.Text.Trim & "','" & Me.txtChallanNo.Text.Trim & "',"
                                strSqlct2qry = strSqlct2qry + "'" & lstrItemCode.Trim & "','" & lstrItemDrgno.Trim & "','" & lblCurrencyDes.Text.Trim & "'," & Val(CStr(lintItemQuantity)) & "," & Val(CStr(ldblItemRate)) & "," & Val(ldblItemToolCost) & ",'" & lstrItemExciseCode.Trim & "'," & dblExcise_Amount & ",'" & txtECESS.Text.Trim & "','" & txtSECESS.Text.Trim & "','" & gstrIpaddressWinSck & "' ) "
                                SqlConnectionclass.ExecuteNonQuery(strSqlct2qry)
                            End If
                            '10736222

                        End If
                        rsItemMst.ResultSetClose()
                        rsCustItemMst.ResultSetClose()
                    Next
                End With
            Case "EDIT"
                strSalesChallan = ""
                strSalesChallan = "UPDATE SalesChallan_Dtl SET Insurance = " & System.Math.Round(Val(ctlInsurance.Text))
                If blnISSalesTaxRoundOff Then
                    strSalesChallan = strSalesChallan & ",Sales_Tax_Amount =" & System.Math.Round(Val(CStr(ldblTotalSaleTaxAmount)))
                    strSalesChallan = strSalesChallan & ",ADDVAT_Type='" & txtAddVAT.Text.Trim() & "',AddVat_Per=" & Val(lblAddVAT.Text) & ",ADDVAT_Amount =" & System.Math.Round(Val(CStr(dblAddVATamount)))
                Else
                    strSalesChallan = strSalesChallan & ",Sales_Tax_Amount =" & Val(CStr(ldblTotalSaleTaxAmount))
                    strSalesChallan = strSalesChallan & ",ADDVAT_Type='" & txtAddVAT.Text.Trim() & "',AddVat_Per=" & Val(lblAddVAT.Text) & ",ADDVAT_Amount =" & Val(CStr(dblAddVATamount))
                End If
                If blnISECESSRoundoff Then
                    strSalesChallan = strSalesChallan & ",ECESS_Amount =" & System.Math.Round(Val(CStr(ldblTotalECESSAmount)))
                    strSalesChallan = strSalesChallan & ",SECESS_Amount =" & System.Math.Round(Val(CStr(ldblTotalSECESSAmount)))
                Else
                    strSalesChallan = strSalesChallan & ",ECESS_Amount =" & Val(CStr(ldblTotalECESSAmount))
                    strSalesChallan = strSalesChallan & ",SECESS_Amount =" & Val(CStr(ldblTotalSECESSAmount))
                End If
                If blnISSurChargeTaxRoundOff Then
                    strSalesChallan = strSalesChallan & ",Surcharge_Sales_Tax_Amount =" & System.Math.Round(Val(CStr(ldblTotalSurchargeTaxAmount)))
                Else
                    strSalesChallan = strSalesChallan & ",Surcharge_Sales_Tax_Amount =" & Val(CStr(ldblTotalSurchargeTaxAmount))
                End If
                If UCase(mstrInvType) <> "JOB" And UCase(mstrInvType) <> "TRF" Then
                    strSalesChallan = strSalesChallan & ", SDTax_Type='" & strSDTType & "', SDTax_Per=" & dblSDT_Per & ", SDTax_Amount = " & dblSDT_Amt
                End If
                strSalesChallan = strSalesChallan & ",Frieght_Amount=" & System.Math.Round(Val(txtFreight.Text))
                strSalesChallan = strSalesChallan & ",SalesTax_Type='" & Trim(txtSaleTaxType.Text) & "'"
                strSalesChallan = strSalesChallan & ",total_amount=" & ldblTotalInvoiceValue
                strSalesChallan = strSalesChallan & ",Packing_amount=" & dblTotalPacking_Amount
                strSalesChallan = strSalesChallan & ",Surcharge_salesTaxType='" & Trim(txtSurchargeTaxType.Text) & "'"
                strSalesChallan = strSalesChallan & ",SalesTax_Per=" & Val(lblSaltax_Per.Text)
                strSalesChallan = strSalesChallan & ",Surcharge_SalesTax_Per=" & Val(lblSurcharge_Per.Text)
                strSalesChallan = strSalesChallan & ",PerValue=" & ctlPerValue.Text & ",Remarks = '" & txtRemarks.Text & "' "
                strSalesChallan = strSalesChallan & ",SRVDINO = '" & Trim(txtSRVDI.Text) & "',"
                strSalesChallan = strSalesChallan & " SRVLocation = '" & Trim(txtSRVLoc.Text)
                strSalesChallan = strSalesChallan & "',USLOC = '" & Trim(txtUsLoc.Text) & "',"
                strSalesChallan = strSalesChallan & " schTime = '" & Trim(txtSchTime.Text) & "'"
                strSalesChallan = strSalesChallan & ",ECESS_Type = '" & Trim(txtECESS.Text) & "',"
                strSalesChallan = strSalesChallan & "ECESS_Per = " & Val(lblECESS_Per.Text) & ","
                strSalesChallan = strSalesChallan & "PAYMENT_TERMS = '" & Trim(lblCreditTerm.Text) & "',"
                strSalesChallan = strSalesChallan & "SECESS_Type = '" & Trim(txtSECESS.Text) & "',"
                strSalesChallan = strSalesChallan & "SECESS_Per = " & Val(lblSECESS_Per.Text) & ","
                strSalesChallan = strSalesChallan & "TotalInvoiceAmtRoundOff_diff = " & ldblTotalInvoiceValueRoundOff
                strSalesChallan = strSalesChallan & ",CGST_TOTAL_AMT=" & dblCGSTAMT & " , SGST_TOTAL_AMT=" & dblSGSTAMT & ""
                strSalesChallan = strSalesChallan & ",IGST_TOTAL_AMT=" & dblIGSTAMT & " , UTGST_TOTAL_AMT=" & dblUTGSTAMT & ""
                strSalesChallan = strSalesChallan & ",CCESS_TOTAL_AMT=" & dblCOMPCESSAMT & ""
                strSalesChallan = strSalesChallan & " WHERE  UNIT_CODE = '" & gstrUnitId & "' and Location_Code ='" & Trim(txtLocationCode.Text) & "'"
                strSalesChallan = strSalesChallan & " and Doc_No ='" & Val(txtChallanNo.Text) & "'"
                strSalesDtl = ""
                strSalesDtlDelete = ""
                With SpChEntry
                    For lintLoopCounter = 1 To .MaxRows
                        .Row = lintLoopCounter
                        .Col = 1
                        lstrItemCode = Trim(.Text)
                        .Col = 3
                        ldblItemRate = Val(.Text) / CDbl(ctlPerValue.Text)
                        .Row = lintLoopCounter
                        .Col = 5
                        lintItemQuantity = Val(.Text)
                        .Col = 6
                        strPackingCode = Trim(.Text)
                        rsPacking_Tax = New ClsResultSetDB
                        rsPacking_Tax.GetResult("Select Txrt_Rate_no,TxRt_Percentage from Gen_TaxRate where  UNIT_CODE = '" & gstrUnitId & "' and Tx_TaxeID = 'PKT' and Txrt_Rate_no = '" & Trim(strPackingCode) & "'")
                        If rsPacking_Tax.GetNoRows > 0 Then
                            ldblItemPacking = rsPacking_Tax.GetValue("TxRt_Percentage")
                        End If
                        rsPacking_Tax.ResultSetClose()
                        .Col = 2
                        lstrItemDrgno = Trim(.Text)
                        .Col = 14
                        lstrItemDelete = Trim(.Text)
                        .Col = 15
                        ldblItemToolCost = Val(.Text) / CDbl(ctlPerValue.Text)
                        .Col = 7
                        lstrItemExciseCode = Trim(.Text)
                        .Col = 8
                        lstrItemCVDCode = Trim(.Text)
                        .Col = 9
                        lstrItemSADCode = Trim(.Text)
                        .Col = 15
                        ldblItemToolCost = Val(.Text)
                        .Col = 22
                        dblBinQty = Val(.Text)
                        If dblBinQty <= 0 Then
                            MsgBox("Bin Quantity can't be zero.", MsgBoxStyle.Information, "eMpro")
                            SaveDataGST = False
                            Exit Function
                        End If
                        .Col = 11
                        ldblItemFromBox = Val(.Text)
                        .Col = 12
                        ldblItemToBox = Val(.Text)
                        If Val(CStr(ldblItemCustMtrl)) > 0 Then
                            strQry = ""
                            strQry = "Select Cust_Mtrl from Cust_ord_dtl WHERE "
                            strQry = strQry & "Account_Code ='" & txtCustCode.Text & "'and Cust_ref ='"
                            strQry = strQry & txtRefNo.Text & "' and Amendment_No = '" & Trim(txtAmendNo.Text) & "'and "
                            strQry = strQry & " Active_flag ='A' "
                            strQry = strQry & " and Cust_DrgNo = '" & Trim(lstrItemDrgno) & "'"
                            strQry = strQry & " and Item_Code ='" & Trim(lstrItemCode) & "' and UNIT_CODE = '" & gstrUnitId & "'"
                            rsCustOrdDtl = New ClsResultSetDB
                            rsCustOrdDtl.GetResult(strQry)
                            If rsCustOrdDtl.GetNoRows > 0 Then
                                dblCustMtrl_SO = rsCustOrdDtl.GetValue("Cust_Mtrl")
                            End If
                            If Val(CStr(dblCustMtrl_SO)) = 0 Then
                                If Val(CStr(ldblItemCustMtrl)) > 0 Then
                                    ldblItemCustMtrl = 0
                                End If
                            End If
                            rsCustOrdDtl.ResultSetClose()
                            rsCustOrdDtl = Nothing
                        End If
                        TempAccessibleVal = CalculateAccessibleValue(lintLoopCounter, ldblNetInsurenceValue, blnISInsExcisable)
                        .Col = HSN_SAC_CODE
                        strHSNSACCode = .Text
                        .Col = IS_HSN_SAC
                        strHSNSACType = .Text
                        .Col = CGST_TYPE
                        CGSTType = .Text
                        .Col = SGST_TYPE
                        SGSTType = .Text
                        .Col = IGST_TYPE
                        IGSTType = .Text
                        .Col = UTGST_TYPE
                        UTGSTType = .Text
                        .Col = COMP_CESS_TYPE
                        COMPCESSType = .Text
                        dtGSTTaxPercent = New DataTable()
                        dtGSTTaxPercent = GetGSTTaxesPercentage(CGSTType, SGSTType, IGSTType, UTGSTType, COMPCESSType)
                        If dtGSTTaxPercent IsNot Nothing AndAlso dtGSTTaxPercent.Rows.Count > 0 Then
                            dblCGSTPercentLine = dtGSTTaxPercent.Rows(0)("CGST_PERCENT")
                            dblSGSTPercentLine = dtGSTTaxPercent.Rows(0)("SGST_PERCENT")
                            dblIGSTPercentLine = dtGSTTaxPercent.Rows(0)("IGST_PERCENT")
                            dblUTGSTPercentLine = dtGSTTaxPercent.Rows(0)("UTGST_PERCENT")
                            dblCCESSPercentLine = dtGSTTaxPercent.Rows(0)("COMPENSATION_CESS_PERCENT")
                            If blnGSTTAXroundoff Then
                                dblCGSTAmtLine = CalculateGSTTaxes(TempAccessibleVal, dblCGSTPercentLine)
                                dblSGSTAmtLine = CalculateGSTTaxes(TempAccessibleVal, dblSGSTPercentLine)
                                dblIGSTAmtLine = CalculateGSTTaxes(TempAccessibleVal, dblIGSTPercentLine)
                                dblUTGSTAmtLine = CalculateGSTTaxes(TempAccessibleVal, dblUTGSTPercentLine)
                                dblCCESSAmtLine = CalculateGSTTaxes(TempAccessibleVal, dblCCESSPercentLine)
                            Else
                                dblCGSTAmtLine = System.Math.Round(CalculateGSTTaxes(TempAccessibleVal, dblCGSTPercentLine), intGSTTAXroundoff_decimal)
                                dblSGSTAmtLine = System.Math.Round(CalculateGSTTaxes(TempAccessibleVal, dblSGSTPercentLine), intGSTTAXroundoff_decimal)
                                dblUTGSTAmtLine = System.Math.Round(CalculateGSTTaxes(TempAccessibleVal, dblUTGSTPercentLine), intGSTTAXroundoff_decimal)
                                dblIGSTAmtLine = System.Math.Round(CalculateGSTTaxes(TempAccessibleVal, dblIGSTPercentLine), intGSTTAXroundoff_decimal)
                                dblCCESSAmtLine = System.Math.Round(CalculateGSTTaxes(TempAccessibleVal, dblCCESSPercentLine), intGSTTAXroundoff_decimal)
                            End If
                        End If
                        dblItemTotalLine = TempAccessibleVal + dblCGSTAmtLine + dblSGSTAmtLine + dblIGSTAmtLine + dblUTGSTAmtLine + dblCCESSAmtLine

                        If UCase(lstrItemDelete) <> "D" Then
                            If UCase(lstrItemDelete) <> "A" Then
                                strSalesDtl = Trim(strSalesDtl) & "UPDATE Sales_dtl SET Sales_Quantity ='" & Val(CStr(lintItemQuantity)) & "',BinQuantity=" & dblBinQty & " ,Sales_Tax =" & Trim(lblSaltax_Per.Text) & ","
                                strSalesDtl = Trim(strSalesDtl) & " TOOL_COST = " & ldblItemToolCost & ","
                                strSalesDtl = Trim(strSalesDtl) & "CustMtrl_Amount= " & Val(CStr(lintItemQuantity * ldblItemCustMtrl)) & ",ToolCost_Amount=" & Val(CStr(lintItemQuantity * ldblItemToolCost))
                                dblItemPacking_Amount = CalculatePackingValue(lintLoopCounter, blnPackingRoundoff)
                                If blnISExciseRoundOff Then
                                    '10736222
                                    dblExcise_Amount = System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal + dblItemPacking_Amount, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff))
                                    strSalesDtl = Trim(strSalesDtl) & ",Excise_Tax=" & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal + dblItemPacking_Amount, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff))
                                Else
                                    dblExcise_Amount = CalculateExciseValue(lintLoopCounter, TempAccessibleVal + dblItemPacking_Amount, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff)
                                    strSalesDtl = Trim(strSalesDtl) & ",Excise_Tax=" & CalculateExciseValue(lintLoopCounter, TempAccessibleVal + dblItemPacking_Amount, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff)
                                End If
                                If blnISExciseRoundOff Then
                                    strSalesDtl = Trim(strSalesDtl) & ",TotalExciseAmount =" & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal + dblItemPacking_Amount, enumExciseType.RETURN_ALLExcise, blnEOUFlag, blnISExciseRoundOff))
                                Else
                                    strSalesDtl = Trim(strSalesDtl) & ",TotalExciseAmount =" & CalculateExciseValue(lintLoopCounter, TempAccessibleVal + dblItemPacking_Amount, enumExciseType.RETURN_ALLExcise, blnEOUFlag, blnISExciseRoundOff)
                                End If
                                strSalesDtl = Trim(strSalesDtl) & ",Excise_type='" & lstrItemExciseCode & "',SalesTax_type='" & Trim(txtSaleTaxType.Text) & "'"
                                strSalesDtl = Trim(strSalesDtl) & ", Packing=" & Val(CStr(ldblItemPacking)) & ",ItemPacking_Amount=" & Val(CStr(dblItemPacking_Amount)) & ""
                                strSalesDtl = Trim(strSalesDtl) & ",CVD_type='" & Trim(lstrItemCVDCode) & "',SAD_type='" & Trim(lstrItemSADCode) & "',Basic_Amount=" & CalculateBasicValue(lintLoopCounter, blnISBasicRoundOff)
                                strSalesDtl = Trim(strSalesDtl) & ",Accessible_amount=" & Val(CStr(TempAccessibleVal))
                                If blnISExciseRoundOff Then
                                    strSalesDtl = Trim(strSalesDtl) & ",CVD_Amount=" & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_CVD, blnEOUFlag, blnISExciseRoundOff)) & ",SVD_amount=" & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_SAD, blnEOUFlag, blnISExciseRoundOff))
                                Else
                                    strSalesDtl = Trim(strSalesDtl) & ",CVD_Amount=" & CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_CVD, blnEOUFlag, blnISExciseRoundOff) & ",SVD_amount=" & CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_SAD, blnEOUFlag, blnISExciseRoundOff)
                                End If
                                strSalesDtl = Trim(strSalesDtl) & ",Excise_per=" & GetTaxRate(lstrItemExciseCode, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='EXC'")
                                strSalesDtl = Trim(strSalesDtl) & ",CVD_per=" & GetTaxRate(lstrItemCVDCode, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='CVD'")
                                strSalesDtl = Trim(strSalesDtl) & ",SVD_per=" & GetTaxRate(lstrItemSADCode, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='SAD'")
                                strSalesDtl = Trim(strSalesDtl) & ",Rate=" & ldblItemRate & ",FROM_BOX = " & ldblItemFromBox & ",To_box = " & ldblItemToBox
                                strSalesDtl = Trim(strSalesDtl) & ",CGST_AMT=" & dblCGSTAmtLine & ",SGST_AMT=" & dblSGSTAmtLine & ""
                                strSalesDtl = Trim(strSalesDtl) & ",IGST_AMT=" & dblIGSTAmtLine & ",UTGST_AMT=" & dblUTGSTAmtLine & ""
                                strSalesDtl = Trim(strSalesDtl) & ",COMPENSATION_CESS_AMT=" & dblCCESSAmtLine & ""
                                strSalesDtl = Trim(strSalesDtl) & ",ITEM_VALUE=" & dblItemTotalLine & ""
                                strSalesDtl = Trim(strSalesDtl) & " WHERE  UNIT_CODE = '" & gstrUnitId & "' and Location_Code ='" & Trim(txtLocationCode.Text) & "'"
                                strSalesDtl = Trim(strSalesDtl) & " and Doc_No =" & Val(txtChallanNo.Text) & " and Cust_Item_Code='"
                                strSalesDtl = Trim(strSalesDtl) & Trim(lstrItemDrgno) & "'" & vbCrLf
                            ElseIf UCase(lstrItemDelete) = "A" Then
                                strSalesDtl = strSalesDtl & vbCrLf & InsertinSalesDtlinEditModeGST(lintLoopCounter)
                            End If
                        Else
                            strSalesDtlDelete = Trim(strSalesDtlDelete) & "DELETE Sales_dtl "
                            strSalesDtlDelete = Trim(strSalesDtlDelete) & " WHERE  UNIT_CODE = '" & gstrUnitId & "'  and Location_Code ='" & Trim(txtLocationCode.Text) & "'"
                            strSalesDtlDelete = Trim(strSalesDtlDelete) & " and Doc_No =" & Val(txtChallanNo.Text) & " and Cust_Item_Code='"
                            strSalesDtlDelete = Trim(strSalesDtlDelete) & Trim(lstrItemDrgno) & "'" & vbCrLf
                        End If
                        '10736222
                        strsql = "select dbo.UDF_ISCT2INVOICE( '" & gstrUnitId & "','" & txtCustCode.Text.Trim & "','" & CmbInvType.Text.Trim & "','" & CmbInvSubType.Text.Trim & "','" & txtRefNo.Text.Trim & "')"
                        If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strsql)) = True Then
                            blnIsCt2 = True
                            strSqlct2qry = "INSERT INTO TMP_CT2_INVOICE_KNOCKOFF ([UNIT_CODE],[CUST_CODE],[SONO],[AMENDMENT_NO],[TMP_INVOICE_NO],[ITEM_CODE],[CUST_DRG_NO],[CURRENCY_CODE],[QTY],[RATE],[TOOL_COST],[EXCISE_TAX],[EXCISE_AMOUNT],[ECESS_TYPE],[SECESS_TYPE],[IP_ADDRESS]) "
                            strSqlct2qry = strSqlct2qry + " Values('" & gstrUnitId & "','" & txtCustCode.Text.Trim & "','" & txtRefNo.Text.Trim & "','" & txtAmendNo.Text.Trim & "','" & Me.txtChallanNo.Text.Trim & "',"
                            strSqlct2qry = strSqlct2qry + "'" & lstrItemCode.Trim & "','" & lstrItemDrgno.Trim & "','" & lblCurrencyDes.Text.Trim & "'," & Val(CStr(lintItemQuantity)) & "," & Val(CStr(ldblItemRate)) & "," & Val(ldblItemToolCost) & ",'" & lstrItemExciseCode.Trim & "'," & dblExcise_Amount & ",'" & txtECESS.Text.Trim & "','" & txtSECESS.Text.Trim & "','" & gstrIpaddressWinSck & "' ) "
                            SqlConnectionclass.ExecuteNonQuery(strSqlct2qry)
                        End If
                        '10736222

                    Next
                End With
        End Select

        If blnIsCt2 = True Then
            '10736222
            Dim objValidateCmd As New ADODB.Command

            With objValidateCmd
                .ActiveConnection = mP_Connection
                .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                .CommandText = "USP_VALIDATE_CT2_INVOICE_KNOCKOFF"
                .CommandTimeout = 0
                .Parameters.Append(.CreateParameter("@MODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, IIf(CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, "A", "E")))
                .Parameters.Append(.CreateParameter("@UNIT_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, gstrUnitId))
                .Parameters.Append(.CreateParameter("@INVOICE_NO", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, , txtChallanNo.Text.Trim))
                .Parameters.Append(.CreateParameter("@IP_ADDRESS", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, gstrIpaddressWinSck))
                .Parameters.Append(.CreateParameter("@ERRMSG", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInputOutput, 8000, ""))
                .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End With

            If objValidateCmd.Parameters(objValidateCmd.Parameters.Count - 1).Value.ToString().Trim.Length <> 0 Then
                MsgBox("Unable To Save CT2 Invoice Knock Off Details." & vbCr & objValidateCmd.Parameters(objValidateCmd.Parameters.Count - 1).Value.ToString(), MsgBoxStyle.Information, ResolveResString(100))
                objValidateCmd = Nothing
                SaveDataGST = False
                Exit Function
            End If
            objValidateCmd = Nothing
            '10736222
        End If

        With mP_Connection
            ResetDatabaseConnection()
            .BeginTrans()
            .Execute(strSalesChallan, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            If Len(Trim(strupSalechallan)) > 0 Then
                .Execute(strupSalechallan, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End If
            .Execute(strSalesDtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            If Len(Trim(mstrUpdDispatchSql)) > 0 Then
                .Execute(mstrUpdDispatchSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End If
            If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                If Len(Trim(strSalesDtlDelete)) > 0 Then
                    .Execute(strSalesDtlDelete, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                End If
            End If
            If blnIsCt2 = True Then
                '10736222
                Dim objCmd As New ADODB.Command

                With objCmd
                    .ActiveConnection = mP_Connection
                    .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                    .CommandText = "USP_SAVE_CT2_INVOICE_KNOCKOFFDTL"
                    .CommandTimeout = 0
                    .Parameters.Append(.CreateParameter("@MODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, IIf(CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, "A", "E")))
                    .Parameters.Append(.CreateParameter("@UNIT_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, gstrUnitId))
                    .Parameters.Append(.CreateParameter("@INVOICE_NO", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, , txtChallanNo.Text.Trim))
                    .Parameters.Append(.CreateParameter("@IP_ADDRESS", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, gstrIpaddressWinSck))
                    .Parameters.Append(.CreateParameter("@ERRMSG", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInputOutput, 8000, ""))
                    .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                End With

                If objCmd.Parameters(objCmd.Parameters.Count - 1).Value.ToString().Trim.Length <> 0 Then
                    MsgBox("Unable To Save CT2 Invoice Knock Off Details.", MsgBoxStyle.Information, ResolveResString(100))
                    objCmd = Nothing
                    mP_Connection.RollbackTrans()
                    SaveDataGST = False
                    Exit Function
                End If
                objCmd = Nothing
                '10736222
            End If

            .CommitTrans()
        End With
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        mP_Connection.RollbackTrans()
        SaveDataGST = False
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

    Public Function InsertinSalesDtlinEditModeGST(ByRef pintRow As Short) As String
        Dim ldblTotalBasicValue As Double
        Dim ldblTotalAccessibleValue As Double
        Dim ldblTempAccessibleVal As Double
        Dim ldblTotalExciseValue As Double
        Dim ldblTotalSaleTaxAmount As Double
        Dim ldblTotalSurchargeTaxAmount As Double
        Dim ldblNetInsurenceValue As Double
        Dim ldblTotalInvoiceValue As Double
        Dim ldblTotalOthersValues As Double
        Dim rsParameterData As ClsResultSetDB
        Dim strParamQuery As String
        Dim rsStockLocation As ClsResultSetDB
        Dim rsCustItemMst As ClsResultSetDB
        Dim rsItemMst As ClsResultSetDB
        Dim strSalesDtl As String
        Dim strInvDes As String
        Dim strInvSubTypeDes As String
        Dim lintItemQuantity As Double
        Dim lstrItemDrgno As String
        Dim lstrItemCode As String
        Dim ldblItemRate As Double
        Dim ldblItemCustMtrl As Double
        Dim ldblItemPacking As Double
        Dim ldblItemOthers As Double
        Dim ldblItemFromBox As Double
        Dim ldblItemToBox As Double
        Dim lstrItemDelete As String
        Dim lintItemPresQty As Double
        Dim lstrItemExciseCode As String
        Dim lstrItemCVDCode As String
        Dim lstrItemSADCode As String
        Dim ldblItemToolCost As Double
        Dim TempAccessibleVal As Double
        Dim ldblTotalCustMatrlValue As Double
        Dim blnISInsExcisable As Boolean
        Dim blnEOUFlag As Boolean
        Dim blnISExciseRoundOff As Boolean
        Dim blnISSalesTaxRoundOff As Boolean
        Dim blnISSurChargeTaxRoundOff As Boolean
        Dim blnAddCustMatrl As Boolean
        Dim blnISBasicRoundOff As Boolean
        Dim CGSTType As String = String.Empty
        Dim SGSTType As String = String.Empty
        Dim IGSTType As String = String.Empty
        Dim UTGSTType As String = String.Empty
        Dim COMPCESSType As String = String.Empty
        Dim dtGSTTaxPercent As New DataTable
        Dim strHSNSACCode As String = String.Empty
        Dim strHSNSACType As String = String.Empty
        Dim dblCGSTPercentLine As Double = 0
        Dim dblSGSTPercentLine As Double = 0
        Dim dblIGSTPercentLine As Double = 0
        Dim dblUTGSTPercentLine As Double = 0
        Dim dblCCESSPercentLine As Double = 0
        Dim dblCGSTAmtLine As Double = 0
        Dim dblSGSTAmtLine As Double = 0
        Dim dblIGSTAmtLine As Double = 0
        Dim dblUTGSTAmtLine As Double = 0
        Dim dblCCESSAmtLine As Double = 0
        Dim dblItemTotalLine As Double = 0
        Dim blnGSTTAXroundoff As Boolean
        Dim intGSTTAXroundoff_decimal As Short
        On Error GoTo ErrHandler
        strParamQuery = "SELECT InsExc_Excise,CustSupp_Inc,EOU_Flag,Basic_Roundoff,SalesTax_Roundoff,Excise_Roundoff,SST_Roundoff,salesTax_Roundoff_decimal,ECESSRoundoff_decimal,GSTTAX_ROUNDOFF_DECIMAL,GSTTAX_ROUNDOFF FROM Sales_Parameter where UNIT_CODE = '" & gstrUnitId & "'"
        rsParameterData = New ClsResultSetDB
        rsParameterData.GetResult(strParamQuery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        blnISBasicRoundOff = rsParameterData.GetValue("Basic_Roundoff")
        blnGSTTAXroundoff = rsParameterData.GetValue("GSTTAX_ROUNDOFF")
        intGSTTAXroundoff_decimal = rsParameterData.GetValue("GSTTAX_ROUNDOFF_DECIMAL")
        rsParameterData.ResultSetClose()
        rsStockLocation = New ClsResultSetDB
        rsStockLocation.GetResult("Select Description,Sub_type_Description from SaleConf a,SalesChallan_dtl b where a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = '" & gstrUnitId & "' AND b.Invoice_type = a.Invoice_type and b.Sub_Category = a.Sub_type and a.Location_Code =b.Location_code and b.Location_code ='" & txtLocationCode.Text & "' and b.Doc_No = " & txtChallanNo.Text & " and (fin_start_date <= getdate() and fin_end_date >= getdate())")
        strInvDes = rsStockLocation.GetValue("Description")
        strInvSubTypeDes = rsStockLocation.GetValue("Sub_type_Description")
        rsStockLocation.ResultSetClose()
        strSalesDtl = ""
        With SpChEntry
            .Row = pintRow
            .Col = 1
            lstrItemCode = Trim(.Text)
            .Col = 2
            lstrItemDrgno = Trim(.Text)
            .Col = 3
            ldblItemRate = Val(.Text) / CDbl(ctlPerValue.Text)
            .Col = 4
            ldblItemCustMtrl = Val(.Text) / CDbl(ctlPerValue.Text)
            .Col = 5
            lintItemQuantity = Val(.Text)
            .Col = 6
            ldblItemPacking = Val(.Text)
            .Col = 7
            lstrItemExciseCode = Trim(.Text)
            .Col = 8
            lstrItemCVDCode = Trim(.Text)
            .Col = 9
            lstrItemSADCode = Trim(.Text)
            .Col = 10
            ldblItemOthers = Val(.Text) / CDbl(ctlPerValue.Text) * lintItemQuantity
            .Col = 11
            ldblItemFromBox = Val(.Text)
            .Col = 12
            ldblItemToBox = Val(.Text)
            .Col = 14
            lstrItemDelete = Trim(.Text)
            .Col = 15
            ldblItemToolCost = Val(.Text) / CDbl(ctlPerValue.Text)

            TempAccessibleVal = CalculateAccessibleValue(pintRow, ldblNetInsurenceValue, blnISInsExcisable)
            .Col = HSN_SAC_CODE
            strHSNSACCode = .Text
            .Col = IS_HSN_SAC
            strHSNSACType = .Text
            .Col = CGST_TYPE
            CGSTType = .Text
            .Col = SGST_TYPE
            SGSTType = .Text
            .Col = IGST_TYPE
            IGSTType = .Text
            .Col = UTGST_TYPE
            UTGSTType = .Text
            .Col = COMP_CESS_TYPE
            COMPCESSType = .Text

            dblCGSTPercentLine = 0
            dblCGSTAmtLine = 0
            dblSGSTPercentLine = 0
            dblSGSTAmtLine = 0
            dblIGSTPercentLine = 0
            dblIGSTAmtLine = 0
            dblUTGSTPercentLine = 0
            dblUTGSTAmtLine = 0
            dblCCESSPercentLine = 0
            dblCCESSAmtLine = 0

            dtGSTTaxPercent = New DataTable()
            dtGSTTaxPercent = GetGSTTaxesPercentage(CGSTType, SGSTType, IGSTType, UTGSTType, COMPCESSType)
            If dtGSTTaxPercent IsNot Nothing AndAlso dtGSTTaxPercent.Rows.Count > 0 Then
                dblCGSTPercentLine = dtGSTTaxPercent.Rows(0)("CGST_PERCENT")
                dblSGSTPercentLine = dtGSTTaxPercent.Rows(0)("SGST_PERCENT")
                dblIGSTPercentLine = dtGSTTaxPercent.Rows(0)("IGST_PERCENT")
                dblUTGSTPercentLine = dtGSTTaxPercent.Rows(0)("UTGST_PERCENT")
                dblCCESSPercentLine = dtGSTTaxPercent.Rows(0)("COMPENSATION_CESS_PERCENT")
                If blnGSTTAXroundoff Then
                    dblCGSTAmtLine = CalculateGSTTaxes(TempAccessibleVal, dblCGSTPercentLine)
                    dblSGSTAmtLine = CalculateGSTTaxes(TempAccessibleVal, dblSGSTPercentLine)
                    dblIGSTAmtLine = CalculateGSTTaxes(TempAccessibleVal, dblIGSTPercentLine)
                    dblUTGSTAmtLine = CalculateGSTTaxes(TempAccessibleVal, dblUTGSTPercentLine)
                    dblCCESSAmtLine = CalculateGSTTaxes(TempAccessibleVal, dblCCESSPercentLine)
                Else
                    dblCGSTAmtLine = System.Math.Round(CalculateGSTTaxes(TempAccessibleVal, dblCGSTPercentLine), intGSTTAXroundoff_decimal)
                    dblSGSTAmtLine = System.Math.Round(CalculateGSTTaxes(TempAccessibleVal, dblSGSTPercentLine), intGSTTAXroundoff_decimal)
                    dblUTGSTAmtLine = System.Math.Round(CalculateGSTTaxes(TempAccessibleVal, dblUTGSTPercentLine), intGSTTAXroundoff_decimal)
                    dblIGSTAmtLine = System.Math.Round(CalculateGSTTaxes(TempAccessibleVal, dblIGSTPercentLine), intGSTTAXroundoff_decimal)
                    dblCCESSAmtLine = System.Math.Round(CalculateGSTTaxes(TempAccessibleVal, dblCCESSPercentLine), intGSTTAXroundoff_decimal)
                End If
                dtGSTTaxPercent.Dispose()
            End If
            dblItemTotalLine = TempAccessibleVal + dblCGSTAmtLine + dblSGSTAmtLine + dblIGSTAmtLine + dblUTGSTAmtLine + dblCCESSAmtLine

            rsCustItemMst = New ClsResultSetDB
            rsItemMst = New ClsResultSetDB
            rsItemMst.GetResult("SELECT Description FROM Item_Mst WHERE Item_Code ='" & Trim(lstrItemCode) & "' and UNIT_CODE = '" & gstrUnitId & "'", ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            rsCustItemMst.GetResult("SELECT Drg_desc FROM CustItem_Mst WHERE Account_code ='" & Trim(txtCustCode.Text) & "'and Cust_DrgNo='" & lstrItemDrgno & "'and Item_code ='" & lstrItemCode & "' and UNIT_CODE = '" & gstrUnitId & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If UCase(Trim(lstrItemDelete)) <> "D" Then
                strSalesDtl = Trim(strSalesDtl) & "INSERT INTO sales_Dtl(UNIT_CODE,Cust_Ref,Amendment_No,Location_Code,Doc_No,Suffix,Item_Code,Sales_Quantity,"
                strSalesDtl = strSalesDtl & "From_Box,To_Box,Rate,Sales_Tax,Excise_Tax,Packing,Others,Cust_Mtrl,"
                strSalesDtl = strSalesDtl & "Year,Cust_Item_Code,Cust_Item_Desc,Tool_Cost,Measure_Code,Excise_type,SalesTax_type,CVD_type,SAD_type,Basic_Amount,Accessible_amount,CVD_Amount,SVD_amount,"
                strSalesDtl = strSalesDtl & "Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,Excise_per,CVD_per,SVD_per,CustMtrl_Amount,ToolCost_Amount,"
                strSalesDtl = strSalesDtl & "HSNSACCODE,ISHSNORSAC,CGSTTXRT_TYPE,CGST_PERCENT,CGST_AMT,SGSTTXRT_TYPE,SGST_PERCENT,"
                strSalesDtl = strSalesDtl & "SGST_AMT,IGSTTXRT_TYPE,IGST_PERCENT,IGST_AMT,UTGSTTXRT_TYPE,UTGST_PERCENT,UTGST_AMT,"
                strSalesDtl = strSalesDtl & "COMPENSATION_CESS_TYPE,COMPENSATION_CESS_PERCENT,COMPENSATION_CESS_AMT,Discount_perc,Discount_amt,ITEM_VALUE)"
                strSalesDtl = strSalesDtl & "values ('" & gstrUnitId & "','" & Trim(txtRefNo.Text) & "','" & Trim(txtAmendNo.Text) & "','" & Trim(txtLocationCode.Text) & "','"
                strSalesDtl = strSalesDtl & Trim(txtChallanNo.Text) & "','','" & Trim(lstrItemCode) & "','" & Val(CStr(lintItemQuantity)) & "','"
                strSalesDtl = strSalesDtl & Val(CStr(ldblItemFromBox)) & "','" & Val(CStr(ldblItemToBox)) & "'," & Val(CStr(ldblItemRate)) & "," & Trim(lblSaltax_Per.Text) & ","

                If blnISExciseRoundOff Then
                    strSalesDtl = strSalesDtl & System.Math.Round(CalculateExciseValue(pintRow, TempAccessibleVal, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff))
                Else
                    strSalesDtl = strSalesDtl & CalculateExciseValue(pintRow, TempAccessibleVal, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff)
                End If
                strSalesDtl = strSalesDtl & "," & Val(CStr(ldblItemPacking)) & "," & Val(CStr(ldblItemOthers)) & "," & Val(CStr(ldblItemCustMtrl)) & ",'"
                strSalesDtl = strSalesDtl & Trim(CStr(Year(dtpDateDesc.Value))) & "','" & Trim(lstrItemDrgno) & "','" & IIf((Len(Trim(rsCustItemMst.GetValue("Drg_Desc"))) <= 0 Or Trim(CStr(rsCustItemMst.GetValue("Drg_Desc") = "Unknown"))), Trim(rsItemMst.GetValue("Description")), Trim(rsCustItemMst.GetValue("Drg_Desc"))) & "',"
                'If UCase(strInvDes) = "NORMAL INVOICE" Or UCase(strInvDes) = "EXPORT INVOICE" Then
                If UCase(strInvDes) = "NORMAL INVOICE" Or UCase(strInvDes) = "EXPORT INVOICE" Or UCase(strInvDes) = "TRANSFER INVOICE" Then
                    If UCase(CmbInvSubType.Text) <> "SCRAP" Then
                        strSalesDtl = strSalesDtl & ldblItemToolCost & ",'','"
                    End If
                Else
                    strSalesDtl = strSalesDtl & ldblItemToolCost & ",'','"
                End If
                strSalesDtl = strSalesDtl & lstrItemExciseCode & "','" & Trim(txtSaleTaxType.Text) & "','" & lstrItemCVDCode & "','" & lstrItemSADCode & "',"
                strSalesDtl = strSalesDtl & CalculateBasicValue(pintRow, blnISBasicRoundOff) & ","
                strSalesDtl = strSalesDtl & TempAccessibleVal & ","
                If blnISExciseRoundOff Then
                    strSalesDtl = strSalesDtl & System.Math.Round(CalculateExciseValue(pintRow, TempAccessibleVal, enumExciseType.RETURN_CVD, blnEOUFlag, blnISExciseRoundOff))
                    strSalesDtl = strSalesDtl & "," & System.Math.Round(CalculateExciseValue(pintRow, TempAccessibleVal, enumExciseType.RETURN_SAD, blnEOUFlag, blnISExciseRoundOff))
                Else
                    strSalesDtl = strSalesDtl & (CalculateExciseValue(pintRow, TempAccessibleVal, enumExciseType.RETURN_CVD, blnEOUFlag, blnISExciseRoundOff))
                    strSalesDtl = strSalesDtl & "," & (CalculateExciseValue(pintRow, TempAccessibleVal, enumExciseType.RETURN_SAD, blnEOUFlag, blnISExciseRoundOff))
                End If
                strSalesDtl = strSalesDtl & ",GetDate(),'"
                strSalesDtl = strSalesDtl & Trim(mP_User) & "', GetDate(),'" & Trim(mP_User) & "'," & GetTaxRate(lstrItemExciseCode, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='EXC'") & "," & GetTaxRate(lstrItemCVDCode, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='CVD'") & "," & GetTaxRate(lstrItemSADCode, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='SAD'") & "," & System.Math.Round(Val(CStr(lintItemQuantity * ldblItemCustMtrl)), 2) & "," & System.Math.Round(Val(CStr(lintItemQuantity * ldblItemToolCost)), 2) & ",'" & strHSNSACCode & "','" & strHSNSACType & "','" & CGSTType & "'," & dblCGSTPercentLine & "," & dblCGSTAmtLine & ",'" & SGSTType & "'," & dblSGSTPercentLine & "," & dblSGSTAmtLine & ",'" & IGSTType & "'," & dblIGSTPercentLine & "," & dblIGSTAmtLine & ",'" & UTGSTType & "'," & dblUTGSTPercentLine & "," & dblUTGSTAmtLine & ",'" & COMPCESSType & "'," & dblCCESSPercentLine & "," & dblCCESSAmtLine & ",0,0," & dblItemTotalLine & ")" & vbCrLf
            End If
        End With
        rsItemMst.ResultSetClose()
        rsCustItemMst.ResultSetClose()
        InsertinSalesDtlinEditModeGST = strSalesDtl
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Function
    End Function
    Private Function ValidateGSTTaxes() As Boolean
        Dim result As Boolean = True
        Dim cgst As String = String.Empty
        Dim sgst As String = String.Empty
        Dim igst As String = String.Empty
        Dim utgst As String = String.Empty
        Dim hsnCode As String = String.Empty
        With SpChEntry
            For i As Integer = 1 To .MaxRows
                .Row = i
                .Col = HSN_SAC_CODE
                hsnCode = .Text
                .Col = CGST_TYPE
                cgst = .Text
                .Col = SGST_TYPE
                sgst = .Text
                .Col = IGST_TYPE
                igst = .Text
                .Col = UTGST_TYPE
                utgst = .Text
                If Len(Trim(hsnCode)) = 0 Then
                    MsgBox("HSN/SAC Codes can't be blank", MsgBoxStyle.Information, "eMPro")
                    result = False
                    .Row = i
                    .Col = 5
                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                    .Focus()
                    Exit For
                End If
                If Len(Trim(cgst & sgst & igst & utgst)) = 0 Then
                    MsgBox("GST Types can't be blank", MsgBoxStyle.Information, "eMPro")
                    result = False
                    .Row = i
                    .Col = 5
                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                    .Focus()
                    Exit For
                End If
            Next
        End With
        Return result
    End Function
    Private Sub cmdHelpTCSTax_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelpTCSTax.Click
        On Error GoTo ErrHandler
        Dim strHelp As String
        Dim rssalechallan As ClsResultSetDB
        Dim salechallan As String
        Dim strInvoiceType As Object
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                If CmdGrpChEnt.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                    rssalechallan = New ClsResultSetDB
                    salechallan = ""
                    salechallan = "select b.Description, b.Sub_type_Description from SalesChallan_dtl a,saleconf b where doc_no = " & Trim(txtChallanNo.Text)
                    salechallan = salechallan & " and a.Location_code = b.Location_code  and a.Unit_Code = b.Unit_Code and a.unit_code='" + gstrUNITID + "' and a.Invoice_type = b.invoice_type and a.sub_category = b.Sub_type"
                    rssalechallan.GetResult(salechallan, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                    If rssalechallan.GetNoRows > 0 Then
                        rssalechallan.MoveFirst()
                        strInvoiceType = rssalechallan.GetValue("Description")
                    End If
                    rssalechallan.ResultSetClose()
                Else
                    strInvoiceType = CmbInvType.Text
                End If
                If Len(Me.txtTCSTaxCode.Text) = 0 Then 'To check if There is No Text Then Show All Help
                    strHelp = ShowList(1, (txtTCSTaxCode.MaxLength), "", "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND (Tx_TaxeID='TCS')  and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))")
                    If strHelp = "-1" Then 'If No Record Exists In The Table
                        Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Exit Sub
                    Else
                        txtTCSTaxCode.Text = strHelp
                    End If
                Else
                    strHelp = ShowList(1, (txtTCSTaxCode.MaxLength), txtTCSTaxCode.Text, "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND (Tx_TaxeID='TCS')  and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))")
                    If strHelp = "-1" Then 'If No Record Exists In The Table
                        Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Exit Sub
                    Else
                        txtTCSTaxCode.Text = strHelp
                    End If
                End If
                Call txtTCSTaxCode_Validating(txtTCSTaxCode, New System.ComponentModel.CancelEventArgs(False))
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub txtTCSTaxCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles TxtTCSTaxcode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim strInvoiceType As String
        Dim rsChallanEntry As ClsResultSetDB
        On Error GoTo ErrHandler
        If Len(txtTCSTaxCode.Text) > 0 Then
            If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                strInvoiceType = UCase(Trim(CmbInvType.Text))
            ElseIf (CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT) Or (CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW) Then
                rsChallanEntry = New ClsResultSetDB
                rsChallanEntry.GetResult("Select a.Description,a.Sub_Type_Description from SaleConf a,SalesChallan_Dtl b where A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" + gstrUNITID + "' AND Doc_No = " & txtChallanNo.Text & " and a.Invoice_Type = b.Invoice_type and a.Sub_type = b.Sub_Category and a.Location_code = b.Location_code and (fin_start_date <= getdate() and fin_end_date >= getdate())")
                strInvoiceType = UCase(rsChallanEntry.GetValue("Description"))
                rsChallanEntry.ResultSetClose()
            End If
            If CheckExistanceOfFieldData((txtTCSTaxCode.Text), "TxRt_Rate_No", "Gen_TaxRate", " (Tx_TaxeID='TCS')  and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))") Then
                'lblTCSTaxPerDes.Text = CStr(GetTaxRate((txtTCSTaxCode.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " (Tx_TaxeID='TCS')"))

            Else
                Call ConfirmWindow(10248, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                Cancel = True
                txtTCSTaxCode.Text = ""
                If txtTCSTaxCode.Enabled Then txtTCSTaxCode.Focus()
            End If
        End If
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    '101188073
End Class