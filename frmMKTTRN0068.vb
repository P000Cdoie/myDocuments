Option Strict Off
Option Explicit On
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports VB = Microsoft.VisualBasic
Imports System.IO

Friend Class frmMKTTRN0068
	Inherits System.Windows.Forms.Form
	'===================================================================================
	' (c) MothersonSumi INfotech & Design Ltd. All rights reserverd.
    ' File Name         :   frmMKTTRN0068.frm
	' Function          :   Used to add sale deatails
	' Created By        :   Arshad Ali (Created as separate for SunVac due to nagare related changes)
	' Created On        :   06 June, 2004
	' Revision History  :   Nisha Rai
    '21/09/2001 MARKED CHECKED BY  changed on version 3
    '03/10/2001 MARKED CHECKED BY BCs  for jobwork invoice changed on version 7
    '09/10/2001  changed on version 8 for schedule Status
    '09/01/2002 changed fof Smiel Chennei to add CVD_PER,SVD_Per,Insurance
    '25/01/2002 changed for decimal 4 places on ChBCsacked Out Form No = 4019
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
	'11/06/2002 for from s box size changes in Quantity Check variable type int to double
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
	'    -   Nitin Sood
	'01/07/2002 Change for More then one Selected RGPs Nos
	'05/07/02 changed for rounding off data
	'CHANGED ON 15/07/2002 FOR EXPORT OPTION ADDING AND CALCULATION SAME AS NORMAL INVOICE
	'23/07/2002 changed to add Grin Linking in Rejection Invoice
	'07/08/2002 changed for Jobwork invoice to check Customer supplied from Vendor Bom
	'Changed by nisha on 26/08/2002
	'changes done by nisha to check SO Qty in Challan Entry
	'CHANGES DONE BY NISHA ON 13/03/2003
	'1.FOR FINAL MERGING & FOR FROM BOX & TO bOX UPDATION WHILE EDITING INVOICE
	'2.For Grin Cancellation flag
	'3.SAMPLE INVOICE TOOL COST COLUMN
	'4.CUNSUMABLES & MISC. SALE IN CASE OF NORMAL RAW MATERIAL INVOICE
	'15/03/2003 added per value on form
	'17/04/2003 to add insert string of daily despatchs by nisha
	'13/05/2003 changes done by nisha on 16/05/2003 21/05/2003  for summit issues
	'13/05/2003 changes done by nisha on04/06/2003 for BIN CARD Printing
	'12/06/2003 changes done by nisha on04/06/2003 for BIN CARD Printing & sumit issue exchange rate multiplication
	'24/06/2003 changes for insurance in S Tax claculation25/06/2003
	'cHANGES DON BY NISHA ON 27/06/2003 FOR SERVICE TAX SELECTIO IN VALIDATE EVENT
	'Changes done by nisha on 08/07/2003 for loading charges09/07/20031,10/07/200311/07/2003
	'Changed By Nisha on 17/07/2003
	'Changed By Nisha on 22/08/2003 for allowing loading Charges in Normal invoice
	'changed by nisha on 29/08/2003 for DTA Calculation
	'Changes Done By nisha on 04/09/2003 for
	'1.To Increase the size of Integer from 3 to 6 in case of insurance & Freight
	'2.To set decimal places in Quanity acc to No of Decimals allowed in Measure Master
	'Changes By Preety on 12/09/2003 for
	'1. making the provision of Discount Entry
	'Changes By Nisha on 19/09/2003 for
	'1. To Add Discount_per As well
	'2. To Correct Others Calculations
	'Changes By Nisha on 07/11/2003 for
	'1.to Add Total Excise value in other then EOU Units
	'Changes Done By Pooja on 30/01/2004
	' 1.sales tax enabled Box Enabled in Case of Transfer , focus Setting
	' 2. Allowed Trading & Finished Good in Transfer - Inputs
	'Changes Added by nisha on 16/02/2004
	'1.to add tool Amortisation
	'Changes Added by nisha on 20/02/2004
	'1.to add To add eNagare Items
	'Changes Added by nisha on 25/02/2004
	'1.to add To add TCS Tax
	'changes done by nisha for Disablein TSC Tax in Case of other then normal Scrap invoice on 19/02/2003
	'---------------------------------------------------------------------------------------
	'15/06/2004 By Arshad
	'Saving has been changed to consider roundoff on the basis of parameters from Sale_parameter
	'if Roundoff_flag is 1 then amount will be rounded on 0 decimal places
	'if Roundoff_flag is 0 then amount will be rounded on decimal places defined in sales_parameter
	'Doing so resolve the problem of accounts ie. DR amount is not equal to CR amount
	'---------------------------------------------------------------------------------------
	'01/07/2004 By Arshad Ali
	'Incase of Sample Invoice User should be allowed to enter tool cost and it should be printed on Invoice
	'---------------------------------------------------------------------------------------
	'08/07/2004 By Arshad Ali
	'Query changed by Arshad ie. sum of sales_quantity
	'---------------------------------------------------------------------------------------
	'08/07/2004 By Arshad Ali
	'ECESS Tax Type field added
	'ECESS is to be calculated on total excise value
	'when calculating Sale Tax ECESS Amount to be considered along with basic value, excise value etc.
	'---------------------------------------------------------------------------------------
	'Added by Arshad on 12/07/2004 to allow user to enter tool cost incase of Sample invoice
	'---------------------------------------------------------------------------------------
	'Revision  By       : Sandeep Chadha
	'Revision On        : 13 Jan, 2005
	'History            : 'A)Correct Calculation of Balance Qty KanBan No.
	'  By considering both the table Printedsrv_dtl, Sales_DTL
	'---------------------------------------------------------------------------------------
	'Revision  By       : Arshad Ali
	'Revision On        : 05 April, 2005
	'History            : Parametrized E-nagare help for item code and cust drgno in Invoice Entry Against E-nagare form.
	'=======================================================================================
	'Revision  By       : Sandeep Chadha
	'Revision On        : 15 Apr 2005
	'History            : Add SO No In ENagare Help
	'=======================================================================================
	'=======================================================================================
	'Revision  By       : Sandeep Chadha
	'Revision On        : 21 Apr 2005
	'History            : Add new Parameter in Sales_Parameter (Enagare_TextPrinting)
    ' Use it for Printing Invoice in frmMKTTRN0068
	'=======================================================================================
	'=======================================================================================
	'Revision  By       : Sandeep Chadha
	'Revision On        : 24-May-2005
	'History            : Show the Invoice Type & Sub Type in View & Edit mode
	'=======================================================================================
	'=======================================================================================
	'Revision  By       : Arshad Ali
	'Revision On        : 01-August-2005
	'History            : Provision to make invoice against Multiple SO
	'=======================================================================================
	'Revision  By       : Ashutosh Verma
	'Revision On        : 25-08-2005
	'History            : Resolve Error log for Issue Id:14999
	'=======================================================================================
	'Revision  By       : Ashutosh Verma
	'Revision On        : 30-08-2005
	'History            : Resolve Error log for Issue Id:14999, reported on 30-08-2005
	'=======================================================================================
	'Revision  By       : Ashutosh Verma
	'Revision On        : 19-01-2006
	'History            : Bug fix - After cancellation user can't recreate the invoice,issue Id:16907.
	'=======================================================================================
	'Revision  By       : Ashutosh Verma
	'Revision On        : 09-03-2006 ,issue id :17229.
	'History            : Calculate dispatches from Printedsrv & 57F4 challan at the time of invoice saving.
	'=======================================================================================
	'Revision  By       : Ashutosh , Issue Id :17610
	'Revision On        : 24-04-2006
	'History            : Save Bin Quantity in invoice.
	'                   : Save Stock Location in Invoice.
	'                   : Validate UOM for sales Quantity & Bin Quantity  from measurement master.
	'-----------------------------------------------------------------------------------
	'Revised By      : Davinder Singh
	'Issue ID        : 19573
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
    'Issue ID        : eMpro -20080425 - 17845
    'Revision Date   : 25-Apr-2007
    'History         : Stock Location variable is not getiing refresh due to this Normal Invoice is saved with 01M1 Location
    '***********************************************************************************
    'Revised By      : Manoj Kr. Vaish
    'Issue ID        : eMpro-20090223-27780
    'Revision Date   : 23 Feb 2009
    'History         : Wrong values was showing in From Box and To Box field
    '***********************************************************************************
    'Revised By      : Manoj Kr. Vaish
    'Issue ID        : eMpro-20090820-35232
    'Revision Date   : 20 Aug 2009
    'History         : Currnecy description was not refreshing in Labels
    '***********************************************************************************
    'Revised By      : Siddharth Ranjan
    'Issue ID        : eMpro-20090930-36886
    'Revision Date   : 30 Sep 2009
    'History         : Restrict CSM editing on Invoice against Nagare
    'Modified By Sanchi on 24 May 2011
    '   Modified to support MultiUnit functionality
    '***********************************************************************************
    '----------------------------------------------------------------------------------------------
    ' REVISION DATE     : 24 FEB 2014
    ' REVISED BY        : VINOD SINGH
    ' ISSUE ID          : 10533478
    ' REVISION HISTORY  : WIP FG FUNCTIONALITY CHANGES
    '----------------------------------------------------------------------------------------------
    ' REVISION DATE     : 01 JULY 2014
    ' REVISED BY        : PRASHANT RAJPAL
    ' ISSUE ID          : 10623079
    ' REVISION HISTORY  : WIP CHANGES FOR VACUFORM CHANGES 
    '************************************************************************
    'REVISED BY     :  PRASHANT RAJPAL
    'REVISED DATE   :  13-JAN-2015
    'ISSUE ID       :  10736222
    'PURPOSE        :  TO INTEGRATE CT2 AR3 FUNCTIONALITY 
    '****************************************************************************************
    'Created By     : Parveen Kumar
    'Created On     : 13 FEB 2015
    'Description    : eMPro Vehicle BOM
    'Issue ID       : 10737738 
    '-----------------------------------------------------------------------------------------
    '****************************************************************************************
    'REVISED BY     :  PRASHANT RAJPAL
    'REVISED DATE   :  29-JUN-2015
    'ISSUE ID       :  10808160 
    'PURPOSE        :  EOP FUNCTIONALITY
    '****************************************************************************************
    'REVISED BY     :  ASHISH SHARMA
    'REVISED DATE   :  25-03-2020
    'PURPOSE        :  TCS GST CHANGES
    '****************************************************************************************


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
	Dim intDiscountType As Short
	Dim mstrSONo As String
    Dim blnInvoiceAgainstMultipleSO As Boolean
    Dim mbln_CSM_Edit_Req As Boolean
    Dim RdAddSold As ReportDocument
    Dim Frm As New eMProCrystalReportViewer
    Dim SchUpdFlag As Boolean = False   '10737738
    '10623079
    Dim mblnskipdacinvoicebincheck As Boolean

    Dim gblnGSTUnit As Boolean 'by abhijit
    Dim dtSalesParameter As New DataTable 'by Abhijit

    Dim HSNSAC_Check As Boolean = False
    Dim GST_TYPE_Check As Boolean = False
    Dim blnGSTTAXroundoff As Boolean
    Dim intGSTTAXroundoff_decimal As Short


    '10623079
    	Private Enum GridHeader
		InternalPartNo = 1
        CustPartNo
        RatePerUnit
        CustSuppMatPerUnit
        Quantity
        CustRefNo
        AmendmentNo
        srvdino
        SRVLocation
        USLOC
        SChTime
        Packing
        EXC
        CVD
        SAD
        OthersPerUnit
        FromBox
        ToBox
        CumulativeBoxes
        delete
        ToolCostPerUnit
        Rate
        CustMtrl
        Others
        ToolCost
        model
        BinQty

        'TT#101188073 Added by Abhijit
        HSNSACCODE
        ISHSNORSAC

        Basic_Value
        Discount_Percent
        Discount_Amt
        Accessible_Amt

        'CGST_SAC_No
        'HSN_SAC_Type

        IGST_Type
        IGST_Percent
        IGST_Amt

        CGST_Type
        CGST_Percent
        CGST_Amt

        SGST_Type
        SGST_Percent
        SGST_Amt

        UTGST_Type
        UTGST_Percent
        UTGST_Amt

        CSESS_TAX_Type
        CSESS_TAX_Percent
        CSESS_TAX_VAL

        Item_Total

        'TT#101188073 Added by Abhijit

	End Enum
 	
	Private Enum enumExciseType
		RETURN_EXCISE = 1
		RETURN_CVD = 2
		RETURN_SAD = 3
		RETURN_ALLExcise = 4
	End Enum
    Dim objInvoicePrint As New prj_InvoicePrinting.clsInvoicePrinting(gstrDateFormat)
    Dim intNoCopies As Short
    Dim strStockLocation As String
	Dim mAmortization As Double
	Dim mStrCustMst As String
	Dim mblnEOUUnit As Boolean
	Dim mAssessableValue As Double
	Dim mOpeeningBalance As Double
	Dim strsaledetails As String
	Dim strupdateGrinhdr As String
	Dim strupdateitbalmst As String
	Dim strupdatecustodtdtl As String
	Dim strUpdateAmorDtl As String
	Dim salesconf As String
	Dim msubTotal, mInvNo, mExDuty, mBasicAmt, mOtherAmt As Double
	Dim mFrAmt, mGrTotal, mStAmt, mCustmtrl As Double
	Dim mDoc_No As Short
	Dim mAccount_Code, mInvType, mSubCat, mlocation As String
	Dim mstrAnnex As String
	Dim mCust_Ref, mAmendment_No As String
	Dim saleschallan As String
	Dim arrCustAnnex() As Object
	Dim ref57f4 As String 'used in BomCheck() insertupdateAnnex()
	Dim dblFinishedQty As Double 'To get Qty of Finished Item from Spread
	Dim strCustCode As String 'used in BomCheck() insertupdateAnnex()
	Dim strItemCode As String 'used in BomCheck() insertupdateAnnex()
	Dim updatestockflag, updatePOflag As Boolean
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
 	Dim mstrNagareDate As String
     Dim mstrQuantity As String
	Dim mstrSRVDINo As String
	Dim mstrSRVLocation As String
	Dim mstrUSLoc As String
	Dim mstrSchTime As String
	Dim mIntRecordCount As Short
	Dim strupdateamordtlbom As String ' changes done by nisha to impliment tool amortization
    Dim blnGridStatus As Boolean
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
    'Changes against 10737738 
    Private Sub ChkVBSchUpdFlag()
        Dim strSql As String = String.Empty

        Try

            strSql = " select top 1 1 from sales_parameter where Unit_Code='" & gstrUNITID & "' and SCHEDULE_UPLOAD_CONFIG = 1  "
            SchUpdFlag = IsRecordExists(strSql)

        Catch ex As Exception
            Throw ex
        End Try

    End Sub

 	Private Sub chkExciseExumpted_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles chkExciseExumpted.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		'****************************************************
		'Created By     -  Tapan
		'Description    -  At Enter Key Press Set Focus To Next Control
		'****************************************************
		On Error GoTo ErrHandler
		Select Case KeyAscii
			Case System.Windows.Forms.Keys.Return
				Select Case Me.CmdGrpChEnt.mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        OptDiscountValue.Focus()
                End Select
			Case 39, 34, 96
				KeyAscii = 0
		End Select
		GoTo EventExitSub
ErrHandler: 'The Error Handling Code Starts here
		Call gobjError.RaiseError(Err.Number, err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub: 
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
    Private Sub CmbInvSubType_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmbInvSubType.SelectedIndexChanged
        On Error GoTo ErrHandler
        If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then Exit Sub
        Call SelectInvTypeSubTypeFromSaleConf((CmbInvType.Text), (CmbInvSubType.Text))
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
	Private Sub CmbInvSubType_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles CmbInvSubType.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		'****************************************************
		'Created By     -  Nisha
		'Description    -  At Enter Key Press Set Focus To Next Control
		'****************************************************
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
		If UCase(CmbInvType.Text) = "NORMAL INVOICE" Then
			If UCase(CmbInvSubType.Text) = "SCRAP" Then
                 ctlPerValue.Enabled = True
                 ctlPerValue.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
				txtTCSTaxCode.Enabled = True : txtTCSTaxCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : cmdHelpTCSTax.Enabled = True
				txtTCSTaxCode.Text = ""
				If blnInvoiceAgainstMultipleSO Then Exit Sub
				txtRefNo.Enabled = False : txtRefNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : CmdRefNoHelp.Enabled = False : txtRefNo.Text = ""
				txtAmendNo.Enabled = False : txtAmendNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : txtAmendNo.Text = ""
			Else
                ctlPerValue.Enabled = False
                ctlPerValue.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                txtTCSTaxCode.Enabled = True : txtTCSTaxCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdHelpTCSTax.Enabled = True
                txtTCSTaxCode.Text = ""
                If blnInvoiceAgainstMultipleSO Then Exit Sub
                txtRefNo.Enabled = True : txtRefNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : CmdRefNoHelp.Enabled = True
                txtAmendNo.Enabled = True : txtAmendNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            End If
        End If
        SpChEntry.MaxRows = 0
    End Sub
    Private Sub CmbInvType_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmbInvType.SelectedIndexChanged
        On Error GoTo ErrHandler

        If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then Exit Sub
        Call SelectInvoiceSubTypeFromSaleConf((CmbInvType.Text))
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub cmbInvType_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles CmbInvType.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '****************************************************
        'Created By     -  Nisha
        'Description    -  At Enter Key Press Set Focus To Next Control
        '****************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        If CmbInvSubType.Enabled Then CmbInvSubType.Focus()
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
         Select Case CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                Select Case UCase(CmbInvType.Text)
                     Case "NORMAL INVOICE", "EXPORT INVOICE", "SERVICE INVOICE"
                     
                        If UCase(CmbInvType.Text) = "SERVICE INVOICE" Then
                            lblSaleTaxType.Text = "Service Tax Code"
                        Else
                            lblSaleTaxType.Text = "Sale Tax    Code"
                        End If
                        If blnInvoiceAgainstMultipleSO Then Exit Sub
                        ctlPerValue.Enabled = False
                        ctlPerValue.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        txtRefNo.Enabled = True : txtRefNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : CmdRefNoHelp.Enabled = True
                         ctlInsurance.Enabled = True
                        ctlInsurance.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                         txtSaleTaxType.Enabled = True : txtSaleTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        lblSaltax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        CmdSaleTaxType.Enabled = True
                        txtSurchargeTaxType.Enabled = True : txtSurchargeTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        lblSurcharge_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        cmdSurchargeTaxCode.Enabled = True
                         txtSurchargeTaxType.Text = "" : txtRefNo.Text = ""
                         ctlInsurance.Text = "" : txtSaleTaxType.Text = ""
                        lblCurrencyDes.Text = ""
                          With SpChEntry
                            .Col = GridHeader.ToolCostPerUnit : .Col2 = GridHeader.ToolCostPerUnit : .BlockMode = True : .ColHidden = True : .BlockMode = False
                        End With
                     
                        If UCase(CmbInvType.Text) = "NORMAL INVOICE" Then
                            If UCase(CmbInvSubType.Text) = "SCRAP" Then
 txtTCSTaxCode.Enabled = True : txtTCSTaxCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : cmdHelpTCSTax.Enabled = True
 txtTCSTaxCode.Text = ""
                            Else
                                txtTCSTaxCode.Enabled = True : txtTCSTaxCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdHelpTCSTax.Enabled = True
 txtTCSTaxCode.Text = ""
                            End If
                        Else
                            txtTCSTaxCode.Enabled = True : txtTCSTaxCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdHelpTCSTax.Enabled = True
                            txtTCSTaxCode.Text = ""
                        End If
                     Case "JOBWORK INVOICE"
                        txtTCSTaxCode.Enabled = True : txtTCSTaxCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdHelpTCSTax.Enabled = True
                        txtTCSTaxCode.Text = ""
                        ctlPerValue.Enabled = False
                        ctlPerValue.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        If blnInvoiceAgainstMultipleSO Then Exit Sub
                        txtRefNo.Enabled = True : txtRefNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : CmdRefNoHelp.Enabled = True
                        ctlInsurance.Enabled = False
                        ctlInsurance.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        txtSaleTaxType.Enabled = True : txtSaleTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        lblSaltax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        CmdSaleTaxType.Enabled = False
                        txtSurchargeTaxType.Enabled = True : txtSurchargeTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        lblSurcharge_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        cmdSurchargeTaxCode.Enabled = True
                         txtSurchargeTaxType.Text = "" : txtRefNo.Text = ""
                         ctlInsurance.Text = "" : txtSaleTaxType.Text = ""
                        lblCurrencyDes.Text = ""
                        With SpChEntry
                            .Col = GridHeader.ToolCostPerUnit : .Col2 = GridHeader.ToolCostPerUnit : .BlockMode = True : .ColHidden = True : .BlockMode = False
                        End With
                     Case "SAMPLE INVOICE", "TRANSFER INVOICE"
                        txtTCSTaxCode.Enabled = True : txtTCSTaxCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdHelpTCSTax.Enabled = True
                        txtTCSTaxCode.Text = ""
                        ctlPerValue.Enabled = False
                        ctlPerValue.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        If blnInvoiceAgainstMultipleSO Then Exit Sub
                        txtRefNo.Enabled = False : txtRefNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : CmdRefNoHelp.Enabled = False
                         txtSurchargeTaxType.Text = "" : txtRefNo.Text = ""
                         ctlInsurance.Text = "" : txtSaleTaxType.Text = ""
                        lblCurrencyDes.Text = ""
                        
                        If UCase(CmbInvType.Text) = "TRANSFER INVOICE" Then
                            ctlInsurance.Enabled = True
                            ctlInsurance.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                            txtSaleTaxType.Enabled = True : txtSaleTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                            lblSaltax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                            CmdSaleTaxType.Enabled = True
                            txtSurchargeTaxType.Enabled = False : txtSurchargeTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                            lblSurcharge_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                            cmdSurchargeTaxCode.Enabled = False
                            With SpChEntry
 .Col = GridHeader.ToolCostPerUnit : .Col2 = GridHeader.ToolCostPerUnit : .BlockMode = True : .ColHidden = True : .BlockMode = False
                            End With
                          Else
                            ctlInsurance.Enabled = False
                            ctlInsurance.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                            txtSaleTaxType.Enabled = True : txtSaleTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                            lblSaltax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                            CmdSaleTaxType.Enabled = True
                            txtSurchargeTaxType.Enabled = True : txtSurchargeTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                            lblSurcharge_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                            cmdSurchargeTaxCode.Enabled = True
                            With SpChEntry
 .Col = GridHeader.ToolCostPerUnit : .Col2 = GridHeader.ToolCostPerUnit : .BlockMode = True : .ColHidden = False : .BlockMode = False
                            End With
                          End If
                    Case "REJECTION"
                        txtTCSTaxCode.Enabled = True : txtTCSTaxCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdHelpTCSTax.Enabled = True
                        txtTCSTaxCode.Text = ""
                        ctlPerValue.Enabled = False
                         ctlPerValue.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        txtSurchargeTaxType.Text = "" : txtRefNo.Text = ""
                         ctlInsurance.Text = "" : txtSaleTaxType.Text = ""
                        lblCurrencyDes.Text = ""
                          ctlInsurance.Enabled = True
                        ctlInsurance.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        txtSaleTaxType.Enabled = True : txtSaleTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        lblSaltax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        CmdSaleTaxType.Enabled = True
                        txtSurchargeTaxType.Enabled = True : txtSurchargeTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        lblSurcharge_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        cmdSurchargeTaxCode.Enabled = True
                        
                        With SpChEntry
                            .Col = GridHeader.ToolCostPerUnit : .Col2 = GridHeader.ToolCostPerUnit : .BlockMode = True : .ColHidden = True : .BlockMode = False
                        End With
                 End Select
        End Select
        SpChEntry.MaxRows = 0
         lblCreditTerm.Text = ""
        lblCreditTermDesc.Text = ""
      End Sub
    Private Sub CmbTransType_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles CmbTransType.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '****************************************************
        'Created By     -  Nisha
        'Description    -  At Enter Key Press Set Focus To Next Control
        '****************************************************
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
        '****************************************************
        'Created By     -  Nisha
        'Description    -  To Display Help From SalesChallan_Dtl
        '****************************************************
        On Error GoTo ErrHandler
        Dim strHelpString As String
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                'Check Location Code Field
                If Trim(txtLocationCode.Text) = "" Then
                    Call ConfirmWindow(10239, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO, 100)
                    If txtLocationCode.Enabled Then txtLocationCode.Focus()
                    Exit Sub
                End If
                If Len(Trim(txtChallanNo.Text)) = 0 Then
                    If blnEOU_FLAG = True Then
                        strHelpString = ShowList(1, (txtChallanNo.MaxLength), "", "Doc_No", DateColumnNameInShowList("Invoice_Date") & "As Invoice_Date", "SalesChallan_Dtl ", "AND Location_Code='" & Trim(txtLocationCode.Text) & "' and invoice_type <> 'EXP' and cancel_flag = 0")
                    Else
                        strHelpString = ShowList(1, (txtChallanNo.MaxLength), "", "Doc_No", DateColumnNameInShowList("Invoice_Date") & "As Invoice_Date", "SalesChallan_Dtl ", "AND Location_Code='" & Trim(txtLocationCode.Text) & "'")
                    End If
                    If strHelpString = "-1" Then 'If No Record Found
                        Call ConfirmWindow(10253, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        If txtChallanNo.Enabled Then txtChallanNo.Focus()
                    Else
                        txtChallanNo.Text = strHelpString
                    End If
                Else
                    If blnEOU_FLAG = False Then
                        strHelpString = ShowList(1, (txtChallanNo.MaxLength), txtChallanNo.Text, "Doc_No", DateColumnNameInShowList("Invoice_Date") & "As Invoice_Date", "SalesChallan_Dtl ", "AND Location_Code='" & Trim(txtLocationCode.Text) & "'")
                    Else
                        strHelpString = ShowList(1, (txtChallanNo.MaxLength), txtChallanNo.Text, "Doc_No", DateColumnNameInShowList("Invoice_Date") & "As Invoice_Date", "SalesChallan_Dtl ", "AND Location_Code='" & Trim(txtLocationCode.Text) & "' and invoice_type <> 'EXP'")
                    End If
                    If strHelpString = "-1" Then 'If No Record Found
                        Call ConfirmWindow(10253, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        If txtChallanNo.Enabled Then txtChallanNo.Focus()
                    Else
                        txtChallanNo.Text = strHelpString
                    End If
                End If
        End Select
        txtChallanNo.Focus()
         If Val(txtChallanNo.Text) > 99000000 Then
            If Not blnInvoiceAgainstMultipleSO Then
                Cmditems.Enabled = True
            End If
        Else
            Cmditems.Enabled = False
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub cmdConsigneeCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdConsigneeCancel.Click
        fraconsigneeDetails.Visible = False
        If CmdGrpChEnt.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            txtContactPerson.Text = "" : txtECC.Text = "" : txtLST.Text = "" : txtAddress1.Text = "" : txtAddress2.Text = "" : txtAddress3.Text = ""
        End If
        cmdConsigneeDetails.Focus()
    End Sub
    Private Sub cmdConsigneeDetails_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdConsigneeDetails.Click
        fraconsigneeDetails.Visible = True
        If txtContactPerson.Enabled = True Then
            txtContactPerson.Focus()
        Else
            cmdConsigneeOK.Enabled = True : cmdConsigneeCancel.Enabled = True
            cmdConsigneeOK.Focus()
        End If
    End Sub
    Private Sub cmdConsigneeOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdConsigneeOK.Click
        fraconsigneeDetails.Visible = False : cmdConsigneeDetails.Focus()
    End Sub
    Private Sub CmdCustCodeHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdCustCodeHelp.Click
        '****************************************************
        'Created By     -  Nisha
        'Description    -  To Display Help From Customer_Mst
        '****************************************************
        Dim strCustMst As String
        Dim rsCustMst As ClsResultSetDB
        On Error GoTo ErrHandler
        Dim strHelpString As String
        If Len(Trim(txtLocationCode.Text)) = 0 Then
            Call ConfirmWindow(10116, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
            If txtLocationCode.Enabled Then txtLocationCode.Focus()
            Exit Sub
        End If
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                'Changes against 10737738 
                If UCase(Trim(mstrInvoiceType)) = "INV" Or UCase(Trim(mstrInvoiceType)) = "SMP" Or UCase(Trim(mstrInvoiceType)) = "TRF" Or UCase(Trim(mstrInvoiceType)) = "JOB" Or UCase(Trim(mstrInvoiceType)) = "EXP" Or UCase(Trim(mstrInvoiceType)) = "SRC" Then
                    If Len(Trim(txtCustCode.Text)) = 0 Then
                        If SchUpdFlag = True Then
                            strHelpString = ShowList(1, (txtCustCode.MaxLength), "", "Customer_Code", "Cust_Name", "Customer_Mst", " AND SCH_UPLOAD_CODE='KENN' AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))")
                        Else
                            strHelpString = ShowList(1, (txtCustCode.MaxLength), "", "Customer_Code", "Cust_Name", "Customer_Mst", " AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))")
                        End If
                        If strHelpString = "-1" Then 'If No Record Found
                            Call ConfirmWindow(10225, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Else
                            txtCustCode.Text = strHelpString
                            Call SelectDescriptionForField("Cust_Name", "Customer_Code", "Customer_Mst", lblCustCodeDes, (txtCustCode.Text))
                        End If
                        'Else

                        '    strHelpString = ShowList(1, (txtCustCode.MaxLength), "", "Customer_Code", "Cust_Name", "Customer_Mst", " AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))")
                        '    If strHelpString = "-1" Then 'If No Record Found
                        '        Call ConfirmWindow(10225, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        '    Else
                        '        txtCustCode.Text = strHelpString
                        '        Call SelectDescriptionForField("Cust_Name", "Customer_Code", "Customer_Mst", lblCustCodeDes, (txtCustCode.Text))
                        '    End If
                    End If
                Else 'Select Help From Vendor Master
                    If Len(Trim(txtCustCode.Text)) = 0 Then

                        strHelpString = ShowList(1, (txtCustCode.MaxLength), "", "Vendor_Code", "Vendor_name", "Vendor_Mst")
                        If strHelpString = "-1" Then 'If No Record Found
                            Call ConfirmWindow(10225, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Else
                            txtCustCode.Text = strHelpString
                            Call SelectDescriptionForField("Vendor_name", "Vendor_Code", "Vendor_Mst", lblCustCodeDes, (txtCustCode.Text))
                        End If
                    Else
                        strHelpString = ShowList(1, (txtCustCode.MaxLength), "", "Vendor_Code", "Vendor_name", "Vendor_Mst")
                        If strHelpString = "-1" Then 'If No Record Found
                            Call ConfirmWindow(10225, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Else
                            txtCustCode.Text = strHelpString
                            Call SelectDescriptionForField("Vendor_name", "Vendor_Code", "Vendor_Mst", lblCustCodeDes, (txtCustCode.Text))
                        End If
                    End If
                End If
        End Select

        If Len(Trim(txtCustCode.Text)) > 0 Then
            rsCustMst = New ClsResultSetDB

            strCustMst = "SELECT Bill_Address1 + ', '  +  Bill_Address2 + ', ' + Bill_City + ' - ' + Bill_Pin as  invoiceAddress from Customer_Mst where Customer_code ='" & txtCustCode.Text & "' and UNIT_CODE='" & gstrUNITID & "'"
            ''---------------------------Addition Ends----------------------------------
            rsCustMst.GetResult(strCustMst)
            If rsCustMst.GetNoRows > 0 Then
                lblAddressDes.Text = rsCustMst.GetValue("InvoiceAddress")
            End If
            rsCustMst = Nothing
        End If
        Call txtCustCode_Validating(txtCustCode, New System.ComponentModel.CancelEventArgs(False))
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub CmdSECSSTaxType_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdSECSSTaxType.Click
        '-----------------------------------------------------------------------------------
        'Created By      : Davinder Singh
        'Issue ID        : 19575
        'Creation Date   : 27 Feb 2007
        'Function        : To Show help for new Tax SEcess
        '-----------------------------------------------------------------------------------
        Dim strHelp As String
        On Error GoTo ErrHandler
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                If Len(txtSECSSTaxType.Text) = 0 Then 'To check if There is No Text Then Show All Help

                    strHelp = ShowList(1, (txtSECSSTaxType.MaxLength), "", "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", " AND (Tx_TaxeID='ECSSH') AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))")
                    If strHelp = "-1" Then 'If No Record Exists In The Table
                        Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Exit Sub
                    Else
                        txtSECSSTaxType.Text = strHelp
                    End If
                Else

                    strHelp = ShowList(1, (txtSECSSTaxType.MaxLength), txtSECSSTaxType.Text, "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", " AND (Tx_TaxeID='ECSSH') AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))")
                    If strHelp = "-1" Then 'If No Record Exists In The Table
                        Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Exit Sub
                    Else
                        txtSECSSTaxType.Text = strHelp
                    End If
                End If
                Call txtSECSSTaxType_Validating(txtSECSSTaxType, New System.ComponentModel.CancelEventArgs(False))
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtECSSTaxType_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtECSSTaxType.TextChanged
        If Len(txtECSSTaxType.Text) = 0 Then
            lblECSStax_Per.Text = "0.00"
        End If
    End Sub
    Private Sub txtECSSTaxType_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtECSSTaxType.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '****************************************************
        'Created By     -  Arshad Ali
        'Description    -  At Enter Key Press Set Focus To Next Control
        '****************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        txtTCSTaxCode.Focus()
                        txtECSSTaxType_Validating(txtECSSTaxType, New System.ComponentModel.CancelEventArgs(False))
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
    Private Sub txtECSSTaxType_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtECSSTaxType.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F1 Then
            Call CmdECSSTaxType_Click(CmdECSSTaxType, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtECSSTaxType_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtECSSTaxType.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        If Len(txtECSSTaxType.Text) > 0 Then
            If CheckExistanceOfFieldData((txtECSSTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", "(Tx_TaxeID='ECS') and UNIT_CODE='" & gstrUNITID & "' AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))") Then
                lblECSStax_Per.Text = CStr(GetTaxRate((txtECSSTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", "(Tx_TaxeID='ECS') and UNIT_CODE='" & gstrUNITID & "'"))
                If txtTCSTaxCode.Enabled Then txtTCSTaxCode.Focus()
            Else
                Call ConfirmWindow(10248, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                Cancel = True
                txtECSSTaxType.Text = ""
                If txtECSSTaxType.Enabled Then txtECSSTaxType.Focus()
            End If
        End If
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub CmdECSSTaxType_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdECSSTaxType.Click
        '****************************************************
        'Created By     -  Arshad Ali
        'Description    -  To Display Help From Tax Rate Master
        '****************************************************
        Dim strHelp As String
        On Error GoTo ErrHandler
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                If Len(txtECSSTaxType.Text) = 0 Then 'To check if There is No Text Then Show All Help

                    strHelp = ShowList(1, (txtECSSTaxType.MaxLength), "", "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", " AND (Tx_TaxeID='ECS') AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))")
                    If strHelp = "-1" Then 'If No Record Exists In The Table
                        Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Exit Sub
                    Else
                        txtECSSTaxType.Text = strHelp
                    End If
                Else

                    strHelp = ShowList(1, (txtECSSTaxType.MaxLength), txtECSSTaxType.Text, "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND (Tx_TaxeID='ECS') AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))")
                    If strHelp = "-1" Then 'If No Record Exists In The Table
                        Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Exit Sub
                    Else
                        txtECSSTaxType.Text = strHelp
                    End If
                End If
                Call txtECSSTaxType_Validating(txtECSSTaxType, New System.ComponentModel.CancelEventArgs(False))
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
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
        Dim salechallan As String
        Dim intNoOfDecimal As Short
        Dim rsCustItemMst As ClsResultSetDB
        Dim rsSaleConf As ClsResultSetDB
        Dim rsItemMst As ClsResultSetDB
        Dim rsSalesChallandtl As ClsResultSetDB
        Dim rsInvoiceType As ClsResultSetDB
        Dim strNewCurrencyCode As String
        Dim intLoopCounter As Short
        Dim strChallanNo As String
        Dim rsReportName As String
        Dim rsECess As ClsResultSetDB
        Dim intLoop As Short
        Dim strMakeDate As String
        Dim rssalechallan As ClsResultSetDB
        Dim strInvoiceType As String
        Select Case e.Button
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD
                Call EnableControls(True, Me, True)
                lblSaltax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                OptDiscountValue.Checked = True
                '***
                lblSurcharge_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                lblLoadingcharge_per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)

                chkExciseExumpted.CheckState = System.Windows.Forms.CheckState.Unchecked
                Call SelectChallanNoFromSalesChallanDtl()
                txtChallanNo.Enabled = False : txtChallanNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                CmdChallanNo.Enabled = False : txtChallanNo.Enabled = False
                txtRefNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : CmdRefNoHelp.Enabled = False
                lblLocCodeDes.Text = "" : lblCustCodeDes.Text = ""
                lblCustPartDesc.Text = ""
                lblCurrencyDes.Text = ""
                lblExchangeRateValue.Text = ""
                lblCreditTerm.Text = "" : lblCreditTermDesc.Text = ""
                ctlPerValue.Text = "" : lblAddressDes.Text = ""
                Me.SpChEntry.Enabled = True
                If blnEOU_FLAG = False Then

                    For intLoopCounter = 0 To CmbInvType.Items.Count - 1 'Selecting Normal Invoice as default
                        If UCase(Trim(ObsoleteManagement.GetItemString(CmbInvType, intLoopCounter))) = "NORMAL INVOICE" Then
                            Exit For
                        End If
                    Next
                    'added by sanchi 
                    If CmbInvType.Items.Count > 0 Then
                        CmbInvType.SelectedIndex = intLoopCounter
                    End If
                    'end 
                    For intLoopCounter = 0 To CmbInvSubType.Items.Count - 1 'Selecting Finished Goods as default
                        If UCase(Trim(ObsoleteManagement.GetItemString(CmbInvSubType, intLoopCounter))) = "FINISHED GOODS" Then
                            If CmbInvSubType.Items.Count > 0 Then
                                CmbInvSubType.SelectedIndex = intLoopCounter
                            End If
                            Exit For
                        End If
                    Next
                    'added by sanchi
                    If CmbTransType.Items.Count > 0 Then
                        CmbTransType.SelectedIndex = 0
                    End If
                    'end
                Else
                    For intLoopCounter = 0 To CmbInvType.Items.Count - 1 'Selecting Normal Invoice as default
                        If UCase(Trim(ObsoleteManagement.GetItemString(CmbInvType, intLoopCounter))) = "EXPORT INVOICE" Then
                            Exit For
                        End If
                    Next
                    'added by sanchi 
                    If CmbTransType.Items.Count > 0 Then
                        CmbTransType.SelectedIndex = 0
                    End If
                    'end 
                End If
                With Me.SpChEntry
                    .MaxRows = 1
                    .Row = 1 : .Row2 = 1 : .Col = GridHeader.InternalPartNo : .Col2 = .MaxCols : .BlockMode = True : .Text = "" : .Lock = False : .BlockMode = False
                End With
                If Not UCase(CStr(Trim(CmbInvType.Text))) = "NORMAL INVOICE" Or UCase(CStr(Trim(CmbInvType.Text))) = "JOBWORK INVOICE" Or UCase(CStr(Trim(CmbInvType.Text))) = "EXPORT INVOICE" Or UCase(CStr(Trim(CmbInvType.Text))) = "SERVICE INVOICE" Then
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
                lblDateDes.Text = VB6.Format(GetServerDate(), gstrDateFormat)
                With dtpDateDesc
                    .Value = ConvertToDate(VB6.Format(lblDateDes.Text, gstrDateFormat))
                    .Visible = True 'Show DatePicker
                End With
                Call SetMaxLengthInSpread(0)
                Call ChangeCellTypeStaticText()
                lblRGPDes.Text = ""
                txtLocationCode.Text = Find_Value("SELECT l.Location_Code FROM Location_mst l,SaleConf s WHERE s.Location_Code = l.Location_Code and s.UNIT_CODE=l.UNIT_CODE AND s.UNIT_CODE='" & gstrUNITID & "'")
                If Len(gStrLocationCode) > 0 Then
                    txtLocationCode.Text = gStrLocationCode
                    txtLocationCode_Validating(txtLocationCode, New System.ComponentModel.CancelEventArgs(False))
                End If
                If Len(gStrCustomerCode) > 0 Then
                    txtCustCode.Text = gStrCustomerCode
                    txtCustCode_Validating(txtCustCode, New System.ComponentModel.CancelEventArgs(False))
                End If
                txtLocationCode_Validating(txtLocationCode, New System.ComponentModel.CancelEventArgs(False))
                If Len(gStrVehicleNo) > 0 Then
                    txtVehNo.Text = gStrVehicleNo
                End If
                If txtRefNo.Enabled And txtRefNo.Visible Then txtRefNo.Focus()
                txtSRVDINO.Focus()
                txtTCSTaxCode.Enabled = True
                lblTCSTaxPerDes.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                cmdHelpTCSTax.Enabled = True
                txtECSSTaxType.Enabled = True
                lblECSStax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                lblSECSStax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                CmdECSSTaxType.Enabled = True

                'txtECSSTaxType.Text = "EC2"
                'Call txtECSSTaxType_Validating(txtECSSTaxType, New System.ComponentModel.CancelEventArgs(False))
                'If txtSRVDINO.Enabled Then txtSRVDINO.Focus()
                'rsECess = New ClsResultSetDB
                'rsECess.GetResult("Select TxRt_Rate_No,TxRt_Percentage from Gen_TaxRate where tx_TaxeID ='ECSSH' and TxRt_Percentage > 0 AND UNIT_CODE='" & gstrUNITID & "' AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))")
                'If Not rsECess.EOFRecord Then
                '    rsECess.MoveFirst()
                '    txtSECSSTaxType.Text = rsECess.GetValue("TxRt_Rate_No")
                '    lblSECSStax_Per.Text = rsECess.GetValue("TxRt_Percentage")
                'End If
                'rsECess = Nothing

                '------------------Satvir Handa------------------------
                Dim strSql As String = ""

                strSql = "select txrt_Rate_No from Gen_TaxRate where Tx_TaxeID in ('ECS') and DEFAULT_FOR_INVOICE =1 And Unit_Code='" & gstrUNITID & "' and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"
                txtECSSTaxType.Text = Convert.ToString(SqlConnectionclass.ExecuteScalar(strSql))

                Call txtECSSTaxType_Validating(txtECSSTaxType, New System.ComponentModel.CancelEventArgs(False))
                If txtSRVDINO.Enabled Then txtSRVDINO.Focus()
                rsECess = New ClsResultSetDB
                'rsECess.GetResult("Select TxRt_Rate_No,TxRt_Percentage from Gen_TaxRate where UNIT_CODE = '" & gstrUNITID & "' AND tx_TaxeID ='ECSSH' and TxRt_Percentage > 0 and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date)) ")
                rsECess.GetResult("Select TxRt_Rate_No,TxRt_Percentage from Gen_TaxRate where UNIT_CODE = '" & gstrUNITID & "' AND tx_TaxeID ='ECSSH' and DEFAULT_FOR_INVOICE=1 and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date)) ")
                If Not rsECess.EOFRecord Then
                    rsECess.MoveFirst()
                    txtSECSSTaxType.Text = rsECess.GetValue("TxRt_Rate_No")
                    lblSECSStax_Per.Text = rsECess.GetValue("TxRt_Percentage")
                End If
                rsECess = Nothing
                '------------------Satvir Handa------------------------

                CmbInvType.DropDownStyle = ComboBoxStyle.DropDownList
                CmbInvType.DropDownStyle = ComboBoxStyle.DropDownList
                If blnInvoiceAgainstMultipleSO Then
                    Me.SSTab1.Controls.Remove(Me._SSTab1_TabPage2)
                    txtCustCode.Enabled = False
                    txtCustCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    CmdCustCodeHelp.Enabled = False
                    txtRefNo.Enabled = False
                    txtRefNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    CmdRefNoHelp.Enabled = False
                    txtAmendNo.Enabled = False
                    txtAmendNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    Cmditems.Focus()
                End If
                If gblnGSTUnit Then
                    EnableDisable_GST()
                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT
                Call EnableControls(False, Me)
                rsSalesChallandtl = New ClsResultSetDB
                rsSalesChallandtl.GetResult("select Invoice_type,Sub_Category,Currency_code from Saleschallan_dtl where UNIT_CODE='" & gstrUNITID & "' AND doc_no = " & txtChallanNo.Text, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                strNewCurrencyCode = rsSalesChallandtl.GetValue("Currency_code")
                If (UCase(Trim(rsSalesChallandtl.GetValue("Invoice_type"))) = "INV") And (UCase(Trim(rsSalesChallandtl.GetValue("Sub_Category"))) = "L") Then
                    txtTCSTaxCode.Enabled = True : txtTCSTaxCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : cmdHelpTCSTax.Enabled = True
                Else
                    txtTCSTaxCode.Enabled = True : txtTCSTaxCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdHelpTCSTax.Enabled = True
                End If
                If rsSalesChallandtl.GetValue("Invoice_type") <> "JOB" Then
                    ctlInsurance.Enabled = True
                    ctlInsurance.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                End If
                If UCase(rsSalesChallandtl.GetValue("Invoice_type")) = "INV" Or UCase(CStr(rsSalesChallandtl.GetValue("Invoice_type") = "REJ")) Or UCase(CStr(rsSalesChallandtl.GetValue("Invoice_type") = "EXP")) Or UCase(CStr(rsSalesChallandtl.GetValue("Invoice_type") = "SRC")) Then
                    txtSaleTaxType.Enabled = True : txtSaleTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    CmdSaleTaxType.Enabled = True
                    lblSaltax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    lblSurcharge_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    ctlInsurance.Enabled = True
                    ctlInsurance.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                End If
                txtLoadingTaxType.Enabled = True : txtLoadingTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                cmdLoadinfChageHelp.Enabled = True : lblLoadingcharge_per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                fraDiscountType.Enabled = True
                OptDiscountValue.Enabled = True
                OptDiscountPercentage.Enabled = True
                txtDiscountAmt.Enabled = True
                txtDiscountAmt.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                txtFreight.Enabled = True
                txtFreight.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                txtSurchargeTaxType.Enabled = True : txtSurchargeTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                txtRemarks.Enabled = True : txtRemarks.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                txtSRVDINO.Enabled = True : txtSRVDINO.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                txtSRVLocation.Enabled = True : txtSRVLocation.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : cmdhelpSRVDI.Enabled = True
                txtUsLoc.Enabled = True : txtUsLoc.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                txtSchTime.Enabled = True : txtSchTime.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                cmdSurchargeTaxCode.Enabled = True
                chkExciseExumpted.Enabled = True
                SpChEntry.Enabled = True
                SpChEntry.Row = 1 : SpChEntry.Row2 = SpChEntry.MaxRows : SpChEntry.Col = GridHeader.InternalPartNo : SpChEntry.Col2 = SpChEntry.MaxCols
                SpChEntry.BlockMode = True : SpChEntry.Lock = False : SpChEntry.BlockMode = False
                intNoOfDecimal = ToGetDecimalPlaces(Trim(strNewCurrencyCode))
                If intNoOfDecimal < 2 Then
                    intNoOfDecimal = 2
                End If
                Call SetMaxLengthInSpread(intNoOfDecimal)
                Call ChangeCellTypeStaticText()
                ReDim mdblPrevQty(SpChEntry.MaxRows - 1) ' To get value of Quantity in Array for updation in despatch
                For intLoop = 1 To SpChEntry.MaxRows
                    mdblPrevQty(intLoop - 1) = Nothing
                    Call SpChEntry.GetText(GridHeader.Quantity, intLoop, mdblPrevQty(intLoop - 1))
                Next
                SSTab1.SelectedIndex = 0
                If ctlInsurance.Enabled Then ctlInsurance.Focus()
                If cmdConsigneeDetails.Visible Then cmdConsigneeDetails.Enabled = True
                txtContactPerson.Enabled = True : txtContactPerson.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                txtECC.Enabled = True : txtECC.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                txtLST.Enabled = True : txtLST.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                txtAddress1.Enabled = True : txtAddress1.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                txtAddress2.Enabled = True : txtAddress2.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                txtAddress3.Enabled = True : txtAddress3.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                cmdConsigneeOK.Enabled = True : cmdConsigneeCancel.Enabled = True
                txtTCSTaxCode.Enabled = True
                txtTCSTaxCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                lblTCSTaxPerDes.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                cmdHelpTCSTax.Enabled = True
                txtECSSTaxType.Enabled = True
                txtECSSTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                lblECSStax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                CmdECSSTaxType.Enabled = True
                CmdSECSSTaxType.Enabled = True
                txtSECSSTaxType.Enabled = True
                txtSECSSTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                lblSECSStax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                Call SetInvoicecategory(mstrInvType, mstrInvoiceSubType)
                'ADDED AGAINST ISSUE ID : 10533478
                Call SetMaxLengthInSpread(intNoOfDecimal)
                WIPFG_Screen_Setting()
                'END OF CHANGES
                If gblnGSTUnit Then
                    EnableDisable_GST()
                End If

            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD

                        If gblnGSTUnit Then
                            If HSNCODE_GSTTYPE_CHECK_01() = False Then
                                Exit Sub
                            End If

                            If HSNCODE_GSTTYPE_CHECK_02() = False Then
                                Exit Sub
                            End If
                        Else
                            If ValidateTariff_CESS() = False Then Exit Sub
                        End If

                        If Not ValidatebeforeSave("ADD") Then
                            gblnCancelUnload = True
                            gblnFormAddEdit = True
                            Exit Sub
                        End If
                        If QuantityCheck() Then
                            Exit Sub
                        End If
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
                        gStrLocationCode = txtLocationCode.Text
                        gStrCustomerCode = txtCustCode.Text
                        gStrVehicleNo = txtVehNo.Text

                        'If Not SaveData("ADD") Then Exit Sub

                        If gblnGSTUnit Then
                            If Not SaveData_GST("ADD") Then Exit Sub
                        Else
                            If Not SaveData("ADD") Then Exit Sub
                        End If

                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT

                        If gblnGSTUnit Then
                            If HSNCODE_GSTTYPE_CHECK_01() = False Then
                                Exit Sub
                            End If
                            If HSNCODE_GSTTYPE_CHECK_02() = False Then
                                Exit Sub
                            End If
                        Else
                            If ValidateTariff_CESS() = False Then Exit Sub
                        End If

                        If Not ValidatebeforeSave("EDIT") Then
                            gblnCancelUnload = True
                            gblnFormAddEdit = True
                            Exit Sub
                        End If
                        'Check Quantity Schedule
                        If QuantityCheck() Then
                            Exit Sub
                        End If
                        rsInvoiceType = New ClsResultSetDB
                        rsInvoiceType.GetResult("select Invoice_type from Saleschallan_dtl where UNIT_CODE='" & gstrUNITID & "' AND doc_no = " & txtChallanNo.Text)
                        If UCase(rsInvoiceType.GetValue("Invoice_type")) = "REJ" Then
                            If Len(Trim(txtRefNo.Text)) > 0 Then
                                If ItemQtyCaseRejGrin() = False Then
                                    Exit Sub
                                End If
                            End If
                        End If
                        If UCase(rsInvoiceType.GetValue("Invoice_type")) = "EXP" Then
                            If CheckExchangeRate() = False Then
                                Exit Sub
                            End If
                        End If
                        rsInvoiceType = Nothing
                        gStrLocationCode = txtLocationCode.Text
                        gStrCustomerCode = txtCustCode.Text
                        gStrVehicleNo = txtVehNo.Text

                        'If Not SaveData("EDIT") Then Exit Sub

                        If gblnGSTUnit Then
                            If Not SaveData_GST("EDIT") Then Exit Sub
                        Else
                            If Not SaveData("EDIT") Then Exit Sub
                        End If




                End Select
                Call ConfirmWindow(10049, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                gblnCancelUnload = False : gblnFormAddEdit = False
                Call EnableControls(False, Me)
                lblSaltax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                lblSurcharge_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                lblLoadingcharge_per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                SpChEntry.Enabled = True
                SpChEntry.Row = 1 : SpChEntry.Row2 = SpChEntry.MaxRows : SpChEntry.Col = GridHeader.InternalPartNo : SpChEntry.Col2 = SpChEntry.MaxCols
                SpChEntry.BlockMode = True : SpChEntry.Lock = True : SpChEntry.BlockMode = False
                CmbInvType.Visible = False : CmbInvSubType.Visible = False
                lblInvSubType.Visible = False : lblInvType.Visible = False
                txtLocationCode.Enabled = True : txtLocationCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                txtChallanNo.Enabled = True : txtChallanNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                CmdLocCodeHelp.Enabled = True : CmdChallanNo.Enabled = True
                lblDateDes.Text = dtpDateDesc.Text
                dtpDateDesc.Visible = False
                chkDTRemoval.Enabled = True
                Me.CmdGrpChEnt.Revert()
                Me.CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                CmdGrpChEnt.Focus()
                txtTCSTaxCode.Enabled = False
                txtTCSTaxCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                lblTCSTaxPerDes.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                cmdHelpTCSTax.Enabled = False
                txtECSSTaxType.Enabled = False
                txtECSSTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                lblECSStax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                CmdECSSTaxType.Enabled = False
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL
                Call frmMKTTRN0068_KeyPress(Me, New System.Windows.Forms.KeyPressEventArgs(Chr(System.Windows.Forms.Keys.Escape)))
                chkDTRemoval.Enabled = True
                chkDTRemoval.CheckState = System.Windows.Forms.CheckState.Unchecked
                dtpRemoval.Enabled = False
                dtpRemovalTime.Enabled = False
                dtpRemoval.Value = GetServerDate()
                dtpRemovalTime.Value = GetServerDate()
                Me.SSTab1.Controls.Add(Me._SSTab1_TabPage2)
                mIntRecordCount = 0
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE
                If ConfirmWindow(10054, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                    mstrUpdDispatchSql = ""
                    If Len(Trim(txtSRVDINO.Text)) > 0 Then
                        If Val(CStr(Month(ConvertToDate(lblDateDes.Text)))) < 10 Then
                            strMakeDate = Year(ConvertToDate(mstrNagareDate)) & "0" & Month(ConvertToDate(mstrNagareDate))
                        Else
                            strMakeDate = Year(ConvertToDate(mstrNagareDate)) & Month(ConvertToDate(mstrNagareDate))
                        End If
                    Else
                        If Val(CStr(Month(ConvertToDate(lblDateDes.Text)))) < 10 Then
                            strMakeDate = Year(ConvertToDate(lblDateDes.Text)) & "0" & Month(ConvertToDate(lblDateDes.Text))
                        Else
                            strMakeDate = Year(ConvertToDate(lblDateDes.Text)) & Month(ConvertToDate(lblDateDes.Text))
                        End If
                    End If
                    rssalechallan = New ClsResultSetDB
                    salechallan = ""
                    salechallan = "SELECT Invoice_type,SUB_CATEGORY FROM saleschallan_dtl WHERE UNIT_CODE='" & gstrUNITID & "' AND doc_No = "
                    salechallan = salechallan & Val(txtChallanNo.Text)
                    rssalechallan.GetResult(salechallan, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                    If rssalechallan.GetNoRows > 0 Then
                        rssalechallan.MoveFirst()
                        strInvoiceType = rssalechallan.GetValue("Invoice_type")
                    End If
                    If UCase(strInvoiceType) <> "SRC" Then
                        For intLoopCount = 1 To SpChEntry.MaxRows
                            varDrgNo = Nothing
                            varItemCode = Nothing
                            PresQty = Nothing
                            Call Me.SpChEntry.GetText(GridHeader.CustPartNo, intLoopCount, varDrgNo)
                            Call Me.SpChEntry.GetText(GridHeader.InternalPartNo, intLoopCount, varItemCode)
                            Call Me.SpChEntry.GetText(GridHeader.Quantity, intLoopCount, PresQty)
                            If Len(Trim(txtSRVDINO.Text)) > 0 Then
                                mstrUpdDispatchSql = Trim(mstrUpdDispatchSql) & "Update DailyMktSchedule set Despatch_qty ="
                                mstrUpdDispatchSql = mstrUpdDispatchSql & "isnull(Despatch_Qty,0) -  " & Val(PresQty) & ",Schedule_flag =1 "
                                mstrUpdDispatchSql = mstrUpdDispatchSql & " Where UNIT_CODE='" & gstrUNITID & "' AND Account_Code='" & Trim(txtCustCode.Text) & "' and "
                                mstrUpdDispatchSql = mstrUpdDispatchSql & " datepart(yyyy,Trans_Date)='" & Year(ConvertToDate(Trim(mstrNagareDate))) & "'"
                                mstrUpdDispatchSql = mstrUpdDispatchSql & " and datepart(mm,Trans_Date)='" & Month(ConvertToDate(Trim(mstrNagareDate))) & "'"
                                mstrUpdDispatchSql = mstrUpdDispatchSql & " and datepart(dd,Trans_Date)='" & VB.Day(ConvertToDate(Trim(mstrNagareDate))) & "'"
                                mstrUpdDispatchSql = mstrUpdDispatchSql & " and Cust_DrgNo ='" & Trim(varDrgNo) & "' and Item_code = '" & varItemCode & "' and Status =1" & vbCrLf
                                mstrUpdDispatchSql = Trim(mstrUpdDispatchSql) & " Update MonthlyMktSchedule set Despatch_qty ="
                                mstrUpdDispatchSql = mstrUpdDispatchSql & "isnull(Despatch_Qty,0)  - " & Val(PresQty) & ",Schedule_flag =1 "
                                mstrUpdDispatchSql = mstrUpdDispatchSql & " Where UNIT_CODE='" & gstrUNITID & "' AND  Account_Code='" & Trim(txtCustCode.Text) & "' and "
                                mstrUpdDispatchSql = mstrUpdDispatchSql & " Year_Month=" & Val(Trim(strMakeDate)) & ""
                                mstrUpdDispatchSql = mstrUpdDispatchSql & " and Cust_DrgNo ='" & Trim(varDrgNo) & "' and Item_code = '" & varItemCode & "' and Status =1 " & vbCrLf
                            Else
                                mstrUpdDispatchSql = Trim(mstrUpdDispatchSql) & "Update DailyMktSchedule set Despatch_qty ="
                                mstrUpdDispatchSql = mstrUpdDispatchSql & "isnull(Despatch_Qty,0) -  " & Val(PresQty) & ",Schedule_flag =1 "
                                mstrUpdDispatchSql = mstrUpdDispatchSql & " Where UNIT_CODE='" & gstrUNITID & "' AND Account_Code='" & Trim(txtCustCode.Text) & "' and "
                                mstrUpdDispatchSql = mstrUpdDispatchSql & " datepart(yyyy,Trans_Date)='" & Year(ConvertToDate(Trim(lblDateDes.Text))) & "'"
                                mstrUpdDispatchSql = mstrUpdDispatchSql & " and datepart(mm,Trans_Date)='" & Month(ConvertToDate(Trim(lblDateDes.Text))) & "'"
                                mstrUpdDispatchSql = mstrUpdDispatchSql & " and datepart(dd,Trans_Date)='" & VB.Day(ConvertToDate(Trim(lblDateDes.Text))) & "'"
                                mstrUpdDispatchSql = mstrUpdDispatchSql & " and Cust_DrgNo ='" & Trim(varDrgNo) & "' and Item_code = '" & varItemCode & "' and Status =1" & vbCrLf
                                mstrUpdDispatchSql = Trim(mstrUpdDispatchSql) & " Update MonthlyMktSchedule set Despatch_qty ="
                                mstrUpdDispatchSql = mstrUpdDispatchSql & "isnull(Despatch_Qty,0)  - " & Val(PresQty) & ",Schedule_flag =1 "
                                mstrUpdDispatchSql = mstrUpdDispatchSql & " Where UNIT_CODE='" & gstrUNITID & "' AND Account_Code='" & Trim(txtCustCode.Text) & "' and "
                                mstrUpdDispatchSql = mstrUpdDispatchSql & " Year_Month=" & Val(Trim(strMakeDate)) & ""
                                mstrUpdDispatchSql = mstrUpdDispatchSql & " and Cust_DrgNo ='" & Trim(varDrgNo) & "' and Item_code = '" & varItemCode & "' and Status =1 " & vbCrLf
                            End If
                        Next
                    End If
                    Call DeleteRecords()
                    mP_Connection.BeginTrans()
                    '10736222
                    Dim objCmd As New ADODB.Command

                    With objCmd
                        .ActiveConnection = mP_Connection
                        .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                        .CommandText = "USP_SAVE_CT2_INVOICE_KNOCKOFFDTL"
                        .CommandTimeout = 0
                        .Parameters.Append(.CreateParameter("@MODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, "D"))
                        .Parameters.Append(.CreateParameter("@UNIT_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, gstrUNITID))
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
                    If Len(Trim(mstrUpdDispatchSql)) > 0 Then
                        mP_Connection.Execute(mstrUpdDispatchSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End If
                    mP_Connection.CommitTrans()
                    ''10533478
                    If gBlnWIPFGProcess AndAlso WIP_FG_Customer(txtCustCode.Text.Trim) Then
                        If mblnskipdacinvoicebincheck = False Then
                            Dim svcWIPFGInv As New eMPRO_WIP_FG_INVOICE_COMPLETION
                            Dim strResponse As String = String.Empty
                            Dim intRow As Integer
                            Dim blnItemExists As Boolean
                            If IsRecordExists("SELECT TOP 1 1 FROM WIP_FG_PICKLIST_INV_DTL  WHERE TMP_INVOICENO = " & txtChallanNo.Text.Trim & " AND UNIT_CODE ='" & gstrUNITID & "' ") Then
                                strResponse = svcWIPFGInv.ProcessInvoice(gstrUNITID, "D", Trim(txtChallanNo.Text))
                                If strResponse.ToUpper() <> "SUCCESS|" Then
                                    WIP_FG_SAVE_EXCEPTION_LOG(Convert.ToInt32(txtChallanNo.Text), "svcWIPFGInv.ProcessInvoice", strResponse, gstrUNITID, "D", Trim(txtChallanNo.Text), "")
                                End If
                            End If
                            svcWIPFGInv.Dispose()
                        End If
                    End If
                    ''
                    Call EnableControls(False, Me, True)
                    txtLocationCode.Enabled = True
                    txtLocationCode.BackColor = System.Drawing.Color.White
                    CmdLocCodeHelp.Enabled = True
                    txtChallanNo.Enabled = True
                    txtChallanNo.BackColor = System.Drawing.Color.White
                    CmdChallanNo.Enabled = True
                    mIntRecordCount = 0

                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT
                    If Trim(txtChallanNo.Text) = "" Then
                        MsgBox("Please select a Challan Number first.", MsgBoxStyle.Information, "eMpro")
                        Exit Sub
                    End If
                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                    If CBool(Find_Value("select Enagare_TextPrinting from sales_parameter WHERE UNIT_CODE='" & gstrUNITID & "'")) Then
                        Call PrintingInvoice()
                    Else
                        'With Me.crReportInvoicePrinting
                        '    .Reset()
                        '    .DiscardSavedData = True
                        '    .WindowShowSearchBtn = True
                        '    .WindowShowPrintSetupBtn = True
                        '    .WindowShowPrintBtn = True
                        '    .WindowShowExportBtn = True
                        '    .WindowState = Crystal.WindowStateConstants.crptMaximized
                        '    .Connect = gstrREPORTCONNECT
                        ' Call PrintingInvoiceRPT()
                        '    .WindowTitle = Me.ctlFormHeader1.HeaderString()
                        '    On Error Resume Next
                        '    .Action = 1
                        'End With
                        'Dim RdAddSold As ReportDocument
                        'Dim Frm As New eMProCrystalReportViewer
                        RdAddSold = Frm.GetReportDocument()
                        Frm.ReportHeader = Me.ctlFormHeader1.HeaderString()
                        Call PrintingInvoiceRPT()
                        On Error Resume Next
                        Frm.Show()
                    End If
                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                    Me.Close()
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub cmdhelpSRVDI_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles cmdhelpSRVDI.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then
            CmdGrpChEnt.Focus()
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub cmdHelpTCSTax_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelpTCSTax.Click
        '****************************************************
        'Created By     -  Nisha
        'Description    -  To Display Help From SaleTax Master
        '****************************************************
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
                    salechallan = salechallan & " and a.Location_code = b.Location_code and a.Invoice_type = b.invoice_type and a.sub_category = b.Sub_type AND a.UNIT_CODE=b.UNIT_CODE AND a.UNIT_CODE='" & gstrUNITID & "'"
                    rssalechallan.GetResult(salechallan, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                    If rssalechallan.GetNoRows > 0 Then
                        rssalechallan.MoveFirst()
                        strInvoiceType = rssalechallan.GetValue("Description")
                    End If
                Else
                    strInvoiceType = CmbInvType.Text
                End If
                If Len(Me.txtTCSTaxCode.Text) = 0 Then 'To check if There is No Text Then Show All Help
                    strHelp = ShowList(1, (txtTCSTaxCode.MaxLength), "", "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND (Tx_TaxeID='TCS') AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))")
                    If strHelp = "-1" Then 'If No Record Exists In The Table
                        Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Exit Sub
                    Else
                        txtTCSTaxCode.Text = strHelp
                    End If
                Else
                    strHelp = ShowList(1, (txtTCSTaxCode.MaxLength), txtTCSTaxCode.Text, "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND (Tx_TaxeID='TCS') AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))")
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
    Private Sub Cmditems_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Cmditems.Click
        '****************************************************
        'Created By     -  Nisha
        'Description    -  Display Another Form for User To Select Item Code >From CustOrd_Dtl
        '                  And After Selecting Item Code Select Data From Sales_Dtl and Display
        '                  That Details In The Spread
        '****************************************************
        On Error GoTo ErrHandler
        Dim rssalechallan As ClsResultSetDB
        Dim salechallan As String
        Dim strItemNotIn As String
        Dim varItemCode As Object
        Dim varKanbanNo As Object
        Dim rsSaleConf As ClsResultSetDB
        Dim strStockLocation As String
        Dim rsCurrencyType As ClsResultSetDB
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        Dim strMain() As String
        Dim strDet() As String
        Dim intCount As Short
        Dim varItemAlready As Object
        Dim strSQL As String
        Dim blntcscheck As Boolean = False
        With Me.SpChEntry
            If CmdGrpChEnt.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                .MaxRows = 1
                .Row = 1 : .Row2 = .MaxRows : .Col = GridHeader.InternalPartNo : .Col2 = .MaxCols : .BlockMode = True : .Text = "" : .BlockMode = False
            End If
        End With
        Dim blnCurrentInvoice As Boolean
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                blnCurrentInvoice = CBool(Find_Value("select isnull(invoiceAgainstMultipleSo,0) from salesChallan_dtl where Location_code='" & Trim(txtLocationCode.Text) & "' and doc_no='" & Trim(txtChallanNo.Text) & "' AND UNIT_CODE='" & gstrUNITID & "'"))
                If blnInvoiceAgainstMultipleSO And blnCurrentInvoice Then
                    If SpChEntry.MaxRows > 0 Then
                        intMaxLoop = SpChEntry.MaxRows
                        strItemNotIn = ""
                        For intLoopCounter = 1 To intMaxLoop
                            With SpChEntry
                                .Row = intLoopCounter : .Col = GridHeader.Rate
                                If .Text <> "D" Then
                                    varItemCode = Nothing
                                    varKanbanNo = Nothing
                                    Call .GetText(GridHeader.CustPartNo, intLoopCounter, varItemCode)
                                    Call .GetText(GridHeader.srvdino, intLoopCounter, varKanbanNo)
                                    strItemNotIn = strItemNotIn & varItemCode & "|" & varKanbanNo & "^"
                                End If
                            End With
                        Next
                    End If
                    If Len(Trim(strItemNotIn)) > 0 Then
                        mstrItemCode = frmMKTTRN0068a.SelectDatafromItem_Mst(strItemNotIn, CInt(Trim(txtChallanNo.Text)))
                    Else
                        mstrItemCode = frmMKTTRN0068a.SelectDatafromItem_Mst()
                    End If
                    With SpChEntry
                        strMain = Split(mstrItemCode, "^")
                        SpChEntry.MaxRows = 0
                        mIntRecordCount = 0
                        For intLoopCounter = 0 To UBound(strMain) - 1
                            strDet = Split(strMain(intLoopCounter), "|")
                            If intLoopCounter = 0 Then
                                txtCustCode.Text = strDet(11)
                                txtCustCode_Validating(txtCustCode, New System.ComponentModel.CancelEventArgs(True))
                            End If
                            mstrRefNo = strDet(5)
                            mstrAmmNo = strDet(6)
                            mstrItemCode = "'" & strDet(2) & "'"
                            mstrQuantity = strDet(4)
                            mstrSRVDINo = strDet(0)
                            mstrSRVLocation = strDet(9)
                            mstrUSLoc = strDet(10)
                            mstrSchTime = strDet(8)
                            Call displayDeatilsfromCustOrdHdrandDtl()
                        Next
                        Call SpChEntry_Change(SpChEntry, New AxFPSpreadADO._DSpreadEvents_ChangeEvent(5, 1))
                    End With

                    Exit Sub
                End If
                'Code Ends here
                rssalechallan = New ClsResultSetDB
                salechallan = ""
                salechallan = "SELECT Invoice_type,SUB_CATEGORY FROM saleschallan_dtl WHERE UNIT_CODE='" & gstrUNITID & "' AND doc_No = "
                salechallan = salechallan & Val(txtChallanNo.Text)
                rssalechallan.GetResult(salechallan, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                If rssalechallan.GetNoRows > 0 Then
                    rssalechallan.MoveFirst()
                    strInvType = rssalechallan.GetValue("Invoice_type")
                    strInvSubType = rssalechallan.GetValue("sub_category")
                End If
                strStockLocation = StockLocationSalesConf(strInvType, strInvSubType, "TYPE")
                If Len(Trim(strStockLocation)) > 0 Then
                    If (UCase(strInvType) = "INV") Or (UCase(strInvType) = "EXP") Or (UCase(strInvType) = "SRC") Then
                        mstrItemCode = frmMKTTRN0021.SelectDatafromsaleDtl(Trim(txtChallanNo.Text))
                        If Len(Trim(mstrItemCode)) = 0 Then SpChEntry.MaxRows = 0 : frmMKTTRN0021.Close()
                    Else
                        mstrItemCode = frmMKTTRN0021.SelectDatafromsaleDtl(Trim(txtChallanNo.Text))
                        If Len(Trim(mstrItemCode)) = 0 Then SpChEntry.MaxRows = 0 : frmMKTTRN0021.Close()
                    End If
                Else
                    MsgBox("Please Define Stock Location in Sales Conf")
                    Exit Sub
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
                                txtTCSTaxCode.Enabled = True : txtTCSTaxCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : cmdHelpTCSTax.Enabled = True
                            End If
                        Else
                            txtTCSTaxCode.Enabled = False : txtTCSTaxCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdHelpTCSTax.Enabled = False : txtTCSTaxCode.Text = ""
                        End If
                    End If
                End If

            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                If blnInvoiceAgainstMultipleSO Then
                    If SpChEntry.MaxRows > 0 Then
                        intMaxLoop = SpChEntry.MaxRows
                        strItemNotIn = ""
                        For intLoopCounter = 1 To intMaxLoop
                            With SpChEntry
                                .Row = intLoopCounter : .Col = GridHeader.Rate
                                If .Text <> "D" Then
                                    varItemCode = Nothing
                                    varKanbanNo = Nothing
                                    Call .GetText(GridHeader.CustPartNo, intLoopCounter, varItemCode)
                                    Call .GetText(GridHeader.srvdino, intLoopCounter, varKanbanNo)
                                    strItemNotIn = strItemNotIn & varItemCode & "|" & varKanbanNo & "^"
                                End If
                            End With
                        Next
                    End If
                    If Len(Trim(strItemNotIn)) > 0 Then
                        frmMKTTRN0068a.ShowDialog()
                        mstrItemCode = mstrItemText
                    Else
                        frmMKTTRN0068a.ShowDialog()
                        mstrItemCode = mstrItemText
                        If Len(mstrItemCode) = 0 Then Exit Sub
                    End If
                    With SpChEntry
                        strMain = Split(mstrItemCode, "^")
                        SpChEntry.MaxRows = 0
                        mIntRecordCount = 0
                        For intLoopCounter = 0 To UBound(strMain) - 1
                            strDet = Split(strMain(intLoopCounter), "|")
                            If intLoopCounter = 0 Then
                                txtCustCode.Text = strDet(11)
                                txtCustCode_Validating(txtCustCode, New System.ComponentModel.CancelEventArgs(True))
                            End If
                            mstrRefNo = strDet(5)
                            mstrAmmNo = strDet(6)
                            mstrItemCode = "'" & strDet(2) & "'"
                            mstrQuantity = strDet(4)
                            mstrSRVDINo = strDet(0)
                            mstrSRVLocation = strDet(9)
                            mstrUSLoc = strDet(10)
                            mstrSchTime = strDet(8)
                            Call displayDeatilsfromCustOrdHdrandDtl()
                        Next
                        Call SpChEntry_Change(SpChEntry, New AxFPSpreadADO._DSpreadEvents_ChangeEvent(5, 1))
                    End With
                    SSTab1.SelectedIndex = 1
                    With SpChEntry
                        .Row = 1
                        .Col = 5
                        .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                        .Focus()
                    End With
                    Exit Sub
                End If
                If UCase(CStr(Trim(CmbInvType.Text))) = "NORMAL INVOICE" Or UCase(CStr(Trim(CmbInvType.Text))) = "JOBWORK INVOICE" Or UCase(CStr(Trim(CmbInvType.Text))) = "EXPORT INVOICE" Or UCase(CStr(Trim(CmbInvType.Text))) = "SERVICE INVOICE" Then
                    If UCase(Trim(CmbInvSubType.Text)) <> "SCRAP" Then
                        If Len(Trim(txtRefNo.Text)) = 0 Then
                            Call ConfirmWindow(10240, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                            If txtRefNo.Enabled Then txtRefNo.Focus()
                            Exit Sub
                        ElseIf Len(Trim(txtAmendNo.Text)) = 0 Then

                            If OriginalRefNoOVER(Trim(txtRefNo.Text)) Then
                                MsgBox("Enter Amendment No.", MsgBoxStyle.Information, "eMPro")
                                If txtAmendNo.Enabled Then txtAmendNo.Focus() : Exit Sub
                            Else
                                mstrRefNo = Trim(txtRefNo.Text)
                                mstrAmmNo = ""
                            End If
                        Else

                            mstrRefNo = Trim(txtRefNo.Text)
                            mstrAmmNo = Trim(txtAmendNo.Text)
                        End If
                    End If
                End If
                If (Trim(CmbInvType.Text) = "JOBWORK INVOICE") Then
                    If blnFIFO = False Then
                        If Len(Trim(mstrRGP)) = 0 Then
                            MsgBox("First Select RGP No. ", MsgBoxStyle.OkOnly, "eMPro")
                            Call AddDataTolstRGPs()
                            fraRGPs.Visible = True
                            Exit Sub
                        End If
                    End If
                End If
                If SpChEntry.MaxRows > 0 Then
                    intMaxLoop = SpChEntry.MaxRows
                    strItemNotIn = ""
                    For intLoopCounter = 1 To intMaxLoop
                        With SpChEntry
                            varItemCode = Nothing
                            Call .GetText(1, intLoopCounter, varItemCode)
                            If Len(Trim(strItemNotIn)) > 0 Then

                                strItemNotIn = Trim(strItemNotIn) & ",'" & Trim(varItemCode) & "'"
                            Else

                                strItemNotIn = "'" & Trim(varItemCode) & "'"
                            End If
                        End With
                    Next
                End If

                If UCase(CStr(Trim(CmbInvType.Text))) = "NORMAL INVOICE" Or UCase(CStr(Trim(CmbInvType.Text))) = "EXPORT INVOICE" Or UCase(CStr(Trim(CmbInvType.Text))) = "SERVICE INVOICE" Then
                    strStockLocation = StockLocationSalesConf((CmbInvType.Text), (CmbInvSubType.Text), "DESCRIPTION")
                    If Len(Trim(strStockLocation)) > 0 Then
                        If CBool(UCase(CStr(Trim(CmbInvSubType.Text) = "SCRAP"))) Then
                            If Len(Trim(strItemNotIn)) > 0 Then
                                mstrItemCode = frmMKTTRN0021.SelectDatafromItem_Mst(Trim(CmbInvType.Text), Trim(CmbInvSubType.Text), strStockLocation, , strItemNotIn, SpChEntry.MaxRows)
                            Else
                                mstrItemCode = frmMKTTRN0021.SelectDatafromItem_Mst(Trim(CmbInvType.Text), Trim(CmbInvSubType.Text), strStockLocation)
                            End If
                        Else
                            If Len(Trim(strItemNotIn)) > 0 Then
                                mstrItemCode = frmMKTTRN0021.SelectDataFromCustOrd_Dtl(Trim(txtCustCode.Text), Trim(txtRefNo.Text), mstrAmmNo, Trim(CmbInvSubType.Text), Trim(CmbInvType.Text), strStockLocation, strItemNotIn, SpChEntry.MaxRows)
                            Else
                                mstrItemCode = frmMKTTRN0021.SelectDataFromCustOrd_Dtl(Trim(txtCustCode.Text), Trim(txtRefNo.Text), mstrAmmNo, Trim(CmbInvSubType.Text), Trim(CmbInvType.Text), strStockLocation)
                            End If
                        End If
                    Else
                        MsgBox("Please Define Stock Location in Sales Conf", MsgBoxStyle.Information, "eMPro")
                        Exit Sub
                    End If
                    If Len(Trim(mstrItemCode)) = 0 Then SpChEntry.MaxRows = 0
                ElseIf (Trim(CmbInvType.Text) = "JOBWORK INVOICE") Then
                    strStockLocation = StockLocationSalesConf((CmbInvType.Text), (CmbInvSubType.Text), "DESCRIPTION")
                    If Len(Trim(strStockLocation)) > 0 Then
                        If Len(Trim(strItemNotIn)) > 0 Then
                            mstrItemCode = frmMKTTRN0021.SelectDataFromCustOrd_Dtl(Trim(txtCustCode.Text), Trim(txtRefNo.Text), mstrAmmNo, Trim(CmbInvSubType.Text), Trim(CmbInvType.Text), strStockLocation, strItemNotIn, SpChEntry.MaxRows)
                        Else
                            mstrItemCode = frmMKTTRN0021.SelectDataFromCustOrd_Dtl(Trim(txtCustCode.Text), Trim(txtRefNo.Text), mstrAmmNo, Trim(CmbInvSubType.Text), Trim(CmbInvType.Text), strStockLocation)
                        End If
                    Else
                        MsgBox("Please Define Stock Location in Sales Conf", MsgBoxStyle.Information, "eMPro")
                        Exit Sub
                    End If
                    If Len(Trim(mstrItemCode)) = 0 Then SpChEntry.MaxRows = 0
                Else
                    rsSaleConf = New ClsResultSetDB
                    rsSaleConf.GetResult("select Stock_Location From saleconf where UNIT_CODE='" & gstrUNITID & "' AND Description ='" & Trim(CmbInvType.Text) & "' and sub_type_description ='" & Trim(CmbInvSubType.Text) & "' AND Location_Code ='" & Trim(txtLocationCode.Text) & "' and (fin_start_date <= getdate() and fin_end_date >= getdate())")
                    If ((Len(Trim(rsSaleConf.GetValue("Stock_Location"))) = 0) Or (Trim(CStr(rsSaleConf.GetValue("Stock_Location") = "Unknown")))) Then
                        MsgBox("Please Select Stock Location in SalesConf first", MsgBoxStyle.Information, "eMPro")
                        If Cmditems.Enabled Then Cmditems.Focus()
                        Exit Sub
                    End If
                    If CBool(UCase(CStr(Trim(CmbInvType.Text) = "REJECTION"))) Then
                        If Len(Trim(txtRefNo.Text)) > 0 Then
                            If Len(Trim(strItemNotIn)) > 0 Then
                                mstrItemCode = frmMKTTRN0021.AddDataFromGrinDtl(Trim(txtCustCode.Text), CDbl(Trim(txtRefNo.Text)), rsSaleConf.GetValue("Stock_Location"), SpChEntry.MaxRows, strItemNotIn)
                            Else
                                mstrItemCode = frmMKTTRN0021.AddDataFromGrinDtl(Trim(txtCustCode.Text), CDbl(Trim(txtRefNo.Text)), rsSaleConf.GetValue("Stock_Location"))
                            End If
                        Else
                            If Len(Trim(strItemNotIn)) > 0 Then
                                mstrItemCode = frmMKTTRN0021.SelectDatafromItem_Mst(Trim(CmbInvType.Text), Trim(CmbInvSubType.Text), rsSaleConf.GetValue("Stock_Location"), Trim(txtCustCode.Text), strItemNotIn, SpChEntry.MaxRows)
                            Else
                                mstrItemCode = frmMKTTRN0021.SelectDatafromItem_Mst(Trim(CmbInvType.Text), Trim(CmbInvSubType.Text), rsSaleConf.GetValue("Stock_Location"), Trim(txtCustCode.Text))
                            End If
                        End If
                    Else
                        If Len(Trim(strItemNotIn)) > 0 Then
                            mstrItemCode = frmMKTTRN0021.SelectDatafromItem_Mst(Trim(CmbInvType.Text), Trim(CmbInvSubType.Text), rsSaleConf.GetValue("Stock_Location"), Trim(txtCustCode.Text), strItemNotIn, SpChEntry.MaxRows)
                        Else
                            mstrItemCode = frmMKTTRN0021.SelectDatafromItem_Mst(Trim(CmbInvType.Text), Trim(CmbInvSubType.Text), rsSaleConf.GetValue("Stock_Location"), Trim(txtCustCode.Text))
                        End If
                    End If
                    If Len(Trim(mstrItemCode)) = 0 And Len(Trim(strItemNotIn)) = 0 Then
                        SpChEntry.MaxRows = 0
                    Else
                        If Len(Trim(mstrItemCode)) = 0 Then
                        End If
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
                                txtTCSTaxCode.Enabled = True : txtTCSTaxCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : cmdHelpTCSTax.Enabled = True
                            End If
                        Else
                            txtTCSTaxCode.Enabled = False : txtTCSTaxCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdHelpTCSTax.Enabled = False : txtTCSTaxCode.Text = ""
                        End If
                    End If
                End If

        End Select
        Dim intDecimalPlace As Short
        Dim strCurrency As String
        If Len(mstrItemCode) > 0 Then
            mstrItemCode = Mid(mstrItemCode, 1, Len(mstrItemCode) - 1)
            Select Case Me.CmdGrpChEnt.Mode
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                    rsCurrencyType = New ClsResultSetDB
                    rsCurrencyType.GetResult("Select Currency_code from saleschallan_dtl where UNIT_CODE='" & gstrUNITID & "' AND doc_No = " & Val(txtChallanNo.Text))
                    If rsCurrencyType.GetNoRows > 0 Then
                        rsCurrencyType.MoveFirst()
                        strCurrency = rsCurrencyType.GetValue("Currency_code")
                    End If
                    intDecimalPlace = ToGetDecimalPlaces(strCurrency)
                    If intDecimalPlace < 2 Then
                        intDecimalPlace = 2
                    End If
                    DisplayDetailsInSpread(strCurrency) 'Procedure Call To Select Data >From Sales_Dtl
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                    Call displayDeatilsfromCustOrdHdrandDtl()
                    System.Windows.Forms.Application.DoEvents()
            End Select
            If Me.CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                If CDbl(Trim(txtChallanNo.Text)) > 99000000 Then

                    Me.CmdGrpChEnt.Enabled(1) = True

                    Me.CmdGrpChEnt.Enabled(2) = True
                End If
            End If
            If Me.CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                SSTab1.SelectedIndex = 1

                If ctlInsurance.Enabled Then

                    If ctlInsurance.Enabled Then ctlInsurance.Focus()
                Else
                    System.Windows.Forms.SendKeys.Send("{tab}")
                End If
            Else
                Me.CmdGrpChEnt.Focus()
            End If
        Else
            frmMKTTRN0021.Close()
        End If
        'Set Cell Type In Spread
        Call ChangeCellTypeStaticText()
        If Me.CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            SSTab1.SelectedIndex = 1

            If ctlInsurance.Enabled Then ctlInsurance.Focus()
        Else
            Me.CmdGrpChEnt.Focus()
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub cmdLoadinfChageHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdLoadinfChageHelp.Click
        '****************************************************
        'Created By     -  Nisha
        'Description    -  To Display Help From SaleTax Master
        '****************************************************
        Dim strHelp As String
        On Error GoTo ErrHandler
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                If Len(txtLoadingTaxType.Text) = 0 Then 'To check if There is No Text Then Show All Help
                    strHelp = ShowList(1, (txtLoadingTaxType.MaxLength), "", "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND (Tx_TaxeID='LDT') AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))")
                    If strHelp = "-1" Then 'If No Record Exists In The Table
                        Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Exit Sub
                    Else
                        txtLoadingTaxType.Text = strHelp
                    End If
                Else
                    strHelp = ShowList(1, (txtLoadingTaxType.MaxLength), txtLoadingTaxType.Text, "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND (Tx_TaxeID='LDT') AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))")
                    If strHelp = "-1" Then 'If No Record Exists In The Table
                        Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Exit Sub
                    Else
                        txtLoadingTaxType.Text = strHelp
                    End If
                End If
                Call txtLoadingTaxType_Validating(txtLoadingTaxType, New System.ComponentModel.CancelEventArgs(False))
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub CmdLocCodeHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdLocCodeHelp.Click
        '****************************************************
        'Created By     -  Nisha
        'Description    -  To Display Help From Location Master
        '****************************************************
        On Error GoTo ErrHandler
        Dim strHelp As String
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                If Len(Me.txtLocationCode.Text) = 0 Then 'To check if There is No Text Then Show All Help
                      strHelp = ShowList(1, (txtLocationCode.MaxLength), "", "s.Location_Code", "l.Description", "Location_mst l,SaleConf s", "and s.Location_Code=l.Location_Code AND s.UNIT_CODE=l.UNIT_CODE", , , , , , "s.UNIT_CODE")
                    If strHelp = "-1" Then 'If No Record Exists In The Table
                        Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Exit Sub
                    Else
                        txtLocationCode.Text = strHelp
                     End If
                Else
                    'To Display All Possible Help Starting With Text in TextField
                      strHelp = ShowList(1, (txtLocationCode.MaxLength), txtLocationCode.Text, "s.Location_Code", "l.Description", "Location_mst l,SaleConf s", "and s.Location_Code=l.Location_Code AND s.UNIT_CODE=l.UNIT_CODE", , , , , , "s.UNIT_CODE")
                    If strHelp = "-1" Then 'If No Record Exists In The Table
                        Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Exit Sub
                    Else
                        txtLocationCode.Text = strHelp
                     End If
                End If
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
        End Select
        'Procedure Call To Select The Location Code Description
        Call SelectDescriptionForField("Description", "Location_Code", "Location_Mst", lblLocCodeDes, (txtLocationCode.Text))
        Call txtLocationCode_Validating(txtLocationCode, New System.ComponentModel.CancelEventArgs(False))
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub CmdRefNoHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdRefNoHelp.Click
        '****************************************************
        'Created By     -  Nisha
        'Description    -  To Display Details Of Customer Order
        '****************************************************
        Dim frmMKTTRN0020 As New frmMKTTRN0020
        On Error GoTo ErrHandler
        If Len(txtCustCode.Text) = 0 Then
            Call ConfirmWindow(10416, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
            txtCustCode.Focus()
            Exit Sub
        End If
        Dim strRefAmm As String
        Dim intPos As Short
        If UCase(CmbInvType.Text) <> "REJECTION" Then
            strRefAmm = frmMKTTRN0020.SelectDataFromCustOrd_Dtl((txtCustCode.Text), (CmbInvType.Text))
            ' frmMKTTRN0020.ShowDialog()
        Else
            strRefAmm = frmMKTTRN0020.SelectDataFromGrinDtl((txtCustCode.Text))
         End If
        If Len(strRefAmm) > 0 Then
            intPos = InStr(1, Trim(strRefAmm), ",", CompareMethod.Text)
            mstrRefNo = Mid(Trim(strRefAmm), 2, intPos - 3)
            mstrAmmNo = Mid(strRefAmm, intPos + 2, ((Len(Trim(strRefAmm))) - intPos) - 2)
            txtRefNo.Text = Trim(mstrRefNo)
            txtAmendNo.Text = mstrAmmNo 'Added By   -   Nitin
          
            txtSRVDINO.Focus()
            '***
         Else
         
            txtSRVDINO.Focus()
         End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub cmdRGPCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdRGPCancel.Click
        mstrRGP = ""
        lblRGPDes.Text = ""
        fraRGPs.Visible = False
        txtCustCode.Focus()
    End Sub
    Private Sub cmdRGPOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdRGPOK.Click
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        mstrRGP = ""
        intMaxLoop = lvwRGPs.Items.Count
        For intLoopCounter = 1 To intMaxLoop
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
    End Sub
    Private Sub CmdSaleTaxType_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdSaleTaxType.Click
        '****************************************************
        'Created By     -  Nisha
        'Description    -  To Display Help From SaleTax Master
        '****************************************************
        On Error GoTo ErrHandler
        Dim strHelp As String
        Dim rssalechallan, rsadditionaltax, rsadditionalsurcharge, rsadditionalVAT As ClsResultSetDB
        Dim salechallan, strsql As String
        Dim strInvoiceType As Object
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                If CmdGrpChEnt.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                    txtSurchargeTaxType.Text = ""
                    rssalechallan = New ClsResultSetDB
                    salechallan = ""
                    salechallan = "select b.Description, b.Sub_type_Description from SalesChallan_dtl a,saleconf b where doc_no = " & Trim(txtChallanNo.Text)
                    salechallan = salechallan & " and a.Location_code = b.Location_code and a.Invoice_type = b.invoice_type and a.sub_category = b.Sub_type AND a.UNIT_CODE=b.UNIT_CODE AND a.UNIT_CODE='" & gstrUNITID & "'"
                    rssalechallan.GetResult(salechallan, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                    If rssalechallan.GetNoRows > 0 Then
                        rssalechallan.MoveFirst()
                        
                        
                        strInvoiceType = rssalechallan.GetValue("Description")
                    End If
                Else
                    
                    strInvoiceType = CmbInvType.Text
                End If
                 If Len(Me.txtSaleTaxType.Text) = 0 Then 'To check if There is No Text Then Show All Help
                    
                    If UCase(strInvoiceType) <> "SERVICE INVOICE" Then
                        
                        
                        strHelp = ShowList(1, (txtSaleTaxType.MaxLength), "", "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND (Tx_TaxeID='CST' OR Tx_TaxeID='LST' OR Tx_TaxeID='VAT') AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))")
                    Else
                        
                        
                        strHelp = ShowList(1, (txtSaleTaxType.MaxLength), "", "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND (Tx_TaxeID='SRT') AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))")
                    End If
                    If strHelp = "-1" Then 'If No Record Exists In The Table
                        Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Exit Sub
                    Else
                        txtSaleTaxType.Text = strHelp
                          If UCase(Trim(GetPlantName)) = "MATM" And UCase(strInvoiceType) = "NORMAL INVOICE" Then
                            strsql = " select TxRt_Rate_No,TxRt_Percentage from Gen_TaxRate where (Tx_TaxeID='CST' OR Tx_TaxeID='LST') and txrt_percentage > 2.0 and TxRt_Rate_No='" & txtSaleTaxType.Text & "' AND UNIT_CODE='" & gstrUNITID & "'"
                            rsadditionaltax = New ClsResultSetDB
                            rsadditionaltax.GetResult(strsql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                             If rsadditionaltax.GetNoRows > 0 Then
 rsadditionalsurcharge = New ClsResultSetDB
                                strsql = " select TxRt_Rate_No,TxRt_Percentage from Gen_TaxRate where Tx_TaxeID='SsT' and txrt_percentage=5.0 AND UNIT_CODE='" & gstrUNITID & "' AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"
 rsadditionalsurcharge.GetResult(strsql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
 If rsadditionalsurcharge.GetNoRows > 0 Then
     txtSurchargeTaxType.Text = rsadditionalsurcharge.GetValue("TxRt_Rate_No")
     lblSurcharge_Per.Text = rsadditionalsurcharge.GetValue("TxRt_Percentage")
 End If
 rsadditionalsurcharge.ResultSetClose()
 rsadditionalsurcharge = Nothing
                            End If
                             rsadditionaltax.ResultSetClose()
                            rsadditionaltax = Nothing
                            strsql = " select TxRt_Rate_No,TxRt_Percentage from Gen_TaxRate where (Tx_TaxeID='VAT') and TxRt_Rate_No='" & txtSaleTaxType.Text & "' AND UNIT_CODE='" & gstrUNITID & "' "
                            rsadditionalVAT = New ClsResultSetDB
                            rsadditionalVAT.GetResult(strsql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                             If rsadditionalVAT.GetNoRows > 0 Then
 rsadditionalsurcharge = New ClsResultSetDB
                                strsql = " select TxRt_Rate_No,TxRt_Percentage from Gen_TaxRate where Tx_TaxeID='SsT' and txrt_percentage=5.0 AND UNIT_CODE='" & gstrUNITID & "' AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"
 rsadditionalsurcharge.GetResult(strsql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
 If rsadditionalsurcharge.GetNoRows > 0 Then
     txtSurchargeTaxType.Text = rsadditionalsurcharge.GetValue("TxRt_Rate_No")
     lblSurcharge_Per.Text = rsadditionalsurcharge.GetValue("TxRt_Percentage")
 End If
 rsadditionalsurcharge.ResultSetClose()
 rsadditionalsurcharge = Nothing
                            End If
                            rsadditionalVAT.ResultSetClose()
                            rsadditionalVAT = Nothing
                         End If
                     End If
                Else
                    
                    
                    If UCase(strInvoiceType) <> "SERVICE INVOICE" Then
                        
                        
                        strHelp = ShowList(1, (txtSaleTaxType.MaxLength), txtSaleTaxType.Text, "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND (Tx_TaxeID='CST' OR Tx_TaxeID='LST' OR Tx_TaxeID='VAT') AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))")
                    Else
                        
                        
                        strHelp = ShowList(1, (txtSaleTaxType.MaxLength), txtSaleTaxType.Text, "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND (Tx_TaxeID='SRT') AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))")
                    End If
                    If strHelp = "-1" Then 'If No Record Exists In The Table
                        Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Exit Sub
                    Else
                        txtSaleTaxType.Text = strHelp
                        If UCase(Trim(GetPlantName)) = "MATM" And UCase(strInvoiceType) = "NORMAL INVOICE" Then
                            strsql = " select TxRt_Rate_No,TxRt_Percentage from Gen_TaxRate where (Tx_TaxeID='CST' OR Tx_TaxeID='LST') and txrt_percentage > 2.0 and TxRt_Rate_No='" & txtSaleTaxType.Text & "' AND UNIT_CODE='" & gstrUNITID & "'"
                            rsadditionaltax = New ClsResultSetDB
                            rsadditionaltax.GetResult(strsql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                             If rsadditionaltax.GetNoRows > 0 Then
 rsadditionalsurcharge = New ClsResultSetDB
                                strsql = " select TxRt_Rate_No,TxRt_Percentage from Gen_TaxRate where Tx_TaxeID='SST' and txrt_percentage=5.0 AND UNIT_CODE='" & gstrUNITID & "' AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"
 rsadditionalsurcharge.GetResult(strsql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
 If rsadditionalsurcharge.GetNoRows > 0 Then
     txtSurchargeTaxType.Text = rsadditionalsurcharge.GetValue("TxRt_Rate_No")
     lblSurcharge_Per.Text = rsadditionalsurcharge.GetValue("TxRt_Percentage")
 End If
 rsadditionalsurcharge.ResultSetClose()
 rsadditionalsurcharge = Nothing
                            End If
                             rsadditionaltax.ResultSetClose()
                            rsadditionaltax = Nothing
                             strsql = " select TxRt_Rate_No,TxRt_Percentage from Gen_TaxRate where (Tx_TaxeID='VAT') and TxRt_Rate_No='" & txtSaleTaxType.Text & "' AND UNIT_CODE='" & gstrUNITID & "'"
                            rsadditionalVAT = New ClsResultSetDB
                            rsadditionalVAT.GetResult(strsql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                             If rsadditionalVAT.GetNoRows > 0 Then
 rsadditionalsurcharge = New ClsResultSetDB
                                strsql = " select TxRt_Rate_No,TxRt_Percentage from Gen_TaxRate where Tx_TaxeID='SsT' and txrt_percentage=5.0 AND UNIT_CODE='" & gstrUNITID & "' AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"
 rsadditionalsurcharge.GetResult(strsql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
 If rsadditionalsurcharge.GetNoRows > 0 Then
     txtSurchargeTaxType.Text = rsadditionalsurcharge.GetValue("TxRt_Rate_No")
     lblSurcharge_Per.Text = rsadditionalsurcharge.GetValue("TxRt_Percentage")
 End If
 rsadditionalsurcharge.ResultSetClose()
 rsadditionalsurcharge = Nothing
                            End If
                            rsadditionalVAT.ResultSetClose()
                            rsadditionalVAT = Nothing
                         End If
                     End If
                End If
                Call txtSaleTaxType_Validating(txtSaleTaxType, New System.ComponentModel.CancelEventArgs(False))
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub cmdSurchargeTaxCode_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSurchargeTaxCode.Click
        '****************************************************
        'Created By     -  Tapan
        'Description    -  To Display Help From Gen_TaxRate
        '****************************************************
        On Error GoTo ErrHandler
        Dim strHelp As String
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                If Len(Me.txtSurchargeTaxType.Text) = 0 Then 'To check if There is No Text Then Show All Help
                    
                    
                    strHelp = ShowList(1, (txtSurchargeTaxType.MaxLength), "", "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND Tx_TaxeID='SST' AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))")
                    If strHelp = "-1" Then 'If No Record Exists In The Table
                        Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Exit Sub
                    Else
                        txtSurchargeTaxType.Text = strHelp
                    End If
                Else
                    
                    
                    strHelp = ShowList(1, (txtSurchargeTaxType.MaxLength), txtSurchargeTaxType.Text, "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND Tx_TaxeID='SST' AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))")
                    If strHelp = "-1" Then 'If No Record Exists In The Table
                        Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Exit Sub
                    Else
                        txtSurchargeTaxType.Text = strHelp
                    End If
                End If
                Call txtSurchargeTaxType_Validating(txtSurchargeTaxType, New System.ComponentModel.CancelEventArgs(False))
        End Select
        
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0068_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
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
    Private Sub frmMKTTRN0068_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Escape
                'If user press the ESC Key ,the Form will be in View Mode
                If Me.CmdGrpChEnt.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                    If ConfirmWindow(10053, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                        Call Me.CmdGrpChEnt.Revert()
                        Call EnableControls(False, Me, True)
                        'In View Mode Disable Combo Of Invoice Type and Inv. Sub type
                        With SpChEntry
                            .Col = GridHeader.ToolCostPerUnit : .Col2 = GridHeader.ToolCostPerUnit : .BlockMode = True : .ColHidden = True : .BlockMode = False
                        End With
                        CmbInvType.Visible = False : CmbInvSubType.Visible = False
                        lblInvSubType.Visible = False : lblInvType.Visible = False
                        txtLocationCode.Enabled = True : txtLocationCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : lblLocCodeDes.Text = ""
                        txtChallanNo.Enabled = True : txtChallanNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        CmdLocCodeHelp.Enabled = True : CmdChallanNo.Enabled = True : Me.SpChEntry.Enabled = False
                        Me.CmdGrpChEnt.Enabled(1) = False
                        Me.CmdGrpChEnt.Enabled(2) = False
                        Me.CmdGrpChEnt.Enabled(5) = False
                        lblSaltax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        lblSurcharge_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        lblLoadingcharge_per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        lblTCSTaxPerDes.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        lblECSStax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        lblSECSStax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        chkExciseExumpted.CheckState = System.Windows.Forms.CheckState.Unchecked
                        If cmdConsigneeDetails.Visible Then cmdConsigneeDetails.Enabled = True
                        cmdConsigneeOK.Enabled = True : cmdConsigneeCancel.Enabled = True
                        If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                            If blnEOU_FLAG = False Then
                                CmbInvType.SelectedIndex = 2 : CmbInvSubType.SelectedIndex = 2
                            Else
                                CmbInvType.SelectedIndex = 1 : CmbInvSubType.SelectedIndex = 2
                            End If
                        End If
                        '***
                        gblnCancelUnload = False
                        gblnFormAddEdit = False
                        With Me.SpChEntry
                            .MaxRows = 1 : .set_RowHeight(1, 300)
                            .Row = 1 : .Row2 = 1 : .Col = GridHeader.InternalPartNo : .Col2 = .MaxCols : .BlockMode = True : .Text = "" : .BlockMode = False
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
    Private Sub frmMKTTRN0068_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Dim rsSalesParameter As New ClsResultSetDB
        Dim rsParameterData As ClsResultSetDB
        Dim strParamQuery As String
        Dim gobjdb As New ClsResultSetDB
        On Error GoTo ErrHandler
        mintIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader1.Tag)
        Call FitToClient(Me, FraChEnt, ctlFormHeader1, CmdGrpChEnt, 500)
        'Set Help Pictures At Command Button
        CmdLocCodeHelp.Image = My.Resources.ico111.ToBitmap
        CmdChallanNo.Image = My.Resources.ico111.ToBitmap
        CmdCustCodeHelp.Image = My.Resources.ico111.ToBitmap
        CmdSaleTaxType.Image = My.Resources.ico111.ToBitmap
        CmdRefNoHelp.Image = My.Resources.ico111.ToBitmap
        If gobjdb.GetResult("Select EOU_FLAG From Company_Mst WHERE UNIT_CODE='" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic) Then
            If gobjdb.GetNoRows > 0 Then
                blnEOU_FLAG = gobjdb.GetValue("EOU_FLAG")
            End If
        End If
        'Initially Disable All Controls
        Call EnableControls(False, Me, True)

        'TT#101188073 Added by Abhijit
        Call Check_GST(gstrUNITID)

        If gblnGSTUnit Then
            EnableDisable_GST()
        End If
        'TT#101188073 Added by Abhijit

        lblSaltax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        lblSurcharge_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        lblLoadingcharge_per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        lblTCSTaxPerDes.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        lblECSStax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        lblSECSStax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        lblDateDes.Text = Format(GetServerDate(), gstrDateFormat)
        With dtpDateDesc
            .Format = DateTimePickerFormat.Custom
            .CustomFormat = gstrDateFormat
            .Value = GetServerDate()
            .Visible = False
        End With
        Call AddTransPortTypeToCombo()
        txtLocationCode.Enabled = True : txtLocationCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        txtChallanNo.Enabled = True : txtChallanNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        CmdLocCodeHelp.Enabled = True : CmdChallanNo.Enabled = True
        Me.SpChEntry.Enabled = False
        Me.CmdGrpChEnt.Enabled(1) = False
        Me.CmdGrpChEnt.Enabled(2) = False
        Me.CmdGrpChEnt.Enabled(5) = False
        blnInvoiceAgainstMultipleSO = CBool(Find_Value("SELECT ISNULL(InvoiceAgainstMultipleSO,0) FROM SALES_PARAMETER WHERE UNIT_CODE='" & gstrUNITID & "'"))
        '10623079
        mblnskipdacinvoicebincheck = CBool(Find_Value("SELECT ISNULL(SKIPDACINVOICEBINCHECK,0)  FROM ProductionConf WHERE UNIT_CODE='" + gstrUNITID + "'"))
        '10623079
        Call SetGridHeader()
        'Function Call To Add Invoice Types In The Inv. Type Combo Box
        Call SelectInvoiceTypeFromSaleConf()
        'In View Mode Disable Combo Of Invoice Type and Inv. Sub type
        CmbInvType.Visible = False : CmbInvSubType.Visible = False
        lblInvSubType.Visible = False : lblInvType.Visible = False
        'Add Row
        Call addRowAtEnterKeyPress(1)
        fraRGPs.Visible = False
        lblRGPDes.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        chkExciseExumpted.CheckState = System.Windows.Forms.CheckState.Unchecked
        fraconsigneeDetails.Visible = False
        rsSalesParameter.GetResult("Select ConsigneeDetails,NAGARE_CSM_EDIT_REQ from Sales_Parameter WHERE UNIT_CODE='" & gstrUNITID & "'")
        If rsSalesParameter.GetValue("ConsigneeDetails") = True Then
            cmdConsigneeDetails.Visible = True
        Else
            cmdConsigneeDetails.Visible = False
        End If
        mbln_CSM_Edit_Req = rsSalesParameter.GetValue("NAGARE_CSM_EDIT_REQ")
        chkDTRemoval.Enabled = True
        chkDTRemoval.CheckState = System.Windows.Forms.CheckState.Unchecked
        dtpRemoval.Format = DateTimePickerFormat.Custom
        dtpRemoval.CustomFormat = gstrDateFormat
        dtpRemoval.Value = GetServerDate()
        dtpRemovalTime.Value = GetServerDate()
        dtpRemoval.Enabled = False
        dtpRemovalTime.Enabled = False
        ctlPerValue.Text = 1
        ChkVBSchUpdFlag()
        If Directory.Exists(gstrLocalCDrive + "EmproInv") = False Then
            Directory.CreateDirectory(gstrLocalCDrive + "EmproInv")
        End If
        blnGSTTAXroundoff = CBool(Find_Value("select GSTTAX_ROUNDOFF from sales_parameter WHERE UNIT_CODE = '" & gstrUNITID & "'"))
        intGSTTAXroundoff_decimal = Val(Find_Value("select GSTTAX_ROUNDOFF_DECIMAL from sales_parameter WHERE UNIT_CODE = '" & gstrUNITID & "'"))
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0068_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo ErrHandler
        mdifrmMain.CheckFormName = mintIndex
        If txtLocationCode.Enabled = True Then
            txtLocationCode.Focus()
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0068_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Me.Tag) = False
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0068_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
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
    Private Sub frmMKTTRN0068_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error GoTo ErrHandler
        frmMKTTRN0020 = Nothing 'Assign form to nothing
        frmMKTTRN0021 = Nothing 'Assign form to nothing
        frmModules.NodeFontBold(Me.Tag) = False
        mdifrmMain.RemoveFormNameFromWindowList = mintIndex
        Me.Dispose()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub addRowAtEnterKeyPress(ByRef pintRows As Short)
        '****************************************************
        'Created By     -  Nisha
        'Description    -  Add Row At Enter Key Press Of Last Column Of Spread
        '****************************************************
        On Error GoTo ErrHandler
        Dim intRowHeight As Short
        With Me.SpChEntry
            .CursorStyle = FPSpreadADO.CursorStyleConstants.CursorStyleArrow
            For intRowHeight = 1 To pintRows
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                .set_RowHeight(.Row, 300)
            Next intRowHeight
            If .MaxRows > 3 Then .ScrollBars = FPSpreadADO.ScrollBarsConstants.ScrollBarsBoth
        End With
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Function SelectInvoiceTypeFromSaleConf() As Object
        '****************************************************
        'Created By     -  Nisha
        'Description    -  Select Invoice Type,Invoice SubTypeDescription From SaleConf
        '****************************************************
        On Error GoTo ErrHandler
        Dim strSaleConfSql As String
        Dim rsSaleConf As ClsResultSetDB
        Dim intRecCount As Short
        Dim intLoopCounter As Short
      
        If blnEOU_FLAG = False Then
            strSaleConfSql = "Select Distinct(Description) from SaleConf where UNIT_CODE='" & gstrUNITID & "'  and Invoice_Type Not in('EXP','STX','CPV') and (fin_start_date <= getdate() and fin_end_date >= getdate()) "
        Else
            strSaleConfSql = "Select Distinct(Description) from SaleConf where UNIT_CODE='" & gstrUNITID & "' and Invoice_Type Not in('EXP','STX','CPV') and (fin_start_date <= getdate() and fin_end_date >= getdate()) "
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
        '****************************************************
        'Created By     -  Nisha
        'Description    -  Select Invoice SubTypeDescription From SaleConf Acc. to Inv. Type
        '****************************************************
        On Error GoTo ErrHandler
        Dim strSaleConfSql As String
        Dim rsSaleConf As ClsResultSetDB
        Dim intRecCount As Short
        Dim intLoopCounter As Short
        strSaleConfSql = "Select Distinct(Sub_Type_Description) from SaleConf where UNIT_CODE='" & gstrUNITID & "' and sub_type not in ('Z') and Description='" & Trim(pstrInvType) & "' and  (fin_start_date <= getdate() and fin_end_date >= getdate()) "
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
        strDesSql = "Select " & Trim(pstrFieldName1) & " from " & Trim(pstrTableName) & " where " & Trim(pstrFieldName2) & "='" & Trim(pstrControlText) & "' AND UNIT_CODE='" & gstrUNITID & "'"
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
    Private Sub OptDiscountPercentage_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OptDiscountPercentage.CheckedChanged
        If eventSender.Checked Then
            '*** set the value of Discount variable to percentage 0->value 1->percentage
            intDiscountType = 1
            If OptDiscountPercentage.Checked = True And Val(txtDiscountAmt.Text) > 100 Then
                MsgBox("Discount cannot be Greater than value.", MsgBoxStyle.Information, "eMPro")
                txtDiscountAmt.Text = ""
                txtDiscountAmt.Focus()
            End If
            ''***
        End If
    End Sub
    Private Sub OptDiscountPercentage_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles OptDiscountPercentage.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo Error_Handler
        If KeyAscii = 13 Then
            txtDiscountAmt.Focus()
        End If
        GoTo EventExitSub
Error_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub OptDiscountValue_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OptDiscountValue.CheckedChanged
        If eventSender.Checked Then
            '*** set the value of Discount variable to value. 0->value 1->percentage
            intDiscountType = 0
        End If
    End Sub
    Private Sub OptDiscountValue_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles OptDiscountValue.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo Error_Handler
        If KeyAscii = 13 Then
            txtDiscountAmt.Focus()
        End If
        GoTo EventExitSub
Error_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub SpChEntry_Change(ByVal eventSender As System.Object, ByVal eventArgs As AxFPSpreadADO._DSpreadEvents_ChangeEvent) Handles SpChEntry.Change
        Dim intRowCount As Short
        Dim intmaxrows As Short
        Dim rsItemMst As ClsResultSetDB
        Dim varFromBox As Object
        Dim varItem As Object
        Dim VarToBox As Object
        Dim varQty As Object
        Dim boxqty As Double
        Dim varCumulativeBoxes As Object

        'added by abhijit
        Dim varItemRate As Double = 0.0
        Dim varQtyRate As Double = 0.0
        Dim varDiscountRate As Double = 0.0
        Dim varDiscountValue As Double = 0.0
        Dim varBaiscValue As Double = 0.0
        Dim varAccessibleValue As Double = 0.0
        Dim varIGSTRate As Double = 0.0
        Dim varIGSTValue As Double = 0.0
        Dim varCGSTRate As Double = 0.0
        Dim varCGSTValue As Double = 0.0
        Dim varSGSTRate As Double = 0.0
        Dim varSGSTValue As Double = 0.0
        Dim varUTGSTRate As Double = 0.0
        Dim varUTGSTValue As Double = 0.0
        Dim varCcessRate As Double = 0.0
        Dim varCcessValue As Double = 0.0
        Dim varItemTotal As Double = 0.0

        'Dim varItemRate As Object = Nothing
        'Dim varQtyRate As Object = Nothing
        'Dim varDiscountRate As Object = Nothing
        'Dim varDiscountValue As Object = Nothing
        'Dim varBaiscValue As Object = Nothing
        'Dim varAccessibleValue As Object = Nothing
        'Dim varIGSTRate As Object = Nothing
        'Dim varIGSTValue As Object = Nothing
        'Dim varCGSTRate As Object = Nothing
        'Dim varCGSTValue As Object = Nothing
        'Dim varSGSTRate As Object = Nothing
        'Dim varSGSTValue As Object = Nothing
        'Dim varUTGSTRate As Object = Nothing
        'Dim varUTGSTValue As Object = Nothing
        'Dim varCcessRate As Object = Nothing
        'Dim varCcessValue As Object = Nothing
        'Dim varItemTotal As Object = Nothing
        'added by abhijit

        With SpChEntry
            If (eventArgs.col = GridHeader.Quantity Or eventArgs.col = GridHeader.BinQty) Then
                If Not RefreshBoxes(eventArgs.row) Then
                    blnGridStatus = True
                    Exit Sub
                End If
            End If
            If (eventArgs.col = GridHeader.FromBox) Or (eventArgs.col = GridHeader.ToBox) Or (eventArgs.col = GridHeader.Quantity) Or (eventArgs.col = GridHeader.RatePerUnit) Then
                intmaxrows = SpChEntry.MaxRows
                For intRowCount = 1 To intmaxrows
                    varFromBox = Nothing
                    VarToBox = Nothing
                    Call .GetText(GridHeader.FromBox, intRowCount, varFromBox)
                    Call .GetText(GridHeader.ToBox, intRowCount, VarToBox)
                    Select Case eventArgs.col
                        Case GridHeader.RatePerUnit
                            .Row = intRowCount
                            varItem = Nothing
                            Call .GetText(GridHeader.InternalPartNo, intRowCount, varItem)
                            varQty = Nothing
                            Call .GetText(GridHeader.Quantity, intRowCount, varQty)
                            If Len(varItem) = 0 Then Exit For
                            'Added by Abhijit
                            If gblnGSTUnit Then
                                CalculateGSTTaxes(intRowCount)
                                '.Col = GridHeader.RatePerUnit : varItemRate = Val(.Text)

                                'varBaiscValue = varItemRate * varQty
                                'Call .SetText(GridHeader.Basic_Value, intRowCount, varBaiscValue)

                                '.Col = GridHeader.Discount_Percent : varDiscountRate = Val(.Text)
                                'varDiscountValue = varBaiscValue * (varDiscountRate / 100)
                                'Call .SetText(GridHeader.Discount_Amt, intRowCount, varDiscountValue)

                                'varAccessibleValue = varBaiscValue - varDiscountValue
                                'Call .SetText(GridHeader.Accessible_Amt, intRowCount, varAccessibleValue)

                                '.Col = GridHeader.IGST_Percent : varIGSTRate = Val(.Text)
                                'varIGSTValue = varAccessibleValue * (varIGSTRate / 100)
                                'Call .SetText(GridHeader.IGST_Amt, intRowCount, varIGSTValue)

                                '.Col = GridHeader.CGST_Percent : varCGSTRate = Val(.Text)
                                'varCGSTValue = varAccessibleValue * (varCGSTRate / 100)
                                'Call .SetText(GridHeader.CGST_Amt, intRowCount, varCGSTValue)

                                '.Col = GridHeader.SGST_Percent : varSGSTRate = Val(.Text)
                                'varSGSTValue = varAccessibleValue * (varSGSTRate / 100)
                                'Call .SetText(GridHeader.SGST_Amt, intRowCount, varSGSTValue)


                                '.Col = GridHeader.UTGST_Percent : varUTGSTRate = Val(.Text)
                                'varUTGSTValue = varAccessibleValue * (varUTGSTRate / 100)
                                'Call .SetText(GridHeader.UTGST_Amt, intRowCount, varUTGSTValue)

                                '.Col = GridHeader.CSESS_TAX_Percent : varCcessRate = Val(.Text)
                                'varCcessValue = varAccessibleValue * (varCcessRate / 100)
                                'Call .SetText(GridHeader.CSESS_TAX_VAL, intRowCount, varCcessValue)

                                'varItemTotal = 0.0
                                'varItemTotal = varAccessibleValue + varIGSTValue + varCGSTValue + varSGSTValue + varUTGSTValue + varCcessValue
                                'Call .SetText(GridHeader.Item_Total, intRowCount, varItemTotal)
                                'Added by Abhijit
                            End If

                        Case GridHeader.Quantity
                            varItem = Nothing
                            Call .GetText(GridHeader.InternalPartNo, intRowCount, varItem)
                            varQty = Nothing
                            Call .GetText(GridHeader.Quantity, intRowCount, varQty)
                            If Len(varItem) = 0 Then Exit For
                            'Added by Abhijit
                            If gblnGSTUnit Then

                                .Col = GridHeader.RatePerUnit : varItemRate = Val(.Text)

                                varBaiscValue = varItemRate * varQty
                                Call .SetText(GridHeader.Basic_Value, intRowCount, varBaiscValue)

                                .Col = GridHeader.Discount_Percent : varDiscountRate = Val(.Text)
                                varDiscountValue = varBaiscValue * (varDiscountRate / 100)
                                Call .SetText(GridHeader.Discount_Amt, intRowCount, varDiscountValue)

                                varAccessibleValue = varBaiscValue - varDiscountValue
                                Call .SetText(GridHeader.Accessible_Amt, intRowCount, varAccessibleValue)

                                .Col = GridHeader.IGST_Percent : varIGSTRate = Val(.Text)
                                varIGSTValue = varAccessibleValue * (varIGSTRate / 100)
                                Call .SetText(GridHeader.IGST_Amt, intRowCount, varIGSTValue)

                                .Col = GridHeader.CGST_Percent : varCGSTRate = Val(.Text)
                                varCGSTValue = varAccessibleValue * (varCGSTRate / 100)
                                Call .SetText(GridHeader.CGST_Amt, intRowCount, varCGSTValue)

                                .Col = GridHeader.SGST_Percent : varSGSTRate = Val(.Text)
                                varSGSTValue = varAccessibleValue * (varSGSTRate / 100)
                                Call .SetText(GridHeader.SGST_Amt, intRowCount, varSGSTValue)


                                .Col = GridHeader.UTGST_Percent : varUTGSTRate = Val(.Text)
                                varUTGSTValue = varAccessibleValue * (varUTGSTRate / 100)
                                Call .SetText(GridHeader.UTGST_Amt, intRowCount, varUTGSTValue)

                                .Col = GridHeader.CSESS_TAX_Percent : varCcessRate = Val(.Text)
                                varCcessValue = varAccessibleValue * (varCcessRate / 100)
                                Call .SetText(GridHeader.CSESS_TAX_VAL, intRowCount, varCcessValue)

                                varItemTotal = 0.0
                                varItemTotal = varAccessibleValue + varIGSTValue + varCGSTValue + varSGSTValue + varUTGSTValue + varCcessValue
                                Call .SetText(GridHeader.Item_Total, intRowCount, varItemTotal)
                                'Added by Abhijit                
                            End If

                        Case GridHeader.FromBox

                            If varFromBox > VarToBox Then
                                MsgBox("From Boxes can't be greater than To Boxes", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                                eventArgs.row = eventArgs.row
                                eventArgs.col = GridHeader.FromBox
                                If eventArgs.row <> 1 Then

                                    .Text = VarToBox
                                Else
                                    .Text = CStr(1)
                                End If
                                .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                .Focus()
                                Exit Sub
                            End If
                        Case GridHeader.ToBox


                            If VarToBox < varFromBox Then
                                MsgBox("To Boxes can't be less than From Boxes", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                                eventArgs.row = eventArgs.row
                                eventArgs.col = GridHeader.ToBox

                                .Text = varFromBox
                                .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                .Focus()
                                Exit Sub
                            End If
                    End Select
                    If intRowCount = 1 Then

                        If Len(Trim(varFromBox)) Then

                            If Len(Trim(VarToBox)) Then


                                Call .SetText(GridHeader.CumulativeBoxes, intRowCount, (Val(VarToBox) - Val(varFromBox)) + 1)
                            End If
                        End If
                    Else
                        varCumulativeBoxes = Nothing
                        Call .GetText(GridHeader.CumulativeBoxes, intRowCount - 1, varCumulativeBoxes)

                        If Len(Trim(varCumulativeBoxes)) Then

                            If Len(Trim(varFromBox)) Then

                                If Len(Trim(VarToBox)) Then



                                    Call .SetText(GridHeader.CumulativeBoxes, intRowCount, varCumulativeBoxes + ((Val(VarToBox) - Val(varFromBox)) + 1))
                                End If
                            End If
                        End If
                    End If
                Next
            End If
        End With
    End Sub
    Private Sub LoadSalesParameter()

        Dim strSql As String = String.Empty
        Try
            strSql = "SELECT  InsExc_Excise,CustSupp_Inc,EOU_Flag, Basic_Roundoff, Basic_Roundoff_decimal," & _
                                "InsInc_SalesTax, TotalToolCostRoundoff, TotalToolCostRoundoff_Decimal, " & _
                                "TotalInvoiceAmount_Roundoff,TotalInvoiceAmountRoundOff_Decimal " & _
                                "FROM Sales_Parameter where UNIT_CODE = '" & gstrUNITID & "'"
            If dtSalesParameter Is Nothing OrElse dtSalesParameter.Rows.Count = 0 Then
                dtSalesParameter = SqlConnectionclass.GetDataTable(strSql)
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub CalculateGSTTaxes(ByVal rowIndex As Integer)
        If Not gblnGSTUnit Then Exit Sub
        Dim dblDiscountAmt As Double = 0
        Dim dblBasicValue As Double = 0
        Dim dblAssessableValue As Double = 0
        Dim dblCGSTPercent As Double = 0
        Dim dblSGSTPercent As Double = 0
        Dim dblIGSTPercent As Double = 0
        Dim dblUTGSTPercent As Double = 0
        Dim dblCCESSPercent As Double = 0
        Dim dblCGSTAmt As Double = 0
        Dim dblSGSTAmt As Double = 0
        Dim dblIGSTAmt As Double = 0
        Dim dblUTGSTAmt As Double = 0
        Dim dblCCESSAmt As Double = 0
        Try
            LoadSalesParameter()
            With SpChEntry
                .Row = rowIndex
                ' dblBasicValue = CalculateBasicValue(rowIndex, CBool(dtSalesParameter.Rows(0)("Basic_Roundoff")), SpChEntry)
                dblBasicValue = CalculateBasicValue(rowIndex, CBool(dtSalesParameter.Rows(0)("Basic_Roundoff")))
                .Col = GridHeader.Discount_Amt : dblDiscountAmt = Val(.Text)
                dblAssessableValue = CalculateAccessibleValue(rowIndex, Math.Round(Val(ctlInsurance.Text)), CBool(dtSalesParameter.Rows(0)("InsExc_Excise"))) - dblDiscountAmt
                .Col = GridHeader.Basic_Value : .Text = dblBasicValue
                .Col = GridHeader.Accessible_Amt : .Text = dblAssessableValue

                .Col = GridHeader.CGST_Percent : dblCGSTPercent = Val(.Text)
                .Col = GridHeader.SGST_Percent : dblSGSTPercent = Val(.Text)
                .Col = GridHeader.IGST_Percent : dblIGSTPercent = Val(.Text)
                .Col = GridHeader.UTGST_Percent : dblUTGSTPercent = Val(.Text)
                .Col = GridHeader.CSESS_TAX_Percent : dblCCESSPercent = Val(.Text)


                dblCGSTAmt = (dblAssessableValue * dblCGSTPercent) / 100
                dblSGSTAmt = (dblAssessableValue * dblSGSTPercent) / 100
                dblIGSTAmt = (dblAssessableValue * dblIGSTPercent) / 100
                dblUTGSTAmt = (dblAssessableValue * dblUTGSTPercent) / 100
                dblCCESSAmt = (dblAssessableValue * dblCCESSPercent) / 100
                If blnGSTTAXroundoff Then
                    .Col = GridHeader.CGST_Amt : .Text = dblCGSTAmt
                    .Col = GridHeader.SGST_Amt : .Text = dblSGSTAmt
                    .Col = GridHeader.IGST_Amt : .Text = dblIGSTAmt
                    .Col = GridHeader.UTGST_Amt : .Text = dblUTGSTAmt
                    .Col = GridHeader.CSESS_TAX_VAL : .Text = dblCCESSAmt
                Else
                    dblCGSTAmt = System.Math.Round(dblCGSTAmt, intGSTTAXroundoff_decimal)
                    dblSGSTAmt = System.Math.Round(dblSGSTAmt, intGSTTAXroundoff_decimal)
                    dblIGSTAmt = System.Math.Round(dblIGSTAmt, intGSTTAXroundoff_decimal)
                    dblUTGSTAmt = System.Math.Round(dblUTGSTAmt, intGSTTAXroundoff_decimal)
                    dblCCESSAmt = System.Math.Round(dblCCESSAmt, intGSTTAXroundoff_decimal)

                    .Col = GridHeader.CGST_Amt : .Text = dblCGSTAmt
                    .Col = GridHeader.SGST_Amt : .Text = dblSGSTAmt
                    .Col = GridHeader.IGST_Amt : .Text = dblIGSTAmt
                    .Col = GridHeader.UTGST_Amt : .Text = dblUTGSTAmt
                    .Col = GridHeader.CSESS_TAX_VAL : .Text = dblCCESSAmt

                End If
                
                ' .Row = 0 : .Col = EnumInv.CSESS_TAX_VAL : .Text = "CSESS TAX (Val)" : .set_ColWidth(EnumInv.CSESS_TAX_VAL, 1200)


                .Col = GridHeader.Item_Total : .Text = dblAssessableValue + dblCGSTAmt + dblSGSTAmt + dblIGSTAmt + dblUTGSTAmt + dblCCESSAmt
            End With
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub CalculateGSTTaxes_onGRID(ByVal rowIndex As Integer)
        If Not gblnGSTUnit Then Exit Sub
        Dim dblDiscountAmt As Double = 0
        Dim dblBasicValue As Double = 0
        Dim dblAssessableValue As Double = 0
        Dim dblCGSTPercent As Double = 0
        Dim dblSGSTPercent As Double = 0
        Dim dblIGSTPercent As Double = 0
        Dim dblUTGSTPercent As Double = 0
        Dim dblCCESSPercent As Double = 0
        Dim dblCGSTAmt As Double = 0
        Dim dblSGSTAmt As Double = 0
        Dim dblIGSTAmt As Double = 0
        Dim dblUTGSTAmt As Double = 0
        Dim dblCCESSAmt As Double = 0
        Try
            LoadSalesParameter()
            With SpChEntry
                .Row = rowIndex
                ' dblBasicValue = CalculateBasicValue(rowIndex, CBool(dtSalesParameter.Rows(0)("Basic_Roundoff")), SpChEntry)
                dblBasicValue = CalculateBasicValue(rowIndex, CBool(dtSalesParameter.Rows(0)("Basic_Roundoff")))
                .Col = GridHeader.Discount_Amt : dblDiscountAmt = Val(.Text)
                dblAssessableValue = dblBasicValue - dblDiscountAmt
                .Col = GridHeader.Basic_Value : .Text = dblBasicValue
                .Col = GridHeader.Accessible_Amt : .Text = dblAssessableValue

                .Col = GridHeader.CGST_Percent : dblCGSTPercent = Val(.Text)
                .Col = GridHeader.SGST_Percent : dblSGSTPercent = Val(.Text)
                .Col = GridHeader.IGST_Percent : dblIGSTPercent = Val(.Text)
                .Col = GridHeader.UTGST_Percent : dblUTGSTPercent = Val(.Text)
                .Col = GridHeader.CSESS_TAX_Percent : dblCCESSPercent = Val(.Text)


                dblCGSTAmt = (dblAssessableValue * dblCGSTPercent) / 100
                dblSGSTAmt = (dblAssessableValue * dblSGSTPercent) / 100
                dblIGSTAmt = (dblAssessableValue * dblIGSTPercent) / 100
                dblUTGSTAmt = (dblAssessableValue * dblUTGSTPercent) / 100
                dblCCESSAmt = (dblAssessableValue * dblCCESSPercent) / 100

                .Col = GridHeader.CGST_Amt : .Text = dblCGSTAmt
                .Col = GridHeader.SGST_Amt : .Text = dblSGSTAmt
                .Col = GridHeader.IGST_Amt : .Text = dblIGSTAmt
                .Col = GridHeader.UTGST_Amt : .Text = dblUTGSTAmt
                .Col = GridHeader.CSESS_TAX_VAL : .Text = dblCCESSAmt

                ' .Row = 0 : .Col = EnumInv.CSESS_TAX_VAL : .Text = "CSESS TAX (Val)" : .set_ColWidth(EnumInv.CSESS_TAX_VAL, 1200)


                .Col = GridHeader.Item_Total : .Text = dblAssessableValue + dblCGSTAmt + dblSGSTAmt + dblIGSTAmt + dblUTGSTAmt + dblCCESSAmt
            End With
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub CalculateDiscount(ByVal gridEnum As GridHeader, ByVal rowIndex As Integer)
        If Not gblnGSTUnit Then Exit Sub
        Dim dblBasicValue As Double = 0
        Dim dblDiscountPercent As Double = 0
        Dim dblDiscountAmt As Double = 0
        Try
            LoadSalesParameter()
            If gridEnum = GridHeader.Discount_Percent Then
                With SpChEntry
                    .Row = rowIndex
                    dblBasicValue = CalculateBasicValue(rowIndex, CBool(dtSalesParameter.Rows(0)("Basic_Roundoff")))
                    .Col = GridHeader.Discount_Percent : dblDiscountPercent = Val(.Text)
                    .Col = GridHeader.Discount_Amt : .Text = (dblBasicValue * dblDiscountPercent) / 100
                End With
                CalculateGSTTaxes(rowIndex)
            ElseIf gridEnum = GridHeader.Discount_Amt Then
                With SpChEntry
                    .Row = rowIndex
                    dblBasicValue = CalculateBasicValue(rowIndex, CBool(dtSalesParameter.Rows(0)("Basic_Roundoff")))
                    .Col = GridHeader.Discount_Amt : dblDiscountAmt = Val(.Text)
                    .Col = GridHeader.Discount_Percent : .Text = (dblDiscountAmt / dblBasicValue) * 100
                End With
                CalculateGSTTaxes(rowIndex)
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub SpChEntry_EditChange(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_EditChangeEvent) Handles SpChEntry.EditChange
        If Not gblnGSTUnit Then Exit Sub
        Dim dblBasicValue As Double = 0
        Dim dblDiscountPercent As Double = 0
        Dim dblDiscountAmt As Double = 0
        Try
            LoadSalesParameter()
            If e.col = GridHeader.Discount_Percent Then
                With SpChEntry
                    .Row = e.row
                    dblBasicValue = CalculateBasicValue(e.row, CBool(dtSalesParameter.Rows(0)("Basic_Roundoff")))
                    .Col = GridHeader.Discount_Percent : dblDiscountPercent = Val(.Text)
                    .Col = GridHeader.Discount_Amt : .Text = (dblBasicValue * dblDiscountPercent) / 100
                End With
                CalculateGSTTaxes(e.row)
            ElseIf e.col = GridHeader.Discount_Amt Then
                With SpChEntry
                    .Row = e.row
                    dblBasicValue = CalculateBasicValue(e.row, CBool(dtSalesParameter.Rows(0)("Basic_Roundoff")))
                    .Col = GridHeader.Discount_Amt : dblDiscountAmt = Val(.Text)
                    .Col = GridHeader.Discount_Percent : .Text = (dblDiscountAmt / dblBasicValue) * 100
                End With
                CalculateGSTTaxes(e.row)
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    
    Private Sub SpChEntry_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles SpChEntry.GotFocus
    End Sub
    Private Sub SpChEntry_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxFPSpreadADO._DSpreadEvents_KeyDownEvent) Handles SpChEntry.KeyDownEvent
        Dim strHelp As String
        Dim strCondition As String
        Dim strItemCode As String
        Dim strpartcode As String

        If CmdGrpChEnt.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            If eventArgs.keyCode = System.Windows.Forms.Keys.F1 And SpChEntry.ActiveCol = GridHeader.CVD Then
                'by Abhijit
                If gblnGSTUnit Then
                    Exit Sub
                End If
                'by Abhijit
                With SpChEntry
                    .Row = .ActiveRow
                    .Col = .ActiveCol
                    If Len(Trim(.Text)) = 0 Then 'To check if There is No Text Then Show All Help

                        strHelp = ShowList(1, 6, "", "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND Tx_TaxeID='CVD' AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))")
                        If strHelp = "-1" Then 'If No Record Exists In The Table
                            Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                            Exit Sub
                        Else
                            .Text = strHelp
                        End If
                    Else
                        'To Display All Possible Help Starting With Text in TextField

                        strHelp = ShowList(1, 6, "", "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND Tx_TaxeID='CVD' AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))")
                        If strHelp = "-1" Then 'If No Record Exists In The Table
                            Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                            Exit Sub
                        Else
                            .Text = strHelp
                        End If
                    End If
                End With
            ElseIf eventArgs.keyCode = System.Windows.Forms.Keys.F1 And SpChEntry.ActiveCol = GridHeader.SAD Then
                'by Abhijit
                If gblnGSTUnit Then
                    Exit Sub
                End If
                'by Abhijit
                With SpChEntry
                    .Row = .ActiveRow
                    .Col = .ActiveCol
                    If Len(Trim(.Text)) = 0 Then 'To check if There is No Text Then Show All Help

                        strHelp = ShowList(1, 6, "", "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND Tx_TaxeID='SAD' AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))")
                        If strHelp = "-1" Then 'If No Record Exists In The Table
                            Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                            Exit Sub
                        Else
                            .Text = strHelp
                        End If
                    Else
                        'To Display All Possible Help Starting With Text in TextField

                        strHelp = ShowList(1, 6, Trim(.Text), "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND Tx_TaxeID='SAD' AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))")
                        If strHelp = "-1" Then 'If No Record Exists In The Table
                            Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                            Exit Sub
                        Else
                            .Text = strHelp
                        End If
                    End If
                End With
                '10808160
            ElseIf eventArgs.keyCode = System.Windows.Forms.Keys.F1 And SpChEntry.ActiveCol = GridHeader.model Then
                With SpChEntry
                    .Row = .ActiveRow
                    .Col = GridHeader.InternalPartNo
                    strItemCode = Trim(.Text)

                    .Row = .ActiveRow
                    .Col = GridHeader.CustPartNo
                    strpartcode = Trim(.Text)

                    .Col = .ActiveCol
                    .Row = .ActiveRow

                    If Len(Trim(.Text)) = 0 Then 'To check if There is No Text Then Show All Help
                        strHelp = ShowList(1, 6, "", "MODEL_CODE", "ENDDATE", "BUDGETITEM_MST ", "  AND CUST_DRGNO='" & strpartcode & "' AND ITEM_CODE='" & strItemCode & "' AND ENDDATE >= '" & GetServerDateNew().ToString("dd MMM yyyy") & "'")
                        If strHelp = "-1" Then 'If No Record Exists In The Table
                            Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                            Exit Sub
                        Else
                            .Text = strHelp
                        End If
                    Else
                        'To Display All Possible Help Starting With Text in TextField
                        strHelp = ShowList(1, 6, "", "MODEL_CODE", "ENDDATE", "BUDGETITEM_MST ", " AND CUST_DRGNO='" & strpartcode & "' AND ITEM_CODE='" & strItemCode & "' AND ENDDATE >= '" & GetServerDateNew().ToString("dd MMM yyyy") & "'")
                        If strHelp = "-1" Then 'If No Record Exists In The Table
                            Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                            Exit Sub
                        Else
                            .Text = strHelp
                        End If
                    End If
                End With
                '10808160

            ElseIf eventArgs.keyCode = System.Windows.Forms.Keys.F1 And SpChEntry.ActiveCol = GridHeader.EXC Then
                'by Abhijit
                If gblnGSTUnit Then
                    Exit Sub
                End If
                'by Abhijit
                With SpChEntry
                    .Row = .ActiveRow
                    .Col = GridHeader.InternalPartNo
                    strItemCode = Trim(.Text)
                    .Col = .ActiveCol
                    strCondition = "AND Tx_TaxeID='EXC' AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date)) " & PrepareQueryForShowingExcise(False, strItemCode)
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
            End If
        End If
        If eventArgs.keyCode = 13 And SpChEntry.ActiveCol = GridHeader.Quantity Then
             If blnInvoiceAgainstMultipleSO Then
                If SpChEntry.MaxRows = SpChEntry.ActiveRow Then
                    CmdGrpChEnt.Focus()
                End If
            Else
                CmdGrpChEnt.Focus()
            End If
         End If
    End Sub
    Private Sub SpChEntry_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxFPSpreadADO._DSpreadEvents_KeyPressEvent) Handles SpChEntry.KeyPressEvent
        On Error GoTo ErrHandler
        Select Case eventArgs.keyAscii
            Case 39, 34, 96, 45
                eventArgs.keyAscii = 0
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
	Private Sub SpChEntry_KeyUpEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxFPSpreadADO._DSpreadEvents_KeyUpEvent) Handles SpChEntry.KeyUpEvent
		Dim intRow As Short
		Dim intDelete As Short
		Dim intLoopCount As Short
		Dim intMaxLoop As Short
		Dim VarDelete As Object
		If ((eventArgs.Shift = 2) And (eventArgs.KeyCode = System.Windows.Forms.Keys.D)) Then
            If CmdGrpChEnt.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                With SpChEntry
                    If .MaxRows > 1 Then
                        intRow = .ActiveRow : intMaxLoop = SpChEntry.MaxRows
                        For intLoopCount = 1 To intMaxLoop
                            If intLoopCount <> intRow Then
 VarDelete = Nothing
 Call .GetText(GridHeader.delete, intLoopCount, VarDelete)
 
 If UCase(VarDelete) = "D" Then
     intDelete = intDelete + 1
 End If
                            End If
                        Next
                        If (intMaxLoop - intDelete) > 1 Then
                            Call .SetText(GridHeader.delete, intRow, "D")
                            .Row = .ActiveRow : .Row2 = .ActiveRow : .BlockMode = True : .RowHidden = True : .BlockMode = False
                        End If
                    End If
                End With
            End If
		End If
	End Sub
	Private Sub SpChEntry_LeaveCell(ByVal eventSender As System.Object, ByVal eventArgs As AxFPSpreadADO._DSpreadEvents_LeaveCellEvent) Handles SpChEntry.LeaveCell
		Dim lstrReturnVal As String
		Dim strWhereClause As String
 		Item_Description((eventArgs.NewRow))
 		lstrReturnVal = ""
        If CmdGrpChEnt.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            With SpChEntry
                If (eventArgs.col = GridHeader.Quantity Or eventArgs.col = GridHeader.BinQty) Then
                    If blnGridStatus = True Then
                        System.Windows.Forms.Application.DoEvents()
                        eventArgs.row = eventArgs.row
                        .Row2 = eventArgs.row
                        eventArgs.col = eventArgs.col
                        .Col2 = eventArgs.col
                        .BlockMode = True
                        .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                        .Focus()
                        .BlockMode = False
                        blnGridStatus = False
                        eventArgs.cancel = True
                        Exit Sub
                    End If
                 
                End If
                 If eventArgs.col = GridHeader.CVD Then
                    eventArgs.col = GridHeader.CVD
                    eventArgs.row = .ActiveRow
                    If Trim(.Text) <> "" Then
                        strWhereClause = " WHERE TxRt_Rate_No='" & Trim(.Text) & "' AND Tx_TaxeID='CVD' and UNIT_CODE='" & gstrUNITID & "' AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"
                        lstrReturnVal = SelectDataFromTable("TxRt_Rate_No", "Gen_TaxRate", strWhereClause)
                        If Len(lstrReturnVal) = 0 Then
                            .Text = ""
                            MsgBox("Invalid Tax Code", MsgBoxStyle.Critical, "eMPro")
                        End If
                    End If
                ElseIf eventArgs.col = GridHeader.SAD Then
                    eventArgs.Col = GridHeader.SAD
                    eventArgs.row = .ActiveRow
                    If Trim(.Text) <> "" Then
                        strWhereClause = " WHERE TxRt_Rate_No='" & Trim(.Text) & "' AND Tx_TaxeID='SAD' and UNIT_CODE='" & gstrUNITID & "' AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"
                        lstrReturnVal = SelectDataFromTable("TxRt_Rate_No", "Gen_TaxRate", strWhereClause)
                        If Len(lstrReturnVal) = 0 Then
                            .Text = ""
                            MsgBox("Invalid Tax Code", MsgBoxStyle.Critical, "eMPro")
                        End If
                    End If
                End If
            End With
        End If
	End Sub
	Private Sub txtAddress1_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtAddress1.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		'****************************************************
		'Created By     -  Nisha
		'Description    -  At Enter Key Press Set Focus To Next Control
		'****************************************************
		On Error GoTo ErrHandler
		Select Case KeyAscii
			Case System.Windows.Forms.Keys.Return
				txtAddress2.Focus()
			Case 39, 34, 96
				KeyAscii = 0
		End Select
		GoTo EventExitSub
ErrHandler: 'The Error Handling Code Starts here
		Call gobjError.RaiseError(Err.Number, err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub: 
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	Private Sub txtAddress2_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtAddress2.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		'****************************************************
		'Created By     -  Nisha
		'Description    -  At Enter Key Press Set Focus To Next Control
		'****************************************************
		On Error GoTo ErrHandler
		Select Case KeyAscii
			Case System.Windows.Forms.Keys.Return
				txtAddress3.Focus()
			Case 39, 34, 96
				KeyAscii = 0
		End Select
		GoTo EventExitSub
ErrHandler: 'The Error Handling Code Starts here
		Call gobjError.RaiseError(Err.Number, err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
		
EventExitSub: 
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	Private Sub txtAddress3_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtAddress3.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		'****************************************************
		'Created By     -  Nisha
		'Description    -  At Enter Key Press Set Focus To Next Control
		'****************************************************
		On Error GoTo ErrHandler
		Select Case KeyAscii
			Case System.Windows.Forms.Keys.Return
				cmdConsigneeOK.Focus()
			Case 39, 34, 96
				KeyAscii = 0
		End Select
		GoTo EventExitSub
ErrHandler: 'The Error Handling Code Starts here
		Call gobjError.RaiseError(Err.Number, err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub: 
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
 	Private Sub txtamendno_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtamendno.TextChanged
		If Trim(txtAmendNo.Text) = "" Then
            If Me.CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                SpChEntry.MaxRows = 0
                mstrItemCode = ""
                 lblCreditTerm.Text = ""
                lblCreditTermDesc.Text = ""
             End If
		End If
	End Sub
	Private Sub txtAmendNo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAmendNo.Enter
		With txtAmendNo
			.SelectionStart = 0
			.SelectionLength = Len(txtAmendNo.Text)
		End With
	End Sub
	Private Sub txtAmendNo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtAmendNo.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		'*****************
		'***********************************
		'Created By     -  Nitin Sood
		'Description    -  At Enter Key Press Set Focus To Next Control
		'****************************************************
		On Error GoTo ErrHandler
		Select Case KeyAscii
			Case System.Windows.Forms.Keys.Return
				Select Case Me.CmdGrpChEnt.mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        If Len(txtAmendNo.Text) > 0 Then
                            Call txtAmendNo_Validating(txtAmendNo, New System.ComponentModel.CancelEventArgs(False))
                        Else
                            If (CmbInvType.Text = "JOBWORK INVOICE") Then
                           
                            Else
 txtCarrServices.Focus()
                            End If
                         End If
                End Select
			Case 39, 34, 96
				KeyAscii = 0
		End Select
		GoTo EventExitSub
ErrHandler: 'The Error Handling Code Starts here
		Call gobjError.RaiseError(Err.Number, err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub: 
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
    Private Sub txtAmendNo_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtAmendNo.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '****************************************************
        'Created By     -  Nitin Sood
        'Description    -  If F1 Key Press Then Display Help From Customer Master/Vendor Master
        '****************************************************
        On Error GoTo ErrHandler
        If KeyCode = 112 Then
            If CmdRefNoHelp.Enabled Then Call CmdRefNoHelp_Click(CmdRefNoHelp, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtAmendNo_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtAmendNo.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        '****************************************************
        'Created By     -  Nitin Sood
        'Description    -  Validate Reference Number Entered
        '****************************************************
        On Error GoTo ErrHandler
        'Only if Some Ref No. is Added
        If Trim(txtRefNo.Text) <> "" Then
            'if Some Amend No is Entered
            If Trim(txtAmendNo.Text) <> "" Then
                If SelectDataFromTable("Amendment_No", "Cust_ORD_HDR", " Where UNIT_CODE='" & gstrUNITID & "' and Account_Code = '" & Trim(txtCustCode.Text) & "' And Cust_Ref = '" & Trim(txtRefNo.Text) & "' And Active_Flag = 'A'  AND  Amendment_No <> '' AND  Amendment_No = '" & Trim(txtAmendNo.Text) & "'") <> "" Then
                    If (CmbInvType.Text = "JOBWORK INVOICE") Then
                    Else
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
        '****************************************************
        'Created By     -  Nisha
        'Description    -  At Enter Key Press Set Focus To Next Control
        '****************************************************
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
        '****************************************************
        'Created By     -  Nisha
        'Description    -  At Enter Key Press Set Focus To Next Control
        '****************************************************
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
        '****************************************************
        'Created By     -  Nisha
        'Description    -  If F1 Key Press Then Display Help From SalesChallan_Dtl
        '****************************************************
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
        '****************************************************
        'Created By     -  Nisha
        'Description    -  Check Validity Of Challan No. In SalesChallan_Dtl
        '****************************************************
        Dim strCondition As String
        Dim rsChallanEntry As ClsResultSetDB
        Dim strInvoiceType As String
        On Error GoTo ErrHandler
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                If Len(txtChallanNo.Text) > 0 Then
                    'Check Existance Of Doc No In The SalesChallan_Dtl
                    If blnEOU_FLAG = True Then
                        strCondition = "Invoice_type <> 'EXP' and UNIT_CODE='" & gstrUNITID & "'"
                    Else
                        strCondition = "UNIT_CODE='" & gstrUNITID & "'"
                    End If
                    If CheckExistanceOfFieldData((txtChallanNo.Text), "Doc_No", "SalesChallan_Dtl", strCondition) Then
                        'If Challan No. Exists
                        'Get Data From Challan_Dtl,Cust_Ord_Dtl,Sales_Dtl
                        If Len(txtLocationCode.Text) > 0 Then
                            If GetDataInViewMode() Then 'if record found
                                rsChallanEntry = New ClsResultSetDB
                                rsChallanEntry.GetResult("Select a.Description,a.Sub_Type_Description from SaleConf a,SalesChallan_Dtl b where a.UNIT_CODE=b.UNIT_CODE AND a.UNIT_CODE= '" & gstrUNITID & "' AND Doc_No = " & txtChallanNo.Text & " and a.Invoice_Type = b.Invoice_type and a.Sub_type = b.Sub_Category and a.Location_code = b.Location_code and (fin_start_date <= getdate() and fin_end_date >= getdate())")
                                strInvoiceType = UCase(rsChallanEntry.GetValue("Description"))
                                If UCase(strInvoiceType) = "SERVICE INVOICE" Then
                                    lblSaleTaxType.Text = "Service Tax Code"
                                Else
                                    lblSaleTaxType.Text = "Sale Tax    Code"
                                End If
                                CmbInvType.Enabled = False
                                CmbInvSubType.Enabled = False
                                If UCase(strInvoiceType) <> "SAMPLE INVOICE" Then
                                    With SpChEntry
                                        .Col = GridHeader.ToolCostPerUnit : .Col2 = GridHeader.ToolCostPerUnit : .BlockMode = True : .ColHidden = True : .BlockMode = False
                                    End With
                                Else
                                    With SpChEntry
                                        .Col = GridHeader.ToolCostPerUnit : .Col2 = GridHeader.ToolCostPerUnit : .BlockMode = True : .ColHidden = False : .BlockMode = False
                                        .Col = GridHeader.ToolCostPerUnit : .Col2 = GridHeader.ToolCostPerUnit : .BlockMode = True : .Lock = False : .BlockMode = False
                                    End With
                                End If
                                Cmditems.Enabled = True
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
        '******Check For Temporary Challan No.
        If Val(txtChallanNo.Text) > 9900000 Then
            If blnInvoiceAgainstMultipleSO Then
                Cmditems.Enabled = False
                DisplayDetailsInSpread(gstrCURRENCYCODE) 'Procedure Call To Select Data From Sales_Dtl
                If Me.CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                    If CDbl(Trim(txtChallanNo.Text)) > 99000000 Then
                        Me.CmdGrpChEnt.Enabled(1) = True
                        Me.CmdGrpChEnt.Enabled(2) = True
                    End If
                End If
                Me.SSTab1.Controls.Remove(Me._SSTab1_TabPage2)
            Else
                Cmditems.Enabled = True
            End If
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
    Private Sub txtContactPerson_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtContactPerson.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '****************************************************
        'Created By     -  Nisha
        'Description    -  At Enter Key Press Set Focus To Next Control
        '****************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                txtECC.Focus()
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
    Private Sub txtCustCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCustCode.TextChanged
        On Error GoTo ErrHandler
        If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            If Not blnInvoiceAgainstMultipleSO Then
                lblCustCodeDes.Text = ""
                txtRefNo.Text = ""
                SpChEntry.MaxRows = 0
                mstrItemCode = ""
                lblAddressDes.Text = ""
                fraRGPs.Visible = False
            End If
        End If
        If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            If txtCustCode.Enabled Then txtCustCode.Focus()
        ElseIf CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            CmdGrpChEnt.Focus()
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtCustCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCustCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '****************************************************
        'Created By     -  Nisha
        'Description    -  At Enter Key Press Set Focus To Next Control
        '****************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        If Len(txtCustCode.Text) > 0 Then
                            Call txtCustCode_Validating(txtCustCode, New System.ComponentModel.CancelEventArgs(False))
                        Else
                            If (UCase(CmbInvType.Text) = "NORMAL INVOICE") Or (UCase(CmbInvType.Text) = "JOBWORK INVOICE") Or (UCase(CmbInvType.Text) = "EXPORT INVOICE") Or (UCase(CmbInvType.Text) = "SERVICE INVOICE") Then
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
        '****************************************************
        'Created By     -  Nisha
        'Description    -  If F1 Key Press Then Display Help From Customer Master/Vendor Master
        '****************************************************
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
        On Error GoTo ErrHandler
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                If Len(txtCustCode.Text) > 0 Then
                    'Changes against 10737738 
                    If UCase(Trim(mstrInvoiceType)) = "INV" Or UCase(Trim(mstrInvoiceType)) = "SMP" Or UCase(Trim(mstrInvoiceType)) = "TRF" Or UCase(Trim(mstrInvoiceType)) = "JOB" Or UCase(Trim(mstrInvoiceType)) = "EXP" Or UCase(Trim(mstrInvoiceType)) = "SRC" Then
                        If SchUpdFlag = True Then
                            If CheckExistanceOfFieldData((txtCustCode.Text), "Customer_Code", "Customer_Mst", "UNIT_CODE='" & gstrUNITID & "' AND (SCH_UPLOAD_CODE ='KENN') AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))") Then
                                Call SelectDescriptionForField("Cust_Name", "Customer_Code", "Customer_Mst", lblCustCodeDes, Trim(txtCustCode.Text))
                                If (UCase(CmbInvType.Text) = "NORMAL INVOICE") Or (UCase(CmbInvType.Text) = "JOBWORK INVOICE") Or (UCase(CmbInvType.Text) = "EXPORT INVOICE") Or (UCase(CmbInvType.Text) = "SERVICE INVOICE") Then
                                    If UCase(CmbInvSubType.Text) <> "SCRAP" Then
                                        If txtRefNo.Enabled Then txtRefNo.Focus()
                                    Else
                                        txtCarrServices.Focus()
                                    End If
                                Else
                                    txtCarrServices.Focus()
                                End If
                            Else
                                Cancel = True
                                Call ConfirmWindow(10417, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                                txtCustCode.Text = ""
                                If txtCustCode.Enabled Then txtCustCode.Focus()
                            End If
                        Else
                            If CheckExistanceOfFieldData((txtCustCode.Text), "Customer_Code", "Customer_Mst", "UNIT_CODE='" & gstrUNITID & "' AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))") Then
                                Call SelectDescriptionForField("Cust_Name", "Customer_Code", "Customer_Mst", lblCustCodeDes, Trim(txtCustCode.Text))
                                If (UCase(CmbInvType.Text) = "NORMAL INVOICE") Or (UCase(CmbInvType.Text) = "JOBWORK INVOICE") Or (UCase(CmbInvType.Text) = "EXPORT INVOICE") Or (UCase(CmbInvType.Text) = "SERVICE INVOICE") Then
                                    If UCase(CmbInvSubType.Text) <> "SCRAP" Then
                                        If txtRefNo.Enabled Then txtRefNo.Focus()
                                    Else
                                        txtCarrServices.Focus()
                                    End If
                                Else
                                    txtCarrServices.Focus()
                                End If
                            Else
                                Cancel = True
                                Call ConfirmWindow(10417, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                                txtCustCode.Text = ""
                                If txtCustCode.Enabled Then txtCustCode.Focus()
                            End If
                        End If
                        If Len(Trim(txtCustCode.Text)) > 0 Then
                            rsCustMst = New ClsResultSetDB
                            strCustMst = "SELECT Bill_Address1 + ', '  +  Bill_Address2 + ', ' + Bill_City + ' - ' + Bill_Pin as  invoiceAddress from Customer_Mst where Customer_code ='" & txtCustCode.Text & "' AND UNIT_CODE='" & gstrUNITID & "'"
                            rsCustMst.GetResult(strCustMst)
                            If rsCustMst.GetNoRows > 0 Then
                                lblAddressDes.Text = rsCustMst.GetValue("InvoiceAddress")
                            End If
                            rsCustMst = Nothing
                        End If
                        '***
                    Else
                        If CheckExistanceOfFieldData((txtCustCode.Text), "Vendor_Code", "Vendor_Mst", "UNIT_CODE='" & gstrUNITID & "'") Then
                            Call SelectDescriptionForField("Vendor_name", "Vendor_Code", "Vendor_Mst", lblCustCodeDes, Trim(txtCustCode.Text))
                            If txtRefNo.Enabled Then
                                txtRefNo.Focus()
                            Else
                                txtCarrServices.Focus()
                            End If
                        Else
                            Cancel = True
                            Call ConfirmWindow(10417, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                            txtCustCode.Text = ""
                            txtCustCode.Focus()
                        End If
                    End If
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
                    '***
                End If
        End Select
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtECC_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtECC.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '****************************************************
        'Created By     -  Nisha
        'Description    -  At Enter Key Press Set Focus To Next Control
        '****************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                txtLST.Focus()
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
    Private Sub txtLoadingTaxType_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLoadingTaxType.TextChanged
        If Len(txtLoadingTaxType.Text) = 0 Then
            lblLoadingcharge_per.Text = "0.00"
        End If
    End Sub
    Private Sub txtLoadingTaxType_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtLoadingTaxType.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '****************************************************
        'Created By     -  Tapan
        'Description    -  At Enter Key Press Set Focus To Next Control
        '****************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        If txtECSSTaxType.Enabled Then txtECSSTaxType.Focus()
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
    Private Sub txtLoadingTaxType_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtLoadingTaxType.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F1 Then
            Call cmdLoadinfChageHelp_Click(cmdLoadinfChageHelp, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtLoadingTaxType_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtLoadingTaxType.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        If Len(txtLoadingTaxType.Text) > 0 Then
            If CheckExistanceOfFieldData((txtLoadingTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", " (Tx_TaxeID='LDT') AND UNIT_CODE='" & gstrUNITID & "' AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))") Then
                lblLoadingcharge_per.Text = CStr(GetTaxRate((txtLoadingTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " (Tx_TaxeID='LDT') AND UNIT_CODE='" & gstrUNITID & "'"))
                If txtECSSTaxType.Enabled Then txtECSSTaxType.Focus()
            Else
                Call ConfirmWindow(10248, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                Cancel = True
                txtLoadingTaxType.Text = ""
                If txtLoadingTaxType.Enabled Then txtLoadingTaxType.Focus()
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
        '****************************************************
        'Created By     -  Nisha
        'Description    -  At Enter Key Press Set Focus To Next Control
        '****************************************************
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
        '****************************************************
        'Created By     -  Nisha
        'Description    -  At F1 Key Press Display Help From Location Master
        '****************************************************
        On Error GoTo ErrHandler
        If KeyCode = 112 Then
            If CmdLocCodeHelp.Enabled Then Call CmdLocCodeHelp_Click(CmdLocCodeHelp, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub SelectInvTypeSubTypeFromSaleConf(ByRef pstrInvType As String, ByRef pstrInvSubtype As String)
        '****************************************************
        'Created By     -  Nisha
        'Description    -  Select Invoice Type,Sub Type From Sale Conf
        '****************************************************
        On Error GoTo ErrHandler
        Dim strSaleConfSql As String
        Dim rsSaleConf As ClsResultSetDB
        strSaleConfSql = "Select Invoice_Type,Sub_Type from SaleConf where UNIT_CODE='" & gstrUNITID & "' AND Description='" & Trim(pstrInvType) & "'"
        strSaleConfSql = strSaleConfSql & " and Sub_Type_Description='" & Trim(pstrInvSubtype) & "' and (fin_start_date <= getdate() and fin_end_date >= getdate()) "
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
        '****************************************************
        'Created By     -  Nisha
        'Description    -  Check Validity Of Location Code In The Location_Mst
        '****************************************************
        On Error GoTo ErrHandler
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                If Len(txtLocationCode.Text) > 0 Then
                    If CheckExistanceOfFieldData((txtLocationCode.Text), "Location_Code", "SalesChallan_Dtl", "UNIT_CODE='" & gstrUNITID & "'") Then
                        Call SelectDescriptionForField("Description", "Location_Code", "Location_Mst", lblLocCodeDes, (txtLocationCode.Text))
                        If txtChallanNo.Enabled Then
                            txtChallanNo.Focus()
                        Else
                            If txtCustCode.Enabled And txtCustCode.Visible Then txtCustCode.Focus()
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
                    If CheckExistanceOfFieldData((txtLocationCode.Text), "Location_Code", "SaleConf", "UNIT_CODE='" & gstrUNITID & "'") Then
                        Call SelectDescriptionForField("Description", "Location_Code", "Location_Mst", lblLocCodeDes, (txtLocationCode.Text))
                        If txtChallanNo.Enabled Then
                            txtChallanNo.Focus()
                        Else
                            If txtCustCode.Enabled Then txtCustCode.Focus()
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
    Private Sub txtLST_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtLST.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '****************************************************
        'Created By     -  Nisha
        'Description    -  At Enter Key Press Set Focus To Next Control
        '****************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                txtAddress1.Focus()
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
    Private Sub txtRefNo_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRefNo.TextChanged
        If Me.CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            SpChEntry.MaxRows = 0 : mstrItemCode = "" : txtRefNo.Focus()
        End If
        txtAmendNo.Text = ""
    End Sub
    Private Sub txtRefNo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtRefNo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '****************************************************
        'Created By     -  Nisha
        'Description    -  At Enter Key Press Set Focus To Next Control
        '****************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        ' focus set to carrier services instead of GRID
                        txtCarrServices.Focus()
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
        '****************************************************
        'Created By     -  Nisha
        'Description    -  If F1 Key Press Then Display Help From Customer Master/Vendor Master
        '****************************************************
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
        On Error GoTo ErrHandler
        If Len(txtLocationCode.Text) > 0 Then
            If Len(txtRefNo.Text) > 0 Then
                If SelectDataFromCustOrd_Dtl((txtCustCode.Text), (CmbInvType.Text)) Then
                    If CmbInvType.Text <> "REJECTION" Then
                        txtSRVDINO.Focus()
                        '****
                    Else
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
        '****************************************************
        'Created By     -  Nisha
        'Description    -  At Enter Key Press Set Focus To Next Control
        '****************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        With Me.SpChEntry
                            .Row = 1 : .Col = GridHeader.Quantity : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                        End With
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
    Private Sub txtSaleTaxType_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSaleTaxType.TextChanged
        If Len(txtSaleTaxType.Text) = 0 Then
            lblSaltax_Per.Text = "0.00"
        End If
    End Sub
    Private Sub txtSaleTaxType_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtSaleTaxType.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '****************************************************
        'Created By     -  Nisha
        'Description    -  At Enter Key Press Set Focus To Next Control
        '****************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        If Len(txtSaleTaxType.Text) > 0 Then
                            Call txtSaleTaxType_Validating(txtSaleTaxType, New System.ComponentModel.CancelEventArgs(False))
                        Else
                            If txtSurchargeTaxType.Enabled Then
                                txtSurchargeTaxType.Focus()
                            Else
                                txtLoadingTaxType.Focus()
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
        Dim strInvoiceType, strsql As String
        Dim rsChallanEntry, rsadditionalsurcharge, rsadditionaltax, rsadditionalVAT As ClsResultSetDB
        Dim flag As Boolean = False
        On Error GoTo ErrHandler
        txtSurchargeTaxType.Text = ""
        If Len(txtSaleTaxType.Text) > 0 Then
            If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                strInvoiceType = UCase(Trim(CmbInvType.Text))
            ElseIf (CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT) Or (CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW) Then
                rsChallanEntry = New ClsResultSetDB
                rsChallanEntry.GetResult("Select a.Description,a.Sub_Type_Description from SaleConf a,SalesChallan_Dtl b where a.UNIT_CODE=b.UNIT_CODE AND a.UNIT_CODE='" & gstrUNITID & "' AND Doc_No = " & txtChallanNo.Text & " and a.Invoice_Type = b.Invoice_type and a.Sub_type = b.Sub_Category and a.Location_code = b.Location_code and (fin_start_date <= getdate() and fin_end_date >= getdate())")
                strInvoiceType = UCase(rsChallanEntry.GetValue("Description"))
            End If
            If UCase(Trim(strInvoiceType)) <> "SERVICE INVOICE" Then
                If CheckExistanceOfFieldData((txtSaleTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", " (Tx_TaxeID='CST' OR Tx_TaxeID='LST' OR Tx_TaxeID='VAT') AND UNIT_CODE='" & gstrUNITID & "' AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))") Then
                    lblSaltax_Per.Text = CStr(GetTaxRate((txtSaleTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " (Tx_TaxeID='CST' OR Tx_TaxeID='LST' OR Tx_TaxeID='VAT') AND UNIT_CODE='" & gstrUNITID & "'"))
                    If txtSurchargeTaxType.Enabled Then
                        txtSurchargeTaxType.Focus()
                    Else
                        If txtLoadingTaxType.Enabled = True Then txtLoadingTaxType.Focus()
                    End If
                Else
                    Call ConfirmWindow(10248, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                    Cancel = True
                    txtSaleTaxType.Text = ""
                    If txtSaleTaxType.Enabled Then txtSaleTaxType.Focus()
                End If
            Else
                If CheckExistanceOfFieldData((txtSaleTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", " (Tx_TaxeID='SRT') AND UNIT_CODE='" & gstrUNITID & "' AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))") Then
                    lblSaltax_Per.Text = CStr(GetTaxRate((txtSaleTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " (Tx_TaxeID='SRT') AND UNIT_CODE='" & gstrUNITID & "'"))
                    If txtSurchargeTaxType.Enabled Then
                        txtSurchargeTaxType.Focus()
                    Else
                        txtLoadingTaxType.Focus()
                    End If
                Else
                    Call ConfirmWindow(10248, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                    Cancel = True
                    txtSaleTaxType.Text = ""
                    If txtSaleTaxType.Enabled Then txtSaleTaxType.Focus()
                End If
            End If
        End If
        If UCase(Trim(GetPlantName)) = "MATM" And UCase(strInvoiceType) = "NORMAL INVOICE" Then
            strsql = " select TxRt_Rate_No,TxRt_Percentage from Gen_TaxRate where (Tx_TaxeID='CST' OR Tx_TaxeID='LST') and txrt_percentage > 2.0 and TxRt_Rate_No='" & txtSaleTaxType.Text & "' AND UNIT_CODE='" & gstrUNITID & "' AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"
            rsadditionaltax = New ClsResultSetDB
            rsadditionaltax.GetResult(strsql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            If rsadditionaltax.GetNoRows > 0 Then
                rsadditionalsurcharge = New ClsResultSetDB
                strsql = " select TxRt_Rate_No,TxRt_Percentage from Gen_TaxRate where Tx_TaxeID='SsT' and txrt_percentage=5.0 AND UNIT_CODE='" & gstrUNITID & "' AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"
                rsadditionalsurcharge.GetResult(strsql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                If rsadditionalsurcharge.GetNoRows > 0 Then
                    txtSurchargeTaxType.Text = rsadditionalsurcharge.GetValue("TxRt_Rate_No")
                    lblSurcharge_Per.Text = rsadditionalsurcharge.GetValue("TxRt_Percentage")
                End If
                rsadditionalsurcharge.ResultSetClose()
                rsadditionalsurcharge = Nothing
            End If
            rsadditionaltax.ResultSetClose()
            rsadditionaltax = Nothing
            strsql = " select TxRt_Rate_No,TxRt_Percentage from Gen_TaxRate where (Tx_TaxeID='VAT') and TxRt_Rate_No='" & txtSaleTaxType.Text & "' AND UNIT_CODE='" & gstrUNITID & "' AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"
            rsadditionalVAT = New ClsResultSetDB
            rsadditionalVAT.GetResult(strsql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            If rsadditionalVAT.GetNoRows > 0 Then
                rsadditionalsurcharge = New ClsResultSetDB
                strsql = " select TxRt_Rate_No,TxRt_Percentage from Gen_TaxRate where Tx_TaxeID='SsT' and txrt_percentage=5.0 AND UNIT_CODE='" & gstrUNITID & "' AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"
                rsadditionalsurcharge.GetResult(strsql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                If rsadditionalsurcharge.GetNoRows > 0 Then
                    txtSurchargeTaxType.Text = rsadditionalsurcharge.GetValue("TxRt_Rate_No")
                    lblSurcharge_Per.Text = rsadditionalsurcharge.GetValue("TxRt_Percentage")
                End If
                rsadditionalsurcharge.ResultSetClose()
                rsadditionalsurcharge = Nothing
            End If
            rsadditionalVAT.ResultSetClose()
            rsadditionalVAT = Nothing
        End If
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtSECSSTaxType_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSECSSTaxType.TextChanged
        '-----------------------------------------------------------------------------------
        'Created By      : Davinder Singh
        'Issue ID        : 19575
        'Creation Date   : 27 Feb 2007
        '-----------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        If Len(Trim(txtSECSSTaxType.Text)) = 0 Then
            lblSECSStax_Per.Text = CStr(0)
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtSECSSTaxType_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtSECSSTaxType.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '-----------------------------------------------------------------------------------
        'Created By      : Davinder Singh
        'Issue ID        : 19575
        'Creation Date   : 27 Feb 2007
        'Function        : Help for New Tax SEcess
        '-----------------------------------------------------------------------------------
        If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then CmdSECSSTaxType.PerformClick()
    End Sub
    Private Sub txtSECSSTaxType_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtSECSSTaxType.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        If Len(txtSECSSTaxType.Text) > 0 Then
            If CheckExistanceOfFieldData((txtSECSSTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", " (Tx_TaxeID='ECSSH') AND UNIT_CODE='" & gstrUNITID & "' AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))") Then
                lblSECSStax_Per.Text = CStr(GetTaxRate((txtSECSSTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " (Tx_TaxeID='ECSSH') AND UNIT_CODE='" & gstrUNITID & "'"))
                If OptDiscountValue.Enabled Then OptDiscountValue.Focus()
            Else
                Call ConfirmWindow(10248, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                Cancel = True
                txtSECSSTaxType.Text = ""
                If txtSECSSTaxType.Enabled Then txtSECSSTaxType.Focus()
            End If
        End If
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtSRVDINO_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtSRVDINO.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '****************************************************
        'Created By     -  Tapan
        'Description    -  At Enter Key Press Set Focus To Next Control
        '****************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        With Me.SpChEntry
                            txtSRVLocation.Focus()
                        End With
                End Select
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        If KeyAscii = 13 Then
            CmdGrpChEnt.Focus()
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
    Private Sub txtSRVDINO_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtSRVDINO.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.F1 Then
            Call cmdhelpSRVDI_Click(cmdhelpSRVDI, New System.EventArgs())
        End If
    End Sub
    Private Sub txtSRVLocation_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtSRVLocation.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '****************************************************
        'Created By     -  Nisha
        'Description    -  At Enter Key Press Set Focus To Next Control
        '****************************************************
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
        If Trim(txtSurchargeTaxType.Text) = "" Then
            lblSurcharge_Per.Text = "0.00"
        End If
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
        '****************************************************
        'Created By     -  Tapan
        'Description    -  At Enter Key Press Set Focus To Next Control
        '****************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        With Me.SpChEntry
                            If txtLoadingTaxType.Enabled Then
                                txtLoadingTaxType.Focus()
                            Else
                                txtRemarks.Focus()
                            End If
                        End With
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
        If Trim(txtSurchargeTaxType.Text) <> "" Then
            If CheckExistanceOfFieldData((txtSurchargeTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", " Tx_TaxeID='SST' AND UNIT_CODE='" & gstrUNITID & "' AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))") Then
                lblSurcharge_Per.Text = CStr(GetTaxRate((txtSurchargeTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='SST' AND UNIT_CODE='" & gstrUNITID & "'"))
                If txtLoadingTaxType.Enabled = True Then txtLoadingTaxType.Focus()
            Else
                Call ConfirmWindow(10248, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                Cancel = True
                txtSurchargeTaxType.Text = ""
                If txtSurchargeTaxType.Enabled Then txtSurchargeTaxType.Focus()
            End If
        Else
            If txtLoadingTaxType.Enabled = True Then txtLoadingTaxType.Focus()
        End If
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtTCSTaxCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtTCSTaxCode.TextChanged
        If Len(txtTCSTaxCode.Text) = 0 Then
            lblTCSTaxPerDes.Text = "0.00"
        End If
    End Sub
    Private Sub txtTCSTaxCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtTCSTaxCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '****************************************************
        'Created By     -  Nisha
        'Description    -  At Enter Key Press Set Focus To Next Control
        '****************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        If chkExciseExumpted.Enabled Then chkExciseExumpted.Focus()
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
    Private Sub txtTCSTaxCode_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtTCSTaxCode.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = 112 Then
            If cmdHelpTCSTax.Enabled Then Call cmdHelpTCSTax_Click(cmdHelpTCSTax, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtTCSTaxCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtTCSTaxCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim strInvoiceType As String
        Dim rsChallanEntry As ClsResultSetDB
        On Error GoTo ErrHandler
        If Len(txtTCSTaxCode.Text) > 0 Then
            If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                strInvoiceType = UCase(Trim(CmbInvType.Text))
            ElseIf (CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT) Or (CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW) Then
                rsChallanEntry = New ClsResultSetDB
                rsChallanEntry.GetResult("Select a.Description,a.Sub_Type_Description from SaleConf a,SalesChallan_Dtl b where a.UNIT_CODE=b.UNIT_CODE AND a.UNIT_CODE='" & gstrUNITID & "' AND Doc_No = " & txtChallanNo.Text & " and a.Invoice_Type = b.Invoice_type and a.Sub_type = b.Sub_Category and a.Location_code = b.Location_code and (fin_start_date <= getdate() and fin_end_date >= getdate())")
                strInvoiceType = UCase(rsChallanEntry.GetValue("Description"))
            End If
            If CheckExistanceOfFieldData((txtTCSTaxCode.Text), "TxRt_Rate_No", "Gen_TaxRate", " (Tx_TaxeID='TCS') AND UNIT_CODE='" & gstrUNITID & "' AND ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))") Then
                lblTCSTaxPerDes.Text = CStr(GetTaxRate((txtTCSTaxCode.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " (Tx_TaxeID='TCS') AND UNIT_CODE='" & gstrUNITID & "'"))
                If chkExciseExumpted.Enabled Then chkExciseExumpted.Focus()
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
    Private Sub txtVehNo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtVehNo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '****************************************************
        'Created By     -  Nisha
        'Description    -  At Enter Key Press Set Focus To Next Control
        '****************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        Cmditems.Focus()
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
            strTableSql = "select " & Trim(pstrColumnName) & " from " & Trim(pstrTableName) & " where " & Trim(pstrColumnName) & "='" & Trim(pstrFieldText) & "' and " & pstrCondition
        Else
            strTableSql = "select " & Trim(pstrColumnName) & " from " & Trim(pstrTableName) & " where " & Trim(pstrColumnName) & "='" & Trim(pstrFieldText) & "'"
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
        Dim strSalesChallanDtl As String
        Dim strRGPNOs As String
        Dim strCustMst As String
        strSalesChallanDtl = "SELECT Transport_type,Vehicle_No,Account_Code,Cust_ref,Amendment_No,SalesTax_Type,"
        strSalesChallanDtl = strSalesChallanDtl & "Insurance,Invoice_Date,"
        strSalesChallanDtl = strSalesChallanDtl & "Invoice_Type,Sub_Category,Cust_Name,Carriage_Name,Frieght_Amount, "
        strSalesChallanDtl = strSalesChallanDtl & "Surcharge_salesTaxType,Amendment_No,ref_doc_no,Currency_Code,Exchange_Rate,"
        strSalesChallanDtl = strSalesChallanDtl & "Remarks,PerValue,SRVDINO,SRVLocation,LoadingChargeTaxType,discount_type,discount_amount,Discount_per,"
        strSalesChallanDtl = strSalesChallanDtl & "LoadingChargeTax_Per,ExciseExumpted,"
        strSalesChallanDtl = strSalesChallanDtl & "ConsigneeContactPerson,ConsigneeECCNo,ConsigneeLST,"
        strSalesChallanDtl = strSalesChallanDtl & "ConsigneeAddress1,ConsigneeAddress2,ConsigneeAddress3,"
        strSalesChallanDtl = strSalesChallanDtl & "USLOC,Schtime,TCSTax_Type,TCSTax_Per,TCSTaxAmount, ECESS_Type, ECESS_Per, ECESS_Amount, SECESS_Type, SECESS_Per, SECESS_Amount,payment_terms From Saleschallan_dtl"
        strSalesChallanDtl = strSalesChallanDtl & " WHERE UNIT_CODE='" & gstrUNITID & "' AND Location_Code ='"
        strSalesChallanDtl = strSalesChallanDtl & Trim(txtLocationCode.Text) & "' and Doc_No = " & Val(txtChallanNo.Text)
        rsGetData = New ClsResultSetDB
        rsGetData.GetResult(strSalesChallanDtl, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsGetData.GetNoRows > 0 Then
            GetDataInViewMode = True
            txtCustCode.Text = rsGetData.GetValue("Account_Code")
            If Not blnInvoiceAgainstMultipleSO Then
                txtRefNo.Text = rsGetData.GetValue("Cust_ref")
                txtAmendNo.Text = rsGetData.GetValue("Amendment_No")
                mstrAmmendmentNo = rsGetData.GetValue("Amendment_No")
            End If
            txtCarrServices.Text = rsGetData.GetValue("Carriage_Name")
            ctlInsurance.Text = rsGetData.GetValue("Insurance")
            txtFreight.Text = rsGetData.GetValue("Frieght_Amount")
            txtSaleTaxType.Text = rsGetData.GetValue("SalesTax_Type")
            Call txtSaleTaxType_Validating(txtSaleTaxType, New System.ComponentModel.CancelEventArgs(False))
            txtSurchargeTaxType.Text = rsGetData.GetValue("Surcharge_salesTaxType")
            Call txtSurchargeTaxType_Validating(txtSurchargeTaxType, New System.ComponentModel.CancelEventArgs(False))
            txtTCSTaxCode.Text = rsGetData.GetValue("TCSTax_Type")
            Call txtTCSTaxCode_Validating(txtTCSTaxCode, New System.ComponentModel.CancelEventArgs(False))
            strRGPNOs = rsGetData.GetValue("ref_doc_no")
            strRGPNOs = Replace(strRGPNOs, "§", ", ", 1)
            lblRGPDes.Text = strRGPNOs
            lblCustCodeDes.Text = rsGetData.GetValue("Cust_Name")
            lblDateDes.Text = VB6.Format(rsGetData.GetValue("Invoice_Date"), gstrDateFormat)
            mstrInvType = rsGetData.GetValue("Invoice_Type")
            mstrInvSubType = rsGetData.GetValue("Sub_Category")
            Call SetInvoicecategory(rsGetData.GetValue("Invoice_Type"), rsGetData.GetValue("Sub_Category"))
            ctlPerValue.Text = rsGetData.GetValue("PerValue")
            If Not blnInvoiceAgainstMultipleSO Then
                txtSRVDINO.Text = rsGetData.GetValue("SRVDINO")
                txtSRVLocation.Text = rsGetData.GetValue("SRVLocation")
                txtUsLoc.Text = rsGetData.GetValue("USLoc")
                txtSchTime.Text = rsGetData.GetValue("Schtime")
            End If
            txtECSSTaxType.Text = rsGetData.GetValue("ECESS_Type")
            Call txtECSSTaxType_Validating(txtECSSTaxType, New System.ComponentModel.CancelEventArgs(False))
            txtSECSSTaxType.Text = rsGetData.GetValue("SECESS_Type")
            Call txtSECSSTaxType_Validating(txtSECSSTaxType, New System.ComponentModel.CancelEventArgs(False))
            txtVehNo.Text = rsGetData.GetValue("vehicle_no")
            mstrNagareDate = Find_Value("select convert(varchar,sch_date,103) as sch_date  from mkt_enagaredtl where kanbanNo='" & txtSRVDINO.Text & "' AND UNIT_CODE='" & gstrUNITID & "'")
            mstrNagareDate = VB6.Format(mstrNagareDate, gstrDateFormat)
            If rsGetData.GetValue("Discount_Type") = False Then
                OptDiscountValue.Checked = True
                txtDiscountAmt.Text = rsGetData.GetValue("Discount_Amount")
            Else
                OptDiscountPercentage.Checked = True
                txtDiscountAmt.Text = rsGetData.GetValue("Discount_Per")
            End If
            If UCase(mstrInvType) = "EXP" Then
                lblCurrency.Visible = True : lblCurrencyDes.Visible = True
                lblCurrencyDes.Text = rsGetData.GetValue("Currency_code")
                lblExchangeRateLable.Visible = True : lblExchangeRateValue.Visible = True
            Else
            End If
            txtRemarks.Text = rsGetData.GetValue("Remarks")
            txtLoadingTaxType.Text = rsGetData.GetValue("LoadingChargeTaxType")
            lblLoadingcharge_per.Text = rsGetData.GetValue("LoadingChargeTax_Per")
            If rsGetData.GetValue("ExciseExumpted") Then
                chkExciseExumpted.CheckState = System.Windows.Forms.CheckState.Checked
            Else
                chkExciseExumpted.CheckState = System.Windows.Forms.CheckState.Unchecked
            End If
            txtContactPerson.Text = rsGetData.GetValue("ConsigneeContactPerson")
            txtECC.Text = rsGetData.GetValue("ConsigneeECCNo")
            txtLST.Text = rsGetData.GetValue("ConsigneeLST")
            txtAddress1.Text = rsGetData.GetValue("ConsigneeAddress1")
            txtAddress2.Text = rsGetData.GetValue("ConsigneeAddress2")
            txtAddress3.Text = rsGetData.GetValue("ConsigneeAddress3")
            cmdConsigneeDetails.Enabled = True
            mstrInvoiceType = rsGetData.GetValue("Invoice_Type")
            mstrInvoiceSubType = rsGetData.GetValue("Sub_Category")
            lblCurrencyDes.Text = rsGetData.GetValue("currency_code")
            lblCreditTerm.Text = IIf(IsDBNull(rsGetData.GetValue("payment_terms")), "", rsGetData.GetValue("payment_terms"))
            If Len(Trim(lblCreditTerm.Text)) > 0 Then
                Call SelectDescriptionForField("crTrm_desc", "crtrm_termID", "Gen_CreditTrmMaster", lblCreditTermDesc, Trim(lblCreditTerm.Text))
            Else
                lblCreditTermDesc.Text = ""
            End If
        Else
            GetDataInViewMode = False
        End If
        If Len(Trim(txtCustCode.Text)) > 0 Then
            rsCustMst = New ClsResultSetDB
            strCustMst = "SELECT Bill_Address1 + ', '  +  Bill_Address2 + ', ' + Bill_City + ' - ' + Bill_Pin as  invoiceAddress from Customer_Mst where Customer_code ='" & txtCustCode.Text & "' AND UNIT_CODE='" & gstrUNITID & "'"
            rsCustMst.GetResult(strCustMst)
            If rsCustMst.GetNoRows > 0 Then
                lblAddressDes.Text = rsCustMst.GetValue("InvoiceAddress")
            End If
            rsCustMst = Nothing
        End If
        '***
        rsGetData.ResultSetClose()
        rsGetData = Nothing
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Sub SetInvoicecategory(ByVal pstrInvType As String, ByVal pstrInvSubtype As String)
        'Created By       : Sandeep Chadha
        'Created On       : 24-May-2005
        'Description      : Show the Invoice Type & Sub Type in View & Edit mode
        On Error GoTo ErrHandler
        Dim intIndex As Short
        Dim strInvType As String
        Dim blnSelected As Boolean
        strInvType = UCase(Find_Value("Select Description from SaleConf where Invoice_Type='" & pstrInvType & "' AND UNIT_CODE='" & gstrUNITID & "'"))
        CmbInvType.Visible = True
        CmbInvType.Enabled = True
        CmbInvSubType.Visible = True
        CmbInvSubType.Enabled = True
        CmbInvType.SelectedIndex = -1
        CmbInvSubType.SelectedIndex = -1
        CmbInvSubType.Items.Clear()
        blnSelected = False
        For intIndex = 0 To CmbInvType.Items.Count - 1
            CmbInvType.SelectedIndex = intIndex
            If UCase(CmbInvType.Text) = strInvType Then
                blnSelected = True
                Call SelectInvoiceSubTypeFromSaleConf((CmbInvType.Text))
                Exit For
            End If
        Next
        If blnSelected = False Then
            CmbInvType.SelectedIndex = -1
        End If
        strInvType = UCase(Find_Value("Select Sub_Type_Description from SaleConf where Invoice_Type='" & pstrInvType & "' and Sub_Type='" & pstrInvSubtype & "' AND UNIT_CODE='" & gstrUNITID & "'"))
        blnSelected = False
        For intIndex = 0 To CmbInvSubType.Items.Count - 1
            CmbInvSubType.SelectedIndex = intIndex
            If UCase(CmbInvSubType.Text) = strInvType Then
                blnSelected = True
                Exit For
            End If
        Next
        If blnSelected = False Then
            CmbInvSubType.SelectedIndex = -1
        End If
        CmbInvType.DropDownStyle = ComboBoxStyle.DropDown
        CmbInvType.DropDownStyle = ComboBoxStyle.DropDown
        lblInvSubType.Visible = True
        lblInvType.Visible = True
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Function DisplayDetailsInSpread(ByRef pstrCurrency As String) As Boolean
        '****************************************************
        'Created By     -  Nisha
        'Description    -  To display Details From Sales_Dtl Acc To Location Code,Challan No and Drawing No
        '****************************************************
        On Error GoTo ErrHandler
        Dim intLoopCounter As Short
        Dim intRecordCount As Short
        Dim inti As Short
        Dim strsaledtl As String
        Dim dblPacking As Double
        Dim varItem_Code As Object
        Dim varItemAlready As Object
        Dim rsSalesDtl As ClsResultSetDB
        Dim rsTariffMst As ClsResultSetDB
        Dim rsSalesParameter As ClsResultSetDB
        Dim blnQtyChkAccToMeasureCode As Boolean
        Dim intDecimal As Short
        Dim strMin As String
        Dim strMax As String
        Dim intloopcounter1 As Short
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                strsaledtl = ""
                If blnInvoiceAgainstMultipleSO Then
                    strsaledtl = "SELECT * from Sales_Dtl WHERE UNIT_CODE='" & gstrUNITID & "' AND Location_Code='" & Trim(txtLocationCode.Text) & "'"
                    strsaledtl = strsaledtl & " and Doc_No=" & Val(txtChallanNo.Text)
                Else
                    strsaledtl = "SELECT * from Sales_Dtl WHERE UNIT_CODE='" & gstrUNITID & "' AND Location_Code='" & Trim(txtLocationCode.Text) & "'"
                    strsaledtl = strsaledtl & " and Doc_No=" & Val(txtChallanNo.Text) & " and Cust_Item_Code in(" & Trim(mstrItemCode) & ")"
                End If
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                If UCase(CStr(Trim(CmbInvType.Text))) = "NORMAL INVOICE" Or UCase(CStr(Trim(CmbInvType.Text))) = "JOBWORK INVOICE" Or UCase(CStr(Trim(CmbInvType.Text))) = "EXPORT INVOICE" Or UCase(CStr(Trim(CmbInvType.Text))) = "SERVICE INVOICE" Then
                    If UCase(Trim(CmbInvSubType.Text)) <> "SCRAP" Then
                        strsaledtl = ""
                        If blnInvoiceAgainstMultipleSO Then
                            strsaledtl = "Select Item_Code,Cust_DrgNo,Rate,Cust_Mtrl,Packing,Others,tool_Cost,Excise_Duty from Cust_ord_dtl WHERE "
                            strsaledtl = strsaledtl & "Account_Code ='" & txtCustCode.Text & "'and Cust_ref ='"
                            strsaledtl = strsaledtl & mstrRefNo & "' and Amendment_No = '" & mstrAmmNo & "'and "
                            strsaledtl = strsaledtl & " Active_flag ='A' and Cust_DrgNo in(" & mstrItemCode & ") AND UNIT_CODE='" & gstrUNITID & "'"
                        Else
                            strsaledtl = "Select Item_Code,Cust_DrgNo,Rate,Cust_Mtrl,Packing,Others,tool_Cost,Excise_Duty from Cust_ord_dtl WHERE "
                            strsaledtl = strsaledtl & "Account_Code ='" & txtCustCode.Text & "'and Cust_ref ='"
                            strsaledtl = strsaledtl & txtRefNo.Text & "' and Amendment_No = '" & mstrAmmNo & "'and "
                            strsaledtl = strsaledtl & " Active_flag ='A' and Cust_DrgNo in(" & mstrItemCode & ") AND UNIT_CODE='" & gstrUNITID & "'"
                        End If

                        'by abhijit
                        If gblnGSTUnit Then
                            '   strsaledtl = "Select Item_Code,Cust_DrgNo,Rate,Cust_Mtrl,Packing,Packing_Type,Others,tool_Cost,Excise_Duty,Markup_type ,MarkUp_value  from Cust_ord_dtl WHERE  UNIT_CODE='" + gstrUNITID + "' AND  "
                            'strsaledtl = "select Item_Code,Cust_DrgNo,Rate,Cust_Mtrl,Packing,Packing_Type,Others,tool_Cost,Excise_Duty,Markup_type ,MarkUp_value,IGSTTXRT_TYPE,IGSTTXRT_Percent=(SELECT TxRt_Percentage from Gen_TaxRate where TxRt_Rate_No= (select IGSTTXRT_TYPE FROM Cust_Ord_Dtl WHERE UNIT_CODE ='" + gstrUNITID + "' AND Account_Code = '" & txtCustCode.Text & "' AND Cust_Ref = '" & txtRefNo.Text & "' AND Amendment_No = '" & mstrAmmNo & "' AND Active_Flag = 'A' AND Cust_DrgNo=" & mstrItemCode & ")),CGSTTXRT_TYPE, CGSTTXRT_Percent=(SELECT TxRt_Percentage from Gen_TaxRate where TxRt_Rate_No= (select CGSTTXRT_TYPE FROM Cust_Ord_Dtl WHERE UNIT_CODE = '" + gstrUNITID + "' AND Account_Code = '" & txtCustCode.Text & "' AND Cust_Ref = '" & txtRefNo.Text & "' AND Amendment_No = '' AND Active_Flag = 'A' AND Cust_DrgNo=" & mstrItemCode & ")),SGSTTXRT_TYPE, SGSTTXRT_Percent=(SELECT TxRt_Percentage from Gen_TaxRate where TxRt_Rate_No= (select SGSTTXRT_TYPE FROM Cust_Ord_Dtl WHERE UNIT_CODE = '" + gstrUNITID + "' AND Account_Code = '" & txtCustCode.Text & "' AND Cust_Ref = '" & txtRefNo.Text & "' AND Amendment_No = '' AND Active_Flag = 'A' AND Cust_DrgNo=" & mstrItemCode & ")),UTGSTTXRT_TYPE, UTGSTTXRT_Percent=(SELECT TxRt_Percentage from Gen_TaxRate where TxRt_Rate_No= (select UTGSTTXRT_TYPE FROM Cust_Ord_Dtl WHERE UNIT_CODE ='" + gstrUNITID + "' AND Account_Code = '" & txtCustCode.Text & "' AND Cust_Ref = '" & txtRefNo.Text & "' AND Amendment_No = '' AND Active_Flag = 'A' AND Cust_DrgNo=" & mstrItemCode & ")),COMPENSATION_CESS,COMPENSATION_Percent=(SELECT TxRt_Percentage from Gen_TaxRate where TxRt_Rate_No= (select COMPENSATION_CESS FROM Cust_Ord_Dtl WHERE UNIT_CODE = '" + gstrUNITID + "' AND Account_Code = '" & txtCustCode.Text & "' AND Cust_Ref = '" & txtRefNo.Text & "' AND Amendment_No = '' AND Active_Flag = 'A' AND Cust_DrgNo=" & mstrItemCode & ")) from Cust_Ord_Dtl where UNIT_CODE = '" + gstrUNITID + "' AND Account_Code = '" & txtCustCode.Text & "' AND Cust_Ref = '" & txtRefNo.Text & "' AND Amendment_No = '' AND Active_Flag = 'A' AND Cust_DrgNo=" & mstrItemCode & ""
                            ''strsaledtl = "select Item_Code,Cust_DrgNo,Rate,Cust_Mtrl,Packing,Packing_Type,Others,tool_Cost,Excise_Duty,Markup_type ,MarkUp_value,HSNSACCODE,ISHSNORSAC,IGSTTXRT_TYPE,IGSTTXRT_Percent=(SELECT TxRt_Percentage from Gen_TaxRate where UNIT_CODE = '" + gstrUNITID + "' AND TxRt_Rate_No= (select IGSTTXRT_TYPE FROM Cust_Ord_Dtl WHERE UNIT_CODE ='" + gstrUNITID + "' AND Account_Code = '" & txtCustCode.Text & "' AND Active_Flag = 'A' AND Cust_DrgNo=" & mstrItemCode & ")),CGSTTXRT_TYPE, CGSTTXRT_Percent=(SELECT TxRt_Percentage from Gen_TaxRate where UNIT_CODE = '" + gstrUNITID + "' AND TxRt_Rate_No= (select CGSTTXRT_TYPE FROM Cust_Ord_Dtl WHERE UNIT_CODE = '" + gstrUNITID + "' AND Account_Code = '" & txtCustCode.Text & "' AND Active_Flag = 'A' AND Cust_DrgNo=" & mstrItemCode & ")),SGSTTXRT_TYPE, SGSTTXRT_Percent=(SELECT TxRt_Percentage from Gen_TaxRate where UNIT_CODE = '" + gstrUNITID + "' AND TxRt_Rate_No= (select SGSTTXRT_TYPE FROM Cust_Ord_Dtl WHERE UNIT_CODE = '" + gstrUNITID + "' AND Account_Code = '" & txtCustCode.Text & "' AND Active_Flag = 'A' AND Cust_DrgNo=" & mstrItemCode & ")),UTGSTTXRT_TYPE, UTGSTTXRT_Percent=(SELECT TxRt_Percentage from Gen_TaxRate where UNIT_CODE = '" + gstrUNITID + "' AND TxRt_Rate_No= (select UTGSTTXRT_TYPE FROM Cust_Ord_Dtl WHERE UNIT_CODE ='" + gstrUNITID + "' AND Account_Code = '" & txtCustCode.Text & "' AND Active_Flag = 'A' AND Cust_DrgNo=" & mstrItemCode & ")),COMPENSATION_CESS,COMPENSATION_Percent=(SELECT TxRt_Percentage from Gen_TaxRate where UNIT_CODE = '" + gstrUNITID + "' AND TxRt_Rate_No= (select COMPENSATION_CESS FROM Cust_Ord_Dtl WHERE UNIT_CODE = '" + gstrUNITID + "' AND Account_Code = '" & txtCustCode.Text & "' AND Active_Flag = 'A' AND Cust_DrgNo=" & mstrItemCode & ")) from Cust_Ord_Dtl where UNIT_CODE = '" + gstrUNITID + "' AND Account_Code = '" & txtCustCode.Text & "' AND Active_Flag = 'A' AND Cust_DrgNo=" & mstrItemCode & ""
                            strsaledtl = "Select D.Item_Code,D.Cust_DrgNo,D.Rate,D.Cust_Mtrl,D.Packing,D.Packing_Type,D.Others,D.tool_Cost,D.Excise_Duty,D.Markup_type ,D.MarkUp_value, D.HSNSACCODE, D.ISHSNORSAC, D.CGSTTXRT_TYPE, D.SGSTTXRT_TYPE, D.UTGSTTXRT_TYPE, D.IGSTTXRT_TYPE, D.COMPENSATION_CESS from Cust_ord_dtl D INNER JOIN CUST_ORD_HDR H ON D.UNIT_CODE=H.UNIT_CODE AND D.ACCOUNT_CODE=H.ACCOUNT_CODE AND D.CUST_REF=H.CUST_REF AND D.Amendment_No=H.Amendment_No WHERE "
                            strsaledtl = strsaledtl & "D.Account_Code ='" & txtCustCode.Text & "'and D.Cust_ref ='"
                            strsaledtl = strsaledtl & mstrRefNo & "' and D.Amendment_No = '" & mstrAmmNo & "'and "
                            strsaledtl = strsaledtl & " D.Active_flag ='A' and D.Cust_DrgNo in(" & mstrItemCode & ") AND D.UNIT_CODE='" & gstrUnitId & "'"
                        End If
                        'by abhijit

                    Else
                        strsaledtl = ""
                        strsaledtl = "SELECT Item_Code,standard_Rate from Item_Mst where "
                        strsaledtl = strsaledtl & " Status = 'A' and Hold_flag <> 1 and Item_Code in (" & mstrItemCode & ") AND UNIT_CODE='" & gstrUNITID & "'"
                    End If
                Else
                    If UCase(Trim(CmbInvType.Text)) = "REJECTION" Then
                        If Len(Trim(txtRefNo.Text)) > 0 Then
                            strsaledtl = ""
                            strsaledtl = "SELECT Item_Code,standard_Rate = item_Rate from grn_Dtl where "
                            strsaledtl = strsaledtl & " Item_Code in (" & mstrItemCode & ") and UNIT_CODE='" & gstrUNITID & "' AND Doc_No =" & txtRefNo.Text
                        Else
                            strsaledtl = ""
                            strsaledtl = "SELECT Item_Code,standard_Rate from Item_Mst where "
                            strsaledtl = strsaledtl & " Status = 'A' and Hold_flag <> 1 and Item_Code in (" & mstrItemCode & ") AND UNIT_CODE='" & gstrUNITID & "'"
                        End If
                    ElseIf UCase(Trim(CmbInvType.Text)) = "TRANSFER INVOICE" And UCase(Trim(CmbInvSubType.Text)) = "FINISHED GOODS" Then
                        strsaledtl = ""
                        strsaledtl = "SELECT Distinct a.Item_Code,c.Cust_drgNo,a.standard_Rate FROM Item_Mst a,Itembal_Mst b,CustItem_Mst c "
                        strsaledtl = strsaledtl & " where a.Item_Code=b.Item_Code and a.Item_Code = c.ITem_Code "
                        strsaledtl = strsaledtl & " and a.Status ='A' and a.Hold_Flag <> 1 and c.Account_code ='" & txtCustCode.Text & "'"
                        strsaledtl = strsaledtl & " and a.Item_code in (" & mstrItemCode & ") AND a.UNIT_CODE=b.UNIT_CODE AND a.UNIT_CODE=c.UNIT_CODE AND a.UNIT_CODE='" & gstrUNITID & "'"
                    Else
                        strsaledtl = ""
                        strsaledtl = "SELECT Item_Code,standard_Rate from Item_Mst where "
                        strsaledtl = strsaledtl & " Status = 'A' and Hold_flag <> 1 and Item_Code in (" & mstrItemCode & ") AND UNIT_CODE='" & gstrUNITID & "'"
                    End If
                End If
        End Select
        rsSalesDtl = New ClsResultSetDB
        rsSalesDtl.GetResult(strsaledtl, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        Dim intLoopCount As Short
        Dim varCumulative As Object
        Dim strCustDrgNo As Object
        Dim strSqlBins As String
        Dim dblBins As Double
        Dim rsBinQty As ClsResultSetDB
        If rsSalesDtl.GetNoRows > 0 Then
            If blnInvoiceAgainstMultipleSO Then
                mIntRecordCount = mIntRecordCount + rsSalesDtl.GetNoRows
                intRecordCount = rsSalesDtl.GetNoRows
                ReDim Preserve mdblPrevQty(mIntRecordCount - 1) ' To get value of Quantity in Arrey for updation in despatch
                ReDim Preserve mdblToolCost(mIntRecordCount - 1) ' To get value of Quantity i
            Else
                intRecordCount = rsSalesDtl.GetNoRows
                ReDim mdblPrevQty(intRecordCount - 1) ' To get value of Quantity in Arrey for updation in despatch
                ReDim mdblToolCost(intRecordCount - 1) ' To get value of Quantity i
            End If
            If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                If SpChEntry.MaxRows > 0 Then
                    varItemAlready = Nothing
                    Call SpChEntry.GetText(GridHeader.InternalPartNo, 1, varItemAlready)
                    If Len(Trim(varItemAlready)) = 0 Then
                        Call addRowAtEnterKeyPress(intRecordCount - 1)
                    End If
                Else
                    Call addRowAtEnterKeyPress(intRecordCount)
                End If
            Else
                If blnInvoiceAgainstMultipleSO Then
                    SpChEntry.MaxRows = 0
                    Call addRowAtEnterKeyPress(intRecordCount)
                Else
                    Call addRowAtEnterKeyPress(intRecordCount - 1)
                End If
            End If
            rsSalesDtl.MoveFirst()
            If CmbInvType.Text = "NORMAL INVOICE" Or CmbInvType.Text = "JOBWORK INVOICE" Or CmbInvType.Text = "EXPORT INVOICE" Or CmbInvType.Text = "SERVICE INVOICE" Then
                If UCase(Trim(CmbInvSubType.Text)) <> "SCRAP" Then
                    If blnInvoiceAgainstMultipleSO Then
                        For intLoopCount = UBound(mdblToolCost) + 1 To mIntRecordCount
                            If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                                mdblToolCost(intLoopCount - 1) = Val(rsSalesDtl.GetValue("Tool_Cost"))
                            Else
                                mdblToolCost(intLoopCount - 1) = Val(rsSalesDtl.GetValue("Tool_Cost"))
                            End If
                            rsSalesDtl.MoveNext()
                        Next
                    Else
                        For intLoopCount = 1 To intRecordCount
                            If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                                mdblToolCost(intLoopCount - 1) = Val(rsSalesDtl.GetValue("Tool_Cost"))
                            Else
                                mdblToolCost(intLoopCount - 1) = Val(rsSalesDtl.GetValue("Tool_Cost"))
                            End If
                            rsSalesDtl.MoveNext()
                        Next
                    End If
                End If
            End If
            rsSalesDtl.MoveFirst()
            intDecimal = ToGetDecimalPlaces(pstrCurrency)
            If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                If SpChEntry.MaxRows > 0 Then
                    varItemAlready = Nothing
                    Call SpChEntry.GetText(GridHeader.InternalPartNo, 1, varItemAlready)
                    If Len(Trim(varItemAlready)) > 0 Then
                        inti = SpChEntry.MaxRows + 1
                        SpChEntry.MaxRows = SpChEntry.MaxRows + intRecordCount
                        intRecordCount = SpChEntry.MaxRows
                    Else
                        inti = 1
                    End If
                Else
                    inti = 1
                    SpChEntry.MaxRows = intRecordCount
                End If
            Else
                inti = 1
            End If
            Call ChangeCellTypeStaticText()
            Call SetMaxLengthInSpread(intDecimal)
            For intLoopCounter = inti To intRecordCount
                With Me.SpChEntry
                    Select Case Me.CmdGrpChEnt.Mode
                        Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                            .Row = 1 : .Row2 = .MaxRows : .Col = 0 : .Col2 = .MaxCols
                            .Enabled = True : .BlockMode = True : .Lock = True : .BlockMode = False
                            Call .SetText(GridHeader.InternalPartNo, intLoopCounter, rsSalesDtl.GetValue("Item_Code"))
                            Call .SetText(GridHeader.CustPartNo, intLoopCounter, rsSalesDtl.GetValue("Cust_Item_Code"))
                            Call .SetText(GridHeader.RatePerUnit, intLoopCounter, rsSalesDtl.GetValue("Rate") * Val(ctlPerValue.Text))
                            Call .SetText(GridHeader.Rate, intLoopCounter, rsSalesDtl.GetValue("Rate"))
                            Call .SetText(GridHeader.CustSuppMatPerUnit, intLoopCounter, rsSalesDtl.GetValue("Cust_Mtrl") * Val(ctlPerValue.Text))
                            Call .SetText(GridHeader.CustMtrl, intLoopCounter, rsSalesDtl.GetValue("Cust_Mtrl"))
                            Call .SetText(GridHeader.Quantity, intLoopCounter, rsSalesDtl.GetValue("Sales_Quantity"))
                            mdblPrevQty(intLoopCounter - 1) = Nothing
                            Call .GetText(GridHeader.Quantity, intLoopCounter, mdblPrevQty(intLoopCounter - 1))
                            If blnInvoiceAgainstMultipleSO Then
                                Call .SetText(GridHeader.CustRefNo, intLoopCounter, rsSalesDtl.GetValue("Cust_ref"))
                                Call .SetText(GridHeader.AmendmentNo, intLoopCounter, rsSalesDtl.GetValue("Amendment_No"))
                                Call .SetText(GridHeader.srvdino, intLoopCounter, rsSalesDtl.GetValue("SRVDINO"))
                                Call .SetText(GridHeader.SRVLocation, intLoopCounter, rsSalesDtl.GetValue("SRVLocation"))
                                Call .SetText(GridHeader.USLOC, intLoopCounter, rsSalesDtl.GetValue("USLOC"))
                                Call .SetText(GridHeader.SChTime, intLoopCounter, rsSalesDtl.GetValue("SchTime"))
                            End If
                            Call .SetText(GridHeader.Packing, intLoopCounter, rsSalesDtl.GetValue("Packing"))
                            Call .SetText(GridHeader.EXC, intLoopCounter, rsSalesDtl.GetValue("Excise_Type"))
                            Call .SetText(GridHeader.CVD, intLoopCounter, rsSalesDtl.GetValue("CVD_type"))
                            Call .SetText(GridHeader.SAD, intLoopCounter, rsSalesDtl.GetValue("SAD_type"))
                            Call .SetText(GridHeader.OthersPerUnit, intLoopCounter, (rsSalesDtl.GetValue("Others") * Val(ctlPerValue.Text)) / rsSalesDtl.GetValue("Sales_Quantity"))
                            Call .SetText(GridHeader.Others, intLoopCounter, rsSalesDtl.GetValue("Others"))
                            Call .SetText(GridHeader.FromBox, intLoopCounter, rsSalesDtl.GetValue("From_Box"))
                            Call .SetText(GridHeader.ToBox, intLoopCounter, rsSalesDtl.GetValue("To_Box"))
                            Call .SetText(GridHeader.ToolCostPerUnit, intLoopCounter, rsSalesDtl.GetValue("tool_Cost") * Val(ctlPerValue.Text))
                            Call .SetText(GridHeader.ToolCost, intLoopCounter, rsSalesDtl.GetValue("tool_Cost"))
                            If intLoopCounter = 1 Then
                                Call .SetText(GridHeader.CumulativeBoxes, intLoopCounter, (rsSalesDtl.GetValue("To_Box") - rsSalesDtl.GetValue("From_Box")) + 1)
                            Else
                                varCumulative = Nothing
                                Call .GetText(GridHeader.CumulativeBoxes, intLoopCounter - 1, varCumulative)
                                Call .SetText(GridHeader.CumulativeBoxes, intLoopCounter, varCumulative + ((rsSalesDtl.GetValue("To_Box") - rsSalesDtl.GetValue("From_Box")) + 1))
                            End If
                            Call .SetText(GridHeader.BinQty, intLoopCounter, rsSalesDtl.GetValue("BinQuantity"))

                            'Abhijit()
                            If gblnGSTUnit Then

                                Call .SetText(GridHeader.ISHSNORSAC, intLoopCounter, rsSalesDtl.GetValue("ISHSNORSAC"))
                                Call .SetText(GridHeader.HSNSACCODE, intLoopCounter, rsSalesDtl.GetValue("HSNSACCODE"))

                                Call .SetText(GridHeader.Basic_Value, intLoopCounter, rsSalesDtl.GetValue("Basic_Amount"))
                                Call .SetText(GridHeader.Discount_Percent, intLoopCounter, rsSalesDtl.GetValue("Discount_perc"))
                                Call .SetText(GridHeader.Discount_Amt, intLoopCounter, rsSalesDtl.GetValue("Discount_amt"))

                                Call .SetText(GridHeader.Accessible_Amt, intLoopCounter, rsSalesDtl.GetValue("Accessible_amount"))

                                Call .SetText(GridHeader.CGST_Type, intLoopCounter, rsSalesDtl.GetValue("CGSTTXRT_TYPE"))
                                Call .SetText(GridHeader.CGST_Percent, intLoopCounter, rsSalesDtl.GetValue("CGST_PERCENT"))
                                Call .SetText(GridHeader.CGST_Amt, intLoopCounter, rsSalesDtl.GetValue("CGST_AMT"))

                                Call .SetText(GridHeader.IGST_Type, intLoopCounter, rsSalesDtl.GetValue("IGSTTXRT_TYPE"))
                                Call .SetText(GridHeader.IGST_Percent, intLoopCounter, rsSalesDtl.GetValue("IGST_PERCENT"))
                                Call .SetText(GridHeader.IGST_Amt, intLoopCounter, rsSalesDtl.GetValue("IGST_AMT"))

                                Call .SetText(GridHeader.SGST_Type, intLoopCounter, rsSalesDtl.GetValue("SGSTTXRT_TYPE"))
                                Call .SetText(GridHeader.SGST_Percent, intLoopCounter, rsSalesDtl.GetValue("SGST_PERCENT"))
                                Call .SetText(GridHeader.SGST_Amt, intLoopCounter, rsSalesDtl.GetValue("SGST_AMT"))

                                Call .SetText(GridHeader.UTGST_Type, intLoopCounter, rsSalesDtl.GetValue("UTGSTTXRT_TYPE"))
                                Call .SetText(GridHeader.UTGST_Percent, intLoopCounter, rsSalesDtl.GetValue("UTGST_PERCENT"))
                                Call .SetText(GridHeader.UTGST_Amt, intLoopCounter, rsSalesDtl.GetValue("UTGST_AMT"))

                                Call .SetText(GridHeader.CSESS_TAX_Type, intLoopCounter, rsSalesDtl.GetValue("COMPENSATION_CESS_TYPE"))
                                Call .SetText(GridHeader.CSESS_TAX_Percent, intLoopCounter, rsSalesDtl.GetValue("COMPENSATION_CESS_PERCENT"))
                                Call .SetText(GridHeader.CSESS_TAX_VAL, intLoopCounter, rsSalesDtl.GetValue("COMPENSATION_CESS_AMT"))

                                Call .SetText(GridHeader.Item_Total, intLoopCounter, rsSalesDtl.GetValue("ITEM_VALUE"))
                            End If


                            ''Abhijit()

                        Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                            .Enabled = True
                            .Row = 1 : .Row2 = .MaxRows : .Col = 0 : .Col2 = .MaxCols : .BlockMode = True : .Lock = False : .BlockMode = False
                            If (Trim(CmbInvType.Text) = "NORMAL INVOICE") Or (Trim(CmbInvType.Text) = "JOBWORK INVOICE") Or (Trim(CmbInvType.Text) = "EXPORT INVOICE") Or (Trim(CmbInvType.Text) = "SERVICE INVOICE") Then
                                If CBool(UCase(CStr((Trim(CmbInvSubType.Text)) <> "SCRAP"))) Then
                                    Call .SetText(GridHeader.InternalPartNo, intLoopCounter, rsSalesDtl.GetValue("Item_Code"))
                                    Call .SetText(GridHeader.CustPartNo, intLoopCounter, rsSalesDtl.GetValue("Cust_DrgNo"))
                                    Call .SetText(GridHeader.RatePerUnit, intLoopCounter, (Val(rsSalesDtl.GetValue("Rate")) * Val(ctlPerValue.Text)))
                                    Call .SetText(GridHeader.Rate, intLoopCounter, Val(rsSalesDtl.GetValue("Rate")))
                                    If blnInvoiceAgainstMultipleSO Then
                                        Call .SetText(GridHeader.Quantity, intLoopCounter, mstrQuantity)
                                        Call .SetText(GridHeader.CustRefNo, intLoopCounter, mstrRefNo)
                                        Call .SetText(GridHeader.AmendmentNo, intLoopCounter, mstrAmmNo)
                                        Call .SetText(GridHeader.srvdino, intLoopCounter, mstrSRVDINo)
                                        Call .SetText(GridHeader.SRVLocation, intLoopCounter, mstrSRVLocation)
                                        Call .SetText(GridHeader.USLOC, intLoopCounter, mstrUSLoc)
                                        Call .SetText(GridHeader.SChTime, intLoopCounter, mstrSchTime)
                                    End If
                                    Call .SetText(GridHeader.CustSuppMatPerUnit, intLoopCounter, (Val(rsSalesDtl.GetValue("Cust_Mtrl")) * Val(ctlPerValue.Text)))
                                    Call .SetText(GridHeader.CustMtrl, intLoopCounter, Val(rsSalesDtl.GetValue("Cust_Mtrl")))
                                    Call .SetText(GridHeader.Packing, intLoopCounter, rsSalesDtl.GetValue("Packing"))
                                    Call .SetText(GridHeader.EXC, intLoopCounter, rsSalesDtl.GetValue("Excise_duty")) ''Changed By Tapan due to column changed in Grid
                                    Call .SetText(GridHeader.OthersPerUnit, intLoopCounter, (Val(rsSalesDtl.GetValue("Others")) * Val(ctlPerValue.Text)))
                                    Call .SetText(GridHeader.Others, intLoopCounter, Val(rsSalesDtl.GetValue("Others")))
                                    Call .SetText(GridHeader.ToolCostPerUnit, intLoopCounter, (Val(rsSalesDtl.GetValue("tool_cost")) * Val(ctlPerValue.Text))) ''Changed By Tapan due to column changed in Grid
                                    Call .SetText(GridHeader.ToolCost, intLoopCounter, Val(rsSalesDtl.GetValue("tool_cost"))) ''Changed By Tapan due to column changed in Grid
                                Else
                                    Call .SetText(GridHeader.InternalPartNo, intLoopCounter, rsSalesDtl.GetValue("Item_Code"))
                                    Call .SetText(GridHeader.CustPartNo, intLoopCounter, rsSalesDtl.GetValue("Item_code"))
                                    Call .SetText(GridHeader.RatePerUnit, intLoopCounter, (Val(rsSalesDtl.GetValue("Standard_Rate")) * Val(ctlPerValue.Text)))
                                    Call .SetText(GridHeader.Rate, intLoopCounter, Val(rsSalesDtl.GetValue("Standard_Rate")))
                                End If
                            Else
                                Call .SetText(GridHeader.InternalPartNo, intLoopCounter, rsSalesDtl.GetValue("Item_Code"))
                                If UCase(CmbInvType.Text) = "TRANSFER INVOICE" And UCase(CmbInvSubType.Text) = "FINISHED GOODS" Then
                                    Call .SetText(GridHeader.CustPartNo, intLoopCounter, rsSalesDtl.GetValue("cust_DrgNo"))
                                Else
                                    Call .SetText(GridHeader.CustPartNo, intLoopCounter, rsSalesDtl.GetValue("Item_code"))
                                End If
                                Call .SetText(GridHeader.RatePerUnit, intLoopCounter, (rsSalesDtl.GetValue("Standard_Rate") * Val(ctlPerValue.Text)))
                                Call .SetText(GridHeader.Rate, intLoopCounter, rsSalesDtl.GetValue("Standard_Rate"))
                            End If
                            rsBinQty = New ClsResultSetDB
                            strCustDrgNo = Nothing
                            Call SpChEntry.GetText(GridHeader.CustPartNo, intLoopCounter, strCustDrgNo)
                            strSqlBins = "Select isnull(BinQuantity,1) as BinQuantity from custitem_mst where cust_drgno= '" & strCustDrgNo & "' and Account_code='" & Trim(Me.txtCustCode.Text) & "' AND UNIT_CODE='" & gstrUNITID & "' "
                            rsBinQty.GetResult(strSqlBins)
                            If rsBinQty.GetNoRows > 0 Then
                                dblBins = rsBinQty.GetValue("BinQuantity")
                            Else
                                dblBins = 1
                            End If
                            Call SpChEntry.SetText(GridHeader.BinQty, intLoopCounter, dblBins)

                            'Abhijit()
                            'If gblnGSTUnit Then

                            '    'Call .SetText(GridHeader.HSNSACCODE, intLoopCounter, rsSalesDtl.GetValue("HSNSACCODE"))
                            '    'Call .SetText(GridHeader.ISHSNORSAC, intLoopCounter, rsSalesDtl.GetValue("ISHSNORSAC"))
                            '    'Call .SetText(GridHeader.CGST_Type, intLoopCounter, rsSalesDtl.GetValue("CGSTTXRT_TYPE"))
                            '    'Call .SetText(GridHeader.CGST_Percent, intLoopCounter, rsSalesDtl.GetValue("CGSTTXRT_Percent"))
                            '    'Call .SetText(GridHeader.IGST_Type, intLoopCounter, rsSalesDtl.GetValue("IGSTTXRT_TYPE"))
                            '    'Call .SetText(GridHeader.IGST_Percent, intLoopCounter, rsSalesDtl.GetValue("IGSTTXRT_Percent"))
                            '    'Call .SetText(GridHeader.SGST_Type, intLoopCounter, rsSalesDtl.GetValue("SGSTTXRT_TYPE"))
                            '    'Call .SetText(GridHeader.SGST_Percent, intLoopCounter, rsSalesDtl.GetValue("SGSTTXRT_Percent"))
                            '    'Call .SetText(GridHeader.UTGST_Type, intLoopCounter, rsSalesDtl.GetValue("UTGSTTXRT_TYPE"))
                            '    'Call .SetText(GridHeader.UTGST_Percent, intLoopCounter, rsSalesDtl.GetValue("UTGSTTXRT_Percent"))
                            '    'Call .SetText(GridHeader.CSESS_TAX_Type, intLoopCounter, rsSalesDtl.GetValue("COMPENSATION_CESS"))
                            '    'Call .SetText(GridHeader.CSESS_TAX_Percent, intLoopCounter, rsSalesDtl.GetValue("COMPENSATION_Percent"))
                            '    'CalculateGSTTaxes_onGRID(intLoopCounter)
                            'End If


                            If gblnGSTUnit Then

                                Dim HSNSAC_CODE_TYPE_so As String = ""
                                Dim ISHSNORSAC_so As String = ""
                                Dim CGST_TXRT_HEAD_so As String = ""
                                Dim SGST_TXRT_HEAD_so As String = ""
                                Dim UGST_TXRT_HEAD_so As String = ""
                                Dim IGST_TXRT_HEAD_so As String = ""
                                Dim COMPENSATION_CESS_so As String = ""

                                HSNSAC_CODE_TYPE_so = rsSalesDtl.GetValue("HSNSACCODE")
                                ISHSNORSAC_so = rsSalesDtl.GetValue("ISHSNORSAC")

                                CGST_TXRT_HEAD_so = rsSalesDtl.GetValue("CGSTTXRT_TYPE")
                                SGST_TXRT_HEAD_so = rsSalesDtl.GetValue("SGSTTXRT_TYPE")
                                UGST_TXRT_HEAD_so = rsSalesDtl.GetValue("UTGSTTXRT_TYPE")
                                IGST_TXRT_HEAD_so = rsSalesDtl.GetValue("IGSTTXRT_TYPE")
                                COMPENSATION_CESS_so = rsSalesDtl.GetValue("COMPENSATION_CESS")
                                'txtTCSTaxCode.Text = rsSalesDtl.GetValue("TCSTAX_TYPE")
                                'Call txtTCSTaxCode_Validating(txtTCSTaxCode, New System.ComponentModel.CancelEventArgs(False))

                                Dim GST_rate_percent_so As ClsResultSetDB
                                Dim strSql_GST_rate_percent_so As String = ""
                                GST_rate_percent_so = New ClsResultSetDB
                                strSql_GST_rate_percent_so = "select * from UDF_GST_TAX_RATE_PERCENT('" & gstrUNITID & "','" & CGST_TXRT_HEAD_so & "','" & SGST_TXRT_HEAD_so & "', '" & IGST_TXRT_HEAD_so & "', '" & UGST_TXRT_HEAD_so & "', '" & COMPENSATION_CESS_so & "')"

                                Call .SetText(GridHeader.HSNSACCODE, intLoopCounter, rsSalesDtl.GetValue("HSNSACCODE"))
                                Call .SetText(GridHeader.ISHSNORSAC, intLoopCounter, rsSalesDtl.GetValue("ISHSNORSAC"))
                                Call .SetText(GridHeader.CGST_Type, intLoopCounter, rsSalesDtl.GetValue("CGSTTXRT_TYPE"))
                                Call .SetText(GridHeader.SGST_Type, intLoopCounter, rsSalesDtl.GetValue("SGSTTXRT_TYPE"))
                                Call .SetText(GridHeader.IGST_Type, intLoopCounter, rsSalesDtl.GetValue("IGSTTXRT_TYPE"))
                                Call .SetText(GridHeader.UTGST_Type, intLoopCounter, rsSalesDtl.GetValue("UTGSTTXRT_TYPE"))
                                Call .SetText(GridHeader.CSESS_TAX_Type, intLoopCounter, rsSalesDtl.GetValue("COMPENSATION_CESS"))

                                GST_rate_percent_so.GetResult(strSql_GST_rate_percent_so)

                                If GST_rate_percent_so.GetNoRows > 0 Then
                                    Call .SetText(GridHeader.CGST_Percent, intLoopCounter, GST_rate_percent_so.GetValue("CGST_PERCENT"))
                                    Call .SetText(GridHeader.IGST_Percent, intLoopCounter, GST_rate_percent_so.GetValue("IGST_PERCENT"))
                                    Call .SetText(GridHeader.SGST_Percent, intLoopCounter, GST_rate_percent_so.GetValue("SGST_PERCENT"))
                                    Call .SetText(GridHeader.UTGST_Percent, intLoopCounter, GST_rate_percent_so.GetValue("UTGST_PERCENT"))
                                    Call .SetText(GridHeader.CSESS_TAX_Percent, intLoopCounter, GST_rate_percent_so.GetValue("COMPENSATION_CESS_PERCENT"))
                                Else
                                    Call .SetText(GridHeader.CGST_Percent, intLoopCounter, 0)
                                    Call .SetText(GridHeader.IGST_Percent, intLoopCounter, 0)
                                    Call .SetText(GridHeader.SGST_Percent, intLoopCounter, 0)
                                    Call .SetText(GridHeader.UTGST_Percent, intLoopCounter, 0)
                                    Call .SetText(GridHeader.CSESS_TAX_Percent, intLoopCounter, 0)
                                End If
                                CalculateGSTTaxes(intLoopCounter)
                            End If

                            '''''''ABHIJIT

                            'Abhijit()
                    End Select
                End With
                rsSalesDtl.MoveNext()
            Next intLoopCounter
            Call SetMaxLengthInSpread(intDecimal)
        End If
        If SpChEntry.MaxRows > 3 Then
            SpChEntry.ScrollBars = FPSpreadADO.ScrollBarsConstants.ScrollBarsBoth
        End If
        rsSalesDtl.ResultSetClose()
        rsSalesDtl = Nothing
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Function ValidatebeforeSave(ByRef pstrMode As String) As Boolean
        '*****************************************************
        'Created By     -  Kapil
        'Description    -  To Check the Blank Fields In The Form
        '*****************************************************
        On Error GoTo ErrHandler
        Dim lstrControls As String
        Dim lNo As Integer
        Dim lctrFocus As System.Windows.Forms.Control
        Dim strSQL As String
        Dim intiLoopCount As Integer
        Dim strCustDrgNoLists As String = ""

        ValidatebeforeSave = True
        lNo = 1
        lstrControls = ResolveResString(10059)
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
                If (Len(Me.txtCustCode.Text) = 0) Then
                    lstrControls = lstrControls & vbCrLf & lNo & ". Customer Code."
                    lNo = lNo + 1
                    If lctrFocus Is Nothing Then
                        lctrFocus = Me.txtCustCode
                    End If
                    ValidatebeforeSave = False
                End If
                If Not DateIsAppropriate() Then
                    lstrControls = lstrControls & vbCrLf & lNo & ". Date specified either Falls Before the LAST Invoice Date or is Greater than Todays Date."
                    lNo = lNo + 1
                    If lctrFocus Is Nothing Then
                        lctrFocus = Me.txtCustCode
                    End If
                    ValidatebeforeSave = False
                End If
                If Val(lblECSStax_Per.Text) > 0 Then
                    If Len(Trim(txtSECSSTaxType.Text)) = 0 Then
                        lstrControls = lstrControls & vbCrLf & lNo & ". Secondary ECESS"
                        lNo = lNo + 1
                        If lctrFocus Is Nothing Then
                            lctrFocus = txtSECSSTaxType
                        End If
                        ValidatebeforeSave = False
                    End If
                End If
                If gblnGSTUnit = True And txtTCSTaxCode.Enabled = True Then
                    If txtTCSTaxCode.Text.Trim.ToString.Length = 0 Then
                        lstrControls = lstrControls & vbCrLf & lNo & ". TCS Tax."
                        lNo = lNo + 1
                        If lctrFocus Is Nothing Then
                            lctrFocus = Me.txtTCSTaxCode
                        End If
                        ValidatebeforeSave = False
                    End If
                End If

                If (UCase(Trim(CmbInvType.Text)) = "NORMAL INVOICE") Or (UCase(Trim(CmbInvType.Text)) = "JOBWORK INVOICE") Or (UCase(Trim(CmbInvType.Text)) = "EXPORT INVOICE") Or (UCase(Trim(CmbInvType.Text)) = "SERVICE INVOICE") Then
                    If (Trim(CmbInvSubType.Text) <> "SCRAP") Then
                        If Not blnInvoiceAgainstMultipleSO Then
                            If (Len(Me.txtRefNo.Text) = 0) Then
                                lstrControls = lstrControls & vbCrLf & lNo & ". Reference No.."
                                lNo = lNo + 1
                                If lctrFocus Is Nothing Then
                                    lctrFocus = Me.CmdRefNoHelp
                                End If
                                ValidatebeforeSave = False
                            End If
                        End If
                    End If
                    If blnFIFO = False Then
                        If (Trim(CmbInvType.Text) = "JOBWORK INVOICE") Then
                            If (Len(mstrRGP) = 0) Then
                                lstrControls = lstrControls & vbCrLf & lNo & ". RGP No.."
                                lNo = lNo + 1
                                If lctrFocus Is Nothing Then
                                    lctrFocus = Me.CmdRefNoHelp
                                End If
                                ValidatebeforeSave = False
                            End If
                        End If
                    End If
                End If
                If SpChEntry.MaxRows = 0 Then
                    lstrControls = lstrControls & vbCrLf & lNo & ". Select Items"
                    lNo = lNo + 1
                    If lctrFocus Is Nothing Then
                        lctrFocus = Cmditems
                    End If
                    ValidatebeforeSave = False
                End If
                If (Len(Me.txtFreight.Text) = 0) Then
                    txtFreight.Text = "0.00"
                End If
                If (Len(Me.txtSurchargeTaxType.Text) = 0) Then
                    ''            txtSurchargeTaxType.Text = "0.00"
                End If
                If (Len(Me.ctlInsurance.Text) = 0) Then
                    ctlInsurance.Text = "0.00"
                End If
                If (Len(lblCurrencyDes.Text) = 0) Then
                    lblCurrencyDes.Text = gstrCURRENCYCODE
                End If
                '10808160 
                strSQL = "select dbo.UDF_ISEOPINVOICE( '" & gstrUNITID & "','" & txtCustCode.Text.Trim & "','" & CmbInvType.Text.Trim & "','" & CmbInvSubType.Text.Trim & "','" & txtRefNo.Text.Trim & "')"
                If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSQL)) = True Then
                    intiLoopCount = 0
                    For intiLoopCount = 1 To SpChEntry.MaxRows
                        With SpChEntry
                            .Col = GridHeader.model
                            .Row = intiLoopCount
                            If .Text.Trim.Length = 0 Then
                                .Col = GridHeader.CustPartNo
                                .Row = intiLoopCount
                                strCustDrgNoLists = strCustDrgNoLists + .Text.Trim + ","
                            End If
                        End With
                    Next

                    If strCustDrgNoLists.Trim.Length > 0 Then
                        lstrControls = lstrControls & vbCrLf & lNo & ". Model Code can't be blank for below Part Number(s) :" & vbCrLf & strCustDrgNoLists
                        lNo = lNo + 1
                        ValidatebeforeSave = False
                    End If
                End If
                '10808160 CHANGES DONE
            Case "EDIT"
                '10736222
                strSQL = "DELETE FROM TMP_CT2_INVOICE_KNOCKOFF where UNIT_CODE='" + gstrUNITID + "' and IP_ADDRESS='" & gstrIpaddressWinSck & "'"
                SqlConnectionclass.ExecuteNonQuery(strSQL)
                '10736222

                If Val(lblECSStax_Per.Text) > 0 Then
                    If Len(Trim(txtSECSSTaxType.Text)) = 0 Then
                        lstrControls = lstrControls & vbCrLf & lNo & ". Secondary ECESS"
                        lNo = lNo + 1
                        If lctrFocus Is Nothing Then
                            lctrFocus = txtSECSSTaxType
                        End If
                        ValidatebeforeSave = False
                    End If
                End If
                If CmbInvType.Visible = True Then
                End If
                If (Len(Me.txtFreight.Text) = 0) Then
                    txtFreight.Text = "0.00"
                End If
                If (Len(Me.txtSurchargeTaxType.Text) = 0) Then
                End If
                '10808160 
                strSQL = "select dbo.UDF_ISEOPINVOICE( '" & gstrUNITID & "','" & txtCustCode.Text.Trim & "','" & CmbInvType.Text.Trim & "','" & CmbInvSubType.Text.Trim & "','" & txtRefNo.Text.Trim & "')"
                If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSQL)) = True Then
                    intiLoopCount = 0
                    For intiLoopCount = 1 To SpChEntry.MaxRows
                        With SpChEntry
                            .Col = GridHeader.model
                            .Row = intiLoopCount
                            If .Text.Trim.Length = 0 Then
                                .Col = GridHeader.CustPartNo
                                .Row = intiLoopCount
                                strCustDrgNoLists = strCustDrgNoLists + .Text.Trim + ","
                            End If
                        End With
                    Next

                    If strCustDrgNoLists.Trim.Length > 0 Then
                        lstrControls = lstrControls & vbCrLf & lNo & ". Model Code can't be blank for below Part Number(s) :" & vbCrLf & strCustDrgNoLists
                        lNo = lNo + 1
                        ValidatebeforeSave = False
                    End If
                End If

                If gblnGSTUnit = True And txtTCSTaxCode.Enabled = True Then
                    If txtTCSTaxCode.Text.Trim.ToString.Length = 0 Then
                        lstrControls = lstrControls & vbCrLf & lNo & ". TCS Tax."
                        lNo = lNo + 1
                        If lctrFocus Is Nothing Then
                            lctrFocus = Me.txtTCSTaxCode
                        End If
                        ValidatebeforeSave = False
                    End If
                End If

                '10808160 CHANGES DONE
        End Select
        If Not ValidatebeforeSave Then
            MsgBox(lstrControls, MsgBoxStyle.Information, ResolveResString(10059))
            If lctrFocus.Enabled Then lctrFocus.Focus()
        End If
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        gblnCancelUnload = True 'Initialize the Variable
        Exit Function
    End Function
    Private Sub ChangeCellTypeStaticText()
        '*****************************************************
        'Created By     -  Kapil
        'Description    -  To Change The Cell Type In Spread Control to Cell Type Static Text to
        '               -  Make Cell Type UnEditable
        '*****************************************************
        On Error GoTo ErrHandler
        Dim intRow As Short
        Dim intcol As Short
        Dim varItemCode As Object
        Dim blnQtyChkAccToMeasureCode As Boolean
        Dim rsSalesParameter As ClsResultSetDB
        Dim strMin As String
        Dim strMax As String
        Dim intDecimal As Short
        Dim intloopcounter1 As Short
        With Me.SpChEntry
            Select Case Me.CmdGrpChEnt.Mode
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                    If (UCase(Trim(CmbInvType.Text)) = "NORMAL INVOICE") Or (UCase(Trim(CmbInvType.Text)) = "EXPORT INVOICE") Or (UCase(Trim(CmbInvType.Text)) = "SERVICE INVOICE") Then
                        If UCase(Trim(CmbInvSubType.Text)) <> "SCRAP" Then
                            For intRow = 1 To .MaxRows
                                .Row = intRow
                                For intcol = 1 To .MaxCols
                                    .Col = intcol
                                    If intcol = GridHeader.Quantity Or intcol = GridHeader.BinQty Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                    ElseIf intcol = GridHeader.FromBox Or intcol = GridHeader.ToBox Or intcol = GridHeader.ToolCostPerUnit Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                    ElseIf intcol = GridHeader.CVD Or intcol = GridHeader.SAD Or intcol = GridHeader.delete Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
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
                                    If intcol = GridHeader.Quantity Or intcol = GridHeader.BinQty Or intcol = GridHeader.FromBox Or intcol = GridHeader.ToBox Or intcol = GridHeader.ToolCostPerUnit Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                    ElseIf intcol = GridHeader.RatePerUnit Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat ''' : .TypeFloatDecimalPlaces = 4: .TypeFloatMin = "0.0000": .TypeFloatMax = "99999999999999.99999"
                                    ElseIf intcol = GridHeader.delete Or intcol = GridHeader.CVD Or intcol = GridHeader.SAD Or intcol = GridHeader.EXC Then
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
                                If intcol = GridHeader.Quantity Or intcol = GridHeader.BinQty Or intcol = GridHeader.FromBox Or intcol = GridHeader.ToBox Or intcol = GridHeader.ToolCostPerUnit Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                ElseIf intcol = GridHeader.RatePerUnit Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat ''' : .TypeFloatDecimalPlaces = 4: .TypeFloatMin = "0.0000": .TypeFloatMax = "99999999999999.99999"
                                ElseIf intcol = GridHeader.delete Or intcol = GridHeader.CVD Or intcol = GridHeader.SAD Or intcol = GridHeader.EXC Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                                Else
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                                End If
                            Next intcol
                        Next intRow
                    End If
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                    If (UCase(strInvType) = "INV") Or (UCase(strInvType) = "EXP") Or (UCase(strInvType) = "SRC") Then
                        If (UCase(strInvSubType) <> "L") Then
                            For intRow = 1 To .MaxRows
                                .Row = intRow
                                For intcol = 1 To .MaxCols
                                    .Col = intcol
                                    If intcol = GridHeader.Quantity Or intcol = GridHeader.BinQty Or intcol = GridHeader.FromBox Or intcol = GridHeader.ToBox Or intcol = GridHeader.ToolCostPerUnit Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                    ElseIf intcol = GridHeader.delete Or intcol = GridHeader.CVD Or intcol = GridHeader.SAD Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
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
                                    If intcol = GridHeader.Quantity Or intcol = GridHeader.BinQty Or intcol = GridHeader.FromBox Or intcol = GridHeader.ToBox Or intcol = GridHeader.ToolCostPerUnit Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                    ElseIf intcol = GridHeader.RatePerUnit Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat ''' : .TypeFloatDecimalPlaces = 4: .TypeFloatMin = "0.0000": .TypeFloatMax = "99999999999999.99999"
                                    ElseIf intcol = GridHeader.delete Or intcol = GridHeader.CVD Or intcol = GridHeader.SAD Or intcol = GridHeader.EXC Then
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
                                If intcol = GridHeader.Quantity Or intcol = GridHeader.BinQty Or intcol = GridHeader.FromBox Or intcol = GridHeader.ToBox Or intcol = GridHeader.ToolCostPerUnit Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                ElseIf intcol = GridHeader.RatePerUnit Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat ''' : .TypeFloatDecimalPlaces = 4: .TypeFloatMin = "0.0000": .TypeFloatMax = "99999999999999.99999"
                                ElseIf intcol = GridHeader.delete Or intcol = GridHeader.CVD Or intcol = GridHeader.SAD Or intcol = GridHeader.EXC Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                                Else
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                                End If
                            Next intcol
                        Next intRow
                    End If
            End Select
            rsSalesParameter = New ClsResultSetDB
            rsSalesParameter.GetResult("Select QtyChkAccToMeasureCode from Sales_parameter WHERE UNIT_CODE='" & gstrUNITID & "'")
            If rsSalesParameter.GetNoRows > 0 Then
                If rsSalesParameter.GetValue("QtyChkAccToMeasureCode") = False Then
                    blnQtyChkAccToMeasureCode = False
                Else
                    blnQtyChkAccToMeasureCode = True
                End If
            End If
            rsSalesParameter = Nothing
            If blnQtyChkAccToMeasureCode = True Then
                For intRow = 1 To .MaxRows
                    varItemCode = Nothing
                    Call .GetText(GridHeader.InternalPartNo, intRow, varItemCode)
                    If Len(Trim(varItemCode)) > 0 Then
                        intDecimal = ReturnNoOfDecimals(CStr(varItemCode))
                        strMin = "0." : strMax = "99999999999999."
                        For intloopcounter1 = 1 To intDecimal
                            strMin = strMin & "0"
                            strMax = strMax & "9"
                        Next
                        If intDecimal = 0 Then
                            strMin = "0" : strMax = "99999999999999"
                        End If
                        .Row = intRow : .Row2 = intRow : .Col = GridHeader.Quantity : .Col2 = GridHeader.Quantity : .BlockMode = True '.CellType = CellTypeFloat
                        .TypeFloatDecimalPlaces = intDecimal ': .TypeFloatMin = strMin: .TypeFloatMax = strMax:
                        .BlockMode = False
                    End If
                Next
            End If
        End With
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Function QuantityCheck() As Boolean
        '*****************************************************
        'Created By     -  Nisha
        'Description    -  To Check Schedule Quantity From DailyMktSchedule/MonthlyMktSchedule
        '*****************************************************
        On Error GoTo ErrHandler
        QuantityCheck = False
        Dim strScheduleSql As String
        Dim strScheduleSql1 As String
        Dim rsMktSchedule As ClsResultSetDB
        Dim rsMktSchedule1 As ClsResultSetDB
        Dim rsChallanEntry As ClsResultSetDB
        Dim rsSaleConf As ClsResultSetDB
        Dim rsSalesParameter As New ClsResultSetDB
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
        Dim rsMktDailySchedule As ClsResultSetDB
        Dim intFromBox As Double
        Dim varCustRefNo As Object
        Dim varAmendmentNo As Object
        Dim varSRVDINo As Object
        Dim varKanbanNo As Object
        Dim rsKanBan As ClsResultSetDB
        Dim rsbom As New ClsResultSetDB
        Dim irowcount As Short
        Dim intRwCount1 As Short
        Dim strToolCode As String
        Dim varItemQty1 As Object
        Dim varBinQty As Object
        rsMktSchedule = New ClsResultSetDB
        rsMktSchedule1 = New ClsResultSetDB
        mstrUpdDispatchSql = ""
        For intRwCount = 1 To SpChEntry.MaxRows
            VarDelete = Nothing
            Call SpChEntry.GetText(GridHeader.delete, intRwCount, VarDelete)
            If UCase(VarDelete) <> "D" Then
                For intcol = 1 To SpChEntry.MaxCols
                    SpChEntry.Col = intcol
                    If (SpChEntry.Col = GridHeader.Quantity) Or (SpChEntry.Col = GridHeader.RatePerUnit) Or (SpChEntry.Col = GridHeader.ToBox) Or (SpChEntry.Col = GridHeader.FromBox) Then ''Column Changed By Tapan
                        SpChEntry.Row = intRwCount
                        If (Val(Trim(SpChEntry.Text)) = 0) Then
                            QuantityCheck = True
                            Call ConfirmWindow(10419, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                            SpChEntry.Row = intRwCount : SpChEntry.Col = intcol : SpChEntry.Action = 0 : SpChEntry.Focus()
                            Exit Function
                        End If
                        If (SpChEntry.Col = GridHeader.ToBox) Then
                            SpChEntry.Row = intRwCount : SpChEntry.Col = GridHeader.FromBox : intFromBox = Val(Trim(SpChEntry.Text))
                            SpChEntry.Row = intRwCount : SpChEntry.Col = GridHeader.ToBox
                            If Val(Trim(SpChEntry.Text)) < intFromBox Then
                                QuantityCheck = True
                                Call ConfirmWindow(10235, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                                SpChEntry.Row = intRwCount : SpChEntry.Col = GridHeader.ToBox : SpChEntry.Action = 0 : SpChEntry.Focus()
                                Exit Function
                            End If
                        End If
                    End If
                Next intcol
            End If
        Next intRwCount
        If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            strInvoiceType = UCase(Trim(CmbInvType.Text))
            strInvoiceSubType = UCase(Trim(CmbInvSubType.Text))
        ElseIf CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
            rsChallanEntry = New ClsResultSetDB
            rsChallanEntry.GetResult("Select a.Description,a.Sub_Type_Description from SaleConf a,SalesChallan_Dtl b where a.UNIT_CODE=b.UNIT_CODE AND a.UNIT_CODE='" & gstrUNITID & "' AND Doc_No = " & txtChallanNo.Text & " and a.Invoice_Type = b.Invoice_type and a.Sub_type = b.Sub_Category and a.Location_code = b.Location_code and (fin_start_date <= getdate() and fin_end_date >= getdate())")
            strInvoiceType = UCase(rsChallanEntry.GetValue("Description"))
            strInvoiceSubType = UCase(rsChallanEntry.GetValue("sub_type_Description"))
        End If
        Dim strSRVNo As String
        Dim strMakeDate As String
        If ((UCase(Trim(strInvoiceType)) = "NORMAL INVOICE") And (UCase(CStr((Trim(strInvoiceSubType)) = "FINISHED GOODS")) Or (UCase(Trim(strInvoiceSubType)) = "TRADING GOODS"))) Or (UCase(Trim(strInvoiceType)) = "JOBWORK INVOICE") Or (UCase(Trim(strInvoiceType)) = "EXPORT INVOICE") Or (UCase(Trim(strInvoiceType)) = "SERVICE INVOICE") Then
            For intRwCount = 1 To SpChEntry.MaxRows
                varItemCode = Nothing
                varDrgNo = Nothing
                varItemQty = Nothing
                VarDelete = Nothing
                varBinQty = Nothing
                Call SpChEntry.GetText(GridHeader.InternalPartNo, intRwCount, varItemCode)
                Call SpChEntry.GetText(GridHeader.CustPartNo, intRwCount, varDrgNo)
                Call SpChEntry.GetText(GridHeader.Quantity, intRwCount, varItemQty)
                Call SpChEntry.GetText(GridHeader.delete, intRwCount, VarDelete)
                Call SpChEntry.GetText(GridHeader.BinQty, intRwCount, varBinQty)
                If blnInvoiceAgainstMultipleSO Then
                    varCustRefNo = Nothing
                    varAmendmentNo = Nothing
                    varSRVDINo = Nothing
                    Call SpChEntry.GetText(GridHeader.CustRefNo, intRwCount, varCustRefNo)
                    Call SpChEntry.GetText(GridHeader.AmendmentNo, intRwCount, varAmendmentNo)
                    Call SpChEntry.GetText(GridHeader.srvdino, intRwCount, varSRVDINo)
                    rsKanBan = New ClsResultSetDB
                    rsKanBan.GetResult("Select sch_date from mkt_EnagareDtl where Kanbanno='" & varSRVDINo & "' AND UNIT_CODE='" & gstrUNITID & "'")
                    If rsKanBan.GetNoRows > 0 Then
                        mstrNagareDate = VB6.Format(rsKanBan.GetValue("sch_date"), gstrDateFormat)
                    Else
                        mstrNagareDate = ""
                    End If
                    rsKanBan = Nothing
                End If
                If UCase(VarDelete) <> "D" Then
                    If CheckMeasurmentUnit(varItemCode, varItemQty, intRwCount, True) = False Then
                        QuantityCheck = True
                        Exit Function
                    End If
                End If
                If UCase(VarDelete) <> "D" Then
                    If Val(varBinQty) = 0 Then
                        MsgBox("Bin Quantity can not be zero for Item-- " & varItemCode, MsgBoxStyle.Information, "eMpro")
                        QuantityCheck = True
                        Call SpChEntry.SetText(GridHeader.BinQty, intRwCount, varBinQty)
                        SpChEntry.Col = GridHeader.BinQty
                        SpChEntry.Row = SpChEntry.ActiveRow
                        SpChEntry.Focus()
                        Exit Function
                    End If
                    If CheckMeasurmentUnit(varItemCode, varBinQty, intRwCount, False) = False Then
                        QuantityCheck = True
                        Exit Function
                    End If
                End If
                If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                    If UCase(VarDelete) <> "D" Then
                        If blnInvoiceAgainstMultipleSO Then
                            If CheckcustorddtlQty("ADD", CStr(varItemCode), CStr(varDrgNo), CDbl(varItemQty), CStr(varCustRefNo), CStr(varAmendmentNo)) = True Then
                                QuantityCheck = False
                            Else
                                QuantityCheck = True
                                SpChEntry.Col = GridHeader.Quantity : SpChEntry.Row = intRwCount : SpChEntry.Focus()
                                Exit Function
                            End If
                        Else
                            If CheckcustorddtlQty("ADD", CStr(varItemCode), CStr(varDrgNo), CDbl(varItemQty)) = True Then
                                QuantityCheck = False
                            Else
                                QuantityCheck = True
                                SpChEntry.Col = GridHeader.Quantity : SpChEntry.Row = intRwCount : SpChEntry.Focus()
                                Exit Function
                            End If
                        End If
                    End If
                ElseIf CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                    If blnInvoiceAgainstMultipleSO Then
                        If CheckcustorddtlQty("EDIT", CStr(varItemCode), CStr(varDrgNo), CDbl(varItemQty), CStr(varCustRefNo), CStr(varAmendmentNo)) = True Then
                            QuantityCheck = False
                        Else
                            QuantityCheck = True
                            SpChEntry.Col = GridHeader.Quantity : SpChEntry.Row = intRwCount : SpChEntry.Focus()
                            Exit Function
                        End If
                    Else
                        If CheckcustorddtlQty("EDIT", CStr(varItemCode), CStr(varDrgNo), CDbl(varItemQty)) = True Then
                            QuantityCheck = False
                        Else
                            QuantityCheck = True
                            SpChEntry.Col = GridHeader.Quantity : SpChEntry.Row = intRwCount : SpChEntry.Focus()
                            Exit Function
                        End If
                    End If
                End If
                If UCase(Trim(strInvoiceType)) <> "SERVICE INVOICE" Then
                    If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                        If blnInvoiceAgainstMultipleSO Then
                            If Len(CStr(varSRVDINo)) > 0 Then
                                ldblNetDispatchQty = GetTotalDispatchQuantityFromDailySchedule(Trim(txtCustCode.Text), Trim(varDrgNo), Trim(varItemCode), Trim(mstrNagareDate), "ADD", 0, CStr(varSRVDINo))
                            Else
                                ldblNetDispatchQty = GetTotalDispatchQuantityFromDailySchedule(Trim(txtCustCode.Text), Trim(varDrgNo), Trim(varItemCode), Trim(lblDateDes.Text), "ADD", 0)
                            End If
                        Else
                            If Len(Trim(txtSRVDINO.Text)) > 0 Then
                                ldblNetDispatchQty = GetTotalDispatchQuantityFromDailySchedule(Trim(txtCustCode.Text), Trim(varDrgNo), Trim(varItemCode), Trim(mstrNagareDate), "ADD", 0, Trim(txtSRVDINO.Text))
                            Else
                                ldblNetDispatchQty = GetTotalDispatchQuantityFromDailySchedule(Trim(txtCustCode.Text), Trim(varDrgNo), Trim(varItemCode), Trim(lblDateDes.Text), "ADD", 0)
                            End If
                        End If
                    ElseIf CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                        If blnInvoiceAgainstMultipleSO Then
                            If Len(CStr(varSRVDINo)) > 0 Then
                                ldblNetDispatchQty = GetTotalDispatchQuantityFromDailySchedule(Trim(txtCustCode.Text), Trim(varDrgNo), Trim(varItemCode), Trim(mstrNagareDate), "EDIT", mdblPrevQty(intRwCount - 1), CStr(varSRVDINo))
                            Else
                                ldblNetDispatchQty = GetTotalDispatchQuantityFromDailySchedule(Trim(txtCustCode.Text), Trim(varDrgNo), Trim(varItemCode), Trim(lblDateDes.Text), "EDIT", mdblPrevQty(intRwCount - 1))
                            End If
                        Else
                            If Len(Trim(txtSRVDINO.Text)) > 0 Then
                                ldblNetDispatchQty = GetTotalDispatchQuantityFromDailySchedule(Trim(txtCustCode.Text), Trim(varDrgNo), Trim(varItemCode), Trim(mstrNagareDate), "EDIT", mdblPrevQty(intRwCount - 1), Trim(txtSRVDINO.Text))
                            Else
                                ldblNetDispatchQty = GetTotalDispatchQuantityFromDailySchedule(Trim(txtCustCode.Text), Trim(varDrgNo), Trim(varItemCode), Trim(lblDateDes.Text), "EDIT", mdblPrevQty(intRwCount - 1))
                            End If
                        End If
                    End If
                    If ldblNetDispatchQty <> -1 Then
                        If Len(Trim(varDrgNo)) > 0 Then
                            If Val(varItemQty) > Val(CStr(ldblNetDispatchQty)) Then
                                QuantityCheck = True
                                MsgBox("Quantity should not be Greater then Schedule Quantity " & CStr(ldblNetDispatchQty) & "  for Item " & CStr(varDrgNo), MsgBoxStyle.Information, "eMPro")
                                With Me.SpChEntry
                                    .Row = intRwCount : .Col = GridHeader.Quantity : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                                End With
                                Exit Function
                            Else
                                QuantityCheck = False
                                If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                                    If UCase(VarDelete) <> "D" Then
                                        rsMktDailySchedule = New ClsResultSetDB
                                        If blnInvoiceAgainstMultipleSO Then
                                            strSRVNo = Trim(CStr(varSRVDINo))
                                        Else
                                            strSRVNo = Trim(txtSRVDINO.Text)
                                        End If
                                        If Len(Trim(strSRVNo)) > 0 Then
                                            rsMktDailySchedule.GetResult("Select * from DailyMktSchedule where Account_Code='" & Trim(txtCustCode.Text) & "' and  datepart(yyyy,Trans_Date)='" & Year(ConvertToDate(Trim(mstrNagareDate))) & "' and datepart(mm,Trans_Date)='" & Month(ConvertToDate(Trim(mstrNagareDate))) & "' and datepart(dd,Trans_Date)='" & VB.Day(ConvertToDate(Trim(mstrNagareDate))) & "' and Cust_DrgNo ='" & Trim(varDrgNo) & "'and Item_code = '" & varItemCode & "' and Status =1 AND UNIT_CODE='" & gstrUNITID & "'")
                                        Else
                                            rsMktDailySchedule.GetResult("Select * from DailyMktSchedule where Account_Code='" & Trim(txtCustCode.Text) & "' and  datepart(yyyy,Trans_Date)='" & Year(ConvertToDate(Trim(lblDateDes.Text))) & "' and datepart(mm,Trans_Date)='" & Month(ConvertToDate(Trim(lblDateDes.Text))) & "' and datepart(dd,Trans_Date)='" & VB.Day(ConvertToDate(Trim(lblDateDes.Text))) & "' and Cust_DrgNo ='" & Trim(varDrgNo) & "'and Item_code = '" & varItemCode & "' and Status =1 AND UNIT_CODE='" & gstrUNITID & "'")
                                        End If
                                        If rsMktDailySchedule.GetNoRows > 0 Then
                                            If Len(Trim(strSRVNo)) > 0 Then
                                                mstrUpdDispatchSql = Trim(mstrUpdDispatchSql) & "Update DailyMktSchedule set Despatch_qty ="
                                                mstrUpdDispatchSql = mstrUpdDispatchSql & "isnull(Despatch_Qty,0) + (" & Val(varItemQty) & ")"
                                                mstrUpdDispatchSql = mstrUpdDispatchSql & " Where UNIT_CODE='" & gstrUNITID & "' AND Account_Code='" & Trim(txtCustCode.Text) & "' and "
                                                mstrUpdDispatchSql = mstrUpdDispatchSql & " datepart(yyyy,Trans_Date)='" & Year(ConvertToDate(Trim(mstrNagareDate))) & "'"
                                                mstrUpdDispatchSql = mstrUpdDispatchSql & " and datepart(mm,Trans_Date)='" & Month(ConvertToDate(Trim(mstrNagareDate))) & "'"
                                                mstrUpdDispatchSql = mstrUpdDispatchSql & " and datepart(dd,Trans_Date)='" & VB.Day(ConvertToDate(Trim(mstrNagareDate))) & "'"
                                                mstrUpdDispatchSql = mstrUpdDispatchSql & " and Cust_DrgNo ='" & Trim(varDrgNo) & "'and Item_code = '" & varItemCode & "' and Status =1 " & vbCrLf
                                            Else
                                                mstrUpdDispatchSql = Trim(mstrUpdDispatchSql) & "Update DailyMktSchedule set Despatch_qty ="
                                                mstrUpdDispatchSql = mstrUpdDispatchSql & "isnull(Despatch_Qty,0) + (" & Val(varItemQty) & ")"
                                                mstrUpdDispatchSql = mstrUpdDispatchSql & " Where UNIT_CODE='" & gstrUNITID & "' AND Account_Code='" & Trim(txtCustCode.Text) & "' and "
                                                mstrUpdDispatchSql = mstrUpdDispatchSql & " datepart(yyyy,Trans_Date)='" & Year(ConvertToDate(Trim(lblDateDes.Text))) & "'"
                                                mstrUpdDispatchSql = mstrUpdDispatchSql & " and datepart(mm,Trans_Date)='" & Month(ConvertToDate(Trim(lblDateDes.Text))) & "'"
                                                mstrUpdDispatchSql = mstrUpdDispatchSql & " and datepart(dd,Trans_Date)='" & VB.Day(ConvertToDate(Trim(lblDateDes.Text))) & "'"
                                                mstrUpdDispatchSql = mstrUpdDispatchSql & " and Cust_DrgNo ='" & Trim(varDrgNo) & "'and Item_code = '" & varItemCode & "' and Status =1 " & vbCrLf
                                            End If
                                        Else
                                            If Len(Trim(strSRVNo)) > 0 Then
                                                mstrUpdDispatchSql = Trim(mstrUpdDispatchSql) & " Insert into dailymktschedule "
                                                mstrUpdDispatchSql = mstrUpdDispatchSql & "(Account_Code,Trans_date,cust_drgno,"
                                                mstrUpdDispatchSql = mstrUpdDispatchSql & "Schedule_Flag,Item_Code,Schedule_Quantity,Despatch_qty,"
                                                mstrUpdDispatchSql = mstrUpdDispatchSql & "Status,Ent_UserId,Upd_UserId,Ent_dt,Upd_dt,"
                                                mstrUpdDispatchSql = mstrUpdDispatchSql & "RevisionNo,UNIT_CODE) values ('" & Trim(txtCustCode.Text) & "',"
                                                mstrUpdDispatchSql = mstrUpdDispatchSql & "'" & getDateForDB(mstrNagareDate) & "', '" & Trim(varDrgNo)
                                                mstrUpdDispatchSql = mstrUpdDispatchSql & "',1,'" & varItemCode & "',0," & Val(varItemQty) & ",1,'" & mP_User & "',"
                                                mstrUpdDispatchSql = mstrUpdDispatchSql & "'" & mP_User & "',getdate(),getdate(),0,'" & gstrUNITID & "')" & vbCrLf
                                            Else
                                                mstrUpdDispatchSql = Trim(mstrUpdDispatchSql) & " Insert into dailymktschedule "
                                                mstrUpdDispatchSql = mstrUpdDispatchSql & "(Account_Code,Trans_date,cust_drgno,"
                                                mstrUpdDispatchSql = mstrUpdDispatchSql & "Schedule_Flag,Item_Code,Schedule_Quantity,Despatch_qty,"
                                                mstrUpdDispatchSql = mstrUpdDispatchSql & "Status,Ent_UserId,Upd_UserId,Ent_dt,Upd_dt,"
                                                mstrUpdDispatchSql = mstrUpdDispatchSql & "RevisionNo,UNIT_CODE) values ('" & Trim(txtCustCode.Text) & "',"
                                                mstrUpdDispatchSql = mstrUpdDispatchSql & "'" & dtpDateDesc.Value & "', '" & Trim(varDrgNo)
                                                mstrUpdDispatchSql = mstrUpdDispatchSql & "',1,'" & varItemCode & "',0," & Val(varItemQty) & ",1,'" & mP_User & "',"
                                                mstrUpdDispatchSql = mstrUpdDispatchSql & "'" & mP_User & "',getdate(),getdate(),0,'" & gstrUNITID & "')" & vbCrLf
                                            End If
                                        End If
                                    End If
                                ElseIf CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                                    If UCase(VarDelete) <> "D" Then
                                        If Len(Trim(strSRVNo)) > 0 Then
                                            mstrUpdDispatchSql = Trim(mstrUpdDispatchSql) & "Update DailyMktSchedule set Despatch_qty ="
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & "isnull(Despatch_Qty,0) + (" & Val(varItemQty) & ") - (" & mdblPrevQty(intRwCount - 1) & ")"
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & " Where UNIT_CODE='" & gstrUNITID & "' AND Account_Code='" & Trim(txtCustCode.Text) & "' and "
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & " datepart(yyyy,Trans_Date)='" & Year(ConvertToDate(Trim(mstrNagareDate))) & "'"
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & " and datepart(mm,Trans_Date)='" & Month(ConvertToDate(Trim(mstrNagareDate))) & "'"
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & " and datepart(dd,Trans_Date)='" & VB.Day(ConvertToDate(Trim(mstrNagareDate))) & "'"
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & " and Cust_DrgNo ='" & Trim(varDrgNo) & "'and Item_code = '" & varItemCode & "' and Status =1 " & vbCrLf
                                        Else
                                            mstrUpdDispatchSql = Trim(mstrUpdDispatchSql) & "Update DailyMktSchedule set Despatch_qty ="
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & "isnull(Despatch_Qty,0) + (" & Val(varItemQty) & ") - (" & mdblPrevQty(intRwCount - 1) & ")"
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & " Where UNIT_CODE='" & gstrUNITID & "' AND Account_Code='" & Trim(txtCustCode.Text) & "' and "
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & " datepart(yyyy,Trans_Date)='" & Year(ConvertToDate(Trim(lblDateDes.Text))) & "'"
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & " and datepart(mm,Trans_Date)='" & Month(ConvertToDate(Trim(lblDateDes.Text))) & "'"
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & " and datepart(dd,Trans_Date)='" & VB.Day(ConvertToDate(Trim(lblDateDes.Text))) & "'"
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & " and Cust_DrgNo ='" & Trim(varDrgNo) & "'and Item_code = '" & varItemCode & "' and Status =1 " & vbCrLf
                                        End If
                                    Else
                                        If Len(Trim(strSRVNo)) > 0 Then
                                            mstrUpdDispatchSql = Trim(mstrUpdDispatchSql) & "Update DailyMktSchedule set Despatch_qty ="
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & "isnull(Despatch_Qty,0)  - (" & mdblPrevQty(intRwCount - 1) & ")"
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & " Where UNIT_CODE='" & gstrUNITID & "' AND Account_Code='" & Trim(txtCustCode.Text) & "' and "
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & " datepart(yyyy,Trans_Date)='" & Year(ConvertToDate(Trim(mstrNagareDate))) & "'"
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & " and datepart(mm,Trans_Date)='" & Month(ConvertToDate(Trim(mstrNagareDate))) & "'"
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & " and datepart(dd,Trans_Date)='" & VB.Day(ConvertToDate(Trim(mstrNagareDate))) & "'"
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & " and Cust_DrgNo ='" & Trim(varDrgNo) & "'and Item_code = '" & varItemCode & "' and Status =1 " & vbCrLf
                                        Else
                                            mstrUpdDispatchSql = Trim(mstrUpdDispatchSql) & "Update DailyMktSchedule set Despatch_qty ="
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & "isnull(Despatch_Qty,0)  - (" & mdblPrevQty(intRwCount - 1) & ")"
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & " Where UNIT_CODE='" & gstrUNITID & "' AND Account_Code='" & Trim(txtCustCode.Text) & "' and "
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & " datepart(yyyy,Trans_Date)='" & Year(ConvertToDate(Trim(lblDateDes.Text))) & "'"
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & " and datepart(mm,Trans_Date)='" & Month(ConvertToDate(Trim(lblDateDes.Text))) & "'"
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & " and datepart(dd,Trans_Date)='" & VB.Day(ConvertToDate(Trim(lblDateDes.Text))) & "'"
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & " and Cust_DrgNo ='" & Trim(varDrgNo) & "'and Item_code = '" & varItemCode & "' and Status =1 " & vbCrLf
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                            If Len(Trim(strSRVNo)) > 0 Then
                                ldblNetDispatchQty = GetTotalDispatchQuantityFromMonthlySchedule(Trim(txtCustCode.Text), Trim(varDrgNo), Trim(varItemCode), Trim(mstrNagareDate), "ADD", 0)
                            Else
                                ldblNetDispatchQty = GetTotalDispatchQuantityFromMonthlySchedule(Trim(txtCustCode.Text), Trim(varDrgNo), Trim(varItemCode), Trim(lblDateDes.Text), "ADD", 0)
                            End If
                        ElseIf CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                            If Len(Trim(strSRVNo)) > 0 Then
                                ldblNetDispatchQty = GetTotalDispatchQuantityFromMonthlySchedule(Trim(txtCustCode.Text), Trim(varDrgNo), Trim(varItemCode), Trim(mstrNagareDate), "EDIT", mdblPrevQty(intRwCount - 1))
                            Else
                                ldblNetDispatchQty = GetTotalDispatchQuantityFromMonthlySchedule(Trim(txtCustCode.Text), Trim(varDrgNo), Trim(varItemCode), Trim(lblDateDes.Text), "EDIT", mdblPrevQty(intRwCount - 1))
                            End If
                        End If
                        If ldblNetDispatchQty <> -1 Then
                            If Len(Trim(varDrgNo)) > 0 Then
                                If Val(varItemQty) > Val(CStr(ldblNetDispatchQty)) Then
                                    QuantityCheck = True
                                    MsgBox("Quantity should not be Greater then Schedule Quantity " & CStr(ldblNetDispatchQty) & "  for Item " & CStr(varDrgNo), MsgBoxStyle.Information, "eMPro")
                                    With Me.SpChEntry
                                        .Row = intRwCount : .Col = GridHeader.Quantity : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                                    End With
                                    Exit Function
                                Else
                                    QuantityCheck = False
                                    If Len(Trim(txtSRVDINO.Text)) > 0 Then
                                        If Val(CStr(Month(CDate(lblDateDes.Text)))) < 10 Then
                                            strMakeDate = Year(ConvertToDate(mstrNagareDate)) & "0" & Month(ConvertToDate(mstrNagareDate))
                                        Else
                                            strMakeDate = Year(ConvertToDate(mstrNagareDate)) & Month(ConvertToDate(mstrNagareDate))
                                        End If
                                    Else
                                        If Val(CStr(Month(ConvertToDate(lblDateDes.Text)))) < 10 Then
                                            strMakeDate = Year(ConvertToDate(lblDateDes.Text)) & "0" & Month(ConvertToDate(lblDateDes.Text))
                                        Else
                                            strMakeDate = Year(ConvertToDate(lblDateDes.Text)) & Month(ConvertToDate(lblDateDes.Text))
                                        End If
                                    End If
                                    If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                                        If UCase(VarDelete) <> "D" Then
                                            mstrUpdDispatchSql = Trim(mstrUpdDispatchSql) & "Update MonthlyMktSchedule set Despatch_qty ="
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & "isnull(Despatch_Qty,0) + (" & Val(varItemQty) & ")"
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & " Where UNIT_CODE='" & gstrUNITID & "' AND Account_Code='" & Trim(txtCustCode.Text) & "' and "
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & " Year_Month=" & Val(Trim(strMakeDate)) & ""
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & " and Cust_DrgNo ='" & Trim(varDrgNo) & "'and Item_code = '" & varItemCode & "' and status =1 " & vbCrLf
                                        End If
                                    ElseIf CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                                        If UCase(VarDelete) <> "D" Then
                                            mstrUpdDispatchSql = Trim(mstrUpdDispatchSql) & "Update MonthlyMktSchedule set Despatch_qty ="
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & "isnull(Despatch_Qty,0) + (" & Val(varItemQty) & ") - (" & mdblPrevQty(intRwCount - 1) & ") "
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & " Where UNIT_CODE='" & gstrUNITID & "' AND  Account_Code='" & Trim(txtCustCode.Text) & "' and "
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & " Year_Month=" & Val(Trim(strMakeDate)) & ""
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & " and Cust_DrgNo ='" & Trim(varDrgNo) & "'and Item_code = '" & varItemCode & "' and status =1 " & vbCrLf
                                        Else
                                            mstrUpdDispatchSql = Trim(mstrUpdDispatchSql) & "Update MonthlyMktSchedule set Despatch_qty ="
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & "isnull(Despatch_Qty,0)  - (" & mdblPrevQty(intRwCount - 1) & ") "
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & " Where UNIT_CODE='" & gstrUNITID & "' AND Account_Code='" & Trim(txtCustCode.Text) & "' and "
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & " Year_Month=" & Val(Trim(strMakeDate)) & ""
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & " and Cust_DrgNo ='" & Trim(varDrgNo) & "'and Item_code = '" & varItemCode & "' and status =1 " & vbCrLf
                                        End If
                                    End If
                                End If
                            End If
                        Else
                            If VarDelete <> "D" Then
                                MsgBox("No Schedule Defined For " & varItemCode & " Item.", MsgBoxStyle.Information, "eMPro")
                                QuantityCheck = True
                                Cmditems.Focus()
                                Exit Function
                            End If
                        End If
                    End If
                End If
            Next
        End If
        For intRwCount = 1 To SpChEntry.MaxRows
            varItemCode = Nothing
            varDrgNo = Nothing
            varItemQty = Nothing
            VarDelete = Nothing
            varBinQty = Nothing
            Call SpChEntry.GetText(GridHeader.InternalPartNo, intRwCount, varItemCode)
            Call SpChEntry.GetText(GridHeader.CustPartNo, intRwCount, varDrgNo)
            Call SpChEntry.GetText(GridHeader.Quantity, intRwCount, varItemQty)
            Call SpChEntry.GetText(GridHeader.delete, intRwCount, VarDelete)
            Call SpChEntry.GetText(GridHeader.BinQty, intRwCount, varBinQty)
            If UCase(VarDelete) <> "D" Then
                If CheckMeasurmentUnit(varItemCode, varItemQty, intRwCount, True) = False Then
                    QuantityCheck = True
                    Exit Function
                End If
            End If
            If UCase(VarDelete) <> "D" Then
                If Val(varBinQty) = 0 Then
                    MsgBox("Bin Quantity can not be zero for Item-- " & varItemCode, MsgBoxStyle.Information, "eMpro")
                    QuantityCheck = True
                    Call SpChEntry.SetText(GridHeader.BinQty, intRwCount, varBinQty)
                    SpChEntry.Col = GridHeader.BinQty
                    SpChEntry.Row = SpChEntry.ActiveRow
                    SpChEntry.Focus()
                    Exit Function
                End If
                If CheckMeasurmentUnit(varItemCode, varBinQty, intRwCount, False) = False Then
                    QuantityCheck = True
                    Exit Function
                End If
            End If
        Next intRwCount
        Dim strItCode As String 'To Make Item Code String
        For intRwCount = 1 To Me.SpChEntry.MaxRows
            VarDelete = Nothing
            Call Me.SpChEntry.GetText(GridHeader.delete, intRwCount, VarDelete)
            If UCase(VarDelete) <> "D" Then
                varItemCode = Nothing
                Call Me.SpChEntry.GetText(GridHeader.InternalPartNo, intRwCount, varItemCode)
                strItCode = strItCode & "'" & Trim(varItemCode) & "',"
            End If
        Next intRwCount
        strItCode = Mid(strItCode, 1, Len(strItCode) - 1)
        rsSaleConf = New ClsResultSetDB
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT, UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                rsSaleConf.GetResult(" Select Invoice_type,Sub_Category from SalesChallan_Dtl Where UNIT_CODE='" & gstrUNITID & "' AND  Doc_No=" & txtChallanNo.Text)
                mstrInvoiceType = rsSaleConf.GetValue("Invoice_Type")
                mstrInvSubType = rsSaleConf.GetValue("Sub_Category")
                rsSaleConf.GetResult("select Stock_Location From saleconf where UNIT_CODE='" & gstrUNITID & "' AND Description ='" & Trim(strInvoiceType) & "' and sub_type_description ='" & Trim(strInvoiceSubType) & "' AND Location_Code ='" & Trim(txtLocationCode.Text) & "' and (fin_start_date <= getdate() and fin_end_date >= getdate())")
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                rsSaleConf.GetResult("select Stock_Location From saleconf where UNIT_CODE='" & gstrUNITID & "' AND Description ='" & Trim(CmbInvType.Text) & "' and sub_type_Description ='" & Trim(CmbInvSubType.Text) & "' AND Location_Code ='" & Trim(txtLocationCode.Text) & "' and (fin_start_date <= getdate() and fin_end_date >= getdate())")
        End Select
        If Len(Trim(rsSaleConf.GetValue("Stock_Location"))) = 0 Then
            MsgBox("Please Define Stock Location in Sales Conf First", MsgBoxStyle.OkOnly, "eMPro")
            QuantityCheck = True
            Exit Function
        End If
        Dim varItemCodeinVeiw As Object
        For intRwCount = 1 To Me.SpChEntry.MaxRows
            varItemCodeinVeiw = Nothing
            varDrgNo = Nothing
            VarDelete = Nothing
            Call SpChEntry.GetText(GridHeader.InternalPartNo, intRwCount, varItemCodeinVeiw)
            Call SpChEntry.GetText(GridHeader.CustPartNo, intRwCount, varDrgNo)
            ''Suspected 1 or 11
            Call SpChEntry.GetText(GridHeader.delete, intRwCount, VarDelete)
            If UCase(VarDelete) <> "D" Then
                strItembal = "Select Cur_Bal From ItemBal_Mst where UNIT_CODE='" & gstrUNITID & "' AND Location_Code ='" & Trim(rsSaleConf.GetValue("Stock_Location")) & "' and item_Code ='" & varItemCodeinVeiw & "'"
                rsMktSchedule.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                strQuantity = CStr(Val(rsMktSchedule.GetValue("Cur_Bal")))
                varItemQty = Nothing
                Call Me.SpChEntry.GetText(GridHeader.Quantity, intRwCount, varItemQty)
                If Val(varItemQty) > Val(strQuantity) Then
                    QuantityCheck = True
                    If CDbl(strQuantity) = 0 Then
                        MsgBox("No Balance Available for Item (" & varDrgNo & ")", MsgBoxStyle.OkOnly, "eMPro")
                    Else
                        MsgBox("Quantity should not be Greater then Current Balance (" & strQuantity & ") at location  " & rsSaleConf.GetValue("Stock_Location") & "  for Item " & varDrgNo, MsgBoxStyle.OkOnly, "eMPro")
                    End If
                    With Me.SpChEntry
                        .Row = intRwCount : .Col = GridHeader.Quantity : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                    End With
                    Exit Function
                Else
                    QuantityCheck = False
                End If
            End If
        Next intRwCount
        rsSaleConf.ResultSetClose()
        rsSaleConf = Nothing
        rsSalesParameter.GetResult("Select CheckToolAmortisation from Sales_Parameter WHERE UNIT_CODE='" & gstrUNITID & "'")
        If rsSalesParameter.GetNoRows > 0 Then
            rsSalesParameter.MoveFirst()
            If Len(Trim(rsSalesParameter.GetValue("CheckToolAmortisation"))) = 0 Then
                MsgBox("First define Check Tool Amortisation in Sales Parameter", MsgBoxStyle.Information, "eMPro")
                QuantityCheck = True
                Exit Function
            End If
            If rsSalesParameter.GetValue("CheckToolAmortisation") = True Then
                For intRwCount = 1 To Me.SpChEntry.MaxRows
                    varItemCodeinVeiw = Nothing
                    varDrgNo = Nothing
                    varToolCost = Nothing
                    VarDelete = Nothing
                    Call SpChEntry.GetText(1, intRwCount, varItemCodeinVeiw)
                    Call SpChEntry.GetText(2, intRwCount, varDrgNo)
                    Call SpChEntry.GetText(15, intRwCount, varToolCost)
                    Call SpChEntry.GetText(14, intRwCount, VarDelete)
                    If UCase(VarDelete) <> "D" Then
                        With mP_Connection
                            '.Execute("Truncate Table tmpBOM", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            .Execute("DELETE FROM tmpBOM WHERE UNIT_CODE='" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            .Execute("BOMExplosion '" & Trim(varItemCodeinVeiw) & "','" & Trim(varItemCodeinVeiw) & "',1,0,0,0,'" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        End With
                        rsbom.GetResult("select * from tmpBOM WHERE UNIT_CODE='" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        If rsbom.GetNoRows > 0 Then
                            irowcount = rsbom.GetNoRows
                            rsbom.MoveFirst()
                            For intRwCount1 = 1 To irowcount
                                strItembal = "select BalanceQty = isnull(a.proj_qty,0) - isnull(a.ClosingValueSMIEL,0),a.Tool_c from Amor_dtl a, tool_mst b "
                                strItembal = strItembal & " where a.UNIT_CODE=b.UNIT_CODE and a.UNIT_CODE='" & gstrUNITID & "' AND account_code = '" & Trim(txtCustCode.Text) & "'"
                                strItembal = strItembal & " and Item_code = '" & rsbom.GetValue("item_code") & "' and a.Tool_c = b.Tool_c and a.Item_code = b.Product_No order by a.tool_c"
                                rsMktSchedule.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                                If rsMktSchedule.GetNoRows > 0 Then
                                    rsMktSchedule.MoveFirst()
                                    strQuantity = CStr(Val(rsMktSchedule.GetValue("BalanceQty")))
                                    strToolCode = rsMktSchedule.GetValue("Tool_c")
                                    varItemQty = Nothing
                                    Call Me.SpChEntry.GetText(5, intRwCount, varItemQty)
                                    varItemQty1 = (varItemQty * Val(rsbom.GetValue("grossweight")))
                                    strItembal = "select BalanceQty = sum(isnull(UsedProjQty,0)) from Amor_dtl "
                                    strItembal = strItembal & " where "
                                    strItembal = strItembal & " Item_code = '" & rsbom.GetValue("item_code") & "' and tool_c = '" & strToolCode & "' AND UNIT_CODE='" & gstrUNITID & "'"
                                    rsMktSchedule.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                                    rsMktSchedule.MoveFirst()
                                    strQuantity = CStr(CDbl(strQuantity) - Val(rsMktSchedule.GetValue("BalanceQty")))
                                    If Val(varItemQty1) > Val(strQuantity) Then
                                        QuantityCheck = True
                                        If CDbl(strQuantity) = 0 Then
                                            MsgBox("No Balance Available for Item (" & rsbom.GetValue("item_code") & ") and customer Part Code (" & varDrgNo & ") For Amortisation Calculations. ", MsgBoxStyle.OkOnly, "eMPro")
                                        Else
                                            MsgBox("Quantity should not be Greater then available Balance Quantity for Amortisarion of this Item (" & rsbom.GetValue("item_code") & ")" & strQuantity, MsgBoxStyle.OkOnly, "eMPro")
                                            With Me.SpChEntry
                                                .Row = intRwCount : .Col = 5 : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                                            End With
                                            Exit Function
                                        End If
                                        Exit Function
                                    Else
                                        QuantityCheck = False
                                    End If
                                Else
                                End If
                                rsbom.MoveNext()
                            Next
                            '-------------------------------------------------------
                        End If
                        strItembal = "select BalanceQty = isnull(a.proj_qty,0) - isnull(a.ClosingValueSMIEL,0),a.Tool_c from Amor_dtl a,Tool_Mst b"
                        strItembal = strItembal & " where a.UNIT_CODE=b.UNIT_CODE AND a.UNIT_CODE='" & gstrUNITID & "' AND account_code = '" & Trim(txtCustCode.Text) & "'"
                        strItembal = strItembal & " and Item_code = '" & varItemCodeinVeiw & "' and a.Tool_c = b.tool_c and a.item_code = b.Product_No order by a.tool_c"
                        rsMktSchedule.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        If rsMktSchedule.GetNoRows > 0 Then
                            rsMktSchedule.MoveFirst()
                            strQuantity = CStr(Val(rsMktSchedule.GetValue("BalanceQty")))
                            strToolCode = rsMktSchedule.GetValue("Tool_c")
                            varItemQty = Nothing
                            Call Me.SpChEntry.GetText(5, intRwCount, varItemQty)
                            strItembal = "select BalanceQty = sum(isnull(UsedProjQty,0)) from Amor_dtl "
                            strItembal = strItembal & " where "
                            strItembal = strItembal & " Item_code = '" & rsbom.GetValue("item_code") & "' and tool_c = '" & strToolCode & "' AND UNIT_CODE='" & gstrUNITID & "'"
                            rsMktSchedule.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                            rsMktSchedule.MoveFirst()
                            strQuantity = CStr(CDbl(strQuantity) - Val(rsMktSchedule.GetValue("BalanceQty")))
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
                            strItembal = "select BalanceQty = isnull(a.proj_qty,0) - isnull(a.ClosingValueSMIEL,0) from Amor_dtl a"
                            strItembal = strItembal & " where a.UNIT_CODE='" & gstrUNITID & "' AND account_code = '" & Trim(txtCustCode.Text) & "'"
                            strItembal = strItembal & " and Item_code = '" & varItemCodeinVeiw & "'"
                            rsMktSchedule.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                            If rsMktSchedule.GetNoRows > 0 Then
                                rsMktSchedule.MoveFirst()
                                strQuantity = CStr(Val(rsMktSchedule.GetValue("BalanceQty")))
                                varItemQty = Nothing
                                Call Me.SpChEntry.GetText(5, intRwCount, varItemQty)
                                strItembal = "select BalanceQty = sum(isnull(UsedProjQty,0)) from Amor_dtl "
                                strItembal = strItembal & " where "
                                strItembal = strItembal & " Item_code = '" & varItemCodeinVeiw & "' AND UNIT_CODE='" & gstrUNITID & "'"
                                rsMktSchedule.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                                rsMktSchedule.MoveFirst()
                                strQuantity = CStr(CDbl(strQuantity) - Val(rsMktSchedule.GetValue("BalanceQty")))
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
                            End If
                        End If
                    End If
                Next intRwCount
            End If
        End If
        rsMktSchedule.ResultSetClose()
        rsMktSchedule = Nothing
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
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Sub RefreshForm(ByRef pstrType As String)
        '*****************************************************
        'Created By     -  Kapil
        'Description    -  To Refresh All The Fields
        '*****************************************************
        On Error GoTo ErrHandler
        Select Case UCase(pstrType)
            Case "LOCATION"
                txtLocationCode.Text = "" : lblLocCodeDes.Text = "" : lblRGPDes.Text = ""
                txtChallanNo.Text = "" : txtCustCode.Text = "" : lblCustCodeDes.Text = "" : lblAddressDes.Text = ""
                txtCarrServices.Text = "" : txtVehNo.Text = ""
                txtFreight.Text = "" : txtSaleTaxType.Text = "" : lblSaltax_Per.Text = "0.00"
                txtSurchargeTaxType.Text = "" : lblSurcharge_Per.Text = "0.00"
                txtLoadingTaxType.Text = "" : lblLoadingcharge_per.Text = "0.00"
                ctlInsurance.Text = "" : lblCurrencyDes.Text = "" : txtRefNo.Text = ""
                CmbInvType.SelectedIndex = -1 : CmbInvSubType.SelectedIndex = -1
                '***
                Me.CmdGrpChEnt.Enabled(1) = False
                Me.CmdGrpChEnt.Enabled(2) = False
                chkExciseExumpted.CheckState = System.Windows.Forms.CheckState.Unchecked
                If cmdConsigneeDetails.Visible Then cmdConsigneeDetails.Enabled = True
                txtContactPerson.Text = "" : txtECC.Text = "" : txtLST.Text = "" : txtAddress1.Text = "" : txtAddress2.Text = ""
                txtAddress3.Text = "" : cmdConsigneeOK.Enabled = True : cmdConsigneeCancel.Enabled = True
                txtTCSTaxCode.Text = ""
            Case "CHALLAN"
                txtChallanNo.Text = "" : txtCustCode.Text = "" : lblCustCodeDes.Text = "" : lblAddressDes.Text = ""
                txtCarrServices.Text = "" : txtVehNo.Text = ""
                txtFreight.Text = "" : txtSaleTaxType.Text = "" : lblSaltax_Per.Text = "0.00"
                txtSurchargeTaxType.Text = "" : lblSurcharge_Per.Text = "0.00"
                ctlInsurance.Text = "" : lblRGPDes.Text = ""
                txtLoadingTaxType.Text = "" : lblLoadingcharge_per.Text = "0.00"
                OptDiscountValue.Checked = True
                txtDiscountAmt.Text = "0.00"
                CmbInvType.SelectedIndex = -1 : CmbInvSubType.SelectedIndex = -1 : lblCurrencyDes.Text = "" : txtRefNo.Text = ""
                Me.CmdGrpChEnt.Enabled(1) = False
                Me.CmdGrpChEnt.Enabled(2) = False
                chkExciseExumpted.CheckState = System.Windows.Forms.CheckState.Unchecked
                If cmdConsigneeDetails.Visible Then cmdConsigneeDetails.Enabled = True
                txtContactPerson.Text = "" : txtECC.Text = "" : txtLST.Text = "" : txtAddress1.Text = "" : txtAddress2.Text = "" : txtAddress3.Text = ""
                cmdConsigneeOK.Enabled = True : cmdConsigneeCancel.Enabled = True : txtTCSTaxCode.Text = ""
        End Select
        With Me.SpChEntry
            .MaxRows = 1
            .Row = 1 : .Row2 = 1 : .Col = GridHeader.InternalPartNo : .Col2 = .MaxCols : .BlockMode = True : .Text = "" : .BlockMode = False
        End With
        strupSalechallan = ""
        strupSaleDtl = ""
        lblCustPartDesc.Text = ""
        lblCreditTerm.Text = ""
        lblCreditTermDesc.Text = ""
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
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
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub SelectChallanNoFromSalesChallanDtl()
        '*****************************************************
        'Created By     -  Kapil
        'Description    -  To Select Max.  Challan No. From SalesChallan_Dtl
        '*****************************************************
        On Error GoTo ErrHandler
        Dim strChallanNo As String
        Dim rsChallanNo As New ClsResultSetDB
        strChallanNo = "SELECT (CURRENT_NO + 1)CURRENT_NO FROM DOCUMENTTYPE_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND DOC_TYPE = 9999 AND GETDATE() BETWEEN FIN_START_DATE AND FIN_END_DATE"
        rsChallanNo.GetResult(strChallanNo, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsChallanNo.GetNoRows > 0 Then
            strChallanNo = rsChallanNo.GetValue("CURRENT_NO").ToString
            While Len(strChallanNo) < 6
                strChallanNo = "0" + strChallanNo
            End While
            strChallanNo = "99" + strChallanNo
            txtChallanNo.Text = strChallanNo
            strChallanNo = "UPDATE DOCUMENTTYPE_MST SET CURRENT_NO = CURRENT_NO + 1 WHERE UNIT_CODE='" & gstrUNITID & "' AND  DOC_TYPE = 9999 AND GETDATE() BETWEEN FIN_START_DATE AND FIN_END_DATE"
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
    Public Sub displayDeatilsfromCustOrdHdrandDtl()
        '*****************************************************
        'Created By     -  Kapil
        'Description    -  To Select Max.  Challan No. From SalesChallan_Dtl
        '*****************************************************
        On Error GoTo ErrHandler
        Dim strCustOrdHdr As String
        Dim rsCustOrdHdr As ClsResultSetDB
        Dim strCurrency As String
        Dim intDecimalPlace As Short
        'To Get Data from Cusft_Ord_hdr
        '***************************************
        Select Case CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT, UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                strCustOrdHdr = "Select max(Order_date),SalesTax_Type,"
                strCustOrdHdr = strCustOrdHdr & "Currency_Code,PerValue,term_payment from Cust_ord_hdr"
                strCustOrdHdr = strCustOrdHdr & " Where UNIT_CODE='" & gstrUNITID & "' AND Account_Code='" & txtCustCode.Text & "' and Cust_Ref ='"
                strCustOrdHdr = strCustOrdHdr & mstrRefNo & "'and Amendment_No ='" & mstrAmmNo & "' Group By salestax_type,currency_code,term_payment"
                rsCustOrdHdr = New ClsResultSetDB
                rsCustOrdHdr.GetResult(strCustOrdHdr, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                strCurrency = rsCustOrdHdr.GetValue("Currency_code")
                intDecimalPlace = ToGetDecimalPlaces(strCurrency)
                If intDecimalPlace < 2 Then
                    intDecimalPlace = 2
                End If
                ctlInsurance.DecSize = intDecimalPlace : txtFreight.DecSize = intDecimalPlace
                txtSaleTaxType.Text = rsCustOrdHdr.GetValue("SalesTax_Type")
                ctlPerValue.Text = rsCustOrdHdr.GetValue("PerValue")
                Call txtSaleTaxType_Validating(txtSaleTaxType, New System.ComponentModel.CancelEventArgs(False))
                rsCustOrdHdr.ResultSetClose()
                rsCustOrdHdr = Nothing
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                If UCase(CStr(Trim(CmbInvType.Text))) = "NORMAL INVOICE" Or UCase(CStr(Trim(CmbInvType.Text))) = "JOBWORK INVOICE" Or UCase(CStr(Trim(CmbInvType.Text))) = "EXPORT INVOICE" Or UCase(CStr(Trim(CmbInvType.Text))) = "SERVICE INVOICE" Then
                    If CBool(UCase(CStr((Trim(CmbInvSubType.Text)) <> "SCRAP"))) Then
                        If Len(Trim(txtRefNo.Text)) > 0 Or blnInvoiceAgainstMultipleSO Then
                            strCustOrdHdr = "Select max(Order_date),SalesTax_Type,Currency_code,PerValue,term_payment, surcharge_code from Cust_ord_hdr"
                            strCustOrdHdr = strCustOrdHdr & " Where UNIT_CODE='" & gstrUNITID & "' AND Account_Code='" & txtCustCode.Text & "' and Cust_Ref ='"
                            strCustOrdHdr = strCustOrdHdr & mstrRefNo & "' and Amendment_No ='" & mstrAmmNo & "'"
                            strCustOrdHdr = strCustOrdHdr & " and active_flag = 'A' Group by salestax_type,currency_code,PerValue,term_payment,surcharge_code"
                            rsCustOrdHdr = New ClsResultSetDB
                            rsCustOrdHdr.GetResult(strCustOrdHdr, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                            txtSaleTaxType.Text = rsCustOrdHdr.GetValue("SalesTax_Type")
                            Call txtSaleTaxType_Validating(txtSaleTaxType, New System.ComponentModel.CancelEventArgs(False))
                            txtSurchargeTaxType.Text = IIf(IsDBNull(rsCustOrdHdr.GetValue("surcharge_code")), "", rsCustOrdHdr.GetValue("surcharge_code"))
                            If txtSurchargeTaxType.Text.Length > 0 Then
                                Call txtSurchargeTaxType_Validating(txtSurchargeTaxType, New System.ComponentModel.CancelEventArgs(False))
                            Else
                                lblSurcharge_Per.Text = "0.00"
                            End If
                            strCurrency = rsCustOrdHdr.GetValue("Currency_code")
                            ctlPerValue.Text = rsCustOrdHdr.GetValue("PerValue")
                            lblCreditTerm.Text = IIf(IsDBNull(rsCustOrdHdr.GetValue("term_payment")), "", rsCustOrdHdr.GetValue("term_payment"))
                            If Len(Trim(lblCreditTerm.Text)) > 0 Then
                                Call SelectDescriptionForField("crTrm_desc", "crtrm_termID", "Gen_CreditTrmMaster", lblCreditTermDesc, Trim(lblCreditTerm.Text))
                            Else
                                lblCreditTermDesc.Text = ""
                            End If
                            lblCurrencyDes.Text = strCurrency
                            intDecimalPlace = ToGetDecimalPlaces(strCurrency)
                            If intDecimalPlace < 2 Then
                                intDecimalPlace = 2
                            End If
                            ctlInsurance.DecSize = intDecimalPlace : txtFreight.DecSize = intDecimalPlace
                            rsCustOrdHdr.ResultSetClose()
                            rsCustOrdHdr = Nothing
                        End If
                    End If
                End If
        End Select
        Call DisplayDetailsInSpread(strCurrency)
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Function SetMaxLengthInSpread(ByRef pintDecimalSize As Short) As Object
        '*****************************************************
        'Created By     -  Kapil
        'Description    -  To Set Max Length Of Columns Of Spread
        '*****************************************************
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
            For intRow = 1 To .MaxRows
                .Row = intRow
                .Col = GridHeader.InternalPartNo : .TypeMaxEditLen = 30
                .Col = GridHeader.CustPartNo : .TypeMaxEditLen = 30
                .Col = GridHeader.RatePerUnit : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = pintDecimalSize : .TypeFloatMin = CDbl(strMin) : .TypeFloatMax = CDbl(strMax)
                .Col = GridHeader.CustSuppMatPerUnit : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = pintDecimalSize : .TypeFloatMin = CDbl(strMin) : .TypeFloatMax = CDbl(strMax)
                .Col2 = GridHeader.CustSuppMatPerUnit
                If mbln_CSM_Edit_Req = False Then
                    .BlockMode = True
                    .Lock = True
                    .BlockMode = False
                End If
                .Col = GridHeader.Quantity : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = pintDecimalSize : .TypeFloatMin = CDbl("0.00") : .TypeFloatMax = CDbl("99999999999999.99")
                .Col = GridHeader.BinQty : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = pintDecimalSize : .TypeFloatMin = CDbl("0.00") : .TypeFloatMax = CDbl("99999999999999.99")
                .Col = GridHeader.Packing : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = pintDecimalSize : .TypeFloatMin = CDbl(strMin) : .TypeFloatMax = CDbl(strMax)
                .Col = GridHeader.CVD : .TypeMaxEditLen = 6
                .Col = GridHeader.SAD : .TypeMaxEditLen = 6
                If CmbInvType.Text = "NORMAL INVOICE" Or CmbInvType.Text = "JOBWORK INVOICE" Or CmbInvType.Text = "EXPORT INVOICE" Or CmbInvType.Text = "SERVICE INVOICE" Then
                    If UCase(Trim(CmbInvSubType.Text)) <> "SCRAP" Then
                        .Col = GridHeader.EXC
                        .CtlEditMode = False
                    Else
                        .Col = GridHeader.EXC
                        .CtlEditMode = True
                    End If
                Else
                    .Col = GridHeader.EXC : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                    .CtlEditMode = True
                End If
                .Col = GridHeader.OthersPerUnit : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = 2 : .TypeFloatMin = CDbl("0.00") : .TypeFloatMax = CDbl("99999999999999.99")
                .Col = GridHeader.CumulativeBoxes : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                .Col = GridHeader.delete : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                .Col = GridHeader.ToolCostPerUnit : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = pintDecimalSize : .TypeFloatMin = CDbl(strMin) : .TypeFloatMax = CDbl(strMax)
                .Col = GridHeader.FromBox : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = pintDecimalSize : .TypeFloatMin = CDbl(strMin) : .TypeFloatMax = CDbl(strMax)
                .Col = GridHeader.ToBox : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = pintDecimalSize : .TypeFloatMin = CDbl(strMin) : .TypeFloatMax = CDbl(strMax)
                'addded by abhijit
                If gblnGSTUnit Then
                    '.BlockMode = True : .Col = GridHeader.Discount_Percent : .Col2 = GridHeader.Discount_Amt : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = pintDecimalSize
                    '.TypeFloatMin = CDbl(strMin) : .TypeFloatMax = CDbl(strMax) : .Lock = False : .BlockMode = False


                    .BlockMode = True : .Col = GridHeader.Basic_Value : .Col2 = GridHeader.Accessible_Amt : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = pintDecimalSize
                    .TypeFloatMin = CDbl(strMin) : .TypeFloatMax = CDbl(strMax) : .Lock = True : .BlockMode = False

                    .BlockMode = True : .Col = GridHeader.CGST_Percent : .Col2 = GridHeader.CGST_Amt : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = pintDecimalSize
                    .TypeFloatMin = CDbl(strMin) : .TypeFloatMax = CDbl(strMax) : .Lock = True : .BlockMode = False

                    .BlockMode = True : .Col = GridHeader.SGST_Percent : .Col2 = GridHeader.SGST_Amt : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = pintDecimalSize
                    .TypeFloatMin = CDbl(strMin) : .TypeFloatMax = CDbl(strMax) : .Lock = True : .BlockMode = False

                    .BlockMode = True : .Col = GridHeader.IGST_Percent : .Col2 = GridHeader.IGST_Amt : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = pintDecimalSize
                    .TypeFloatMin = CDbl(strMin) : .TypeFloatMax = CDbl(strMax) : .Lock = True : .BlockMode = False

                    .BlockMode = True : .Col = GridHeader.UTGST_Percent : .Col2 = GridHeader.UTGST_Amt : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = pintDecimalSize
                    .TypeFloatMin = CDbl(strMin) : .TypeFloatMax = CDbl(strMax) : .Lock = True : .BlockMode = False

                    .BlockMode = True : .Col = GridHeader.CSESS_TAX_Percent : .Col2 = GridHeader.CSESS_TAX_VAL : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = pintDecimalSize
                    .TypeFloatMin = CDbl(strMin) : .TypeFloatMax = CDbl(strMax) : .Lock = True : .BlockMode = False

                    If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Or CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                        .BlockMode = True : .Col = GridHeader.Discount_Percent : .Col2 = GridHeader.Discount_Amt : .Lock = False : .BlockMode = False
                    End If

                End If
                'addded by abhijit
            Next intRow
            

        End With
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Public Function DeleteRecords() As Boolean
        On Error GoTo ErrHandler
        DeleteRecords = False
        strupSalechallan = ""
        strupSaleDtl = ""
        strupSalechallan = "Delete SalesChallan_Dtl where Doc_No =" & Trim(txtChallanNo.Text)
        strupSalechallan = strupSalechallan & " and Location_Code ='" & Trim(txtLocationCode.Text) & "' AND UNIT_CODE='" & gstrUNITID & "'"
        strupSaleDtl = "Delete Sales_Dtl where Doc_No =" & Trim(txtChallanNo.Text)
        strupSaleDtl = strupSaleDtl & " and Location_Code ='" & Trim(txtLocationCode.Text) & "' AND UNIT_CODE='" & gstrUNITID & "'"
        DeleteRecords = True
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Public Function CheckMeasurmentUnit(ByRef strItem As Object, ByRef strQuantity As Object, ByRef intRow As Short, ByRef blnQtyStatus As Boolean) As Boolean
        Dim strMeasure As String
        Dim rsMeasure As ClsResultSetDB
        On Error GoTo ErrHandler
        strMeasure = "select a.Decimal_allowed_flag from Measure_Mst a,Item_Mst b"
        strMeasure = strMeasure & " where a.UNIT_CODE=b.UNIT_CODE AND a.UNIT_CODE='" & gstrUNITID & "' AND b.cons_Measure_Code=a.Measure_Code and b.Item_Code = '" & strItem & "'"
        rsMeasure = New ClsResultSetDB
        rsMeasure.GetResult(strMeasure, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        If rsMeasure.GetValue("Decimal_allowed_flag") = False Then
            If System.Math.Round(strQuantity, 0) - Val(strQuantity) <> 0 Then
                If blnQtyStatus = True Then
                    MsgBox("Quantity can not be in Decimal/Fraction for item-- " & strItem, MsgBoxStyle.Information, "eMpro")
                    CheckMeasurmentUnit = False
                    Call SpChEntry.SetText(GridHeader.Quantity, intRow, strQuantity)
                    SpChEntry.Col = GridHeader.Quantity
                    SpChEntry.Row = SpChEntry.ActiveRow
                    SpChEntry.Focus()
                    Exit Function
                Else
                    MsgBox("Bin quantity can not be in Decimal/Fraction for item-- " & strItem, MsgBoxStyle.Information, "eMpro")
                    CheckMeasurmentUnit = False
                    Call SpChEntry.SetText(GridHeader.BinQty, intRow, strQuantity)
                    SpChEntry.Col = GridHeader.BinQty
                    SpChEntry.Row = SpChEntry.ActiveRow
                    SpChEntry.Focus()
                    Exit Function
                End If
            Else
                CheckMeasurmentUnit = True
            End If
        Else
            CheckMeasurmentUnit = True
        End If
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Public Function ParentQty(ByRef pstrItemCode As String, ByRef pstrfinished As Object) As Double
        On Error GoTo ErrHandler
        Dim strParentQty As String
        Dim rsParentQty As ClsResultSetDB
        strParentQty = "select sum(required_qty + waste_Qty) as TotalQty from Bom_Mst where finished_Product_code ='"
        strParentQty = strParentQty & pstrfinished & "' and rawMaterial_Code ='" & pstrItemCode & "' AND UNIT_CODE='" & gstrUNITID & "'"
        rsParentQty = New ClsResultSetDB
        rsParentQty.GetResult(strParentQty, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        ParentQty = rsParentQty.GetValue("TotalQty")
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Function StockLocationSalesConf(ByRef pstrInvType As String, ByRef pstrInvSubtype As String, ByRef pstrFeild As String) As String
        Dim rsSalesConf As ClsResultSetDB
        Dim StockLocation As String
        On Error GoTo ErrHandler
        rsSalesConf = New ClsResultSetDB
        Select Case pstrFeild
            Case "DESCRIPTION"
                rsSalesConf.GetResult("Select Stock_Location from SaleConf Where UNIT_CODE='" & gstrUNITID & "' AND Description ='" & Trim(pstrInvType) & "' and Sub_type_Description ='" & Trim(pstrInvSubtype) & "' AND Location_Code='" & Trim(txtLocationCode.Text) & "' and (fin_start_date <= getdate() and fin_end_date >= getdate())")
            Case "TYPE"
                rsSalesConf.GetResult("Select Stock_Location from SaleConf Where UNIT_CODE='" & gstrUNITID & "' AND Invoice_type ='" & Trim(pstrInvType) & "' and Sub_type ='" & Trim(pstrInvSubtype) & "' AND Location_Code='" & Trim(txtLocationCode.Text) & "' and (fin_start_date <= getdate() and fin_end_date >= getdate())")
        End Select
        If rsSalesConf.GetNoRows > 0 Then
            StockLocation = rsSalesConf.GetValue("Stock_Location")
        End If
        StockLocationSalesConf = StockLocation
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Function GetServerDate() As Date
        Dim objServerDate As ClsResultSetDB 'Class Object
        Dim strSQL As String 'Stores the SQL statement
        On Error GoTo ErrHandler
        'Build the SQL statement
        strSQL = "SELECT CONVERT(datetime,getdate(),103)"
        'Creating the instance
        objServerDate = New ClsResultSetDB
        With objServerDate
            'Open the recordset
            Call .GetResult(strSQL, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            'If we have a record, then getting the financial year else exiting
            If .GetNoRows <= 0 Then Exit Function
            'GetServerDate = CDate(VB6.Format(DateValue(.GetValueByNo(0)), gstrDateFormat))
            GetServerDate = DateValue(.GetValueByNo(0))
            .ResultSetClose()
        End With
        objServerDate = Nothing
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Public Function ExploreBom(ByRef pstrItemCode As String, ByRef pstrFinishedQty As Object, ByRef pstrSPCurrentRow As Object, ByRef pstrFinishedProduct As String) As Boolean
        '*****************************************************
        'Created By     -  Nisha
        'Description    -  to get the values of required items in Sub assambly bom
        'input Variable -  Item Code to Found, reqquantity of Finished Product,row in spread
        '*****************************************************
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
        rsBomMstRaw = New ClsResultSetDB
        rsCustAnnexDtl = New ClsResultSetDB
        rsItemMst = New ClsResultSetDB
        strBomMstRaw = "Select RawMaterial_Code,Required_qty + Waste_qty "
        strBomMstRaw = strBomMstRaw & " As TotalReqQty,Process_Type from Bom_Mst where "
        strBomMstRaw = strBomMstRaw & " item_Code ='" & strBomItem
        strBomMstRaw = strBomMstRaw & "'and finished_product_code ='"
        strBomMstRaw = strBomMstRaw & pstrItemCode & "' AND UNIT_CODE='" & gstrUNITID & "'"
        rsBomMstRaw.GetResult(strBomMstRaw, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        Dim intArrCount As Short
        Dim blnArrItemFound As Boolean
        If rsBomMstRaw.GetNoRows > 0 Then ' If Item Found in Bom Mst
            intBomMaxRaw = rsBomMstRaw.GetNoRows
            rsBomMstRaw.MoveFirst()
            For intCurrentRaw = 1 To intBomMaxRaw
                strBomItem = rsBomMstRaw.GetValue("RawMaterial_code")
                dblTotalReqQty = rsBomMstRaw.GetValue("TotalReqQty")
                strCustAnnexDtl = "Select Item_Code,Balance_qty = sum(Balance_qty) from CustAnnex_hdr where UNIT_CODE='" & gstrUNITID & "' AND Customer_code ='"
                strCustAnnexDtl = strCustAnnexDtl & Trim(txtCustCode.Text) & "' "
                If blnFIFO = False Then
                    strCustAnnexDtl = strCustAnnexDtl & " and ref57f4_no in ("
                    '***
                    strRGPQuote = Replace(mstrRGP, "§", "','", 1)
                    strRGPQuote = "'" & strRGPQuote & "'"
                    '***
                    strCustAnnexDtl = strCustAnnexDtl & Trim(strRGPQuote) & ") "
                End If
                strCustAnnexDtl = strCustAnnexDtl & " and getdate() <= "
                strCustAnnexDtl = strCustAnnexDtl & " DateAdd(d, 180, ref57f4_date)"
                strCustAnnexDtl = strCustAnnexDtl & " and Item_code ='" & strBomItem & "' group by Item_code"
                rsCustAnnexDtl.GetResult(strCustAnnexDtl, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                If rsCustAnnexDtl.GetNoRows >= 1 Then 'if item Found in CustAnnex then replace that item from Parant string
                    rsVandorBom = New ClsResultSetDB
                    rsVandorBom.GetResult("Select RawMaterial_Code from Vendor_bom where UNIT_CODE='" & gstrUNITID & "' AND Finish_Product_code = '" & pstrFinishedProduct & "'and RawMaterial_code = '" & strBomItem & "' and Vendor_code = '" & txtCustCode.Text & "'")
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
                                    SpChEntry.Col = GridHeader.Quantity
                                    SpChEntry.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                    ExploreBom = False
                                    Exit Function
                                Else
                                    Exit For
                                    ExploreBom = True
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
                                SpChEntry.Col = GridHeader.Quantity
                                SpChEntry.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                ExploreBom = False
                                Exit Function
                            Else
                                ExploreBom = True
                            End If
                        End If
                    End If
                Else
                    rsVandorBom.GetResult("Select RawMaterial_Code from Vendor_bom where UNIT_CODE='" & gstrUNITID & "' AND Finish_Product_code = '" & pstrFinishedProduct & "'and RawMaterial_code = '" & strBomItem & "' and vendor_code = '" & txtCustCode.Text & "'")
                    If rsVandorBom.GetNoRows > 0 Then
                        MsgBox("Item " & strBomItem & " is not supplied.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, "eMPro")
                        SpChEntry.Row = pstrSPCurrentRow
                        SpChEntry.Col = GridHeader.Quantity
                        SpChEntry.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                        ExploreBom = False
                        Exit Function
                    Else 'if not of Process type I then again Explore
                        rsItemMst.GetResult("Select Item_Main_grp from Item_Mst Where Item_code = '" & strBomItem & "' AND UNIT_CODE='" & gstrUNITID & "'")
                        If (UCase(rsItemMst.GetValue("Item_Main_grp")) = "R") Or (UCase(rsItemMst.GetValue("Item_Main_grp")) = "C") Then
                            ExploreBom = True
                        Else
                            pstrFinishedQty = pstrFinishedQty * dblTotalReqQty
                            Call ExploreBom(strBomItem, pstrFinishedQty, pstrSPCurrentRow, pstrFinishedProduct)
                        End If
                    End If
                End If
                rsBomMstRaw.MoveNext()
            Next
        Else
            MsgBox("No BOM Defind for Item (" & pstrItemCode & ") defined in challan", MsgBoxStyle.Information, "eMPro")
            ExploreBom = False
            Exit Function
        End If
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        'if Child Item not found in CustAnnex_dtl
    End Function
    Public Function BomCheck() As Boolean
        '*****************************************************
        'Created By     -  Nisha
        'Description    -  to get the values of required items in Sub assambly bom
        'input Variable -  Item Code to Found, reqquantity of Finished Product,row in spread
        '*****************************************************
        Dim intSpreadRow As Short
        Dim intSpCurrentRow As Short
        Dim intCurrentItem As Short
        Dim VarFinishedItem As Object
        Dim VarFinishedQty As Object
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
        rsItemMst = New ClsResultSetDB
        BomCheck = False
        intSpreadRow = SpChEntry.MaxRows
        inti = 0
        Dim intArrCount As Short
        Dim blnItemFoundinArray As Boolean ' to be used to check if item already exist in Array arrItem where we are storing all item we found in Cust annex
        If SpChEntry.MaxRows >= 1 Then
            For intSpCurrentRow = 1 To intSpreadRow
                With SpChEntry
                    VarFinishedItem = Nothing
                    VarFinishedQty = Nothing
                    Call .GetText(GridHeader.InternalPartNo, intSpCurrentRow, VarFinishedItem)
                    Call .GetText(GridHeader.Quantity, intSpCurrentRow, VarFinishedQty)
                End With
                strBomMst = "Select RawMaterial_Code,Process_type,Required_qty + Waste_qty "
                strBomMst = strBomMst & " As TotalReqQty"
                strBomMst = strBomMst & " from Bom_Mst where UNIT_CODE='" & gstrUNITID & "' AND Finished_Product_code ='"
                strBomMst = strBomMst & VarFinishedItem & "' Order By Bom_Level"
                rsBomMst = New ClsResultSetDB
                rsVandorBom = New ClsResultSetDB
                rsBomMst.GetResult(strBomMst, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                intBomMaxItem = rsBomMst.GetNoRows
                rsBomMst.MoveFirst()
                If intBomMaxItem > 0 Then ' Item Found in Bom_Mst
                    rsVandorBom.GetResult("Select RawMaterial_Code from Vendor_bom where UNIT_CODE='" & gstrUNITID & "' AND Finish_Product_code = '" & VarFinishedItem & "' and Vendor_code = '" & txtCustCode.Text & "'")
                    If rsVandorBom.GetNoRows > 0 Then
                        'Loop for Parent Items of Items at First lavel
                        For intCurrentItem = 1 To intBomMaxItem
                            strBomItem = ""
                            strBomItem = rsBomMst.GetValue("RawMaterial_Code")
                            strCustAnnexDtl = "Select Item_Code,Balance_qty = sum(Balance_qty) from CustAnnex_hdr where UNIT_CODE='" & gstrUNITID & "' AND Customer_code ='"
                            strCustAnnexDtl = strCustAnnexDtl & Trim(txtCustCode.Text) & "'"
                            If blnFIFO = False Then
                                strCustAnnexDtl = strCustAnnexDtl & " and ref57f4_no in ("
                                strRgpsWithQuots = Replace(mstrRGP, "§", "','", 1)
                                strRgpsWithQuots = "'" & strRgpsWithQuots & "'"
                                strCustAnnexDtl = strCustAnnexDtl & Trim(strRgpsWithQuots) & ") "
                            End If
                            strCustAnnexDtl = strCustAnnexDtl & " and getdate() <= "
                            strCustAnnexDtl = strCustAnnexDtl & " DateAdd(d, 180, ref57f4_date)"
                            strCustAnnexDtl = strCustAnnexDtl & " and Item_code ='" & strBomItem & "' group By Item_code"
                            rsCustAnnexDtl = New ClsResultSetDB
                            rsCustAnnexDtl.GetResult(strCustAnnexDtl)
                            If rsCustAnnexDtl.GetNoRows >= 1 Then 'if item Found in Cust Annex
                                rsVandorBom.GetResult("Select RawMaterial_Code from Vendor_bom where UNIT_CODE='" & gstrUNITID & "' AND Finish_Product_code = '" & VarFinishedItem & "'and RawMaterial_code = '" & strBomItem & "' and Vendor_code = '" & txtCustCode.Text & "'")
                                If rsVandorBom.GetNoRows > 0 Then
                                    rsCustAnnexDtl.MoveFirst()
                                    ReDim Preserve arrItem(inti)
                                    ReDim Preserve arrQty(inti)
                                    ReDim Preserve arrReqQty(inti)
                                    dblTotalReqQty = ParentQty(strBomItem, VarFinishedItem)
                                    If inti > 0 Then
                                        blnItemFoundinArray = False
                                        For intArrCount = 0 To UBound(arrItem) - 1
                                            If UCase(Trim(arrItem(intArrCount))) = UCase(rsCustAnnexDtl.GetValue("Item_code")) Then
                                                ' if item already exist in arritem then will sum up its requied Quantity in arrreqQty() and mark blnFoundinarray as true will be used later
                                                blnItemFoundinArray = True
                                                arrReqQty(intArrCount) = arrReqQty(intArrCount) + (dblTotalReqQty * VarFinishedQty)
                                                If arrQty(intArrCount) < arrReqQty(intArrCount) Then 'in case if sum up is less then Quantity suplied in cust annex
                                                    MsgBox("Customer Supplied Materail for Item " & arrItem(inti) & "is" & arrQty(inti) & ".", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, "eMPro")
                                                    SpChEntry.Row = intSpCurrentRow
                                                    SpChEntry.Col = GridHeader.Quantity
                                                    SpChEntry.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                                    BomCheck = False
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
                                                SpChEntry.Col = GridHeader.Quantity
                                                SpChEntry.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                                BomCheck = False
                                                Exit Function
                                            End If
                                        End If
                                    Else
                                        arrItem(inti) = rsCustAnnexDtl.GetValue("Item_code")
                                        arrQty(inti) = rsCustAnnexDtl.GetValue("Balance_qty")
                                        arrReqQty(inti) = dblTotalReqQty * VarFinishedQty
                                        If arrQty(inti) < arrReqQty(inti) Then 'Again Same Check
                                            MsgBox("Customer Supplied Material for Item " & arrItem(inti) & "is" & arrQty(inti) & ".", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, "eMPro")
                                            SpChEntry.Row = intSpCurrentRow
                                            SpChEntry.Col = GridHeader.Quantity
                                            SpChEntry.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                            BomCheck = False
                                            Exit Function
                                        End If
                                    End If
                                End If
                            Else
                                rsVandorBom.GetResult("Select RawMaterial_Code from Vendor_bom where UNIT_CODE='" & gstrUNITID & "' AND Finish_Product_code = '" & VarFinishedItem & "'and RawMaterial_code = '" & strBomItem & "' and Vendor_code = '" & txtCustCode.Text & "'")
                                If rsVandorBom.GetNoRows > 0 Then
                                    MsgBox("Item " & strBomItem & " is not supplied by Customer.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, "eMPro")
                                    SpChEntry.Row = intSpCurrentRow
                                    SpChEntry.Col = GridHeader.Quantity
                                    SpChEntry.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                    BomCheck = False
                                    Exit Function
                                Else
                                    rsItemMst.GetResult("Select Item_Main_grp from Item_Mst Where UNIT_CODE='" & gstrUNITID & "' AND Item_code = '" & strBomItem & "'")
                                    If (UCase(rsItemMst.GetValue("Item_Main_grp")) = "R") Or (UCase(rsItemMst.GetValue("Item_Main_grp")) = "C") Then
                                        BomCheck = True
                                    Else
                                        VarFinishedQty = VarFinishedQty * rsBomMst.GetValue("TotalReqQty")
                                        If ExploreBom(strBomItem, VarFinishedQty, intSpCurrentRow, CStr(VarFinishedItem)) = False Then
                                            BomCheck = False
                                            Exit Function
                                        End If
                                    End If
                                End If
                            End If
                            rsBomMst.MoveNext()
                            inti = inti + 1
                        Next
                    Else
                        MsgBox("No Customer BOM Defind for Item (" & VarFinishedItem & ") defined in challan", MsgBoxStyle.Information, "eMPro")
                        BomCheck = False
                        Exit Function
                    End If
                Else
                    MsgBox("No BOM Defind for Item (" & VarFinishedItem & ") defined in challan", MsgBoxStyle.Information, "eMPro")
                    BomCheck = False
                    Exit Function
                End If
            Next
        End If
        BomCheck = True
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Public Function ToGetDecimalPlaces(ByRef pstrCurrency As String) As Short
        Dim rscurrency As ClsResultSetDB
        rscurrency = New ClsResultSetDB
        rscurrency.GetResult("Select Decimal_Place from Currency_Mst where UNIT_CODE='" & gstrUNITID & "' AND Currency_code ='" & pstrCurrency & "'")
        ToGetDecimalPlaces = Val(rscurrency.GetValue("Decimal_Place"))
    End Function
    Public Function ToGetCurrencyType() As String
        Dim rsCustOrdHdr As ClsResultSetDB
        Dim strcustHdr As String
        On Error GoTo ErrHandler
        rsCustOrdHdr = New ClsResultSetDB
        strcustHdr = "Select Currency_Code from Cust_ord_hdr"
        strcustHdr = strcustHdr & " Where UNIT_CODE='" & gstrUNITID & "' AND Account_Code='" & txtCustCode.Text & "' and Cust_Ref ='"
        strcustHdr = strcustHdr & mstrRefNo & "'and Amendment_No ='" & mstrAmmNo & "'"
        rsCustOrdHdr.GetResult(strcustHdr)
        If rsCustOrdHdr.GetNoRows > 0 Then
            rsCustOrdHdr.MoveFirst()
            ToGetCurrencyType = rsCustOrdHdr.GetValue("Currency_Code")
        End If
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Function SelectDataFromCustOrd_Dtl(ByRef pstrCustCode As String, ByRef pstrInvType As String) As Boolean
        '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
        'Added By   -   Nitin Sood
        'Function Copied From frmMKTTRN0020
        'If User enters Any Ref Number , Returns TRUE if That Ref No. is Validated
        'To CHECK Data From Cust_Ord_Dtl
        '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
        On Error GoTo ErrHandler
        Dim strSelectSql As String 'Declared To Make Select Query
        Dim rsCustOrdDtl As ClsResultSetDB
        SelectDataFromCustOrd_Dtl = False
        If UCase(pstrInvType) = "JOBWORK INVOICE" Then
            strSelectSql = "Select b.Item_Code,b.Cust_DrgNo,b.Amendment_No,b.Cust_Ref from Cust_Ord_hdr a,Cust_Ord_Dtl b"
            strSelectSql = strSelectSql & " where b.Account_Code='" & Trim(pstrCustCode) & "' and b.Active_flag ='A' and "
            strSelectSql = strSelectSql & " a.Account_Code = b.Account_Code and a.Cust_ref =b.Cust_ref and "
            strSelectSql = strSelectSql & " a.Amendment_No = b.amendment_No AND a.Authorized_Flag = 1 and a.PO_type in ('J') "
            strSelectSql = strSelectSql & " and a.Valid_date >'" & getDateForDB(GetServerDate()) & "' and effect_Date <='" & getDateForDB(GetServerDate()) & "'"
            strSelectSql = strSelectSql & " AND b.Cust_Ref = '" & Trim(txtRefNo.Text) & "' AND a.UNIT_CODE=b.UNIT_CODE AND a.UNIT_CODE='" & gstrUNITID & "'"
            strSelectSql = strSelectSql & " order by b.Cust_Ref,b.Amendment_No,b.Cust_DrgNo,b.Item_Code "
        ElseIf UCase(pstrInvType) = "EXPORT INVOICE" Then
            strSelectSql = "Select b.Item_Code,b.Cust_DrgNo,b.Amendment_No,b.Cust_Ref from Cust_Ord_hdr a,Cust_Ord_Dtl b"
            strSelectSql = strSelectSql & " where b.Account_Code='" & Trim(pstrCustCode) & "' and b.Active_flag ='A' and "
            strSelectSql = strSelectSql & " a.Account_Code = b.Account_Code and a.Cust_ref =b.Cust_ref and "
            strSelectSql = strSelectSql & " a.Amendment_No = b.amendment_No AND a.Authorized_Flag = 1 and a.PO_type in ('E') "
            strSelectSql = strSelectSql & " and a.Valid_date >'" & getDateForDB(GetServerDate()) & "' and effect_date <='" & getDateForDB(GetServerDate()) & "'"
            strSelectSql = strSelectSql & " AND b.Cust_Ref = '" & Trim(txtRefNo.Text) & "' AND a.UNIT_CODE=b.UNIT_CODE AND a.UNIT_CODE='" & gstrUNITID & "'"
            strSelectSql = strSelectSql & " order by b.Cust_Ref,b.Amendment_No,b.Cust_DrgNo,b.Item_Code "
        ElseIf UCase(pstrInvType) = "REJECTION" Then
            strSelectSql = "select a.Doc_No,a.Item_code,a.Rejected_Quantity from grn_Dtl a,grn_hdr b Where "
            strSelectSql = strSelectSql & "a.Doc_type = b.Doc_type And a.Doc_No = b.Doc_No and "
            strSelectSql = strSelectSql & "a.From_Location = b.From_Location and a.From_Location ='01R1'"
            strSelectSql = strSelectSql & "and a.Rejected_quantity > 0  and b.Vendor_code = '" & pstrCustCode & "' AND A.Doc_No = " & txtRefNo.Text & "  AND ISNULL(b.GRN_Cancelled,0) = 0 AND a.UNIT_CODE=b.UNIT_CODE AND a.UNIT_CODE='" & gstrUNITID & "' order by a.Doc_No"
        Else
            strSelectSql = "Select b.Item_Code,b.Cust_DrgNo,b.Amendment_No,b.Cust_Ref from Cust_Ord_hdr a,Cust_Ord_Dtl b"
            strSelectSql = strSelectSql & " where b.Account_Code='" & Trim(pstrCustCode) & "' and b.Active_flag ='A' and "
            strSelectSql = strSelectSql & " a.Account_Code = b.Account_Code and a.Cust_ref =b.Cust_ref and "
            strSelectSql = strSelectSql & " a.Amendment_No = b.amendment_No  AND a.Authorized_Flag = 1 and a.PO_type in ('O','S','M') "
            strSelectSql = strSelectSql & " and a.Valid_date >'" & getDateForDB(GetServerDate()) & "' and effect_Date <= '" & getDateForDB(GetServerDate()) & "'"
            strSelectSql = strSelectSql & " AND b.Cust_Ref = '" & Trim(txtRefNo.Text) & "' AND a.UNIT_CODE=b.UNIT_CODE AND a.UNIT_CODE='" & gstrUNITID & "'"
            strSelectSql = strSelectSql & " order by b.Cust_Ref,b.Amendment_No,b.Cust_DrgNo,b.Item_Code "
        End If
        rsCustOrdDtl = New ClsResultSetDB
        rsCustOrdDtl.GetResult(strSelectSql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        If rsCustOrdDtl.GetNoRows > 0 Then '          'if record found
            SelectDataFromCustOrd_Dtl = True 'Return TRUE
            rsCustOrdDtl.ResultSetClose()
            rsCustOrdDtl = Nothing
        End If
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Function OriginalRefNoOVER(ByVal strRefNumber As String) As Boolean
        '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
        'Created By     -   Nitin Sood
        'Creation Date  -   26 June 2002
        'Description    -   Checks for Original Ref Over or Not
        '                   From Cust_Ord_Hdr
        'Arguments      -   By Value strRefNumber As String
        '                   (Reference Code Number For which Amendments are to be checked.
        '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
        On Error GoTo ErrHandler
        If SelectDataFromTable("Active_Flag", "Cust_ORD_HDR", " Where UNIT_CODE='" & gstrUNITID & "' AND Account_Code ='" & Trim(txtCustCode.Text) & "' AND Cust_Ref = '" & txtRefNo.Text & "' AND Amendment_No = ''") = "O" Then
            OriginalRefNoOVER = True
        Else
            OriginalRefNoOVER = False
        End If
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Function SelectDataFromTable(ByRef mstrFieldName As String, ByRef mstrTableName As String, ByRef mstrCondition As String) As String
        '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
        'Created By     -   Nitin Sood
        'Description    -   Get Data from BackEnd
        '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
        Dim StrSQLQuery As String
        Dim GetDataFromTable As ClsResultSetDB
        On Error GoTo ErrHandler
        StrSQLQuery = "Select " & mstrFieldName & " From " & mstrTableName & mstrCondition
        GetDataFromTable = New ClsResultSetDB
        If GetDataFromTable.GetResult(StrSQLQuery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic) Then
            If GetDataFromTable.GetNoRows > 0 Then
                GetDataFromTable.MoveFirst()
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
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function DateIsAppropriate() As Boolean
        '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
        'Created By     -   Nitin Sood
        'Creation Date  -   27 June 2002
        'Description    -   Checks for Specified Date is within LIMITs
        '                   From SalesChallan_DTL
        'Arguments      -   -
        '*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
        On Error GoTo ErrHandler
        Dim MaxInvoiceDate As String 'Get Max Date of Last Invoice made
        Dim CurrentDate As Date
        MaxInvoiceDate = SelectDataFromTable("INVOICE_DATE", "SalesChallan_Dtl", " WHERE UNIT_CODE='" & gstrUNITID & "' AND BILL_FLAG = 1 ORDER BY INVOICE_DATE DESC")
        If Trim(MaxInvoiceDate) = "" Then MaxInvoiceDate = "01/JAN/1900"
        CurrentDate = GetServerDate()
        If Len(MaxInvoiceDate) = 0 Then
            MaxInvoiceDate = CStr(GetServerDate())
        End If
        MaxInvoiceDate = VB6.Format(MaxInvoiceDate, gstrDateFormat)
        If (dtpDateDesc.Value <= CurrentDate) And (dtpDateDesc.Value >= ConvertToDate(MaxInvoiceDate)) Then
            DateIsAppropriate = True
        Else
            DateIsAppropriate = False
        End If
        Exit Function
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Public Function AddDataTolstRGPs() As Boolean
        Dim rsCustAnnex As ClsResultSetDB
        Dim intLoopCounter As Short
        Dim intMaxCounter As Short
        On Error GoTo ErrHandler
        With lvwRGPs
            .GridLines = True : .Items.Clear() : .Columns.Clear()
            Call .Columns.Insert(1, "", "RGP No(s)", CInt(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(lvwRGPs.Width) / 2)))
            Call .Columns.Insert(2, "", "RGP Date", CInt(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(lvwRGPs.Width) / 2 - 700)))
            rsCustAnnex = New ClsResultSetDB
            rsCustAnnex.GetResult("select distinct ref57f4_No,ref57f4_date from custannex_HDR where UNIT_CODE='" & gstrUNITID & "' AND customer_Code='" & Trim(txtCustCode.Text) & "' and getdate() < dateadd(d,180,ref57f4_Date) order by ref57f4_Date")
            If rsCustAnnex.GetNoRows > 0 Then
                AddDataTolstRGPs = True
                intMaxCounter = rsCustAnnex.GetNoRows
                rsCustAnnex.MoveFirst()
                For intLoopCounter = 1 To intMaxCounter
                    Call .Items.Insert(intLoopCounter, rsCustAnnex.GetValue("ref57f4_No"))
                    If .Items.Item(intLoopCounter).SubItems.Count > 1 Then
                        .Items.Item(intLoopCounter).SubItems(1).Text = rsCustAnnex.GetValue("ref57f4_Date")
                    Else
                        .Items.Item(intLoopCounter).SubItems.Insert(1, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsCustAnnex.GetValue("ref57f4_Date")))
                    End If
                    rsCustAnnex.MoveNext()
                Next
            Else
                AddDataTolstRGPs = False
            End If
        End With
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
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
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
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
        rsGrnDtl = New ClsResultSetDB
        intMaxLoop = SpChEntry.MaxRows
        ItemQtyCaseRejGrin = False
        For intLoopCounter = 1 To intMaxLoop
            varItemCode = Nothing
            VarDelete = Nothing
            varItemQty = Nothing
            Call SpChEntry.GetText(GridHeader.InternalPartNo, intLoopCounter, varItemCode)
            Call SpChEntry.GetText(GridHeader.delete, intLoopCounter, VarDelete)
            Call SpChEntry.GetText(GridHeader.Quantity, intLoopCounter, varItemQty)
            If VarDelete <> "D" Then
                strSQL = "select a.Doc_No,a.Item_code, MaxAllowedQty = ((a.Rejected_Quantity + a.excess_po_quantity) - (isnull(a.Despatch_Quantity,0) + isnull(a.Inspected_Quantity,0) + isnull(a.RGP_Quantity,0)))from grn_Dtl a,grn_hdr b Where "
                strSQL = strSQL & "a.Doc_type = b.Doc_type And a.Doc_No = b.Doc_No and "
                strSQL = strSQL & "a.From_Location = b.From_Location and a.From_Location ='01R1'"
                strSQL = strSQL & "and a.Rejected_quantity > 0 and b.Vendor_code = '" & txtCustCode.Text
                strSQL = strSQL & "' and a.Doc_No = " & CDbl(txtRefNo.Text) & " and a.Item_code = '" & varItemCode & "' AND ISNULL(b.GRN_Cancelled,0) = 0 AND a.UNIT_CODE=b.UNIT_CODE AND a.UNIT_CODE='" & gstrUNITID & "'"
                rsGrnDtl.GetResult(strSQL)
                If rsGrnDtl.GetNoRows > 0 Then
                    dblRejQty = rsGrnDtl.GetValue("MaxAllowedQty")
                    If varItemQty > dblRejQty Then
                        MsgBox("Quantity Allowed For This Item is " & dblRejQty & ", cannot Enter More then This.")
                        SpChEntry.Row = intLoopCounter : SpChEntry.Col = GridHeader.Quantity : SpChEntry.Action = FPSpreadADO.ActionConstants.ActionActiveCell : SpChEntry.Focus()
                        ItemQtyCaseRejGrin = False
                        Exit Function
                    Else
                        ItemQtyCaseRejGrin = True
                    End If
                Else
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
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Public Function ScheduleCheckEditMode() As Boolean
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
        If ((UCase(mstrInvType) = "INV") And (UCase(mstrInvSubType) = "F") Or (UCase(mstrInvSubType) = "T")) Or (UCase(Trim(CmbInvType.Text)) = "JOBWORK INVOICE") Or (UCase(mstrInvType) = "EXP") Or (UCase(mstrInvType) = "SRC") Then
            strScheduleSql = "Select Quantity=Schedule_Quantity-isnull(Despatch_Qty,0),Cust_DrgNo,Item_Code from DailyMktSchedule where Account_Code='" & Trim(txtCustCode.Text) & "' and "
            strScheduleSql = strScheduleSql & " datepart(yyyy,Trans_Date)='" & Year(ConvertToDate(Trim(lblDateDes.Text))) & "'"
            strScheduleSql = strScheduleSql & " and datepart(mm,Trans_Date)='" & Month(ConvertToDate(Trim(lblDateDes.Text))) & "'"
            strScheduleSql = strScheduleSql & " and datepart(dd,Trans_Date)='" & VB.Day(ConvertToDate(Trim(lblDateDes.Text))) & "'"
            strScheduleSql = strScheduleSql & " and Cust_DrgNo in(" & Trim(mstrItemCode) & ") and Status =1 and Schedule_Flag =1 AND UNIT_CODE='" & gstrUNITID & "'"
            rsMktSchedule.GetResult(strScheduleSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If rsMktSchedule.GetNoRows > 0 Then 'If Record Found
                rsMktSchedule.MoveFirst()
                For intRwCount = 1 To Me.SpChEntry.MaxRows
                    varItemQty = Nothing
                    VarDelete = Nothing
                    Call Me.SpChEntry.GetText(GridHeader.Quantity, intRwCount, varItemQty)
                    Call Me.SpChEntry.GetText(GridHeader.delete, intRwCount, VarDelete) ''Column Changed By Tapan
                    strQuantity = rsMktSchedule.GetValue("Quantity")
                    If UCase(VarDelete) <> "D" Then
                        If (Val(varItemQty) - Val(mdblPrevQty(intLoopCount))) > Val(CStr(strQuantity)) Then
                            ScheduleCheckEditMode = False
                            MsgBox("Quantity should not be Greater then Schedule Quantity " & strQuantity)
                            With Me.SpChEntry
                                .Row = intRwCount : .Col = GridHeader.Quantity : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                            End With
                            Exit Function
                        Else
                            ScheduleCheckEditMode = True
                            ' Make Update Query For Dispatch
                            mstrUpdDispatchSql = ""
                            For intLoopCount = 1 To SpChEntry.MaxRows
                                varDrgNo = Nothing
                                varItemCode = Nothing
                                PresQty = Nothing
                                Call Me.SpChEntry.GetText(GridHeader.CustPartNo, intLoopCount, varDrgNo)
                                Call Me.SpChEntry.GetText(GridHeader.InternalPartNo, intLoopCount, varItemCode)
                                Call Me.SpChEntry.GetText(GridHeader.Quantity, intLoopCount, PresQty)
                                strScheduleSql = "select Despatch_qty  = isnull(Despatch_Qty,0) - (" & Val(mdblPrevQty(intLoopCount - 1)) - Val(PresQty) & "),SChedule_Quantity from DailyMktSchedule "
                                strScheduleSql = strScheduleSql & " Where UNIT_CODE='" & gstrUNITID & "' AND Account_Code='" & Trim(txtCustCode.Text) & "' and "
                                strScheduleSql = strScheduleSql & " datepart(yyyy,Trans_Date)='" & Year(ConvertToDate(Trim(lblDateDes.Text))) & "'"
                                strScheduleSql = strScheduleSql & " and datepart(mm,Trans_Date)='" & Month(ConvertToDate(Trim(lblDateDes.Text))) & "'"
                                strScheduleSql = strScheduleSql & " and datepart(dd,Trans_Date)='" & VB.Day(ConvertToDate(Trim(lblDateDes.Text))) & "'"
                                strScheduleSql = strScheduleSql & " and Cust_DrgNo ='" & Trim(varDrgNo) & "'and Item_code = '" & varItemCode & "' and Status =1 and Schedule_flag =1" & vbCrLf
                                rsMktSchedule1.GetResult(strScheduleSql)
                                mstrUpdDispatchSql = Trim(mstrUpdDispatchSql) & "Update DailyMktSchedule set Despatch_qty ="
                                mstrUpdDispatchSql = mstrUpdDispatchSql & "isnull(Despatch_Qty,0) - (" & Val(mdblPrevQty(intLoopCount - 1)) - Val(PresQty) & ")"
                                If Val(rsMktSchedule1.GetValue("Despatch_Qty")) = Val(rsMktSchedule1.GetValue("Schedule_Quantity")) Then
                                    mstrUpdDispatchSql = mstrUpdDispatchSql & ", Schedule_Flag = 0 "
                                End If
                                mstrUpdDispatchSql = mstrUpdDispatchSql & " Where UNIT_CODE='" & gstrUNITID & "' AND Account_Code='" & Trim(txtCustCode.Text) & "' and "
                                mstrUpdDispatchSql = mstrUpdDispatchSql & " datepart(yyyy,Trans_Date)='" & Year(ConvertToDate(Trim(lblDateDes.Text))) & "'"
                                mstrUpdDispatchSql = mstrUpdDispatchSql & " and datepart(mm,Trans_Date)='" & Month(ConvertToDate(Trim(lblDateDes.Text))) & "'"
                                mstrUpdDispatchSql = mstrUpdDispatchSql & " and datepart(dd,Trans_Date)='" & VB.Day(ConvertToDate(Trim(lblDateDes.Text))) & "'"
                                mstrUpdDispatchSql = mstrUpdDispatchSql & " and Cust_DrgNo ='" & Trim(varDrgNo) & "'and Item_code = '" & varItemCode & "' and Status =1 and Schedule_flag =1" & vbCrLf
                            Next
                        End If
                    End If
                    rsMktSchedule.MoveNext()
                Next intRwCount
                'If Record Not Found In DailyMktSchedule Then Check From
                'MonthlyMktSchedule
            ElseIf rsMktSchedule.GetNoRows = 0 Then
                If Val(CStr(Month(ConvertToDate(lblDateDes.Text)))) < 10 Then
                    strMakeDate = Year(ConvertToDate(lblDateDes.Text)) & "0" & Month(ConvertToDate(lblDateDes.Text))
                Else
                    strMakeDate = Year(ConvertToDate(lblDateDes.Text)) & Month(ConvertToDate(lblDateDes.Text))
                End If
                strScheduleSql = "Select Quantity=Schedule_Qty-isnull(Despatch_Qty,0) from MonthlyMktSchedule where Account_Code='" & Trim(txtCustCode.Text) & "' and "
                strScheduleSql = strScheduleSql & " Year_Month=" & Val(Trim(strMakeDate)) & ""
                strScheduleSql = strScheduleSql & " and Cust_DrgNo in(" & Trim(mstrItemCode) & ") and status =1 and Schedule_flag =1 AND UNIT_CODE='" & gstrUNITID & "'"
                rsMktSchedule.GetResult(strScheduleSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                If rsMktSchedule.GetNoRows > 0 Then
                    rsMktSchedule.MoveFirst()
                    For intRwCount = 1 To Me.SpChEntry.MaxRows
                        Select Case CmdGrpChEnt.Mode
                            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                                strQuantity = rsMktSchedule.GetValue("Quantity")
                                varItemQty = Nothing
                                Call Me.SpChEntry.GetText(GridHeader.Quantity, intRwCount, varItemQty)
                            Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                                varItemQty = Nothing
                                Call Me.SpChEntry.GetText(GridHeader.Quantity, intRwCount, varItemQty)
                                strQuantity = Val(rsMktSchedule.GetValue("Quantity")) + Val(varItemQty)
                        End Select
                        VarDelete = Nothing
                        Call Me.SpChEntry.GetText(GridHeader.delete, intRwCount, VarDelete)
                        If UCase(VarDelete) <> "D" Then
                            If Val(varItemQty) > Val(CStr(strQuantity)) Then
                                ScheduleCheckEditMode = False
                                MsgBox("Quantity should not be Greater then Schedule Quantity " & strQuantity, MsgBoxStyle.Information, "eMPro")
                                With Me.SpChEntry
                                    .Row = intRwCount : .Col = GridHeader.Quantity : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                                End With
                                Exit Function
                            Else
                                ScheduleCheckEditMode = False
                                mstrUpdDispatchSql = ""
                                For intLoopCount = 1 To SpChEntry.MaxRows
                                    varDrgNo = Nothing
                                    varItemCode = Nothing
                                    PresQty = Nothing
                                    Call Me.SpChEntry.GetText(GridHeader.CustPartNo, intLoopCount, varDrgNo)
                                    Call Me.SpChEntry.GetText(GridHeader.InternalPartNo, intLoopCount, varItemCode)
                                    Call Me.SpChEntry.GetText(GridHeader.Quantity, intLoopCount, PresQty)
                                    '**** To Check schedule Quantity
                                    strScheduleSql = "Select Despatch_qty = "
                                    strScheduleSql = strScheduleSql & "isnull(Despatch_Qty,0) - (" & Val(mdblPrevQty(intLoopCount - 1)) - Val(PresQty) & "),Schedule_Qty"
                                    strScheduleSql = strScheduleSql & " From MonthlyMktSchedule "
                                    strScheduleSql = strScheduleSql & " Where UNIT_CODE='" & gstrUNITID & "' AND Account_Code='" & Trim(txtCustCode.Text) & "' and "
                                    strScheduleSql = strScheduleSql & " Year_Month=" & Val(Trim(strMakeDate)) & ""
                                    strScheduleSql = strScheduleSql & " and Cust_DrgNo ='" & Trim(varDrgNo) & "'and Item_code = '" & varItemCode & "' and status =1 and Schedule_flag =1" & vbCrLf
                                    rsMktSchedule1.GetResult(strScheduleSql)
                                    mstrUpdDispatchSql = Trim(mstrUpdDispatchSql) & "Update MonthlyMktSchedule set Despatch_qty ="
                                    mstrUpdDispatchSql = mstrUpdDispatchSql & "isnull(Despatch_Qty,0) - (" & Val(mdblPrevQty(intLoopCount - 1)) - Val(PresQty) & ")"
                                    If rsMktSchedule1.GetValue("Despatch_Qty") = rsMktSchedule1.GetValue("Schedule_Qty") Then
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & ", Schedule_Flag = 0 "
                                    End If
                                    mstrUpdDispatchSql = mstrUpdDispatchSql & " Where UNIT_CODE='" & gstrUNITID & "' AND Account_Code='" & Trim(txtCustCode.Text) & "' and "
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
                    Cmditems.Focus()
                    Exit Function
                End If
            Else
                MsgBox("No Schedule Defined For Selected Items,Define Schedule First")
                ScheduleCheckEditMode = False
                Cmditems.Focus()
                Exit Function
            End If
        End If
    End Function
    Private Function GetTaxRate(ByRef pstrFieldText As String, ByRef pstrColumnName As String, ByRef pstrTableName As String, ByRef pstrFieldName_WhichValueRequire As String, Optional ByRef pstrCondition As String = "") As Double
        '****************************************************
        'Created By     -  Tapan
        'Description    -  To Check Validity Of Field Data Whether it Exists In The
        '                  Database Or Not and Return it's Tax Rate
        'Arguments      -  pstrFieldText - Field Text,pstrColumnName - Column Name
        '               -  pstrTableName - Table Name,pstrCondition - Optional Parameter For Condition
        '****************************************************
        On Error GoTo ErrHandler
        GetTaxRate = 0
        Dim strTableSql As String 'Declared To Make Select Query
        Dim rsExistData As ClsResultSetDB
        If Len(Trim(pstrCondition)) > 0 Then
            strTableSql = "select " & Trim(pstrFieldName_WhichValueRequire) & " from " & Trim(pstrTableName) & " where " & Trim(pstrColumnName) & "='" & Trim(pstrFieldText) & "' and " & pstrCondition
        Else
            strTableSql = "select " & Trim(pstrFieldName_WhichValueRequire) & " from " & Trim(pstrTableName) & " where " & Trim(pstrColumnName) & "='" & Trim(pstrFieldText) & "'"
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
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Function GetExchangeRate(ByVal pstrCurrencyCode As String, ByVal pstrDate As String, ByVal IsCustomer As Boolean) As Double
        On Error GoTo ErrHandler
        GetExchangeRate = 1.0#
        Dim strTableSql As String 'Declared To Make Select Query
        Dim rsExistData As ClsResultSetDB
        pstrDate = VB6.Format(pstrDate, gstrDateFormat)
        If IsCustomer = True Then
            strTableSql = "SELECT CExch_MultiFactor FROM Gen_CurExchMaster WHERE UNIT_CODE='" & gstrUNITID & "' AND CExch_CurrencyTo='" & Trim(pstrCurrencyCode) & "' AND CExch_InOut=1 AND '" & getDateForDB(pstrDate) & "' BETWEEN CExch_DateFrom AND CExch_DateTo"
        Else
            strTableSql = "SELECT CExch_MultiFactor FROM Gen_CurExchMaster WHERE UNIT_CODE='" & gstrUNITID & "' AND CExch_CurrencyTo='" & Trim(pstrCurrencyCode) & "' AND CExch_InOut=0 AND '" & getDateForDB(pstrDate) & "' BETWEEN CExch_DateFrom AND CExch_DateTo"
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
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Function SaveData(ByVal Button As String) As Boolean
        '---------------------------------------------------------------------------------------
        'Name       :   SaveData
        'Type       :   Function
        'Author     :   Tapan Jain
        'Arguments  :
        'Return     :
        'Purpose    :
        '''***** Changes done by Ashutosh on 08-12-2005,Issue id:16518
        '---------------------------------------------------------------------------------------
        Dim ldblTotalBasicValue As Double
        Dim ldblTotalAccessibleValue As Double
        Dim lintLoopCounter As Short
        Dim ldblTempAccessibleVal As Double
        Dim ldblTotalExciseValue As Double
        Dim ldblTotalCVDValue As Double
        Dim ldblTotalSADValue As Double
        Dim ldbltempTotalExciseValue As Double
        Dim ldblTotalSaleTaxAmount As Double
        Dim ldblTotalSurchargeTaxAmount As Double
        Dim ldblNetInsurenceValue As Double
        Dim ldblTotalInvoiceValue As Double
        Dim ldblTotalOthersValues As Double
        Dim dblTotalLoadingcharges As Double
        Dim dblNetLoadingcharges As Double
        Dim dblTCSTaxAmount As Double
        Dim rsParameterData As ClsResultSetDB
        Dim strParamQuery As String
        Dim dblBasicForLoading As Double
        ''-----------Variable For Saving Data---------
        Dim strSalesChallan As String
        Dim updateSalesChallan As String
        Dim strSalesDtl As String
        Dim strSalesDtlDelete As String
        Dim rsCustItemMst As ClsResultSetDB
        Dim rsSaleConf As ClsResultSetDB
        Dim rsItemMst As ClsResultSetDB
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
        Dim PdblDiscountAmount As Double ' to store the value of Discount
        Dim blnISInsExcisable As Boolean
        Dim blnEOUFlag As Boolean
        Dim blnISExciseRoundOff As Boolean
        Dim blnISSalesTaxRoundOff As Boolean
        Dim blnISSurChargeTaxRoundOff As Boolean
        Dim blnAddCustMatrl As Boolean
        Dim blnISBasicRoundOff As Boolean
        Dim ldblExciseValueForSaleTax As Double
        Dim blnTotalToolCostRoundOff As Boolean
        Dim ldblTotalToolCost As Double
        Dim blnInsIncSTax As Boolean
        Dim blnTCSTax As Boolean
        Dim VarDelete As Object
        Dim intNonDeletedRowCount As Short
        Dim intBasicRoundOffDecimal As Short
        Dim intSaleTaxRoundOffDecimal As Short
        Dim intExciseRoundOffDecimal As Short
        Dim intSSTRoundOffDecimal As Short
        Dim intTCSRoundOffDecimal As Short
        Dim intToolCostRoundOffDecimal As Short
        Dim blnActiveTrans As Boolean
        Dim blnECSSTax As Boolean
        Dim intECSSRoundOffDecimal As Short
        Dim ldblTotalECSSTaxAmount As Double
        Dim ldblTotalSECSSTaxAmount As Double
        Dim strCustRef As String
        Dim StrAmendmentNo As String
        Dim strSrvDINo As String
        Dim strSRVLocation As String
        Dim strUSLoc As String
        Dim strSchTime As String
        Dim blnTotalInvoiceAmount As Boolean
        Dim intTotalInvoiceAmountRoundOffDecimal As Short
        Dim ldblTotalInvoiceValueRoundOff As Double
        Dim dblBinQuantity As Double
        Dim strSql As String
        Dim strSqlct2qry As String
        Dim dblExcise_Amount As Double
        Dim blnIsCt2 As Boolean = False
        Dim strModel As String = ""

        On Error GoTo ErrHandler
        ldblTotalBasicValue = 0
        ldblTotalAccessibleValue = 0
        ldblTotalExciseValue = 0
        ldblTotalCVDValue = 0
        ldblTotalSADValue = 0
        ldblTotalSaleTaxAmount = 0
        ldblTotalSurchargeTaxAmount = 0
        ldblTotalInvoiceValue = 0
        ldblTotalOthersValues = 0
        ldblTotalCustMatrlValue = 0
        ldblExciseValueForSaleTax = 0
        PdblDiscountAmount = 0
        ldblTotalECSSTaxAmount = 0
        dblBinQuantity = 0
        SaveData = True
        strParamQuery = "SELECT InsExc_Excise,CustSupp_Inc,EOU_Flag, Basic_Roundoff, Basic_Roundoff_decimal, SalesTax_Roundoff, SalesTax_Roundoff_decimal, Excise_Roundoff, Excise_Roundoff_decimal, "
        strParamQuery = strParamQuery & " SST_Roundoff, SST_Roundoff_decimal, InsInc_SalesTax, TCSTax_Roundoff, TCSTax_Roundoff_decimal, TotalToolCostRoundoff, TotalToolCostRoundoff_Decimal, ECESS_Roundoff, ECESSRoundoff_Decimal,TotalInvoiceAmount_Roundoff,TotalInvoiceAmountRoundOff_Decimal FROM Sales_Parameter WHERE UNIT_CODE='" & gstrUNITID & "'"
        rsParameterData = New ClsResultSetDB
        rsParameterData.GetResult(strParamQuery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsParameterData.GetNoRows > 0 Then
            blnISInsExcisable = rsParameterData.GetValue("InsExc_Excise")
            blnEOUFlag = rsParameterData.GetValue("EOU_Flag")
            blnISBasicRoundOff = rsParameterData.GetValue("Basic_Roundoff")
            blnISExciseRoundOff = rsParameterData.GetValue("Excise_Roundoff")
            blnISSalesTaxRoundOff = rsParameterData.GetValue("SalesTax_Roundoff")
            blnISSurChargeTaxRoundOff = rsParameterData.GetValue("SST_Roundoff")
            blnAddCustMatrl = rsParameterData.GetValue("CustSupp_Inc")
            blnInsIncSTax = rsParameterData.GetValue("InsInc_SalesTax")
            blnTCSTax = rsParameterData.GetValue("TCSTax_Roundoff")
            blnTotalToolCostRoundOff = rsParameterData.GetValue("TotalToolCostRoundoff")
            intBasicRoundOffDecimal = rsParameterData.GetValue("Basic_Roundoff_decimal")
            intSaleTaxRoundOffDecimal = rsParameterData.GetValue("SalesTax_Roundoff_decimal")
            intExciseRoundOffDecimal = rsParameterData.GetValue("Excise_Roundoff_decimal")
            intSSTRoundOffDecimal = rsParameterData.GetValue("SST_Roundoff_decimal")
            intTCSRoundOffDecimal = rsParameterData.GetValue("TCSTax_Roundoff_decimal")
            intToolCostRoundOffDecimal = rsParameterData.GetValue("TotalToolCostRoundoff_decimal")
            blnECSSTax = rsParameterData.GetValue("ECESS_Roundoff")
            intECSSRoundOffDecimal = rsParameterData.GetValue("ECESSRoundoff_Decimal")
            blnTotalInvoiceAmount = rsParameterData.GetValue("TotalInvoiceAmount_RoundOff")
            intTotalInvoiceAmountRoundOffDecimal = rsParameterData.GetValue("TotalInvoiceAmountRoundOff_Decimal")
        Else
            MsgBox("No data define in Sales_Parameter Table", MsgBoxStyle.Critical, "eMPro")
            SaveData = False
            rsParameterData.ResultSetClose()
            rsParameterData = Nothing
            Exit Function
        End If
        rsParameterData.ResultSetClose()
        rsParameterData = Nothing
        dblBasicForLoading = 0
        intNonDeletedRowCount = 0
        For lintLoopCounter = 1 To SpChEntry.MaxRows
            VarDelete = Nothing
            Call SpChEntry.GetText(GridHeader.delete, lintLoopCounter, VarDelete)
            If UCase(VarDelete) <> "D" Then
                dblBasicForLoading = dblBasicForLoading + CalculateBasicValue(lintLoopCounter, blnISBasicRoundOff)
                intNonDeletedRowCount = intNonDeletedRowCount + 1
            End If
        Next
        dblTotalLoadingcharges = CalculateLoadingchargesAmount(dblBasicForLoading, CDbl(lblLoadingcharge_per.Text))
        ldblNetInsurenceValue = System.Math.Round(Val(ctlInsurance.Text)) / intNonDeletedRowCount
        dblNetLoadingcharges = dblTotalLoadingcharges / intNonDeletedRowCount
        For lintLoopCounter = 1 To SpChEntry.MaxRows
            VarDelete = Nothing
            Call SpChEntry.GetText(GridHeader.delete, lintLoopCounter, VarDelete)
            If UCase(VarDelete) <> "D" Then
                ldblTotalBasicValue = ldblTotalBasicValue + CalculateBasicValue(lintLoopCounter, blnISBasicRoundOff)
                ldblTempAccessibleVal = CalculateAccessibleValue(lintLoopCounter, ldblNetInsurenceValue, blnISInsExcisable)
                If blnISExciseRoundOff Then
                    ldblTotalExciseValue = System.Math.Round(CalculateExciseValue(lintLoopCounter, ldblTempAccessibleVal, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff, dblNetLoadingcharges))
                    ldblTotalCVDValue = System.Math.Round(CalculateExciseValue(lintLoopCounter, ldblTempAccessibleVal, enumExciseType.RETURN_CVD, blnEOUFlag, blnISExciseRoundOff, dblNetLoadingcharges))
                    ldblTotalSADValue = System.Math.Round(CalculateExciseValue(lintLoopCounter, ldblTempAccessibleVal, enumExciseType.RETURN_SAD, blnEOUFlag, blnISExciseRoundOff, dblNetLoadingcharges))
                Else
                    ldblTotalExciseValue = System.Math.Round(CalculateExciseValue(lintLoopCounter, ldblTempAccessibleVal, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff, dblNetLoadingcharges), intExciseRoundOffDecimal)
                    ldblTotalCVDValue = System.Math.Round(CalculateExciseValue(lintLoopCounter, ldblTempAccessibleVal, enumExciseType.RETURN_CVD, blnEOUFlag, blnISExciseRoundOff, dblNetLoadingcharges), intExciseRoundOffDecimal)
                    ldblTotalSADValue = System.Math.Round(CalculateExciseValue(lintLoopCounter, ldblTempAccessibleVal, enumExciseType.RETURN_SAD, blnEOUFlag, blnISExciseRoundOff, dblNetLoadingcharges), intExciseRoundOffDecimal)
                End If
                ldblTotalAccessibleValue = ldblTotalAccessibleValue + ldblTempAccessibleVal
                SpChEntry.Row = lintLoopCounter : SpChEntry.Col = GridHeader.Quantity
                lintItemQuantity = Val(SpChEntry.Text)
                SpChEntry.Row = lintLoopCounter : SpChEntry.Col = GridHeader.OthersPerUnit
                ldblTotalOthersValues = ldblTotalOthersValues + ((Val(SpChEntry.Text) / Val(ctlPerValue.Text)) * lintItemQuantity)
                SpChEntry.Row = lintLoopCounter : SpChEntry.Col = GridHeader.CustSuppMatPerUnit
                ldblTotalCustMatrlValue = ldblTotalCustMatrlValue + ((Val(SpChEntry.Text) / Val(ctlPerValue.Text)) * lintItemQuantity)
                If blnEOU_FLAG Then
                    If blnISExciseRoundOff Then
                        ldblExciseValueForSaleTax = ldblExciseValueForSaleTax + System.Math.Round((ldblTotalExciseValue + ldblTotalCVDValue + ldblTotalSADValue) / 2)
                    Else
                        ldblExciseValueForSaleTax = ldblExciseValueForSaleTax + System.Math.Round((ldblTotalExciseValue + ldblTotalCVDValue + ldblTotalSADValue) / 2, intExciseRoundOffDecimal)
                    End If
                Else
                    If blnISExciseRoundOff Then
                        ldblExciseValueForSaleTax = ldblExciseValueForSaleTax + System.Math.Round(ldblTotalExciseValue)
                    Else
                        ldblExciseValueForSaleTax = ldblExciseValueForSaleTax + System.Math.Round(ldblTotalExciseValue, intExciseRoundOffDecimal)
                    End If
                End If
            End If
        Next
        If blnECSSTax Then
            ldblTotalSECSSTaxAmount = System.Math.Round(CalculateSECSSTaxValue(ldblExciseValueForSaleTax))
            ldblTotalECSSTaxAmount = System.Math.Round(CalculateECSSTaxValue(ldblExciseValueForSaleTax))
        Else
            ldblTotalSECSSTaxAmount = System.Math.Round(CalculateSECSSTaxValue(ldblExciseValueForSaleTax), intECSSRoundOffDecimal)
            ldblTotalECSSTaxAmount = System.Math.Round(CalculateECSSTaxValue(ldblExciseValueForSaleTax), intECSSRoundOffDecimal)
        End If
        If blnISSalesTaxRoundOff Then
            ldblTotalSaleTaxAmount = System.Math.Round(CalculateSalesTaxValue(ldblTotalBasicValue, ldblExciseValueForSaleTax + ldblTotalECSSTaxAmount + ldblTotalSECSSTaxAmount, blnInsIncSTax, CDbl(ctlInsurance.Text)))
        Else
            ldblTotalSaleTaxAmount = System.Math.Round(CalculateSalesTaxValue(ldblTotalBasicValue, ldblExciseValueForSaleTax + ldblTotalECSSTaxAmount + ldblTotalSECSSTaxAmount, blnInsIncSTax, CDbl(ctlInsurance.Text)), intSaleTaxRoundOffDecimal)
        End If
        If blnISSurChargeTaxRoundOff Then
            ldblTotalSurchargeTaxAmount = System.Math.Round(CalculateSurchargeTaxValue(ldblTotalSaleTaxAmount))
        Else
            ldblTotalSurchargeTaxAmount = System.Math.Round(CalculateSurchargeTaxValue(ldblTotalSaleTaxAmount), intSSTRoundOffDecimal)
        End If

        If Val(txtDiscountAmt.Text) > 0 Then
            If OptDiscountValue.Checked = True Then
                PdblDiscountAmount = System.Math.Round(Val(txtDiscountAmt.Text), 0)
            Else
                If chkExciseExumpted.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                    ' to calculate Discount Amount by Percentage if Excise Duty is available
                    PdblDiscountAmount = System.Math.Round(((ldblTotalBasicValue + ldblTotalExciseValue) * Val(txtDiscountAmt.Text)) / 100)
                Else
                    ' to calculate Discount Amount by Percentage if Excise Duty is not available
                    PdblDiscountAmount = System.Math.Round(((ldblTotalBasicValue) * Val(txtDiscountAmt.Text)) / 100)
                End If
            End If
        Else
            ' if discount amount is not available
            PdblDiscountAmount = 0
        End If
        '****************************************************************
        ''-----------------------------------------------------------------------------------------------------------------------------
        If blnAddCustMatrl Then
            If chkExciseExumpted.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                ldblTotalInvoiceValue = (ldblTotalBasicValue + ldblExciseValueForSaleTax + ldblTotalSaleTaxAmount + ldblTotalSurchargeTaxAmount + ldblTotalECSSTaxAmount + ldblTotalSECSSTaxAmount + System.Math.Round(Val(txtFreight.Text)) + System.Math.Round(ldblTotalOthersValues) + System.Math.Round(Val(ctlInsurance.Text)) + System.Math.Round(ldblTotalCustMatrlValue, 2)) - PdblDiscountAmount
            Else
                ldblTotalInvoiceValue = (ldblTotalBasicValue + ldblTotalSaleTaxAmount + ldblTotalSurchargeTaxAmount + ldblTotalECSSTaxAmount + ldblTotalSECSSTaxAmount + System.Math.Round(Val(txtFreight.Text)) + System.Math.Round(ldblTotalOthersValues) + System.Math.Round(Val(ctlInsurance.Text)) + System.Math.Round(ldblTotalCustMatrlValue, 2)) - PdblDiscountAmount
            End If
        Else
            If chkExciseExumpted.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                ldblTotalInvoiceValue = (ldblTotalBasicValue + ldblExciseValueForSaleTax + ldblTotalSaleTaxAmount + ldblTotalSurchargeTaxAmount + ldblTotalECSSTaxAmount + ldblTotalSECSSTaxAmount + System.Math.Round(Val(txtFreight.Text)) + System.Math.Round(ldblTotalOthersValues) + System.Math.Round(Val(ctlInsurance.Text))) - PdblDiscountAmount
            Else
                ldblTotalInvoiceValue = (ldblTotalBasicValue + ldblTotalSaleTaxAmount + ldblTotalSurchargeTaxAmount + ldblTotalECSSTaxAmount + ldblTotalSECSSTaxAmount + System.Math.Round(Val(txtFreight.Text)) + System.Math.Round(ldblTotalOthersValues) + System.Math.Round(Val(ctlInsurance.Text))) - PdblDiscountAmount
            End If
        End If
        If Val(lblTCSTaxPerDes.Text) > 0 Then
            dblTCSTaxAmount = CalculateTCSTax(ldblTotalInvoiceValue, blnTCSTax, Val(lblTCSTaxPerDes.Text))
            ldblTotalInvoiceValue = ldblTotalInvoiceValue + dblTCSTaxAmount
        End If
        If blnTotalInvoiceAmount Then
            ldblTotalInvoiceValueRoundOff = ldblTotalInvoiceValue - System.Math.Round(ldblTotalInvoiceValue)
            ldblTotalInvoiceValue = System.Math.Round(ldblTotalInvoiceValue)
        Else
            ldblTotalInvoiceValueRoundOff = ldblTotalInvoiceValue - System.Math.Round(ldblTotalInvoiceValue, intTotalInvoiceAmountRoundOffDecimal)
            ldblTotalInvoiceValue = System.Math.Round(ldblTotalInvoiceValue, intTotalInvoiceAmountRoundOffDecimal)
        End If
        Dim strStock_Loc As String
        Dim rsLocation As ClsResultSetDB
        rsLocation = New ClsResultSetDB
        strStock_Loc = ""
        rsLocation.GetResult("Select Invoice_Type,Sub_Type from SaleConf where UNIT_CODE='" & gstrUNITID & "' AND Description ='" & Trim(CmbInvType.Text) & "'and Sub_type_Description ='" & Trim(CmbInvSubType.Text) & "' and (fin_start_date <= getdate() and fin_end_date >= getdate())")
        If rsLocation.GetNoRows > 0 Then
            strStock_Loc = StockLocationSalesConf(rsLocation.GetValue("Invoice_Type"), rsLocation.GetValue("Sub_Type"), "TYPE")
        Else
            MsgBox("Stock Location is not defined", vbInformation + vbOKOnly, ResolveResString(100))
            Exit Function
        End If
        Select Case Button
            Case "ADD"
                rsSaleConf = New ClsResultSetDB
                rsSaleConf.GetResult("Select Invoice_Type,Sub_Type from SaleConf where UNIT_CODE='" & gstrUNITID & "' AND Description ='" & Trim(CmbInvType.Text) & "'and Sub_type_Description ='" & Trim(CmbInvSubType.Text) & "' and (fin_start_date <= getdate() and fin_end_date >= getdate())")
                mstrInvType = rsSaleConf.GetValue("Invoice_Type")
                mstrInvoiceSubType = rsSaleConf.GetValue("Sub_Type")
                strSalesChallan = ""
                If UCase(CmbInvType.Text) <> "JOBWORK INVOICE" Then
                    mstrRGP = ""
                End If
                If UCase(CmbInvType.Text) = "NORMAL INVOICE" And UCase(CmbInvSubType.Text) = "FINISHED GOODS" Then
                    If strStock_Loc = "01M1" Then strStock_Loc = "01B1"
                End If
                strSalesChallan = "INSERT INTO SalesChallan_dtl (UNIT_CODE,Location_Code,Doc_No,Suffix,Transport_Type,Vehicle_No,"
                strSalesChallan = strSalesChallan & "From_Station,To_Station,Account_Code,Cust_Ref,"
                strSalesChallan = strSalesChallan & "Amendment_No,Bill_Flag,Discount_type,Discount_Amount,Discount_Per,Form3,Carriage_Name,"
                strSalesChallan = strSalesChallan & "Year,Insurance,invoice_Type,Ref_Doc_No,"
                strSalesChallan = strSalesChallan & "Cust_Name ,Sales_Tax_Amount , Surcharge_Sales_Tax_Amount,"
                strSalesChallan = strSalesChallan & "Frieght_Amount,Sub_Category,SalesTax_Type,SalesTax_FormNo,SalesTax_FormValue,"
                strSalesChallan = strSalesChallan & "Annex_no,Invoice_Date,Currency_code,Ent_dt,"
                strSalesChallan = strSalesChallan & "Ent_UserId,Upd_dt,Upd_UserId,Exchange_Rate,total_amount,Surcharge_salesTaxType,"
                strSalesChallan = strSalesChallan & "SalesTax_Per,Surcharge_SalesTax_Per,Remarks,PerValue,SRVDINO,SRVLocation,"
                strSalesChallan = strSalesChallan & "LoadingChargeTaxType,LoadingChargeTaxAmount,LoadingChargeTax_Per,ExciseExumpted,"
                strSalesChallan = strSalesChallan & "ConsigneeContactPerson,ConsigneeECCNo,ConsigneeLST,ConsigneeAddress1,"
                strSalesChallan = strSalesChallan & "ConsigneeAddress2,ConsigneeAddress3"
                If UCase(CmbInvType.Text) = "JOBWORK INVOICE" Then
                    strSalesChallan = strSalesChallan & ",Fifo_Flag"
                End If
                strSalesChallan = strSalesChallan & ",USLOC,Schtime,TCSTax_Type,TCSTax_Per,TCSTaxAmount,ECESS_Type, ECESS_Per, ECESS_Amount,SECESS_Type, SECESS_Per, SECESS_Amount"
                strSalesChallan = strSalesChallan & " ,TotalInvoiceAmtRoundOff_diff "
                strSalesChallan = strSalesChallan & ",Payment_Terms"
                strSalesChallan = strSalesChallan & ", Invoice_time, "
                strSalesChallan = strSalesChallan & "InvoiceAgainstMultipleSO, TextFileGenerated,From_Location) "
                strSalesChallan = strSalesChallan & " Values ('" & gstrUNITID & "','" & Trim(txtLocationCode.Text)
                strSalesChallan = strSalesChallan & "', '" & Trim(txtChallanNo.Text) & "',''"
                strSalesChallan = strSalesChallan & ",'" & Mid(Trim(CmbTransType.Text), 1, 1) & "', '" & Trim(txtVehNo.Text) & "','"
                strSalesChallan = strSalesChallan & "','','" & Trim(txtCustCode.Text)
                strSalesChallan = strSalesChallan & "','" & Trim(txtRefNo.Text) & "','" & Trim(mstrAmmNo) & "','0',"
                If OptDiscountPercentage.Checked = True Then 'In PerCentage
                    strSalesChallan = strSalesChallan & intDiscountType & "," & PdblDiscountAmount & "," & (Val(txtDiscountAmt.Text))
                Else
                    strSalesChallan = strSalesChallan & intDiscountType & "," & System.Math.Round(Val(txtDiscountAmt.Text)) & ",0"
                End If
                strSalesChallan = strSalesChallan & ",'','" & Trim(txtCarrServices.Text)
                strSalesChallan = strSalesChallan & "','" & Trim(CStr(Year(dtpDateDesc.Value))) & "',"
                strSalesChallan = strSalesChallan & System.Math.Round(Val(ctlInsurance.Text)) & ",'" & Trim(rsSaleConf.GetValue("Invoice_type")) & "','"
                strSalesChallan = strSalesChallan & Trim(mstrRGP) & "','"
                strSalesChallan = strSalesChallan & Trim(lblCustCodeDes.Text) & "',"
                strSalesChallan = strSalesChallan & Val(CStr(ldblTotalSaleTaxAmount)) & "," & Val(CStr(ldblTotalSurchargeTaxAmount)) & "," & System.Math.Round(Val(txtFreight.Text)) & ",'" & Trim(rsSaleConf.GetValue("Sub_Type")) & "','"
                strSalesChallan = strSalesChallan & Trim(txtSaleTaxType.Text) & "','"
                strSalesChallan = strSalesChallan & "0',0,'0','"
                strSalesChallan = strSalesChallan & getDateForDB(dtpDateDesc.Value) & "','" & lblCurrencyDes.Text & "',getdate(),'" & mP_User & "',  getdate() ,'" & mP_User & "','"
                strSalesChallan = strSalesChallan & Val(lblExchangeRateValue.Text) & "'," & ldblTotalInvoiceValue & ",'"
                strSalesChallan = strSalesChallan & Trim(txtSurchargeTaxType.Text) & "'," & Val(lblSaltax_Per.Text) & ","
                strSalesChallan = strSalesChallan & Val(lblSurcharge_Per.Text) & ",'" & Trim(txtRemarks.Text) & "',"
                strSalesChallan = strSalesChallan & ctlPerValue.Text & ",'" & Trim(txtSRVDINO.Text) & "','"
                strSalesChallan = strSalesChallan & Trim(txtSRVLocation.Text) & "','" & Trim(txtLoadingTaxType.Text) & "',"
                strSalesChallan = strSalesChallan & dblTotalLoadingcharges & "," & Val(lblLoadingcharge_per.Text) & ","
                If chkExciseExumpted.CheckState = System.Windows.Forms.CheckState.Checked Then
                    strSalesChallan = strSalesChallan & "1"
                Else
                    strSalesChallan = strSalesChallan & "0"
                End If
                strSalesChallan = strSalesChallan & ",'" & Trim(txtContactPerson.Text) & "','" & Trim(txtECC.Text) & "','" & Trim(txtLST.Text)
                strSalesChallan = strSalesChallan & "','" & Trim(txtAddress1.Text) & "','" & Trim(txtAddress2.Text) & "','" & Trim(txtAddress3.Text) & "'"
                If UCase(CmbInvType.Text) = "JOBWORK INVOICE" Then
                    If blnFIFO = True Then
                        strSalesChallan = strSalesChallan & ",1"
                    Else
                        strSalesChallan = strSalesChallan & ",0"
                    End If
                End If
                strSalesChallan = strSalesChallan & ",'" & Trim(txtUsLoc.Text) & "','" & Trim(txtSchTime.Text) & "'"
                strSalesChallan = strSalesChallan & ",'" & Trim(txtTCSTaxCode.Text) & "'," & Val(lblTCSTaxPerDes.Text)
                strSalesChallan = strSalesChallan & "," & dblTCSTaxAmount
                strSalesChallan = strSalesChallan & ",'" & Trim(txtECSSTaxType.Text) & "'," & Val(lblECSStax_Per.Text)
                strSalesChallan = strSalesChallan & "," & ldblTotalECSSTaxAmount
                strSalesChallan = strSalesChallan & ",'" & Trim(txtSECSSTaxType.Text) & "'," & Val(lblSECSStax_Per.Text)
                strSalesChallan = strSalesChallan & "," & ldblTotalSECSSTaxAmount
                strSalesChallan = strSalesChallan & "," & ldblTotalInvoiceValueRoundOff
                strSalesChallan = strSalesChallan & ",'" & Trim(lblCreditTerm.Text) & "'"
                strSalesChallan = strSalesChallan & ",substring(convert(varchar(20),Getdate()),13,len(getdate()))"
                strSalesChallan = strSalesChallan & "," & IIf(blnInvoiceAgainstMultipleSO, 1, 0) & ",0,'" & Trim(strStock_Loc) & "' "
                strSalesChallan = strSalesChallan & ")"
                rsSaleConf.ResultSetClose()
                rsSaleConf = Nothing
                strSalesDtl = ""
                With SpChEntry
                    For lintLoopCounter = 1 To .MaxRows
                        .Row = lintLoopCounter
                        .Col = GridHeader.InternalPartNo
                        lstrItemCode = Trim(.Text)
                        .Col = GridHeader.CustPartNo
                        lstrItemDrgno = Trim(.Text)
                        .Col = GridHeader.RatePerUnit
                        ldblItemRate = Val(.Text) / Val(ctlPerValue.Text)
                        .Col = GridHeader.CustSuppMatPerUnit
                        ldblItemCustMtrl = Val(.Text) / Val(ctlPerValue.Text)
                        .Col = GridHeader.Quantity
                        lintItemQuantity = Val(.Text)
                        '10808160
                        .Col = GridHeader.model
                        strModel = Trim(.Text)
                        '10808160
                        .Col = GridHeader.BinQty
                        dblBinQuantity = Val(.Text)
                        If blnInvoiceAgainstMultipleSO Then
                            .Col = GridHeader.CustRefNo
                            strCustRef = Trim(.Text)
                            .Col = GridHeader.AmendmentNo
                            StrAmendmentNo = Trim(.Text)
                            .Col = GridHeader.srvdino
                            strSrvDINo = Trim(.Text)
                            .Col = GridHeader.SRVLocation
                            strSRVLocation = Trim(.Text)
                            .Col = GridHeader.USLOC
                            strUSLoc = Trim(.Text)
                            .Col = GridHeader.SChTime
                            strSchTime = Trim(.Text)
                        Else
                            strCustRef = Trim(txtRefNo.Text)
                            StrAmendmentNo = Trim(txtAmendNo.Text)
                            strSrvDINo = Trim(txtSRVDINO.Text)
                            strSRVLocation = Trim(txtSRVLocation.Text)
                            strUSLoc = Trim(txtUsLoc.Text)
                            strSchTime = Trim(txtSchTime.Text)
                        End If
                        .Col = GridHeader.Packing
                        ldblItemPacking = Val(.Text)
                        .Col = GridHeader.EXC
                        lstrItemExciseCode = Trim(.Text)
                        .Col = GridHeader.CVD
                        lstrItemCVDCode = Trim(.Text)
                        .Col = GridHeader.SAD
                        lstrItemSADCode = Trim(.Text)
                        .Col = GridHeader.OthersPerUnit
                        ldblItemOthers = Val(.Text) / Val(ctlPerValue.Text) * lintItemQuantity
                        .Col = GridHeader.FromBox
                        ldblItemFromBox = Val(.Text)
                        .Col = GridHeader.ToBox
                        ldblItemToBox = Val(.Text)
                        .Col = GridHeader.delete
                        lstrItemDelete = Trim(.Text)
                        If UCase(CmbInvType.Text) = "SAMPLE INVOICE" Then
                            .Col = GridHeader.ToolCostPerUnit
                            ldblItemToolCost = Val(.Text) / Val(ctlPerValue.Text)
                        Else
                            .Col = GridHeader.ToolCost
                            ldblItemToolCost = Val(.Text) / Val(ctlPerValue.Text)
                        End If
                        If blnTotalToolCostRoundOff = True Then
                            ldblTotalToolCost = System.Math.Round(Val(CStr(lintItemQuantity * ldblItemToolCost)))
                        Else
                            ldblTotalToolCost = System.Math.Round(lintItemQuantity * ldblItemToolCost, intToolCostRoundOffDecimal)
                        End If
                        rsCustItemMst = New ClsResultSetDB
                        rsItemMst = New ClsResultSetDB
                        rsItemMst.GetResult("SELECT Description FROM Item_Mst WHERE UNIT_CODE='" & gstrUNITID & "' AND Item_Code ='" & Trim(lstrItemCode) & "'", ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                        rsCustItemMst.GetResult("SELECT Drg_desc FROM CustItem_Mst WHERE UNIT_CODE='" & gstrUNITID & "' AND Account_code ='" & Trim(txtCustCode.Text) & "'and Cust_DrgNo='" & lstrItemDrgno & "'and Item_code ='" & lstrItemCode & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        If UCase(Trim(lstrItemDelete)) <> "D" Then
                            strSalesDtl = Trim(strSalesDtl) & "INSERT INTO sales_Dtl(EOP_MODEL,UNIT_CODE,Location_Code,Doc_No,Suffix,Item_Code,Sales_Quantity,BinQuantity,"
                            strSalesDtl = strSalesDtl & "From_Box,To_Box,Rate,Sales_Tax,Excise_Tax,Packing,Others,Cust_Mtrl,"
                            strSalesDtl = strSalesDtl & "Year,Cust_Item_Code,Cust_Item_Desc,Tool_Cost,Measure_Code,Excise_type,SalesTax_type,CVD_type,SAD_type,Basic_Amount,Accessible_amount,CVD_Amount,SVD_amount, "
                            strSalesDtl = strSalesDtl & "Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,Excise_per,CVD_per,SVD_per,CustMtrl_Amount,ToolCost_Amount,PerValue,TotalExciseAmount, "
                            strSalesDtl = strSalesDtl & "Cust_ref, Amendment_No, SRVDINO, SRVLocation, USLOC, SchTime)"
                            strSalesDtl = strSalesDtl & " values ('" & strModel & "','" & gstrUNITID & "','" & Trim(txtLocationCode.Text) & "','"
                            strSalesDtl = strSalesDtl & Trim(txtChallanNo.Text) & "','','" & Trim(lstrItemCode) & "','" & Val(CStr(lintItemQuantity)) & "','" & dblBinQuantity & "','"
                            strSalesDtl = strSalesDtl & Val(CStr(ldblItemFromBox)) & "','" & Val(CStr(ldblItemToBox)) & "'," & Val(CStr(ldblItemRate)) & "," & Trim(lblSaltax_Per.Text) & ","
                            TempAccessibleVal = CalculateAccessibleValue(lintLoopCounter, ldblNetInsurenceValue, blnISInsExcisable)
                            If blnISExciseRoundOff Then
                                dblExcise_Amount = System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff, dblNetLoadingcharges))
                                strSalesDtl = strSalesDtl & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff, dblNetLoadingcharges))
                            Else
                                dblExcise_Amount = System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff, dblNetLoadingcharges), intExciseRoundOffDecimal)
                                strSalesDtl = strSalesDtl & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff, dblNetLoadingcharges), intExciseRoundOffDecimal)
                            End If
                            strSalesDtl = strSalesDtl & "," & Val(CStr(ldblItemPacking)) & "," & Val(CStr(ldblItemOthers)) & "," & Val(CStr(ldblItemCustMtrl)) & ",'"
                            strSalesDtl = strSalesDtl & Trim(CStr(Year(dtpDateDesc.Value))) & "','" & Trim(lstrItemDrgno) & "','" & IIf((Len(Trim(rsCustItemMst.GetValue("Drg_Desc"))) <= 0 Or Trim(CStr(rsCustItemMst.GetValue("Drg_Desc") = "Unknown"))), Trim(rsItemMst.GetValue("Description")), Trim(rsCustItemMst.GetValue("Drg_Desc"))) & "',"
                            If UCase(CmbInvType.Text) = "NORMAL INVOICE" Or UCase(CmbInvType.Text) = "EXPORT INVOICE" Or UCase(CmbInvType.Text) = "SERVICE INVOICE" Then
                                If UCase(CmbInvSubType.Text) <> "SCRAP" Then
                                    strSalesDtl = strSalesDtl & mdblToolCost(lintLoopCounter - 1) & ",'','"
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
                                strSalesDtl = strSalesDtl & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_CVD, blnEOUFlag, blnISExciseRoundOff, dblNetLoadingcharges))
                                strSalesDtl = strSalesDtl & "," & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_SAD, blnEOUFlag, blnISExciseRoundOff, dblNetLoadingcharges))
                            Else
                                strSalesDtl = strSalesDtl & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_CVD, blnEOUFlag, blnISExciseRoundOff, dblNetLoadingcharges), intExciseRoundOffDecimal)
                                strSalesDtl = strSalesDtl & "," & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_SAD, blnEOUFlag, blnISExciseRoundOff, dblNetLoadingcharges), intExciseRoundOffDecimal)
                            End If
                            strSalesDtl = strSalesDtl & ",GetDate(),'"
                            strSalesDtl = strSalesDtl & Trim(mP_User) & "', GetDate(),'" & Trim(mP_User) & "'," & GetTaxRate(lstrItemExciseCode, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='EXC'  and UNIT_CODE='" & gstrUNITID & "'") & "," & GetTaxRate(lstrItemCVDCode, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='CVD' and UNIT_CODE='" & gstrUNITID & "'") & "," & GetTaxRate(lstrItemSADCode, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='SAD' and UNIT_CODE='" & gstrUNITID & "'") & "," & System.Math.Round(Val(CStr(lintItemQuantity * ldblItemCustMtrl))) & "," & ldblTotalToolCost & "," & ctlPerValue.Text & ","


                            If blnISExciseRoundOff Then
                                strSalesDtl = strSalesDtl & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_ALLExcise, blnEOUFlag, blnISExciseRoundOff, dblNetLoadingcharges))
                            Else
                                strSalesDtl = strSalesDtl & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_ALLExcise, blnEOUFlag, blnISExciseRoundOff, dblNetLoadingcharges), intExciseRoundOffDecimal)
                            End If
                            strSalesDtl = strSalesDtl & ",'" & strCustRef & "','" & StrAmendmentNo & "','" & strSrvDINo & "'"
                            strSalesDtl = strSalesDtl & ",'" & strSRVLocation & "','" & strUSLoc & "','" & strSchTime & "')" & vbCrLf
                        End If

                        '10736222
                        strSql = "select dbo.UDF_ISCT2INVOICE( '" & gstrUNITID & "','" & txtCustCode.Text.Trim & "','" & CmbInvType.Text.Trim & "','" & CmbInvSubType.Text.Trim & "','" & txtRefNo.Text.Trim & "')"
                        If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSql)) = True Then
                            blnIsCt2 = True
                            strSqlct2qry = "INSERT INTO TMP_CT2_INVOICE_KNOCKOFF ([UNIT_CODE],[CUST_CODE],[SONO],[AMENDMENT_NO],[TMP_INVOICE_NO],[ITEM_CODE],[CUST_DRG_NO],[CURRENCY_CODE],[QTY],[RATE],[TOOL_COST],[EXCISE_TAX],[EXCISE_AMOUNT],[ECESS_TYPE],[SECESS_TYPE],[IP_ADDRESS]) "
                            strSqlct2qry = strSqlct2qry + " Values('" & gstrUNITID & "','" & txtCustCode.Text.Trim & "','" & txtRefNo.Text.Trim & "','" & txtAmendNo.Text.Trim & "','" & Me.txtChallanNo.Text.Trim & "',"
                            strSqlct2qry = strSqlct2qry + "'" & lstrItemCode.Trim & "','" & lstrItemDrgno.Trim & "','" & lblCurrencyDes.Text.Trim & "'," & Val(CStr(lintItemQuantity)) & "," & Val(CStr(ldblItemRate)) & "," & Val(mdblToolCost(lintLoopCounter - 1)) & ",'" & lstrItemExciseCode.Trim & "'," & dblExcise_Amount & ",'" & txtECSSTaxType.Text.Trim & "','" & txtSECSSTaxType.Text.Trim & "','" & gstrIpaddressWinSck & "' ) "
                            SqlConnectionclass.ExecuteNonQuery(strSqlct2qry)
                        End If
                        '10736222

                    Next
                End With
            Case "EDIT"
                lblCreditTerm.Text = "060"
                If UCase(mstrInvoiceType) = "INV" And UCase(mstrInvSubType) = "F" Then
                    If strStock_Loc = "01M1" Then strStock_Loc = "01B1"
                End If
                strSalesChallan = ""
                strSalesChallan = "UPDATE SalesChallan_Dtl SET Insurance = " & System.Math.Round(Val(ctlInsurance.Text))
                If blnISSalesTaxRoundOff Then
                    strSalesChallan = strSalesChallan & ",Sales_Tax_Amount =" & System.Math.Round(Val(CStr(ldblTotalSaleTaxAmount)))
                Else
                    strSalesChallan = strSalesChallan & ",Sales_Tax_Amount =" & System.Math.Round(Val(CStr(ldblTotalSaleTaxAmount)), intSaleTaxRoundOffDecimal)
                End If
                If blnISSurChargeTaxRoundOff Then
                    strSalesChallan = strSalesChallan & ",Surcharge_Sales_Tax_Amount =" & System.Math.Round(Val(CStr(ldblTotalSurchargeTaxAmount)))
                Else
                    strSalesChallan = strSalesChallan & ",Surcharge_Sales_Tax_Amount =" & System.Math.Round(Val(CStr(ldblTotalSurchargeTaxAmount)), intSSTRoundOffDecimal)
                End If
                strSalesChallan = strSalesChallan & ",Frieght_Amount=" & System.Math.Round(Val(txtFreight.Text))
                strSalesChallan = strSalesChallan & ",Discount_type=" & intDiscountType
                If OptDiscountPercentage.Checked = True Then 'In Percentage
                    strSalesChallan = strSalesChallan & ",Discount_Amount=" & PdblDiscountAmount
                    strSalesChallan = strSalesChallan & ",Discount_Per=" & Val(txtDiscountAmt.Text)
                Else
                    strSalesChallan = strSalesChallan & ",Discount_Amount=" & System.Math.Round(Val(txtDiscountAmt.Text), 0)
                    strSalesChallan = strSalesChallan & ",Discount_Per= 0"
                End If
                strSalesChallan = strSalesChallan & ",SalesTax_Type='" & Trim(txtSaleTaxType.Text) & "'"
                strSalesChallan = strSalesChallan & ",total_amount=" & ldblTotalInvoiceValue
                strSalesChallan = strSalesChallan & ",Surcharge_salesTaxType='" & Trim(txtSurchargeTaxType.Text) & "'"
                strSalesChallan = strSalesChallan & ",SalesTax_Per=" & Val(lblSaltax_Per.Text)
                strSalesChallan = strSalesChallan & ",Surcharge_SalesTax_Per=" & Val(lblSurcharge_Per.Text)
                strSalesChallan = strSalesChallan & ",Remarks = '" & Trim(txtRemarks.Text) & "'"
                strSalesChallan = strSalesChallan & ",SRVDINO = '" & Trim(txtSRVDINO.Text) & "'"
                strSalesChallan = strSalesChallan & ",SRVLocation = '" & Trim(txtSRVLocation.Text) & "'"
                strSalesChallan = strSalesChallan & ",LoadingChargeTaxType = '" & Trim(txtLoadingTaxType.Text) & "'"
                strSalesChallan = strSalesChallan & ",LoadingChargeTaxAmount = " & dblTotalLoadingcharges
                strSalesChallan = strSalesChallan & ",LoadingChargeTax_Per = " & Val(lblLoadingcharge_per.Text)
                If chkExciseExumpted.CheckState = System.Windows.Forms.CheckState.Checked Then
                    strSalesChallan = strSalesChallan & ",ExciseExumpted = " & 1
                Else
                    strSalesChallan = strSalesChallan & ",ExciseExumpted = " & 0
                End If
                strSalesChallan = strSalesChallan & ",ConsigneeContactPerson = '" & Trim(txtContactPerson.Text) & "'"
                strSalesChallan = strSalesChallan & ",ConsigneeECCNo = '" & Trim(txtECC.Text) & "'"
                strSalesChallan = strSalesChallan & ",ConsigneeLST = '" & Trim(txtLST.Text) & "'"
                strSalesChallan = strSalesChallan & ",ConsigneeAddress1 = '" & Trim(txtAddress1.Text) & "'"
                strSalesChallan = strSalesChallan & ",ConsigneeAddress2 = '" & Trim(txtAddress2.Text) & "'"
                strSalesChallan = strSalesChallan & ",ConsigneeAddress3 = '" & Trim(txtAddress3.Text) & "'"
                strSalesChallan = strSalesChallan & ",USLOC = '" & Trim(txtUsLoc.Text) & "'"
                strSalesChallan = strSalesChallan & ",Schtime = '" & Trim(txtSchTime.Text) & "'"
                strSalesChallan = strSalesChallan & ",TCSTax_Type = '" & txtTCSTaxCode.Text & "'"
                strSalesChallan = strSalesChallan & ",TCSTax_Per = " & Val(lblTCSTaxPerDes.Text)
                strSalesChallan = strSalesChallan & ",TCSTaxAmount = " & dblTCSTaxAmount
                strSalesChallan = strSalesChallan & ",ECESS_Type = '" & txtECSSTaxType.Text & "'"
                strSalesChallan = strSalesChallan & ",ECESS_Per = " & Val(lblECSStax_Per.Text)
                strSalesChallan = strSalesChallan & ",ECESS_Amount = " & ldblTotalECSSTaxAmount
                strSalesChallan = strSalesChallan & ",SECESS_Type = '" & txtSECSSTaxType.Text & "'"
                strSalesChallan = strSalesChallan & ",SECESS_Per = " & Val(lblSECSStax_Per.Text)
                strSalesChallan = strSalesChallan & ",SECESS_Amount = " & ldblTotalSECSSTaxAmount
                strSalesChallan = strSalesChallan & ",TotalInvoiceAmtRoundOff_diff = " & ldblTotalInvoiceValueRoundOff
                strSalesChallan = strSalesChallan & ",PAYMENT_TERMS = '" & Trim(lblCreditTerm.Text) & "'"
                strSalesChallan = strSalesChallan & ",Invoice_time = substring(Convert(VarChar(20), getDate()), 13, Len(getDate()))"
                strSalesChallan = strSalesChallan & ",InvoiceAgainstMultipleSO='" & IIf(blnInvoiceAgainstMultipleSO, 1, 0) & "'"
                strSalesChallan = strSalesChallan & ",TextFileGenerated=0 , from_location='" & Trim(strStock_Loc) & "'"
                strSalesChallan = strSalesChallan & " WHERE UNIT_CODE='" & gstrUNITID & "' and Location_Code ='" & Trim(txtLocationCode.Text) & "'"
                strSalesChallan = strSalesChallan & " and Doc_No ='" & Val(txtChallanNo.Text) & "'"
                strSalesDtl = ""
                strSalesDtlDelete = ""
                With SpChEntry
                    For lintLoopCounter = 1 To .MaxRows
                        .Row = lintLoopCounter
                        .Col = GridHeader.Quantity
                        lintItemQuantity = Val(.Text)
                        .Col = GridHeader.BinQty
                        dblBinQuantity = Val(.Text)
                        If dblBinQuantity <= 0 Then
                            MsgBox("Bin Quantity can't be zero.", MsgBoxStyle.Information, "eMpro")
                            SaveData = False
                            Exit Function
                        End If
                        .Col = GridHeader.CustPartNo
                        lstrItemDrgno = Trim(.Text)
                        .Col = GridHeader.delete
                        lstrItemDelete = Trim(.Text)
                        .Col = GridHeader.EXC
                        lstrItemExciseCode = Trim(.Text)
                        '10808160
                        .Col = GridHeader.model
                        strModel = Trim(.Text)
                        '10808160

                        .Col = GridHeader.CVD
                        lstrItemCVDCode = Trim(.Text)
                        .Col = GridHeader.SAD
                        lstrItemSADCode = Trim(.Text)
                        .Col = GridHeader.FromBox
                        ldblItemFromBox = Val(.Text)
                        .Col = GridHeader.ToBox
                        ldblItemToBox = Val(.Text)
                        If UCase(mstrInvoiceType) = "SMP" Then
                            .Col = GridHeader.ToolCostPerUnit
                            ldblItemToolCost = Val(.Text) / Val(ctlPerValue.Text)
                        Else
                            .Col = GridHeader.ToolCost
                            ldblItemToolCost = Val(.Text) / Val(ctlPerValue.Text)
                        End If
                        If blnInvoiceAgainstMultipleSO Then
                            .Col = GridHeader.CustRefNo
                            strCustRef = Trim(.Text)
                            .Col = GridHeader.AmendmentNo
                            StrAmendmentNo = Trim(.Text)
                            .Col = GridHeader.srvdino
                            strSrvDINo = Trim(.Text)
                            .Col = GridHeader.SRVLocation
                            strSRVLocation = Trim(.Text)
                            .Col = GridHeader.USLOC
                            strUSLoc = Trim(.Text)
                            .Col = GridHeader.SChTime
                            strSchTime = Trim(.Text)
                        Else
                            strCustRef = Trim(txtRefNo.Text)
                            StrAmendmentNo = Trim(txtAmendNo.Text)
                            strSrvDINo = Trim(txtSRVDINO.Text)
                            strSRVLocation = Trim(txtSRVLocation.Text)
                            strUSLoc = Trim(txtUsLoc.Text)
                            strSchTime = Trim(txtSchTime.Text)
                        End If
                        If blnTotalToolCostRoundOff = True Then
                            ldblTotalToolCost = System.Math.Round(Val(CStr(lintItemQuantity * ldblItemToolCost)))
                        Else
                            ldblTotalToolCost = System.Math.Round(lintItemQuantity * ldblItemToolCost, intToolCostRoundOffDecimal)
                        End If
                        If UCase(lstrItemDelete) <> "D" Then
                            strSalesDtl = Trim(strSalesDtl) & "UPDATE Sales_dtl SET EOP_MODEL='" & strModel & "',Sales_Quantity ='" & Val(CStr(lintItemQuantity)) & "',BinQuantity='" & dblBinQuantity & "',Sales_Tax =" & Trim(lblSaltax_Per.Text) & ","
                            strSalesDtl = Trim(strSalesDtl) & "CustMtrl_Amount= " & Val(CStr(lintItemQuantity * ldblItemCustMtrl)) & ",ToolCost_Amount=" & Val(CStr(ldblTotalToolCost))
                            TempAccessibleVal = CalculateAccessibleValue(lintLoopCounter, ldblNetInsurenceValue, blnISInsExcisable)
                            If blnISExciseRoundOff Then
                                dblExcise_Amount = System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff, dblNetLoadingcharges))
                                strSalesDtl = Trim(strSalesDtl) & ",Excise_Tax=" & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff, dblNetLoadingcharges))
                            Else
                                dblExcise_Amount = System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff, dblNetLoadingcharges), intExciseRoundOffDecimal)
                                strSalesDtl = Trim(strSalesDtl) & ",Excise_Tax=" & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff, dblNetLoadingcharges), intExciseRoundOffDecimal)
                            End If
                            strSalesDtl = Trim(strSalesDtl) & ",Excise_type='" & lstrItemExciseCode & "',SalesTax_type='" & Trim(txtSaleTaxType.Text) & "'"
                            strSalesDtl = Trim(strSalesDtl) & ",CVD_type='" & Trim(lstrItemCVDCode) & "',SAD_type='" & Trim(lstrItemSADCode) & "',Basic_Amount=" & CalculateBasicValue(lintLoopCounter, blnISBasicRoundOff)
                            strSalesDtl = Trim(strSalesDtl) & ",Accessible_amount=" & Val(CStr(TempAccessibleVal))
                            If blnISExciseRoundOff Then
                                strSalesDtl = Trim(strSalesDtl) & ",CVD_Amount=" & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_CVD, blnEOUFlag, blnISExciseRoundOff, dblNetLoadingcharges)) & ",SVD_amount=" & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_SAD, blnEOUFlag, blnISExciseRoundOff, dblNetLoadingcharges))
                            Else
                                strSalesDtl = Trim(strSalesDtl) & ",CVD_Amount=" & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_CVD, blnEOUFlag, blnISExciseRoundOff, dblNetLoadingcharges)) & ",SVD_amount=" & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_SAD, blnEOUFlag, blnISExciseRoundOff, dblNetLoadingcharges), intExciseRoundOffDecimal)
                            End If
                            strSalesDtl = Trim(strSalesDtl) & ",Excise_per=" & GetTaxRate(lstrItemExciseCode, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='EXC' and UNIT_CODE='" & gstrUNITID & "'")
                            strSalesDtl = Trim(strSalesDtl) & ",CVD_per=" & GetTaxRate(lstrItemCVDCode, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='CVD' and UNIT_CODE='" & gstrUNITID & "'")
                            strSalesDtl = Trim(strSalesDtl) & ",SVD_per=" & GetTaxRate(lstrItemSADCode, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='SAD' and UNIT_CODE='" & gstrUNITID & "'")
                            strSalesDtl = Trim(strSalesDtl) & ",Tool_Cost =" & ldblItemToolCost & ",From_box = " & ldblItemFromBox & ", To_box = " & ldblItemToBox
                            If blnISExciseRoundOff Then
                                strSalesDtl = Trim(strSalesDtl) & ",TotalExciseAmount =" & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_ALLExcise, blnEOUFlag, blnISExciseRoundOff, dblNetLoadingcharges))
                            Else
                                strSalesDtl = Trim(strSalesDtl) & ",TotalExciseAmount =" & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_ALLExcise, blnEOUFlag, blnISExciseRoundOff, dblNetLoadingcharges), intExciseRoundOffDecimal)
                            End If
                            strSalesDtl = Trim(strSalesDtl) & ",Cust_ref='" & strCustRef & "'"
                            strSalesDtl = Trim(strSalesDtl) & ",Amendment_No='" & StrAmendmentNo & "'"
                            strSalesDtl = Trim(strSalesDtl) & ",SRVDINO='" & strSrvDINo & "'"
                            strSalesDtl = Trim(strSalesDtl) & ",SRVLocation='" & strSRVLocation & "'"
                            strSalesDtl = Trim(strSalesDtl) & ",USLOC='" & strUSLoc & "'"
                            strSalesDtl = Trim(strSalesDtl) & ",SchTime='" & strSchTime & "'"
                            strSalesDtl = Trim(strSalesDtl) & " WHERE UNIT_CODE='" & gstrUNITID & "' and Location_Code ='" & Trim(txtLocationCode.Text) & "'"
                            strSalesDtl = Trim(strSalesDtl) & " and Doc_No =" & Val(txtChallanNo.Text) & " and Cust_Item_Code='"
                            strSalesDtl = Trim(strSalesDtl) & Trim(lstrItemDrgno) & "'" & vbCrLf
                        Else
                            strSalesDtlDelete = Trim(strSalesDtlDelete) & "DELETE Sales_dtl "
                            strSalesDtlDelete = Trim(strSalesDtlDelete) & " WHERE UNIT_CODE='" & gstrUNITID & "' and Location_Code ='" & Trim(txtLocationCode.Text) & "'"
                            strSalesDtlDelete = Trim(strSalesDtlDelete) & " and Doc_No =" & Val(txtChallanNo.Text) & " and Cust_Item_Code='"
                            strSalesDtlDelete = Trim(strSalesDtlDelete) & Trim(lstrItemDrgno) & "'" & vbCrLf
                        End If
                        '10736222
                        strSql = "select dbo.UDF_ISCT2INVOICE( '" & gstrUNITID & "','" & txtCustCode.Text.Trim & "','" & CmbInvType.Text.Trim & "','" & CmbInvSubType.Text.Trim & "','" & txtRefNo.Text.Trim & "')"
                        If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSql)) = True Then
                            blnIsCt2 = True
                            strSqlct2qry = "INSERT INTO TMP_CT2_INVOICE_KNOCKOFF ([UNIT_CODE],[CUST_CODE],[SONO],[AMENDMENT_NO],[TMP_INVOICE_NO],[ITEM_CODE],[CUST_DRG_NO],[CURRENCY_CODE],[QTY],[RATE],[TOOL_COST],[EXCISE_TAX],[EXCISE_AMOUNT],[ECESS_TYPE],[SECESS_TYPE],[IP_ADDRESS]) "
                            strSqlct2qry = strSqlct2qry + " Values('" & gstrUNITID & "','" & txtCustCode.Text.Trim & "','" & txtRefNo.Text.Trim & "','" & txtAmendNo.Text.Trim & "','" & Me.txtChallanNo.Text.Trim & "',"
                            strSqlct2qry = strSqlct2qry + "'" & lstrItemCode.Trim & "','" & lstrItemDrgno.Trim & "','" & lblCurrencyDes.Text.Trim & "'," & Val(CStr(lintItemQuantity)) & "," & Val(CStr(ldblItemRate)) & "," & Val(mdblToolCost(lintLoopCounter - 1)) & ",'" & lstrItemExciseCode.Trim & "'," & dblExcise_Amount & ",'" & txtECSSTaxType.Text.Trim & "','" & txtSECSSTaxType.Text.Trim & "','" & gstrIpaddressWinSck & "' ) "
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
                MsgBox("Unable To Save CT2 Invoice Knock Off Details." & vbCr & objValidateCmd.Parameters(objValidateCmd.Parameters.Count - 1).Value.ToString().Trim, MsgBoxStyle.Information, ResolveResString(100))
                objValidateCmd = Nothing
                SaveData = False
                Exit Function
            End If
            objValidateCmd = Nothing
            '10736222
        End If

        With mP_Connection
            blnActiveTrans = False
            .BeginTrans()
            blnActiveTrans = True
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
            blnActiveTrans = False
        End With
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        If blnActiveTrans Then mP_Connection.RollbackTrans()
        SaveData = False
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function SaveData_GST(ByVal Button As String) As Boolean
        '---------------------------------------------------------------------------------------
        'Name       :   SaveData
        'Type       :   Function
        'Author     :   Tapan Jain
        'Arguments  :
        'Return     :
        'Purpose    :
        '''***** Changes done by Ashutosh on 08-12-2005,Issue id:16518
        '---------------------------------------------------------------------------------------
        Dim ldblTotalBasicValue As Double
        Dim ldblTotalAccessibleValue As Double
        Dim lintLoopCounter As Short
        Dim ldblTempAccessibleVal As Double
        Dim ldblTotalExciseValue As Double
        Dim ldblTotalCVDValue As Double
        Dim ldblTotalSADValue As Double
        Dim ldbltempTotalExciseValue As Double
        Dim ldblTotalSaleTaxAmount As Double
        Dim ldblTotalSurchargeTaxAmount As Double
        Dim ldblNetInsurenceValue As Double
        Dim ldblTotalInvoiceValue As Double
        Dim ldblTotalOthersValues As Double
        Dim dblTotalLoadingcharges As Double
        Dim dblNetLoadingcharges As Double
        Dim dblTCSTaxAmount As Double
        Dim rsParameterData As ClsResultSetDB
        Dim strParamQuery As String
        Dim dblBasicForLoading As Double
        ''-----------Variable For Saving Data---------
        Dim strSalesChallan As String
        Dim updateSalesChallan As String
        Dim strSalesDtl As String
        Dim strSalesDtlDelete As String
        Dim rsCustItemMst As ClsResultSetDB
        Dim rsSaleConf As ClsResultSetDB
        Dim rsItemMst As ClsResultSetDB
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
        Dim PdblDiscountAmount As Double ' to store the value of Discount
        PdblDiscountAmount = 0
        Dim blnISInsExcisable As Boolean
        Dim blnEOUFlag As Boolean
        Dim blnISExciseRoundOff As Boolean
        Dim blnISSalesTaxRoundOff As Boolean
        Dim blnISSurChargeTaxRoundOff As Boolean
        Dim blnAddCustMatrl As Boolean
        Dim blnISBasicRoundOff As Boolean
        Dim ldblExciseValueForSaleTax As Double
        Dim blnTotalToolCostRoundOff As Boolean
        Dim ldblTotalToolCost As Double
        Dim blnInsIncSTax As Boolean
        Dim blnTCSTax As Boolean
        Dim VarDelete As Object
        Dim intNonDeletedRowCount As Short
        Dim intBasicRoundOffDecimal As Short
        Dim intSaleTaxRoundOffDecimal As Short
        Dim intExciseRoundOffDecimal As Short
        Dim intSSTRoundOffDecimal As Short
        Dim intTCSRoundOffDecimal As Short
        Dim intToolCostRoundOffDecimal As Short
        Dim blnActiveTrans As Boolean
        Dim blnECSSTax As Boolean
        Dim intECSSRoundOffDecimal As Short
        Dim ldblTotalECSSTaxAmount As Double
        Dim ldblTotalSECSSTaxAmount As Double
        Dim strCustRef As String
        Dim StrAmendmentNo As String
        Dim strSrvDINo As String
        Dim strSRVLocation As String
        Dim strUSLoc As String
        Dim strSchTime As String
        Dim blnTotalInvoiceAmount As Boolean
        Dim intTotalInvoiceAmountRoundOffDecimal As Short
        Dim ldblTotalInvoiceValueRoundOff As Double
        Dim dblBinQuantity As Double
        Dim strSql As String
        Dim strSqlct2qry As String
        Dim dblExcise_Amount As Double
        Dim blnIsCt2 As Boolean = False
        Dim strModel As String = ""

        On Error GoTo ErrHandler
        ldblTotalBasicValue = 0
        ldblTotalAccessibleValue = 0
        ldblTotalExciseValue = 0
        ldblTotalCVDValue = 0
        ldblTotalSADValue = 0
        ldblTotalSaleTaxAmount = 0
        ldblTotalSurchargeTaxAmount = 0
        ldblTotalInvoiceValue = 0
        ldblTotalOthersValues = 0
        ldblTotalCustMatrlValue = 0
        ldblExciseValueForSaleTax = 0
        PdblDiscountAmount = 0
        ldblTotalECSSTaxAmount = 0
        dblBinQuantity = 0

        Dim dblCGSTAmt As Double = 0
        Dim dblSGSTAmt As Double = 0
        Dim dblIGSTAmt As Double = 0
        Dim dblUTGSTAmt As Double = 0
        Dim dblCCESSAmt As Double = 0
        Dim strHSNSACCode As String = String.Empty
        Dim strHSNSACType As String = String.Empty
        Dim strCGSTType As String = String.Empty
        Dim strSGSTType As String = String.Empty
        Dim strIGSTType As String = String.Empty
        Dim strUTGSTType As String = String.Empty
        Dim strCCESSType As String = String.Empty
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
        Dim dblDiscountPercentLine As Double = 0
        Dim dblDiscountAmountLine As Double = 0
        Dim dblItemTotalLine As Double = 0
        Dim dblBasicAmtLine As Double = 0
        Dim dblAssessableAmtLine As Double = 0
        Dim dblTotalItemValue As Double = 0

        SaveData_GST = True

        strParamQuery = "SELECT InsExc_Excise,CustSupp_Inc,EOU_Flag, Basic_Roundoff, Basic_Roundoff_decimal, SalesTax_Roundoff, SalesTax_Roundoff_decimal, Excise_Roundoff, Excise_Roundoff_decimal, "
        strParamQuery = strParamQuery & " SST_Roundoff, SST_Roundoff_decimal, InsInc_SalesTax, TCSTax_Roundoff, TCSTax_Roundoff_decimal, TotalToolCostRoundoff, TotalToolCostRoundoff_Decimal, ECESS_Roundoff, ECESSRoundoff_Decimal,TotalInvoiceAmount_Roundoff,TotalInvoiceAmountRoundOff_Decimal FROM Sales_Parameter WHERE UNIT_CODE='" & gstrUNITID & "'"
        rsParameterData = New ClsResultSetDB
        rsParameterData.GetResult(strParamQuery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsParameterData.GetNoRows > 0 Then
            blnISInsExcisable = rsParameterData.GetValue("InsExc_Excise")
            blnEOUFlag = rsParameterData.GetValue("EOU_Flag")
            blnISBasicRoundOff = rsParameterData.GetValue("Basic_Roundoff")
            blnISExciseRoundOff = rsParameterData.GetValue("Excise_Roundoff")
            blnISSalesTaxRoundOff = rsParameterData.GetValue("SalesTax_Roundoff")
            blnISSurChargeTaxRoundOff = rsParameterData.GetValue("SST_Roundoff")
            blnAddCustMatrl = rsParameterData.GetValue("CustSupp_Inc")
            blnInsIncSTax = rsParameterData.GetValue("InsInc_SalesTax")
            blnTCSTax = rsParameterData.GetValue("TCSTax_Roundoff")
            blnTotalToolCostRoundOff = rsParameterData.GetValue("TotalToolCostRoundoff")
            intBasicRoundOffDecimal = rsParameterData.GetValue("Basic_Roundoff_decimal")
            intSaleTaxRoundOffDecimal = rsParameterData.GetValue("SalesTax_Roundoff_decimal")
            intExciseRoundOffDecimal = rsParameterData.GetValue("Excise_Roundoff_decimal")
            intSSTRoundOffDecimal = rsParameterData.GetValue("SST_Roundoff_decimal")
            intTCSRoundOffDecimal = rsParameterData.GetValue("TCSTax_Roundoff_decimal")
            intToolCostRoundOffDecimal = rsParameterData.GetValue("TotalToolCostRoundoff_decimal")
            blnECSSTax = rsParameterData.GetValue("ECESS_Roundoff")
            intECSSRoundOffDecimal = rsParameterData.GetValue("ECESSRoundoff_Decimal")
            blnTotalInvoiceAmount = rsParameterData.GetValue("TotalInvoiceAmount_RoundOff")
            intTotalInvoiceAmountRoundOffDecimal = rsParameterData.GetValue("TotalInvoiceAmountRoundOff_Decimal")
        Else
            MsgBox("No data define in Sales_Parameter Table", MsgBoxStyle.Critical, "eMPro")
            SaveData_GST = False
            rsParameterData.ResultSetClose()
            rsParameterData = Nothing
            Exit Function
        End If
        rsParameterData.ResultSetClose()
        rsParameterData = Nothing
        dblBasicForLoading = 0
        intNonDeletedRowCount = 0
        For lintLoopCounter = 1 To SpChEntry.MaxRows
            VarDelete = Nothing
            Call SpChEntry.GetText(GridHeader.delete, lintLoopCounter, VarDelete)
            If UCase(VarDelete) <> "D" Then
                dblBasicForLoading = dblBasicForLoading + CalculateBasicValue(lintLoopCounter, blnISBasicRoundOff)
                intNonDeletedRowCount = intNonDeletedRowCount + 1
            End If
        Next
        dblTotalLoadingcharges = CalculateLoadingchargesAmount(dblBasicForLoading, CDbl(lblLoadingcharge_per.Text))
        ldblNetInsurenceValue = System.Math.Round(Val(ctlInsurance.Text)) / intNonDeletedRowCount
        dblNetLoadingcharges = dblTotalLoadingcharges / intNonDeletedRowCount

        For lintLoopCounter = 1 To SpChEntry.MaxRows
            VarDelete = Nothing
            Call SpChEntry.GetText(GridHeader.delete, lintLoopCounter, VarDelete)
            If UCase(VarDelete) <> "D" Then
                SpChEntry.Row = lintLoopCounter : SpChEntry.Col = GridHeader.Accessible_Amt
                ldblTotalAccessibleValue += Val(SpChEntry.Text)
                SpChEntry.Row = lintLoopCounter : SpChEntry.Col = GridHeader.Quantity
                lintItemQuantity = Val(SpChEntry.Text)
                SpChEntry.Row = lintLoopCounter : SpChEntry.Col = GridHeader.OthersPerUnit
                ldblTotalOthersValues += ((Val(SpChEntry.Text) / Val(ctlPerValue.Text)) * lintItemQuantity)
                SpChEntry.Row = lintLoopCounter : SpChEntry.Col = GridHeader.CustSuppMatPerUnit
                ldblTotalCustMatrlValue += ((Val(SpChEntry.Text) / Val(ctlPerValue.Text)) * lintItemQuantity)
                SpChEntry.Row = lintLoopCounter : SpChEntry.Col = GridHeader.Discount_Amt
                PdblDiscountAmount += Val(SpChEntry.Text)
                SpChEntry.Row = lintLoopCounter : SpChEntry.Col = GridHeader.CGST_Amt
                dblCGSTAmt += Val(SpChEntry.Text)
                SpChEntry.Row = lintLoopCounter : SpChEntry.Col = GridHeader.SGST_Amt
                dblSGSTAmt += Val(SpChEntry.Text)
                SpChEntry.Row = lintLoopCounter : SpChEntry.Col = GridHeader.IGST_Amt
                dblIGSTAmt += Val(SpChEntry.Text)
                SpChEntry.Row = lintLoopCounter : SpChEntry.Col = GridHeader.UTGST_Amt
                dblUTGSTAmt += Val(SpChEntry.Text)
                SpChEntry.Row = lintLoopCounter : SpChEntry.Col = GridHeader.CSESS_TAX_VAL
                dblCCESSAmt += Val(SpChEntry.Text)
                SpChEntry.Row = lintLoopCounter : SpChEntry.Col = GridHeader.Item_Total
                dblTotalItemValue += Val(SpChEntry.Text)
            End If
        Next

        For lintLoopCounter = 1 To SpChEntry.MaxRows
            VarDelete = Nothing
            Call SpChEntry.GetText(GridHeader.delete, lintLoopCounter, VarDelete)
            If UCase(VarDelete) <> "D" Then
                ldblTotalBasicValue = ldblTotalBasicValue + CalculateBasicValue(lintLoopCounter, blnISBasicRoundOff)
                ' ldblTempAccessibleVal = CalculateAccessibleValue(lintLoopCounter, ldblNetInsurenceValue, blnISInsExcisable)
                If blnISExciseRoundOff Then
                    ldblTotalExciseValue = System.Math.Round(CalculateExciseValue(lintLoopCounter, ldblTempAccessibleVal, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff, dblNetLoadingcharges))
                    ldblTotalCVDValue = System.Math.Round(CalculateExciseValue(lintLoopCounter, ldblTempAccessibleVal, enumExciseType.RETURN_CVD, blnEOUFlag, blnISExciseRoundOff, dblNetLoadingcharges))
                    ldblTotalSADValue = System.Math.Round(CalculateExciseValue(lintLoopCounter, ldblTempAccessibleVal, enumExciseType.RETURN_SAD, blnEOUFlag, blnISExciseRoundOff, dblNetLoadingcharges))
                Else
                    ldblTotalExciseValue = System.Math.Round(CalculateExciseValue(lintLoopCounter, ldblTempAccessibleVal, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff, dblNetLoadingcharges), intExciseRoundOffDecimal)
                    ldblTotalCVDValue = System.Math.Round(CalculateExciseValue(lintLoopCounter, ldblTempAccessibleVal, enumExciseType.RETURN_CVD, blnEOUFlag, blnISExciseRoundOff, dblNetLoadingcharges), intExciseRoundOffDecimal)
                    ldblTotalSADValue = System.Math.Round(CalculateExciseValue(lintLoopCounter, ldblTempAccessibleVal, enumExciseType.RETURN_SAD, blnEOUFlag, blnISExciseRoundOff, dblNetLoadingcharges), intExciseRoundOffDecimal)
                End If
                ldblTotalAccessibleValue = ldblTotalAccessibleValue + ldblTempAccessibleVal
                SpChEntry.Row = lintLoopCounter : SpChEntry.Col = GridHeader.Quantity
                lintItemQuantity = Val(SpChEntry.Text)
                SpChEntry.Row = lintLoopCounter : SpChEntry.Col = GridHeader.OthersPerUnit
                ldblTotalOthersValues = ldblTotalOthersValues + ((Val(SpChEntry.Text) / Val(ctlPerValue.Text)) * lintItemQuantity)
                SpChEntry.Row = lintLoopCounter : SpChEntry.Col = GridHeader.CustSuppMatPerUnit
                ldblTotalCustMatrlValue = ldblTotalCustMatrlValue + ((Val(SpChEntry.Text) / Val(ctlPerValue.Text)) * lintItemQuantity)
                If blnEOU_FLAG Then
                    If blnISExciseRoundOff Then
                        ldblExciseValueForSaleTax = ldblExciseValueForSaleTax + System.Math.Round((ldblTotalExciseValue + ldblTotalCVDValue + ldblTotalSADValue) / 2)
                    Else
                        ldblExciseValueForSaleTax = ldblExciseValueForSaleTax + System.Math.Round((ldblTotalExciseValue + ldblTotalCVDValue + ldblTotalSADValue) / 2, intExciseRoundOffDecimal)
                    End If
                Else
                    If blnISExciseRoundOff Then
                        ldblExciseValueForSaleTax = ldblExciseValueForSaleTax + System.Math.Round(ldblTotalExciseValue)
                    Else
                        ldblExciseValueForSaleTax = ldblExciseValueForSaleTax + System.Math.Round(ldblTotalExciseValue, intExciseRoundOffDecimal)
                    End If
                End If
            End If
        Next
        If blnECSSTax Then
            ldblTotalSECSSTaxAmount = System.Math.Round(CalculateSECSSTaxValue(ldblExciseValueForSaleTax))
            ldblTotalECSSTaxAmount = System.Math.Round(CalculateECSSTaxValue(ldblExciseValueForSaleTax))
        Else
            ldblTotalSECSSTaxAmount = System.Math.Round(CalculateSECSSTaxValue(ldblExciseValueForSaleTax), intECSSRoundOffDecimal)
            ldblTotalECSSTaxAmount = System.Math.Round(CalculateECSSTaxValue(ldblExciseValueForSaleTax), intECSSRoundOffDecimal)
        End If
        If blnISSalesTaxRoundOff Then
            ldblTotalSaleTaxAmount = System.Math.Round(CalculateSalesTaxValue(ldblTotalBasicValue, ldblExciseValueForSaleTax + ldblTotalECSSTaxAmount + ldblTotalSECSSTaxAmount, blnInsIncSTax, CDbl(ctlInsurance.Text)))
        Else
            ldblTotalSaleTaxAmount = System.Math.Round(CalculateSalesTaxValue(ldblTotalBasicValue, ldblExciseValueForSaleTax + ldblTotalECSSTaxAmount + ldblTotalSECSSTaxAmount, blnInsIncSTax, CDbl(ctlInsurance.Text)), intSaleTaxRoundOffDecimal)
        End If
        If blnISSurChargeTaxRoundOff Then
            ldblTotalSurchargeTaxAmount = System.Math.Round(CalculateSurchargeTaxValue(ldblTotalSaleTaxAmount))
        Else
            ldblTotalSurchargeTaxAmount = System.Math.Round(CalculateSurchargeTaxValue(ldblTotalSaleTaxAmount), intSSTRoundOffDecimal)
        End If
        'If Val(txtDiscountAmt.Text) > 0 Then
        '    If OptDiscountValue.Checked = True Then
        '        PdblDiscountAmount = System.Math.Round(Val(txtDiscountAmt.Text), 0)
        '    Else
        '        If chkExciseExumpted.CheckState = System.Windows.Forms.CheckState.Unchecked Then
        '            ' to calculate Discount Amount by Percentage if Excise Duty is available
        '            PdblDiscountAmount = System.Math.Round(((ldblTotalBasicValue + ldblTotalExciseValue) * Val(txtDiscountAmt.Text)) / 100)
        '        Else
        '            ' to calculate Discount Amount by Percentage if Excise Duty is not available
        '            PdblDiscountAmount = System.Math.Round(((ldblTotalBasicValue) * Val(txtDiscountAmt.Text)) / 100)
        '        End If
        '    End If
        'Else
        ' if discount amount is not available
        'PdblDiscountAmount = 0
        'End If
        '****************************************************************
        ''-----------------------------------------------------------------------------------------------------------------------------
        'If blnAddCustMatrl Then
        '    If chkExciseExumpted.CheckState = System.Windows.Forms.CheckState.Unchecked Then
        '        ldblTotalInvoiceValue = (ldblTotalBasicValue + ldblExciseValueForSaleTax + ldblTotalSaleTaxAmount + ldblTotalSurchargeTaxAmount + ldblTotalECSSTaxAmount + ldblTotalSECSSTaxAmount + System.Math.Round(Val(txtFreight.Text)) + System.Math.Round(ldblTotalOthersValues) + System.Math.Round(Val(ctlInsurance.Text)) + System.Math.Round(ldblTotalCustMatrlValue, 2)) - PdblDiscountAmount
        '    Else
        '        ldblTotalInvoiceValue = (ldblTotalBasicValue + ldblTotalSaleTaxAmount + ldblTotalSurchargeTaxAmount + ldblTotalECSSTaxAmount + ldblTotalSECSSTaxAmount + System.Math.Round(Val(txtFreight.Text)) + System.Math.Round(ldblTotalOthersValues) + System.Math.Round(Val(ctlInsurance.Text)) + System.Math.Round(ldblTotalCustMatrlValue, 2)) - PdblDiscountAmount
        '    End If
        'Else
        '    If chkExciseExumpted.CheckState = System.Windows.Forms.CheckState.Unchecked Then
        '        ldblTotalInvoiceValue = (ldblTotalBasicValue + ldblExciseValueForSaleTax + ldblTotalSaleTaxAmount + ldblTotalSurchargeTaxAmount + ldblTotalECSSTaxAmount + ldblTotalSECSSTaxAmount + System.Math.Round(Val(txtFreight.Text)) + System.Math.Round(ldblTotalOthersValues) + System.Math.Round(Val(ctlInsurance.Text))) - PdblDiscountAmount
        '    Else
        '        ldblTotalInvoiceValue = (ldblTotalBasicValue + ldblTotalSaleTaxAmount + ldblTotalSurchargeTaxAmount + ldblTotalECSSTaxAmount + ldblTotalSECSSTaxAmount + System.Math.Round(Val(txtFreight.Text)) + System.Math.Round(ldblTotalOthersValues) + System.Math.Round(Val(ctlInsurance.Text))) - PdblDiscountAmount
        '    End If
        'End If

        If blnAddCustMatrl Then
            ldblTotalInvoiceValue = (ldblTotalBasicValue + dblCGSTAmt + dblSGSTAmt + dblIGSTAmt + dblUTGSTAmt + dblCCESSAmt + Math.Round(Val(txtFreight.Text)) + Math.Round(ldblTotalOthersValues) + Math.Round(Val(ctlInsurance.Text)) + Math.Round(ldblTotalCustMatrlValue, 2)) - PdblDiscountAmount
        Else
            ldblTotalInvoiceValue = (ldblTotalBasicValue + dblCGSTAmt + dblSGSTAmt + dblIGSTAmt + dblUTGSTAmt + dblCCESSAmt + Math.Round(Val(txtFreight.Text)) + Math.Round(ldblTotalOthersValues) + Math.Round(Val(ctlInsurance.Text))) - PdblDiscountAmount
        End If

        If Val(lblTCSTaxPerDes.Text) > 0 Then
            dblTCSTaxAmount = CalculateTCSTax(ldblTotalInvoiceValue, blnTCSTax, Val(lblTCSTaxPerDes.Text))
            ldblTotalInvoiceValue = ldblTotalInvoiceValue + dblTCSTaxAmount
        End If
        If blnTotalInvoiceAmount Then
            ldblTotalInvoiceValueRoundOff = ldblTotalInvoiceValue - System.Math.Round(ldblTotalInvoiceValue)
            ldblTotalInvoiceValue = System.Math.Round(ldblTotalInvoiceValue)
        Else
            ldblTotalInvoiceValueRoundOff = ldblTotalInvoiceValue - System.Math.Round(ldblTotalInvoiceValue, intTotalInvoiceAmountRoundOffDecimal)
            ldblTotalInvoiceValue = System.Math.Round(ldblTotalInvoiceValue, intTotalInvoiceAmountRoundOffDecimal)
        End If

        Dim strStock_Loc As String
        Dim rsLocation As ClsResultSetDB
        rsLocation = New ClsResultSetDB
        strStock_Loc = ""
        rsLocation.GetResult("Select Invoice_Type,Sub_Type from SaleConf where UNIT_CODE='" & gstrUNITID & "' AND Description ='" & Trim(CmbInvType.Text) & "'and Sub_type_Description ='" & Trim(CmbInvSubType.Text) & "' and (fin_start_date <= getdate() and fin_end_date >= getdate())")
        If rsLocation.GetNoRows > 0 Then
            strStock_Loc = StockLocationSalesConf(rsLocation.GetValue("Invoice_Type"), rsLocation.GetValue("Sub_Type"), "TYPE")
        Else
            MsgBox("Stock Location is not defined", vbInformation + vbOKOnly, ResolveResString(100))
            Exit Function
        End If
        Select Case Button
            Case "ADD"
                rsSaleConf = New ClsResultSetDB
                rsSaleConf.GetResult("Select Invoice_Type,Sub_Type from SaleConf where UNIT_CODE='" & gstrUNITID & "' AND Description ='" & Trim(CmbInvType.Text) & "'and Sub_type_Description ='" & Trim(CmbInvSubType.Text) & "' and (fin_start_date <= getdate() and fin_end_date >= getdate())")
                mstrInvType = rsSaleConf.GetValue("Invoice_Type")
                mstrInvoiceSubType = rsSaleConf.GetValue("Sub_Type")
                strSalesChallan = ""
                If UCase(CmbInvType.Text) <> "JOBWORK INVOICE" Then
                    mstrRGP = ""
                End If
                If UCase(CmbInvType.Text) = "NORMAL INVOICE" And UCase(CmbInvSubType.Text) = "FINISHED GOODS" Then
                    If strStock_Loc = "01M1" Then strStock_Loc = "01B1"
                End If
                strSalesChallan = "INSERT INTO SalesChallan_dtl (UNIT_CODE,Location_Code,Doc_No,Suffix,Transport_Type,Vehicle_No,"
                strSalesChallan = strSalesChallan & "From_Station,To_Station,Account_Code,Cust_Ref,"
                strSalesChallan = strSalesChallan & "Amendment_No,Bill_Flag,Discount_type,Discount_Amount,Discount_Per,Form3,Carriage_Name,"
                strSalesChallan = strSalesChallan & "Year,Insurance,invoice_Type,Ref_Doc_No,"
                strSalesChallan = strSalesChallan & "Cust_Name ,Sales_Tax_Amount , Surcharge_Sales_Tax_Amount,"
                strSalesChallan = strSalesChallan & "Frieght_Amount,Sub_Category,SalesTax_Type,SalesTax_FormNo,SalesTax_FormValue,"
                strSalesChallan = strSalesChallan & "Annex_no,Invoice_Date,Currency_code,Ent_dt,"
                strSalesChallan = strSalesChallan & "Ent_UserId,Upd_dt,Upd_UserId,Exchange_Rate,total_amount,Surcharge_salesTaxType,"
                strSalesChallan = strSalesChallan & "SalesTax_Per,Surcharge_SalesTax_Per,Remarks,PerValue,SRVDINO,SRVLocation,"
                strSalesChallan = strSalesChallan & "LoadingChargeTaxType,LoadingChargeTaxAmount,LoadingChargeTax_Per,ExciseExumpted,"
                strSalesChallan = strSalesChallan & "ConsigneeContactPerson,ConsigneeECCNo,ConsigneeLST,ConsigneeAddress1,"
                strSalesChallan = strSalesChallan & "ConsigneeAddress2,ConsigneeAddress3"
                If UCase(CmbInvType.Text) = "JOBWORK INVOICE" Then
                    strSalesChallan = strSalesChallan & ",Fifo_Flag"
                End If
                strSalesChallan = strSalesChallan & ",USLOC,Schtime,TCSTax_Type,TCSTax_Per,TCSTaxAmount,ECESS_Type, ECESS_Per, ECESS_Amount,SECESS_Type, SECESS_Per, SECESS_Amount"
                strSalesChallan = strSalesChallan & " ,TotalInvoiceAmtRoundOff_diff "
                '
                strSalesChallan = strSalesChallan & ",CGST_TOTAL_AMT,SGST_TOTAL_AMT,IGST_TOTAL_AMT,UTGST_TOTAL_AMT,CCESS_TOTAL_AMT,ITEM_TOTAL_VALUE"
                '
                strSalesChallan = strSalesChallan & ",Payment_Terms"
                strSalesChallan = strSalesChallan & ", Invoice_time, "
                strSalesChallan = strSalesChallan & "InvoiceAgainstMultipleSO, TextFileGenerated,From_Location) "
                strSalesChallan = strSalesChallan & " Values ('" & gstrUNITID & "','" & Trim(txtLocationCode.Text)
                strSalesChallan = strSalesChallan & "', '" & Trim(txtChallanNo.Text) & "',''"
                strSalesChallan = strSalesChallan & ",'" & Mid(Trim(CmbTransType.Text), 1, 1) & "', '" & Trim(txtVehNo.Text) & "','"
                strSalesChallan = strSalesChallan & "','','" & Trim(txtCustCode.Text)
                strSalesChallan = strSalesChallan & "','" & Trim(txtRefNo.Text) & "','" & Trim(mstrAmmNo) & "','0',"

                'If OptDiscountPercentage.Checked = True Then 'In PerCentage
                '    strSalesChallan = strSalesChallan & intDiscountType & "," & PdblDiscountAmount & "," & (Val(txtDiscountAmt.Text))
                'Else
                '    strSalesChallan = strSalesChallan & intDiscountType & "," & System.Math.Round(Val(txtDiscountAmt.Text)) & ",0"
                'End If
                intDiscountType = 0
                txtDiscountAmt.Text = 0
                strSalesChallan = strSalesChallan & intDiscountType & "," & PdblDiscountAmount & "," & (Val(txtDiscountAmt.Text))

                strSalesChallan = strSalesChallan & ",'','" & Trim(txtCarrServices.Text)
                strSalesChallan = strSalesChallan & "','" & Trim(CStr(Year(dtpDateDesc.Value))) & "',"
                strSalesChallan = strSalesChallan & System.Math.Round(Val(ctlInsurance.Text)) & ",'" & Trim(rsSaleConf.GetValue("Invoice_type")) & "','"
                strSalesChallan = strSalesChallan & Trim(mstrRGP) & "','"
                strSalesChallan = strSalesChallan & Trim(lblCustCodeDes.Text) & "',"
                strSalesChallan = strSalesChallan & Val(CStr(ldblTotalSaleTaxAmount)) & "," & Val(CStr(ldblTotalSurchargeTaxAmount)) & "," & System.Math.Round(Val(txtFreight.Text)) & ",'" & Trim(rsSaleConf.GetValue("Sub_Type")) & "','"
                strSalesChallan = strSalesChallan & Trim(txtSaleTaxType.Text) & "','"
                strSalesChallan = strSalesChallan & "0',0,'0','"
                strSalesChallan = strSalesChallan & getDateForDB(dtpDateDesc.Value) & "','" & lblCurrencyDes.Text & "',getdate(),'" & mP_User & "',  getdate() ,'" & mP_User & "','"
                strSalesChallan = strSalesChallan & Val(lblExchangeRateValue.Text) & "'," & ldblTotalInvoiceValue & ",'"
                strSalesChallan = strSalesChallan & Trim(txtSurchargeTaxType.Text) & "'," & Val(lblSaltax_Per.Text) & ","
                strSalesChallan = strSalesChallan & Val(lblSurcharge_Per.Text) & ",'" & Trim(txtRemarks.Text) & "',"
                strSalesChallan = strSalesChallan & ctlPerValue.Text & ",'" & Trim(txtSRVDINO.Text) & "','"
                strSalesChallan = strSalesChallan & Trim(txtSRVLocation.Text) & "','" & Trim(txtLoadingTaxType.Text) & "',"
                strSalesChallan = strSalesChallan & dblTotalLoadingcharges & "," & Val(lblLoadingcharge_per.Text) & ","
                If chkExciseExumpted.CheckState = System.Windows.Forms.CheckState.Checked Then
                    strSalesChallan = strSalesChallan & "1"
                Else
                    strSalesChallan = strSalesChallan & "0"
                End If
                strSalesChallan = strSalesChallan & ",'" & Trim(txtContactPerson.Text) & "','" & Trim(txtECC.Text) & "','" & Trim(txtLST.Text)
                strSalesChallan = strSalesChallan & "','" & Trim(txtAddress1.Text) & "','" & Trim(txtAddress2.Text) & "','" & Trim(txtAddress3.Text) & "'"
                If UCase(CmbInvType.Text) = "JOBWORK INVOICE" Then
                    If blnFIFO = True Then
                        strSalesChallan = strSalesChallan & ",1"
                    Else
                        strSalesChallan = strSalesChallan & ",0"
                    End If
                End If
                strSalesChallan = strSalesChallan & ",'" & Trim(txtUsLoc.Text) & "','" & Trim(txtSchTime.Text) & "'"
                strSalesChallan = strSalesChallan & ",'" & Trim(txtTCSTaxCode.Text) & "'," & Val(lblTCSTaxPerDes.Text)
                strSalesChallan = strSalesChallan & "," & dblTCSTaxAmount
                strSalesChallan = strSalesChallan & ",'" & Trim(txtECSSTaxType.Text) & "'," & Val(lblECSStax_Per.Text)
                strSalesChallan = strSalesChallan & "," & ldblTotalECSSTaxAmount
                strSalesChallan = strSalesChallan & ",'" & Trim(txtSECSSTaxType.Text) & "'," & Val(lblSECSStax_Per.Text)
                strSalesChallan = strSalesChallan & "," & ldblTotalSECSSTaxAmount
                strSalesChallan = strSalesChallan & "," & ldblTotalInvoiceValueRoundOff

                '
                strSalesChallan = strSalesChallan & "," & Val(dblCGSTAmt) & "," & Val(dblSGSTAmt) & "," & Val(dblIGSTAmt) & "," & Val(dblUTGSTAmt) & "," & Val(dblCCESSAmt) & "," & dblTotalItemValue
                '
                strSalesChallan = strSalesChallan & ",'" & Trim(lblCreditTerm.Text) & "'"
                strSalesChallan = strSalesChallan & ",substring(convert(varchar(20),Getdate()),13,len(getdate()))"
                strSalesChallan = strSalesChallan & "," & IIf(blnInvoiceAgainstMultipleSO, 1, 0) & ",0,'" & Trim(strStock_Loc) & "' "
                strSalesChallan = strSalesChallan & ")"
                rsSaleConf.ResultSetClose()
                rsSaleConf = Nothing
                strSalesDtl = ""
                With SpChEntry
                    For lintLoopCounter = 1 To .MaxRows
                        .Row = lintLoopCounter
                        .Col = GridHeader.InternalPartNo
                        lstrItemCode = Trim(.Text)
                        .Col = GridHeader.CustPartNo
                        lstrItemDrgno = Trim(.Text)
                        .Col = GridHeader.RatePerUnit
                        ldblItemRate = Val(.Text) / Val(ctlPerValue.Text)
                        .Col = GridHeader.CustSuppMatPerUnit
                        ldblItemCustMtrl = Val(.Text) / Val(ctlPerValue.Text)
                        .Col = GridHeader.Quantity
                        lintItemQuantity = Val(.Text)
                        '10808160
                        .Col = GridHeader.model
                        strModel = Trim(.Text)
                        '10808160
                        .Col = GridHeader.BinQty
                        dblBinQuantity = Val(.Text)
                        If blnInvoiceAgainstMultipleSO Then
                            .Col = GridHeader.CustRefNo
                            strCustRef = Trim(.Text)
                            .Col = GridHeader.AmendmentNo
                            StrAmendmentNo = Trim(.Text)
                            .Col = GridHeader.srvdino
                            strSrvDINo = Trim(.Text)
                            .Col = GridHeader.SRVLocation
                            strSRVLocation = Trim(.Text)
                            .Col = GridHeader.USLOC
                            strUSLoc = Trim(.Text)
                            .Col = GridHeader.SChTime
                            strSchTime = Trim(.Text)
                        Else
                            strCustRef = Trim(txtRefNo.Text)
                            StrAmendmentNo = Trim(txtAmendNo.Text)
                            strSrvDINo = Trim(txtSRVDINO.Text)
                            strSRVLocation = Trim(txtSRVLocation.Text)
                            strUSLoc = Trim(txtUsLoc.Text)
                            strSchTime = Trim(txtSchTime.Text)
                        End If
                        .Col = GridHeader.Packing
                        ldblItemPacking = Val(.Text)
                        .Col = GridHeader.EXC
                        lstrItemExciseCode = Trim(.Text)
                        .Col = GridHeader.CVD
                        lstrItemCVDCode = Trim(.Text)
                        .Col = GridHeader.SAD
                        lstrItemSADCode = Trim(.Text)
                        .Col = GridHeader.OthersPerUnit
                        ldblItemOthers = Val(.Text) / Val(ctlPerValue.Text) * lintItemQuantity
                        .Col = GridHeader.FromBox
                        ldblItemFromBox = Val(.Text)
                        .Col = GridHeader.ToBox
                        ldblItemToBox = Val(.Text)
                        .Col = GridHeader.delete
                        lstrItemDelete = Trim(.Text)
                        If UCase(CmbInvType.Text) = "SAMPLE INVOICE" Then
                            .Col = GridHeader.ToolCostPerUnit
                            ldblItemToolCost = Val(.Text) / Val(ctlPerValue.Text)
                        Else
                            .Col = GridHeader.ToolCost
                            ldblItemToolCost = Val(.Text) / Val(ctlPerValue.Text)
                        End If
                        If blnTotalToolCostRoundOff = True Then
                            ldblTotalToolCost = System.Math.Round(Val(CStr(lintItemQuantity * ldblItemToolCost)))
                        Else
                            ldblTotalToolCost = System.Math.Round(lintItemQuantity * ldblItemToolCost, intToolCostRoundOffDecimal)
                        End If

                        '
                        .Col = GridHeader.HSNSACCODE
                        strHSNSACCode = .Text
                        .Col = GridHeader.ISHSNORSAC
                        strHSNSACType = .Text
                        .Col = GridHeader.CGST_Type
                        strCGSTType = .Text
                        .Col = GridHeader.CGST_Percent
                        dblCGSTPercentLine = Val(.Text)
                        .Col = GridHeader.CGST_Amt
                        dblCGSTAmtLine = Val(.Text)
                        .Col = GridHeader.SGST_Type
                        strSGSTType = .Text
                        .Col = GridHeader.SGST_Percent
                        dblSGSTPercentLine = Val(.Text)
                        .Col = GridHeader.SGST_Amt
                        dblSGSTAmtLine = Val(.Text)
                        .Col = GridHeader.IGST_Type
                        strIGSTType = .Text
                        .Col = GridHeader.IGST_Percent
                        dblIGSTPercentLine = Val(.Text)
                        .Col = GridHeader.IGST_Amt
                        dblIGSTAmtLine = Val(.Text)
                        .Col = GridHeader.UTGST_Type
                        strUTGSTType = .Text
                        .Col = GridHeader.UTGST_Percent
                        dblUTGSTPercentLine = Val(.Text)
                        .Col = GridHeader.UTGST_Amt
                        dblUTGSTAmtLine = Val(.Text)
                        .Col = GridHeader.Discount_Percent
                        dblDiscountPercentLine = Val(.Text)
                        .Col = GridHeader.Discount_Amt
                        dblDiscountAmountLine = Val(.Text)
                        .Col = GridHeader.Item_Total
                        dblItemTotalLine = Val(.Text)
                        .Col = GridHeader.Basic_Value
                        dblBasicAmtLine = Val(.Text)
                        .Col = GridHeader.Accessible_Amt
                        dblAssessableAmtLine = Val(.Text)
                        '
                        rsCustItemMst = New ClsResultSetDB
                        rsItemMst = New ClsResultSetDB
                        rsItemMst.GetResult("SELECT Description FROM Item_Mst WHERE UNIT_CODE='" & gstrUNITID & "' AND Item_Code ='" & Trim(lstrItemCode) & "'", ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                        rsCustItemMst.GetResult("SELECT Drg_desc FROM CustItem_Mst WHERE UNIT_CODE='" & gstrUNITID & "' AND Account_code ='" & Trim(txtCustCode.Text) & "'and Cust_DrgNo='" & lstrItemDrgno & "'and Item_code ='" & lstrItemCode & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        If UCase(Trim(lstrItemDelete)) <> "D" Then
                            strSalesDtl = Trim(strSalesDtl) & "INSERT INTO sales_Dtl(EOP_MODEL,UNIT_CODE,Location_Code,Doc_No,Suffix,Item_Code,Sales_Quantity,BinQuantity,"
                            strSalesDtl = strSalesDtl & "From_Box,To_Box,Rate,Sales_Tax,Excise_Tax,Packing,Others,Cust_Mtrl,"
                            strSalesDtl = strSalesDtl & "Year,Cust_Item_Code,Cust_Item_Desc,Tool_Cost,Measure_Code,Excise_type,SalesTax_type,CVD_type,SAD_type,Basic_Amount,Accessible_amount,CVD_Amount,SVD_amount, "
                            strSalesDtl = strSalesDtl & "Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,Excise_per,CVD_per,SVD_per,CustMtrl_Amount,ToolCost_Amount,PerValue,TotalExciseAmount, "
                            strSalesDtl = strSalesDtl & "HSNSACCODE,ISHSNORSAC,CGSTTXRT_TYPE,CGST_PERCENT,CGST_AMT,SGSTTXRT_TYPE,SGST_PERCENT,"
                            strSalesDtl = strSalesDtl & "SGST_AMT,IGSTTXRT_TYPE,IGST_PERCENT,IGST_AMT,UTGSTTXRT_TYPE,UTGST_PERCENT,UTGST_AMT,"
                            strSalesDtl = strSalesDtl & "COMPENSATION_CESS_TYPE,COMPENSATION_CESS_PERCENT,COMPENSATION_CESS_AMT,Discount_perc,Discount_amt,ITEM_VALUE,"
                            strSalesDtl = strSalesDtl & "Cust_ref, Amendment_No, SRVDINO, SRVLocation, USLOC, SchTime)"
                            strSalesDtl = strSalesDtl & " values ('" & strModel & "','" & gstrUNITID & "','" & Trim(txtLocationCode.Text) & "','"
                            strSalesDtl = strSalesDtl & Trim(txtChallanNo.Text) & "','','" & Trim(lstrItemCode) & "','" & Val(CStr(lintItemQuantity)) & "','" & dblBinQuantity & "','"
                            strSalesDtl = strSalesDtl & Val(CStr(ldblItemFromBox)) & "','" & Val(CStr(ldblItemToBox)) & "'," & Val(CStr(ldblItemRate)) & "," & Trim(lblSaltax_Per.Text) & ","
                            TempAccessibleVal = CalculateAccessibleValue(lintLoopCounter, ldblNetInsurenceValue, blnISInsExcisable)
                            If blnISExciseRoundOff Then
                                dblExcise_Amount = System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff, dblNetLoadingcharges))
                                strSalesDtl = strSalesDtl & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff, dblNetLoadingcharges))
                            Else
                                dblExcise_Amount = System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff, dblNetLoadingcharges), intExciseRoundOffDecimal)
                                strSalesDtl = strSalesDtl & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff, dblNetLoadingcharges), intExciseRoundOffDecimal)
                            End If
                            strSalesDtl = strSalesDtl & "," & Val(CStr(ldblItemPacking)) & "," & Val(CStr(ldblItemOthers)) & "," & Val(CStr(ldblItemCustMtrl)) & ",'"
                            strSalesDtl = strSalesDtl & Trim(CStr(Year(dtpDateDesc.Value))) & "','" & Trim(lstrItemDrgno) & "','" & IIf((Len(Trim(rsCustItemMst.GetValue("Drg_Desc"))) <= 0 Or Trim(CStr(rsCustItemMst.GetValue("Drg_Desc") = "Unknown"))), Trim(rsItemMst.GetValue("Description")), Trim(rsCustItemMst.GetValue("Drg_Desc"))) & "',"
                            If UCase(CmbInvType.Text) = "NORMAL INVOICE" Or UCase(CmbInvType.Text) = "EXPORT INVOICE" Or UCase(CmbInvType.Text) = "SERVICE INVOICE" Then
                                If UCase(CmbInvSubType.Text) <> "SCRAP" Then
                                    strSalesDtl = strSalesDtl & mdblToolCost(lintLoopCounter - 1) & ",'','"
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
                                strSalesDtl = strSalesDtl & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_CVD, blnEOUFlag, blnISExciseRoundOff, dblNetLoadingcharges))
                                strSalesDtl = strSalesDtl & "," & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_SAD, blnEOUFlag, blnISExciseRoundOff, dblNetLoadingcharges))
                            Else
                                strSalesDtl = strSalesDtl & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_CVD, blnEOUFlag, blnISExciseRoundOff, dblNetLoadingcharges), intExciseRoundOffDecimal)
                                strSalesDtl = strSalesDtl & "," & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_SAD, blnEOUFlag, blnISExciseRoundOff, dblNetLoadingcharges), intExciseRoundOffDecimal)
                            End If
                            strSalesDtl = strSalesDtl & ",GetDate(),'"
                            strSalesDtl = strSalesDtl & Trim(mP_User) & "', GetDate(),'" & Trim(mP_User) & "'," & GetTaxRate(lstrItemExciseCode, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='EXC'  and UNIT_CODE='" & gstrUNITID & "'") & "," & GetTaxRate(lstrItemCVDCode, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='CVD' and UNIT_CODE='" & gstrUNITID & "'") & "," & GetTaxRate(lstrItemSADCode, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='SAD' and UNIT_CODE='" & gstrUNITID & "'") & "," & System.Math.Round(Val(CStr(lintItemQuantity * ldblItemCustMtrl))) & "," & ldblTotalToolCost & "," & ctlPerValue.Text & ","


                            If blnISExciseRoundOff Then
                                strSalesDtl = strSalesDtl & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_ALLExcise, blnEOUFlag, blnISExciseRoundOff, dblNetLoadingcharges))
                            Else
                                strSalesDtl = strSalesDtl & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_ALLExcise, blnEOUFlag, blnISExciseRoundOff, dblNetLoadingcharges), intExciseRoundOffDecimal)
                            End If

                            strSalesDtl = strSalesDtl & ",'" & strHSNSACCode & "','" & strHSNSACType & "'"
                            strSalesDtl = strSalesDtl & ",'" & strCGSTType & "'," & dblCGSTPercentLine & "," & dblCGSTAmtLine & ""
                            strSalesDtl = strSalesDtl & ",'" & strSGSTType & "'," & dblSGSTPercentLine & "," & dblSGSTAmtLine & ""
                            strSalesDtl = strSalesDtl & ",'" & strIGSTType & "'," & dblIGSTPercentLine & "," & dblIGSTAmtLine & ""
                            strSalesDtl = strSalesDtl & ",'" & strUTGSTType & "'," & dblUTGSTPercentLine & "," & dblUTGSTAmtLine & ""
                            strSalesDtl = strSalesDtl & ",'" & strCCESSType & "'," & dblCCESSPercentLine & "," & dblCCESSAmtLine & "," & dblDiscountPercentLine & "," & dblDiscountAmountLine & "," & dblItemTotalLine & ""

                            strSalesDtl = strSalesDtl & ",'" & strCustRef & "','" & StrAmendmentNo & "','" & strSrvDINo & "'"
                            strSalesDtl = strSalesDtl & ",'" & strSRVLocation & "','" & strUSLoc & "','" & strSchTime & "')" & vbCrLf
                        End If

                        '10736222
                        strSql = "select dbo.UDF_ISCT2INVOICE( '" & gstrUNITID & "','" & txtCustCode.Text.Trim & "','" & CmbInvType.Text.Trim & "','" & CmbInvSubType.Text.Trim & "','" & txtRefNo.Text.Trim & "')"
                        If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSql)) = True Then
                            blnIsCt2 = True
                            strSqlct2qry = "INSERT INTO TMP_CT2_INVOICE_KNOCKOFF ([UNIT_CODE],[CUST_CODE],[SONO],[AMENDMENT_NO],[TMP_INVOICE_NO],[ITEM_CODE],[CUST_DRG_NO],[CURRENCY_CODE],[QTY],[RATE],[TOOL_COST],[EXCISE_TAX],[EXCISE_AMOUNT],[ECESS_TYPE],[SECESS_TYPE],[IP_ADDRESS]) "
                            strSqlct2qry = strSqlct2qry + " Values('" & gstrUNITID & "','" & txtCustCode.Text.Trim & "','" & txtRefNo.Text.Trim & "','" & txtAmendNo.Text.Trim & "','" & Me.txtChallanNo.Text.Trim & "',"
                            strSqlct2qry = strSqlct2qry + "'" & lstrItemCode.Trim & "','" & lstrItemDrgno.Trim & "','" & lblCurrencyDes.Text.Trim & "'," & Val(CStr(lintItemQuantity)) & "," & Val(CStr(ldblItemRate)) & "," & Val(mdblToolCost(lintLoopCounter - 1)) & ",'" & lstrItemExciseCode.Trim & "'," & dblExcise_Amount & ",'" & txtECSSTaxType.Text.Trim & "','" & txtSECSSTaxType.Text.Trim & "','" & gstrIpaddressWinSck & "' ) "
                            SqlConnectionclass.ExecuteNonQuery(strSqlct2qry)
                        End If
                        '10736222

                    Next
                End With
            Case "EDIT"
                lblCreditTerm.Text = "060"
                If UCase(mstrInvoiceType) = "INV" And UCase(mstrInvSubType) = "F" Then
                    If strStock_Loc = "01M1" Then strStock_Loc = "01B1"
                End If
                strSalesChallan = ""
                strSalesChallan = "UPDATE SalesChallan_Dtl SET Insurance = " & System.Math.Round(Val(ctlInsurance.Text))
                If blnISSalesTaxRoundOff Then
                    strSalesChallan = strSalesChallan & ",Sales_Tax_Amount =" & System.Math.Round(Val(CStr(ldblTotalSaleTaxAmount)))
                Else
                    strSalesChallan = strSalesChallan & ",Sales_Tax_Amount =" & System.Math.Round(Val(CStr(ldblTotalSaleTaxAmount)), intSaleTaxRoundOffDecimal)
                End If
                If blnISSurChargeTaxRoundOff Then
                    strSalesChallan = strSalesChallan & ",Surcharge_Sales_Tax_Amount =" & System.Math.Round(Val(CStr(ldblTotalSurchargeTaxAmount)))
                Else
                    strSalesChallan = strSalesChallan & ",Surcharge_Sales_Tax_Amount =" & System.Math.Round(Val(CStr(ldblTotalSurchargeTaxAmount)), intSSTRoundOffDecimal)
                End If
                strSalesChallan = strSalesChallan & ",Frieght_Amount=" & System.Math.Round(Val(txtFreight.Text))

                'strSalesChallan = strSalesChallan & ",Discount_type=" & intDiscountType

                'If OptDiscountPercentage.Checked = True Then 'In Percentage
                '    strSalesChallan = strSalesChallan & ",Discount_Amount=" & PdblDiscountAmount
                '    strSalesChallan = strSalesChallan & ",Discount_Per=" & Val(txtDiscountAmt.Text)
                'Else
                '    strSalesChallan = strSalesChallan & ",Discount_Amount=" & System.Math.Round(Val(txtDiscountAmt.Text), 0)
                '    strSalesChallan = strSalesChallan & ",Discount_Per= 0"
                'End If

                strSalesChallan = strSalesChallan & ",Discount_type=0"
                strSalesChallan = strSalesChallan & ",Discount_Amount=" & PdblDiscountAmount
                strSalesChallan = strSalesChallan & ",Discount_Per= 0"


                strSalesChallan = strSalesChallan & ",SalesTax_Type='" & Trim(txtSaleTaxType.Text) & "'"
                strSalesChallan = strSalesChallan & ",total_amount=" & ldblTotalInvoiceValue
                strSalesChallan = strSalesChallan & ",Surcharge_salesTaxType='" & Trim(txtSurchargeTaxType.Text) & "'"
                strSalesChallan = strSalesChallan & ",SalesTax_Per=" & Val(lblSaltax_Per.Text)
                strSalesChallan = strSalesChallan & ",Surcharge_SalesTax_Per=" & Val(lblSurcharge_Per.Text)
                strSalesChallan = strSalesChallan & ",Remarks = '" & Trim(txtRemarks.Text) & "'"
                strSalesChallan = strSalesChallan & ",SRVDINO = '" & Trim(txtSRVDINO.Text) & "'"
                strSalesChallan = strSalesChallan & ",SRVLocation = '" & Trim(txtSRVLocation.Text) & "'"
                strSalesChallan = strSalesChallan & ",LoadingChargeTaxType = '" & Trim(txtLoadingTaxType.Text) & "'"
                strSalesChallan = strSalesChallan & ",LoadingChargeTaxAmount = " & dblTotalLoadingcharges
                strSalesChallan = strSalesChallan & ",LoadingChargeTax_Per = " & Val(lblLoadingcharge_per.Text)
                If chkExciseExumpted.CheckState = System.Windows.Forms.CheckState.Checked Then
                    strSalesChallan = strSalesChallan & ",ExciseExumpted = " & 1
                Else
                    strSalesChallan = strSalesChallan & ",ExciseExumpted = " & 0
                End If
                strSalesChallan = strSalesChallan & ",ConsigneeContactPerson = '" & Trim(txtContactPerson.Text) & "'"
                strSalesChallan = strSalesChallan & ",ConsigneeECCNo = '" & Trim(txtECC.Text) & "'"
                strSalesChallan = strSalesChallan & ",ConsigneeLST = '" & Trim(txtLST.Text) & "'"
                strSalesChallan = strSalesChallan & ",ConsigneeAddress1 = '" & Trim(txtAddress1.Text) & "'"
                strSalesChallan = strSalesChallan & ",ConsigneeAddress2 = '" & Trim(txtAddress2.Text) & "'"
                strSalesChallan = strSalesChallan & ",ConsigneeAddress3 = '" & Trim(txtAddress3.Text) & "'"
                strSalesChallan = strSalesChallan & ",USLOC = '" & Trim(txtUsLoc.Text) & "'"
                strSalesChallan = strSalesChallan & ",Schtime = '" & Trim(txtSchTime.Text) & "'"
                strSalesChallan = strSalesChallan & ",TCSTax_Type = '" & txtTCSTaxCode.Text & "'"
                strSalesChallan = strSalesChallan & ",TCSTax_Per = " & Val(lblTCSTaxPerDes.Text)
                strSalesChallan = strSalesChallan & ",TCSTaxAmount = " & dblTCSTaxAmount
                strSalesChallan = strSalesChallan & ",ECESS_Type = '" & txtECSSTaxType.Text & "'"
                strSalesChallan = strSalesChallan & ",ECESS_Per = " & Val(lblECSStax_Per.Text)
                strSalesChallan = strSalesChallan & ",ECESS_Amount = " & ldblTotalECSSTaxAmount
                strSalesChallan = strSalesChallan & ",SECESS_Type = '" & txtSECSSTaxType.Text & "'"
                strSalesChallan = strSalesChallan & ",SECESS_Per = " & Val(lblSECSStax_Per.Text)
                strSalesChallan = strSalesChallan & ",SECESS_Amount = " & ldblTotalSECSSTaxAmount
                strSalesChallan = strSalesChallan & ",TotalInvoiceAmtRoundOff_diff = " & ldblTotalInvoiceValueRoundOff
                strSalesChallan = strSalesChallan & ",PAYMENT_TERMS = '" & Trim(lblCreditTerm.Text) & "'"
                strSalesChallan = strSalesChallan & ",Invoice_time = substring(Convert(VarChar(20), getDate()), 13, Len(getDate()))"
                strSalesChallan = strSalesChallan & ",InvoiceAgainstMultipleSO='" & IIf(blnInvoiceAgainstMultipleSO, 1, 0) & "'"
                strSalesChallan = strSalesChallan & ",TextFileGenerated=0 , from_location='" & Trim(strStock_Loc) & "'"

                strSalesChallan = strSalesChallan & ",CGST_TOTAL_AMT = " & Val(dblCGSTAmt)
                strSalesChallan = strSalesChallan & ",SGST_TOTAL_AMT = " & Val(dblSGSTAmt)
                strSalesChallan = strSalesChallan & ",IGST_TOTAL_AMT = " & Val(dblIGSTAmt)
                strSalesChallan = strSalesChallan & ",UTGST_TOTAL_AMT = " & Val(dblUTGSTAmt)
                strSalesChallan = strSalesChallan & ",CCESS_TOTAL_AMT = " & Val(dblCCESSAmt)
                strSalesChallan = strSalesChallan & ",ITEM_TOTAL_VALUE = " & dblTotalItemValue

                strSalesChallan = strSalesChallan & " WHERE UNIT_CODE='" & gstrUNITID & "' and Location_Code ='" & Trim(txtLocationCode.Text) & "'"
                strSalesChallan = strSalesChallan & " and Doc_No ='" & Val(txtChallanNo.Text) & "'"
                strSalesDtl = ""
                strSalesDtlDelete = ""
                With SpChEntry
                    For lintLoopCounter = 1 To .MaxRows
                        .Row = lintLoopCounter
                        .Col = GridHeader.Quantity
                        lintItemQuantity = Val(.Text)
                        .Col = GridHeader.BinQty
                        dblBinQuantity = Val(.Text)
                        If dblBinQuantity <= 0 Then
                            MsgBox("Bin Quantity can't be zero.", MsgBoxStyle.Information, "eMpro")
                            SaveData_GST = False
                            Exit Function
                        End If
                        .Col = GridHeader.CustPartNo
                        lstrItemDrgno = Trim(.Text)
                        .Col = GridHeader.delete
                        lstrItemDelete = Trim(.Text)
                        .Col = GridHeader.EXC
                        lstrItemExciseCode = Trim(.Text)
                        '10808160
                        .Col = GridHeader.model
                        strModel = Trim(.Text)
                        '10808160

                        .Col = GridHeader.CVD
                        lstrItemCVDCode = Trim(.Text)
                        .Col = GridHeader.SAD
                        lstrItemSADCode = Trim(.Text)
                        .Col = GridHeader.FromBox
                        ldblItemFromBox = Val(.Text)
                        .Col = GridHeader.ToBox
                        ldblItemToBox = Val(.Text)
                        If UCase(mstrInvoiceType) = "SMP" Then
                            .Col = GridHeader.ToolCostPerUnit
                            ldblItemToolCost = Val(.Text) / Val(ctlPerValue.Text)
                        Else
                            .Col = GridHeader.ToolCost
                            ldblItemToolCost = Val(.Text) / Val(ctlPerValue.Text)
                        End If
                        If blnInvoiceAgainstMultipleSO Then
                            .Col = GridHeader.CustRefNo
                            strCustRef = Trim(.Text)
                            .Col = GridHeader.AmendmentNo
                            StrAmendmentNo = Trim(.Text)
                            .Col = GridHeader.srvdino
                            strSrvDINo = Trim(.Text)
                            .Col = GridHeader.SRVLocation
                            strSRVLocation = Trim(.Text)
                            .Col = GridHeader.USLOC
                            strUSLoc = Trim(.Text)
                            .Col = GridHeader.SChTime
                            strSchTime = Trim(.Text)
                        Else
                            strCustRef = Trim(txtRefNo.Text)
                            StrAmendmentNo = Trim(txtAmendNo.Text)
                            strSrvDINo = Trim(txtSRVDINO.Text)
                            strSRVLocation = Trim(txtSRVLocation.Text)
                            strUSLoc = Trim(txtUsLoc.Text)
                            strSchTime = Trim(txtSchTime.Text)
                        End If
                        If blnTotalToolCostRoundOff = True Then
                            ldblTotalToolCost = System.Math.Round(Val(CStr(lintItemQuantity * ldblItemToolCost)))
                        Else
                            ldblTotalToolCost = System.Math.Round(lintItemQuantity * ldblItemToolCost, intToolCostRoundOffDecimal)
                        End If

                        .Col = GridHeader.HSNSACCODE
                        strHSNSACCode = .Text
                        .Col = GridHeader.ISHSNORSAC
                        strHSNSACType = .Text
                        .Col = GridHeader.CGST_Type
                        strCGSTType = .Text
                        .Col = GridHeader.CGST_Percent
                        dblCGSTPercentLine = Val(.Text)
                        .Col = GridHeader.CGST_Amt
                        dblCGSTAmtLine = Val(.Text)
                        .Col = GridHeader.SGST_Type
                        strSGSTType = .Text
                        .Col = GridHeader.SGST_Percent
                        dblSGSTPercentLine = Val(.Text)
                        .Col = GridHeader.SGST_Amt
                        dblSGSTAmtLine = Val(.Text)
                        .Col = GridHeader.IGST_Type
                        strIGSTType = .Text
                        .Col = GridHeader.IGST_Percent
                        dblIGSTPercentLine = Val(.Text)
                        .Col = GridHeader.IGST_Amt
                        dblIGSTAmtLine = Val(.Text)
                        .Col = GridHeader.UTGST_Type
                        strUTGSTType = .Text
                        .Col = GridHeader.UTGST_Percent
                        dblUTGSTPercentLine = Val(.Text)
                        .Col = GridHeader.UTGST_Amt
                        dblUTGSTAmtLine = Val(.Text)
                        .Col = GridHeader.Discount_Percent
                        dblDiscountPercentLine = Val(.Text)
                        .Col = GridHeader.Discount_Amt
                        dblDiscountAmountLine = Val(.Text)
                        .Col = GridHeader.Item_Total
                        dblItemTotalLine = Val(.Text)
                        .Col = GridHeader.Basic_Value
                        dblBasicAmtLine = Val(.Text)
                        .Col = GridHeader.Accessible_Amt
                        dblAssessableAmtLine = Val(.Text)

                        If UCase(lstrItemDelete) <> "D" Then
                            strSalesDtl = Trim(strSalesDtl) & "UPDATE Sales_dtl SET EOP_MODEL='" & strModel & "',Sales_Quantity ='" & Val(CStr(lintItemQuantity)) & "',BinQuantity='" & dblBinQuantity & "',Sales_Tax =" & Trim(lblSaltax_Per.Text) & ","
                            strSalesDtl = Trim(strSalesDtl) & "CustMtrl_Amount= " & Val(CStr(lintItemQuantity * ldblItemCustMtrl)) & ",ToolCost_Amount=" & Val(CStr(ldblTotalToolCost))
                            TempAccessibleVal = CalculateAccessibleValue(lintLoopCounter, ldblNetInsurenceValue, blnISInsExcisable)


                            If blnISExciseRoundOff Then
                                dblExcise_Amount = System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff, dblNetLoadingcharges))
                                strSalesDtl = Trim(strSalesDtl) & ",Excise_Tax=" & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff, dblNetLoadingcharges))
                            Else
                                dblExcise_Amount = System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff, dblNetLoadingcharges), intExciseRoundOffDecimal)
                                strSalesDtl = Trim(strSalesDtl) & ",Excise_Tax=" & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff, dblNetLoadingcharges), intExciseRoundOffDecimal)
                            End If
                            strSalesDtl = Trim(strSalesDtl) & ",Excise_type='" & lstrItemExciseCode & "',SalesTax_type='" & Trim(txtSaleTaxType.Text) & "'"
                            strSalesDtl = Trim(strSalesDtl) & ",CVD_type='" & Trim(lstrItemCVDCode) & "',SAD_type='" & Trim(lstrItemSADCode) & "',Basic_Amount=" & CalculateBasicValue(lintLoopCounter, blnISBasicRoundOff)
                            ' strSalesDtl = Trim(strSalesDtl) & ",Accessible_amount=" & Val(CStr(TempAccessibleVal))
                            strSalesDtl = Trim(strSalesDtl) & ",Accessible_amount=" & Val(dblAssessableAmtLine)
                            If blnISExciseRoundOff Then
                                strSalesDtl = Trim(strSalesDtl) & ",CVD_Amount=" & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_CVD, blnEOUFlag, blnISExciseRoundOff, dblNetLoadingcharges)) & ",SVD_amount=" & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_SAD, blnEOUFlag, blnISExciseRoundOff, dblNetLoadingcharges))
                            Else
                                strSalesDtl = Trim(strSalesDtl) & ",CVD_Amount=" & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_CVD, blnEOUFlag, blnISExciseRoundOff, dblNetLoadingcharges)) & ",SVD_amount=" & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_SAD, blnEOUFlag, blnISExciseRoundOff, dblNetLoadingcharges), intExciseRoundOffDecimal)
                            End If
                            strSalesDtl = Trim(strSalesDtl) & ",Excise_per=" & GetTaxRate(lstrItemExciseCode, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='EXC' and UNIT_CODE='" & gstrUNITID & "'")
                            strSalesDtl = Trim(strSalesDtl) & ",CVD_per=" & GetTaxRate(lstrItemCVDCode, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='CVD' and UNIT_CODE='" & gstrUNITID & "'")
                            strSalesDtl = Trim(strSalesDtl) & ",SVD_per=" & GetTaxRate(lstrItemSADCode, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='SAD' and UNIT_CODE='" & gstrUNITID & "'")
                            strSalesDtl = Trim(strSalesDtl) & ",Tool_Cost =" & ldblItemToolCost & ",From_box = " & ldblItemFromBox & ", To_box = " & ldblItemToBox
                            If blnISExciseRoundOff Then
                                strSalesDtl = Trim(strSalesDtl) & ",TotalExciseAmount =" & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_ALLExcise, blnEOUFlag, blnISExciseRoundOff, dblNetLoadingcharges))
                            Else
                                strSalesDtl = Trim(strSalesDtl) & ",TotalExciseAmount =" & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_ALLExcise, blnEOUFlag, blnISExciseRoundOff, dblNetLoadingcharges), intExciseRoundOffDecimal)
                            End If

                            strSalesDtl = Trim(strSalesDtl) & ",HSNSACCODE='" & strHSNSACCode & "'"
                            strSalesDtl = Trim(strSalesDtl) & ",ISHSNORSAC='" & strHSNSACType & "'"

                            strSalesDtl = Trim(strSalesDtl) & ",CGSTTXRT_TYPE='" & strCGSTType & "'"
                            strSalesDtl = Trim(strSalesDtl) & ",SGSTTXRT_TYPE='" & strSGSTType & "'"
                            strSalesDtl = Trim(strSalesDtl) & ",IGSTTXRT_TYPE='" & strIGSTType & "'"
                            strSalesDtl = Trim(strSalesDtl) & ",UTGSTTXRT_TYPE='" & strUTGSTType & "'"
                            strSalesDtl = Trim(strSalesDtl) & ",COMPENSATION_CESS_TYPE='" & strCCESSType & "'"

                            strSalesDtl = Trim(strSalesDtl) & ",CGST_PERCENT=" & dblCGSTPercentLine
                            strSalesDtl = Trim(strSalesDtl) & ",SGST_PERCENT=" & dblSGSTPercentLine
                            strSalesDtl = Trim(strSalesDtl) & ",IGST_PERCENT=" & dblIGSTPercentLine
                            strSalesDtl = Trim(strSalesDtl) & ",UTGST_PERCENT=" & dblUTGSTPercentLine
                            strSalesDtl = Trim(strSalesDtl) & ",COMPENSATION_CESS_PERCENT=" & dblCCESSPercentLine

                            strSalesDtl = Trim(strSalesDtl) & ",CGST_AMT=" & dblCGSTAmtLine
                            strSalesDtl = Trim(strSalesDtl) & ",SGST_AMT=" & dblSGSTAmtLine
                            strSalesDtl = Trim(strSalesDtl) & ",IGST_AMT=" & dblIGSTAmtLine
                            strSalesDtl = Trim(strSalesDtl) & ",UTGST_AMT=" & dblUTGSTAmtLine
                            strSalesDtl = Trim(strSalesDtl) & ",COMPENSATION_CESS_AMT=" & dblCCESSAmtLine

                            strSalesDtl = Trim(strSalesDtl) & ",Discount_perc=" & dblDiscountPercentLine
                            strSalesDtl = Trim(strSalesDtl) & ",Discount_amt=" & dblDiscountAmountLine
                            strSalesDtl = Trim(strSalesDtl) & ",ITEM_VALUE=" & dblItemTotalLine

                            strSalesDtl = Trim(strSalesDtl) & ",Cust_ref='" & strCustRef & "'"
                            strSalesDtl = Trim(strSalesDtl) & ",Amendment_No='" & StrAmendmentNo & "'"
                            strSalesDtl = Trim(strSalesDtl) & ",SRVDINO='" & strSrvDINo & "'"
                            strSalesDtl = Trim(strSalesDtl) & ",SRVLocation='" & strSRVLocation & "'"
                            strSalesDtl = Trim(strSalesDtl) & ",USLOC='" & strUSLoc & "'"
                            strSalesDtl = Trim(strSalesDtl) & ",SchTime='" & strSchTime & "'"
                            strSalesDtl = Trim(strSalesDtl) & " WHERE UNIT_CODE='" & gstrUNITID & "' and Location_Code ='" & Trim(txtLocationCode.Text) & "'"
                            strSalesDtl = Trim(strSalesDtl) & " and Doc_No =" & Val(txtChallanNo.Text) & " and Cust_Item_Code='"
                            strSalesDtl = Trim(strSalesDtl) & Trim(lstrItemDrgno) & "'" & vbCrLf
                        Else
                            strSalesDtlDelete = Trim(strSalesDtlDelete) & "DELETE Sales_dtl "
                            strSalesDtlDelete = Trim(strSalesDtlDelete) & " WHERE UNIT_CODE='" & gstrUNITID & "' and Location_Code ='" & Trim(txtLocationCode.Text) & "'"
                            strSalesDtlDelete = Trim(strSalesDtlDelete) & " and Doc_No =" & Val(txtChallanNo.Text) & " and Cust_Item_Code='"
                            strSalesDtlDelete = Trim(strSalesDtlDelete) & Trim(lstrItemDrgno) & "'" & vbCrLf
                        End If
                        '10736222
                        strSql = "select dbo.UDF_ISCT2INVOICE( '" & gstrUNITID & "','" & txtCustCode.Text.Trim & "','" & CmbInvType.Text.Trim & "','" & CmbInvSubType.Text.Trim & "','" & txtRefNo.Text.Trim & "')"
                        If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSql)) = True Then
                            blnIsCt2 = True
                            strSqlct2qry = "INSERT INTO TMP_CT2_INVOICE_KNOCKOFF ([UNIT_CODE],[CUST_CODE],[SONO],[AMENDMENT_NO],[TMP_INVOICE_NO],[ITEM_CODE],[CUST_DRG_NO],[CURRENCY_CODE],[QTY],[RATE],[TOOL_COST],[EXCISE_TAX],[EXCISE_AMOUNT],[ECESS_TYPE],[SECESS_TYPE],[IP_ADDRESS]) "
                            strSqlct2qry = strSqlct2qry + " Values('" & gstrUNITID & "','" & txtCustCode.Text.Trim & "','" & txtRefNo.Text.Trim & "','" & txtAmendNo.Text.Trim & "','" & Me.txtChallanNo.Text.Trim & "',"
                            strSqlct2qry = strSqlct2qry + "'" & lstrItemCode.Trim & "','" & lstrItemDrgno.Trim & "','" & lblCurrencyDes.Text.Trim & "'," & Val(CStr(lintItemQuantity)) & "," & Val(CStr(ldblItemRate)) & "," & Val(mdblToolCost(lintLoopCounter - 1)) & ",'" & lstrItemExciseCode.Trim & "'," & dblExcise_Amount & ",'" & txtECSSTaxType.Text.Trim & "','" & txtSECSSTaxType.Text.Trim & "','" & gstrIpaddressWinSck & "' ) "
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
                MsgBox("Unable To Save CT2 Invoice Knock Off Details." & vbCr & objValidateCmd.Parameters(objValidateCmd.Parameters.Count - 1).Value.ToString().Trim, MsgBoxStyle.Information, ResolveResString(100))
                objValidateCmd = Nothing
                SaveData_GST = False
                Exit Function
            End If
            objValidateCmd = Nothing
            '10736222
        End If

        With mP_Connection
            blnActiveTrans = False
            .BeginTrans()
            blnActiveTrans = True
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
                    SaveData_GST = False
                    Exit Function
                End If
                objCmd = Nothing
                '10736222
            End If

            .CommitTrans()
            blnActiveTrans = False
        End With
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        If blnActiveTrans Then mP_Connection.RollbackTrans()
        SaveData_GST = False
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function CalculateBasicValue(ByVal pintRowNo As Short, ByVal blnRoundoff As Boolean) As Double
        '---------------------------------------------------------------------------------------
        'Name       :   CalculateBasicValue
        'Type       :   Function
        'Author     :   Tapan Jain
        'Arguments  :
        'Return     :
        'Purpose    :
        '---------------------------------------------------------------------------------------
        Dim ldblPkg_Per As Double
        Dim ldblRate As Double
        Dim lintQty As Double
        Dim intBasicRoundOffDecimal As Short
        On Error GoTo ErrHandler
        With SpChEntry
            .Row = pintRowNo
            .Col = GridHeader.RatePerUnit
            ldblRate = Val(.Text) / Val(ctlPerValue.Text)
            .Col = GridHeader.Packing
            ldblPkg_Per = Val(.Text)
            .Col = GridHeader.Quantity
            lintQty = Val(.Text)
            .Col = GridHeader.CustSuppMatPerUnit
            intBasicRoundOffDecimal = Val(Find_Value("select basic_roundoff_decimal from sales_parameter where UNIT_CODE='" & gstrUNITID & "'"))
            If blnRoundoff = True Then
                CalculateBasicValue = System.Math.Round((ldblRate + ((ldblPkg_Per * ldblRate) / 100)) * lintQty)
            Else
                CalculateBasicValue = System.Math.Round((ldblRate + ((ldblPkg_Per * ldblRate) / 100)) * lintQty, intBasicRoundOffDecimal)
            End If
        End With
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function CalculateAccessibleValue(ByVal pintRowNo As Short, ByVal pdblInsurenceValue As Double, ByVal pblnISInsAdd As Boolean) As Double
        '---------------------------------------------------------------------------------------
        'Name       :   CalculateAccessibleValue
        'Type       :   Function
        'Author     :   Tapan Jain
        'Arguments  :
        'Return     :
        'Purpose    :
        '---------------------------------------------------------------------------------------
        Dim ldblRate As Double
        Dim ldblCustMat As Double
        Dim ldblToolCost As Double
        Dim ldblPkg_Per As Double
        Dim lintQty As Double
        Dim RSAccessibleVal As ClsResultSetDB
        Dim strSQL As String
        Dim dblMRP As Double
        Dim dblAbatment As Double
        On Error GoTo ErrHandler
        With SpChEntry
            .Row = pintRowNo
            .Col = GridHeader.RatePerUnit
            ldblRate = Val(.Text) / Val(ctlPerValue.Text)
            .Col = GridHeader.Packing
            ldblPkg_Per = Val(.Text)
            .Col = GridHeader.Quantity
            lintQty = Val(.Text)
            .Col = GridHeader.CustSuppMatPerUnit
            ldblCustMat = Val(.Text) / Val(ctlPerValue.Text)
            If UCase(mstrInvoiceType) = "SMP" Then
                .Col = GridHeader.ToolCostPerUnit
                ldblToolCost = Val(.Text) / Val(ctlPerValue.Text)
            Else
                .Col = GridHeader.ToolCost
                ldblToolCost = Val(.Text) / Val(ctlPerValue.Text)
            End If
            If CheckSOType(pintRowNo) = "M" Then
                RSAccessibleVal = New ClsResultSetDB
                .Col = GridHeader.CustRefNo
                strSQL = "select isnull(MRP,0) as MRP,TxRt_Percentage from Cust_Ord_Dtl COH,Gen_TaxRate GT where Account_code = '" & Trim(txtCustCode.Text) & "' and Cust_Ref='" & .Text & "'"
                .Col = GridHeader.AmendmentNo
                strSQL = strSQL & " and COH.Amendment_No='" & .Text & "'"
                .Col = GridHeader.InternalPartNo
                strSQL = strSQL & " and COH.Item_Code = '" & .Text & "'"
                .Col = GridHeader.CustPartNo
                strSQL = strSQL & " and COH.Cust_Drgno = '" & .Text & "' AND COH.Abantment_Code = GT.TxRt_Rate_No and COH.UNIT_CODE=GT.UNIT_CODE AND COH.UNIT_CODE='" & gstrUNITID & "'"
                RSAccessibleVal.GetResult(strSQL, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                If RSAccessibleVal.GetNoRows > 0 Then
                    dblMRP = RSAccessibleVal.GetValue("MRP")
                    dblAbatment = RSAccessibleVal.GetValue("TxRt_Percentage")
                    CalculateAccessibleValue = System.Math.Round((dblMRP * lintQty) - ((dblMRP * lintQty) * dblAbatment / 100), 2)
                End If
                RSAccessibleVal.ResultSetClose()
                RSAccessibleVal = Nothing
            Else


                '
                If pblnISInsAdd = True Then
                    If gblnGSTUnit Then 'added by abhijit
                        CalculateAccessibleValue = System.Math.Round((ldblRate + ldblCustMat + ldblToolCost) * lintQty, 2) 'added by abhijit
                    Else 'added by abhijit
                        CalculateAccessibleValue = System.Math.Round(((ldblRate + ldblCustMat + ldblToolCost + ((ldblPkg_Per * ldblRate) / 100)) * lintQty) + pdblInsurenceValue, 2)
                    End If

                Else
                    CalculateAccessibleValue = System.Math.Round((ldblRate + ldblCustMat + ldblToolCost + ((ldblPkg_Per * ldblRate) / 100)) * lintQty, 2)
                End If

                'If pblnISInsAdd = True Then
                '    CalculateAccessibleValue = System.Math.Round(((ldblRate + ldblCustMat + ldblToolCost + ((ldblPkg_Per * ldblRate) / 100)) * lintQty) + pdblInsurenceValue, 2)
                'Else
                '    CalculateAccessibleValue = System.Math.Round((ldblRate + ldblCustMat + ldblToolCost + ((ldblPkg_Per * ldblRate) / 100)) * lintQty, 2)
                'End If
            End If
        End With
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function CalculateExciseValue(ByVal pintRowNo As Short, ByVal pdblAccessibleValue As Double, ByVal penumTaxType As enumExciseType, ByRef pblnEOU_FLAG As Boolean, ByRef blnExciseFlag As Boolean, ByRef pdblLoadingCharges As Double) As Double
        '---------------------------------------------------------------------------------------
        'Name       :   CalculateExciseValue
        'Type       :   Function
        'Author     :   Tapan Jain
        'Arguments  :
        'Return     :
        'Purpose    :
        '---------------------------------------------------------------------------------------
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
        With SpChEntry
            .Row = pintRowNo
            .Col = GridHeader.EXC
            rsGetTaxRate = New ClsResultSetDB
            strTableSql = "SELECT TxRt_Percentage FROM Gen_TaxRate WHERE TxRt_Rate_No='" & Trim(.Text) & "' AND UNIT_CODE='" & gstrUNITID & "'"
            rsGetTaxRate.GetResult(strTableSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If rsGetTaxRate.GetNoRows > 0 Then
                ldblTaxRate = rsGetTaxRate.GetValue("TxRt_Percentage")
            Else
                ldblTaxRate = 0
            End If
            If pblnEOU_FLAG Then
                .Col = GridHeader.CVD
                strTableSql = "SELECT TxRt_Percentage FROM Gen_TaxRate WHERE TxRt_Rate_No='" & Trim(.Text) & "' AND UNIT_CODE='" & gstrUNITID & "'"
                rsGetTaxRate.GetResult(strTableSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                If rsGetTaxRate.GetNoRows > 0 Then
                    ldblCVDRate = rsGetTaxRate.GetValue("TxRt_Percentage")
                Else
                    ldblCVDRate = 0
                End If
                .Col = GridHeader.SAD
                strTableSql = "SELECT TxRt_Percentage FROM Gen_TaxRate WHERE TxRt_Rate_No='" & Trim(.Text) & "' AND UNIT_CODE='" & gstrUNITID & "'"
                rsGetTaxRate.GetResult(strTableSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                If rsGetTaxRate.GetNoRows > 0 Then
                    ldblSADRate = rsGetTaxRate.GetValue("TxRt_Percentage")
                Else
                    ldblSADRate = 0
                End If
                ldblTempTotalExcise = (((pdblAccessibleValue + pdblLoadingCharges) * ldblTaxRate) / 100)
                If blnExciseFlag = True Then
                    ldblTempTotalExcise = System.Math.Round(ldblTempTotalExcise, 0)
                End If
                ldblTempAllExcise = ldblTempTotalExcise / 2
                ldblTempTotalCVD = (((ldblTempTotalExcise + (pdblAccessibleValue + pdblLoadingCharges)) * ldblCVDRate) / 100)
                If blnExciseFlag = True Then
                    ldblTempTotalCVD = System.Math.Round(ldblTempTotalCVD, 0)
                End If
                ldblTempAllExcise = ldblTempAllExcise + (ldblTempTotalCVD / 2)
                ldblTempTotalSAD = (((ldblTempTotalCVD + ldblTempTotalExcise + (pdblAccessibleValue + pdblLoadingCharges)) * ldblSADRate) / 100)
                If blnExciseFlag = True Then
                    ldblTempTotalSAD = System.Math.Round(ldblTempTotalSAD, 0)
                End If
                ldblTempAllExcise = ldblTempAllExcise + (ldblTempTotalSAD / 2)
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
                    CalculateExciseValue = (((pdblAccessibleValue + pdblLoadingCharges) * ldblTaxRate) / 100)
                ElseIf penumTaxType = enumExciseType.RETURN_CVD Then
                    CalculateExciseValue = 0
                ElseIf penumTaxType = enumExciseType.RETURN_SAD Then
                    CalculateExciseValue = 0
                Else
                    CalculateExciseValue = (((pdblAccessibleValue + pdblLoadingCharges) * ldblTaxRate) / 100)
                End If
            End If
            rsGetTaxRate.ResultSetClose()
            rsGetTaxRate = Nothing
        End With
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function CalculateSalesTaxValue(ByVal pdblTotalBasicValue As Double, ByVal pdblTotalExciseValue As Double, ByRef pblnIncStax As Boolean, ByRef pdblInsurance As Double) As Double
        '---------------------------------------------------------------------------------------
        'Name       :   CalculateSalesTaxValue
        'Type       :   Function
        'Author     :   Tapan Jain
        'Arguments  :
        'Return     :
        'Purpose    :
        '---------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim dbldiscountamount As Double
        '
        If System.Math.Round(Val(txtDiscountAmt.Text)) > 0 Then
            If OptDiscountValue.Checked = True Then
                dbldiscountamount = System.Math.Round(Val(txtDiscountAmt.Text), 0)
            Else
                If chkExciseExumpted.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                    dbldiscountamount = ((pdblTotalBasicValue + pdblTotalExciseValue) * Val(txtDiscountAmt.Text)) / 100
                Else
                    dbldiscountamount = ((pdblTotalBasicValue) * Val(txtDiscountAmt.Text)) / 100
                End If
            End If
        Else
            dbldiscountamount = 0
        End If
        If pblnIncStax = True Then
            If chkExciseExumpted.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                CalculateSalesTaxValue = ((pdblTotalBasicValue + pdblTotalExciseValue + pdblInsurance - dbldiscountamount) * Val(lblSaltax_Per.Text)) / 100
            Else
                CalculateSalesTaxValue = ((pdblTotalBasicValue + pdblInsurance - dbldiscountamount) * Val(lblSaltax_Per.Text)) / 100
            End If
        Else
            If chkExciseExumpted.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                CalculateSalesTaxValue = ((pdblTotalBasicValue + pdblTotalExciseValue - dbldiscountamount) * Val(lblSaltax_Per.Text)) / 100
            Else
                CalculateSalesTaxValue = ((pdblTotalBasicValue - dbldiscountamount) * Val(lblSaltax_Per.Text)) / 100
            End If
        End If
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function CalculateECSSTaxValue(ByVal pdblTotalExciseValue As Double) As Double
        '---------------------------------------------------------------------------------------
        'Name       :   CalculateECSSTaxValue
        'Type       :   Function
        'Author     :   Arshad Ali
        'Return     :   calculated ECSS Amount
        'Purpose    :
        '---------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        CalculateECSSTaxValue = (pdblTotalExciseValue * Val(lblECSStax_Per.Text) / 100)
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function CalculateSurchargeTaxValue(ByVal pdblTotalCSTValue As Double) As Double
        '---------------------------------------------------------------------------------------
        'Name       :   CalculateSurchargeTaxValue
        'Type       :   Function
        'Author     :   Tapan Jain
        'Arguments  :
        'Return     :
        'Purpose    :
        '---------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        CalculateSurchargeTaxValue = (pdblTotalCSTValue * Val(lblSurcharge_Per.Text) / 100)
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function PrepareQueryForShowingExcise(ByVal pblnTarrifCodeReq As Boolean, ByRef pstrItemCode As String) As String
        Dim strSQL As String
        Dim lclsGetTariffCode As ClsResultSetDB
        PrepareQueryForShowingExcise = ""
        If pblnTarrifCodeReq = True Then
            strSQL = "SELECT Tariff_code FROM Item_Mst WHERE UNIT_CODE='" & gstrUNITID & "' AND Item_Code ='" & pstrItemCode & "'"
            lclsGetTariffCode = New ClsResultSetDB
            Call lclsGetTariffCode.GetResult(strSQL, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If lclsGetTariffCode.GetNoRows > 0 Then
                strSQL = "SELECT Excise_duty FROM Tax_Tariff_Mst WHERE UNIT_CODE='" & gstrUNITID & "' AND Tariff_SubHead='" & lclsGetTariffCode.GetValue("Tariff_code") & "'"
                Call lclsGetTariffCode.GetResult(strSQL, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                If lclsGetTariffCode.GetNoRows > 0 Then
                    PrepareQueryForShowingExcise = " AND TxRt_Rate_No='" & lclsGetTariffCode.GetValue("Excise_duty") & "'"
                End If
            End If
            lclsGetTariffCode.ResultSetClose()
            lclsGetTariffCode = Nothing
        Else
            PrepareQueryForShowingExcise = ""
        End If
    End Function
    Private Function GetBOMCheckFlagValue(ByVal pstrFieldName As String) As Boolean
        '---------------------------------------------------------------------------------------
        'Name       :   GetBOMCheckFlagValue
        'Type       :   Function
        'Author     :   Tapan Jain
        'Arguments  :
        'Return     :
        'Purpose    :
        '---------------------------------------------------------------------------------------
        Dim strSQL As String
        Dim rsObj As New ADODB.Recordset
        On Error GoTo ErrHandler
        strSQL = ""
        strSQL = "SELECT " & pstrFieldName & " FROM Sales_Parameter WHERE UNIT_CODE='" & gstrUNITID & "'"
        If rsObj.State = 1 Then rsObj.Close()
        rsObj.Open(strSQL, mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
        If rsObj.EOF Or rsObj.BOF Then
            MsgBox("No Data define in Sales_Parameter Table", MsgBoxStyle.Critical, "eMPro")
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
    Private Function GetTotalDispatchQuantityFromDailySchedule(ByVal pstrAccountCode As String, ByVal pstrCustomerDrawNo As String, ByVal pstrItemCode As String, ByVal pstrDate As String, ByVal pstrMode As String, ByVal pdblPrevQty As Double, Optional ByRef pstrSRVDINo As String = "") As Double
        '---------------------------------------------------------------------------------------
        'Name       :   GetTotalDispatchQuantityFromDailySchedule
        'Type       :   Function
        'Author     :   Tapan Jain
        'Arguments  :
        'Return     :
        'Purpose    :
        '---------------------------------------------------------------------------------------
        'Revision  By       : Ashutosh Verma
        'Revision On        : 19-01-2006
        'History            : Bug fix - After cancellation user can't recreate the invoice,issue Id:16907.
        '=======================================================================================
        'Revision  By       : Ashutosh Verma
        'Revision On        : 09-03-2006 ,issue id :17229.
        'History            : Calculate dispatches from Printedsrv & 57F4 challan at the time of invoice saveing.
        '=======================================================================================
        Dim strScheduleSql As String
        Dim objRsForSchedule As New ADODB.Recordset
        Dim ldblTotalDispatchQuantity As Double
        Dim ldblTotalScheduleQuantity As Double
        Dim lintLoopCounter As Short
        On Error GoTo ErrHandler
        ldblTotalDispatchQuantity = 0
        ldblTotalScheduleQuantity = 0
        mP_Connection.Execute("SET DATEFORMAT 'mdy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        If pstrMode = "ADD" Then
            If Len(Trim(pstrSRVDINo)) > 0 Then
                strScheduleSql = " select isnull(c.quantity,0) as schedule_quantity "
                strScheduleSql = strScheduleSql & " from mkt_enagareDtl c "
                strScheduleSql = strScheduleSql & " where c.UNIT_CODE='" & gstrUNITID & "' AND c.kanbanNo ='" & Trim(pstrSRVDINo) & "'"
            Else
                strScheduleSql = "Select Schedule_Quantity,Despatch_Qty from DailyMktSchedule where UNIT_CODE='" & gstrUNITID & "' AND Account_Code='" & pstrAccountCode & "' and "
                strScheduleSql = strScheduleSql & " datepart(yyyy,Trans_Date)='" & Year(ConvertToDate(pstrDate)) & "'"
                strScheduleSql = strScheduleSql & " and datepart(mm,Trans_Date)='" & Month(ConvertToDate(pstrDate)) & "'"
                strScheduleSql = strScheduleSql & " and Trans_Date <='" & getDateForDB(pstrDate) & "'"
                strScheduleSql = strScheduleSql & " and Cust_DrgNo = '" & pstrCustomerDrawNo & "' AND ITEM_CODE  = '" & pstrItemCode & "' and Status =1  ORDER BY Trans_Date DESC" '''and Schedule_Flag =1   ( Now Not Consider)
            End If
            If objRsForSchedule.State = 1 Then objRsForSchedule.Close()
            objRsForSchedule.CursorLocation = ADODB.CursorLocationEnum.adUseClient
            objRsForSchedule.Open(strScheduleSql, mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
            If objRsForSchedule.EOF Or objRsForSchedule.BOF Then
                GetTotalDispatchQuantityFromDailySchedule = -1
                mP_Connection.Execute("SET DATEFORMAT 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                objRsForSchedule.Close()
                Exit Function
            Else
                ldblTotalScheduleQuantity = Val(objRsForSchedule.Fields("Schedule_Quantity").Value)
                ldblTotalDispatchQuantity = GetTotalDispatchForKanban(pstrSRVDINo, pstrMode)
                GetTotalDispatchQuantityFromDailySchedule = Val(CStr(ldblTotalScheduleQuantity - ldblTotalDispatchQuantity))
                mP_Connection.Execute("SET DATEFORMAT 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                objRsForSchedule.Close()
                Exit Function
            End If
        Else
            If Len(Trim(pstrSRVDINo)) > 0 Then
                strScheduleSql = " select isnull(c.quantity,0) as schedule_quantity "
                strScheduleSql = strScheduleSql & " from mkt_enagareDtl c "
                strScheduleSql = strScheduleSql & " where c.UNIT_CODE='" & gstrUNITID & "' AND c.kanbanNo ='" & Trim(pstrSRVDINo) & "'"
            Else
                strScheduleSql = "Select Schedule_Quantity,Despatch_Qty from DailyMktSchedule where UNIT_CODE='" & gstrUNITID & "' AND Account_Code='" & pstrAccountCode & "' and "
                strScheduleSql = strScheduleSql & " datepart(yyyy,Trans_Date)='" & Year(ConvertToDate(pstrDate)) & "'"
                strScheduleSql = strScheduleSql & " and datepart(mm,Trans_Date)='" & Month(ConvertToDate(pstrDate)) & "'"
                strScheduleSql = strScheduleSql & " and Trans_Date <='" & getDateForDB(pstrDate) & "'"
                strScheduleSql = strScheduleSql & " and Cust_DrgNo = '" & pstrCustomerDrawNo & "' AND ITEM_CODE  = '" & pstrItemCode & "' and Status =1  ORDER BY Trans_Date DESC" '''and Schedule_Flag =1   ( Now Not Consider)
            End If
            If objRsForSchedule.State = 1 Then objRsForSchedule.Close()
            objRsForSchedule.CursorLocation = ADODB.CursorLocationEnum.adUseClient
            objRsForSchedule.Open(strScheduleSql, mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
            If objRsForSchedule.EOF Or objRsForSchedule.BOF Then
                GetTotalDispatchQuantityFromDailySchedule = -1
                mP_Connection.Execute("SET DATEFORMAT 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                objRsForSchedule.Close()
                Exit Function
            Else
                ldblTotalScheduleQuantity = Val(objRsForSchedule.Fields("Schedule_Quantity").Value)
                ldblTotalDispatchQuantity = GetTotalDispatchForKanban(pstrSRVDINo, pstrMode)
                GetTotalDispatchQuantityFromDailySchedule = Val(CStr(ldblTotalScheduleQuantity - ldblTotalDispatchQuantity)) ''+ val(pdblPrevQty)
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
        '---------------------------------------------------------------------------------------
        'Name       :   GetTotalDispatchQuantityFromDailySchedule
        'Type       :   Function
        'Author     :   Tapan Jain
        'Arguments  :
        'Return     :
        'Purpose    :
        '---------------------------------------------------------------------------------------
        Dim strScheduleSql As String
        Dim objRsForSchedule As New ADODB.Recordset
        Dim ldblTotalDispatchQuantity As Double
        Dim ldblTotalScheduleQuantity As Double
        Dim lintLoopCounter As Short
        Dim strMakeDate As String
        On Error GoTo ErrHandler
        ldblTotalDispatchQuantity = 0
        ldblTotalScheduleQuantity = 0
        mP_Connection.Execute("SET DATEFORMAT 'mdy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        If Val(CStr(Month(ConvertToDate(pstrDate)))) < 10 Then
            strMakeDate = Year(ConvertToDate(pstrDate)) & "0" & Month(ConvertToDate(pstrDate))
        Else
            strMakeDate = Year(ConvertToDate(pstrDate)) & Month(ConvertToDate(pstrDate))
        End If
        If pstrMode = "ADD" Then
            If objRsForSchedule.State = 1 Then objRsForSchedule.Close()
            objRsForSchedule.CursorLocation = ADODB.CursorLocationEnum.adUseClient
            strScheduleSql = "Select Schedule_Qty,isnull(Despatch_Qty,0) AS Despatch_qty  from MonthlyMktSchedule where UNIT_CODE='" & gstrUNITID & "' AND Account_Code='" & pstrAccountCode & "' and "
            strScheduleSql = strScheduleSql & " Year_Month=" & Val(Trim(strMakeDate)) & ""
            strScheduleSql = strScheduleSql & " and Cust_DrgNo = '" & pstrCustomerDrawNo & "' AND Item_code = '" & pstrItemCode & "' and status =1 "
            objRsForSchedule.Open(strScheduleSql, mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
            If objRsForSchedule.EOF Or objRsForSchedule.BOF Then
                GetTotalDispatchQuantityFromMonthlySchedule = -1
                mP_Connection.Execute("SET DATEFORMAT 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                objRsForSchedule.Close()
                Exit Function
            Else
                objRsForSchedule.MoveFirst()
                For lintLoopCounter = 1 To objRsForSchedule.RecordCount
                    ldblTotalScheduleQuantity = ldblTotalScheduleQuantity + Val(objRsForSchedule.Fields("Schedule_Qty").Value)
                    ldblTotalDispatchQuantity = ldblTotalDispatchQuantity + Val(objRsForSchedule.Fields("Despatch_Qty").Value)
                    objRsForSchedule.MoveNext()
                Next
                GetTotalDispatchQuantityFromMonthlySchedule = Val(CStr(ldblTotalScheduleQuantity - ldblTotalDispatchQuantity))
                mP_Connection.Execute("SET DATEFORMAT 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                objRsForSchedule.Close()
                Exit Function
            End If
        Else
            strScheduleSql = "Select Schedule_Qty,isnull(Despatch_Qty,0)AS Despatch_qty from MonthlyMktSchedule where UNIT_CODE='" & gstrUNITID & "' AND Account_Code='" & pstrAccountCode & "' and "
            strScheduleSql = strScheduleSql & " Year_Month=" & Val(Trim(strMakeDate)) & ""
            strScheduleSql = strScheduleSql & " and Cust_DrgNo = '" & pstrCustomerDrawNo & "' AND Item_code = '" & pstrItemCode & "' and status =1 "
            If objRsForSchedule.State = 1 Then objRsForSchedule.Close()
            objRsForSchedule.CursorLocation = ADODB.CursorLocationEnum.adUseClient
            objRsForSchedule.Open(strScheduleSql, mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
            If objRsForSchedule.EOF Or objRsForSchedule.BOF Then
                GetTotalDispatchQuantityFromMonthlySchedule = -1
                mP_Connection.Execute("SET DATEFORMAT 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                objRsForSchedule.Close()
                Exit Function
            Else
                objRsForSchedule.MoveFirst()
                For lintLoopCounter = 1 To objRsForSchedule.RecordCount
                    ldblTotalScheduleQuantity = ldblTotalScheduleQuantity + Val(objRsForSchedule.Fields("Schedule_Qty").Value)
                    ldblTotalDispatchQuantity = ldblTotalDispatchQuantity + Val(objRsForSchedule.Fields("Despatch_Qty").Value)
                    objRsForSchedule.MoveNext()
                Next
                GetTotalDispatchQuantityFromMonthlySchedule = Val(CStr(ldblTotalScheduleQuantity - ldblTotalDispatchQuantity)) + Val(CStr(pdblPrevQty))
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
    Public Function CheckcustorddtlQty(ByRef pstrMode As String, ByRef pstrItemCode As String, ByRef pstrDrgno As String, ByRef pdblQty As Double, Optional ByRef pstrCustRef As String = "", Optional ByRef pstrAmendment As String = "") As Boolean
        Dim rsCustOrdDtl As ClsResultSetDB
        Dim rssaledtl As ClsResultSetDB
        Dim dblSaleQuantity As Double
        Dim strCustOrdDtl As String
        On Error GoTo ErrHandler
        rsCustOrdDtl = New ClsResultSetDB
        If blnInvoiceAgainstMultipleSO Then
            strCustOrdDtl = "Select openso,balance_Qty = order_qty - Despatch_qty from Cust_ord_dtl where "
            strCustOrdDtl = strCustOrdDtl & "Account_code ='" & txtCustCode.Text & "'" & " and Item_code ='"
            strCustOrdDtl = strCustOrdDtl & pstrItemCode & "' and cust_drgNo ='" & pstrDrgno
            strCustOrdDtl = strCustOrdDtl & "' and Authorized_flag = 1 and cust_ref = '" & pstrCustRef
            strCustOrdDtl = strCustOrdDtl & "' and Amendment_no = '" & pstrAmendment & "' AND UNIT_CODE='" & gstrUNITID & "'"
        Else
            strCustOrdDtl = "Select openso,balance_Qty = order_qty - Despatch_qty from Cust_ord_dtl where "
            strCustOrdDtl = strCustOrdDtl & "Account_code ='" & txtCustCode.Text & "'" & " and Item_code ='"
            strCustOrdDtl = strCustOrdDtl & pstrItemCode & "' and cust_drgNo ='" & pstrDrgno
            strCustOrdDtl = strCustOrdDtl & "' and Authorized_flag = 1 and cust_ref = '" & txtRefNo.Text
            strCustOrdDtl = strCustOrdDtl & "' and Amendment_no = '" & txtAmendNo.Text & "' AND UNIT_CODE='" & gstrUNITID & "'"
        End If
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
                    rssaledtl.GetResult("Select Sales_Quantity from Sales_Dtl where UNIT_CODE='" & gstrUNITID & "' AND doc_no = " & txtChallanNo.Text & " and item_code = '" & pstrItemCode & "' and cust_ITem_code = '" & pstrDrgno & "'")
                    dblSaleQuantity = rssaledtl.GetValue("Sales_Quantity")
                    If (rsCustOrdDtl.GetValue("Balance_Qty")) < pdblQty Then
                        MsgBox("Balance Quantity available in SO for Customer Part code [ " & pstrDrgno & "] is " & rsCustOrdDtl.GetValue("Balance_Qty") & ".", MsgBoxStyle.Information, "eMPro")
                        CheckcustorddtlQty = False
                    Else
                        CheckcustorddtlQty = True
                    End If
            End Select
        End If
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Public Function CalculateLoadingchargesAmount(ByRef pdblaccessible As Double, ByRef pdblLoadingTax_per As Double) As Double
        Dim dblLoadingCharges As Double
        Dim rsRoundData As New ClsResultSetDB
        Dim intRound As Short
        rsRoundData.GetResult("Select TX_ROUNDOFFPLACE from gen_TaxHeadMaster where Tx_taxeID = 'LDT' AND UNIT_CODE='" & gstrUNITID & "'")
        intRound = 0
        If rsRoundData.GetNoRows > 0 Then
            intRound = rsRoundData.GetValue("TX_ROUNDOFFPLACE")
        End If
        dblLoadingCharges = System.Math.Round((pdblaccessible * pdblLoadingTax_per) / 100, intRound)
        CalculateLoadingchargesAmount = dblLoadingCharges
    End Function
    Public Function ReturnNoOfDecimals(ByRef pstrItemCode As String) As Short
        '***************************************************************************************
        'Name       :   ReturnNoOfDecimals
        'Type       :   Function
        'Author     :   Nisha Rai
        'Arguments  :   Item Code
        'Return     :   No of Decimals as intiger
        'Purpose    :   Fetches measure code of given item code and then according to decimal
        '               allowed fkag it returns No decimals allowed
        '***************************************************************************************
        Dim rsMeasurementUnit As ClsResultSetDB
        Dim rsNoOfDecimal As ClsResultSetDB
        Dim strMeasurementUnit As String
        Dim intNoofDecimals As Short
        On Error GoTo ErrHandler
        rsMeasurementUnit = New ClsResultSetDB
        rsMeasurementUnit.GetResult("Select Cons_Measure_Code from Item_Mst where UNIT_CODE='" & gstrUNITID & "' AND item_code = '" & pstrItemCode & "'")
        If rsMeasurementUnit.GetNoRows > 0 Then
            rsMeasurementUnit.MoveFirst()
            strMeasurementUnit = rsMeasurementUnit.GetValue("Cons_Measure_Code")
            rsNoOfDecimal = New ClsResultSetDB
            rsNoOfDecimal.GetResult("select Decimal_Allowed_Flag,NoOFDecimal from Measure_Mst where UNIT_CODE='" & gstrUNITID & "' AND Measure_Code = '" & strMeasurementUnit & "'")
            If rsNoOfDecimal.GetNoRows > 0 Then
                rsNoOfDecimal.MoveFirst()
                If rsNoOfDecimal.GetValue("Decimal_Allowed_Flag") = True Then
                    intNoofDecimals = Val(rsNoOfDecimal.GetValue("NoOFDecimal"))
                    If intNoofDecimals = 0 Then
                        intNoofDecimals = 2
                    End If
                    ReturnNoOfDecimals = intNoofDecimals
                Else
                    ReturnNoOfDecimals = 0
                End If
            End If
        End If
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub txtSRVDINO_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSRVDINO.TextChanged
        If Len(Trim(txtSRVDINO.Text)) = 0 Then
            txtSchTime.Text = "" : txtSRVLocation.Text = "" : txtUsLoc.Text = ""
        End If
    End Sub
    Private Sub txtSRVLocation_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSRVLocation.TextChanged
        If Len(Trim(txtSRVLocation.Text)) = 0 Then
            txtSRVDINO.Text = "" : txtSchTime.Text = "" : txtUsLoc.Text = ""
        End If
    End Sub
    Private Sub txtUsLoc_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUsLoc.TextChanged
        If Len(Trim(txtUsLoc.Text)) = 0 Then
            txtSRVDINO.Text = "" : txtSRVLocation.Text = "" : txtSchTime.Text = ""
        End If
    End Sub
    Private Sub txtUsLoc_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtUsLoc.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '****************************************************
        'Created By     -  Nisha
        'Description    -  At Enter Key Press Set Focus To Next Control
        '****************************************************
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
    Private Sub txtSchTime_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSchTime.TextChanged
        If Len(Trim(txtSchTime.Text)) = 0 Then
            'txtSRVDINO.Text = "": txtSRVLocation.Text = "": txtUsLoc.Text = ""
        End If
    End Sub
    Private Sub txtSchTime_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtSchTime.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '****************************************************
        'Created By     -  Nisha
        'Description    -  At Enter Key Press Set Focus To Next Control
        '****************************************************
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
    Private Sub cmdhelpSRVDI_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdhelpSRVDI.Click
        '****************************************************
        'Created By     -  Nisha Rai
        'Description    -  To fetch the help on MktNagare Details.
        '****************************************************
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
                Call .GetText(GridHeader.delete, intLoopCounter, VarDelete)
                If UCase(Trim(VarDelete)) <> "D" Then
                    intmaxitems = intmaxitems + 1
                End If
            Next
            intMaxLoop = .MaxRows
            For intLoopCounter = 1 To intMaxLoop
                VarDelete = Nothing
                Call .GetText(GridHeader.delete, intLoopCounter, VarDelete)
                If UCase(Trim(VarDelete)) <> "D" Then
                    varItemCode = Nothing
                    varDrgNo = Nothing
                    Call .GetText(GridHeader.InternalPartNo, intLoopCounter, varItemCode)
                    Call .GetText(GridHeader.CustPartNo, intLoopCounter, varDrgNo)
                    Exit For
                End If
            Next
        End With
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                If CBool(Find_Value("SELECT showItemInEnagareHelp FROM SALES_PARAMETER WHERE UNIT_CODE='" & gstrUNITID & "'")) Then
                    StrHelpSql = "select m.cust_drgNo, m.item_code, KanbanNo, UNLOC,USLOC, case when Sch_Time = '23:59' then '' else Sch_Time end as sch_time,Sch_Date,Quantity, Cust_Ref from vw_Enagaredtl_Help m "
                Else
                    StrHelpSql = "select Cust_drgNo, KanbanNo, UNLOC,USLOC, case when Sch_Time = '23:59' then '' else Sch_Time end as sch_time,Sch_Date,Quantity from MKT_Enagaredtl m "
                End If
                StrHelpSql = StrHelpSql & " where m.UNIT_CODE='" & gstrUNITID & "' AND  m.quantity > ((select isnull(sum(b.sales_quantity),0) from salesChallan_dtl a inner join sales_dtl b on a.UNIT_CODE=b.UNIT_CODE AND a.location_code = b.location_code and a.doc_no=b.doc_no where m.kanbanNo = b.srvdino and a.cancel_flag <> 1 AND a.UNIT_CODE='" & gstrUNITID & "')" & " + (select IsNull(sum(sales_quantity),0) as sales_quantity " & " from printedsrv_dtl p where p.KanBan_No=m.KanBanNo AND p.UNIT_CODE=m.UNIT_CODE AND p.UNIT_CODE='" & gstrUNITID & "') +(Select isnull(Sum(b.quantity),0) as sales_quantity From mkt_57F4challankanban_dtl B inner join mkt_57F4challan_hdr A on b.UNIT_CODE=a.UNIT_CODE AND B.doc_type=A.doc_type and B.doc_no = A.doc_no where A.cancel_flag = 0 and B.Kanban_no=m.KanBanNo AND B.UNIT_CODE=m.UNIT_CODE AND B.UNIT_CODE='" & gstrUNITID & "')) "
                If Len(txtCustCode.Text) > 0 Then
                    StrHelpSql = StrHelpSql & " and m.Account_code = '" & Trim(txtCustCode.Text) & "'"
                End If
                StrHelpSql = StrHelpSql & " order by sch_date desc, Sch_time asc"
                strMktNagare = Me.ctlEMPHelpInvoiceEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, StrHelpSql, "eNagare Details")
                If UBound(strMktNagare) < 0 Then Exit Sub
                If strMktNagare(0) = "0" Then
                    MsgBox("No Record Available to Display", MsgBoxStyle.Information, "eMPro") : txtSRVDINO.Text = "" : txtSRVDINO.Focus() : Exit Sub
                Else
                    If CBool(Find_Value("SELECT showItemInEnagareHelp FROM SALES_PARAMETER WHERE UNIT_CODE='" & gstrUNITID & "'")) Then
                        txtUsLoc.Text = IIf(IsDBNull(strMktNagare(4)), "", strMktNagare(4))
                        txtSRVDINO.Text = IIf(IsDBNull(strMktNagare(2)), "", strMktNagare(2))
                        txtSRVLocation.Text = IIf(IsDBNull(strMktNagare(3)), "", strMktNagare(3))
                        txtSchTime.Text = IIf(IsDBNull(strMktNagare(5)), "", strMktNagare(5))
                        mstrSONo = IIf(IsDBNull(strMktNagare(8)), "", strMktNagare(8))
                        Call FillDetails(True, IIf(IsDBNull(strMktNagare(1)), "", strMktNagare(1)))
                    Else
                        txtUsLoc.Text = IIf(IsDBNull(strMktNagare(3)), "", strMktNagare(3))
                        txtSRVDINO.Text = IIf(IsDBNull(strMktNagare(1)), "", strMktNagare(1))
                        txtSRVLocation.Text = IIf(IsDBNull(strMktNagare(2)), "", strMktNagare(2))
                        txtSchTime.Text = IIf(IsDBNull(strMktNagare(4)), "", strMktNagare(4))
                        Call FillDetails(False)
                    End If
                End If
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Public Function CalculateTCSTax(ByRef pdblTotalValue As Double, ByRef pblnTCSRoundOFF As Boolean, ByRef pintTCSPer As Double) As Double
        Dim dblTCSTax As Double
        If pblnTCSRoundOFF = True Then
            dblTCSTax = System.Math.Round((pdblTotalValue * pintTCSPer) / 100, 0)
        Else
            dblTCSTax = System.Math.Round((pdblTotalValue * pintTCSPer) / 100, 2)
        End If
        CalculateTCSTax = dblTCSTax
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
    Private Sub cmdPrint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdPrint.Click
        On Error GoTo ErrHandler
        Dim intCount As Short
        Dim varTemp As Object
        Dim strFileName As String
        Dim dblWaitingTime As Double
        'Kill(gstrLocalCDrive & "TypeToPrn.bat")
        Kill(gstrLocalCDrive & "EmproInv\TypeToPrn.bat")
        If Len(objInvoicePrint.FileName) > 0 Then
            strFileName = objInvoicePrint.FileName
        End If
        If intNoCopies = 0 Then intNoCopies = 1
        dblWaitingTime = Val(Find_Value("select waitingTime from sales_parameter WHERE UNIT_CODE='" & gstrUNITID & "'"))
        If dblWaitingTime = 0 Then
            dblWaitingTime = 5000
        End If
TypeFileNotFoundCreateRetry:
        For intCount = 1 To intNoCopies
            varTemp = Shell("cmd.exe /c " & gstrLocalCDrive & "EmproInv\TypeToPrn.bat " & strFileName, AppWinStyle.Hide)
            Sleep(dblWaitingTime)
            Call printBarCode(objInvoicePrint.BCFileName)
            Sleep(dblWaitingTime)
            varTemp = Shell("cmd.exe /c " & gstrLocalCDrive & "EmproInv\TypeToPrn.bat " & gstrLocalCDrive & "EmproInv\PageFeed.txt", AppWinStyle.Hide)
        Next
        Exit Sub
ErrHandler:
        If Err.Number = 53 Then
            FileOpen(1, gstrLocalCDrive & "EmproInv\TypeToPrn.bat", OpenMode.Append)
            PrintLine(1, "Type %1> prn") '& Printer.Port
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
        '----------------------------------------------------------------------------
        'Author         :   Arshad Ali
        'Argument       :   Non
        'Return Value   :   Non
        'Function       :   Removes all special characters used for formating from text file
        'Comments       :   Nil
        '----------------------------------------------------------------------------
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
        'Author         :   Arshad Ali
        'Argument       :
        'Return Value   :
        'Function       :
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        Dim varTemp As Object
        Dim strString As String
        'strString = gstrLocalCDrive & "pdf-dot.bat " & gstrLocalCDrive & "BarCode.txt 4 2 2 1"
        strString = gstrLocalCDrive & "EmproInv\pdf-dot.bat " & gstrLocalCDrive & "EmproInv\BarCode.txt 4 2 2 1"
        'strString = gstrLocalCDrive & "pdf-dot.bat " & pstrFileName & " 4 2 2 1"
        strString = gstrLocalCDrive & "EmproInv\pdf-dot.bat " & pstrFileName & " 4 2 2 1"
        varTemp = Shell("cmd.exe /c " & strString)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Sub PrintingInvoice()
        On Error GoTo ErrHandler
        If Val(txtChallanNo.Text) < 99000000 Then 'Check For Temporary Challan No.
            objInvoicePrint.ConnectionString = gstrCONNECTIONSTRING
            objInvoicePrint.Connection()
            'objInvoicePrint.FileName = gstrLocalCDrive & "InvoicePrint.txt"
            objInvoicePrint.FileName = gstrLocalCDrive & "EmproInv\InvoicePrint.txt"
            'objInvoicePrint.BCFileName = gstrLocalCDrive & "BarCode.txt"
            objInvoicePrint.BCFileName = gstrLocalCDrive & "EmproInv\BarCode.txt"
            objInvoicePrint.CompanyName = gstrCOMPANY
            objInvoicePrint.Address1 = gstr_RGN_ADDRESS1
            objInvoicePrint.Address2 = gstr_RGN_ADDRESS2
            If chkDTRemoval.CheckState = System.Windows.Forms.CheckState.Checked Then
                objInvoicePrint.Print_Invoice(gstrUNITID, True, (txtLocationCode.Text), (txtChallanNo.Text), dtpRemoval.Value & " " & VB6.Format(dtpRemovalTime.Value.Hour, "00") & ":" & VB6.Format(dtpRemovalTime.Value.Minute, "00"))
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
        On Error GoTo Err_Handler
        rssaledtl = New ClsResultSetDB
        rsItembal = New ClsResultSetDB
        rssaledtl = New ClsResultSetDB
        rsCompany = New ClsResultSetDB
        SALEDTL = "select * from Saleschallan_Dtl where Doc_No =" & txtChallanNo.Text & "  and Location_Code='" & Trim(txtLocationCode.Text) & "' and UNIT_CODE='" & gstrUNITID & "'"
        rsSalesChallan = New ClsResultSetDB
        rssaledtl.GetResult(SALEDTL, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        strAccountCode = rssaledtl.GetValue("Account_code")
        strCustRef = rssaledtl.GetValue("Cust_ref")
        StrAmendmentNo = rssaledtl.GetValue("Amendment_No")
        strInvoiceDate = VB6.Format(rssaledtl.GetValue("Invoice_Date"), gstrDateFormat)
        strSalesconf = "Select UpdatePO_Flag,UpdateStock_Flag,Stock_Location,OpenningBal,Preprinted_Flag,NoCopies from saleconf where "
        strSalesconf = strSalesconf & "Invoice_type = '" & mstrInvoiceType & "' and sub_type = '"
        strSalesconf = strSalesconf & mstrInvoiceSubType & "' and Location_Code='" & Trim(txtLocationCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0 and UNIT_CODE='" & gstrUNITID & "'"
        rsSalesConf = New ClsResultSetDB
        rsSalesConf.GetResult(strSalesconf, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        updatePOflag = rsSalesConf.GetValue("UpdatePO_Flag")
        updatestockflag = rsSalesConf.GetValue("UpdateStock_Flag")
        strStockLocation = rsSalesConf.GetValue("Stock_Location")
        mOpeeningBalance = Val(rsSalesConf.GetValue("OpenningBal"))
        mIntNoCopies = rsSalesConf.GetValue("NoCopies")
        If Len(Trim(strStockLocation)) = 0 Then
            MsgBox("Please Define Stock Location in Sales Configuration. ")
            Exit Sub
        End If
        '***********To check if Tool Cost Deduction will be done or Not on 16/02/2004
        rsSalesParameter.GetResult("Select CheckToolAmortisation from Sales_Parameter where UNIT_CODE='" & gstrUNITID & "'")
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
        '*************
        SALEDTL = "Select Sales_Quantity,Item_code,Cust_Item_Code,toolcost_amount from sales_Dtl where Doc_No = " & txtChallanNo.Text & " and Location_Code='" & Trim(txtLocationCode.Text) & "' and UNIT_CODE='" & gstrUNITID & "'"
        rssaledtl.GetResult(SALEDTL, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        intRow = rssaledtl.GetNoRows
        rssaledtl.MoveFirst()
        '******Check for balance & despatch in Cust_ord_dtl
        For intLoopCount = 1 To intRow
            ItemCode = rssaledtl.GetValue("Item_code")
            salesQuantity = rssaledtl.GetValue("Sales_quantity")
            strDrgNo = rssaledtl.GetValue("Cust_Item_code")
            dblToolCost = rssaledtl.GetValue("ToolCost_amount")
            rsItembal.GetResult("Select Cur_bal from Itembal_Mst where Item_code = '" & ItemCode & "'and Location_code ='" & strStockLocation & "' and UNIT_CODE='" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            If rsItembal.GetNoRows > 0 Then
                If salesQuantity > rsItembal.GetValue("Cur_Bal") Then
                    MsgBox("Balance for item " & ItemCode & " at Location " & strStockLocation & " not available. ", MsgBoxStyle.Information, "eMPro")
                    Exit Sub
                End If
            Else
                MsgBox("No Item in ItemMaster for Location " & strStockLocation & ".", MsgBoxStyle.OkOnly, "eMPro")
                rsSalesConf.ResultSetClose()
                rsSalesConf = Nothing
                Exit Sub
            End If
            If Len(Trim(strCustRef)) > 0 Then
                If UCase(mstrInvoiceType) <> "REJ" Then
                    rsItembal.GetResult("Select balanceQty = order_qty - despatch_Qty,OpenSO from Cust_ord_dtl where account_code ='" & strAccountCode & "' and Cust_ref ='" & strCustRef & "' and Amendment_No = '" & StrAmendmentNo & "' and Item_code ='" & ItemCode & "' and Cust_drgNo ='" & strDrgNo & "' and Active_flag ='A' and UNIT_CODE='" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                    If rsItembal.GetNoRows > 0 Then
                        'Changed by nisha for Open So Check
                        If rsItembal.GetValue("OpenSO") = False Then
                            If salesQuantity > rsItembal.GetValue("BalanceQty") Then
                                MsgBox("Balance Quantity in SO for item " & ItemCode & " is " & rsItembal.GetValue("BalanceQty") & ".Check Quantity of Item in Challan.", MsgBoxStyle.Information, "eMPro")
                                Exit Sub
                            End If
                        End If
                    Else
                        MsgBox("No Item (" & strItemCode & ") exist in SO - " & strCustRef & ".", MsgBoxStyle.Information, "eMPro")
                        Exit Sub
                    End If
                End If
            End If
            '************To Check for Tool Cost
            If blnCheckToolCost = True Then
                If dblToolCost > 0 Then
                    strItembal = "select BalanceQty = isnull(proj_qty,0) - isnull(UsedProjQty,0) from Amor_dtl "
                    strItembal = strItembal & " where account_code = '" & strAccountCode & "'"
                    strItembal = strItembal & " and Item_code = '" & ItemCode & "' and Cust_drgNo = '" & strDrgNo & "' and UNIT_CODE='" & gstrUNITID & "'"
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
                        Exit Sub
                    End If
                End If
            End If
            rssaledtl.MoveNext()
        Next
        If UCase(mstrInvoiceType) = "REJ" Then
            If Len(Trim(strCustRef)) > 0 Then
                If CheckDataFromGrin(CDbl(Trim(strCustRef)), strAccountCode) = False Then
                    Exit Sub
                End If
            End If
        End If
        rssaledtl.ResultSetClose()
        rssaledtl = Nothing
        '-------------------------------------------------
        If Not (InvoiceGeneration() = True) Then
            Exit Sub
        End If
        If ConfirmWindow(10344, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
            If Len(Find_Value("select doc_no from SalesChallan_dtl where location_code='" & Trim(txtLocationCode.Text) & "' and doc_no='" & mInvNo & "' and UNIT_CODE='" & gstrUNITID & "'")) > 0 Then
                MsgBox("Next Invoice number already generated." & vbCrLf & "Please skip current no either backward or forward" & vbCrLf & "in Sales Configuration Master Form.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "eMPro")
                Exit Sub
            End If
            mP_Connection.BeginTrans()
            mP_Connection.Execute("set Dateformat 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            mP_Connection.Execute(salesconf, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            If Len(Trim(mstrExcisePriorityUpdationString)) > 0 Then
                mP_Connection.Execute("update Saleschallan_dtl set Excise_type = '" & mstrExcisePriorityUpdationString & "' where UNIT_CODE='" & gstrUNITID & "' and Doc_no = " & txtChallanNo.Text, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End If
            mP_Connection.Execute(saleschallan, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            If updatePOflag = True Then
                mP_Connection.Execute(strupdatecustodtdtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End If
            mP_Connection.Execute("update i set cur_bal = Cur_bal - Sales_Quantity from itembal_mst i INNER JOIN InvoiceStock_dtl s ON i.item_code = s.item_code and i.Location_code = s.from_Location and i.UNIT_CODE=s.UNIT_CODE where i.UNIT_CODE='" & gstrUNITID & "' AND Doc_no = '" & txtChallanNo.Text & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            If blnCheckToolCost = True Then
                If Len(Trim(strUpdateAmorDtl)) > 0 Then
                    mP_Connection.Execute(strUpdateAmorDtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    If Len(Trim(strupdateamordtlbom)) > 0 Then
                        mP_Connection.Execute(strupdateamordtlbom, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End If
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
                CmdGrpChEnt.Enabled(1) = False
                CmdGrpChEnt.Enabled(2) = False
            End If
            txtChallanNo.Text = CStr(mInvNo)
            txtChallanNo_Validating(txtChallanNo, New System.ComponentModel.CancelEventArgs(False))
        Else
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
        End If
        rsItembal.ResultSetClose()
        rsItembal = Nothing
        Exit Sub
Err_Handler:
        If Err.Number = 20545 Then
            Resume Next
        Else
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub
    Public Function InvoiceGeneration() As Boolean
        Dim rsCompMst As ClsResultSetDB
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
        gobjDB.GetResult("SELECT EOU_Flag, CustSupp_Inc,InsExc_Excise,postinfin,Excise_RoundOFF FROM sales_parameter WHERE UNIT_CODE='" & gstrUNITID & "'")
        If gobjDB.GetValue("EOU_Flag") = True Then
            mStrCustMst = "Select Doc_No,Invoice_type from SalesChallan_Dtl where Invoice_Type <> 'EXP' and Location_Code='" & Trim(txtLocationCode.Text) & "' AND UNIT_CODE='" & gstrUNITID & "'"
            mblnEOUUnit = True
        Else
            mStrCustMst = "Select Doc_No,Invoice_type from SalesChallan_Dtl where Location_Code='" & Trim(txtLocationCode.Text) & "' AND UNIT_CODE='" & gstrUNITID & "'"
            mblnEOUUnit = False
        End If
        mblnAddCustomerMaterial = gobjDB.GetValue("CustSupp_Inc")
        mblnInsuranceFlag = gobjDB.GetValue("InsExc_Excise")
        mblnpostinfin = gobjDB.GetValue("postinfin")
        mblnExciseRoundOFFFlag = gobjDB.GetValue("Excise_RoundOFF")
        rsSalesConf.Open("SELECT * FROM SaleConf WHERE UNIT_CODE='" & gstrUNITID & "' AND Invoice_Type='" & mstrInvoiceType & "' AND Sub_Type ='" & mstrInvoiceSubType & "' AND Location_Code='" & Trim(txtLocationCode.Text) & "' and datediff(dd,getdate(),fin_start_date)<=0  and datediff(dd,fin_end_date,getdate())<=0 ", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
        If Not rsSalesConf.EOF Then
            mstrPurposeCode = IIf(IsDBNull(rsSalesConf.Fields("inv_GLD_prpsCode").Value), "", Trim(rsSalesConf.Fields("inv_GLD_prpsCode").Value))
            mblnSameSeries = rsSalesConf.Fields("Single_Series").Value
            mstrReportFilename = IIf(IsDBNull(rsSalesConf.Fields("Report_filename").Value), "", Trim(rsSalesConf.Fields("Report_filename").Value))
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
        'objInvoicePrint.FileName = gstrLocalCDrive & "InvoicePrint.txt"
        objInvoicePrint.FileName = gstrLocalCDrive & "EmproInv\InvoicePrint.txt"
        'objInvoicePrint.BCFileName = gstrLocalCDrive & "BarCode.txt"
        objInvoicePrint.BCFileName = gstrLocalCDrive & "EmproInv\BarCode.txt"
        objInvoicePrint.CompanyName = gstrCOMPANY
        objInvoicePrint.Address1 = gstr_RGN_ADDRESS1
        objInvoicePrint.Address2 = gstr_RGN_ADDRESS2
        If chkDTRemoval.CheckState = System.Windows.Forms.CheckState.Checked Then
            objInvoicePrint.Print_Invoice(gstrUNITID, True, (txtLocationCode.Text), (txtChallanNo.Text), dtpRemoval.Value & " " & VB6.Format(dtpRemovalTime.Value.Hour, "00") & ":" & VB6.Format(dtpRemovalTime.Value.Minute, "00"))
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
        mDoc_No = 0 : mCustmtrl = 0 : mAmortization = 0 : mstrAnnex = "" : strupdateGrinhdr = "" ': mblnCustSupp = False
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
        strSQL = "select INVOICE_DATE from Saleschallan_Dtl where UNIT_CODE='" & gstrUNITID & "' and Doc_No =" & txtChallanNo.Text & "  and Location_Code='" & Trim(txtLocationCode.Text) & "'"
        rsSalesChallan = New ClsResultSetDB
        rsSalesChallan.GetResult(strSQL, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        strInvoiceDate = VB6.Format(rsSalesChallan.GetValue("Invoice_Date"), gstrDateFormat)
        mInvNo = CDbl(GenerateInvoiceNo(mstrInvoiceType, mstrInvoiceSubType, strInvoiceDate))
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Function CreateStringForAccounts() As Boolean
        '-------------------------------------------------------------------------------------
        'Revised By      : Manoj Kr. Vaish
        'Issue ID        : 19992
        'Revision Date   : 28 June 2007
        'History         : Display Credit Term from Cust_Ord_Dtl and save into saleschallan_dtl,
        '                  During Invoice Posting, fetch credit term from saleschallan_dtl for saving in ar_docmaster
        '-----------------------------------------------------------------------------------
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
        Dim dblInvoiceAmtRoundOff_diff As Double
        rsFULLExciseAmount = New ClsResultSetDB
        mstrExcisePriorityUpdationString = ""
        blnMsgBox = False
        On Error GoTo ErrHandler
        objRecordSet.Open("SELECT * FROM  saleschallan_dtl WHERE UNIT_CODE='" & gstrUNITID & "' AND Doc_No='" & Trim(txtChallanNo.Text) & "' and Location_Code='" & Trim(txtLocationCode.Text) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
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
        strInvoiceDate = VB6.Format(objRecordSet.Fields("Invoice_Date").Value, "dd-MMM-yyyy hh:mm")
        strCurrencyCode = Trim(IIf(IsDBNull(objRecordSet.Fields("Currency_Code").Value), "", objRecordSet.Fields("Currency_Code").Value))
        dblInvoiceAmt = IIf(IsDBNull(objRecordSet.Fields("total_amount").Value), 0, objRecordSet.Fields("total_amount").Value)
        dblInvoiceAmtRoundOff_diff = IIf(IsDBNull(objRecordSet.Fields("TotalInvoiceAmtRoundOff_diff").Value), 0, objRecordSet.Fields("TotalInvoiceAmtRoundOff_diff").Value)
        dblExchangeRate = IIf(IsDBNull(objRecordSet.Fields("Exchange_Rate").Value), 1, objRecordSet.Fields("Exchange_Rate").Value)
        dblTCStaxAmt = IIf(IsDBNull(objRecordSet.Fields("TCSTaxAmount").Value), 1, objRecordSet.Fields("TCSTaxAmount").Value)
        strCustCode = Trim(objRecordSet.Fields("Account_Code").Value)
        strCustRef = Trim(IIf(IsDBNull(objRecordSet.Fields("cust_ref").Value), "", objRecordSet.Fields("cust_ref").Value))
        blnExciseExumpted = objRecordSet.Fields("ExciseExumpted").Value
        strCreditTermsID = Trim(IIf(IsDBNull(objRecordSet.Fields("payment_terms").Value), "", objRecordSet.Fields("payment_terms").Value))
        mstrCreditTermId = strCreditTermsID
        Dim objCreditTerms As New prj_CreditTerm.clsCR_Term_Resolver
        If UCase(mstrInvoiceType) <> "SMP" Then 'if invoice type is not sample sales then
            If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
            If UCase(Trim(mstrInvoiceType)) = "REJ" Then
                objTmpRecordset.Open("SELECT ISNULL(SUM(Basic_Amount),0) AS Basic_Amt FROM sales_dtl WHERE UNIT_CODE='" & gstrUNITID & "' AND doc_no =" & txtChallanNo.Text, mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
                If Not objTmpRecordset.EOF Then
                    dblBasicAmount = objTmpRecordset.Fields("Basic_Amt").Value
                End If
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
                If (UCase(Trim(mstrInvoiceType)) = "REJ" And strCustRef <> "") Then 'In case of non line rejections Basic posting is not done
                    dblInvoiceAmt = dblInvoiceAmt - dblBasicAmount
                End If
                dblBasicAmount = 0
                objTmpRecordset.Open("SELECT GL_AccountID, Ven_slCode, CrTrm_Termid FROM Pur_VendorMaster where UNIT_CODE='" & gstrUNITID & "' AND Prty_PartyID='" & strCustCode & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
            Else
                objTmpRecordset.Open("SELECT Cst_ArCode, Cst_slCode, Cst_CreditTerm FROM Sal_CustomerMaster where UNIT_CODE='" & gstrUNITID & "' AND Prty_PartyID='" & strCustCode & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
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
        Else
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
        Dim rsSalesParameter As New ADODB.Recordset
        Dim blnTotalInvoiceAmountRoundOff As Boolean
        Dim intTotalInvoiceAmountRoundOff As Short
        If rsSalesParameter.State = ADODB.ObjectStateEnum.adStateOpen Then rsSalesParameter.Close()
        rsSalesParameter.Open("SELECT TotalInvoiceAmount_RoundOff, TotalInvoiceAmountRoundOff_Decimal FROM SALES_PARAMETER WHERE UNIT_CODE='" & gstrUNITID & "'", mP_Connection)
        If Not rsSalesParameter.EOF Then
            blnTotalInvoiceAmountRoundOff = rsSalesParameter.Fields("TotalInvoiceAmount_RoundOff").Value
            intTotalInvoiceAmountRoundOff = rsSalesParameter.Fields("TotalInvoiceAmountRoundOff_Decimal").Value
        End If
        If rsSalesParameter.State = ADODB.ObjectStateEnum.adStateOpen Then
            rsSalesParameter.Close()
            rsSalesParameter = Nothing
        End If
        If UCase(Trim(mstrInvoiceType)) <> "REJ" Then
            mstrMasterString = "I»" & strInvoiceNo & "»Dr»»" & strInvoiceDate & "»»»»»SAL»I»" & strInvoiceNo & "»" & strInvoiceDate & "»"
            If UCase(mstrInvoiceType) <> "SMP" Then
                mstrMasterString = mstrMasterString & Trim(strCustCode) & "»" & gstrUNITID & "»" & strCurrencyCode & "»"
            Else
                mstrMasterString = mstrMasterString & "»" & gstrUNITID & "»" & strCurrencyCode & "»"
            End If
            If blnTotalInvoiceAmountRoundOff Then
                mstrMasterString = mstrMasterString & System.Math.Round(dblInvoiceAmt, 0) & "»" & System.Math.Round(dblInvoiceAmt * dblExchangeRate, 0) & "»" & dblExchangeRate & "»" & strCreditTermsID & "»" & strBasicDueDate & "»" & strPaymentDueDate & "»" & strExpectedDueDate & "»" & strCustomerGL & "»" & strCustomerSL & "»" & mP_User & "»getdate()»»"
            Else
                mstrMasterString = mstrMasterString & System.Math.Round(dblInvoiceAmt, intTotalInvoiceAmountRoundOff) & "»" & System.Math.Round(dblInvoiceAmt * dblExchangeRate, intTotalInvoiceAmountRoundOff) & "»" & dblExchangeRate & "»" & strCreditTermsID & "»" & strBasicDueDate & "»" & strPaymentDueDate & "»" & strExpectedDueDate & "»" & strCustomerGL & "»" & strCustomerSL & "»" & mP_User & "»getdate()»»"
            End If
        Else
            If blnTotalInvoiceAmountRoundOff Then
                mstrMasterString = "M»»" & VB6.Format(GetServerDate, "dd-MMM-yyyy") & "»0»»" & gstrUNITID & "»" & Trim(strCustCode) & "»" & strInvoiceNo & "»" & strInvoiceDate & "»" & strBasicDueDate & "»" & strPaymentDueDate & "»" & strExpectedDueDate & "»" & strCurrencyCode & "»" & dblExchangeRate & "»" & System.Math.Round(dblInvoiceAmt) & "»0»»»Rej. Inv. " & strInvoiceNo & "»" & strCustomerGL & "»" & strCustomerSL & "»DR»" & strCustomerGL & "»" & strCustomerSL & "»»" & gstrCURRENCYCODE & "»" & mP_User & "»getdate()»0»AP»»»»0»»¦"
            Else
                mstrMasterString = "M»»" & VB6.Format(GetServerDate, "dd-MMM-yyyy") & "»0»»" & gstrUNITID & "»" & Trim(strCustCode) & "»" & strInvoiceNo & "»" & strInvoiceDate & "»" & strBasicDueDate & "»" & strPaymentDueDate & "»" & strExpectedDueDate & "»" & strCurrencyCode & "»" & dblExchangeRate & "»" & System.Math.Round(dblInvoiceAmt, intTotalInvoiceAmountRoundOff) & "»0»»»Rej. Inv. " & strInvoiceNo & "»" & strCustomerGL & "»" & strCustomerSL & "»DR»" & strCustomerGL & "»" & strCustomerSL & "»»" & gstrCURRENCYCODE & "»" & mP_User & "»getdate()»0»AP»»»»0»»¦"
            End If
        End If
        iCtr = 1
        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
        If Trim(IIf(IsDBNull(objRecordSet.Fields("SalesTax_Type").Value), "", objRecordSet.Fields("SalesTax_Type").Value)) <> "" Then
            objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE UNIT_CODE='" & gstrUNITID & "' AND TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("SalesTax_Type").Value), "", objRecordSet.Fields("SalesTax_Type").Value)) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
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
        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
        If Trim(IIf(IsDBNull(objRecordSet.Fields("ECESS_Type").Value), "", objRecordSet.Fields("ECESS_Type").Value)) <> "" Then
            objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE UNIT_CODE='" & gstrUNITID & "' AND TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("ECESS_Type").Value), "", objRecordSet.Fields("ECESS_Type").Value)) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
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
        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
        If Trim(IIf(IsDBNull(objRecordSet.Fields("SECESS_Type").Value), "", objRecordSet.Fields("SECESS_Type").Value)) <> "" Then
            objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE UNIT_CODE='" & gstrUNITID & "' AND TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("SECESS_Type").Value), "", objRecordSet.Fields("SECESS_Type").Value)) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
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
        'Freight Posting
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
        objRecordSet.Open("SELECT sales_dtl.*, item_mst.GlGrp_code FROM sales_dtl, item_mst WHERE sales_dtl.Doc_No='" & Trim(txtChallanNo.Text) & "' and sales_dtl.Item_Code=item_mst.Item_Code and sales_dtl.Location_Code='" & Trim(txtLocationCode.Text) & "' and sales_dtl.UNIT_CODE=item_mst.UNIT_CODE AND sales_dtl.UNIT_CODE='" & gstrUNITID & "'")
        If objRecordSet.EOF Then
            MsgBox("Item details not found.", MsgBoxStyle.Information, "eMPro")
            CreateStringForAccounts = False
            If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                objRecordSet.Close()
                objRecordSet = Nothing
            End If
            Exit Function
        End If
        Dim blnFOC As Boolean
        While Not objRecordSet.EOF
            strGlGroupId = Trim(IIf(IsDBNull(objRecordSet.Fields("GlGrp_code").Value), "", objRecordSet.Fields("GlGrp_code").Value))
            blnFOC = CBool(Find_Value("select foc_invoice from salesChallan_dtl where UNIT_CODE='" & gstrUNITID & "' AND Location_Code='" & Trim(txtLocationCode.Text) & "' and doc_no='" & Trim(txtChallanNo.Text) & "'"))
            If UCase(Trim(mstrInvoiceType)) = "SMP" And blnFOC Then
            ElseIf (UCase(Trim(mstrInvoiceType)) = "REJ" And strCustRef = "") Or UCase(Trim(mstrInvoiceType)) <> "REJ" Then  'In case of non line rejections Basic posting is not done
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
                    objTmpRecordset.Open("SELECT * FROM invcc_dtl WHERE UNIT_CODE='" & gstrUNITID & "' AND Invoice_Type='" & mstrInvoiceType & "' AND Sub_Type = '" & mstrInvoiceSubType & "' AND Location_Code ='" & Trim(txtLocationCode.Text) & "' AND ccM_cc_Percentage > 0", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
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
                    '*********************************************************
                End If
            End If
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
                    rsFULLExciseAmount.GetResult("Select Sum(isnull(TotalExciseAmount,0)) as TotalExciseAmount from Sales_dtl where UNIT_CODE='" & gstrUNITID & "' AND Doc_no =" & txtChallanNo.Text)
                    dblFullExciseAmount = rsFULLExciseAmount.GetValue("TotalExciseAmount")
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
            dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("Others").Value), 0, objRecordSet.Fields("Others").Value)
            dblBaseCurrencyAmount = dblTaxAmt
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
        dblBaseCurrencyAmount = dblInvoiceAmtRoundOff_diff
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
        strSQL = "select *  from Saleschallan_dtl where UNIT_CODE='" & gstrUNITID & "' AND Doc_No = " & txtChallanNo.Text
        strSQL = strSQL & " and Invoice_type = '" & mstrInvoiceType & "'  and  sub_category =  '" & mstrInvoiceSubType & "' and Location_Code='" & Trim(txtLocationCode.Text) & "'"
        rsSalesChallan = New ClsResultSetDB
        rsSalesChallan.GetResult(strSQL, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsSalesChallan.GetNoRows > 0 Then
            mAccount_Code = rsSalesChallan.GetValue("Account_Code")
            mCust_Ref = rsSalesChallan.GetValue("Cust_ref")
            mAmendment_No = rsSalesChallan.GetValue("Amendment_No")
            dblInvoiceAmt = rsSalesChallan.GetValue("total_amount")
            strInvoiceDate = VB6.Format(rsSalesChallan.GetValue("Invoice_Date"), gstrDateFormat)
        End If
        rsSalesChallan.ResultSetClose()
        rsSalesChallan = Nothing
        If mblnEOUUnit = True Then
            If UCase(mstrInvoiceType) <> "EXP" Then
                If Not mblnSameSeries Then
                    salesconf = "update saleconf set current_No = " & mSaleConfNo & ", OpenningBal = openningBal - " & mAssessableValue & " where UNIT_CODE='" & gstrUNITID & "' and Invoice_type <> 'EXP' and Location_Code='" & Trim(txtLocationCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
                Else
                    salesconf = "update saleconf set current_No = " & mSaleConfNo & " where UNIT_CODE='" & gstrUNITID & "' and Single_Series = 1 and Invoice_type <> 'EXP' and Location_Code='" & Trim(txtLocationCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0" & vbCrLf
                    salesconf = salesconf & "update saleconf set OpenningBal = openningBal - " & mAssessableValue & " where UNIT_CODE='" & gstrUNITID & "' and Invoice_type <> 'EXP' and Location_Code='" & Trim(txtLocationCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
                End If
            Else
                salesconf = "update saleconf set current_No = " & mSaleConfNo & " where UNIT_CODE='" & gstrUNITID & "' and Invoice_type = 'EXP' and Location_Code='" & Trim(txtLocationCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
            End If
        Else
            If Not mblnSameSeries Then
                salesconf = "update saleconf set current_No = " & mSaleConfNo & " where UNIT_CODE='" & gstrUNITID & "' and Invoice_type = '" & mstrInvoiceType & "' and Location_Code='" & Trim(txtLocationCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
            Else
                salesconf = "update saleconf set current_No = " & mSaleConfNo & " where UNIT_CODE='" & gstrUNITID & "' and Single_Series = 1 and Location_Code='" & Trim(txtLocationCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
            End If
        End If
        Dim intInvoicePostingFlag As Short
        If mblnpostinfin = True Then
            intInvoicePostingFlag = 1
        Else
            intInvoicePostingFlag = 0
        End If
        saleschallan = "UPDATE SalesChallan_Dtl SET doc_no=" & mInvNo & ", Bill_Flag=1,print_flag = 1 , postingFlag=" & intInvoicePostingFlag & ",Payment_terms='" & mstrCreditTermId & "',Upd_dt=getdate(),Upd_Userid='" & mP_User & "' WHERE UNIT_CODE='" & gstrUNITID & "' and Doc_No=" & txtChallanNo.Text & " and Location_Code='" & Trim(txtLocationCode.Text) & "' " & vbCrLf
        saleschallan = saleschallan & "UPDATE Sales_Dtl SET doc_no=" & mInvNo & " ,Upd_dt=getdate(),Upd_Userid='" & mP_User & "' WHERE UNIT_CODE='" & gstrUNITID & "' and Doc_No=" & txtChallanNo.Text & " and Location_Code='" & Trim(txtLocationCode.Text) & "'" & vbCrLf
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
        Dim rsSalesParameter As New ClsResultSetDB
        Dim intRow, intLoopCount As Short
        Dim mItem_Code, mCust_Item_Code As String
        Dim mSales_Quantity As Double
        Dim mToolCost As Double
        Dim blnCheckToolCost As Boolean
        Dim strAccountCode As String
        Dim strItembal As String ' update in enagare invoice entry for tool amortisation
        Dim rsMktSchedule As New ClsResultSetDB ' update in enagare invoice entry for tool amortisation
        Dim strQuantity As String ' update in enagare invoice entry for tool amortisation
        Dim strToolCode As String ' update in enagare invoice entry for tool amortisation
        Dim rsbom As New ClsResultSetDB ' update in enagare invoice entry for tool amortisation
        Dim irowcount As Short ' update in enagare invoice entry for tool amortisation
        Dim intRwCount1 As Short ' update in enagare invoice entry for tool amortisation
        Dim varItemQty1 As String ' update in enagare invoice entry for tool amortisation
        strupdateitbalmst = ""
        strupdatecustodtdtl = ""
        strUpdateAmorDtl = ""
        strupdateamordtlbom = ""
        On Error GoTo Err_Handler
        'CODE ADDED BY NISHA ON 21/03/2003 FOR FINANCIAL ROLLOVER
        mP_Connection.Execute("Delete from InvoiceStock_dtl where UNIT_CODE='" & gstrUNITID & "' and Doc_no = '" & txtChallanNo.Text & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        strSQL = "select * from Saleschallan_Dtl where UNIT_CODE='" & gstrUNITID & "' and Doc_No =" & txtChallanNo.Text & "  and Location_Code='" & Trim(txtLocationCode.Text) & "'"
        rsSalesChallan = New ClsResultSetDB
        rsSalesChallan.GetResult(strSQL, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        strInvoiceDate = VB6.Format(rsSalesChallan.GetValue("Invoice_Date"), gstrDateFormat)
        strAccountCode = rsSalesChallan.GetValue("Account_code")
        rsSaleConf = New ClsResultSetDB
        rsSaleConf.GetResult("Select Stock_Location from saleconf where UNIT_CODE='" & gstrUNITID & "' and Description = '" & CmbInvType.Text & "' and Sub_Type_Description ='" & Me.CmbInvSubType.Text & "' and Location_Code='" & Trim(txtLocationCode.Text) & "'and datediff(dd,'" & getDateForDB(strInvoiceDate) & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0")
        strStockLocCode = rsSaleConf.GetValue("Stock_Location")
        strSQL = "Select * from sales_Dtl where UNIT_CODE='" & gstrUNITID & "' and Doc_No = " & txtChallanNo.Text & " and Location_Code='" & Trim(txtLocationCode.Text) & "'"
        rsSalesParameter.GetResult("Select CheckToolAmortisation from Sales_Parameter where UNIT_CODE='" & gstrUNITID & "'")
        If rsSalesParameter.GetNoRows > 0 Then
            rsSalesParameter.MoveFirst()
            If Len(Trim(rsSalesParameter.GetValue("CheckToolAmortisation"))) = 0 Then
                MsgBox("First define Check Tool Amortisation in Sales Parameter", MsgBoxStyle.Information, "eMPro")
                Exit Sub
            End If
            blnCheckToolCost = rsSalesParameter.GetValue("CheckToolAmortisation")
        End If
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
                    mP_Connection.Execute("Insert into InvoiceStock_dtl values('" & txtChallanNo.Text & "','" & mItem_Code & "'," & mSales_Quantity & ",'" & strStockLocation & "','" & gstrUNITID & "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    strupdateitbalmst = Trim(strupdateitbalmst) & "Update Itembal_mst set cur_bal= cur_bal-"
                    strupdateitbalmst = strupdateitbalmst & mSales_Quantity & " where UNIT_CODE='" & gstrUNITID & "' and Location_code = '" & strStockLocation
                    strupdateitbalmst = strupdateitbalmst & "' and item_code = '" & mItem_Code & "'"
                    strupdatecustodtdtl = Trim(strupdatecustodtdtl) & "Update Cust_ord_dtl set Despatch_Qty = Despatch_Qty + "
                    strupdatecustodtdtl = strupdatecustodtdtl & mSales_Quantity & " where Account_code ='"
                    strupdatecustodtdtl = strupdatecustodtdtl & mAccount_Code & "'and Cust_DrgNo = '"
                    strupdatecustodtdtl = strupdatecustodtdtl & mCust_Item_Code & "' and Cust_ref = '" & mCust_Ref
                    strupdatecustodtdtl = strupdatecustodtdtl & "'and amendment_no = '" & mAmendment_No & "' and active_Flag ='A' and UNIT_CODE='" & gstrUNITID & "'"
                    If blnCheckToolCost = True Then
                        strItembal = "select BalanceQty = isnull(a.proj_qty,0) - isnull(a.ClosingValueSMIEL,0),a.Tool_C from Amor_dtl a,Tool_Mst b"
                        strItembal = strItembal & " where account_code = '" & strAccountCode & "'"
                        strItembal = strItembal & " and Item_code = '" & mItem_Code & "' and a.Tool_c = b.tool_c and a.Item_code = b.Product_No and A.unit_code=B.UNIT_CODE AND a.UNIT_CODE='" & gstrUNITID & "' order by a.tool_c"
                        rsMktSchedule.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        If rsMktSchedule.GetNoRows > 0 Then
                            rsMktSchedule.MoveFirst()
                            strQuantity = CStr(Val(rsMktSchedule.GetValue("BalanceQty")))
                            strToolCode = rsMktSchedule.GetValue("tool_c")
                            strItembal = "select BalanceQty = sum(isnull(usedProjQty,0)) from Amor_dtl a "
                            strItembal = strItembal & " Where a.UNIT_CODE='" & gstrUNITID & "' AND Item_code = '" & mItem_Code & "' and a.Tool_c = '" & strToolCode & "'"
                            rsMktSchedule.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                            strQuantity = CStr(CDbl(strQuantity) - Val(rsMktSchedule.GetValue("BalanceQty")))
                            If Val(CStr(mSales_Quantity)) > Val(strQuantity) Then
                                If CDbl(strQuantity) = 0 Then
                                    MsgBox("No Balance Available for Item (" & mItem_Code & ") and customer Part Code (" & mCust_Item_Code & ") For Amortisation Calculations. ", MsgBoxStyle.OkOnly, "eMPro")
                                Else
                                    MsgBox("Quantity should not be Greater then available Balance Quantity for Amortisarion " & strQuantity, MsgBoxStyle.OkOnly, "eMPro")
                                End If
                                Exit Sub
                            Else
                                strUpdateAmorDtl = Trim(strUpdateAmorDtl) & " Update Amor_dtl set usedProjQty = "
                                strUpdateAmorDtl = Trim(strUpdateAmorDtl) & " isnull(usedProjQty,0) + " & mSales_Quantity
                                strUpdateAmorDtl = Trim(strUpdateAmorDtl) & " where UNIT_CODE='" & gstrUNITID & "' AND account_code = '" & strAccountCode
                                strUpdateAmorDtl = Trim(strUpdateAmorDtl) & "' and tool_c = '" & strToolCode & "'"
                                strUpdateAmorDtl = Trim(strUpdateAmorDtl) & " and item_code = '" & mItem_Code & "'"
                            End If
                        Else
                            strItembal = "select BalanceQty = isnull(proj_qty,0) - isnull(ClosingValueSMIEL,0) from Amor_dtl "
                            strItembal = strItembal & " where UNIT_CODE='" & gstrUNITID & "' AND account_code = '" & strAccountCode & "'"
                            strItembal = strItembal & " and Item_code = '" & mItem_Code & "' "
                            rsMktSchedule.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                            If rsMktSchedule.GetNoRows > 0 Then
                                rsMktSchedule.MoveFirst()
                                strQuantity = CStr(Val(rsMktSchedule.GetValue("BalanceQty")))
                                strItembal = "select BalanceQty = sum(isnull(usedProjQty,0)) from Amor_dtl "
                                strItembal = strItembal & " Where UNIT_CODE='" & gstrUNITID & "' AND Item_code = '" & mItem_Code & "'"
                                rsMktSchedule.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                                strQuantity = CStr(CDbl(strQuantity) - Val(rsMktSchedule.GetValue("BalanceQty")))
                                If Val(CStr(mSales_Quantity)) > Val(strQuantity) Then
                                    If CDbl(strQuantity) = 0 Then
                                        MsgBox("No Balance Available for Item (" & mItem_Code & ") and customer Part Code (" & mCust_Item_Code & ") For Amortisation Calculations. ", MsgBoxStyle.OkOnly, "eMPro")
                                    Else
                                        MsgBox("Quantity should not be Greater then available Balance Quantity for Amortisarion " & strQuantity, MsgBoxStyle.OkOnly, "eMPro")
                                    End If
                                    Exit Sub
                                Else
                                    strUpdateAmorDtl = Trim(strUpdateAmorDtl) & " Update Amor_dtl set usedProjQty = "
                                    strUpdateAmorDtl = Trim(strUpdateAmorDtl) & " isnull(usedProjQty,0) + " & mSales_Quantity
                                    strUpdateAmorDtl = Trim(strUpdateAmorDtl) & " where UNIT_CODE='" & gstrUNITID & "' AND account_code = '" & strAccountCode & "'"
                                    strUpdateAmorDtl = Trim(strUpdateAmorDtl) & " and item_code = '" & mItem_Code & "'"
                                End If
                            End If
                        End If
                        With mP_Connection
                            '.Execute("Truncate Table tmpBOM", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            .Execute("DELETE tmpBOM WHERE UNIT_CODE='" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            .Execute("BOMExplosion '" & Trim(mItem_Code) & "','" & Trim(mItem_Code) & "',1,0,0,0,'" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        End With
                        rsbom.GetResult("select * from tmpBOM WHERE UNIT_CODE='" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        If rsbom.GetNoRows > 0 Then
                            irowcount = rsbom.GetNoRows
                            rsbom.MoveFirst()
                            For intRwCount1 = 1 To irowcount
                                strItembal = "select BalanceQty = isnull(a.proj_qty,0) - isnull(a.ClosingValueSMIEL,0),a.tool_C from Amor_dtl a, tool_mst b "
                                strItembal = strItembal & " where account_code = '" & Trim(strAccountCode) & "'"
                                strItembal = strItembal & " and Item_code = '" & rsbom.GetValue("item_code") & "' and a.Tool_c = b.Tool_c and a.ITem_code = b.Product_no AND a.UNIT_CODE=b.UNIT_CODE AND a.UNIT_CODE='" & gstrUNITID & "' order by a.tool_c"
                                rsMktSchedule.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                                If rsMktSchedule.GetNoRows > 0 Then
                                    rsMktSchedule.MoveFirst()
                                    strQuantity = CStr(Val(rsMktSchedule.GetValue("BalanceQty")))
                                    strToolCode = rsMktSchedule.GetValue("tool_c")
                                    varItemQty1 = CStr(mSales_Quantity * Val(rsbom.GetValue("grossweight")))
                                    strItembal = "select BalanceQty = sum(isnull(usedProjQty,0)) from Amor_dtl a "
                                    strItembal = strItembal & " where a.UNIT_CODE='" & gstrUNITID & "' AND account_code = '" & Trim(strAccountCode) & "'"
                                    strItembal = strItembal & " and Item_code = '" & rsbom.GetValue("item_code") & "' and a.Tool_c '" & strToolCode & "'"
                                    strQuantity = CStr(CDbl(strQuantity) - Val(rsMktSchedule.GetValue("BalanceQty")))
                                    rsMktSchedule.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                                    If Val(varItemQty1) > Val(strQuantity) Then
                                        If CDbl(strQuantity) = 0 Then
                                            MsgBox("No Balance Available for Item (" & rsbom.GetValue("item_code") & ") and customer Part Code (" & mCust_Item_Code & ") For Amortisation Calculations. ", MsgBoxStyle.OkOnly, "eMPro")
                                        Else
                                            MsgBox("Quantity should not be Greater then available Balance Quantity for Amortisarion of this Item (" & rsbom.GetValue("item_code") & ")" & mSales_Quantity, MsgBoxStyle.OkOnly, "eMPro")
                                        End If
                                        Exit Sub
                                    Else
                                        strupdateamordtlbom = Trim(strupdateamordtlbom) & " Update Amor_dtl set usedProjQty = "
                                        strupdateamordtlbom = Trim(strupdateamordtlbom) & " isnull(usedProjQty,0) + " & mSales_Quantity * Val(rsbom.GetValue("grossweight"))
                                        strupdateamordtlbom = Trim(strupdateamordtlbom) & " where UNIT_CODE='" & gstrUNITID & "' AND account_code = '" & strAccountCode
                                        strupdateamordtlbom = Trim(strupdateamordtlbom) & "' and Item_code = '" & rsbom.GetValue("item_code")
                                        strupdateamordtlbom = Trim(strupdateamordtlbom) & "' and tool_c = '" & strToolCode & "'"
                                    End If
                                End If
                                rsbom.MoveNext()
                            Next
                        End If
                    End If
                    rssaledtl.MoveNext()
                End If
            Next
        End If
        rssaledtl.ResultSetClose()
        rssaledtl = Nothing
        rsSaleConf.ResultSetClose()
        rsSaleConf = Nothing
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
        rsSalesDtl.GetResult("select * from sales_dtl where UNIT_CODE='" & gstrUNITID & "' AND Doc_No = " & txtChallanNo.Text & " and Location_Code='" & Trim(txtLocationCode.Text) & "'")
        If rsSalesDtl.GetNoRows > 0 Then
            intMaxLoop = rsSalesDtl.GetNoRows
            rsSalesDtl.MoveFirst()
            strupdateGrinhdr = ""
            For intLoopCount = 1 To intMaxLoop
                strItemCode = rsSalesDtl.GetValue("ITem_code")
                dblqty = rsSalesDtl.GetValue("Sales_Quantity")
                If Len(Trim(strupdateGrinhdr)) = 0 Then
                    strupdateGrinhdr = "Update Grn_Dtl Set Despatch_Quantity = isnull(Despatch_Quantity,0) +" & dblqty
                    strupdateGrinhdr = strupdateGrinhdr & " Where UNIT_CODE='" & gstrUNITID & "' AND ITem_Code = '" & strItemCode & "' and Doc_No = " & pdblGrinNo
                Else
                    strupdateGrinhdr = strupdateGrinhdr & vbCrLf & "Update Grn_Dtl Set Despatch_Quantity = isnull(Despatch_Quantity,0) + " & dblqty
                    strupdateGrinhdr = strupdateGrinhdr & " Where UNIT_CODE='" & gstrUNITID & "' AND ITem_Code = '" & strItemCode & "' and Doc_No = " & pdblGrinNo
                End If
                rsSalesDtl.MoveNext()
            Next
        Else
            MsgBox("No Items Available in Invoice " & txtChallanNo.Text)
        End If
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
            strSQL = "Select Current_No,Suffix,Fin_start_date,Fin_end_Date From saleConf Where "
            strSQL = strSQL & "Invoice_Type ='" & pstrInvoiceType & "' and  sub_type='" & pstrInvoiceSubType & "' AND Location_Code ='" & Trim(txtLocationCode.Text) & "' and datediff(dd,'" & getDateForDB(pstrRequiredDate) & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(pstrRequiredDate) & "')<=0 AND UNIT_CODE='" & gstrUNITID & "'"
            With clsInstEMPDBDbase.CConnection
                .OpenConnection(gstrDSNName, gstrDatabaseName)
                .ExecuteSQL("Set Dateformat 'dmy'")
            End With
            clsInstEMPDBDbase.CRecordset.OpenRecordset(strSQL, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
            If clsInstEMPDBDbase.CRecordset.Recordcount > 0 Then
                strCheckDOcNo = CStr(clsInstEMPDBDbase.CRecordset.GetFieldValue("Current_No", EMPDataBase.EMPDB.ADODataType.ADONumeric, EMPDataBase.EMPDB.ADOCustomFormat.CustomZeroDecimal))
                strSuffix = CStr(clsInstEMPDBDbase.CRecordset.GetFieldValue("suffix", EMPDataBase.EMPDB.ADODataType.ADONumeric, EMPDataBase.EMPDB.ADOCustomFormat.CustomZeroDecimal))
                strFin_Start_Date = CStr(clsInstEMPDBDbase.CRecordset.GetFieldValue("Fin_Start_Date", EMPDataBase.EMPDB.ADODataType.ADODate, EMPDataBase.EMPDB.ADOCustomFormat.CustomDate))
                strFin_End_Date = CStr(clsInstEMPDBDbase.CRecordset.GetFieldValue("Fin_End_Date", EMPDataBase.EMPDB.ADODataType.ADODate, EMPDataBase.EMPDB.ADOCustomFormat.CustomDate))
            Else
                Err.Raise(vbObjectError + 20008, "[GenerateDocNo]", "Incorrect Parameters Passed Invoice Number cannot be Generated.")
            End If
            clsInstEMPDBDbase.CRecordset.CloseRecordset() 'Close Recordset
        Else
            Err.Raise(vbObjectError + 20007, "[GenerateDocNo]", "Wanted Date Information not Passed")
        End If
        If Len(Trim(strCheckDOcNo)) > 0 Then 'That is the Document is Made for that Perio
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
            GenerateInvoiceNo = strTempSeries
        End If
        Exit Function
ErrHandler:
        'Logging the ERROR at Application's Path
        Dim clsErrorInst As New EMPDataBase.EMPDB(gstrUNITID)
        clsErrorInst.CError.RaiseError(20008, "[frmMKTTRN0068]", "[GenerateInvoiceNo]", "", "No. Not Generated For DocType = " & pstrInvoiceType & " due to [ " & Err.Description & " ].", My.Application.Info.DirectoryPath, gstrDSNName, gstrDatabaseName)
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
        rsGrnDtl = New ClsResultSetDB
        rsSalesDtl = New ClsResultSetDB
        rsSalesDtl.GetResult("Select Item_Code,Sales_Quantity from Sales_dtl where UNIT_CODE='" & gstrUNITID & "' AND doc_No =" & txtChallanNo.Text & " and Location_Code='" & Trim(txtLocationCode.Text) & "'")
        intMaxLoop = rsSalesDtl.GetNoRows : rsSalesDtl.MoveFirst()
        CheckDataFromGrin = False
        For intLoopCounter = 1 To intMaxLoop
            strItemCode = rsSalesDtl.GetValue("Item_code")
            dblItemQty = rsSalesDtl.GetValue("Sales_quantity")
            strSQL = "select a.Doc_No,a.Item_code,a.Rejected_Quantity, a.excess_po_quantity ,"
            strSQL = strSQL & "Despatch_Quantity = isnull(a.Despatch_Quantity,0),"
            strSQL = strSQL & " Inspected_Quantity = isnull(Inspected_Quantity,0),"
            strSQL = strSQL & "RGP_Quantity = isnull(RGP_Quantity,0) from grn_Dtl a,grn_hdr b Where "
            strSQL = strSQL & "a.Doc_type = b.Doc_type And a.Doc_No = b.Doc_No and "
            strSQL = strSQL & "a.From_Location = b.From_Location and a.From_Location ='01R1'"
            strSQL = strSQL & "and a.Rejected_quantity > 0 and b.Vendor_code = '" & pstrCustCode
            strSQL = strSQL & "' and a.Doc_No = " & pdblDocNo & " and a.Item_code = '" & strItemCode & "' AND a.UNIT_CODE=b.UNIT_CODE AND a.UNIT_CODE='" & gstrUNITID & "'"
            rsGrnDtl.GetResult(strSQL)
            dblRejQty = rsGrnDtl.GetValue("Rejected_Quantity") + rsGrnDtl.GetValue("excess_po_Quantity") - rsGrnDtl.GetValue("Despatch_Quantity") - rsGrnDtl.GetValue("Inspected_Quantity") - rsGrnDtl.GetValue("RGP_Quantity")
            If rsGrnDtl.GetNoRows > 0 Then
                If dblItemQty > (dblRejQty) Then
                    MsgBox("Max. Quantity Allowed For Item " & strItemCode & " is " & dblRejQty & ", Quantity Entered in Invoice is : " & dblItemQty)
                    CheckDataFromGrin = False
                    Exit Function
                Else
                    CheckDataFromGrin = True
                End If
            End If
            rsSalesDtl.MoveNext()
        Next
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
        objRecordSet.Open("SELECT invGld_glcode, invGld_slcode FROM fin_InvGLGrpDtl WHERE UNIT_CODE='" & gstrUNITID & "' AND invGld_prpsCode = '" & PurposeCode & "' AND invGld_invGLGrpId = '" & InventoryGlGroup & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
        If objRecordSet.EOF Then
            objRecordSet.Close()
            objRecordSet.Open("SELECT gbl_glCode, gbl_slCode FROM fin_globalGL WHERE UNIT_CODE='" & gstrUNITID & "' AND gbl_prpsCode = '" & PurposeCode & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
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
        objRecordSet.Open("SELECT tx_glCode, tx_slCode FROM fin_TaxGlRel where UNIT_CODE='" & gstrUNITID & "' AND tx_rowType = 'ARTAX' AND tx_taxId ='" & TaxType & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
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
        strSQL = "Select * from Tax_PriorityMst where UNIT_CODE='" & gstrUNITID & "'"
        rsTaxPriority.GetResult(strSQL)
        If rsTaxPriority.GetNoRows > 0 Then
            rsTaxPriority.MoveFirst()
            CheckExcPriority = True
            If Len(Trim(rsTaxPriority.GetValue("VarExPriority1"))) = 0 Then
                If Len(Trim(rsTaxPriority.GetValue("VarExPriority2"))) = 0 Then
                    If Len(Trim(rsTaxPriority.GetValue("VarExPriority3"))) = 0 Then
                        CheckExcPriority = False
                        Exit Function
                    End If
                End If
            End If
        Else
            CheckExcPriority = False
        End If
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
        rsCheckBalance = New ClsResultSetDB
        strSQL = "Select VarExPriority1,VarExGL1,VarExSL1,VarExPriority2,VarExGL2,VarExSL2,VarExPriority3,VarExGL3,VarExSL3 from Tax_PriorityMst WHERE UNIT_CODE='" & gstrUNITID & "'"
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
                        strBalance = strBalance & Trim(txtLocationCode.Text) & "' and br_slCode = '" & strExcSL & "'"
                        strBalance = strBalance & " and br_glCode = '" & strExcGL & "'"
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
                        strBalance = strBalance & Trim(txtLocationCode.Text) & "' and br_slCode = '" & strExcSL & "'"
                        strBalance = strBalance & " and br_glCode = '" & strExcGL & "'"
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
                        strBalance = strBalance & " and br_glCode = '" & strExcGL & "'"
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
                    Else
                        ReturnGLSLAccExcPriority = VB6.CopyArray(StrData)
                    End If
                Else
                    ReturnGLSLAccExcPriority = VB6.CopyArray(StrData)
                End If
        End Select
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
        rsVandBom = New ClsResultSetDB
        On Error GoTo ErrHandler
        '****
        For intLoopCount = 0 To intMaxCount
            rsVandBom.GetResult("Select RawMaterial_Code from Vendor_bom where UNIT_CODE='" & gstrUNITID & "' AND Finish_Product_code = '" & pstrFinishedItem & "' and Vendor_code = '" & strCustCode & "' and rawMaterial_code ='" & parrCustAnnex(0, intLoopCount) & "'")
            If rsVandBom.GetNoRows > 0 Then
                strRef57F4 = Replace(ref57f4, "§", "','")
                strRef57F4 = "'" & strRef57F4 & "'"
                strannex = "Select Balance_qty,Ref57f4_No,ref57f4_Date from CustAnnex_HDR "
                strannex = strannex & " WHERE UNIT_CODE='" & gstrUNITID & "' AND Item_code ='" & parrCustAnnex(0, intLoopCount) & "' and Customer_code ='"
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
                        str57f4Date = VB6.Format(rsCustAnnex.GetValue("ref57f4_Date"), gstrDateFormat)
                        mstrAnnex = Trim(mstrAnnex) & " Update CustAnnex_HDR "
                        If dblbalanceqty < parrCustAnnex(1, intLoopCount) Then
                            mstrAnnex = Trim(mstrAnnex) & " Set Balance_Qty = 0 "
                        Else
                            mstrAnnex = Trim(mstrAnnex) & " Set Balance_Qty = Balance_Qty - " & parrCustAnnex(1, intLoopCount)
                        End If
                        mstrAnnex = mstrAnnex & " WHERE UNIT_CODE='" & gstrUNITID & "' AND Item_code ='" & parrCustAnnex(0, intLoopCount) & "' and Customer_code ='"
                        mstrAnnex = mstrAnnex & strCustCode & "' and Ref57f4_No ='" & strRef57F4 & "' "
                        mstrAnnex = mstrAnnex & "Insert into CustAnnex_dtl (Doc_Ty,"
                        mstrAnnex = mstrAnnex & "Invoice_No,Invoice_Date,ref57f4_Date,Ref57f4_No,"
                        mstrAnnex = mstrAnnex & "Item_Code,Quantity,"
                        mstrAnnex = mstrAnnex & "Customer_Code,"
                        mstrAnnex = mstrAnnex & "Location_Code,Product_Code,Ent_Userid,Ent_dt,"
                        mstrAnnex = mstrAnnex & "Upd_Userid,Upd_dt,UNIT_CODE) values ('O'," & mInvNo & ",GetDate(),'" & getDateForDB(str57f4Date) & "','"
                        mstrAnnex = mstrAnnex & ref57f4 & "','" & parrCustAnnex(0, intLoopCount) & "'," & parrCustAnnex(1, intLoopCount) & ","
                        mstrAnnex = mstrAnnex & "'" & strCustCode & "',"
                        mstrAnnex = mstrAnnex & "'SMIL','" & pstrFinishedItem & "','" & mP_User & "',GETDATE(),'"
                        mstrAnnex = mstrAnnex & mP_User & "',GETDATE(),'" & gstrUNITID & "')"
                        If dblbalanceqty < parrCustAnnex(1, intLoopCount) Then
                            mP_Connection.Execute(" insert into tempCustAnnex (Ref57f4_No,Annex_No,ref57f4_date,Item_code,Quantity,Balance_qty,finishedItem,UNIT_CODE) values ('" & strRef57F4 & "',0,'" & getDateForDB(str57f4Date) & "','" & parrCustAnnex(0, intLoopCount) & "'," & dblbalanceqty & ",0,'" & pstrFinishedItem & "','" & gstrUNITID & "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            parrCustAnnex(1, intLoopCount) = parrCustAnnex(1, intLoopCount) - dblbalanceqty
                        Else
                            mP_Connection.Execute(" insert into tempCustAnnex (Ref57f4_No,Annex_No,ref57f4_date,Item_code,Quantity,Balance_qty,finishedItem,UNIT_CODE) values ('" & strRef57F4 & "',0,'" & getDateForDB(str57f4Date) & "','" & parrCustAnnex(0, intLoopCount) & "'," & parrCustAnnex(1, intLoopCount) & "," & dblbalanceqty - parrCustAnnex(1, intLoopCount) & ",'" & pstrFinishedItem & "','" & gstrUNITID & "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            parrCustAnnex(1, intLoopCount) = 0
                        End If
                        rsCustAnnex.MoveNext()
                    Else
                        Exit For
                    End If
                Next
            End If
        Next
        rsCustAnnex = Nothing
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Public Function Generate_Unique_FileName(ByRef pstrFileName As String) As Object
        '----------------------------------------------------------------------------
        'Author         :   Arshad Ali
        'Argument       :   name of user
        'Return Value   :   return unique file name
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim lStrFileName As String
        Dim lIntRandomNo As Short
        Dim unique_char As String
        Dim unique_int As String
        Dim unique_no As String
        Dim lStrTempFileName As String
        Dim intCount As Short
        lStrTempFileName = pstrFileName
        lStrTempFileName = Replace(lStrTempFileName, " ", "") ' remove blank spaces in name
        If Len(lStrTempFileName) > 10 Then
            lStrFileName = Mid(lStrTempFileName, 1, 9)
        Else
            lStrFileName = Mid(lStrTempFileName, 1, Len(lStrTempFileName))
        End If
        Randomize()
        For intCount = 1 To 2
            If Len(lStrTempFileName) > 10 Then
                lIntRandomNo = Int((9 * Rnd()) + 1) ' Generate random value between 1 and 9.
            Else
                lIntRandomNo = Int((Len(lStrTempFileName) * Rnd()) + 1) ' Generate random value between 1 and len of name
            End If
            unique_char = unique_char & Mid(lStrFileName, lIntRandomNo, 1)
        Next
        For intCount = 1 To 3
            unique_int = unique_int & CStr(Int((9 * Rnd()) + 1)) ' Generate random value between 1 and 9.
        Next
        unique_no = unique_char & unique_int
        Generate_Unique_FileName = UCase(unique_no) & ".txt"
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    '*****************************************PRINTING WITH CRYSTAL REPORT**********************************
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
        On Error GoTo Err_Handler
        rssaledtl = New ClsResultSetDB
        rsItembal = New ClsResultSetDB
        rssaledtl = New ClsResultSetDB
        rsCompany = New ClsResultSetDB
        SALEDTL = "select * from Saleschallan_Dtl where UNIT_CODE='" & gstrUNITID & "' AND  Doc_No =" & txtChallanNo.Text & "  and Location_Code='" & Trim(txtLocationCode.Text) & "'"
        rsSalesChallan = New ClsResultSetDB
        rssaledtl.GetResult(SALEDTL, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        strAccountCode = rssaledtl.GetValue("Account_code")
        strCustRef = rssaledtl.GetValue("Cust_ref")
        StrAmendmentNo = rssaledtl.GetValue("Amendment_No")
        strInvoiceDate = VB6.Format(rssaledtl.GetValue("Invoice_Date"), gstrDateFormat)
        strSalesconf = "Select UpdatePO_Flag,UpdateStock_Flag,Stock_Location,OpenningBal, report_filename, Single_Series ,Preprinted_Flag,NoCopies from saleconf where "
        strSalesconf = strSalesconf & "Invoice_type = '" & mstrInvoiceType & "' and sub_type = '"
        strSalesconf = strSalesconf & mstrInvoiceSubType & "' and Location_Code='" & Trim(txtLocationCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0 AND UNIT_CODE='" & gstrUNITID & "'"
        rsSalesConf = New ClsResultSetDB
        rsSalesConf.GetResult(strSalesconf, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        updatePOflag = rsSalesConf.GetValue("UpdatePO_Flag")
        updatestockflag = rsSalesConf.GetValue("UpdateStock_Flag")
        strStockLocation = rsSalesConf.GetValue("Stock_Location")
        mOpeeningBalance = Val(rsSalesConf.GetValue("OpenningBal"))
        mIntNoCopies = rsSalesConf.GetValue("NoCopies")
        mstrReportFilename = rsSalesConf.GetValue("Report_Filename")
        If Len(Trim(strStockLocation)) = 0 Then
            MsgBox("Please Define Stock Location in Sales Configuration. ")
            Exit Sub
        End If
        'Changed for Issue ID 19992 Starts (Check Temporary No. for Balance Check)
        If Val(txtChallanNo.Text) > 99000000 Then
            '***********To check if Tool Cost Deduction will be done or Not on 16/02/2004
            rsSalesParameter.GetResult("Select CheckToolAmortisation from Sales_Parameter WHERE UNIT_CODE='" & gstrUNITID & "'")
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
            '*************
            SALEDTL = "Select Sales_Quantity,Item_code,Cust_Item_Code,toolcost_amount from sales_Dtl where Doc_No = " & txtChallanNo.Text & " and Location_Code='" & Trim(txtLocationCode.Text) & "' AND UNIT_CODE='" & gstrUNITID & "'"
            rssaledtl.GetResult(SALEDTL, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            intRow = rssaledtl.GetNoRows
            rssaledtl.MoveFirst()
            '******Check for balance & despatch in Cust_ord_dtl
            For intLoopCount = 1 To intRow
                ItemCode = rssaledtl.GetValue("Item_code")
                salesQuantity = rssaledtl.GetValue("Sales_quantity")
                strDrgNo = rssaledtl.GetValue("Cust_Item_code")
                dblToolCost = rssaledtl.GetValue("ToolCost_amount")
                rsItembal.GetResult("Select Cur_bal from Itembal_Mst where Item_code = '" & ItemCode & "'and Location_code ='" & strStockLocation & "' AND UNIT_CODE='" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                If rsItembal.GetNoRows > 0 Then
                    If salesQuantity > rsItembal.GetValue("Cur_Bal") Then
                        MsgBox("Balance for item " & ItemCode & " at Location " & strStockLocation & " not available. ", MsgBoxStyle.Information, "eMPro")
                        Exit Sub
                    End If
                Else
                    MsgBox("No Item in ItemMaster for Location " & strStockLocation & ".", MsgBoxStyle.OkOnly, "eMPro")
                    rsSalesConf.ResultSetClose()
                    rsSalesConf = Nothing
                    Exit Sub
                End If
                If Len(Trim(strCustRef)) > 0 Then
                    If UCase(mstrInvoiceType) <> "REJ" Then
                        rsItembal.GetResult("Select balanceQty = order_qty - despatch_Qty,OpenSO from Cust_ord_dtl where account_code ='" & strAccountCode & "' and Cust_ref ='" & strCustRef & "' and Amendment_No = '" & StrAmendmentNo & "' and Item_code ='" & ItemCode & "' and Cust_drgNo ='" & strDrgNo & "' and Active_flag ='A' AND UNIT_CODE='" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                        If rsItembal.GetNoRows > 0 Then
                            If rsItembal.GetValue("OpenSO") = False Then
                                If salesQuantity > rsItembal.GetValue("BalanceQty") Then
                                    MsgBox("Balance Quantity in SO for item " & ItemCode & " is " & rsItembal.GetValue("BalanceQty") & ".Check Quantity of Item in Challan.", MsgBoxStyle.Information, "eMPro")
                                    Exit Sub
                                End If
                            End If
                        Else
                            MsgBox("No Item (" & strItemCode & ") exist in SO - " & strCustRef & ".", MsgBoxStyle.Information, "eMPro")
                            Exit Sub
                        End If
                    End If
                End If

                If blnCheckToolCost = True Then
                    If dblToolCost > 0 Then
                        strItembal = "select BalanceQty = isnull(proj_qty,0) - isnull(UsedProjQty,0) from Amor_dtl "
                        strItembal = strItembal & " where account_code = '" & strAccountCode & "'"
                        strItembal = strItembal & " and Item_code = '" & ItemCode & "' and Cust_drgNo = '" & strDrgNo & "' AND UNIT_CODE='" & gstrUNITID & "'"
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
                        End If
                    End If
                End If
                rssaledtl.MoveNext()
            Next
            If UCase(mstrInvoiceType) = "REJ" Then
                If Len(Trim(strCustRef)) > 0 Then
                    If CheckDataFromGrin(CDbl(Trim(strCustRef)), strAccountCode) = False Then
                        Exit Sub
                    End If
                End If
            End If
            rssaledtl.ResultSetClose()
            rssaledtl = Nothing
        End If
        If Not (InvoiceGenerationRPT() = True) Then
            Exit Sub
        End If
        If Val(txtChallanNo.Text) > 99000000 Then
            If ConfirmWindow(10344, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                If Len(Find_Value("select doc_no from SalesChallan_dtl where location_code='" & Trim(txtLocationCode.Text) & "' and doc_no='" & mInvNo & "' AND UNIT_CODE='" & gstrUNITID & "'")) > 0 Then
                    MsgBox("Next Invoice number already generated." & vbCrLf & "Please skip current no either backward or forward" & vbCrLf & "in Sales Configuration Master Form.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "eMPro")
                    Exit Sub
                End If
                mP_Connection.BeginTrans()
                mP_Connection.Execute("set Dateformat 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                mP_Connection.Execute(salesconf, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                If Len(Trim(mstrExcisePriorityUpdationString)) > 0 Then
                    mP_Connection.Execute("update Saleschallan_dtl set Excise_type = '" & mstrExcisePriorityUpdationString & "' where UNIT_CODE='" & gstrUNITID & "' AND Doc_no = " & txtChallanNo.Text, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                End If
                mP_Connection.Execute(saleschallan, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                If updatePOflag = True Then
                    mP_Connection.Execute(strupdatecustodtdtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                End If
                mP_Connection.Execute("update i set cur_bal = Cur_bal - Sales_Quantity from itembal_mst i INNER JOIN InvoiceStock_dtl s ON i.item_code = s.item_code and i.Location_code = s.from_Location AND i.UNIT_CODE=s.UNIT_CODE where Doc_no = '" & txtChallanNo.Text & "' AND i.UNIT_CODE='" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                If blnCheckToolCost = True Then
                    If Len(Trim(strUpdateAmorDtl)) > 0 Then
                        mP_Connection.Execute(strUpdateAmorDtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        If Len(Trim(strupdateamordtlbom)) > 0 Then
                            mP_Connection.Execute(strupdateamordtlbom, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        End If
                        '*************************************
                    End If
                End If
                If UCase(mstrInvType) = "JOB" And GetBOMCheckFlagValue("BomCheck_Flag") Then
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
                    If UCase(Trim(mstrInvType)) <> "REJ" Then
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
                    mP_Connection.Execute("update InvoiceStock_dtl set Doc_no = " & mInvNo & " where UNIT_CODE='" & gstrUNITID & "' and Doc_no = '" & Me.txtChallanNo.Text & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    mP_Connection.CommitTrans()
                    MsgBox("Invoice has been locked successfully with number " & mInvNo, MsgBoxStyle.Information, "eMPro")
                    CmdGrpChEnt.Enabled(1) = False
                    CmdGrpChEnt.Enabled(2) = False
                End If
                txtChallanNo.Text = CStr(mInvNo)
                txtChallanNo_Validating(txtChallanNo, New System.ComponentModel.CancelEventArgs(False))
                strSelection = "{SalesChallan_Dtl.Location_Code}='" & Trim(txtLocationCode.Text) & "' and {SalesChallan_Dtl.Doc_No} =" & Trim(txtChallanNo.Text) & ""
                strSelection = strSelection & "  and {SalesChallan_Dtl.Invoice_Type} = '" & Trim(mstrInvoiceType) & "'  and {SalesChallan_Dtl.Sub_Category} = '" & Trim(mstrInvoiceSubType) & "' and {SalesChallan_Dtl.UNIT_CODE} = '" & gstrUNITID & "'"
                '.ReplaceSelectionFormula((strSelection))
                RdAddSold.RecordSelectionFormula = strSelection
            End If
            rsItembal.ResultSetClose()
            rsItembal = Nothing
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
        Dim gobjDB As New ClsResultSetDB
        Dim rsSalesConf1 As New ADODB.Recordset
        Dim TinNo As String
        Dim blnPrintTinNo As Boolean
        gobjDB.GetResult("SELECT EOU_Flag, CustSupp_Inc,InsExc_Excise,postinfin,Excise_RoundOFF FROM sales_parameter WHERE UNIT_CODE='" & gstrUNITID & "'")
        If gobjDB.GetValue("EOU_Flag") = True Then
            mStrCustMst = "Select Doc_No,Invoice_type from SalesChallan_Dtl where Invoice_Type <> 'EXP' and Location_Code='" & Trim(txtLocationCode.Text) & "' AND UNIT_CODE='" & gstrUNITID & "'"
            mblnEOUUnit = True
        Else
            mStrCustMst = "Select Doc_No,Invoice_type from SalesChallan_Dtl where UNIT_CODE='" & gstrUNITID & "' AND Location_Code='" & Trim(txtLocationCode.Text) & "'"
            mblnEOUUnit = False
        End If
        mblnAddCustomerMaterial = gobjDB.GetValue("CustSupp_Inc")
        mblnInsuranceFlag = gobjDB.GetValue("InsExc_Excise")
        mblnpostinfin = gobjDB.GetValue("postinfin")
        mblnExciseRoundOFFFlag = gobjDB.GetValue("Excise_RoundOFF")
        rsSalesConf1.Open("SELECT * FROM SaleConf WHERE UNIT_CODE='" & gstrUNITID & "' AND Invoice_Type='" & mstrInvoiceType & "' AND Sub_Type ='" & mstrInvoiceSubType & "' AND Location_Code='" & Trim(txtLocationCode.Text) & "' and datediff(dd,getdate(),fin_start_date)<=0  and datediff(dd,fin_end_date,getdate())<=0 ", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
        If Not rsSalesConf1.EOF Then
            mstrPurposeCode = IIf(IsDBNull(rsSalesConf1.Fields("inv_GLD_prpsCode").Value), "", Trim(rsSalesConf1.Fields("inv_GLD_prpsCode").Value))
            mblnSameSeries = rsSalesConf1.Fields("Single_Series").Value
            mstrReportFilename = IIf(IsDBNull(rsSalesConf1.Fields("Report_filename").Value), "", Trim(rsSalesConf1.Fields("Report_filename").Value))
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
        rsSalesConf1.Close()
        rsSalesConf1 = Nothing
        On Error GoTo Err_Handler
        rsCompMst = New ClsResultSetDB
        strSQL = "{SalesChallan_Dtl.Location_Code}='" & Trim(txtLocationCode.Text) & "' and {SalesChallan_Dtl.Doc_No} =" & Trim(txtChallanNo.Text) & ""
        strSQL = strSQL & "  and {SalesChallan_Dtl.Invoice_Type} = '" & Trim(mstrInvoiceType) & "'  and {SalesChallan_Dtl.Sub_Category} = '" & Trim(mstrInvoiceSubType) & "' and {SalesChallan_Dtl.UNIT_CODE} = '" & gstrUNITID & "'"
        strCompMst = "Select * from Company_Mst WHERE UNIT_CODE='" & gstrUNITID & "'"
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
            TinNo = rsCompMst.GetValue("Tin_no")
        End If
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
            'mP_Connection.Execute("Truncate Table tempCustAnnex", , ADODB.ExecuteOptionEnum.adExecuteNoRecords) ' to delete all the records from table before inserting new one for selected invoice
            mP_Connection.Execute("DELETE FROM tempCustAnnex WHERE UNIT_CODE='" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            If BomCheck() = False Then
                InvoiceGenerationRPT = False
                Exit Function
            End If
        End If
        'End If
        Address = gstr_WRK_ADDRESS1 & gstr_WRK_ADDRESS2
        rsCompMst.ResultSetClose()
        rsCompMst = New ClsResultSetDB
        rsCompMst.GetResult("Select ConsigneeDetails from Sales_parameter WHERE UNIT_CODE='" & gstrUNITID & "'")
        If rsCompMst.GetValue("ConsigneeDetails") = False Then
            rsCompMst.GetResult("Select a.* from Customer_Mst a, saleschallan_dtl b where a.UNIT_CODE=b.UNIT_CODE AND a.UNIT_CODE='" & gstrUNITID & "' AND a.Customer_code = b.Account_code and b.Doc_No = " & txtChallanNo.Text & " and b.Location_Code='" & Trim(txtLocationCode.Text) & "'")
            If rsCompMst.GetNoRows > 0 Then
                DeliveredAdd = Trim(rsCompMst.GetValue("Ship_address1"))
                If Len(Trim(DeliveredAdd)) Then
                    DeliveredAdd = Trim(DeliveredAdd) & "," & Trim(rsCompMst.GetValue("Ship_address2"))
                Else
                    DeliveredAdd = Trim(rsCompMst.GetValue("Ship_address2"))
                End If
            End If
        Else
            rsCompMst.GetResult("Select ConsigneeAddress1,ConsigneeAddress2,ConsigneeAddress3 from Saleschallan_dtl where UNIT_CODE='" & gstrUNITID & "' AND Doc_No = " & txtChallanNo.Text & " and Location_Code='" & Trim(txtLocationCode.Text) & "'")
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
        End If
        rsCompMst.ResultSetClose()
        'crReportInvoicePrinting.Connect = gstrREPORTCONNECT
        'crReportInvoicePrinting.DiscardSavedData = True
        With RdAddSold
            .Load(My.Application.Info.DirectoryPath & "\Reports\" & mstrReportFilename & ".rpt")
            .RecordSelectionFormula = strSQL
            If UCase(mstrInvoiceType) <> "JOB" Then
                .DataDefinition.FormulaFields("Category").Text = "'" & mstrInvoiceType & "'"
            End If
            .DataDefinition.FormulaFields("Registration").Text = "'" & RegNo & "'"
            .DataDefinition.FormulaFields("ECC").Text = "'" & EccNo & "'"
            .DataDefinition.FormulaFields("Range").Text = "'" & Range & "'"
            .DataDefinition.FormulaFields("CompanyName").Text = "'" & gstrCOMPANY & "'"
            .DataDefinition.FormulaFields("CompanyAddress").Text = "'" & Address & "'"
            .DataDefinition.FormulaFields("Phone").Text = "'" & Phone & "'"
            .DataDefinition.FormulaFields("Fax").Text = "'" & Fax & "'"
            If UCase(mstrInvoiceType) <> "JOB" Then
                .DataDefinition.FormulaFields("EMail").Text = "'" & EMail & "'"
            End If
            .DataDefinition.FormulaFields("PLA").Text = "'" & PLA & "'"
            .DataDefinition.FormulaFields("UPST").Text = "'" & UPST & "'"
            .DataDefinition.FormulaFields("CST").Text = "'" & CST & "'"
            .DataDefinition.FormulaFields("Division").Text = "'" & Division & "'"
            .DataDefinition.FormulaFields("commissionerate").Text = "'" & Commissionerate & "'"
            .DataDefinition.FormulaFields("InvoiceRule").Text = "'" & Invoice_Rule & "'"
            .DataDefinition.FormulaFields("EOUFlag").Text = "'" & mblnEOUUnit & "'"
            If Val(txtChallanNo.Text) > 99000000 Then
                .DataDefinition.FormulaFields("DeliveredAt").Text = "' Delivered At '"
                .DataDefinition.FormulaFields("Address2").Text = "'" & DeliveredAdd & "'"
            Else
                .DataDefinition.FormulaFields("DeliveredAt").Text = "''"
                .DataDefinition.FormulaFields("Address2").Text = "''" 'to pass blanck Address in this case will overwrite this Formula written in Crystal Report for else case
            End If
            '   crReportInvoicePrinting.Formulas(16) = "EOUFlag='" & blnEOUFlag & "'"
            'crReportInvoicePrinting.Formulas(19) = "PLADuty='" & Trim(txtPLA) & "'"
            .DataDefinition.FormulaFields("InsuranceFlag").Text = "'" & mblnInsuranceFlag & "'"
            .DataDefinition.FormulaFields("StringYear").Text = "'" & Year(GetServerDate) & "'"
            .DataDefinition.FormulaFields("DateOfRemoval").Text = "'" & dtpRemoval.Text & "'"
            Dim strInvoiceDate As String
            Dim dblExistingInvNo As Double
            Dim strSql1 As String
            If Val(txtChallanNo.Text) > 99000000 Then
                .DataDefinition.FormulaFields("InvoiceNo").Text = "'" & mSaleConfNo & "'"
            Else
                strSql1 = "select * from Saleschallan_Dtl where UNIT_CODE='" & gstrUNITID & "' and Doc_No =" & Me.txtChallanNo.Text & "  and Location_Code='" & Trim(txtLocationCode.Text) & "'"
                rsSalesInvoiceDate = New ClsResultSetDB
                rsSalesInvoiceDate.GetResult(strSql1, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                strInvoiceDate = VB6.Format(rsSalesInvoiceDate.GetValue("Invoice_Date"), gstrDateFormat)
                rsSalesConf = New ClsResultSetDB
                rsSalesConf.GetResult("Select Suffix from SaleConf Where UNIT_CODE='" & gstrUNITID & "' and invoice_type ='" & mstrInvoiceType & "' AND Location_Code ='" & Trim(txtLocationCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0")
                strSuffix = rsSalesConf.GetValue("Suffix")
                If Len(Trim(strSuffix)) > 0 Then
                    If Val(strSuffix) > 0 Then
                        dblExistingInvNo = Val(Mid(CStr(txtChallanNo.Text), Len(Trim(strSuffix)) + 1))
                    Else
                        dblExistingInvNo = CDbl(txtChallanNo.Text)
                    End If
                Else
                    dblExistingInvNo = CDbl(txtChallanNo.Text)
                End If
                .DataDefinition.FormulaFields("InvoiceNo").Text = "'" & dblExistingInvNo & "'"
            End If
            blnPrintTinNo = CBool(Find_Value("Select isnull(PrintTinNO,0) as PrintTinNO from sales_parameter where UNIT_CODE='" & gstrUNITID & "' "))
            If blnPrintTinNo = True Then
                .DataDefinition.FormulaFields("TinNo").Text = "'" & TinNo & "'"
            End If
            'crReportInvoicePrinting.WindowState = Crystal.WindowStateConstants.crptMaximized
            'crReportInvoicePrinting.WindowShowPrintBtn = True
            If UCase(mstrInvoiceType) = "REJ" Then
                rsGrnHdr = New ClsResultSetDB
                strGRNDate = "" : strVendorInvDate = "" : strVendorInvNo = "" : strCustRefForGrn = ""
                rsGrnHdr.GetResult("Select Cust_ref from salesChallan_dtl where UNIT_CODE='" & gstrUNITID & "' and doc_No = " & txtChallanNo.Text)
                If rsGrnHdr.GetNoRows > 0 Then
                    rsGrnHdr.MoveFirst()
                    strCustRefForGrn = rsGrnHdr.GetValue("Cust_ref")
                End If
                rsGrnHdr.ResultSetClose()
                If Len(Trim(strCustRefForGrn)) > 0 Then
                    rsGrnHdr.GetResult("select grn_date,Invoice_no,Invoice_date from grn_hdr where UNIT_CODE='" & gstrUNITID & "' and From_Location ='01R1' and doc_No = " & strCustRefForGrn)
                    If rsGrnHdr.GetNoRows > 0 Then
                        rsGrnHdr.MoveFirst()
                        strGRNDate = IIf(IsDBNull(rsGrnHdr.GetValue("grn_date")), "", VB6.Format(rsGrnHdr.GetValue("grn_date"), gstrDateFormat))
                        strVendorInvDate = IIf(IsDBNull(rsGrnHdr.GetValue("invoice_date")), "", VB6.Format(rsGrnHdr.GetValue("invoice_date"), gstrDateFormat))
                        strVendorInvNo = rsGrnHdr.GetValue("Invoice_No")
                    End If
                End If
                .DataDefinition.FormulaFields("GrinDate").Text = "'" & strGRNDate & "'"
                .DataDefinition.FormulaFields("GrinInvoiceNo").Text = "'" & strVendorInvNo & "'"
                .DataDefinition.FormulaFields("GrinInvoiceDate").Text = "'" & strVendorInvDate & "'"
                rsGrnHdr = Nothing
            End If
            If CBool(Find_Value("select TextPrinting from sales_parameter where UNIT_CODE='" & gstrUNITID & "'")) Then
            Else
                If mstrReportFilename = "" Then
                    MsgBox("No Report filename selected for the invoice. Invoice cannot be printed", MsgBoxStyle.Information, "eMPro")
                    Exit Function
                End If
            End If
            '.Load(My.Application.Info.DirectoryPath & "\Reports\" & mstrReportFilename & ".rpt")
            'crReportInvoicePrinting.ReplaceSelectionFormula((strSQL))
        End With
        'rsCompMst.ResultSetClose()
        rsCompMst = Nothing
        InvoiceGenerationRPT = True
        Exit Function
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Sub FillDetails(ByRef ShowIemCode As Boolean, Optional ByRef SelectedItemCode As String = "")
        On Error GoTo ErrHandler
        Dim rsNagare As New ADODB.Recordset
        Dim rsCustref As New ADODB.Recordset
        Dim rsCurrencyType As ClsResultSetDB
        Dim strSQL As String
        SpChEntry.MaxRows = 0
        strSQL = "select * from MKT_enagareDTL where UNIT_CODE='" & gstrUNITID & "' and kanbanNo = '" & Trim(txtSRVDINO.Text) & "'"
        rsNagare.Open(strSQL, mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        Dim intDecimalPlace As Short
        Dim strCurrency As String
        If rsNagare.RecordCount > 0 Then
            txtCustCode.Text = IIf(IsDBNull(rsNagare.Fields("Account_Code").Value), "", rsNagare.Fields("Account_Code").Value)
            Call txtCustCode_Validating(txtCustCode, New System.ComponentModel.CancelEventArgs(False))
            If ShowIemCode Then
                strSQL = Query2GetDataFromCustOrd_Dtl(IIf(IsDBNull(rsNagare.Fields("Account_Code").Value), "", rsNagare.Fields("Account_Code").Value), (CmbInvType.Text), SelectedItemCode, IIf(IsDBNull(rsNagare.Fields("Cust_DrgNo").Value), "", rsNagare.Fields("Cust_DrgNo").Value))
            Else
                strSQL = Query2GetDataFromCustOrd_Dtl(IIf(IsDBNull(rsNagare.Fields("Account_Code").Value), "", rsNagare.Fields("Account_Code").Value), (CmbInvType.Text), IIf(IsDBNull(rsNagare.Fields("Cust_DrgNo").Value), "", rsNagare.Fields("Cust_DrgNo").Value), IIf(IsDBNull(rsNagare.Fields("Cust_DrgNo").Value), "", rsNagare.Fields("Cust_DrgNo").Value))
            End If
            rsCustref.Open(strSQL, mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If rsCustref.RecordCount > 0 Then
                If CBool(Find_Value("Select showItemInEnagareHelp  from Sales_Parameter where UNIT_CODE='" & gstrUNITID & "'")) Then
                    mstrRefNo = mstrSONo
                    txtRefNo.Text = mstrSONo
                Else
                    mstrRefNo = IIf(IsDBNull(rsCustref.Fields("Cust_Ref").Value), "", rsCustref.Fields("Cust_Ref").Value)
                    txtRefNo.Text = IIf(IsDBNull(rsCustref.Fields("Cust_Ref").Value), "", rsCustref.Fields("Cust_Ref").Value)
                End If
                Call txtRefNo_Validating(txtRefNo, New System.ComponentModel.CancelEventArgs(False))
                txtAmendNo.Text = IIf(IsDBNull(rsCustref.Fields("Amendment_no").Value), "", rsCustref.Fields("Amendment_no").Value)
                Call txtAmendNo_Validating(txtAmendNo, New System.ComponentModel.CancelEventArgs(False))
                mstrAmmNo = IIf(IsDBNull(rsCustref.Fields("Amendment_no").Value), "", rsCustref.Fields("Amendment_no").Value)
                mstrItemCode = "'" & IIf(IsDBNull(rsNagare.Fields("Cust_DrgNo").Value), "", rsNagare.Fields("Cust_DrgNo").Value) & "'"
                mstrNagareDate = VB6.Format(rsNagare.Fields("sch_date").Value, gstrDateFormat)
                If Len(mstrItemCode) > 0 Then
                    Select Case Me.CmdGrpChEnt.Mode
                        Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                            rsCurrencyType = New ClsResultSetDB
                            rsCurrencyType.GetResult("Select Currency_code from saleschallan_dtl where UNIT_CODE='" & gstrUNITID & "' and doc_No = " & Val(txtChallanNo.Text))
                            If rsCurrencyType.GetNoRows > 0 Then
                                rsCurrencyType.MoveFirst()
                                strCurrency = rsCurrencyType.GetValue("Currency_code")
                            End If
                            intDecimalPlace = ToGetDecimalPlaces(strCurrency)
                            If intDecimalPlace < 2 Then
                                intDecimalPlace = 2
                            End If
                            DisplayDetailsInSpread(strCurrency) 'Procedure Call To Select Data >From Sales_Dtl
                        Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                            Call displayDeatilsfromCustOrdHdrandDtl()
                            With SpChEntry
                                SSTab1.SelectedIndex = 1
                                .Row = 1
                                .Col = GridHeader.Quantity
                                'Query to pick remaining quantity against selected kanban no
                                'Query changed by Arshad ie. sum of sales_quantity
                                strSQL = "select isnull(c.quantity,0) - sum(isnull(b.sales_quantity,0)) - sum(isnull(p.sales_quantity,0))-sum(isnull(f.quantity,0)) as Balance"
                                strSQL = strSQL & " from mkt_enagareDtl c"
                                strSQL = strSQL & " left outer join salesChallan_dtl a on a.bill_flag= 1 and a.UNIT_CODE=c.UNIT_CODE"
                                strSQL = strSQL & " left outer join sales_dtl b on b.srvdino = c.kanbanNo and a.location_code = b.location_code and a.doc_no=b.doc_no AND b.UNIT_CODE=c.UNIT_CODE AND a.UNIT_CODE=b.UNIT_CODE"
                                strSQL = strSQL & " Left Outer join PrintedSRV_Dtl as p on c.kanbanno=p.kanban_no AND c.UNIT_CODE=p.UNIT_CODE"
                                strSQL = strSQL & " Left Outer join mkt_57F4challankanban_dtl as f on c.kanbanno=f.Kanban_no AND c.UNIT_CODE=f.UNIT_CODE "
                                strSQL = strSQL & " Inner join mkt_57F4challan_hdr as h on h.doc_type=f.doc_type and h.doc_no = f.doc_no and h.invoice_lock= 1 and h.cancel_flag = 0 AND h.UNIT_CODE=f.UNIT_CODE"
                                strSQL = strSQL & " where c.kanbanNo ='" & Trim(txtSRVDINO.Text) & "' AND c.UNIT_CODE='" & gstrUNITID & "' group by b.srvdino, c.quantity"
                                .Text = CStr(Val(Find_Value(strSQL)))
                                Call SpChEntry_Change(SpChEntry, New AxFPSpreadADO._DSpreadEvents_ChangeEvent(5, 1))
                                .Action = 0
                                .Focus()
                            End With
                            System.Windows.Forms.Application.DoEvents()
                    End Select
                    If Me.CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                        If CDbl(Trim(txtChallanNo.Text)) > 99000000 Then
                            Me.CmdGrpChEnt.Enabled(1) = True
                            Me.CmdGrpChEnt.Enabled(2) = True
                        End If
                    End If
                End If
                'Set Cell Type In Spread
                Call ChangeCellTypeStaticText()
            Else
                MsgBox("Sales Order not found.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "eMPro")
                txtRefNo.Text = ""
            End If
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Public Function Query2GetDataFromCustOrd_Dtl(ByRef pstrCustCode As String, ByRef pstrInvType As String, ByRef pstrItemCode As String, ByRef pstrDrgCode As String) As String
        '***********************************
        'To Get string to retrieve data from Cust_Ord_Dtl
        '***********************************
        On Error GoTo ErrHandler
        Dim strSelectSql As String 'Declared To Make Select Query
        If UCase(pstrInvType) = "JOBWORK INVOICE" Then
            strSelectSql = "Select b.Item_Code,b.Cust_DrgNo,b.Amendment_No,b.Cust_Ref from Cust_Ord_hdr a,Cust_Ord_Dtl b"
            strSelectSql = strSelectSql & " where b.Account_Code='" & Trim(pstrCustCode) & "' and a.Active_flag ='A'  and b.Active_flag ='A' and "
            strSelectSql = strSelectSql & " a.Account_Code = b.Account_Code and a.Cust_ref =b.Cust_ref and "
            strSelectSql = strSelectSql & " a.Amendment_No = b.amendment_No AND a.Authorized_Flag = 1 and a.PO_type in ('J') "
            strSelectSql = strSelectSql & " and a.Valid_date >'" & getDateForDB(GetServerDate()) & "' and effect_Date <= (select max(effect_date) from cust_ord_hdr where account_code = a.account_code and cust_ref = a.cust_ref and unit_code=a.unit_code and Active_flag <>'L' and effect_date <='" & getDateForDB(GetServerDate()) & "' and unit_code='" & gstrUNITID & "' )"
            strSelectSql = strSelectSql & " and b.Item_Code ='" & Trim(pstrItemCode) & "' and b.Cust_DrgNo = '" & Trim(pstrDrgCode) & "' AND a.UNIT_CODE=b.UNIT_CODE AND a.UNIT_CODE='" & gstrUNITID & "'"
            strSelectSql = strSelectSql & " order by b.Cust_Ref,b.Amendment_No,b.Cust_DrgNo,b.Item_Code "
        ElseIf UCase(pstrInvType) = "EXPORT INVOICE" Then
            strSelectSql = "Select b.Item_Code,b.Cust_DrgNo,b.Amendment_No,b.Cust_Ref from Cust_Ord_hdr a,Cust_Ord_Dtl b"
            strSelectSql = strSelectSql & " where b.Account_Code='" & Trim(pstrCustCode) & "' and a.Active_flag ='A'  and b.Active_flag ='A' and "
            strSelectSql = strSelectSql & " a.Account_Code = b.Account_Code and a.Cust_ref =b.Cust_ref and "
            strSelectSql = strSelectSql & " a.Amendment_No = b.amendment_No AND a.Authorized_Flag = 1 and a.PO_type in ('E') "
            strSelectSql = strSelectSql & " and a.Valid_date >'" & getDateForDB(GetServerDate()) & "' and effect_Date <= (select max(effect_date) from cust_ord_hdr where account_code = a.account_code and cust_ref = a.cust_ref and unit_code=a.unit_code and Active_flag <>'L' and effect_date <='" & getDateForDB(GetServerDate()) & "' and unit_code='" & gstrUNITID & "' )"
            strSelectSql = strSelectSql & " and b.Item_Code ='" & Trim(pstrItemCode) & "' and b.Cust_DrgNo = '" & Trim(pstrDrgCode) & "' AND a.UNIT_CODE=b.UNIT_CODE AND a.UNIT_CODE='" & gstrUNITID & "'"
            strSelectSql = strSelectSql & " order by b.Cust_Ref,b.Amendment_No,b.Cust_DrgNo,b.Item_Code "
        Else
            strSelectSql = "Select b.Item_Code,b.Cust_DrgNo,b.Amendment_No,b.Cust_Ref from Cust_Ord_hdr a,Cust_Ord_Dtl b"
            strSelectSql = strSelectSql & " where b.Account_Code='" & Trim(pstrCustCode) & "' and a.Active_flag ='A' and b.Active_flag ='A' and "
            strSelectSql = strSelectSql & " a.Account_Code = b.Account_Code and a.Cust_ref =b.Cust_ref and "
            strSelectSql = strSelectSql & " a.Amendment_No = b.amendment_No  AND a.Authorized_Flag = 1 and a.PO_type in ('O','S','M') "
            strSelectSql = strSelectSql & " and a.Valid_date >'" & getDateForDB(GetServerDate()) & "' and effect_Date <= (select max(effect_date) from cust_ord_hdr where UNIT_CODE=a.UNIT_CODE AND UNIT_CODE='" & gstrUNITID & "' AND account_code = a.account_code and cust_ref = a.cust_ref and Active_flag <>'L' and effect_date <='" & getDateForDB(GetServerDate()) & "' )"
            strSelectSql = strSelectSql & " and b.Item_Code ='" & Trim(pstrItemCode) & "' and b.Cust_DrgNo = '" & Trim(pstrDrgCode) & "' AND a.UNIT_CODE=b.UNIT_CODE AND a.UNIT_CODE='" & gstrUNITID & "'"
            strSelectSql = strSelectSql & " order by b.Cust_Ref,b.Amendment_No,b.Cust_DrgNo,b.Item_Code "
        End If
        Query2GetDataFromCustOrd_Dtl = strSelectSql
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Sub SetGridHeader()
        With Me.SpChEntry
            .DisplayRowHeaders = True
            If gblnGSTUnit = True Then
                .MaxCols = GridHeader.Item_Total
            Else
                .MaxCols = GridHeader.BinQty
            End If
            .Row = 0 : .Col = GridHeader.InternalPartNo : .Text = "Internal Part No."
            .Row = 0 : .Col = GridHeader.CustPartNo : .Text = "Cust.Part No."
            .Row = 0 : .Col = GridHeader.RatePerUnit : .Text = "Rate (Per Unit)"
            .Row = 0 : .Col = GridHeader.CustSuppMatPerUnit : .Text = "Cust supp. Mat (Per Unit)"
            .Row = 0 : .Col = GridHeader.Quantity : .Text = "Quantity"
            .Row = 0 : .Col = GridHeader.model : .Text = "Model"
            If blnInvoiceAgainstMultipleSO Then
                .Row = 0 : .Col = GridHeader.CustRefNo : .Text = "Ref No."
                .Row = 0 : .Col = GridHeader.AmendmentNo : .Text = "Amendment No."
                .Row = 0 : .Col = GridHeader.srvdino : .Text = "SRVDI No"
                .Row = 0 : .Col = GridHeader.SRVLocation : .Text = "SRV Location"
                .Row = 0 : .Col = GridHeader.USLOC : .Text = "US Loc"
                .Row = 0 : .Col = GridHeader.SChTime : .Text = "Sch Time"
                .Row = 0 : .Col = GridHeader.CustRefNo : .Col2 = GridHeader.SChTime
                .BlockMode = True
                .ColHidden = False
                .BlockMode = False
            Else
                .Row = 0 : .Col = GridHeader.CustRefNo : .Text = "Ref No."
                .Row = 0 : .Col = GridHeader.AmendmentNo : .Text = "Amendment No."
                .Row = 0 : .Col = GridHeader.srvdino : .Text = "SRVDI No"
                .Row = 0 : .Col = GridHeader.SRVLocation : .Text = "SRV Location"
                .Row = 0 : .Col = GridHeader.USLOC : .Text = "US Loc"
                .Row = 0 : .Col = GridHeader.SChTime : .Text = "Sch Time"
                .Row = 0 : .Col = GridHeader.CustRefNo : .Col2 = GridHeader.SChTime
                .BlockMode = True
                .ColHidden = True
                .BlockMode = False
            End If
            .Row = 0 : .Col = GridHeader.Packing : .Text = "Packing(%)"
            .Row = 0 : .Col = GridHeader.EXC : .Text = "EXC(%)"
            .Row = 0 : .Col = GridHeader.CVD : .Text = "CVD(%)"
            .Row = 0 : .Col = GridHeader.SAD : .Text = "SAD(%)"
            If Not blnEOU_FLAG Then
                .Col = GridHeader.CVD : .Col2 = GridHeader.CVD
                .BlockMode = True
                .ColHidden = True
                .BlockMode = False
                .Col = GridHeader.SAD : .Col2 = GridHeader.SAD
                .BlockMode = True
                .ColHidden = True
                .BlockMode = False
            End If
            .Row = 0 : .Col = GridHeader.OthersPerUnit : .Text = "Others (Per Unit)"
            .Row = 0 : .Col = GridHeader.FromBox : .Text = "From Box"
            .Row = 0 : .Col = GridHeader.ToBox : .Text = "To Box"
            .Row = 0 : .Col = GridHeader.CumulativeBoxes : .Text = "Cumulative Boxes" : .set_ColWidth(10, 1500)
            .Row = 0 : .Col = GridHeader.delete : .Text = "Delete"
            .Col = GridHeader.delete : .Col2 = GridHeader.delete : .BlockMode = True : .ColHidden = True : .BlockMode = False
            .Row = 0 : .Col = GridHeader.ToolCostPerUnit : .Text = "Tool Cost (Per Unit)"
            .Col = GridHeader.ToolCostPerUnit : .Col2 = GridHeader.ToolCostPerUnit : .BlockMode = True : .ColHidden = True : .BlockMode = False
            .Row = 0 : .Col = GridHeader.Rate : .Text = "Rate"
            .Col = GridHeader.Rate : .Col2 = GridHeader.Rate : .BlockMode = True : .ColHidden = True : .BlockMode = False
            .Row = 0 : .Col = GridHeader.CustMtrl : .Text = "Cust Mtrl"
            .Col = GridHeader.CustMtrl : .Col2 = GridHeader.CustMtrl : .BlockMode = True : .ColHidden = True : .BlockMode = False
            .Row = 0 : .Col = GridHeader.Others : .Text = "Others"
            .Col = GridHeader.Others : .Col2 = GridHeader.Others : .BlockMode = True : .ColHidden = True : .BlockMode = False
            .Row = 0 : .Col = GridHeader.ToolCost : .Text = "Tool Cost"
            .Col = GridHeader.ToolCost : .Col2 = GridHeader.ToolCost : .BlockMode = True : .ColHidden = True : .BlockMode = False
            .Row = 0 : .Col = GridHeader.BinQty : .Text = "Bin Quantity"
            .set_ColWidth(GridHeader.InternalPartNo, 1500)
            .set_ColWidth(GridHeader.CustPartNo, 1500)
            .set_ColWidth(GridHeader.Rate, 1200)
            .set_ColWidth(GridHeader.CustMtrl, 1200)
            .set_ColWidth(GridHeader.Quantity, 1200)
            .set_ColWidth(GridHeader.CustRefNo, 1500)
            .set_ColWidth(GridHeader.AmendmentNo, 1500)
            .set_ColWidth(GridHeader.srvdino, 1500)
            .set_ColWidth(GridHeader.SRVLocation, 1000)
            .set_ColWidth(GridHeader.USLOC, 1000)
            .set_ColWidth(GridHeader.SChTime, 1000)
            .set_ColWidth(GridHeader.Packing, 1000)
            .set_ColWidth(GridHeader.EXC, 1000)
            .set_ColWidth(GridHeader.CVD, 1000)
            .set_ColWidth(GridHeader.SAD, 1000)
            .set_ColWidth(GridHeader.Others, 1000)
            .set_ColWidth(GridHeader.FromBox, 1000)
            .set_ColWidth(GridHeader.ToBox, 1000)
            .set_ColWidth(GridHeader.CumulativeBoxes, 1000)
            .set_ColWidth(GridHeader.BinQty, 1000)

           
            .Row = 0 : .Col = GridHeader.delete : .Text = "Delete"
            .Col = GridHeader.delete : .Col2 = GridHeader.delete : .BlockMode = True : .ColHidden = True : .BlockMode = False
            .Row = 0 : .Col = GridHeader.ToolCostPerUnit : .Text = "Tool Cost (Per Unit)"
            .Col = GridHeader.ToolCostPerUnit : .Col2 = GridHeader.ToolCostPerUnit : .BlockMode = True : .ColHidden = True : .BlockMode = False
            .Row = 0 : .Col = GridHeader.Rate : .Text = "Rate"
            .Col = GridHeader.Rate : .Col2 = GridHeader.Rate : .BlockMode = True : .ColHidden = True : .BlockMode = False
            .Row = 0 : .Col = GridHeader.CustMtrl : .Text = "Cust Mtrl"
            .Col = GridHeader.CustMtrl : .Col2 = GridHeader.CustMtrl : .BlockMode = True : .ColHidden = True : .BlockMode = False
            .Row = 0 : .Col = GridHeader.Others : .Text = "Others"
            .Col = GridHeader.Others : .Col2 = GridHeader.Others : .BlockMode = True : .ColHidden = True : .BlockMode = False
            .Row = 0 : .Col = GridHeader.ToolCost : .Text = "Tool Cost"
            .Col = GridHeader.ToolCost : .Col2 = GridHeader.ToolCost : .BlockMode = True : .ColHidden = True : .BlockMode = False


            'TT#101188073 Added by Abhijit
            If gblnGSTUnit = True Then

                .Row = 0 : .Col = GridHeader.Basic_Value : .Text = "Basic Value" : .set_ColWidth(GridHeader.Basic_Value, 1200)
                .Row = 0 : .Col = GridHeader.Discount_Percent : .Text = "Discount(%)" : .set_ColWidth(GridHeader.Discount_Percent, 1200)
                .Row = 0 : .Col = GridHeader.Discount_Amt : .Text = "Discount(Val)" : .set_ColWidth(GridHeader.Discount_Amt, 1200)
                .Row = 0 : .Col = GridHeader.Accessible_Amt : .Text = "Accessible Value" : .set_ColWidth(GridHeader.Accessible_Amt, 1800)
                .Row = 0 : .Col = GridHeader.HSNSACCODE : .Text = "HSNSACCODE" : .set_ColWidth(GridHeader.HSNSACCODE, 1200)
                .Row = 0 : .Col = GridHeader.ISHSNORSAC : .Text = "ISHSNORSAC" : .set_ColWidth(GridHeader.ISHSNORSAC, 1200)
                .Row = 0 : .Col = GridHeader.CGST_Type : .Text = "CGST Type" : .set_ColWidth(GridHeader.CGST_Type, 1200)
                .Row = 0 : .Col = GridHeader.CGST_Percent : .Text = "CGST(%)" : .set_ColWidth(GridHeader.CGST_Percent, 1200)
                .Row = 0 : .Col = GridHeader.CGST_Amt : .Text = "CGST(Val)" : .set_ColWidth(GridHeader.CGST_Amt, 1200)
                .Row = 0 : .Col = GridHeader.SGST_Type : .Text = "SGST Type" : .set_ColWidth(GridHeader.SGST_Type, 1200)
                .Row = 0 : .Col = GridHeader.SGST_Percent : .Text = "SGST(%)" : .set_ColWidth(GridHeader.SGST_Percent, 1200)
                .Row = 0 : .Col = GridHeader.SGST_Amt : .Text = "SGST(Val)" : .set_ColWidth(GridHeader.SGST_Amt, 1200)
                .Row = 0 : .Col = GridHeader.IGST_Type : .Text = "IGST Type" : .set_ColWidth(GridHeader.IGST_Type, 1200)
                .Row = 0 : .Col = GridHeader.IGST_Percent : .Text = "IGST(%)" : .set_ColWidth(GridHeader.IGST_Percent, 1200)
                .Row = 0 : .Col = GridHeader.IGST_Amt : .Text = "IGST(Val)" : .set_ColWidth(GridHeader.IGST_Amt, 1200)
                .Row = 0 : .Col = GridHeader.UTGST_Type : .Text = "UTGST type" : .set_ColWidth(GridHeader.UTGST_Type, 1200)
                .Row = 0 : .Col = GridHeader.UTGST_Percent : .Text = "UTGST(%)" : .set_ColWidth(GridHeader.UTGST_Percent, 1200)
                .Row = 0 : .Col = GridHeader.UTGST_Amt : .Text = "UTGST(Val)" : .set_ColWidth(GridHeader.UTGST_Amt, 1200)

                .Row = 0 : .Col = GridHeader.CSESS_TAX_Type : .Text = "CSESS TAX Type" : .set_ColWidth(GridHeader.CSESS_TAX_Type, 1200)
                .Row = 0 : .Col = GridHeader.CSESS_TAX_Percent : .Text = "CSESS TAX(%)" : .set_ColWidth(GridHeader.CSESS_TAX_Percent, 1200)
                .Row = 0 : .Col = GridHeader.CSESS_TAX_VAL : .Text = "CSESS TAX (Val)" : .set_ColWidth(GridHeader.CSESS_TAX_VAL, 1200)

                .Row = 0 : .Col = GridHeader.Item_Total : .Text = "Item Total" : .set_ColWidth(GridHeader.Item_Total, 1200)

            End If

            'TT#101188073 Added by Abhijit

        End With
    End Sub
    Private Sub Item_Description(ByVal varRow As Short)
        '----------------------------------------------------------------------------------
        'Created By     -   Ashutosh Verma
        'Description    -   To show item description that is selected in grid.
        'Created On     -   31-08-2005
        'Arguments      -   Current Row Number
        '----------------------------------------------------------------------------------
        Dim bflag1, bFlag, bflag2 As Object
        Dim bflag3 As Boolean
        Dim varCustRef, varCustPartCode, varItemCode, varAmendNo As Object
        Dim Rs As New ClsResultSetDB
        Dim m_strSql As String
        lblCustPartDesc.Text = ""
        If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Or CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
            With SpChEntry

                varCustPartCode = Nothing
                bFlag = .GetText(2, varRow, varCustPartCode)

                varItemCode = Nothing
                bflag1 = .GetText(1, varRow, varItemCode)

                varCustRef = Nothing
                bflag2 = .GetText(6, varRow, varCustRef)
                varAmendNo = Nothing
                bflag3 = .GetText(7, varRow, varAmendNo)
            End With




            m_strSql = "Select cust_drg_desc from Cust_ord_dtl where UNIT_CODE='" & gstrUNITID & "' AND Item_code ='" & Trim(varItemCode) & "' and cust_drgNo ='" & Trim(varCustPartCode) & "' and cust_ref ='" & Trim(varCustRef) & "' and amendment_no='" & Trim(varAmendNo) & "' and account_code='" & Trim(txtCustCode.Text) & "'"

            Rs.GetResult(m_strSql)
            If Rs.GetNoRows > 0 Then

                lblCustPartDesc.Text = Rs.GetValue("cust_drg_desc")
            End If

            Rs = Nothing
        End If
    End Sub
    Private Function GetTotalDispatchForKanban(ByRef strSrvDINo As Object, ByRef strMode As Object) As Double
        '----------------------------------------------------------------------------------
        'Created By     -   Ashutosh Verma
        'Description    -   Calculate total dispatch for a Kanban number.
        'Created On     -   09-03-2006, Issue Id :17229
        'Arguments      -   Kanban No.
        '----------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim strSalesDispQty As String
        Dim strPrintedSrvDispQty As String
        Dim str57F4DispQty As String
        Dim rsItembal As ClsResultSetDB
        Dim dblSalesDispatch As Double
        Dim dblSRVDispatch As Double
        Dim dbl54F4Dispatch As Double
        Dim intRecordCount As Integer
        Dim intCount As Short

        If strMode = "ADD" Then
            strSalesDispQty = "Select isnull(sum(b.sales_quantity),0) as sales_quantity from salesChallan_dtl a inner join sales_dtl b on a.UNIT_CODE=b.UNIT_CODE AND a.location_code = b.location_code and a.doc_no=b.doc_no where  a.cancel_flag <> 1 AND a.UNIT_CODE='" & gstrUNITID & "'"

            strSalesDispQty = strSalesDispQty & " and b.srvdino = '" & Trim(strSrvDINo) & "' "
            rsItembal = New ClsResultSetDB
            rsItembal.GetResult(strSalesDispQty, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            intRecordCount = rsItembal.GetNoRows 'assign record count to integer variable
            If intRecordCount > 0 Then '          'if record found

                dblSalesDispatch = rsItembal.GetValue("sales_quantity")
            Else
                dblSalesDispatch = 0
            End If
            rsItembal.ResultSetClose()
            rsItembal = New ClsResultSetDB
            strPrintedSrvDispQty = " Select IsNull(sum(sales_quantity),0) as sales_quantity  from printedsrv_dtl p "

            strPrintedSrvDispQty = strPrintedSrvDispQty & " where p.KanBan_No = '" & Trim(strSrvDINo) & "' AND p.UNIT_CODE='" & gstrUNITID & "' "
            rsItembal.GetResult(strPrintedSrvDispQty, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            intRecordCount = rsItembal.GetNoRows 'assign record count to integer variable
            If intRecordCount > 0 Then '          'if record found

                dblSRVDispatch = rsItembal.GetValue("sales_quantity")
            Else
                dblSRVDispatch = 0
            End If
            rsItembal.ResultSetClose()
            rsItembal = New ClsResultSetDB
            str57F4DispQty = "Select isnull(Sum(quantity),0) as sales_quantity From mkt_57F4challankanban_dtl B inner join mkt_57F4challan_hdr A on B.UNIT_CODE=A.UNIT_CODE AND B.doc_type=A.doc_type and B.doc_no = A.doc_no where A.cancel_flag = 0 AND A.UNIT_CODE='" & gstrUNITID & "'"

            str57F4DispQty = str57F4DispQty & " and B.Kanban_No='" & Trim(strSrvDINo) & "' "
            rsItembal.GetResult(str57F4DispQty, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            intRecordCount = rsItembal.GetNoRows 'assign record count to integer variable
            If intRecordCount > 0 Then '          'if record found

                dbl54F4Dispatch = rsItembal.GetValue("sales_quantity")
            Else
                dbl54F4Dispatch = 0
            End If
            rsItembal.ResultSetClose()
        Else
            ''''''''''''''''''''''''
            strSalesDispQty = "Select isnull(sum(b.sales_quantity),0) as sales_quantity from salesChallan_dtl a inner join sales_dtl b on a.UNIT_CODE=b.UNIT_CODE AND a.location_code = b.location_code and a.doc_no=b.doc_no where  a.cancel_flag <> 1 and a.bill_flag=1 AND a.UNIT_CODE='" & gstrUNITID & "'"

            strSalesDispQty = strSalesDispQty & " and b.srvdino = '" & Trim(strSrvDINo) & "' "
            rsItembal = New ClsResultSetDB
            rsItembal.GetResult(strSalesDispQty, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            intRecordCount = rsItembal.GetNoRows 'assign record count to integer variable
            If intRecordCount > 0 Then '          'if record found

                dblSalesDispatch = rsItembal.GetValue("sales_quantity")
            Else
                dblSalesDispatch = 0
            End If
            rsItembal.ResultSetClose()
            rsItembal = New ClsResultSetDB
            strPrintedSrvDispQty = " Select IsNull(sum(sales_quantity),0) as sales_quantity  from printedsrv_dtl p "

            strPrintedSrvDispQty = strPrintedSrvDispQty & " where p.KanBan_No = '" & Trim(strSrvDINo) & "' AND p.UNIT_CODE='" & gstrUNITID & "' "
            rsItembal.GetResult(strPrintedSrvDispQty, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            intRecordCount = rsItembal.GetNoRows 'assign record count to integer variable
            If intRecordCount > 0 Then '          'if record found

                dblSRVDispatch = rsItembal.GetValue("sales_quantity")
            Else
                dblSRVDispatch = 0
            End If
            rsItembal.ResultSetClose()
            rsItembal = New ClsResultSetDB
            str57F4DispQty = "Select isnull(Sum(quantity),0) as sales_quantity From mkt_57F4challankanban_dtl B inner join mkt_57F4challan_hdr A on a.UNIT_CODE=b.UNIT_CODE AND B.doc_type=A.doc_type and B.doc_no = A.doc_no where A.cancel_flag = 0 AND a.UNIT_CODE='" & gstrUNITID & "'"

            str57F4DispQty = str57F4DispQty & " and B.Kanban_No='" & Trim(strSrvDINo) & "' "
            rsItembal.GetResult(str57F4DispQty, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            intRecordCount = rsItembal.GetNoRows 'assign record count to integer variable
            If intRecordCount > 0 Then '          'if record found

                dbl54F4Dispatch = rsItembal.GetValue("sales_quantity")
            Else
                dbl54F4Dispatch = 0
            End If
            rsItembal.ResultSetClose()
            '''''''''''''''''''''
        End If
        GetTotalDispatchForKanban = dblSalesDispatch + dblSRVDispatch + dbl54F4Dispatch
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Function ValidateTariffCode(ByVal strItem As String) As Boolean
        '----------------------------------------------------------------------------------
        'Created By     -   Ashutosh Verma
        'Description    -   Check whether tariff code of the item is defined or not.
        'Created On     -   10 May 2006 , Issue Id: 17410
        'Arguments      -
        '----------------------------------------------------------------------------------
        Dim rsTarriff As ClsResultSetDB
        Dim strSQL As String
        On Error GoTo ErrHandler
        strSQL = "Select Tariff_Code,item_code from item_mst where UNIT_CODE='" & gstrUNITID & "' AND item_code= '" & Trim(strItem) & "' "
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
        '----------------------------------------------------------------------------------
        'Created By     -   Ashutosh Verma
        'Description    -   Check for tariff code and Ecess.
        'Created On     -   10 May 2006 , Issue Id: 17610
        'Arguments      -
        '----------------------------------------------------------------------------------
        Dim intItem As Short
        Dim strItemList As String
        Dim blnExcisableItem As Boolean
        Dim strExciseTax As String
        Dim strECESSTax As String
        Dim rsECESSTax_Percentage As ClsResultSetDB
        Dim rsExcise_Percentage As ClsResultSetDB
        Dim dblExcisePercentage As Double
        Dim dblTemp As Double
        On Error GoTo ErrHandler

        If gblnGSTUnit Then
            ValidateTariff_CESS = True
            Exit Function
        End If

        For intItem = 1 To SpChEntry.MaxRows
            SpChEntry.Col = GridHeader.EXC : SpChEntry.Row = intItem
            strExciseTax = Trim(SpChEntry.Text)
            If Trim(strExciseTax) = "" Then
                MsgBox("Excise Tax Can't be blank for Item. Please enter Valid Excise Tax.", MsgBoxStyle.Information, "eMpro")
                ValidateTariff_CESS = False
                Exit Function
            End If
            rsExcise_Percentage = New ClsResultSetDB
            rsExcise_Percentage.GetResult("SELECT TxRt_Percentage FROM Gen_TaxRate WHERE UNIT_CODE='" & gstrUNITID & "' AND TxRt_Rate_No ='" & Trim(strExciseTax) & "' AND Tx_TaxeID='EXC'  ", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If intItem = 1 Then

                dblTemp = rsExcise_Percentage.GetValue("TxRt_Percentage")
            Else

                dblExcisePercentage = rsExcise_Percentage.GetValue("TxRt_Percentage")
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
                SpChEntry.Col = GridHeader.InternalPartNo : SpChEntry.Row = intItem
                If ValidateTariffCode(Trim(SpChEntry.Text)) = False Then
                    If Len(strItemList) = 0 Then
                        strItemList = Trim(SpChEntry.Text)
                    Else
                        strItemList = strItemList & "," & Trim(SpChEntry.Text)
                    End If
                End If
            End If
        Next intItem
        If Len(strItemList) > 1 Then
            MsgBox("Tariff Code is required for Item(s)-- " & strItemList, MsgBoxStyle.Information, "eMpro")
            ValidateTariff_CESS = False
            Exit Function
        End If
        strECESSTax = (Me.txtECSSTaxType.Text)
        If Trim(strECESSTax) = "" Then
            MsgBox("Ecess Can't be blank. Please enter Valid Ecess.", MsgBoxStyle.Information, "eMpro")
            ValidateTariff_CESS = False
            Exit Function
        End If
        rsECESSTax_Percentage = New ClsResultSetDB
        rsECESSTax_Percentage.GetResult("SELECT TxRt_Percentage FROM Gen_TaxRate WHERE UNIT_CODE='" & gstrUNITID & "' AND TxRt_Rate_No ='" & Trim(strECESSTax) & "' AND Tx_TaxeID='ECS'  ", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
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
                Me.txtECSSTaxType.Text = ""
                Me.txtECSSTaxType.Focus()
                Exit Function
            End If
        End If
        rsECESSTax_Percentage.ResultSetClose()
        ValidateTariff_CESS = True
        Exit Function
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function RefreshBoxes(ByVal Row As Short) As Boolean
        '-----------------------------------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Row no. of the Grid
        ' Return Value  : True if data is correct otherwise return false
        ' Function      : To Refresh the FromBox ToBox columns of the grid
        ' Datetime      : 31 May 2006
        '---------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim intCtr As Short
        Dim dblSaleQty As Double
        Dim dblBinQty As Double
        Dim intFromBox As Short
        Dim intBoxes As Short
        Dim varsalesqty As Object = Nothing
        Dim varBinQty As Object = Nothing
        RefreshBoxes = True
        With SpChEntry
            For intCtr = Row To .MaxRows Step 1
                .Row = intCtr
                If intCtr = 1 Then
                    intFromBox = 1
                Else
                    .Row = intCtr - 1
                    .Col = GridHeader.ToBox
                    intFromBox = Val(.Text) + 1
                End If
                .Row = intCtr
                .Col = GridHeader.FromBox
                .Text = CStr(intFromBox)

                varsalesqty = Nothing
                .GetText(GridHeader.Quantity, intCtr, varsalesqty)
                dblSaleQty = Val(varsalesqty)
                If dblSaleQty <= 0 Then
                    MsgBox("Quantity Can't be Zero", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    RefreshBoxes = False
                    Exit Function
                End If

                varBinQty = Nothing
                .GetText(GridHeader.BinQty, intCtr, varBinQty)
                dblBinQty = Val(varBinQty)
                If dblBinQty <= 0 Then
                    MsgBox("Bin Quantity Can't be Zero", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    RefreshBoxes = False
                    Exit Function
                End If
                intBoxes = Fix(dblSaleQty / dblBinQty)
                .Col = GridHeader.ToBox
                If (dblSaleQty / dblBinQty) > intBoxes Then
                    .Text = CStr(intFromBox + intBoxes)
                    intFromBox = intBoxes
                Else
                    .Text = CStr(intFromBox + intBoxes - 1)
                    intFromBox = intBoxes - 1
                End If
                .Col = GridHeader.CumulativeBoxes
                If intCtr <> 1 Then
                    .Row = intCtr - 1
                    intBoxes = Val(.Text)
                    .Row = intCtr
                    .Text = CStr(intBoxes + intFromBox + 1)
                Else
                    .Row = intCtr
                    .Text = CStr(intFromBox + 1)
                End If
            Next
        End With
        Exit Function
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function CheckSOType(ByVal Row As Short) As String
        '-----------------------------------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Row no. of the Grid
        ' Return Value  : PO_Type fields value of the Cust_Ord_Hdr Table
        ' Function      : To Check the PO_Type of the SO
        ' Datetime      : 31 May 2006
        '---------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim RSchkSoType As ClsResultSetDB
        Dim strSQL As String
        RSchkSoType = New ClsResultSetDB
        With SpChEntry
            .Row = Row
            .Col = GridHeader.CustRefNo
            strSQL = "select Po_Type from Cust_Ord_Hdr where UNIT_CODE='" & gstrUNITID & "' AND Account_code = '" & Trim(txtCustCode.Text) & "' and Cust_Ref='" & .Text & "'"
            .Col = GridHeader.AmendmentNo
            strSQL = strSQL & " and Amendment_No='" & .Text & "'"
            RSchkSoType.GetResult(strSQL, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            If RSchkSoType.GetNoRows > 0 Then

                CheckSOType = Trim(RSchkSoType.GetValue("PO_Type"))
            Else
                CheckSOType = ""
            End If
            RSchkSoType.ResultSetClose()

            RSchkSoType = Nothing
        End With
        Exit Function
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function CalculateSECSSTaxValue(ByVal pdblTotalExciseValue As Double) As Double
        '-----------------------------------------------------------------------------------
        'Created By      : Davinder Singh
        'Issue ID        : 19575
        'Creation Date   : 27 Feb 2007
        'Function        : Calculate New Tax SEcess
        '-----------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        CalculateSECSSTaxValue = (pdblTotalExciseValue * Val(lblSECSStax_Per.Text) / 100)
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub ctlInsurance_KeyPress(ByVal Sender As Object, ByVal e As UCActXCtl.UCctlFloat.KeyPressEventArgs) Handles ctlInsurance.KeyPress
        '****************************************************
        'Created By     -  Nisha
        'Description    -  At Enter Key Press Set Focus To Next Control
        '****************************************************
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
    Private Sub txtDiscountAmt_KeyPress(ByVal Sender As Object, ByVal e As UCActXCtl.UCctlFloat.KeyPressEventArgs) Handles txtDiscountAmt.KeyPress
        '****************************************************
        'Created By     -  Preety
        'Description    -  At Enter Key Press Set Focus To Next Control
        '****************************************************
        On Error GoTo ErrHandler
        Select Case e.KeyAscii
            Case System.Windows.Forms.Keys.Return
                If txtSRVDINO.Enabled Then
                    txtSRVDINO.Focus()
                Else
                    CmdGrpChEnt.Focus()
                End If
            Case 39, 34, 96
                e.KeyAscii = 0
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtFreight_KeyPress(ByVal Sender As Object, ByVal e As UCActXCtl.UCctlFloat.KeyPressEventArgs) Handles txtFreight.KeyPress
        '****************************************************
        'Created By     -  Nisha
        'Description    -  At Enter Key Press Set Focus To Next Control
        '****************************************************
        On Error GoTo ErrHandler
        Select Case e.KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        If (CmbInvType.Text = "SAMPLE INVOICE") Or (CmbInvType.Text = "JOBWORK INVOICE") Then
                            If txtSurchargeTaxType.Enabled = True Then txtSurchargeTaxType.Focus()
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
    Private Sub ctlFormHeader1_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlFormHeader1.Click
        On Error GoTo ErrHandler
        Call ShowHelp("HLPMKTTRN0005.HTM")
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
        lblDateDes.Text = dtpDateDesc.Value
    End Sub
    Private Sub dtpRemoval_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        'If dtpRemoval.value < dtpDateDesc.value Then
        If Len(lblDateDes.Text) > 0 Then
            If dtpRemoval.Value < ConvertToDate(lblDateDes.Text) Then
                dtpRemoval.Value = ConvertToDate(lblDateDes.Text)
            End If
        End If
    End Sub
    Private Sub txtDiscountAmt_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtDiscountAmt.Validating
        Dim Cancel As Boolean = e.Cancel
        '*** code added by preety
        ' to validate that discount must not be more than 100%
        Select Case CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                If OptDiscountPercentage.Checked = True And System.Math.Round(Val(txtDiscountAmt.Text)) > 100 Then
                    MsgBox("Discount cannot be Greater than value.", MsgBoxStyle.Information, "eMPro")
                    Cancel = True
                    txtDiscountAmt.Text = ""
                    txtDiscountAmt.Focus()
                Else
                    txtRemarks.Focus()
                End If
                '***
        End Select
        e.Cancel = Cancel
    End Sub
    Private Sub ctlPerValue_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ctlPerValue.KeyPress
        '****************************************************
        'Created By     -  Tapan
        'Description    -  At Enter Key Press Set Focus To Next Control
        '****************************************************
        On Error GoTo ErrHandler
        Select Case Asc(e.KeyChar)
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        With Me.SpChEntry
                            txtRemarks.Focus()
                        End With
                End Select
            Case 39, 34, 96
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub ctlPerValue_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ctlPerValue.TextChanged
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        Dim varRate As Object
        Dim varCustMtrl As Object
        Dim varOthers As Object
        Dim varToolCost As Object
        Dim varDrgNo As Object
        Dim varItemCode As Object
        With ctlPerValue
            If Val(ctlPerValue.Text) < 1 Then
                ctlPerValue.Text = 1
            End If
            With SpChEntry
                If Val(ctlPerValue.Text) > 1 Then
                    .Row = 0 : .Col = GridHeader.Rate
                    .Text = "Rate (Per " & Val(ctlPerValue.Text) & ")"
                    .Row = 0 : .Col = GridHeader.CustMtrl
                    .Text = "Cust Supp Mat. (Per " & Val(ctlPerValue.Text) & ")"
                    .Row = 0 : .Col = GridHeader.ToolCostPerUnit
                    .Text = "Tool Cost (Per " & Val(ctlPerValue.Text) & ")"
                    .Row = 0 : .Col = GridHeader.OthersPerUnit
                    .Text = "Others (Per " & Val(ctlPerValue.Text) & ")"
                    With SpChEntry
                        intMaxLoop = .MaxRows
                        For intLoopCounter = 1 To intMaxLoop
                            varDrgNo = Nothing
                            varItemCode = Nothing
                            Call .GetText(GridHeader.CustPartNo, intLoopCounter, varDrgNo)
                            Call .GetText(GridHeader.InternalPartNo, intLoopCounter, varItemCode)


                            If (Len(Trim(CStr(varDrgNo))) > 0) And (Len(Trim(CStr(varItemCode))) > 0) Then
                                varRate = Nothing
                                varCustMtrl = Nothing
                                varToolCost = Nothing
                                varOthers = Nothing
                                Call .GetText(GridHeader.Rate, intLoopCounter, varRate)
                                Call .GetText(GridHeader.CustMtrl, intLoopCounter, varCustMtrl)
                                Call .GetText(GridHeader.ToolCost, intLoopCounter, varToolCost)
                                Call .GetText(GridHeader.Others, intLoopCounter, varOthers)

                                Call .SetText(GridHeader.RatePerUnit, intLoopCounter, varRate * CDbl(ctlPerValue.Text))

                                Call .SetText(GridHeader.CustSuppMatPerUnit, intLoopCounter, Val(varCustMtrl) * CDbl(ctlPerValue.Text))

                                Call .SetText(GridHeader.ToolCostPerUnit, intLoopCounter, Val(varToolCost) * CDbl(ctlPerValue.Text))

                                Call .SetText(GridHeader.OthersPerUnit, intLoopCounter, Val(varOthers) * CDbl(ctlPerValue.Text))
                            End If
                        Next
                    End With
                Else
                    .Row = 0 : .Col = GridHeader.Rate : .Text = "Rate (Per Unit)"
                    .Row = 0 : .Col = GridHeader.CustSuppMatPerUnit : .Text = "Cust Supp Mat. (Per Unit)"
                    .Row = 0 : .Col = GridHeader.ToolCostPerUnit : .Text = "Tool Cost (Per Unit)"
                    .Row = 0 : .Col = GridHeader.OthersPerUnit : .Text = "Others (Per Unit)"
                    With SpChEntry
                        intMaxLoop = .MaxRows
                        For intLoopCounter = 1 To intMaxLoop
                            varDrgNo = Nothing
                            varItemCode = Nothing
                            Call .GetText(GridHeader.CustPartNo, intLoopCounter, varDrgNo)
                            Call .GetText(GridHeader.InternalPartNo, intLoopCounter, varItemCode)


                            If (Len(Trim(CStr(varDrgNo))) > 0) And (Len(Trim(CStr(varItemCode))) > 0) Then
                                varRate = Nothing
                                varCustMtrl = Nothing
                                varToolCost = Nothing
                                varOthers = Nothing
                                varRate = Nothing
                                Call .GetText(GridHeader.Rate, intLoopCounter, varRate)
                                Call .GetText(GridHeader.CustMtrl, intLoopCounter, varCustMtrl)
                                Call .GetText(GridHeader.ToolCost, intLoopCounter, varToolCost)
                                Call .GetText(GridHeader.Others, intLoopCounter, varOthers)
                                Call .SetText(GridHeader.RatePerUnit, intLoopCounter, varRate)

                                Call .SetText(GridHeader.CustSuppMatPerUnit, intLoopCounter, Val(varCustMtrl))

                                Call .SetText(GridHeader.ToolCostPerUnit, intLoopCounter, Val(varToolCost))

                                Call .SetText(GridHeader.OthersPerUnit, intLoopCounter, Val(varOthers))
                            End If
                        Next
                    End With
                End If
            End With
        End With
    End Sub
    Private Sub lblCurrencyDes_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblCurrencyDes.TextChanged
        If Trim(lblCurrencyDes.Text) <> "" Then
            If Trim(lblCurrencyDes.Text) = Trim(gstrCURRENCYCODE) Then
                lblExchangeRateValue.Text = CStr(1.0#)
            Else
                If UCase(Trim(mstrInvoiceType)) = "INV" Or UCase(Trim(mstrInvoiceType)) = "SMP" Or UCase(Trim(mstrInvoiceType)) = "TRF" Or UCase(Trim(mstrInvoiceType)) = "JOB" Or UCase(Trim(mstrInvoiceType)) = "EXP" Or UCase(Trim(mstrInvoiceType)) = "SRC" Then
                    lblExchangeRateValue.Text = CStr(GetExchangeRate(lblCurrencyDes.Text, dtpDateDesc.Value, True))
                Else
                    lblExchangeRateValue.Text = CStr(GetExchangeRate(lblCurrencyDes.Text, dtpDateDesc.Value, False))
                End If
                If Val(Trim(lblExchangeRateValue.Text)) = 1 Then
                    MsgBox("Exchange Rate for " & Trim(lblCurrencyDes.Text) & " is not defined on " & dtpDateDesc.Value, MsgBoxStyle.Information, "eMPro")
                    lblExchangeRateValue.Text = ""
                End If
            End If
        Else
            lblExchangeRateValue.Text = ""
        End If
    End Sub
    Public Function Check_GST(ByVal gstrUNITID As String) As Boolean 'by abhijit
        On Error GoTo Err_Handler
        Dim rsGBL_GST As ClsResultSetDB
        rsGBL_GST = New ClsResultSetDB

        rsGBL_GST.GetResult("SELECT TOP 1 1 FROM gen_unitmaster WHERE UNT_CODEID='" & gstrUNITID & "' AND GST_Unit=1")
        If rsGBL_GST.GetNoRows > 0 Then
            gblnGSTUnit = True
        Else
            gblnGSTUnit = False
        End If
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Function
    End Function
    Public Sub EnableDisable_GST()
        Try
            If gblnGSTUnit Then
                CmdSECSSTaxType.Enabled = False
                lblSaleTaxType.Enabled = False
                lblSurchargeCode.Enabled = False
                lblLoadingCgargesType.Enabled = False
                lblECSSCode.Enabled = False
                'lblTCSTaxCode.Enabled = False
                txtSaleTaxType.Enabled = False
                txtSurchargeTaxType.Enabled = False
                txtLoadingTaxType.Enabled = False
                txtECSSTaxType.Enabled = False
                'txtTCSTaxCode.Enabled = False
                txtSECSSTaxType.Enabled = False
                CmdSaleTaxType.Enabled = False
                cmdSurchargeTaxCode.Enabled = False
                cmdLoadinfChageHelp.Enabled = False
                CmdECSSTaxType.Enabled = False
                'cmdHelpTCSTax.Enabled = False
                lblSalesTax.Enabled = False
                lblSurcharge.Enabled = False
                lblLoadingCharge.Enabled = False
                lblECSSPer.Enabled = False
                lblTCSPer.Enabled = False
                Label5.Enabled = False
                lblSaltax_Per.Enabled = False
                lblSurcharge_Per.Enabled = False
                lblLoadingcharge_per.Enabled = False
                lblECSStax_Per.Enabled = False
                lblTCSTaxPerDes.Enabled = False
                lblSECSStax_Per.Enabled = False

                fraDiscountType.Enabled = False
                OptDiscountValue.Checked = True
                OptDiscountPercentage.Checked = False
                txtDiscountAmt.Text = 0.0
                txtDiscountAmt.Enabled = False
                lblDiscountAmt.Enabled = False
                chkExciseExumpted.Enabled = False

            End If
        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Private Sub WIPFG_Screen_Setting()

        Dim intRow As Integer
        Try
            If gBlnWIPFGProcess AndAlso WIP_FG_Customer(txtCustCode.Text.Trim) Then
                With SpChEntry
                    For intRow = 1 To .MaxRows
                        .Row = intRow
                        .Col = GridHeader.InternalPartNo
                        If WIP_FG_ITEM(.Text.Trim) Then
                            .BlockMode = True
                            .Col = GridHeader.Quantity
                            .Col2 = GridHeader.Quantity
                            .Row = intRow
                            .Row2 = intRow
                            .Lock = True
                            .BlockMode = False
                        End If
                    Next
                End With
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Function HSNCODE_GSTTYPE_CHECK_01() As Boolean
        HSNCODE_GSTTYPE_CHECK_01 = True
        Dim intMaxLoop As Integer
        Dim intLoopCounter As Integer
        Dim varDrgNo As Object = Nothing
        Dim varItemCode As Object = Nothing
        Dim sqlHSNCODE_TYPE As String = ""
        Dim rsHSNCODE_TYPE As ClsResultSetDB

        Dim sqlGSTTYPE As String = ""
        Dim rsGSTTYPE As ClsResultSetDB

        Try
            With SpChEntry
                intMaxLoop = .MaxRows
                For intLoopCounter = 1 To intMaxLoop
                    varDrgNo = Nothing
                    varItemCode = Nothing
                    Call .GetText(GridHeader.CustPartNo, intLoopCounter, varDrgNo)
                    Call .GetText(GridHeader.InternalPartNo, intLoopCounter, varItemCode)
                    rsHSNCODE_TYPE = New ClsResultSetDB
                    sqlHSNCODE_TYPE = "select item_code from Item_mst WHERE UNIT_Code='" & gstrUNITID & "' and item_code='" & varItemCode & "' and (len(HSN_SAC_CODE)=0 or HSN_SAC_CODE IS NULL)"
                    rsHSNCODE_TYPE.GetResult(sqlHSNCODE_TYPE)

                    If rsHSNCODE_TYPE.GetNoRows > 0 Then
                        HSNCODE_GSTTYPE_CHECK_01 = False
                        MsgBox("HSNSAC CODE is not Available for Item ('" & varItemCode & "')", MsgBoxStyle.Information)
                        Exit Function
                    End If

                    Dim GST_COLUMNS_AV As ClsResultSetDB
                    Dim HSNSAC_CODE_TYPE_AV As String = ""
                    Dim ISHSNORSAC_AV As String = ""
                    Dim CGST_TXRT_HEAD_AV As String = ""
                    Dim SGST_TXRT_HEAD_AV As String = ""
                    Dim UGST_TXRT_HEAD_AV As String = ""
                    Dim IGST_TXRT_HEAD_AV As String = ""
                    Dim COMPENSATION_CESS_AV As String = ""
                    Dim strSql_GST_AV As String = ""

                    GST_COLUMNS_AV = New ClsResultSetDB
                    strSql_GST_AV = "select * from dbo.UFN_GST_ITEMWISETAXES('" & gstrUNITID & "','" & Trim(Me.txtCustCode.Text) & "','" & varItemCode & "',getdate(),getdate())"
                    GST_COLUMNS_AV.GetResult(strSql_GST_AV)
                    If GST_COLUMNS_AV.GetNoRows > 0 Then
                        HSNSAC_CODE_TYPE_AV = GST_COLUMNS_AV.GetValue("HSNSACCODE")
                        ISHSNORSAC_AV = GST_COLUMNS_AV.GetValue("ISHSNORSAC")
                        CGST_TXRT_HEAD_AV = GST_COLUMNS_AV.GetValue("CGST_TXRT_HEAD")
                        SGST_TXRT_HEAD_AV = GST_COLUMNS_AV.GetValue("SGST_TXRT_HEAD")
                        UGST_TXRT_HEAD_AV = GST_COLUMNS_AV.GetValue("UGST_TXRT_HEAD")
                        IGST_TXRT_HEAD_AV = GST_COLUMNS_AV.GetValue("IGST_TXRT_HEAD")
                        COMPENSATION_CESS_AV = GST_COLUMNS_AV.GetValue("COMPENSATION_CESS")
                        If Len(CGST_TXRT_HEAD_AV) = 0 And Len(SGST_TXRT_HEAD_AV) = 0 And Len(UGST_TXRT_HEAD_AV) = 0 And Len(IGST_TXRT_HEAD_AV) = 0 Then
                            HSNCODE_GSTTYPE_CHECK_01 = False
                            MsgBox("GST Types can't be blank", MsgBoxStyle.Information)
                            Exit Function
                        End If
                    Else
                        HSNCODE_GSTTYPE_CHECK_01 = False
                        MsgBox("GST Types not found", MsgBoxStyle.Information)
                        Exit Function
                    End If
                    '========================================================
                    ''ABHIJIT================================================
                Next
            End With
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Private Function HSNCODE_GSTTYPE_CHECK_02() As Boolean
        Dim result As Boolean = True
        Dim cgst As String = String.Empty
        Dim sgst As String = String.Empty
        Dim igst As String = String.Empty
        Dim utgst As String = String.Empty
        Dim hsnCode As String = String.Empty
        With SpChEntry
            For i As Integer = 1 To .MaxRows
                .Row = i
                .Col = GridHeader.HSNSACCODE
                hsnCode = .Text
                .Col = GridHeader.CGST_Type
                cgst = .Text
                .Col = GridHeader.SGST_Type
                sgst = .Text
                .Col = GridHeader.IGST_Type
                igst = .Text
                .Col = GridHeader.UTGST_Type
                utgst = .Text
                If Len(Trim(hsnCode)) = 0 Then
                    MsgBox("HSN/SAC Codes can't be blank", MsgBoxStyle.Information, "eMPro")
                    result = False
                    .Row = i
                    .Col = GridHeader.Quantity
                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                    .Focus()
                    Exit For
                End If
                If Len(Trim(cgst & sgst & igst & utgst)) = 0 Then
                    MsgBox("GST Types can't be blank", MsgBoxStyle.Information, "eMPro")
                    result = False
                    .Row = i
                    .Col = GridHeader.Quantity
                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                    .Focus()
                    Exit For
                End If
            Next
        End With
        Return result
    End Function
    Public Sub checktcsvalue(ByVal pInvType As String, ByVal pInvSubType As String)
        Dim rsTCSReq As ClsResultSetDB
        Try
            rsTCSReq = New ClsResultSetDB
            rsTCSReq.GetResult("Select isnull(REQD_TCS,0) as REQD_TCS , TCSTXRT_TYPE from saleConf Where UNIT_CODE='" + gstrUNITID + "' AND description ='" & Trim(pInvType) & "' and Sub_Type_Description='" & Trim(pInvSubType) & "' and  (fin_start_date <= getdate() and fin_end_date >= getdate())")
            If rsTCSReq.GetValue("REQD_TCS") = True Then
                lblTCSTaxCode.Enabled = True : txtTCSTaxCode.Enabled = True : txtTCSTaxCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : cmdHelpTCSTax.Enabled = True : txtTCSTaxCode.Text = rsTCSReq.GetValue("TCSTXRT_TYPE").ToString
                If CheckExistanceOfFieldData((txtTCSTaxCode.Text), "TxRt_Rate_No", "Gen_TaxRate", " (Tx_TaxeID='TCS')  and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))") Then
                    lblTCSTaxPerDes.Text = CStr(GetTaxRate((txtTCSTaxCode.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " (Tx_TaxeID='TCS')"))
                End If
            Else
                If (UCase(Trim(pInvType) = "NORMAL INVOICE") And (UCase(Trim(pInvSubType)) = "SCRAP")) Then
                    If gblnGSTUnit = False Then
                        lblTCSTaxCode.Enabled = True : txtTCSTaxCode.Enabled = True : txtTCSTaxCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : cmdHelpTCSTax.Enabled = True
                    End If
                Else
                    lblTCSTaxCode.Enabled = False : txtTCSTaxCode.Enabled = False : txtTCSTaxCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdHelpTCSTax.Enabled = False : txtTCSTaxCode.Text = ""
                End If

            End If
            rsTCSReq.ResultSetClose()

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

End Class