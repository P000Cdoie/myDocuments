Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports System.Data.SqlClient
Friend Class frmMKTTRN0009_SOUTH
	Inherits System.Windows.Forms.Form
	'===================================================================================
	' (c) MothersonSumi INfotech & Design Ltd. All rights reserverd.
    ' File Name         :   FRMMKTTRN0009_south.frm
	' Function          :   Used to add sale deatails
	' Created By        :   Nisha & Kapil
	' Created On        :   16 May, 2001
	' Revision History  :   Nisha Rai
	'21/09/2001 MARKED CHECKED BY BCs changed on version 3
	'03/11/2001 MARKED CHECKED BY BCs  for jobwork invoice changed on version 7
	'09/11/2001  changed on version 9 for schedule Status
	'09/01/2002 changed fof Smiel Chennei to add CVD_PER,SVD_Per,Insurance
	'25/01/2002 changed for decimal 4 places on Chacked Out Form No = 4019
	'28/01/2002 changed for decimal 4 places on Chacked Out Form No = 4033
	'in ChangeCellTypeStaticText()
	'16/01/2002 CHANGED FOR DOCUMENT NO. ON FORM NO. 4066
	'19/04/2002 changed on  for Tariff & TDA Changes
	'22/04/2002 changed for box quantity from ITem_Mst
	'23/04/2002 BOM structure changed
	'30/04/2002 to change on the basis of currency then decimal places
	'30/04/2002 for replacing Mod Function
	'08/05/2002 SCRAP invoice Changes
	'29/05/2002 schedule check
	'03/06/2002 for changes in refresh form to set list index to -1
	'04/06/2002 for enabling all text feilds in Rejection invoice
	'12/06/2002 for from s box size changes in Quantity Check variable type int to double
	'14/06/2002 CHANGE IN BOMCHECK FUNCTION
	'18/06/2002 change label in Grim From Drawing No to Cust Part No & to Show Packing
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
	'CHANGED ON 16/07/2002 FOR EXPORT OPTION ADDING AND CALCULATION SAME AS NORMAL INVOICE
	'23/07/2002 changed to add Grin Linking in Rejection Invoice
	'07/08/2002 changed for Jobwork invoice to check Customer supplied from Vendor Bom
	'Changed by nisha on 26/08/2002
	'changes done by nisha to check SO Qty in Challan Entry
	'CHANGES DONE BY NISHA ON 14/03/2003
	'1.FOR FINAL MERGING & FOR FROM BOX & TO bOX UPDATION WHILE EDITING INVOICE
	'2.For Grin Cancellation flag
	'3.SAMPLE INVOICE TOOL COST COLUMN
	'4.CUNSUMABLES & MISC. SALE IN CASE OF NORMAL RAW MATERIAL INVOICE
	'16/03/2003 added per value on form
	'04/07/2003 changes done by nisha on 04/07/2003 for adding cust Part Code Description.
	'changes done by nisha on 18/07/2003 for sales tax calculation1
	'===================================================================================
	'Code Changed By Arul on 30/03/2004
	'Purpose : To made the invoices for Customer supplied Material Items while Rejected at GRIN and Line Rejection
	'===================================================================================
	'Code added by arul on 27-05-2004 in Customer supplied material execise value calculation function
	'Purpose : Customer supplied excise value calculated all materials that was declared in Cusomer supplied master
	'Bcoz the IF condition had been added in the part
	'==================================================================================
	'Code Changed by Arul on 02/06/2004
	'Purpose : To avoid the BOF and Eof Error in Customer supplied excise value calculatation Function
	'==================================================================================
	'===============================================
	'08/07/2004 By Arshad Ali implemented by Sourabh in Chennai
	'ECESS Tax Type field added
	'ECESS is to be calculated on total excise value
	'when calculating Sale Tax ECESS Amount to be considered along with basic value, excise value etc.
	'===============================================
	'Code Changed By Arul on 09-12-2004
	'Rounding Off Function Added in Sales Tax and Surcharge Amounts
	'Because User getting "Cr amount not equal to Dr Amount" Message while Posting Invoice
	'===================================================================
	'Code Changed By Arul on 02-13-2004
	'Code Changed To Subtract the Customer Supplied Material Ecess Amount from the Total Invoice Amount
	'========================================================================
	'Code Changed by Arul on 20-01-2005
	'Reason : To change the lintQty Variable Datatype As Double insteade of Long
	'---------------------------------------------------------------------
	'Code Changed By Arul on 07-03-2005
	'Reason : Multiple selection of sales Order in one Invoice Option provided to users
	'Emp_InvoiceSOLinkage table used to save the multiple Sales Order for a single invoice
	'--------------------------------------------------------------------------------------
	'Revised By     : Arul mozhi
	'Revised On     : 20-04-2005
	'Reason         : Carrier Name & Vechile No set as the mendatory field
	'--------------------------------------------------------------------------------------
	'Revised By     : Arul mozhi
	'Revised On     : 13-05-2005
	'Reason         : SalesTax_Onerupee_Roundoff Flag introduced on Sales Parameter Table
	'                 To Save the sales tax value by One Rupee If the Value between 0.1 and 1.0 Rupees
	'--------------------------------------------------------------------------------------
	'Revised By     : Arul mozhi
	'Revised On     : 31-08-2005
	'Reason         : Goup company customer names only appear in Customer search button
	'                 For transfer invoice selection
	'--------------------------------------------------------------------------------------
	'Revised By     : Arul mozhi
	'Revised On     : 25-12-2005
	'Reason         : More than one user not able to make invoice at a time. Because earlier the
	'                 new invoice no has papulated in the new button press event.
	'                 Now it is changed in the save button press event.
	'--------------------------------------------------------------------------------------
	'Revised By     : Rajani kant
	'Revised On     : 17-13-2005
	'Reason         : To avail the sales order functionality for Transfer invoice
	'--------------------------------------------------------------------------------------
	'Revised By     : Arul mozhi varman
	'Revised On     : 02-06-2005
	'Reason         : Invisible the invoice date input dtpicker on New mode
	'--------------------------------------------------------------------------------------
	'Revised By      : Manoj Kr. Vaish
	'Issue ID        : 22303
	'Revision Date   : 04 Feb 2008
	'History         : Addition of bar Code Functionaltiy for Chennai
	'***********************************************************************************
	'Revised By     : Manoj Kr. Vaish
	'Revised On     : 11 Feb 2008
	'Issue ID       : 21942
	'Reason         : Higher Education Cess is not being calculated on CSM (Customer Supplied Material)
	'--------------------------------------------------------------------------------------
    'Revised By     : Rajeev Gupta
    'Revised On     : 02 Jul 2008
    'Issue ID       : eMpro-20080702-20009
    'Reason         : Cess on ED, Sale Tax Code, SE. Cess Code Should comes default on all invoices
    '                 For Cess on ED: If it is Defined in SO (Cust_ord_hdr) then Cess on ED = SO's Cess on ED
    '                 If it not defined in SO then it comes from Gen_TaxRate where DEFAULT_FOR_INVOICE = 1
    '                 
    '                 For Sale Tax Code: it comes from Gen_TaxRate where DEFAULT_FOR_INVOICE = 1
    '
    '                 For SE Cess Code: it comes from Gen_TaxRate where DEFAULT_FOR_INVOICE = 1
    '--------------------------------------------------------------------------------------
    'Revised By      : Manoj Kr.Vaish
    'Issue ID        : eMpro-20090209-27201
    'Revision Date   : 09 Feb 2009
    'History         : BatchWise Tracking of Invoices Made from 01M1 Location including BarCode Tracking
    '*******************************************************************************************************
    'Revised By      : Manoj Kr.Vaish
    'Issue ID        : eMpro-20090513-31282
    'Revision Date   : 13 May 2009
    'History         : Intergeration of Ford ASN File Generation for Mate South Units
    '*******************************************************************************************************
    'Revised By      : Manoj Kr. Vaish
    'Issue ID        : eMpro-20090624-32847
    'Revision Date   : 24 Jun 2009
    'History         : Plant Code automatically filled on selection of Customer Code
    '                : Check for Ford ASN Generation only for Normal-Finished Good
    '****************************************************************************************
    'Revised By      : Manoj Kr. Vaish
    'Issue ID        : eMpro-20090703-33213
    'Revision Date   : 03 Jul 2009
    'History         : Credit term should be fetched from Sales Order while locking the invoice
    '****************************************************************************************
    'Revised By        : Siddharth Ranjan
    'Issue ID          : eMpro-20090709-33428
    'Revision Date     : 15 Jul 2009
    'History           : CSI functionality
    '****************************************************************************************
    'Revised By        : Siddharth Ranjan
    'Issue ID          : eMpro-20100108-40881
    'Revision Date     : 09 Dec 2009
    'History           : New CSM FIFO KnockedOff functionality
    '****************************************************************************************
    'Revised By        : Prashant Rajpal
    'Issue ID          : 1096797
    'Revision Date     : 20 May 2011
    'History           : After Locking the INVOICE, EDIT/UPDATE NOT POSSIBLE 
    '****************************************************************************************
    '=====================================================================
    'Revised By     : Amit Kumar (0670)
    'Revision Date  : 30 May 2011
    'Remarks        : Changes Done To Support Multiunit Function
    '=====================================================================
    'Revised By        : Prashant Rajpal
    'Issue ID          : 10118036
    'Revision Date     : 02 Aug 2011
    'History           : CSM Functionality 
    '****************************************************************************************
    'Revised By        : Prashant Rajpal
    'Issue ID          : 10126648 
    'Revision Date     : 30 Aug 2011
    'History           : ASN functionality for Rejection Invoice 
    '----------------------------------------------------------------
    ' Revised By     :   Pankaj Kumar
    ' Revision Date  :   14 Oct 2011
    ' Description    :   Modified for MultiUnit Change Management
    '****************************************************************************************
    '***********************************************************************************
    'Revised By      : PRASHANT RAJPAL
    'Issue ID        : 10192547 
    'Revision Date   : 08 FEB 2012
    'History         : Changes in Invoice Entry FOR barcode process (At Main Store )
    '***********************************************************************************
    '-- Modified by Roshan Singh on 14 FEB 2012 for multiunit change
    '***********************************************************************************
    'Revised By      : PRASHANT RAJPAL
    'Issue ID        : 10239927  
    'Revision Date   : 20 june 2012
    'History         : Changes in Invoice Entry FOR Freight and insurance amount not able to type done
    '***********************************************************************************
    'Revised By      : PRASHANT RAJPAL
    'Issue ID        : 10178530     
    'Revision Date   : 24 july 2012
    'History         : Changes in Invoice Entry :-LRN concept for REJECTION INVOICE
    '***********************************************************************************
    'Revised By      : PRASHANT RAJPAL
    'Issue ID        : 10245698       
    'Revision Date   : 16 Nov 2012-18 Nov 2012
    'History         : Changes in Invoice Entry :-ASN for REJECTION INVOICE
    '***********************************************************************************
    'Revised By      : PRASHANT RAJPAL
    'Issue ID        : 10404992   
    'Revision Date   : 12-June-2013-13-June-2013
    'History         : Including TCS Value in Normal Invoice with sub type SCRAP
    '***********************************************************************************
    'Revised By      : PRASHANT RAJPAL
    'Issue ID        : 10414415    
    'Revision Date   : 02-July-2013- 03-July-2013
    'History         : Customer Part Description not shown in Invoice entry form , Resolved
    '****************************************************************************************
    'Revised By      : PRASHANT RAJPAL
    'Issue ID        : 10229989
    'Revision Date   : 10-aug -2013- 31-aug -2013
    'History         : Multiple So Functionlity 
    '****************************************************************************************
    'REVISED BY     :  VINOD SINGH
    'REVISED DATE   :  30 AUG 2013
    'ISSUE ID       :  10378778
    'PURPOSE        :  GLOBAL TOOL CHAGES
    '****************************************************************************************
    'REVISED BY     :  PRASHANT RAJPAL
    'REVISED DATE   :  20-NOV-2014 - 21-NOV-2014
    'ISSUE ID       :  10706455  
    'PURPOSE        :  TO ADD ADDITIONAL VAT CALCULATION
    '****************************************************************************************
    'REVISED BY     :  PRASHANT RAJPAL
    'REVISED DATE   :  12-JAN-2015
    'ISSUE ID       :  10736222
    'PURPOSE        :  TO INTEGRATE CT2 AR3 FUNCTIONALITY 
    '****************************************************************************************
    'REVISED BY     :  PRASHANT RAJPAL
    'REVISED DATE   :  16-FEB-2015
    'ISSUE ID       :  10727107
    'PURPOSE        :  TO ADD ADDITIONAL EXCISE CALCULATION 
    '****************************************************************************************
    'REVISED BY     :  ABHINAV KUMAR
    'REVISED DATE   :  08-MAY-2015
    'ISSUE ID       :  10778579 
    'PURPOSE        :  REG: PRPC PENDING SHIPMENT
    '****************************************************************************************
    'REVISED BY     :  PRASHANT RAJPAL
    'REVISED DATE   :  23-JUN-2015
    'ISSUE ID       :  10808160 
    'PURPOSE        :  EOP FUNCTIONALITY
    '****************************************************************************************
    'REVISED BY     -  PRASHANT RAJPAL
    'REVISED ON     -  21 JULY 2015
    'PURPOSE        -  10856126 -ASN DOCK CODE FUNCTIONALITY
    '-------------------------------------------------------------------------------------------------------------------------------------------------------
    'REVISED BY     -  PRASHANT RAJPAL
    'REVISED ON     -  18 SEP 2015
    'PURPOSE        -  10869290 -SERVICE INVOICE 
    '-------------------------------------------------------------------------------------------------------------------------------------------------------
    'REVISED BY     -  PRASHANT RAJPAL
    'REVISED ON     -  09 NOV 2015
    'PURPOSE        -  10899126 -ADDVAT TAX INCLUSED IN TCS TAX CALCULATION  IN INVOICE 
    '-------------------------------------------------------------------------------------------------------------------------------------------------------
    'REVISED BY     -  ASHISH SHARMA
    'REVISED ON     -  06 SEP 2017
    'PURPOSE        -  101254587 - Global Tool Master and Tool Master Enhancement Phase-II
    '-------------------------------------------------------------------------------------------------------------------------------------------------------
    'REVISED BY     -  PRASHANT RAJPAL
    'REVISED ON     -  04 OCT 2016
    'PURPOSE        -  10869291 — eMPro- Inter-Division Invoice 
    '-------------------------------------------------------------------------------------------------------------------------------------------------------
    'REVISED BY     -  ASHISH SHARMA
    'REVISED ON     -  10 SEP 2018
    'PURPOSE        -  101375632 - REG Bar code implementation - BM1 UNIT
    '-------------------------------------------------------------------------------------------------------------------------------------------------------

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
    Dim strupSalechallanUpload As String
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
    Dim mstrCusRefno As String
    '-------------added by Arul Mozhi----------------------
    'Added On 30-03-2004
    Dim BLNREJECTION_FLAG As Boolean 'To store the rejection flag (Customer or Vendor)
    '--------------ends here---------------------------------
    'Code Added by Arul on 13-05-2005
    Dim BlnSalesTax_Onerupee_Roundoff As Boolean
    'Addition Ends here
    Dim mstrupdateBarBondedStockFlag As String
    Dim mstrupdateBarBondedStockQty As String
    Dim mblnQuantityCheck As Boolean
    Dim mstrFGDomestic As String
    Dim mInvNo As String
    Dim mblnRejTracking As Boolean
    Dim mstrCompileDocDetails As String
    Public mstrCompileBatchDetails As String
    Dim mblnlinelevelcustomer As Boolean
    Dim mstrEOP_Required As Boolean
    Dim dtPaletteItemQty As DataTable
    Dim mblnAllowTransporterfromMaster As Boolean = False

    Private Enum enumExciseType
        RETURN_EXCISE = 1
        RETURN_CVD = 2
        RETURN_SAD = 3
    End Enum
    Private Structure Batch_Detailsrejection
        Dim Document_No() As String
        Dim Batch_No() As String
        Dim Batch_Date() As Date
        Dim Batch_Quantity() As Double
    End Structure
    Dim mblnBatchTracking As Boolean
    Dim mstrInvsubTypeDesc As String
    Dim mblnbatchfifomode As Boolean
    Dim mblnBatchTrack As Boolean
    Dim mstrBatchData As String
    Dim mstrLocationCode As String
    Dim mblnCSM_Knockingoff_req As Boolean
    Dim mBatchData() As Batch_Details 'Variable To Store Batch Wise Stock
    Dim mBatchDataRejction() As Batch_Detailsrejection 'Variable To Store Batch Wise Stock
    Dim mblnSinglelinelevelso As Boolean
    Dim mblnATN_invoicewise As Boolean
    Dim mstrcurrentATNCode As String = String.Empty
    Dim mstrcustcompcode As String = String.Empty
    Dim mstrexportsotype As String = String.Empty


    Private Sub CmbInvSubType_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmbInvSubType.SelectedIndexChanged
        On Error GoTo ErrHandler
        Call SelectInvTypeSubTypeFromSaleConf((CmbInvType.Text), (CmbInvSubType.Text))
        Call CheckBatchTrackingAllowed(CmbInvType.Text, CmbInvSubType.Text)
        Call PaletteActiveInActive()
        'Call checktcsvalue(CmbInvType.Text, CmbInvSubType.Text)

        'added by priti
        If Len(Trim(txtCustCode.Text)) > 0 Then
            Dim strsql As String
            Dim blntcscheck As Boolean = False
            strsql = "select dbo.UDF_IRN_TCSREQUIRED( '" & gstrUNITID & "','" & txtCustCode.Text.Trim & "')"
            If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strsql)) = True Then
                blntcscheck = True
            Else
                blntcscheck = False
            End If
            If blntcscheck = True Then
                Call checktcsvalue(CmbInvType.Text, CmbInvSubType.Text)
            End If
        End If
        '' code ends
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
                'txtLocationCode.SetFocus
                If dtpDateDesc.Enabled Then dtpDateDesc.Focus()
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

        If UCase(CmbInvType.Text) = "NORMAL INVOICE" Then
            If UCase(CmbInvSubType.Text) = "SCRAP" Then
                txtRefNo.Enabled = False : txtRefNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : CmdRefNoHelp.Enabled = False : txtRefNo.Text = ""
                txtAmendNo.Enabled = False : txtAmendNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : txtAmendNo.Text = ""
                ctlPerValue.Enabled = True
                ctlPerValue.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                'If gblnGSTUnit = False Then 'commented on 21/7/2017
                txtTCSTaxCode.Enabled = True : txtTCSTaxCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : cmdHelpTCSTax.Enabled = True
                'End If 'commented on 21/7/2017 for scrap invoice
            Else
                txtRefNo.Enabled = True : txtRefNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : CmdRefNoHelp.Enabled = True
                txtAmendNo.Enabled = True : txtAmendNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                ctlPerValue.Enabled = False
                ctlPerValue.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                txtTCSTaxCode.Enabled = False : txtTCSTaxCode.Text = "" : txtTCSTaxCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdHelpTCSTax.Enabled = False : lblTCSTaxPerDes.Text = "0.00"
                'txtTCSTaxCode.Enabled = True : txtTCSTaxCode.Text = "" : txtTCSTaxCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : cmdHelpTCSTax.Enabled = True : lblTCSTaxPerDes.Text = "0.00"
            End If

        End If
        'Call checktcsvalue(CmbInvType.Text, CmbInvSubType.Text)

        'added by priti
        Dim strsql As String
        Dim blntcscheck As Boolean = False

        If Len(Trim(txtCustCode.Text)) > 0 Then
            strsql = "select dbo.UDF_IRN_TCSREQUIRED( '" & gstrUNITID & "','" & txtCustCode.Text.Trim & "')"
            If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strsql)) = True Then
                blntcscheck = True
            Else
                blntcscheck = False
            End If
            If blntcscheck = True Then
                Call checktcsvalue(CmbInvType.Text, CmbInvSubType.Text)
            End If
        End If
        '' code ends


        Call CheckATNRequired(CmbInvType.Text, CmbInvSubType.Text)
        SpChEntry.MaxRows = 0
    End Sub

    Private Sub CmbInvType_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmbInvType.SelectedIndexChanged
        On Error GoTo ErrHandler
        'Procedure Call To Select InvoiceSubTypeDescription From Sale Conf Acc. To Invoice Type
        Call SelectInvoiceSubTypeFromSaleConf((CmbInvType.Text))
        If CmbInvType.Text = "REJECTION" Then
            Call CheckBatchTrackingAllowed(CmbInvType.Text, "REJ")
        End If

        '---------------Code added by Arul-------------------------------------------
        'Purpose  : To pick the flag for vendor rejection or customer rejection
        If mblnRejTracking = True And CmbInvType.Text = "REJECTION" Then
            txtRefNo.ReadOnly = True
            txtRefNo.Enabled = True
            txtRefNo.BackColor = System.Drawing.Color.White
            CmdRefNoHelp.Enabled = True
            chkRejType.Visible = True
            txtCustCode.Text = String.Empty
            With SpChEntry
                If gblnGSTUnit = True Then
                    .MaxCols = 35
                Else
                    .MaxCols = 25
                End If

                '.Row = 0 : .Col = 22 : .ColHidden = True
                .Row = 0 : .Col = 23 : .Text = "Max Qty" : .ColHidden = True
                .Row = 0 : .Col = 24 : .Text = "Doc_Detail" : .ColHidden = True
                .Row = 0 : .Col = 25 : .Text = "Rej Detail" : .ColHidden = True
            End With
            If mblnBatchTrack = True Then
                With SpChEntry
                    .Col = 25 : .ColHidden = False
                End With
            End If
        Else
            txtRefNo.ReadOnly = False
            chkRejType.Visible = False
            With SpChEntry
                '.MaxCols = 22
                If gblnGSTUnit = True Then
                    .MaxCols = 35
                    'GST DETAILS
                    .Row = 0 : .Col = 30 : .Text = "HSN/SAC CODE" : .set_ColWidth(30, 1000)
                    .Row = 0 : .Col = 31 : .Text = "CGST TAX" : .set_ColWidth(31, 1000)
                    .Row = 0 : .Col = 32 : .Text = "SGST TAX" : .set_ColWidth(32, 1000)
                    .Row = 0 : .Col = 33 : .Text = "UTGST TAX" : .set_ColWidth(33, 1000)
                    .Row = 0 : .Col = 34 : .Text = "IGST TAX" : .set_ColWidth(34, 1000)
                    .Row = 0 : .Col = 35 : .Text = "COMPENSATION CESS" : .set_ColWidth(35, 1400)
                    'GST DETAILS

                Else
                    .MaxCols = 27
                End If

            End With
        End If

        If mblnRejTracking = False And CmbInvType.Text = "REJECTION" Then
            If MsgBox("Do you want to make a Invoice for Customer", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "EMPOWER") = MsgBoxResult.Yes Then
                BLNREJECTION_FLAG = True
            Else
                BLNREJECTION_FLAG = False
            End If
        Else
            BLNREJECTION_FLAG = False
        End If

        Call CheckATNRequired(CmbInvType.Text, CmbInvSubType.Text)
        lblexportsodetails.Text = ""
        PaletteActiveInActive()
        '----------------------Ends here-----------------------------------
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Public Property SelectedItems() As String
        Get
            SelectedItems = mstrCusRefno
        End Get
        Set(ByVal Value As String)
            mstrCusRefno = Value
        End Set
    End Property
    Public Property CompileDocDetails() As String
        Get
            CompileDocDetails = mstrCompileDocDetails
        End Get
        Set(ByVal Value As String)
            mstrCompileDocDetails = Value
        End Set
    End Property

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
                        'EXPORT WILL BE CHECKED IN CASE OF EOU_FLAG IS FALSE 16/07/02
                        'If UCase(CmbInvType.Text) = "EXPORT INVOICE" Then
                        ctlPerValue.Enabled = False
                        ctlPerValue.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        txtRefNo.Enabled = True : txtRefNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : CmdRefNoHelp.Enabled = True
                        ''    txtExciseDuty.Enabled = True: txtExciseDuty.BackColor = glngCOLOR_ENABLED: txtAddExciseDuty.Enabled = True
                        ''    txtAddExciseDuty.BackColor = glngCOLOR_ENABLED: ctlSVD.Enabled = True: ctlSVD.BackColor = glngCOLOR_ENABLED
                        ctlInsurance.Enabled = True
                        ctlInsurance.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)

                        If gblnGSTUnit = True Then
                            txtSaleTaxType.Enabled = False : txtSaleTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                            txtKKC.Enabled = False : txtKKC.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                            txtSBC.Enabled = False : txtSBC.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                            lblSaltax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                            CmdSaleTaxType.Enabled = False
                            cmdkkccode.Enabled = False
                            cmdSBC.Enabled = False
                            txtSurchargeTaxType.Enabled = False : txtSurchargeTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                            lblSurcharge_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                            cmdSurchargeTaxCode.Enabled = False
                        Else
                            txtSaleTaxType.Enabled = True : txtSaleTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                            lblSaltax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                            CmdSaleTaxType.Enabled = True
                            txtSurchargeTaxType.Enabled = True : txtSurchargeTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                            lblSurcharge_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                            cmdSurchargeTaxCode.Enabled = True
                        End If

                        Cmbtrninvtype.Enabled = False : Cmbtrninvtype.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        aedamnt.Enabled = False : aedamnt.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        txtSurchargeTaxType.Text = "" : txtRefNo.Text = ""
                        ctlInsurance.Text = "" : txtSaleTaxType.Text = ""

                        lblCurrencyDes.Text = ""
                        If UCase(CmbInvType.Text) = "EXPORT INVOICE" Or UCase(CmbInvType.Text) = "SERVICE INVOICE" Then
                            lblCurrency.Visible = True : lblCurrencyDes.Visible = True
                            lblExchangeRateLable.Visible = True : lblExchangeRateValue.Visible = True
                        Else
                            lblCurrency.Visible = False : lblCurrencyDes.Visible = False
                            lblExchangeRateLable.Visible = False : lblExchangeRateValue.Visible = False
                        End If
                        'code added by nisha on 11/03/2003 for showing ToolCost in case of Sample invoice
                        With SpChEntry
                            .Col = 16 : .Col2 = 16 : .BlockMode = True : .ColHidden = True : .BlockMode = False
                        End With
                        '*******changes ends here on 11/03/2003
                    Case "JOBWORK INVOICE"
                        ctlPerValue.Enabled = False
                        ctlPerValue.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        txtRefNo.Enabled = True : txtRefNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : CmdRefNoHelp.Enabled = True
                        ''    txtExciseDuty.Enabled = False: txtExciseDuty.BackColor = glngCOLOR_DISABLED: txtAddExciseDuty.Enabled = False
                        ''    txtAddExciseDuty.BackColor = glngCOLOR_DISABLED: ctlSVD.Enabled = False: ctlSVD.BackColor = glngCOLOR_DISABLED
                        ctlInsurance.Enabled = False
                        ctlInsurance.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        If gblnGSTUnit = False Then
                            txtSaleTaxType.Enabled = True : txtSaleTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                            aedamnt.Enabled = False : aedamnt.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                            Cmbtrninvtype.Enabled = False : Cmbtrninvtype.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                            lblSaltax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                            CmdSaleTaxType.Enabled = False
                            txtSurchargeTaxType.Enabled = True : txtSurchargeTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                            lblSurcharge_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                            cmdSurchargeTaxCode.Enabled = True
                            lblCurrency.Visible = False : lblCurrencyDes.Visible = False
                            lblExchangeRateLable.Visible = False : lblExchangeRateValue.Visible = False
                            txtSurchargeTaxType.Text = "" : txtRefNo.Text = ""
                            ctlInsurance.Text = "" : txtSaleTaxType.Text = ""
                            lblCurrencyDes.Text = ""
                        Else
                            txtSaleTaxType.Enabled = False : txtSaleTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                            aedamnt.Enabled = False : aedamnt.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                            Cmbtrninvtype.Enabled = False : Cmbtrninvtype.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                            lblSaltax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                            CmdSaleTaxType.Enabled = False
                            txtSurchargeTaxType.Enabled = False : txtSurchargeTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                            lblSurcharge_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                            cmdSurchargeTaxCode.Enabled = False
                            lblCurrency.Visible = False : lblCurrencyDes.Visible = False
                            lblExchangeRateLable.Visible = False : lblExchangeRateValue.Visible = False
                            txtSurchargeTaxType.Text = "" : txtRefNo.Text = ""
                            ctlInsurance.Text = "" : txtSaleTaxType.Text = ""
                            lblCurrencyDes.Text = ""
                        End If
                        'code added by nisha on 11/03/2003 for showing ToolCost in case of Sample invoice
                        With SpChEntry
                            .Col = 16 : .Col2 = 16 : .BlockMode = True : .ColHidden = True : .BlockMode = False
                        End With
                        '*******changes ends here on 11/03/2003
                        '10869291
                    Case "SAMPLE INVOICE", "TRANSFER INVOICE", "INTER-DIVISION"
                        ctlPerValue.Enabled = False
                        ctlPerValue.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        'Chage By Rajani Kant
                        'txtRefNo.Enabled = True: txtRefNo.BackColor = glngCOLOR_ENABLED: CmdRefNoHelp.Enabled = True
                        'Code changed by Arul on 19-13-2005 to disable the controls for Sample invoice selection
                        txtRefNo.Enabled = False : txtRefNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : CmdRefNoHelp.Enabled = False
                        'Changes ends here
                        ''    txtExciseDuty.Enabled = True: txtExciseDuty.BackColor = glngCOLOR_ENABLED: txtAddExciseDuty.Enabled = True
                        ''    txtAddExciseDuty.BackColor = glngCOLOR_ENABLED: ctlSVD.Enabled = True: ctlSVD.BackColor = glngCOLOR_ENABLED
                        lblCurrency.Visible = False : lblCurrencyDes.Visible = False
                        lblExchangeRateLable.Visible = False : lblExchangeRateValue.Visible = False
                        aedamnt.Enabled = False : aedamnt.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        Cmbtrninvtype.Enabled = False : Cmbtrninvtype.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        txtSurchargeTaxType.Text = "" : txtRefNo.Text = ""
                        ctlInsurance.Text = "" : txtSaleTaxType.Text = ""
                        lblCurrencyDes.Text = ""
                        '10869291
                        If UCase(CmbInvType.Text) = "TRANSFER INVOICE" Or UCase(CmbInvType.Text) = "INTER-DIVISION" Then
                            'Code Changed by Arul on 19-13-2005 to enable the controls on the above selection only
                            txtRefNo.Enabled = True : txtRefNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : CmdRefNoHelp.Enabled = True
                            'Changes ends here
                            ctlInsurance.Enabled = True
                            ctlInsurance.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                            txtSaleTaxType.Enabled = False : txtSaleTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                            lblSaltax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                            aedamnt.Enabled = False : aedamnt.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                            Cmbtrninvtype.Enabled = True : Cmbtrninvtype.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                            CmdSaleTaxType.Enabled = False
                            txtSurchargeTaxType.Enabled = False : txtSurchargeTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                            lblSurcharge_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                            cmdSurchargeTaxCode.Enabled = False
                            With SpChEntry
                                .Col = 16 : .Col2 = 16 : .BlockMode = True : .ColHidden = True : .BlockMode = False
                            End With
                        Else
                            ctlInsurance.Enabled = False
                            ctlInsurance.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                            If gblnGSTUnit = False Then
                                txtSaleTaxType.Enabled = True : txtSaleTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                                lblSaltax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                                CmdSaleTaxType.Enabled = True
                                txtSurchargeTaxType.Enabled = True : txtSurchargeTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                                txtAddVAT.Enabled = True : txtAddVAT.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                                lblSurcharge_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                                lblAddVAT.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                                cmdSurchargeTaxCode.Enabled = True
                                CmdAddVat.Enabled = True
                            Else
                                txtSaleTaxType.Enabled = False : txtSaleTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                                lblSaltax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                                CmdSaleTaxType.Enabled = False
                                txtSurchargeTaxType.Enabled = False : txtSurchargeTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                                txtAddVAT.Enabled = False : txtAddVAT.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                                lblSurcharge_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                                lblAddVAT.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                                cmdSurchargeTaxCode.Enabled = False
                                CmdAddVat.Enabled = False
                            End If

                            With SpChEntry
                                .Col = 16 : .Col2 = 16 : .BlockMode = True : .ColHidden = False : .BlockMode = False
                            End With
                        End If
                    Case "REJECTION"
                        ctlPerValue.Enabled = False
                        ctlPerValue.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        txtSurchargeTaxType.Text = "" : txtRefNo.Text = ""
                        ctlInsurance.Text = "" : txtSaleTaxType.Text = ""
                        lblCurrencyDes.Text = ""
                        ctlInsurance.Enabled = True
                        ctlInsurance.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        If gblnGSTUnit = False Then
                            txtSaleTaxType.Enabled = True : txtSaleTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                            lblSaltax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                            CmdSaleTaxType.Enabled = True
                            txtSurchargeTaxType.Enabled = True : txtSurchargeTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                            lblSurcharge_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                            cmdSurchargeTaxCode.Enabled = True
                            txtAddVAT.Enabled = True : txtAddVAT.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                            lblAddVAT.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        Else
                            txtSaleTaxType.Enabled = False : txtSaleTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                            lblSaltax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                            CmdSaleTaxType.Enabled = False
                            txtSurchargeTaxType.Enabled = False : txtSurchargeTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                            lblSurcharge_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                            cmdSurchargeTaxCode.Enabled = False
                            txtAddVAT.Enabled = False : txtAddVAT.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                            lblAddVAT.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        End If

                        With SpChEntry
                            .Col = 16 : .Col2 = 16 : .BlockMode = True : .ColHidden = True : .BlockMode = False
                        End With
                End Select

        End Select
        SpChEntry.MaxRows = 0
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
        '******Check For Temporary Challan No.
        If txtChallanNo.Text.Trim <> "" Then
            If Val(txtChallanNo.Text.Trim.Substring(0, 2)) = 99 Then
                Cmditems.Enabled = True
            Else
                Cmditems.Enabled = False
            End If
        Else
            Cmditems.Enabled = False
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
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
        Dim blnNTRF_INV_GROUPCOMP As Boolean
        Dim blntcscheck As Boolean = False
        Dim strSql As String
        blnNTRF_INV_GROUPCOMP = Find_Value("SELECT ISNULL(TRF_INV_GROUPCOMP ,0) AS TRF_INV_GROUPCOMP  FROM SALES_PARAMETER WHERE UNIT_CODE = '" & gstrUNITID & "'")

        If Len(Trim(txtLocationCode.Text)) = 0 Then
            Call ConfirmWindow(10116, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
            If txtLocationCode.Enabled Then txtLocationCode.Focus()
            Exit Sub
        End If
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                If UCase(Trim(mstrInvoiceType)) = "INV" Or UCase(Trim(mstrInvoiceType)) = "SMP" Or UCase(Trim(mstrInvoiceType)) = "TRF" Or UCase(Trim(mstrInvoiceType)) = "JOB" Or UCase(Trim(mstrInvoiceType)) = "EXP" Or BLNREJECTION_FLAG = True Or UCase(Trim(mstrInvoiceType)) = "SRC" Or UCase(Trim(mstrInvoiceType)) = "ITD" Then
                    If Len(Trim(txtCustCode.Text)) = 0 Then
                        If gstrUNITID = "STH" Then
                            strHelpString = ShowList(1, (txtCustCode.MaxLength), "", "Customer_Code", "Cust_Name", "Customer_Mst", "and ( INVOICEAGSTSHIPPING= 0  and INVOICEAGSTACKN= 0 ) and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))")

                        ElseIf blnNTRF_INV_GROUPCOMP = True And UCase(Trim(mstrInvoiceType)) = "TRF" Then
                            strHelpString = ShowList(1, (txtCustCode.MaxLength), "", "Customer_Code", "Cust_Name", "Customer_Mst", "and Group_Customer = 1 and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))")
                        ElseIf blnNTRF_INV_GROUPCOMP = True And UCase(Trim(mstrInvoiceType)) = "ITD" Then
                            strHelpString = ShowList(1, (txtCustCode.MaxLength), "", "Customer_Code", "Cust_Name", "Customer_Mst", "and Group_Customer_InterDivision = 1 and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))")
                        ElseIf blnNTRF_INV_GROUPCOMP = True And UCase(Trim(mstrInvoiceType)) = "INV" Then
                            strHelpString = ShowList(1, (txtCustCode.MaxLength), "", "Customer_Code", "Cust_Name", "Customer_Mst", "and Group_Customer = 0 and Group_Customer_InterDivision = 0 and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))")
                        Else
                            strHelpString = ShowList(1, (txtCustCode.MaxLength), "", "Customer_Code", "Cust_Name", "Customer_Mst", "and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))")
                        End If

                        If strHelpString = "-1" Then 'If No Record Found
                            Call ConfirmWindow(10225, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Else
                            txtCustCode.Text = strHelpString
                            ' 10118036
                            mblnCSM_Knockingoff_req = DataExist("SELECT TOP 1 1 FROM CUSTOMER_MST WHERE CSM_FLAG = 1 AND CUSTOMER_CODE='" & strHelpString & "' and Unit_code='" & gstrUNITID & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))")
                            If mblnCSM_Knockingoff_req Then
                                With SpChEntry
                                    '.MaxCols = 23
                                    .Row = 0 : .Col = 23 : .Text = "CSM Knocking Details" : .set_ColWidth(23, 1700)
                                End With
                            End If
                            ' 10118036
                            Call SetControlsforASNDetails(Trim(txtCustCode.Text))
                            Call SelectDescriptionForField("Cust_Name", "Customer_Code", "Customer_Mst", lblCustCodeDes, (txtCustCode.Text))
                        End If
                    Else
                        strHelpString = ShowList(1, (txtCustCode.MaxLength), "", "Customer_Code", "Cust_Name", "Customer_Mst", "and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))")
                        If strHelpString = "-1" Then 'If No Record Found
                            Call ConfirmWindow(10225, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Else
                            txtCustCode.Text = strHelpString
                            ' 10118036
                            mblnCSM_Knockingoff_req = DataExist("SELECT TOP 1 1 FROM CUSTOMER_MST WHERE CSM_FLAG = 1 AND CUSTOMER_CODE='" & strHelpString & "' and Unit_code='" & gstrUNITID & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))")
                            If mblnCSM_Knockingoff_req Then
                                With SpChEntry
                                    .MaxCols = 23
                                    .Row = 0 : .Col = 23 : .Text = "CSM Knocking Details" : .set_ColWidth(23, 1700)
                                End With
                            End If

                            ' 10118036
                            Call SetControlsforASNDetails(Trim(txtCustCode.Text))
                            Call SelectDescriptionForField("Cust_Name", "Customer_Code", "Customer_Mst", lblCustCodeDes, (txtCustCode.Text))

                        End If
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
                    Else
                        txtTCSTaxCode.Enabled = False : txtTCSTaxCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdHelpTCSTax.Enabled = False : txtTCSTaxCode.Text = ""
                    End If
                Else
                    txtTCSTaxCode.Enabled = False : txtTCSTaxCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdHelpTCSTax.Enabled = False : txtTCSTaxCode.Text = ""
                End If
            End If
        End If

        If Len(Trim(txtCustCode.Text)) > 0 Then
            rsCustMst = New ClsResultSetDB
            strCustMst = "SELECT Bill_Address1 + ', '  +  Bill_Address2 + ', ' + Bill_City + ' - ' + Bill_Pin as  invoiceAddress From Customer_Mst  WHERE UNIT_CODE='" + gstrUNITID + "' AND Customer_code ='" & txtCustCode.Text & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
            rsCustMst.GetResult(strCustMst)
            If rsCustMst.GetNoRows > 0 Then
                lblAddressDes.Text = rsCustMst.GetValue("InvoiceAddress")
            End If
            rsCustMst.ResultSetClose()
            Me.txtRefNo.Focus()
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub Cmditems_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Cmditems.Click
        Dim frmMKTTRN0021_SOUTH As New frmMKTTRN0021_SOUTH
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
        Dim rsSaleConf As ClsResultSetDB
        Dim strStockLocation As String
        Dim rsCurrencyType As ClsResultSetDB
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        Dim strSQL As String
        With Me.SpChEntry
            If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                If CmbInvType.Text = "NORMAL INVOICE" And CmbInvSubType.Text = "FINISHED GOODS" And DataExist("SELECT PCA_CUSTOMER FROM CUSTOMER_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & Trim(txtCustCode.Text) & "' and PCA_CUSTOMER =1 ") = True Then
                    .MaxRows = 0
                    strSQL = "DELETE FROM TMP_PCA_ITEMSELECTION_PACKAGECODE where UNIT_CODE='" + gstrUNITID + "' and IP_ADDRESS='" & gstrIpaddressWinSck & "'"
                    SqlConnectionclass.ExecuteNonQuery(strSQL)
                End If
            End If
            If CmdGrpChEnt.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                .MaxRows = 1
                If mblnCSM_Knockingoff_req Then
                    .Row = 1 : .Row2 = .MaxRows : .Col = 23 : .ColHidden = False
                End If
                .Row = 1 : .Row2 = .MaxRows : .Col = 1 : .Col2 = .MaxCols : .BlockMode = True : .Text = "" : .BlockMode = False
            End If
        End With
        frmMKTTRN0021_SOUTH.Cust_Code = Trim(txtCustCode.Text)
        frmMKTTRN0021_SOUTH.Invoice_type = Trim(CmbInvType.Text)
        frmMKTTRN0021_SOUTH.Invoice_Subtype = Trim(CmbInvSubType.Text)


        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                rssalechallan = New ClsResultSetDB
                salechallan = ""
                salechallan = "SELECT Invoice_type,SUB_CATEGORY FROM saleschallan_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  doc_No = "
                salechallan = salechallan & Val(txtChallanNo.Text)
                rssalechallan.GetResult(salechallan, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                If rssalechallan.GetNoRows > 0 Then
                    rssalechallan.MoveFirst()
                    strInvType = rssalechallan.GetValue("Invoice_type")
                    strInvSubType = rssalechallan.GetValue("sub_category")
                End If
                If strInvType = "REJ" Then
                    SpChEntry.MaxCols = 22
                End If
                rssalechallan.ResultSetClose()
                strStockLocation = StockLocationSalesConf(strInvType, strInvSubType, "TYPE")
                mstrLocationCode = Trim(strStockLocation)
                If Len(Trim(strStockLocation)) > 0 Then
                    If (UCase(strInvType) = "INV") Or (UCase(strInvType) = "EXP") Then
                        mstrItemCode = frmMKTTRN0021_SOUTH.SelectDatafromsaleDtl(Trim(txtChallanNo.Text))
                        If Len(Trim(mstrItemCode)) = 0 Then
                            SpChEntry.MaxRows = 0
                            frmMKTTRN0021_SOUTH = Nothing
                        End If
                    Else
                        mstrItemCode = frmMKTTRN0021_SOUTH.SelectDatafromsaleDtl(Trim(txtChallanNo.Text))
                        If Len(Trim(mstrItemCode)) = 0 Then
                            SpChEntry.MaxRows = 0
                            frmMKTTRN0021_SOUTH = Nothing
                        End If
                    End If
                Else
                    MsgBox("Please Define Stock Location in Sales Conf")
                    Exit Sub
                End If
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD

                If mblnRejTracking = True And UCase(CStr((Trim(CmbInvType.Text)) = "REJECTION")) Then
                    If Len(Trim(txtRefNo.Text)) = 0 Then
                        MsgBox("Enter Reference No", MsgBoxStyle.Information, ResolveResString(100))
                        Exit Sub
                    End If
                End If

                If mblnBatchTrack = True And UCase(CmbInvType.Text) <> "JOBWORK INVOICE" And UCase(CmbInvType.Text) <> "REJECTION INVOICE" Then
                    If MsgBox(" Do You Want To Follow FIFO Wise Batch Tracking ", vbYesNo, ResolveResString(100)) = vbYes Then
                        mblnbatchfifomode = True
                    Else
                        mblnbatchfifomode = False
                    End If
                End If
                '10869291
                If UCase(Trim(CmbInvType.Text)) = "NORMAL INVOICE" Or UCase(Trim(CmbInvType.Text)) = "JOBWORK INVOICE" Or UCase(Trim(CmbInvType.Text)) = "EXPORT INVOICE" Or UCase(Trim(CmbInvType.Text)) = "TRANSFER INVOICE" Or UCase(CmbInvType.Text) = "INTER-DIVISION" Then
                    If UCase(Trim(CmbInvSubType.Text)) <> "SCRAP" Then
                        If Len(Trim(txtRefNo.Text)) = 0 Then
                            Call ConfirmWindow(10240, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                            If txtRefNo.Enabled Then txtRefNo.Focus()
                            Exit Sub
                        ElseIf Len(Trim(txtAmendNo.Text)) = 0 Then
                            'Condition Added 
                            'User Can Enter Ref Code And Amendment From KeyBoard 1.Check If Ref No with Blank Amend is Over Or NOT
                            '   2.If Over Then see y No Amendments are added
                            If OriginalRefNoOVER(Trim(txtRefNo.Text)) Then
                                'Orig Ref No is OVER , So Amendment Number should be added
                                MsgBox("Enter Amendment No.", MsgBoxStyle.Information, "empower")
                                If txtAmendNo.Enabled Then txtAmendNo.Focus() : Exit Sub
                            Else
                                'Original Ref No is Still Active
                                mstrRefNo = Trim(txtRefNo.Text)
                                mstrAmmNo = "" 'As Amend No is Blank
                            End If
                        Else
                            'Reference Number is Already Verfied in Validate Event of txtRefNo Amend No. is Also Verified in Its Validate Event
                            'Then Pass It to Form variables
                            mstrRefNo = Trim(txtRefNo.Text)
                            mstrAmmNo = Trim(txtAmendNo.Text)
                        End If
                    End If
                End If

                If (Trim(CmbInvType.Text) = "JOBWORK INVOICE") Then
                    'jul
                    If blnFIFO = False Then
                        If Len(Trim(mstrRGP)) = 0 Then
                            MsgBox("First Select RGP No. ", MsgBoxStyle.OkOnly, "empower")
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

                frmMKTTRN0021_SOUTH.SHOP_CODE = Find_Value("SELECT TOP 1 SHOP_CODE FROM CUSTITEM_MST WHERE UNIT_CODE='" + gstrUNITID + "' AND ACCOUNT_CODE='" & Me.txtCustCode.Text.Trim & "' AND ITEM_CODE='" & varItemCode & "' AND Active=1 ")
                frmMKTTRN0021_SOUTH.TOTALALREADYITEMINGRID = Me.SpChEntry.MaxRows
                '10869291
                If UCase(Trim(CmbInvType.Text)) = "NORMAL INVOICE" Or UCase(Trim(CmbInvType.Text)) = "EXPORT INVOICE" Or UCase(Trim(CmbInvType.Text)) = "TRANSFER INVOICE" Or UCase(Trim(CmbInvType.Text)) = "INTER-DIVISION" Then
                    strStockLocation = StockLocationSalesConf((CmbInvType.Text), (CmbInvSubType.Text), "DESCRIPTION")
                    mstrLocationCode = Trim(strStockLocation)
                    If Len(Trim(strStockLocation)) > 0 Then
                        If CBool(UCase(CStr(Trim(CmbInvSubType.Text) = "SCRAP"))) Then
                            If Len(Trim(strItemNotIn)) > 0 Then
                                mstrItemCode = frmMKTTRN0021_SOUTH.SelectDatafromItem_Mst(Trim(CmbInvType.Text), Trim(CmbInvSubType.Text), strStockLocation, , strItemNotIn, SpChEntry.MaxRows)
                            Else
                                mstrItemCode = frmMKTTRN0021_SOUTH.SelectDatafromItem_Mst(Trim(CmbInvType.Text), Trim(CmbInvSubType.Text), strStockLocation)
                            End If
                        Else
                            If MULTIPLESO > 1 Then
                                mstrAmmNo = ""
                            End If
                            If Len(Trim(strItemNotIn)) > 0 Then
                                mstrItemCode = frmMKTTRN0021_SOUTH.SelectDataFromCustOrd_Dtl(Trim(txtCustCode.Text), Trim(CUSTREFLIST), mstrAmmNo, Trim(CmbInvSubType.Text), Trim(CmbInvType.Text), strStockLocation, strItemNotIn, SpChEntry.MaxRows)
                            Else
                                mstrItemCode = frmMKTTRN0021_SOUTH.SelectDataFromCustOrd_Dtl(Trim(txtCustCode.Text), Trim(CUSTREFLIST), mstrAmmNo, Trim(CmbInvSubType.Text), Trim(CmbInvType.Text), strStockLocation)
                            End If
                            'Changes ends here
                        End If
                    Else
                        MsgBox("Please Define Stock Location in Sales Conf", MsgBoxStyle.Information, "empower")
                        Exit Sub
                    End If


                    If Len(Trim(mstrItemCode)) = 0 Then SpChEntry.MaxRows = 0
                ElseIf (Trim(CmbInvType.Text) = "JOBWORK INVOICE" Or Trim(CmbInvType.Text) = "SERVICE INVOICE") Then
                    strStockLocation = StockLocationSalesConf((CmbInvType.Text), (CmbInvSubType.Text), "DESCRIPTION")
                    If Len(Trim(strStockLocation)) > 0 Then
                        If Len(Trim(strItemNotIn)) > 0 Then
                            mstrItemCode = frmMKTTRN0021_SOUTH.SelectDataFromCustOrd_Dtl(Trim(txtCustCode.Text), Trim(txtRefNo.Text), mstrAmmNo, Trim(CmbInvSubType.Text), Trim(CmbInvType.Text), strStockLocation, strItemNotIn, SpChEntry.MaxRows)
                        Else
                            mstrItemCode = frmMKTTRN0021_SOUTH.SelectDataFromCustOrd_Dtl(Trim(txtCustCode.Text), Trim(txtRefNo.Text), mstrAmmNo, Trim(CmbInvSubType.Text), Trim(CmbInvType.Text), strStockLocation)
                        End If
                    Else
                        MsgBox("Please Define Stock Location in Sales Conf", MsgBoxStyle.Information, "empower")
                        Exit Sub
                    End If
                    If Len(Trim(mstrItemCode)) = 0 Then SpChEntry.MaxRows = 0
                Else
                    rsSaleConf = New ClsResultSetDB
                    rsSaleConf.GetResult("select Stock_Location From saleconf (nolock) WHERE UNIT_CODE='" + gstrUNITID + "' AND  Description ='" & Trim(CmbInvType.Text) & "' and sub_type_description ='" & Trim(CmbInvSubType.Text) & "' AND Location_Code ='" & Trim(txtLocationCode.Text) & "' and (fin_start_date <= getdate() and fin_end_date >= getdate())")
                    If ((Len(Trim(rsSaleConf.GetValue("Stock_Location"))) = 0) Or (Trim(CStr(rsSaleConf.GetValue("Stock_Location") = "Unknown")))) Then
                        MsgBox("Plese Select Stock Location in SalesConf first", MsgBoxStyle.Information, "empower")
                        If Cmditems.Enabled Then Cmditems.Focus()
                        Exit Sub
                    Else
                        mstrLocationCode = rsSaleConf.GetValue("Stock_Location")
                    End If
                    If CBool(UCase(CStr(Trim(CmbInvType.Text) = "REJECTION"))) Then
                        If mblnRejTracking = False Then
                            If Len(Trim(txtRefNo.Text)) > 0 Then
                                If Len(Trim(strItemNotIn)) > 0 Then
                                    mstrItemCode = frmMKTTRN0021_SOUTH.AddDataFromGrinDtl(Trim(txtCustCode.Text), Trim(txtRefNo.Text), rsSaleConf.GetValue("Stock_Location"), SpChEntry.MaxRows, strItemNotIn)
                                Else
                                    mstrItemCode = frmMKTTRN0021_SOUTH.AddDataFromGrinDtl(Trim(txtCustCode.Text), Trim(txtRefNo.Text), rsSaleConf.GetValue("Stock_Location"))
                                End If
                            Else
                                If Len(Trim(strItemNotIn)) > 0 Then
                                    mstrItemCode = frmMKTTRN0021_SOUTH.SelectDatafromItem_Mst(Trim(CmbInvType.Text), Trim(CmbInvSubType.Text), rsSaleConf.GetValue("Stock_Location"), Trim(txtCustCode.Text), strItemNotIn, SpChEntry.MaxRows)
                                Else
                                    mstrItemCode = frmMKTTRN0021_SOUTH.SelectDatafromItem_Mst(Trim(CmbInvType.Text), Trim(CmbInvSubType.Text), rsSaleConf.GetValue("Stock_Location"), Trim(txtCustCode.Text))
                                End If
                            End If
                        Else
                            If Len(Trim(txtRefNo.Text)) > 0 Then
                                If Len(Trim(strItemNotIn)) > 0 Then
                                    mstrItemCode = frmMKTTRN0021_SOUTH.AddDataFromGRNORLRN(Trim(txtCustCode.Text), Trim(txtRefNo.Text), rsSaleConf.GetValue("Stock_Location"), chkRejType.Text, SpChEntry.MaxRows, strItemNotIn)
                                Else
                                    mstrItemCode = frmMKTTRN0021_SOUTH.AddDataFromGRNORLRN(Trim(txtCustCode.Text), Trim(txtRefNo.Text), rsSaleConf.GetValue("Stock_Location"), chkRejType.Text)
                                End If
                            Else
                                If Len(Trim(strItemNotIn)) > 0 Then
                                    mstrItemCode = frmMKTTRN0021_SOUTH.SelectDatafromItem_Mst(Trim(CmbInvType.Text), Trim(CmbInvSubType.Text), rsSaleConf.GetValue("Stock_Location"), Trim(txtCustCode.Text), strItemNotIn, SpChEntry.MaxRows)
                                Else
                                    mstrItemCode = frmMKTTRN0021_SOUTH.SelectDatafromItem_Mst(Trim(CmbInvType.Text), Trim(CmbInvSubType.Text), rsSaleConf.GetValue("Stock_Location"), Trim(txtCustCode.Text))
                                End If
                            End If
                        End If
                        rsSaleConf.ResultSetClose()
                        If Len(Trim(mstrItemCode)) = 0 And Len(Trim(strItemNotIn)) = 0 Then
                            SpChEntry.MaxRows = 0
                        Else
                            If Len(Trim(mstrItemCode)) = 0 Then
                                '  mstrItemCode = strItemNotIn
                            End If
                        End If
                    Else
                        If Len(Trim(strItemNotIn)) > 0 Then
                            mstrItemCode = frmMKTTRN0021_SOUTH.SelectDatafromItem_Mst(Trim(CmbInvType.Text), Trim(CmbInvSubType.Text), rsSaleConf.GetValue("Stock_Location"), Trim(txtCustCode.Text), strItemNotIn, SpChEntry.MaxRows)
                        Else
                            mstrItemCode = frmMKTTRN0021_SOUTH.SelectDatafromItem_Mst(Trim(CmbInvType.Text), Trim(CmbInvSubType.Text), rsSaleConf.GetValue("Stock_Location"), Trim(txtCustCode.Text))
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
                    '*************** to get refrence detail for curenct
                    rsCurrencyType = New ClsResultSetDB
                    rsCurrencyType.GetResult("Select Currency_code from saleschallan_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  doc_No = " & Val(txtChallanNo.Text))
                    If rsCurrencyType.GetNoRows > 0 Then
                        rsCurrencyType.MoveFirst()
                        strCurrency = rsCurrencyType.GetValue("Currency_code")
                    End If
                    rsCurrencyType.ResultSetClose()
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
                If CDbl(txtChallanNo.Text.Trim.Substring(0, 2)) = 99 Then
                    Me.CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = True
                    Me.CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = True
                End If
            End If
            If Me.CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                If ctlInsurance.Enabled Then
                    If ctlInsurance.Enabled Then ctlInsurance.Focus()
                Else
                    System.Windows.Forms.SendKeys.Send("{tab}")
                End If
            Else
                Me.CmdGrpChEnt.Focus()
            End If
        Else
            frmMKTTRN0021_SOUTH = Nothing
        End If
        'Set Cell Type In Spread
        Call ChangeCellTypeStaticText()
        If Me.CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            If ctlInsurance.Enabled Then ctlInsurance.Focus()
        Else
            Me.CmdGrpChEnt.Focus()
        End If
        PaletteActiveInActive()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub CmdLocCodeHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdLocCodeHelp.Click
        On Error GoTo ErrHandler
        Dim strHelp As String
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                If Len(Me.txtLocationCode.Text) = 0 Then 'To check if There is No Text Then Show All Help
                    strHelp = ShowList(1, (txtLocationCode.MaxLength), "", "s.Location_Code", "l.Description", "Location_mst l,SaleConf s (nolock)", "and s.Location_Code=l.Location_Code AND s.UNIT_CODE=l.UNIT_CODE  ", , , , , , "s.UNIT_CODE")
                    If strHelp = "-1" Then 'If No Record Exists In The Table
                        Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Exit Sub
                    Else
                        txtLocationCode.Text = strHelp
                    End If
                Else
                    'To Display All Possible Help Starting With Text in TextField
                    strHelp = ShowList(1, (txtLocationCode.MaxLength), txtLocationCode.Text, "s.Location_Code", "l.Description", "Location_mst l,SaleConf s (nolock) ", "and s.Location_Code=l.Location_Code and s.UNIT_CODE=l.UNIT_CODE ", , , , , , "s.UNIT_CODE")
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
        Dim frmMKTTRN0020NEW As New frmMKTTRN0020NEW
        Dim frmMKTTRN0009a_SOUTH As New frmMKTTRN0009a_SOUTH
        '****************************************************
        'Created By     -  Nisha
        'Description    -  To Display Details Of Customer Order
        '****************************************************
        On Error GoTo ErrHandler
        If Len(txtCustCode.Text) = 0 Then
            Call ConfirmWindow(10416, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
            txtCustCode.Focus()
            Exit Sub
        End If
        Dim strRefAmm As String
        Dim intPos As Short
        If UCase(CmbInvType.Text) <> "REJECTION" Then
            strRefAmm = frmMKTTRN0020NEW.SelectDataFromCustOrd_Dtl(txtCustCode.Text, CmbInvType.Text)
        Else
            If mblnRejTracking = True Then ' Tracking REQ
                frmMKTTRN0009a_SOUTH.RejectionType = chkRejType.Text
                frmMKTTRN0009a_SOUTH.Vendor_code = txtCustCode.Text
                frmMKTTRN0009a_SOUTH.Batch_Tracking = mblnBatchTracking
                SelectedItems = ""
                Call frmMKTTRN0009a_SOUTH.ShowDialog()
                strRefAmm = SelectedItems
            Else
                strRefAmm = frmMKTTRN0020NEW.SelectDataFromGrinDtl(txtCustCode.Text)
            End If
        End If
        If Len(strRefAmm) > 0 Then
            If mblnRejTracking = True And CmbInvType.Text = "REJECTION" Then ' Tracking Req
                txtRefNo.Text = strRefAmm
            Else
                intPos = InStr(1, Trim(strRefAmm), ",", CompareMethod.Text)
                mstrRefNo = Mid(Trim(strRefAmm), 2, intPos - 3)
                mstrAmmNo = Mid(strRefAmm, intPos + 2, ((Len(Trim(strRefAmm))) - intPos) - 2)
                txtRefNo.Text = Trim(mstrRefNo)
                txtAmendNo.Text = mstrAmmNo
                If CmbInvType.Text.ToUpper = "EXPORT INVOICE" Then
                    mstrexportsotype = Find_Value("SELECT EXPORTSOTYPE FROM CUST_ORD_HDR WHERE UNIT_CODE='" + gstrUNITID + "' AND ACCOUNT_CODE='" & txtCustCode.Text & "' AND cust_ref='" & mstrRefNo & "' and amendment_no='" & mstrAmmNo & "'")
                    lblexportsodetails.Text = mstrexportsotype
                Else
                    lblexportsodetails.Text = ""
                End If
            End If
        Else
            txtAmendNo.Focus()
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
            MsgBox("Select atleast one RGP from List", MsgBoxStyle.Information, "empower")
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
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                If Len(Me.txtSaleTaxType.Text) = 0 Then 'To check if There is No Text Then Show All Help
                    ''Query Changed By Tapan
                    strHelp = ShowList(1, (txtSaleTaxType.MaxLength), "", "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND (Tx_TaxeID='CST' OR Tx_TaxeID='LST' or Tx_TaxeID='VAT')")
                    If strHelp = "-1" Then 'If No Record Exists In The Table
                        Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Exit Sub
                    Else
                        txtSaleTaxType.Text = strHelp
                    End If
                Else
                    'To Display All Possible Help Starting With Text in TextField
                    ''Query Changed By Tapan
                    strHelp = ShowList(1, (txtSaleTaxType.MaxLength), txtSaleTaxType.Text, "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND (Tx_TaxeID='CST' OR Tx_TaxeID='LST' or Tx_TaxeID='VAT')")
                    If strHelp = "-1" Then 'If No Record Exists In The Table
                        Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Exit Sub
                    Else
                        txtSaleTaxType.Text = strHelp
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
                    strHelp = ShowList(1, (txtSurchargeTaxType.MaxLength), "", "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND Tx_TaxeID='SST'")
                    If strHelp = "-1" Then 'If No Record Exists In The Table
                        Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Exit Sub
                    Else
                        txtSurchargeTaxType.Text = strHelp
                    End If
                Else
                    strHelp = ShowList(1, (txtSurchargeTaxType.MaxLength), txtSurchargeTaxType.Text, "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND Tx_TaxeID='SST'")
                    If strHelp = "-1" Then 'If No Record Exists In The Table
                        Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Exit Sub
                    Else
                        txtSurchargeTaxType.Text = strHelp
                    End If
                End If
                Call txtSurchargeTaxType_Validating(txtSurchargeTaxType, New System.ComponentModel.CancelEventArgs(False))
        End Select
        'Procedure Call To Select The Location Code Description
        ''    txtSaleTaxType.SetFocus
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
        '-----------------------------------------------------------------------------------
        'Created By      : Ashutosh Verma
        'Issue ID        :
        'Creation Date   : 06 Mar 2007
        'Function        : To Show help for New Tax SEcess
        '-----------------------------------------------------------------------------------
        Dim strHelp As String
        On Error GoTo ErrHandler
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                If Len(txtSECSSTaxType.Text) = 0 Then 'To check if There is No Text Then Show All Help
                    '------------------Satvir Handa------------------------
                    strHelp = ShowList(1, (txtSECSSTaxType.MaxLength), "", "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND (Tx_TaxeID='ECSSH') and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))")
                    '------------------Satvir Handa------------------------
                    If strHelp = "-1" Then 'If No Record Exists In The Table
                        Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Exit Sub
                    Else
                        txtSECSSTaxType.Text = strHelp
                    End If
                Else
                    '------------------Satvir Handa------------------------
                    strHelp = ShowList(1, (txtSECSSTaxType.MaxLength), txtSECSSTaxType.Text, "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND (Tx_TaxeID='ECSSH')  and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))")
                    '------------------Satvir Handa------------------------
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
    Private Sub frmMKTTRN0009_SOUTH_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
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
    Private Sub frmMKTTRN0009_SOUTH_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Escape
                'If user press the ESC Key ,the Form will be in View Mode
                If Me.CmdGrpChEnt.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                    If ConfirmWindow(10053, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                        CreatePaletteDataTable()
                        Call Me.CmdGrpChEnt.Revert()
                        If mblnCSM_Knockingoff_req Then
                            mP_Connection.Execute("IF NOT EXISTS(SELECT TOP 1 1 FROM CSM_KNOCKOFF_DTL CS INNER JOIN SALES_DTL SC on (SC.UNIT_CODE=CS.UNIT_CODE AND SC.DOC_NO=CS.INV_NO ) AND SC.UNIT_CODE='" + gstrUnitId + "' AND  CS.INV_NO = " & txtChallanNo.Text & ") DELETE FROM CSM_KNOCKOFF_DTL WHERE UNIT_CODE='" + gstrUnitId + "' AND  INV_NO = " & txtChallanNo.Text, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        End If
                        Call EnableControls(False, Me, True)
                        'In View Mode Disable Combo Of Invoice Type and Inv. Sub type
                        With SpChEntry
                            .Col = 16 : .Col2 = 16 : .BlockMode = True : .ColHidden = True : .BlockMode = False
                        End With

                        CmbInvType.Visible = False : CmbInvSubType.Visible = False
                        lblInvSubType.Visible = False : lblInvType.Visible = False

                        txtLocationCode.Enabled = True : txtLocationCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : lblLocCodeDes.Text = ""
                        txtChallanNo.Enabled = True : txtChallanNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        CmdLocCodeHelp.Enabled = True : CmdChallanNo.Enabled = True : Me.SpChEntry.Enabled = False
                        Me.CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                        Me.CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                        Me.CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False

                        lblSaltax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        lblSurcharge_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        lblECSStax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)

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
                        lblDateDes.Text = VB6.Format(GetServerDate(), gstrDateFormat)
                        dtpDateDesc.Visible = False
                        txtLocationCode.Focus()
                        mP_Connection.Execute("if exists(select name from sysobjects where name = '" + frmMKTTRN0020NEW.strTmpTable + "') drop table " + frmMKTTRN0020NEW.strTmpTable, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
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
    Private Sub frmMKTTRN0009_SOUTH_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo ErrHandler
        'Add Form Name To Window List
        mintIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader1.Tag)
        'Fill Lebels From Resource File
        Call FillLabelFromResFile(Me) 'Fill Labels >From Resource File
        Call FitToClient(Me, FraChEnt, ctlFormHeader1, CmdGrpChEnt, 625)
        CmdLocCodeHelp.Image = My.Resources.ico111.ToBitmap
        CmdChallanNo.Image = My.Resources.ico111.ToBitmap
        CmdCustCodeHelp.Image = My.Resources.ico111.ToBitmap
        CmdSaleTaxType.Image = My.Resources.ico111.ToBitmap
        CmdRefNoHelp.Image = My.Resources.ico111.ToBitmap
        'Check If Company is 100% EOU then CVD SVD fields are SHOWN 
        gobjDB = New ClsResultSetDB
        If gobjDB.GetResult("Select EOU_FLAG From Company_Mst WHERE UNIT_CODE='" + gstrUNITID + "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic) Then
            If gobjDB.GetNoRows > 0 Then
                blnEOU_FLAG = gobjDB.GetValue("EOU_FLAG")
            End If
        End If
        gobjDB.ResultSetClose()
        gobjDB = Nothing
        'Initially Disable All Controls
        Call EnableControls(False, Me, True)
        lblSaltax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        lblSurcharge_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        lblTCSTaxPerDes.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        'Date is Also Added in DatePicker,and Its Visible Property is set to False - Nitin Sood
        With dtpDateDesc
            .Format = DateTimePickerFormat.Custom
            .CustomFormat = gstrDateFormat
            .Value = GetServerDate() 'Get Server Date
            .Visible = False
        End With
        lblDateDes.Text = dtpDateDesc.Text
        'Add Transport Type To Combo
        Call AddTransPortTypeToCombo()

        txtLocationCode.Enabled = True : txtLocationCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        txtChallanNo.Enabled = True : txtChallanNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        CmdLocCodeHelp.Enabled = True : CmdChallanNo.Enabled = True
        Me.SpChEntry.Enabled = False
        Me.CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
        Me.CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
        Me.CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
        'txtTCSTaxCode.Enabled = False : txtTCSTaxCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdHelpTCSTax.Enabled = False
        txtTCSTaxCode.Enabled = True : txtTCSTaxCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : cmdHelpTCSTax.Enabled = True
        mblnBatchTracking = False
        mblnbatchfifomode = True
        'Set Column Headers
        With Me.SpChEntry
            .DisplayRowHeaders = True
            .set_ColWidth(0, 300)
            If mblnBatchTracking = True Then
                .MaxCols = 22
            Else
                .MaxCols = 21
            End If
            .MaxCols = 27
            'mblnCSM_Knockingoff_req = DataExist("SELECT TOP 1 1 FROM SALES_PARAMETER WHERE UNIT_CODE='" + gstrUNITID + "' AND  CSM_KNOCKINGOFF_REQ = 1 ")
            'If mblnCSM_Knockingoff_req Then
            '    .MaxCols = 23
            'End If
            .Row = 0 : .Col = 1 : .Text = "Internal Part No." : .set_ColWidth(1, 1650)
            .Row = 0 : .Col = 2 : .Text = "Cust.Part No." : .set_ColWidth(2, 1700)
            '***
            .Row = 0 : .Col = 3 : .Text = "Rate (Per Unit)" : .set_ColWidth(3, 1500)
            .Row = 0 : .Col = 4 : .Text = "Cust supp. Mat (Per Unit)" : .set_ColWidth(4, 1800)
            .Row = 0 : .Col = 5 : .Text = "Quantity"
            .Row = 0 : .Col = 6 : .Text = "Packing(%)" : .set_ColWidth(6, 1500)
            .Row = 0 : .Col = 7 : .Text = "EXC(%)"
            .Row = 0 : .Col = 8 : .Text = "ADD EXC(%)" : .set_ColWidth(8, 1500)
            .Row = 0 : .Col = 9 : .Text = "CVD(%)"
            .Row = 0 : .Col = 10 : .Text = "SAD(%)"
            If Not blnEOU_FLAG Then
                .Col = 9 : .Col2 = 9
                .BlockMode = True
                .ColHidden = True
                .BlockMode = False
                .Col = 10 : .Col2 = 10
                .BlockMode = True
                .ColHidden = True
                .BlockMode = False
            End If
            .Row = 0 : .Col = 11 : .Text = "Others (Per Unit)" : .set_ColWidth(11, 1500)
            .Row = 0 : .Col = 12 : .Text = "From Box"
            .Row = 0 : .Col = 13 : .Text = "To Box"
            .Row = 0 : .Col = 14 : .Text = "Cumulative Boxes" : .set_ColWidth(14, 1500)
            .Row = 0 : .Col = 15 : .Text = "Delete"
            .Col = 15 : .Col2 = 15 : .BlockMode = True : .ColHidden = True : .BlockMode = False
            .Row = 0 : .Col = 16 : .Text = "Tool Cost (Per Unit)"
            .Col = 16 : .Col2 = 16 : .BlockMode = True : .ColHidden = True : .BlockMode = False
            .Row = 0 : .Col = 17 : .Text = "Rate"
            .Col = 17 : .Col2 = 17 : .BlockMode = True : .ColHidden = True : .BlockMode = False
            .Row = 0 : .Col = 18 : .Text = "Cust Mtrl"
            .Col = 18 : .Col2 = 18 : .BlockMode = True : .ColHidden = True : .BlockMode = False
            .Row = 0 : .Col = 19 : .Text = "Others"
            .Col = 19 : .Col2 = 19 : .BlockMode = True : .ColHidden = True : .BlockMode = False
            .Row = 0 : .Col = 20 : .Text = "Tool Cost"
            .Col = 20 : .Col2 = 20 : .BlockMode = True : .ColHidden = True : .BlockMode = False
            If mblnBatchTracking = True And UCase(CmbInvType.Text) <> "JOBWORK INVOICE" And UCase(CmbInvType.Text) <> "REJECTION" Then
                .Row = 0 : .Col = 21 : .Text = "Batch Details" : .set_ColWidth(21, 1300)
                .Row = 0 : .Col = 22 : .Text = "Max Quantity" : .ColHidden = True : .BlockMode = False : .set_ColWidth(22, 1200)
            Else
                .Row = 0 : .Col = 22 : .Text = "Max Quantity" : .ColHidden = True : .BlockMode = False : .set_ColWidth(22, 1200)
            End If
            'If mblnCSM_Knockingoff_req Then
            '    .Row = 0 : .Col = 23 : .Text = "CSM Knocking Details" : .set_ColWidth(23, 1700)
            'End If
            '10808160
            .Row = 0 : .Col = 24 : .Text = "Doc_Detail" : .ColHidden = True
            .Row = 0 : .Col = 25 : .Text = "Rej Detail" : .ColHidden = True
            .Row = 0 : .Col = 26 : .Text = "Detail" : .ColHidden = True
            .Row = 0 : .Col = 27 : .Text = "Model" : .set_ColWidth(27, 1200)
            '10808160
            'ATN CHANGES
            .Row = 0 : .Col = 28 : .Text = "ATN No" : .set_ColWidth(28, 1200)
            'ATN CHANGES
        End With


        'Function Call To Add Invoice Types In The Inv. Type Combo Box
        Call SelectInvoiceTypeFromSaleConf()
        'In View Mode Disable Combo Of Invoice Type and Inv. Sub type
        CmbInvType.Visible = False : CmbInvSubType.Visible = False
        lblInvSubType.Visible = False : lblInvType.Visible = False
        'GST DETAILS
        If gblnGSTUnit = True Then
            With SpChEntry
                .MaxCols = 35
                .Row = 0 : .Col = 30 : .Text = "HSN/SAC CODE" : .set_ColWidth(30, 1000)
                .Row = 0 : .Col = 31 : .Text = "CGST TAX" : .set_ColWidth(31, 1000)
                .Row = 0 : .Col = 32 : .Text = "SGST TAX" : .set_ColWidth(32, 1000)
                .Row = 0 : .Col = 33 : .Text = "UTGST TAX" : .set_ColWidth(33, 1000)
                .Row = 0 : .Col = 34 : .Text = "IGST TAX" : .set_ColWidth(34, 1000)
                .Row = 0 : .Col = 35 : .Text = "COMPENSATION CESS" : .set_ColWidth(35, 1400)

            End With
        End If
        'GST DETAILS

        'Add Row
        Call addRowAtEnterKeyPress(1)
        fraRGPs.Visible = False
        lblRGPDes.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        lblCurrency.Visible = False
        lblCurrencyDes.Visible = False
        lblExchangeRateLable.Visible = False
        lblExchangeRateValue.Visible = False
        mstrFGDomestic = Find_Value("Select FG_DOMESTIC from BarCode_config_mst WHERE UNIT_CODE='" + gstrUNITID + "'")
        '10178530 
        mblnRejTracking = CBool(Find_Value("Select REJINV_Tracking from sales_parameter WHERE UNIT_CODE='" + gstrUNITID + "'"))
        mstrcurrentATNCode = Find_Value("SELECT COMPANYCODE  FROM FA_MCOMPANY WHERE EMPROUNITCODE='" & gstrUNITID & "'")
        Call CheckATNRequired(CmbInvType.Text, CmbInvSubType.Text)
        CreatePaletteDataTable()
        'added by priti on 16 march 2020 to add vehicle help box
        mblnAllowTransporterfromMaster = CBool(Find_Value("SELECT isnull(AllowTransporterfromMaster,0) as AllowTransporterfromMaster  FROM SALES_PARAMETER WHERE UNIT_CODE='" + gstrUNITID + "'"))
        If mblnAllowTransporterfromMaster Then
            txtVehNo.Enabled = True
            cmdVehicleCodeHelp.Visible = True
        Else
            txtVehNo.Enabled = True
            cmdVehicleCodeHelp.Visible = False
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0009_SOUTH_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo ErrHandler
        mdifrmMain.CheckFormName = mintIndex
        txtLocationCode.Focus()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0009_SOUTH_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Me.Tag) = False
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0009_SOUTH_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
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
                    'Set Global VAriable
                    gblnCancelUnload = True
                    gblnFormAddEdit = True
                    Me.CmdGrpChEnt.Focus()
                End If
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
    Private Sub frmMKTTRN0009_SOUTH_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error GoTo ErrHandler
        Me.Dispose()
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
        With Me.SpChEntry
            .CursorStyle = FPSpreadADO.CursorStyleConstants.CursorStyleArrow
            For intRowHeight = 1 To pintRows
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                .set_RowHeight(.Row, 300)
                If CmbInvType.Text = "NORMAL INVOICE" And CmbInvSubType.Text = "FINISHED GOODS" And mblnCSM_Knockingoff_req Then
                    .Col = 23 : .Row = .MaxRows : .ColHidden = False : .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton : .TypeButtonText = "Knocking Details"
                Else
                    .Col = 23 : .ColHidden = True
                End If
            Next intRowHeight
            If .MaxRows > 4 Then .ScrollBars = FPSpreadADO.ScrollBarsConstants.ScrollBarsBoth
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
            strSaleConfSql = "Select Distinct(Description) from SaleConf (nolock) WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_Type Not in('STX') and (fin_start_date <= getdate() and fin_end_date >= getdate()) "
        Else
            strSaleConfSql = "Select Distinct(Description) from SaleConf (nolock) WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_Type Not in('EXP','STX') and (fin_start_date <= getdate() and fin_end_date >= getdate()) "
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
        'Description    -  Select Invoice SubTypeDescription From SaleConf Acc. to Inv. Type
        '****************************************************
        On Error GoTo ErrHandler
        Dim strSaleConfSql As String
        Dim rsSaleConf As ClsResultSetDB
        Dim intRecCount As Short
        Dim intLoopCounter As Short
        strSaleConfSql = "Select Distinct(Sub_Type_Description) from SaleConf (nolock)  WHERE UNIT_CODE='" + gstrUNITID + "' AND  Description='" & Trim(pstrInvType) & "' and  (fin_start_date <= getdate() and fin_end_date >= getdate()) "
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
        'Description    -  To Select The Field Description In The Description Labels
        'Arguments      -  pstrFieldName1 - Field Name1,pstrFieldName2 - Field Name2,pstrTableName - Table Name
        '               -  pContName - Name Of The Control where Caption Is To Be Set
        '               -  pstrControlText - Field Text
        '****************************************************

        On Error GoTo ErrHandler
        Dim strDesSql As String 'Declared to make Select Query
        Dim rsDescription As ClsResultSetDB
        If pstrFieldName2 = "Customer_Code" Then
            strDesSql = "Select " & Trim(pstrFieldName1) & " from " & Trim(pstrTableName) & " WHERE UNIT_CODE='" + gstrUNITID + "'       and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date)) AND  " & Trim(pstrFieldName2) & "='" & Trim(pstrControlText) & "'"
        Else

            strDesSql = "Select " & Trim(pstrFieldName1) & " from " & Trim(pstrTableName) & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  " & Trim(pstrFieldName2) & "='" & Trim(pstrControlText) & "'"
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
    Private Function DecimalAllowedFlag(ByVal pstrItemCode As String) As Boolean
        '----------------------------------------------------------------------------------
        'Description    -   Returns TRUE if Decimals Are Allowed for the Item's UOM
        'Issue ID       -   10 Feb 2009 eMpro-20090209-27201
        '----------------------------------------------------------------------------------
        Dim strUOM As String 'Measurement unit
        Dim bitDecimal As String 'Decimal Allowed 1 or 0
        Dim clsInstMeasure As ClsResultSetDB
        Dim getResultset As New ClsResultSetDB
        Dim strSql As String
        On Error GoTo ErrHandler
        strSql = "Select pur_Measure_Code from Item_Mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  Item_Code = '" & pstrItemCode & "'"
        getResultset.GetResult("Select pur_Measure_Code from Item_Mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  Item_Code = '" & pstrItemCode & "'")
        If getResultset.GetNoRows > 0 Then
            'Get UOM
            strUOM = Trim(getResultset.GetValue("Pur_Measure_Code"))
            clsInstMeasure = New ClsResultSetDB
            clsInstMeasure.GetResult("Select Decimal_Allowed_Flag From Measure_Mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  Measure_Code = '" & strUOM & "'")
            If clsInstMeasure.GetNoRows > 0 Then
                bitDecimal = clsInstMeasure.GetValue("Decimal_Allowed_Flag")
            Else
                bitDecimal = "0"
            End If
            clsInstMeasure.ResultSetClose()
        End If
        getResultset.ResultSetClose() 'Close Resultset
        getResultset = Nothing
        If bitDecimal = "0" Or bitDecimal = "" Then
            'Set Decimal Allowed Flag = False
            DecimalAllowedFlag = False
        Else
            'Set Decimal Allowed Flag = True
            DecimalAllowedFlag = True
        End If
        Exit Function
ErrHandler:
        DecimalAllowedFlag = False
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Public Sub AddBlankRow()
        Dim int_counter As Short
        On Error GoTo ErrHandler
        For int_counter = 1 To frmBatchDetails.fpsprBatch.MaxRows 'Check if Any Blank Rows Previously
            With frmBatchDetails.fpsprBatch
                .Col = 1 : .Col2 = 1 : .Row = int_counter : .Row2 = int_counter
                If .Text = "" Then .Action = FPSpreadADO.ActionConstants.ActionActiveCell : Exit Sub 'If Blank Row is There ,Set focus to its 1st Column
            End With
        Next
        With frmBatchDetails.fpsprBatch
            .MaxRows = .MaxRows + 1 : .set_RowHeight(.MaxRows, 300) ' A new blank row is added and height is set to 300
            .Row = .MaxRows : .Col = 1 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit : .TypeMaxEditLen = 25 : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Row = .MaxRows : .Col = 2 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton : .TypeButtonPicture = My.Resources.ico111.ToBitmap
            .Row = .MaxRows : .Col = 3 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeDate : .TypeDateCentury = True : .TypeDateFormat = FPSpreadADO.TypeDateFormatConstants.TypeDateFormatDDMMYY : .Text = CStr(GetServerDate()) : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            If frmBatchDetails.pDecimalAllowedFlag Then
                .Row = .MaxRows : .Col = 4 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMin = 0.0# : .TypeFloatMax = 99999999.9999 : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .TypeFloatDecimalPlaces = 4 : .Text = CStr(0.0#)
            Else
                .Row = .MaxRows : .Col = 4 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeInteger : .TypeIntegerMin = 0 : .TypeIntegerMax = 99999999 : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight : .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .Text = CStr(0)
            End If
        End With
        With frmBatchDetails.fpsprBatch 'Disable Date Field
            .Col = 3 : .Col2 = 3 : .Row = 1 : .Row2 = .MaxRows : .BlockMode = True : .Protect = True : .Lock = True : .BlockMode = False
            .Col = 1 : .Col2 = 1 : .Row = .ActiveRow : .Action = FPSpreadADO.ActionConstants.ActionActiveCell
        End With
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub SpChEntry_ButtonClicked(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ButtonClickedEvent) Handles SpChEntry.ButtonClicked
        '--------------------------------------------------------------------------------------
        'Created By : Manoj Kr Vaish
        'Issue Id   : eMpro-20090209-27201 10 Feb 2009
        '----------------------------------------------------------------------------------
        Dim frmMKTTRN0009b_SOUTH As New frmMKTTRN0009b_SOUTH
        On Error GoTo Errorhandler
        Dim StrItemCode As String
        Dim dblVal As Double
        Dim BatchStringForItems As String
        Dim intLoopCounter As Short
        Dim varItemCode As Object
        Dim varQty As Object
        Dim varComp_docDetail As Object

        Dim strBatchReq As String 'Counter for For...Next Looping
        If e.col = 21 Then
            With Me.SpChEntry
                .Col = 1 : .Row = e.row : StrItemCode = Trim(.Text)
                If Len(StrItemCode) = 0 Then Exit Sub
                mstrBatchData = ""
                With frmBatchDetails 'Settings According to FORM MODE
                    .pcmdGrpMode = CmdGrpChEnt.Mode
                    .pLocationCode = mstrLocationCode
                    'For Decimal Allowed Flag
                    If DecimalAllowedFlag(StrItemCode) Then
                        'Decimals Are Allowed
                        frmBatchDetails.pDecimalAllowedFlag = True
                    Else
                        'Decimals Are Not Allowed
                        frmBatchDetails.pDecimalAllowedFlag = False
                    End If
                    With Me.SpChEntry
                        .Col = 5 : .Col2 = 5 : .Row = .ActiveRow : frmBatchDetails.lngIssuedQuantity = Val(.Text)
                        .Col = 1 : .Col2 = 1 : .Row = .ActiveRow : .Row2 = .ActiveRow : frmBatchDetails.pItemCodeBat = Trim(.Text)
                    End With
                    .lblItemCode.Text = StrItemCode
                    .Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(Me.FraChEnt.Top)) + 860) 'Adjust The TOP of Form
                    .Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(Me.FraChEnt.Left)) + 4500) 'Adjust The LEFT of Form
                    .fpsprBatch.CursorStyle = FPSpreadADO.CursorStyleConstants.CursorStyleArrow
                    'Check If Their is any Value in BatchDetails UDT
                    If UBound(mBatchData) > 0 Then
                        'Assign This Value to GRID of frmBatchDetails
                        'Check if The Batch Details Values are for Same ItemCode
                        If UBound(mBatchData) >= Me.SpChEntry.ActiveRow Then
                            If mblnbatchfifomode = True Then ReDim mBatchData(Me.SpChEntry.ActiveRow).Batch_No(0)
                            If UBound(mBatchData(Me.SpChEntry.ActiveRow).Batch_No) >= 1 Then
                                For intLoopCounter = 1 To UBound(mBatchData(Me.SpChEntry.ActiveRow).Batch_No)
                                    With frmBatchDetails.fpsprBatch
                                        If Me.CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                                            Call AddBlankRow()
                                        Else
                                            frmBatchDetails.AddBlankRow()
                                            frmBatchDetails.SetSpreadColTypes(intLoopCounter)
                                        End If
                                        .Col = 1 : .Col2 = 1 : .Row = intLoopCounter : .Row2 = intLoopCounter : .Text = mBatchData(Me.SpChEntry.ActiveRow).Batch_No(intLoopCounter)
                                        .Col = 3 : .Col2 = 3 : .Row = intLoopCounter : .Row2 = intLoopCounter : .Text = CStr(mBatchData(Me.SpChEntry.ActiveRow).Batch_Date(intLoopCounter))
                                        .Col = 4 : .Col2 = 4 : .Row = intLoopCounter : .Row2 = intLoopCounter : .Text = CStr(mBatchData(Me.SpChEntry.ActiveRow).Batch_Quantity(intLoopCounter))
                                        frmBatchDetails.lblSumQtyInBatch.Text = CStr(Val(frmBatchDetails.lblSumQtyInBatch.Text) + Val(.Text))
                                        If frmBatchDetails.pDecimalAllowedFlag = True Then
                                            frmBatchDetails.lblSumQtyInBatch.Text = VB6.Format(frmBatchDetails.lblSumQtyInBatch.Text, "0.0000")
                                        Else
                                            frmBatchDetails.lblSumQtyInBatch.Text = VB6.Format(frmBatchDetails.lblSumQtyInBatch.Text, "0")
                                        End If
                                    End With
                                Next
                            ElseIf mblnbatchfifomode = True Then
                                '20 Used For Else condition of IssueBatchDetailsInFIFOMode Select Query
                                dblVal = IssueBatchDetailsInFIFOMode(20, frmBatchDetails.pItemCodeBat, frmBatchDetails.lngIssuedQuantity, Me.SpChEntry.ActiveRow, mstrLocationCode, mBatchData)
                                With Me.SpChEntry
                                    If dblVal > 0 Then
                                        MsgBox("Issued quantity can't exceed current Stock against batches- " & frmBatchDetails.lngIssuedQuantity - dblVal)
                                        .Row = .ActiveRow : .Col = 5 : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus() : Exit Sub
                                    ElseIf dblVal = -1 Then
                                        MsgBox("Batch Details does not exist for the specified item- " & StrItemCode)
                                        .Row = .ActiveRow : .Col = 5 : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus() : Exit Sub
                                    End If
                                End With
                                If UBound(mBatchData) >= Me.SpChEntry.ActiveRow Then
                                    For intLoopCounter = 1 To UBound(mBatchData(Me.SpChEntry.ActiveRow).Batch_No)
                                        With frmBatchDetails.fpsprBatch
                                            If Me.CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                                                Call AddBlankRow()
                                            Else
                                                frmBatchDetails.AddBlankRow()
                                                frmBatchDetails.SetSpreadColTypes(intLoopCounter)
                                            End If
                                            .Col = 1 : .Col2 = 1 : .Row = intLoopCounter : .Row2 = intLoopCounter : .Text = mBatchData(Me.SpChEntry.ActiveRow).Batch_No(intLoopCounter)
                                            .Col = 3 : .Col2 = 3 : .Row = intLoopCounter : .Row2 = intLoopCounter : .Text = CStr(mBatchData(Me.SpChEntry.ActiveRow).Batch_Date(intLoopCounter))
                                            .Col = 4 : .Col2 = 4 : .Row = intLoopCounter : .Row2 = intLoopCounter : .Text = CStr(mBatchData(Me.SpChEntry.ActiveRow).Batch_Quantity(intLoopCounter))
                                            frmBatchDetails.lblSumQtyInBatch.Text = CStr(Val(frmBatchDetails.lblSumQtyInBatch.Text) + Val(.Text))
                                            If frmBatchDetails.pDecimalAllowedFlag = True Then
                                                frmBatchDetails.lblSumQtyInBatch.Text = VB6.Format(frmBatchDetails.lblSumQtyInBatch.Text, "0.0000")
                                            Else
                                                frmBatchDetails.lblSumQtyInBatch.Text = VB6.Format(frmBatchDetails.lblSumQtyInBatch.Text, "0")
                                            End If
                                        End With
                                    Next
                                End If 'End of Check for Change in ItemCodes row
                            ElseIf mblnbatchfifomode = False Then
                                frmBatchDetails.SetSpreadColTypes(1)
                            End If
                        Else
                            If mblnbatchfifomode = True Then
                                dblVal = IssueBatchDetailsInFIFOMode(20, frmBatchDetails.pItemCodeBat, frmBatchDetails.lngIssuedQuantity, Me.SpChEntry.ActiveRow, mstrLocationCode, mBatchData)
                                With Me.SpChEntry
                                    If dblVal > 0 Then
                                        MsgBox("Issued quantity can't exceed current Stock against batches- " & frmBatchDetails.lngIssuedQuantity - dblVal)
                                        .Row = .ActiveRow : .Col = 5 : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus() : Exit Sub
                                    ElseIf dblVal = -1 Then
                                        MsgBox("Batch Details does not exist for the specified item- " & StrItemCode)
                                        .Row = .ActiveRow : .Col = 5 : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus() : Exit Sub
                                    End If
                                End With
                                If UBound(mBatchData) >= Me.SpChEntry.ActiveRow Then
                                    For intLoopCounter = 1 To UBound(mBatchData(Me.SpChEntry.ActiveRow).Batch_No)
                                        With frmBatchDetails.fpsprBatch
                                            If Me.CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                                                Call AddBlankRow()
                                            Else
                                                frmBatchDetails.AddBlankRow()
                                                frmBatchDetails.SetSpreadColTypes(intLoopCounter)
                                            End If
                                            .Col = 1 : .Col2 = 1 : .Row = intLoopCounter : .Row2 = intLoopCounter : .Text = mBatchData(Me.SpChEntry.ActiveRow).Batch_No(intLoopCounter)
                                            .Col = 3 : .Col2 = 3 : .Row = intLoopCounter : .Row2 = intLoopCounter : .Text = CStr(mBatchData(Me.SpChEntry.ActiveRow).Batch_Date(intLoopCounter))
                                            .Col = 4 : .Col2 = 4 : .Row = intLoopCounter : .Row2 = intLoopCounter : .Text = CStr(mBatchData(Me.SpChEntry.ActiveRow).Batch_Quantity(intLoopCounter))
                                            frmBatchDetails.lblSumQtyInBatch.Text = CStr(Val(frmBatchDetails.lblSumQtyInBatch.Text) + Val(.Text))
                                            If frmBatchDetails.pDecimalAllowedFlag = True Then
                                                frmBatchDetails.lblSumQtyInBatch.Text = VB6.Format(frmBatchDetails.lblSumQtyInBatch.Text, "0.0000")
                                            Else
                                                frmBatchDetails.lblSumQtyInBatch.Text = VB6.Format(frmBatchDetails.lblSumQtyInBatch.Text, "0")
                                            End If
                                        End With
                                    Next
                                End If 'End of Check for Change in ItemCodes row
                            End If
                        End If
                    Else
                        frmBatchDetails.SetSpreadColTypes(1)
                    End If
                    .cmdGrpBatch(0).Enabled = False : .cmdGrpBatch(1).Enabled = False : .cmdGrpBatch(2).Enabled = False : .cmdGrpBatch(3).Enabled = True
                    .blncallfrmRGPClose = False 'Calling Form is not RGP closure
                    .blncallfromStckAdjst = False
                    .blnBatchFIFOMode = mblnbatchfifomode
                    frmBatchDetails.pcmdGrpMode = Me.CmdGrpChEnt.Mode
                    Call frmBatchDetails.InitialBatchDetailsSettings()
                    .ShowDialog()
                    System.Windows.Forms.Application.DoEvents()
                    If Me.CmdGrpChEnt.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then 'No Need To Play With Arrays In VIEW Mode
                        If Len(Trim(.mstrBatchRecords)) > 0 Then
                            BatchStringForItems = .mstrBatchRecords
                            Call SplitBatchRecords(BatchStringForItems, Me.SpChEntry.ActiveRow)
                        End If
                    End If
                End With
            End With
        End If
        If e.col = 23 Then
            With SpChEntry
                Dim objFG_Item_code As Object
                objFG_Item_code = Nothing
                varQty = Nothing
                Call .GetText(1, e.row, objFG_Item_code)
                If objFG_Item_code.ToString.Length > 0 Then
                    varQty = Nothing
                    Call .GetText(5, e.row, varQty)
                    If CDbl(varQty) > 0 Then
                        frmCSMDetail.mIntInv_no = CInt(Me.txtChallanNo.Text)
                        frmCSMDetail.mstrFG_item_code = objFG_Item_code.ToString
                        frmCSMDetail.ShowDialog()
                    Else
                        MsgBox("Please Enter Item Qty. First", MsgBoxStyle.Information, ResolveResString(100))
                    End If
                Else
                    MsgBox("Please select item first", MsgBoxStyle.Information, ResolveResString(100))
                End If
            End With
        End If

        If e.col = 25 Then
            With SpChEntry
                varItemCode = Nothing
                .GetText(1, e.row, varItemCode)
                varQty = Nothing
                .GetText(5, e.row, varQty)
                varComp_docDetail = Nothing
                .GetText(25, e.row, varComp_docDetail)
            End With
            If varQty > 0 Then
                CompileDocDetails = ""
                frmMKTTRN0009b_SOUTH.Item_Code = varItemCode
                frmMKTTRN0009b_SOUTH.Item_Desc = Find_Value("Select Description from Item_Mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  Item_Code='" & varItemCode & "'")
                frmMKTTRN0009b_SOUTH.TotalQuantityToBeIssues = varQty
                frmMKTTRN0009b_SOUTH.IsTrans_BatchWise = Find_Value("Select FormLevelBatch_Tracking from FORM_LEVEL_FLAGS WHERE UNIT_CODE='" + gstrUNITID + "' AND  Form_Name='FRMSTRTRN0027'")
                frmMKTTRN0009b_SOUTH.DecimalValue = 4
                frmMKTTRN0009b_SOUTH.Selected_DocNo = txtRefNo.Text
                frmMKTTRN0009b_SOUTH.CompileDocDetails = varComp_docDetail
                If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                    frmMKTTRN0009b_SOUTH.RejectionType = chkRejType.Text
                Else
                    If Find_Value("Select Rej_Type from MKT_INVREJ_DTL WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_No=" & txtChallanNo.Text & " and Item_code='" & varItemCode & "'") = "1" Then
                        frmMKTTRN0009b_SOUTH.RejectionType = "GRN"
                    Else
                        frmMKTTRN0009b_SOUTH.RejectionType = "LRN"
                    End If
                End If
                mstrCompileBatchDetails = ""
                frmMKTTRN0009b_SOUTH.ShowDialog()
                With SpChEntry
                    If Trim(CompileDocDetails) <> "" Then
                        .SetText(24, e.row, CompileDocDetails)
                    End If
                    If mblnBatchTrack = True Then
                        If Len(Trim(CompileDocDetails)) <> 0 Then
                            Call SplitBatchRecords(CompileDocDetails, e.row)
                        End If
                    End If
                End With
            Else
            End If
        End If

        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub SplitBatchRecords(ByVal pstrBatchRecords As String, ByVal mRowNumber As Short)
        '-------------------------------------------------------------------------------------------
        'Created By     -   Sourabh Khatri
        'Description    -   Split Batch Details Information Received from Batch Details Form Row Wise
        '-------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim strBatchRecs() As String
        Dim Intcounter As Short
        strBatchRecs = Split(pstrBatchRecords, "¶")
        If mblnRejTracking = True And mblnBatchTracking = True Then
            If UCase(strInvType) = "REJ" Or UCase(CmbInvType.Text) = "REJECTION" Then
                Call SplitBatchColumnsRejection(strBatchRecs, mRowNumber)
            End If
        Else
            Call SplitBatchColumns(strBatchRecs, mRowNumber)
        End If
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub SplitBatchColumns(ByRef pstrBatchRecords() As String, ByVal mActiveRowNoforArrindex As Short)
        '-------------------------------------------------------------------------------------------
        'Created By     -   Sourabh Khatri
        'Description    -   Split Batch Details Information Received from Batch Details Form Column Wise
        '-------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim strBatchCols() As String
        Dim Intcounter As Short
        If UBound(mBatchData) < mActiveRowNoforArrindex Then 'If New Row is Added In Issue GRID
            ReDim Preserve mBatchData(mActiveRowNoforArrindex)
        End If
        For Intcounter = 0 To UBound(pstrBatchRecords) - 1
            strBatchCols = Split(pstrBatchRecords(Intcounter), "§")
            ReDim Preserve mBatchData(Me.SpChEntry.ActiveRow).Batch_No(Intcounter + 1)
            mBatchData(Me.SpChEntry.ActiveRow).Batch_No(Intcounter + 1) = strBatchCols(0)
            ReDim Preserve mBatchData(Me.SpChEntry.ActiveRow).Batch_Date(Intcounter + 1)
            mBatchData(Me.SpChEntry.ActiveRow).Batch_Date(Intcounter + 1) = CDate(strBatchCols(1))
            ReDim Preserve mBatchData(Me.SpChEntry.ActiveRow).Batch_Quantity(Intcounter + 1)
            mBatchData(Me.SpChEntry.ActiveRow).Batch_Quantity(Intcounter + 1) = CDbl(strBatchCols(2))

        Next
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub SpChEntry_Change(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ChangeEvent) Handles SpChEntry.Change
        On Error GoTo ErrHandler
        Dim intRowCount As Short
        Dim intmaxrows As Short
        Dim rsItemMst As ClsResultSetDB
        Dim varFromBox As Object
        Dim varItem As Object
        Dim VarToBox As Object
        Dim varQty As Object
        Dim boxqty As Double
        Dim varCumulativeBoxes As Object
        Dim varCustItem As Object
        Dim objComm As New ADODB.Command
        Dim varMaxQty As Object
        Dim strExceptionMsg As String
        Dim rsExceptiondtl As ClsResultSetDB
        Dim introws As Short


        With SpChEntry
            If e.col = 5 Then
                With SpChEntry
                    intmaxrows = SpChEntry.MaxRows
                    For intRowCount = 1 To intmaxrows
                        varItem = Nothing
                        Call .GetText(1, intRowCount, varItem)
                        varCustItem = Nothing
                        Call .GetText(2, intRowCount, varCustItem) 'Added by arul on 27-04-2005
                        varQty = Nothing
                        Call .GetText(5, intRowCount, varQty)
                        If mblnRejTracking = True And CmbInvType.Text = "REJECTION" Then
                            varMaxQty = Nothing
                            Call SpChEntry.GetText(22, e.row, varMaxQty)
                            If CDbl(varQty) > CDbl(varMaxQty) Then
                                MsgBox("Quantity should be greater than " & varMaxQty, MsgBoxStyle.Information, ResolveResString(100))
                                Call SpChEntry.SetText(5, e.row, varMaxQty)
                            End If
                        End If
                        rsItemMst = New ClsResultSetDB

                        If UCase(CmbInvType.Text) = "NORMAL INVOICE" Then
                            rsItemMst.GetResult("Select Container_qty Box_Qty from custitem_mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  Account_Code = '" & Trim(txtCustCode.Text) & "' and ITem_code = '" & varItem & "' and Cust_Drgno = '" & varCustItem & "' and Active = 1 ")
                        Else
                            rsItemMst.GetResult("Select Box_Qty from ITem_Mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  ITem_code ='" & varItem & "'")
                        End If
                        'Changes Ends here
                        boxqty = Val(rsItemMst.GetValue("Box_Qty"))
                        rsItemMst.ResultSetClose()
                        If boxqty > 0 Then
                            If varQty > 0 Then
                                If (varQty / boxqty) - Int(varQty / boxqty) > 0 Then
                                    '
                                    If intRowCount = 1 Then
                                        Call .SetText(12, intRowCount, 1)
                                        Call .SetText(13, intRowCount, Int(varQty / boxqty) + 1)
                                        Call .SetText(14, intRowCount, ((Int(varQty / boxqty) + 1) - 1) + 1)
                                    Else
                                        VarToBox = Nothing
                                        Call .GetText(13, intRowCount - 1, VarToBox)
                                        varCumulativeBoxes = Nothing
                                        Call .GetText(14, intRowCount - 1, varCumulativeBoxes)
                                        Call .SetText(12, intRowCount, VarToBox + 1)
                                        Call .SetText(13, intRowCount, VarToBox + ((Int(varQty / boxqty)) + 1))
                                        Call .SetText(14, intRowCount, Val(varCumulativeBoxes) + ((Int((System.Math.Round(varQty / boxqty)) + 1) - Val(varFromBox + 1)) + 1))
                                    End If
                                Else
                                    If intRowCount = 1 Then
                                        Call .SetText(12, intRowCount, 1)
                                        Call .SetText(13, intRowCount, (Int(varQty / boxqty)))
                                        Call .SetText(14, intRowCount, ((Int(varQty / boxqty) + 1) - 1))
                                    Else
                                        VarToBox = Nothing
                                        Call .GetText(13, intRowCount - 1, VarToBox)
                                        Call .SetText(12, intRowCount, VarToBox + 1)
                                        Call .SetText(13, intRowCount, VarToBox + Int(varQty / boxqty))
                                        Call .SetText(14, intRowCount, Val(varCumulativeBoxes) + ((Val(CStr(Int(varQty / boxqty) + 1)) - Val(varFromBox + 1))))
                                    End If
                                End If
                            End If
                        End If
                        rsExceptiondtl = New ClsResultSetDB
                        If ((UCase(CmbInvType.Text) = "NORMAL INVOICE" And UCase(CmbInvSubType.Text) = "FINISHED GOODS") Or
                            (strInvType = "INV" And strInvSubType = "F")) And mblnCSM_Knockingoff_req Then
                            If CInt(varQty) > 0 Then

                                With objComm
                                    .ActiveConnection = mP_Connection
                                    .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                                    .CommandText = "USP_GET_CSM_CALCULATION_FOR_INVOICE"
                                    .CommandTimeout = 0
                                    .Parameters.Append(.CreateParameter("@UNITCODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                                    .Parameters.Append(.CreateParameter("@CUSTOMER_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 16, txtCustCode.Text.ToString))
                                    .Parameters.Append(.CreateParameter("@ITEM_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 16, varItem.ToString))
                                    .Parameters.Append(.CreateParameter("@CUST_DRGNO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 40, varCustItem.ToString))
                                    .Parameters.Append(.CreateParameter("@INV_QTY", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, , CInt(varQty)))
                                    .Parameters.Append(.CreateParameter("@INV_NO", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, , CInt(txtChallanNo.Text)))
                                    .Parameters.Append(.CreateParameter("@CSM_AMT", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInputOutput, , 0))
                                    .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                    If .Parameters(.Parameters.Count - 1).Value = "0" Then

                                        strExceptionMsg = ""

                                        rsExceptiondtl.GetResult("SELECT FG_ITEM_CODE,CSM_ITEM_CODE,REMAINED_QTY FROM CSM_KNOCKOFF_DTL_EXCEPTION  WHERE UNIT_CODE='" + gstrUNITID + "' AND INV_NO = " & txtChallanNo.Text & " ")
                                        If rsExceptiondtl.GetNoRows > 0 Then
                                            rsExceptiondtl.MoveFirst()
                                            While Not rsExceptiondtl.EOFRecord
                                                strExceptionMsg = strExceptionMsg + "FINISHED ITEM_CODE : " & rsExceptiondtl.GetValue("FG_ITEM_CODE") & "  CSI ITEM CODE :" & rsExceptiondtl.GetValue("CSM_ITEM_CODE") & " MORE REQUIRED QTY :" & rsExceptiondtl.GetValue("REMAINED_QTY ") & vbCrLf
                                                rsExceptiondtl.MoveNext()
                                            End While
                                        End If
                                        MsgBox("Insufficent CSM Stock, Invoice Can't Be Saved." & vbCrLf & strExceptionMsg, MsgBoxStyle.Information, ResolveResString(100))
                                        Call SpChEntry.SetText(4, intRowCount, "")
                                        Call SpChEntry.SetText(5, intRowCount, "0.00")
                                    ElseIf .Parameters(.Parameters.Count - 1).Value = "-1" Then
                                        Call SpChEntry.SetText(4, intRowCount, "0")
                                    Else
                                        Call SpChEntry.SetText(4, intRowCount, CDbl(.Parameters(.Parameters.Count - 1).Value))
                                    End If
                                End With
                                objComm.Parameters.Delete(6)
                                objComm.Parameters.Delete(5)
                                objComm.Parameters.Delete(4)
                                objComm.Parameters.Delete(3)
                                objComm.Parameters.Delete(2)
                                objComm.Parameters.Delete(1)
                                objComm.Parameters.Delete(0)
                            End If
                            rsExceptiondtl.ResultSetClose()
                            rsExceptiondtl = Nothing
                        End If
                    Next
                    objComm = Nothing

                End With
            End If
            If (e.col = 12) Or (e.col = 13) Then
                intmaxrows = SpChEntry.MaxRows
                For intRowCount = 1 To intmaxrows
                    varFromBox = Nothing
                    Call .GetText(12, intRowCount, varFromBox)
                    VarToBox = Nothing
                    Call .GetText(13, intRowCount, VarToBox)
                    If intRowCount = 1 Then
                        If Len(Trim(varFromBox)) Then
                            If Len(Trim(VarToBox)) Then
                                Call .SetText(14, intRowCount, (Val(VarToBox) - Val(varFromBox)) + 1)
                            End If
                        End If
                    Else
                        varCumulativeBoxes = Nothing
                        Call .GetText(14, intRowCount - 1, varCumulativeBoxes)
                        If Len(Trim(varCumulativeBoxes)) Then
                            If Len(Trim(varFromBox)) Then
                                If Len(Trim(VarToBox)) Then
                                    Call .SetText(14, intRowCount, varCumulativeBoxes + ((Val(VarToBox) - Val(varFromBox)) + 1))
                                End If
                            End If
                        End If
                    End If
                Next
            End If

            If (e.col = 7) Then
                Dim varExcise As Object
                Dim Validexcisetype As String
                intmaxrows = SpChEntry.MaxRows
                For intRowCount = 1 To intmaxrows
                    varExcise = Nothing
                    Call .GetText(7, intRowCount, varExcise)
                    Validexcisetype = Find_Value("select TxRt_Rate_No from Gen_TaxRate where Tx_TaxeID='EXC' and unit_code='" & gstrUNITID & "' and TxRt_Rate_No = '" & varExcise & "'")
                    If Validexcisetype = "0" Then
                        .Text = ""
                        MsgBox("Invalid Excise Type, Press F1 for help.", MsgBoxStyle.Information, ResolveResString(100))
                        Exit Sub
                    End If
                Next
            End If

        End With

        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub txtamendno_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAmendNo.TextChanged
        If Trim(txtAmendNo.Text) = "" Then
            If Me.CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                SpChEntry.MaxRows = 0
                mstrItemCode = ""
                PaletteActiveInActive()
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
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        If Len(txtAmendNo.Text) > 0 Then
                            Call txtAmendNo_Validating(txtAmendNo, New System.ComponentModel.CancelEventArgs(False))
                        Else
                            If (CmbInvType.Text = "JOBWORK INVOICE") Then
                                'jul
                                'txtAnnex.SetFocus
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
        Dim rsObjTax As New ADODB.Recordset
        'Only if Some Ref No. is Added
        If Trim(txtRefNo.Text) <> "" Then
            'if Some Amend No is Entered
            If Trim(txtAmendNo.Text) <> "" Then
                If SelectDataFromTable("Amendment_No", "Cust_ORD_HDR", " Where UNIT_CODE='" + gstrUNITID + "' AND Account_Code = '" & Trim(txtCustCode.Text) & "' And Cust_Ref = '" & Trim(txtRefNo.Text) & "' And Active_Flag = 'A'  AND  Amendment_No <> '' AND  Amendment_No = '" & Trim(txtAmendNo.Text) & "'") <> "" Then
                    'Verified,Set focus to Another Control
                    If (CmbInvType.Text = "JOBWORK INVOICE") Then

                    Else
                        If rsObjTax.State = ADODB.ObjectStateEnum.adStateOpen Then rsObjTax.Close()
                        'Code Changed By Arul on 12-01-2005 to add the Ecess Field
                        rsObjTax.Open("Select isnull(SalesTax_Type,0),isnull(Surcharge_code,0),isnull(ECESS_Code,'') ECESS_Code from cust_ord_hdr WHERE UNIT_CODE='" + gstrUNITID + "' AND  Amendment_No='" & Trim(txtAmendNo.Text) & "' and Cust_Ref ='" & Trim(txtRefNo.Text) & "' and Account_Code='" & Trim(txtCustCode.Text) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        If Not rsObjTax.EOF Then
                            txtSaleTaxType.Text = Trim(rsObjTax.Fields(0).Value)
                            txtSurchargeTaxType.Text = Trim(rsObjTax.Fields(1).Value)
                            txtECSSTaxType.Text = Trim(rsObjTax.Fields(2).Value)
                            Call FN_ECSSH_TAX_Disp()
                        End If
                        txtCarrServices.Focus()
                    End If
                Else
                    MsgBox("Entered Amendment Number for Ref No." & txtRefNo.Text & vbCr & " does not Exist or is Not Active.", MsgBoxStyle.Information, "empower")
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

    Private Function FN_ECSSH_TAX_Disp() As Object
        On Error GoTo err_Renamed
        Dim sql As String
        Dim adors As New ADODB.Recordset
        '------------------Satvir Handa------------------------
        'sql = "select TxRt_Rate_No,TxRt_Percentage from Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  Tx_TaxeID='ECSSH' and TxRt_Percentage=1"
        sql = "select TxRt_Rate_No,TxRt_Percentage from Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  Tx_TaxeID='ECSSH' and DEFAULT_FOR_INVOICE =1 and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"
        '------------------Satvir Handa------------------------
        adors.Open(sql, mP_Connection, ADODB.CursorTypeEnum.adOpenStatic)
        If adors.EOF = False Then
            txtSECSSTaxType.Text = IIf(IsDBNull(adors.Fields("TxRt_Rate_No").Value), "", adors.Fields("TxRt_Rate_No").Value)
            Call txtSECSSTaxType_Validating(txtSECSSTaxType, New System.ComponentModel.CancelEventArgs(False))
        End If
        If adors.State Then adors.Close()
        adors = Nothing
        Exit Function
err_Renamed:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Function
    End Function
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
        Dim strInvoiceSubType As String
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
                                rsChallanEntry.GetResult("Select a.Description,a.Sub_Type_Description from SaleConf a (nolock) ,SalesChallan_Dtl b where A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" + gstrUNITID + "' AND Doc_No = " & txtChallanNo.Text & " and a.Invoice_Type = b.Invoice_type and a.Sub_type = b.Sub_Category and a.Location_code = b.Location_code and (fin_start_date <= getdate() and fin_end_date >= getdate())")
                                strInvoiceType = UCase(rsChallanEntry.GetValue("Description"))
                                strInvoiceSubType = UCase(rsChallanEntry.GetValue("sub_type_Description"))
                                mstrInvsubTypeDesc = UCase(rsChallanEntry.GetValue("Sub_Type_Description"))
                                Call CheckBatchTrackingAllowed(strInvoiceType, strInvoiceSubType)
                                rsChallanEntry.ResultSetClose()
                                If UCase(strInvoiceType) <> "SAMPLE INVOICE" Then
                                    With SpChEntry
                                        .Col = 16 : .Col2 = 16 : .BlockMode = True : .ColHidden = True : .BlockMode = False
                                    End With
                                Else
                                    With SpChEntry
                                        .Col = 16 : .Col2 = 16 : .BlockMode = True : .ColHidden = False : .BlockMode = False
                                        .Col = 16 : .Col2 = 16 : .BlockMode = True : .Lock = False : .BlockMode = False
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
        If txtChallanNo.Text.Trim <> "" Then
            If Val(txtChallanNo.Text.Trim.Substring(0, 2)) = 99 Then
                Cmditems.Enabled = True
            Else
                CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
            End If
        Else
            CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
            CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
        End If

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
            lblexportsodetails.Text = ""
            fraRGPs.Visible = False
        End If
        If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            txtCustCode.Focus()
            PaletteActiveInActive()
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
                            If (UCase(CmbInvType.Text) = "NORMAL INVOICE") Or (UCase(CmbInvType.Text) = "JOBWORK INVOICE") Or (UCase(CmbInvType.Text) = "EXPORT INVOICE") Then
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
    Private Sub txtCustCode_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCustCode.KeyUp
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
        Dim strSQL As String
        Dim blnNTRF_INV_GROUPCOMP As Boolean = False
        Dim strcondGroupcompany As String
        Dim blntcscheck As Boolean

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
                    If UCase(Trim(mstrInvoiceType)) = "INV" Or UCase(Trim(mstrInvoiceType)) = "SMP" Or UCase(Trim(mstrInvoiceType)) = "TRF" Or UCase(Trim(mstrInvoiceType)) = "JOB" Or UCase(Trim(mstrInvoiceType)) = "EXP" Or BLNREJECTION_FLAG = True Or UCase(Trim(mstrInvoiceType)) = "SRC" Or UCase(Trim(mstrInvoiceType)) = "ITD" Then
                        If CheckExistanceOfFieldData((txtCustCode.Text), "Customer_Code", "Customer_Mst", "  ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))" & strcondGroupcompany & "") Then
                            Call SelectDescriptionForField("Cust_Name", "Customer_Code", "Customer_Mst", lblCustCodeDes, Trim(txtCustCode.Text))
                            Call SetControlsforASNDetails(Trim(txtCustCode.Text))
                            '10856126
                            strSQL = "select dbo.UDF_ISDOCKCODE_INVOICE( '" & gstrUNITID & "','" & txtCustCode.Text.Trim & "','" & CmbInvType.Text.Trim.ToUpper & "','" & CmbInvSubType.Text.Trim.ToUpper & "' )"
                            If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSQL)) = True Then
                                TxtLRNO.MaxLength = 11
                            Else
                                TxtLRNO.MaxLength = 30
                            End If
                            '10856126

                            If (UCase(CmbInvType.Text) = "NORMAL INVOICE") Or (UCase(CmbInvType.Text) = "JOBWORK INVOICE") Or (UCase(CmbInvType.Text) = "EXPORT INVOICE") Or (UCase(CmbInvType.Text) = "TRANSFER INVOICE") Or (UCase(CmbInvType.Text) = "SERVICE INVOICE") Or (UCase(CmbInvType.Text) = "INTER-DIVISION") Then
                                If UCase(CmbInvSubType.Text) <> "SCRAP" Then
                                    txtRefNo.Focus()
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
                            txtCustCode.Focus()
                        End If
                        '***To Display invoice Address of Customer
                        If Len(Trim(txtCustCode.Text)) > 0 Then
                            mblnCSM_Knockingoff_req = DataExist("SELECT TOP 1 1 FROM CUSTOMER_MST WHERE CSM_FLAG = 1 AND CUSTOMER_CODE='" & txtCustCode.Text & "' and Unit_code='" & gstrUNITID & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))")
                            rsCustMst = New ClsResultSetDB
                            strCustMst = "SELECT Bill_Address1 + ', '  +  Bill_Address2 + ', ' + Bill_City + ' - ' + Bill_Pin as  invoiceAddress from Customer_Mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  Customer_code ='" & txtCustCode.Text & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
                            rsCustMst.GetResult(strCustMst)
                            If rsCustMst.GetNoRows > 0 Then
                                lblAddressDes.Text = rsCustMst.GetValue("InvoiceAddress")
                            End If
                            rsCustMst.ResultSetClose()
                        End If
                        '***
                    Else
                        If CheckExistanceOfFieldData((txtCustCode.Text), "Vendor_Code", "Vendor_Mst") Then
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
                            If MsgBox("Would Like to Follow FIFO Method For JobWork Material Process.", MsgBoxStyle.YesNo, "empower") = 7 Then
                                blnFIFO = False
                                mstrRGP = ""
                                If AddDataTolstRGPs() = True Then
                                    fraRGPs.Visible = True
                                Else
                                    MsgBox("No RGP's in last 180 days for this Customer.", MsgBoxStyle.Information, "empower")
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
                If Len(Trim(txtCustCode.Text)) > 0 Then
                    strSQL = "select dbo.UDF_IRN_TCSREQUIRED( '" & gstrUNITID & "','" & txtCustCode.Text.Trim & "')"
                    If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSQL)) = True Then
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
                            Else
                                txtTCSTaxCode.Enabled = False : txtTCSTaxCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdHelpTCSTax.Enabled = False : txtTCSTaxCode.Text = ""
                            End If
                        Else
                            txtTCSTaxCode.Enabled = False : txtTCSTaxCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdHelpTCSTax.Enabled = False : txtTCSTaxCode.Text = ""
                        End If
                    End If
                End If

                PaletteActiveInActive()
        End Select
        mstrcustcompcode = Find_Value("Select isnull(Customer_CompanyCode,'')Customer_CompanyCode from EMPRO_FA_CUSTOMER_UNIT_ATNMAPPING where EMPRO_UNIT='" + gstrUNITID + "' AND  customer_code='" & Trim(txtCustCode.Text) & "'and company_code='" & mstrcurrentATNCode & "'")
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
        Dim STRSQL As String
        '10808160 
        mstrEOP_Required = False
        '10808160 
        strSaleConfSql = "Select Invoice_Type,Sub_Type , EOP_Required from SaleConf (nolock) WHERE UNIT_CODE='" + gstrUNITID + "' AND  Description='" & Trim(pstrInvType) & "'"
        strSaleConfSql = strSaleConfSql & " and Sub_Type_Description='" & Trim(pstrInvSubtype) & "' and (fin_start_date <= getdate() and fin_end_date >= getdate()) "
        rsSaleConf = New ClsResultSetDB
        rsSaleConf.GetResult(strSaleConfSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsSaleConf.GetNoRows > 0 Then
            mstrInvoiceType = rsSaleConf.GetValue("Invoice_Type")
            mstrInvoiceSubType = rsSaleConf.GetValue("Sub_Type")
            '10808160 
            mstrEOP_Required = rsSaleConf.GetValue("EOP_Required")
            '10808160 
        Else
            mstrInvoiceType = ""
            mstrInvoiceSubType = ""
        End If
        rsSaleConf.ResultSetClose()
        rsSaleConf = Nothing
        'Me.SpChEntry.MaxCols = 23
        With SpChEntry
            If gblnGSTUnit = True Then
                .MaxCols = 35
                .Row = 0 : .Col = 30 : .Text = "HSN/SAC CODE" : .set_ColWidth(30, 1000)
                .Row = 0 : .Col = 31 : .Text = "CGST TAX" : .set_ColWidth(31, 1000)
                .Row = 0 : .Col = 32 : .Text = "SGST TAX" : .set_ColWidth(32, 1000)
                .Row = 0 : .Col = 33 : .Text = "UTGST TAX" : .set_ColWidth(33, 1000)
                .Row = 0 : .Col = 34 : .Text = "IGST TAX" : .set_ColWidth(34, 1000)
                .Row = 0 : .Col = 35 : .Text = "COMPENSATION CESS" : .set_ColWidth(35, 1400)

            Else
                Me.SpChEntry.MaxCols = 27
            End If
        End With

        'Me.SpChEntry.MaxCols = 27

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
                    If CheckExistanceOfFieldData((txtLocationCode.Text), "Location_Code", "Saleconf") Then
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
    Private Sub TxtLRNO_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TxtLRNO.Enter
        On Error GoTo ErrHandler
        'Selecting the text in the text box
        With TxtLRNO
            .SelectionStart = 0
            .SelectionLength = Len(Trim(.Text))
        End With
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub TxtLRNO_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles TxtLRNO.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Dim strSQL As String
        On Error GoTo ErrHandler
        If KeyAscii = 14 Then
            System.Windows.Forms.SendKeys.Send("{tab}")
        End If
        strSQL = "select dbo.UDF_ISEOPINVOICE( '" & gstrUNITID & "','" & txtCustCode.Text.Trim & "','" & CmbInvType.Text.Trim & "','" & CmbInvSubType.Text.Trim & "','" & txtRefNo.Text.Trim & "')"
        If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSQL)) = True Then
            AllowNumericValueInTextBox(TxtLRNO, eventArgs)
        End If

        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtRefNo_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRefNo.TextChanged
        If Me.CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            SpChEntry.MaxRows = 0 : mstrItemCode = "" : If txtRefNo.Enabled = True Then txtRefNo.Focus()
            txtSaleTaxType.Text = ""
            txtSurchargeTaxType.Text = ""
            txtAmendNo.Text = ""
            PaletteActiveInActive()
        End If
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
            If mblnRejTracking = True And CmbInvType.Text = "REJECTION" Then
                GoTo EventExitSub
            End If
            If Len(txtRefNo.Text) > 0 Then
                If SelectDataFromCustOrd_Dtl((txtCustCode.Text), (CmbInvType.Text)) Then

                Else
                    If CmbInvType.Text <> "REJECTION" Then
                        Call ConfirmWindow(10436, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                    Else
                        MsgBox("GRIN No Entered by you is inValid,Press F1 for Help.", MsgBoxStyle.Information, "empower")
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
        eventArgs.Cancel = False
    End Sub
    Private Sub txtRemarks_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtRemarks.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        With Me.SpChEntry
                            .Row = 1 : .Col = 5 : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
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
        Dim rsObjTx As New ADODB.Recordset
        If Len(txtSaleTaxType.Text) = 0 Then
            lblSaltax_Per.Text = "0.00"
        Else
            If rsObjTx.State = ADODB.ObjectStateEnum.adStateOpen Then rsObjTx.Close()
            rsObjTx.Open("SELECT txrt_rate_no,txrt_percentage FROM gen_taxrate WHERE UNIT_CODE='" + gstrUNITID + "' AND  txrt_rate_no='" & Trim(txtSaleTaxType.Text) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If Not rsObjTx.EOF Then
                txtSaleTaxType.Text = Trim(rsObjTx.Fields(0).Value)
                lblSaltax_Per.Text = Trim(rsObjTx.Fields(1).Value)
            End If
        End If
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
                            txtSurchargeTaxType.Focus()
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
        If Len(txtSaleTaxType.Text) > 0 Then
            If CheckExistanceOfFieldData((txtSaleTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", " (Tx_TaxeID='CST' OR Tx_TaxeID='LST' or Tx_TaxeID='VAT')") Then
                lblSaltax_Per.Text = CStr(GetTaxRate((txtSaleTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " (Tx_TaxeID='CST' OR Tx_TaxeID='LST' or Tx_TaxeID='VAT')"))
                If txtSurchargeTaxType.Enabled Then txtSurchargeTaxType.Focus()
            Else
                MsgBox("Invalid Sale Tax Code, Press F1 for help.", MsgBoxStyle.Information, ResolveResString(100))
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
    Private Sub txtSECSSTaxType_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSECSSTaxType.TextChanged
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
        If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then Command1.PerformClick()
    End Sub
    Private Sub txtSECSSTaxType_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtSECSSTaxType.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        If Len(txtSECSSTaxType.Text) > 0 And gblnGSTUnit = False Then
            '------------------Satvir Handa------------------------
            If CheckExistanceOfFieldData((txtSECSSTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", " (Tx_TaxeID='ECSSH') and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))") Then
                '------------------Satvir Handa------------------------
                lblSECSStax_Per.Text = CStr(GetTaxRate((txtSECSSTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " (Tx_TaxeID='ECSSH')"))
            Else
                MsgBox("Invalid SE. Cess Code, Press F1 for help.", MsgBoxStyle.Information, ResolveResString(100))
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
    Private Sub txtSurchargeTaxType_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSurchargeTaxType.TextChanged
        Dim rsObjTx As New ADODB.Recordset
        If Trim(txtSurchargeTaxType.Text) = "" Then
            lblSurcharge_Per.Text = "0.00"
        Else
            If rsObjTx.State = ADODB.ObjectStateEnum.adStateOpen Then rsObjTx.Close()
            rsObjTx.Open(" SELECT txrt_rate_no,txrt_percentage FROM gen_taxrate WHERE UNIT_CODE='" + gstrUNITID + "' AND  txrt_rate_no='" & Trim(txtSurchargeTaxType.Text) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If Not rsObjTx.EOF Then
                txtSurchargeTaxType.Text = Trim(rsObjTx.Fields(0).Value)
                lblSurcharge_Per.Text = Trim(rsObjTx.Fields(1).Value)
            End If
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
                            ctlPerValue.Focus()
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
            If CheckExistanceOfFieldData((txtSurchargeTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", " Tx_TaxeID='SST'") Then
                lblSurcharge_Per.Text = CStr(GetTaxRate((txtSurchargeTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='SST'"))
                If SpChEntry.Enabled Then
                    With Me.SpChEntry
                        .Row = 1 : .Col = 5 : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                    End With
                End If
            Else
                MsgBox("Invalid Surcharge Code, Press F1 for help.", MsgBoxStyle.Information, ResolveResString(100))
                Cancel = True
                txtSurchargeTaxType.Text = ""
                If txtSurchargeTaxType.Enabled Then txtSurchargeTaxType.Focus()
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
    Private Sub txtVehNo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtVehNo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        If KeyAscii = 13 Then System.Windows.Forms.SendKeys.Send("{tab}")
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
            strTableSql = "select " & Trim(pstrColumnName) & " from " & Trim(pstrTableName) & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  " & Trim(pstrColumnName) & "='" & Trim(pstrFieldText) & "' and " & pstrCondition
        Else
            strTableSql = "select " & Trim(pstrColumnName) & " from " & Trim(pstrTableName) & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  " & Trim(pstrColumnName) & "='" & Trim(pstrFieldText) & "'"
        End If
        rsExistData = New ClsResultSetDB
        rsExistData.GetResult(strTableSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsExistData.GetNoRows > 0 Then
            CheckExistanceOfFieldData = True
        Else
            CheckExistanceOfFieldData = False
        End If
        rsExistData.ResultSetClose()
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
        Dim strSql As String

        '------------------Satvir Handa------------------------
        strSalesChallanDtl = "SELECT SECESS_Type,SECESS_Per,Transport_type,Vehicle_No,Account_Code,Cust_ref,Amendment_No,SalesTax_Type,"
        strSalesChallanDtl = strSalesChallanDtl & "Insurance,Invoice_Date,tot_add_excise_amt,"
        strSalesChallanDtl = strSalesChallanDtl & "Invoice_Type,Sub_Category,Cust_Name,Carriage_Name,Frieght_Amount, "
        strSalesChallanDtl = strSalesChallanDtl & "Surcharge_salesTaxType,Amendment_No,ref_doc_no,Currency_Code,Exchange_Rate,Remarks,PerValue ,LorryNo_Date,ECESS_Type,payment_terms,TCSTax_Type ,TCSTax_Per ,TCSTaxAmount,ADDVAT_Amount ,ADDVAT_Type ,ADDVAT_Per ,ServiceTax_Type,ServiceTax_Per,ServiceTax_Amount ,DELIVERYNOTENO From Saleschallan_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  Location_Code ='"
        strSalesChallanDtl = strSalesChallanDtl & Trim(txtLocationCode.Text) & "' and Doc_No = " & Val(txtChallanNo.Text)
        rsGetData = New ClsResultSetDB
        rsGetData.GetResult(strSalesChallanDtl, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsGetData.GetNoRows > 0 Then
            GetDataInViewMode = True
            txtCustCode.Text = rsGetData.GetValue("Account_Code")
            txtRefNo.Text = rsGetData.GetValue("Cust_ref")
            txtAmendNo.Text = rsGetData.GetValue("Amendment_No")
            txtCarrServices.Text = rsGetData.GetValue("Carriage_Name")
            ctlInsurance.Text = rsGetData.GetValue("Insurance")
            txtFreight.Text = rsGetData.GetValue("Frieght_Amount")
            txtSaleTaxType.Text = rsGetData.GetValue("SalesTax_Type")
            aedamnt.Text = rsGetData.GetValue("tot_add_excise_amt")
            txtAddVAT.Text = rsGetData.GetValue("ADDVAT_type")
            lblAddVAT.Text = rsGetData.GetValue("ADDVAT_per")
            Call txtSaleTaxType_Validating(txtSaleTaxType, New System.ComponentModel.CancelEventArgs(False))
            txtSurchargeTaxType.Text = rsGetData.GetValue("Surcharge_salesTaxType")
            Call txtSurchargeTaxType_Validating(txtSurchargeTaxType, New System.ComponentModel.CancelEventArgs(False))
            txtECSSTaxType.Text = rsGetData.GetValue("ECESS_Type")
            Call txtECSSTaxType_Validating(txtECSSTaxType, New System.ComponentModel.CancelEventArgs(False))

            '------------------Satvir Handa------------------------
            'Call FN_ECSSH_TAX_Disp()
            txtSECSSTaxType.Text = rsGetData.GetValue("SECESS_Type")
            lblSECSStax_Per.Text = rsGetData.GetValue("SECESS_Per")
            '------------------Satvir Handa------------------------

            txtTCSTaxCode.Text = rsGetData.GetValue("TCSTax_Type")
            lblTCSTaxPerDes.Text = rsGetData.GetValue("TCSTax_Per")
            strRGPNOs = rsGetData.GetValue("ref_doc_no")
            strRGPNOs = Replace(strRGPNOs, "§", ", ", 1)
            lblRGPDes.Text = strRGPNOs
            lblCustCodeDes.Text = rsGetData.GetValue("Cust_Name")
            mstrAmmendmentNo = rsGetData.GetValue("Amendment_No")
            dtpDateDesc.Value = VB6.Format(rsGetData.GetValue("Invoice_Date"), "dd MMM yyyy")
            lblDateDes.Text = VB6.Format(rsGetData.GetValue("Invoice_Date"), gstrDateFormat)

            mstrInvType = rsGetData.GetValue("Invoice_Type")
            mstrInvoiceType = rsGetData.GetValue("Invoice_Type")
            mstrInvSubType = rsGetData.GetValue("Sub_Category")
            mstrInvoiceSubType = rsGetData.GetValue("Sub_Category")
            ctlPerValue.Text = rsGetData.GetValue("PerValue")
            TxtLRNO.Text = rsGetData.GetValue("LorryNo_Date")
            txtVehNo.Text = rsGetData.GetValue("Vehicle_No")
            txtDeliveryNoteNo.Text = rsGetData.GetValue("DELIVERYNOTENO")
            '10869290
            If UCase(mstrInvType) = "SRC" Then
                txtServiceTaxType.Text = rsGetData.GetValue("ServiceTax_Type")
                lblServiceTax_Per.Text = rsGetData.GetValue("ServiceTax_per")
            End If


            If UCase(mstrInvType) = "EXP" Or UCase(mstrInvType) = "SRC" Then
                lblCurrency.Visible = True : lblCurrencyDes.Visible = True
                lblCurrencyDes.Text = rsGetData.GetValue("Currency_code")
                lblExchangeRateLable.Visible = True : lblExchangeRateValue.Visible = True
            Else
                lblCurrency.Visible = False : lblCurrencyDes.Visible = False
                lblExchangeRateLable.Visible = False : lblExchangeRateValue.Visible = False
            End If
            txtRemarks.Text = rsGetData.GetValue("Remarks")
            lblcreditterm.Text = IIf(IsDBNull(rsGetData.GetValue("Payment_terms")), "", rsGetData.GetValue("Payment_terms"))
        Else
            GetDataInViewMode = False
        End If
        '***To Display invoice Address of Customer
        If Len(Trim(txtCustCode.Text)) > 0 Then
            rsCustMst = New ClsResultSetDB
            strCustMst = "SELECT Bill_Address1 + ', '  +  Bill_Address2 + ', ' + Bill_City + ' - ' + Bill_Pin as  invoiceAddress from Customer_Mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  Customer_code ='" & txtCustCode.Text & "'"
            rsCustMst.GetResult(strCustMst)
            If rsCustMst.GetNoRows > 0 Then
                lblAddressDes.Text = rsCustMst.GetValue("InvoiceAddress")
            End If
            strCustMst = "select cust_plantcode,arl_code from mkt_asn_invdtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  doc_no=" & Val(txtChallanNo.Text)
            rsCustMst.GetResult(strCustMst, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            If rsCustMst.GetNoRows > 0 Then
                txtPlantCode.Text = rsCustMst.GetValue("cust_plantcode")
                txtActualReceivingLoc.Text = rsCustMst.GetValue("arl_code")
            End If
            rsCustMst.ResultSetClose()
        End If
        '***
        If mblnRejTracking = True Then
            strSql = "Select Ref_Doc_No from MKT_INVREJ_DTL " & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_No = " & Trim(txtChallanNo.Text)
            rsGetData = New ClsResultSetDB
            rsGetData.GetResult(strSql)
            If rsGetData.RowCount > 0 Then
                txtRefNo.Text = ""
                Do While Not rsGetData.EOFRecord
                    If Len(Trim(txtRefNo.Text)) = 0 Then
                        txtRefNo.Text = rsGetData.GetValue("Ref_Doc_No")
                    Else
                        txtRefNo.Text = Trim(txtRefNo.Text) & "," & rsGetData.GetValue("Ref_Doc_No")
                    End If
                    rsGetData.MoveNext()
                Loop
            End If
            rsGetData.ResultSetClose()
        End If
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
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
        Dim intDecimal As Short
        Dim strSql As String = "", strModel As String = ""
        Dim ATNNo As String
        Dim objSQLConn As SqlConnection
        Dim objReader As SqlDataReader
        Dim objCommand As SqlCommand


        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                strsaledtl = ""
                strsaledtl = "SELECT EOP_MODEL,Location_Code,Doc_No,Suffix,Item_Code,Sales_Quantity,From_Box,To_Box,Rate,Sales_Tax,Excise_Tax,Packing,Others,Cust_Mtrl,Year,Cust_Item_Code,Cust_Item_Desc,Tool_Cost,Measure_Code,Excise_type,SalesTax_type,CVD_type,SAD_type,GL_code,SL_code,Basic_Amount,Accessible_amount,CVD_Amount,SVD_amount,Excise_per,CVD_per,SVD_per,CustMtrl_Amount,ToolCost_amount,pervalue,TotalExciseAmount,SupplementaryInvoiceFlag,To_Location,Discount_type,Discount_amt,Discount_perc,From_Location,Cust_ref,Amendment_No,SRVDINO,SRVLocation,USLOC,SchTime,BinQuantity,Packing_Type,ItemPacking_Amount,Item_remark,pkg_amount,csiexcise_amount,ADD_EXCISE_TYPE,ADD_EXCISE_PER,ADD_EXCISE_AMOUNT,UNIT_CODE,HSNSACCODE,CGSTTXRT_TYPE,CGST_PERCENT,CGST_AMT,SGSTTXRT_TYPE,SGST_PERCENT,SGST_AMT,UTGSTTXRT_TYPE,UTGST_PERCENT,UTGST_AMT,IGSTTXRT_TYPE,IGST_PERCENT,IGST_AMT,COMPENSATION_CESS_TYPE,COMPENSATION_CESS_PERCENT,COMPENSATION_CESS_AMT  from Sales_Dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  Location_Code='" & Trim(txtLocationCode.Text) & "'"
                strsaledtl = strsaledtl & " and Doc_No=" & Val(txtChallanNo.Text) & " and Cust_Item_Code in(" & Trim(mstrItemCode) & ")"
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                '10869291
                If UCase(Trim(CmbInvType.Text)) = "NORMAL INVOICE" Or UCase(Trim(CmbInvType.Text)) = "JOBWORK INVOICE" Or UCase(Trim(CmbInvType.Text)) = "EXPORT INVOICE" Or UCase(Trim(CmbInvType.Text)) = "TRANSFER INVOICE" Or UCase(Trim(CmbInvType.Text)) = "SERVICE INVOICE" Or UCase(Trim(CmbInvType.Text)) = "INTER-DIVISION" Then
                    If UCase(Trim(CmbInvSubType.Text)) <> "SCRAP" Then
                        Dim strmultipleSO As String
                        strsaledtl = ""
                        strmultipleSO = Replace(CUSTREFLIST, "'", "")
                        strsaledtl = Replace(mstrItemCode, "'", "")
                        strsaledtl = "SELECT * FROM DBO.UDF_GET_SO_ITEM_DTL ('" + gstrUNITID + "','" & txtCustCode.Text & "','" & strmultipleSO & "','" & strsaledtl & "','" & txtAmendNo.Text.Trim & "' )"
                    Else
                        strsaledtl = ""
                        strsaledtl = "SELECT Item_Code,standard_Rate from Item_Mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  "
                        strsaledtl = strsaledtl & " Status = 'A' and Hold_flag <> 1 and Item_Code in (" & mstrItemCode & ")"
                    End If
                Else
                    If UCase(Trim(CmbInvType.Text)) = "REJECTION" Then
                        If mblnRejTracking = True Then
                            If Len(Trim(txtRefNo.Text)) > 0 Then
                                If chkRejType.Text <> "LRN" Then
                                    strsaledtl = ""
                                    strsaledtl = "SELECT Distinct Item_Code from grn_Dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  "
                                    strsaledtl = strsaledtl & " Item_Code in (" & mstrItemCode & ") and Doc_No In (" & txtRefNo.Text & ")"
                                Else
                                    strsaledtl = ""
                                    strsaledtl = "SELECT Item_Code,standard_Rate from Item_Mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  "
                                    strsaledtl = strsaledtl & " Status = 'A' and Hold_flag <> 1 and Item_Code in (" & mstrItemCode & ")"
                                End If
                            Else
                                If chkRejType.Text <> "LRN" Then
                                    strsaledtl = ""
                                    strsaledtl = "SELECT Item_Code,standard_Rate from Item_Mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  "
                                    strsaledtl = strsaledtl & " Status = 'A' and Hold_flag <> 1 and Item_Code in (" & mstrItemCode & ")"
                                Else
                                    strsaledtl = ""
                                    strsaledtl = "SELECT Item_Code,standard_Rate from Item_Mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  "
                                    strsaledtl = strsaledtl & " Status = 'A' and Hold_flag <> 1 and Item_Code in (" & mstrItemCode & ")"
                                End If
                            End If
                        Else
                            If Len(Trim(txtRefNo.Text)) > 0 Then
                                strsaledtl = ""
                                strsaledtl = "SELECT Item_Code,standard_Rate = Item_Rate from grn_Dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  "
                                strsaledtl = strsaledtl & " Item_Code in (" & mstrItemCode & ") and Doc_No =" & txtRefNo.Text
                            Else
                                strsaledtl = ""
                                strsaledtl = "SELECT Item_Code,standard_Rate from Item_Mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  "
                                strsaledtl = strsaledtl & " Status = 'A' and Hold_flag <> 1 and Item_Code in (" & mstrItemCode & ")"
                            End If
                        End If
                        '10869291
                    ElseIf (UCase(Trim(CmbInvType.Text)) = "TRANSFER INVOICE" And UCase(Trim(CmbInvSubType.Text)) = "FINISHED GOODS") Or (UCase(Trim(CmbInvType.Text)) = "INTER-DIVISION" And UCase(Trim(CmbInvSubType.Text)) = "FINISHED GOODS") Then
                        strsaledtl = ""
                        strsaledtl = "SELECT Distinct a.Item_Code,c.Cust_drgNo,a.standard_Rate FROM Item_Mst a,Itembal_Mst b,CustItem_Mst c "
                        strsaledtl = strsaledtl & " WHERE A.UNIT_CODE=B.UNIT_CODE AND B.UNIT_CODE=C.UNIT_CODE AND A.UNIT_CODE='" + gstrUNITID + "' AND  a.Item_Code=b.Item_Code and a.Item_Code = c.ITem_Code"
                        strsaledtl = strsaledtl & " and a.Status ='A' and a.Hold_Flag <> 1 and c.Account_code ='" & txtCustCode.Text & "'"
                        strsaledtl = strsaledtl & " and a.Item_code in (" & mstrItemCode & ")"
                    Else
                        strsaledtl = ""
                        strsaledtl = "SELECT Item_Code,standard_Rate from Item_Mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  "
                        strsaledtl = strsaledtl & " Status = 'A' and Hold_flag <> 1 and Item_Code in (" & mstrItemCode & ")"
                    End If
                End If
        End Select
        rsSalesDtl = New ClsResultSetDB
        rsSalesDtl.GetResult(strsaledtl, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        Dim intLoopcount As Short
        Dim varCumulative As Object
        Dim strCompileString As String
        Dim dblqty As Double
        Dim dblStock As Double
        Dim strpono As String
        Dim intMaxSerial_No As Short
        Dim inttotalrows As Integer

        With SpChEntry
            If SpChEntry.MaxRows <= 0 Then
                inttotalrows = 0
            Else
                inttotalrows = SpChEntry.MaxRows
            End If
        End With


        If rsSalesDtl.GetNoRows > 0 Then
            intRecordCount = rsSalesDtl.GetNoRows

            ReDim mdblPrevQty(intRecordCount - 1) ' To get value of Quantity in Arrey for updation in despatch
            ReDim Preserve mdblToolCost(intRecordCount - 1 + inttotalrows) ' To get value of Quantity i

            If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                If SpChEntry.MaxRows > 0 Then
                    varItemAlready = Nothing
                    Call SpChEntry.GetText(1, 1, varItemAlready)
                    If Len(Trim(varItemAlready)) = 0 Then
                        Call addRowAtEnterKeyPress(intRecordCount - 1)
                    End If
                Else
                    Call addRowAtEnterKeyPress(intRecordCount)
                End If
            Else
                Call addRowAtEnterKeyPress(intRecordCount - 1)
            End If
            rsSalesDtl.MoveFirst()
            '10869290
            If CmbInvType.Text = "NORMAL INVOICE" Or CmbInvType.Text = "JOBWORK INVOICE" Or CmbInvType.Text = "EXPORT INVOICE" Or CmbInvType.Text = "TRANSFER INVOICE" Or CmbInvType.Text = "SERVICE INVOICE" Or CmbInvType.Text = "INTER-DIVISION" Then
                If UCase(Trim(CmbInvSubType.Text)) <> "SCRAP" Then
                    For intLoopcount = 1 + inttotalrows To intRecordCount + inttotalrows
                        If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                            mdblToolCost(intLoopcount - 1) = Val(rsSalesDtl.GetValue("Tool_Cost"))
                        Else
                            mdblToolCost(intLoopcount - 1) = Val(rsSalesDtl.GetValue("Tool_Cost"))
                        End If
                        rsSalesDtl.MoveNext()
                    Next
                End If
            End If
            rsSalesDtl.MoveFirst()
            intDecimal = ToGetDecimalPlaces(pstrCurrency)
            Call SetMaxLengthInSpread(intDecimal)
            If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                If SpChEntry.MaxRows > 0 Then
                    varItemAlready = Nothing
                    Call SpChEntry.GetText(1, 1, varItemAlready)
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
            For intLoopCounter = inti To intRecordCount
                With Me.SpChEntry
                    Select Case Me.CmdGrpChEnt.Mode
                        Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                            .Row = 1 : .Row2 = .MaxRows : .Col = 0 : .Col2 = .MaxCols
                            .Enabled = True : .BlockMode = True : .Lock = True : .BlockMode = False
                            Call .SetText(1, intLoopCounter, rsSalesDtl.GetValue("Item_Code"))
                            Call .SetText(2, intLoopCounter, rsSalesDtl.GetValue("Cust_Item_Code"))
                            Call .SetText(3, intLoopCounter, rsSalesDtl.GetValue("Rate") * Val(ctlPerValue.Text))
                            Call .SetText(17, intLoopCounter, rsSalesDtl.GetValue("Rate"))
                            Call .SetText(4, intLoopCounter, rsSalesDtl.GetValue("Cust_Mtrl") * Val(ctlPerValue.Text))
                            Call .SetText(18, intLoopCounter, rsSalesDtl.GetValue("Cust_Mtrl"))
                            Call .SetText(5, intLoopCounter, rsSalesDtl.GetValue("Sales_Quantity"))
                            Call .GetText(5, intLoopCounter, mdblPrevQty(intLoopCounter - 1))

                            If mblnRejTracking = True Then
                                Call AddMaxAllowedQuanity(intLoopCounter)
                            End If
                            Call .SetText(6, intLoopCounter, rsSalesDtl.GetValue("Packing"))
                            Call .SetText(7, intLoopCounter, rsSalesDtl.GetValue("Excise_Type"))
                            Call .SetText(8, intLoopCounter, rsSalesDtl.GetValue("ADD_Excise_Type"))
                            Call .SetText(9, intLoopCounter, rsSalesDtl.GetValue("CVD_type"))
                            Call .SetText(10, intLoopCounter, rsSalesDtl.GetValue("SAD_type"))
                            Call .SetText(11, intLoopCounter, rsSalesDtl.GetValue("Others") * Val(ctlPerValue.Text))
                            Call .SetText(19, intLoopCounter, rsSalesDtl.GetValue("Others"))
                            Call .SetText(12, intLoopCounter, rsSalesDtl.GetValue("From_Box"))
                            Call .SetText(13, intLoopCounter, rsSalesDtl.GetValue("To_Box"))
                            Call .SetText(16, intLoopCounter, rsSalesDtl.GetValue("tool_Cost") * Val(ctlPerValue.Text))
                            Call .SetText(20, intLoopCounter, rsSalesDtl.GetValue("tool_Cost"))

                            If mblnRejTracking = True And mblnBatchTracking = True Then
                                strCompileString = GetDocumentDetail(rsSalesDtl.GetValue("Item_Code"))
                                Call .SetText(23, intLoopCounter, strCompileString)
                            End If
                            ''10808160
                            'strSql = "SELECT Top 1 MODEL_CODE FROM BUDGETITEM_MST WHERE UNIT_CODE= '" & gstrUNITID & "' AND ACCOUNT_CODE = '" & txtCustCode.Text.Trim & "'AND ITEM_CODE = '" & rsSalesDtl.GetValue("Item_Code") & "' AND CUST_DRGNO = '" & rsSalesDtl.GetValue("Cust_DrgNo") & "' AND ENDDATE < '" & GetServerDateNew().ToString("dd MMM yyyy") & "'"
                            'strModel = Convert.ToString(SqlConnectionclass.ExecuteScalar(strSql))
                            'Call .SetText(26, intLoopCounter, strModel)
                            ''10808160
                            '10808160
                            Call .SetText(27, intLoopCounter, rsSalesDtl.GetValue("EOP_MODEL"))
                            '10808160
                            'GST CHANGES
                            If gblnGSTUnit = True Then
                                Call .SetText(30, intLoopCounter, rsSalesDtl.GetValue("HSNSACCODE"))
                                Call .SetText(31, intLoopCounter, rsSalesDtl.GetValue("CGSTTXRT_TYPE"))
                                Call .SetText(32, intLoopCounter, rsSalesDtl.GetValue("SGSTTXRT_TYPE"))
                                Call .SetText(33, intLoopCounter, rsSalesDtl.GetValue("UTGSTTXRT_TYPE"))
                                Call .SetText(34, intLoopCounter, rsSalesDtl.GetValue("IGSTTXRT_TYPE"))
                                Call .SetText(35, intLoopCounter, rsSalesDtl.GetValue("COMPENSATION_CESS_TYPE"))
                            End If

                            'GST CHANGES
                            If intLoopCounter = 1 Then
                                Call .SetText(14, intLoopCounter, (rsSalesDtl.GetValue("To_Box") - rsSalesDtl.GetValue("From_Box")) + 1)
                            Else
                                varCumulative = Nothing
                                Call .GetText(14, intLoopCounter - 1, varCumulative)
                                Call .SetText(14, intLoopCounter, varCumulative + ((rsSalesDtl.GetValue("To_Box") - rsSalesDtl.GetValue("From_Box")) + 1))
                            End If
                            Dim varCustItemCode As Object
                            Dim rsBatch = New ClsResultSetDB
                            Dim Intcounter As Integer
                            If mblnBatchTrack = True And UCase(CmbInvType.Text) <> "JOBWORK INVOICE" And UCase(CmbInvType.Text) <> "REJECTION INVOICE" Then
                                varItem_Code = Nothing
                                Call .GetText(1, intLoopCounter, varItem_Code)
                                varCustItemCode = Nothing
                                Call .GetText(2, intLoopCounter, varCustItemCode)
                                rsBatch = New ClsResultSetDB
                                Call rsBatch.GetResult("Select Batch_No,Batch_Date,Batch_Qty from ItemBatch_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  Item_Code = '" & varItem_Code & "' and Cust_Item_Code = '" & varCustItemCode & "' and  From_Location = '" & mstrLocationCode & "' and Doc_No = " & Trim(Me.txtChallanNo.Text) & " and Doc_type = 9999 ", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                                If rsBatch.RowCount > 0 Then
                                    ReDim Preserve mBatchData(intLoopCounter)
                                    ReDim Preserve mBatchData(intLoopCounter).Batch_No(rsBatch.RowCount)
                                    ReDim Preserve mBatchData(intLoopCounter).Batch_Date(rsBatch.RowCount)
                                    ReDim Preserve mBatchData(intLoopCounter).Batch_Quantity(rsBatch.RowCount)
                                    Intcounter = 1
                                    While Not rsBatch.EOFRecord
                                        mBatchData(intLoopCounter).Batch_No(Intcounter) = rsBatch.GetValue("Batch_No")
                                        mBatchData(intLoopCounter).Batch_Date(Intcounter) = ConvertToDate(VB6.Format(rsBatch.GetValue("Batch_Date"), gstrDateFormat))
                                        mBatchData(intLoopCounter).Batch_Quantity(Intcounter) = rsBatch.GetValue("Batch_Qty")
                                        Intcounter = Intcounter + 1
                                        rsBatch.MoveNext()
                                    End While
                                End If
                                rsBatch.ResultSetClose()
                            End If
                            If mblnRejTracking = True And mblnBatchTracking = True Then
                                If UCase(strInvType) = "REJ" Then
                                    .Col = 24 : .ColHidden = False
                                End If
                            End If
                        Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                            .Enabled = True
                            .Row = 1 : .Row2 = .MaxRows : .Col = 0 : .Col2 = .MaxCols : .BlockMode = True : .Lock = False : .BlockMode = False
                            If CmbInvType.Text = "NORMAL INVOICE" And CmbInvSubType.Text = "FINISHED GOODS" And DataExist("SELECT PCA_CUSTOMER FROM CUSTOMER_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & Trim(txtCustCode.Text) & "' and PCA_CUSTOMER =1 ") = True Then
                                .Row = 1 : .Row2 = .MaxRows : .Col = 5 : .Col2 = 5 : .BlockMode = True : .Lock = True : .BlockMode = False
                            End If
                            ''10869290
                            If (Trim(CmbInvType.Text) = "NORMAL INVOICE") Or (Trim(CmbInvType.Text) = "JOBWORK INVOICE") Or (Trim(CmbInvType.Text) = "EXPORT INVOICE") Or (Trim(CmbInvType.Text) = "TRANSFER INVOICE") Or (Trim(CmbInvType.Text) = "INTER-DIVISION") Or (Trim(CmbInvType.Text) = "SERVICE INVOICE") Then
                                If CBool(UCase(CStr((Trim(CmbInvSubType.Text)) <> "SCRAP"))) Then
                                    Call .SetText(1, intLoopCounter, rsSalesDtl.GetValue("Item_Code"))
                                    Call .SetText(2, intLoopCounter, rsSalesDtl.GetValue("Cust_DrgNo"))
                                    Call .SetText(3, intLoopCounter, (Val(rsSalesDtl.GetValue("Rate")) * Val(ctlPerValue.Text)))
                                    Call .SetText(17, intLoopCounter, Val(rsSalesDtl.GetValue("Rate")))
                                    Call .SetText(4, intLoopCounter, (Val(rsSalesDtl.GetValue("Cust_Mtrl")) * Val(ctlPerValue.Text)))
                                    Call .SetText(18, intLoopCounter, Val(rsSalesDtl.GetValue("Cust_Mtrl")))
                                    Call .SetText(6, intLoopCounter, rsSalesDtl.GetValue("Packing"))
                                    Call .SetText(7, intLoopCounter, rsSalesDtl.GetValue("Excise_duty"))
                                    Call .SetText(8, intLoopCounter, rsSalesDtl.GetValue("ADD_Excise_duty"))
                                    Call .SetText(11, intLoopCounter, (Val(rsSalesDtl.GetValue("Others")) * Val(ctlPerValue.Text)))
                                    Call .SetText(19, intLoopCounter, Val(rsSalesDtl.GetValue("Others")))
                                    Call .SetText(16, intLoopCounter, (Val(rsSalesDtl.GetValue("tool_cost")) * Val(ctlPerValue.Text)))
                                    Call .SetText(20, intLoopCounter, Val(rsSalesDtl.GetValue("tool_cost")))
                                    Call .SetText(26, intLoopCounter, rsSalesDtl.GetValue("external_salesorder_no"))
                                    'GST CHANGES
                                    If CmbInvType.Text = "NORMAL INVOICE" And CmbInvSubType.Text = "FINISHED GOODS" And DataExist("SELECT PCA_CUSTOMER FROM CUSTOMER_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & Trim(txtCustCode.Text) & "' and PCA_CUSTOMER =1 ") = True Then

                                        dblqty = CDbl(Find_Value("Select SUM(QUANTITY) QUANTITY from TMP_PCA_ITEMSELECTION_PACKAGECODE WHERE UNIT_CODE='" + gstrUNITID + "' AND  IP_ADDRESS='" & gstrIpaddressWinSck & "' and Item_code='" & rsSalesDtl.GetValue("Item_code") & "'"))
                                        Call .SetText(5, intLoopCounter, dblqty)
                                    End If

                                    If gblnGSTUnit = True Then
                                        Call .SetText(30, intLoopCounter, rsSalesDtl.GetValue("HSNSACCODE"))
                                        Call .SetText(31, intLoopCounter, rsSalesDtl.GetValue("CGSTTXRT_TYPE"))
                                        Call .SetText(32, intLoopCounter, rsSalesDtl.GetValue("SGSTTXRT_TYPE"))
                                        Call .SetText(33, intLoopCounter, rsSalesDtl.GetValue("UTGSTTXRT_TYPE"))
                                        Call .SetText(34, intLoopCounter, rsSalesDtl.GetValue("IGSTTXRT_TYPE"))
                                        Call .SetText(35, intLoopCounter, rsSalesDtl.GetValue("COMPENSATION_CESS"))
                                    End If

                                    'GST CHANGES

                                    '10808160
                                    strSql = "SELECT Top 1 MODEL_CODE FROM BUDGETITEM_MST WHERE UNIT_CODE= '" & gstrUNITID & "' AND ACCOUNT_CODE = '" & txtCustCode.Text.Trim & "'AND ITEM_CODE = '" & rsSalesDtl.GetValue("Item_Code") & "' AND CUST_DRGNO = '" & rsSalesDtl.GetValue("Cust_DrgNo") & "' AND ENDDATE >= '" & GetServerDateNew().ToString("dd MMM yyyy") & "' AND DefaultModel=1 "
                                    strModel = Convert.ToString(SqlConnectionclass.ExecuteScalar(strSql))
                                    Call .SetText(27, intLoopCounter, strModel)
                                    '10808160
                                    If mblnATN_invoicewise = True Then
                                        ATNNo = Find_Value("select ATNCODE from VW_FA_ATN WHERE companycode='" + mstrcurrentATNCode + "' AND  transferedTo='" & mstrcustcompcode & "' and itemname='" & rsSalesDtl.GetValue("Item_Code") & "'" &
                                                           " AND NOT EXISTS ( SELECT TOP  1 1 FROM SALES_DTL WHERE UNIT_CODE =VW_FA_ATN.UNIT_CODE AND ATNNO = VW_FA_ATN.ATNCODE )")
                                        Call .SetText(28, intLoopCounter, ATNNo)
                                    End If
                                    'atn changes
                                Else
                                    Call .SetText(1, intLoopCounter, rsSalesDtl.GetValue("Item_Code"))
                                    Call .SetText(2, intLoopCounter, rsSalesDtl.GetValue("Item_code"))
                                    Call .SetText(3, intLoopCounter, (Val(rsSalesDtl.GetValue("Standard_Rate")) * Val(ctlPerValue.Text)))
                                    Call .SetText(17, intLoopCounter, Val(rsSalesDtl.GetValue("Standard_Rate")))
                                    '10808160
                                    If CBool(UCase(CStr((Trim(CmbInvSubType.Text)) <> "SCRAP"))) Then
                                        strSql = "SELECT Top 1 MODEL_CODE FROM BUDGETITEM_MST WHERE UNIT_CODE= '" & gstrUNITID & "' AND ACCOUNT_CODE = '" & txtCustCode.Text.Trim & "'AND ITEM_CODE = '" & rsSalesDtl.GetValue("Item_Code") & "' AND CUST_DRGNO = '" & rsSalesDtl.GetValue("Cust_DrgNo") & "' AND ENDDATE >= '" & GetServerDateNew().ToString("dd MMM yyyy") & "' AND DefaultModel=1 "
                                        strModel = Convert.ToString(SqlConnectionclass.ExecuteScalar(strSql))
                                        Call .SetText(27, intLoopCounter, strModel)
                                    End If
                                    '10808160

                                End If
                            Else
                                If CmbInvType.Text = "REJECTION" And mblnRejTracking = True Then
                                    ' In case of Rejection Invoice Entrt Rates/ Taxes Will be Picked from
                                    ' PO Order
                                    If Len(Trim(txtRefNo.Text)) <> 0 Then
                                        Call .SetText(1, intLoopCounter, rsSalesDtl.GetValue("Item_Code"))
                                        Call .SetText(2, intLoopCounter, rsSalesDtl.GetValue("Item_code"))
                                        ' Total Possible Quantity
                                        If chkRejType.Text <> "LRN" Then
                                            strsaledtl = "select MaxAllowedQty = SUM( ((a.Rejected_Quantity + a.excess_po_quantity) - (isnull(a.Despatch_Quantity,0) + isnull(a.Inspected_Quantity,0) + isnull(a.RGP_Quantity,0)))) from grn_Dtl a, grn_hdr b Where A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" + gstrUNITID + "' AND "
                                            strsaledtl = strsaledtl & " a.Doc_type = b.Doc_type And a.Doc_No = b.Doc_No and "
                                            strsaledtl = strsaledtl & " a.From_Location = b.From_Location and a.From_Location ='01R1'"
                                            strsaledtl = strsaledtl & " and a.Rejected_quantity > 0 and b.Vendor_code = '" & txtCustCode.Text
                                            strsaledtl = strsaledtl & "' and a.Doc_No in (" & Trim(txtRefNo.Text) & ") and a.Item_code = '" & rsSalesDtl.GetValue("Item_Code") & "' AND ISNULL(b.GRN_Cancelled,0) = 0"
                                            strsaledtl = strsaledtl & " Group by a.Item_code "
                                            dblqty = CDbl(Find_Value(strsaledtl))
                                        Else
                                            strsaledtl = "Select  (isnull(Sum(rejected_Quantity),0)-isnull(Sum(Quantity),0)) as MaxAllowedQty from LRN_HDR as a " &
                                                                       " Inner Join LRN_DTL as b on a.doc_No=b.doc_no AND A.UNIT_CODE=B.UNIT_CODE and a.Doc_Type=b.doc_Type and a.from_Location=b.from_location " &
                                                                       " Left Outer Join(Select Ref_doc_no,item_code,isnull(Sum(Quantity),0) as Quantity, UNIT_CODE" &
                                                                       " from MKT_INVREJ_DTL WHERE UNIT_CODE='" + gstrUNITID + "'  and Cancel_flag <> 1 and Rej_type = 2 Group by Ref_doc_no,item_code,UNIT_CODE Having IsNull(Sum(Quantity), 0) > 0)g " &
                                                                       " ON a.doc_No=g.Ref_doc_No and a.UNIT_CODE=g.UNIT_CODE and b.item_code=g.item_code and b.UNIT_CODE=g.UNIT_CODE" &
                                                                       " WHERE a.UNIT_CODE='" + gstrUNITID + "' AND  Authorized_Code Is Not Null " &
                                                                       " and a.Doc_No IN (" & Trim(txtRefNo.Text) & ") and B.ITem_code = '" & rsSalesDtl.GetValue("Item_code") & "' Group by B.Item_Code "
                                            dblqty = CDbl(Find_Value(strsaledtl))
                                        End If
                                        dblStock = CDbl(Find_Value("Select Cur_Bal From ItemBal_Mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  Location_code='01J1' and Item_code='" & rsSalesDtl.GetValue("Item_code") & "'"))
                                        If dblStock < dblqty Then
                                            Call .SetText(5, intLoopCounter, dblStock)
                                            Call .SetText(22, intLoopCounter, dblStock)
                                        Else
                                            Call .SetText(5, intLoopCounter, dblqty)
                                            Call .SetText(22, intLoopCounter, dblqty)
                                        End If
                                        '''SpChEntry_Change 5, intLoopCounter
                                        SpChEntry_Change(SpChEntry, New AxFPSpreadADO._DSpreadEvents_ChangeEvent(5, intLoopCounter))
                                        If chkRejType.Text = "LRN" Then
                                            If CBool(Find_Value("Select FormLevelBatch_Tracking from FORM_LEVEL_FLAGS WHERE UNIT_CODE='" + gstrUNITID + "' AND  Form_Name='FRMSTRTRN0027'")) = True Then
                                                'strpono = Find_Value("Select Top 1 PUR_ORDER_NO From vw_INVREJ_LRN_DETAIL WHERE UNIT_CODE='" + gstrUNITID + "' AND  Doc_No in (" & Trim(txtRefNo.Text) & ")")
                                                strpono = Find_Value("Select Top 1 PUR_ORDER_NO From vw_INVREJ_LRN_DETAIL WHERE UNIT_CODE='" + gstrUNITID + "' AND  Doc_No in (" & Trim(txtRefNo.Text) & ")" & " and item_code='" & rsSalesDtl.GetValue("Item_code") & "'")
                                            Else
                                                strpono = Find_Value("Select Top 1 PUR_ORDER_NO From vw_INVREJ_LRN_DETAIL_WB WHERE UNIT_CODE='" + gstrUNITID + "' AND  Doc_No in (" & Trim(txtRefNo.Text) & ")")
                                            End If
                                        Else
                                            strpono = Find_Value("Select Top 1 PUR_Order_No from GRN_HDR WHERE UNIT_CODE='" + gstrUNITID + "' AND  Doc_No in (" & Trim(txtRefNo.Text) & ")")
                                        End If
                                        If Len(Trim(strpono)) = 0 Then
                                            strpono = CStr(0)
                                        End If
                                        intMaxSerial_No = CShort(Val(Find_Value("Select Max(Serial_No) From PoOrds_Hdr WHERE UNIT_CODE='" + gstrUNITID + "' AND  Pur_Order_No='" & strpono & "'")))
                                        If chkRejType.Text = "LRN" Then
                                            If CBool(Find_Value("Select FormLevelBatch_Tracking from FORM_LEVEL_FLAGS WHERE UNIT_CODE='" + gstrUNITID + "' AND  Form_Name='FRMSTRTRN0027'")) = True Then
                                                strsaledtl = " Select  top 1 c.Rate, Excise_TaxID, Sale_TaxID, Tool_cost, Ecess_Percent, TxRt_Rate_No, a.Doc_no, GRN_DATE from GRN_HDR as a  Inner join PoOrds_Hdr as b on a.Pur_Order_No=b.Pur_Order_No and A.UNIT_CODE=B.UNIT_CODE  Inner Join Vend_Item as c On B.Pur_Order_No=C.Pur_Order_No AND B.UNIT_CODE=C.UNIT_CODE  and B.Account_code=c.Account_code and B.Serial_No=C.Serial_No " & " where A.UNIT_CODE='" + gstrUNITID + "' AND a.Pur_Order_No=" & strpono & " and Item_code ='" & rsSalesDtl.GetValue("Item_Code") & "' " & " and a.Doc_No in " & " (Select GRIN_No from vw_INVREJ_LRN_DETAIL WHERE UNIT_CODE='" + gstrUNITID + "' AND  Doc_No in (" & Trim(txtRefNo.Text) & ")) " & " and  B.Serial_No <> 0 "
                                            Else
                                                strsaledtl = " Select  top 1 c.Rate, Excise_TaxID,Sale_TaxID, Tool_cost, Ecess_Percent, TxRt_Rate_No from GRN_HDR as a Inner join PoOrds_Hdr as b on a.Pur_Order_No=b.Pur_Order_No AND A.UNIT_CODE=B.UNIT_CODE Inner Join Vend_Item as c On B.Pur_Order_No=C.Pur_Order_No and B.Account_code=c.Account_code AND B.UNIT_CODE=C.UNIT_CODE and B.Serial_No=C.Serial_No where A.UNIT_CODE='" + gstrUNITID + "' AND  Item_code ='" & rsSalesDtl.GetValue("Item_Code") & "' and B.Serial_No <> 0 "
                                            End If
                                            ' IF PO HAVE AMENDMENTS THEN
                                            If intMaxSerial_No > 1 Then
                                                strsaledtl = strsaledtl & " and DateDiff(n, Isnull(Amendment_date, PO_Date), GRN_DATE ) >= 0 "
                                            End If
                                            If intMaxSerial_No > 1 Then
                                                strsaledtl = strsaledtl & " Order by a.GRN_DATE, DateDiff(n, Isnull(Amendment_date, PO_Date), GRN_DATE ) "
                                            Else
                                                strsaledtl = strsaledtl & " Order by a.GRN_DATE DESC"
                                            End If
                                        Else
                                            strsaledtl = " Select  top 1 c.Rate as Rate, Excise_TaxID, Sale_TaxID, Tool_cost, Ecess_Percent, TxRt_Rate_No, a.Doc_no, GRN_DATE from GRN_HDR as a  Inner join PoOrds_Hdr as b on a.Pur_Order_No=b.Pur_Order_No AND A.UNIT_CODE=B.UNIT_CODE Inner Join Vend_Item as c On B.Pur_Order_No=c.Pur_Order_No AND B.UNIT_CODE=C.UNIT_CODE and B.Account_code=c.Account_code and B.Serial_No=C.Serial_No " & " where A.UNIT_CODE='" + gstrUNITID + "' AND a.Pur_Order_No=" & strpono & " and Item_code ='" & rsSalesDtl.GetValue("Item_Code") & "' " & " and a.Doc_No in " & " (Select Doc_No from GRN_HDR WHERE UNIT_CODE='" + gstrUNITID + "' AND  Doc_No in (" & Trim(txtRefNo.Text) & ")) " & " and  B.Serial_No <> 0 "
                                            ' IF PO HAVE AMENDMENTS THEN
                                            If intMaxSerial_No > 1 Then
                                                strsaledtl = strsaledtl & " and DateDiff(n, Isnull(Amendment_date, PO_Date), GRN_DATE ) >= 0 "
                                            End If
                                            If intMaxSerial_No > 1 Then
                                                strsaledtl = strsaledtl & " Order by a.GRN_DATE, DateDiff(n, Isnull(Amendment_date, PO_Date), GRN_DATE ) "
                                            Else
                                                strsaledtl = strsaledtl & " Order by a.GRN_DATE"
                                            End If
                                        End If
                                        Dim rsTmp As New ClsResultSetDB
                                        rsTmp.GetResult(strsaledtl)
                                        If rsTmp.GetNoRows > 0 Then
                                            Call .SetText(3, intLoopCounter, rsTmp.GetValue("Rate") * CDbl(ctlPerValue.Text))
                                            Call .SetText(16, intLoopCounter, rsTmp.GetValue("Tool_cost") * CDbl(ctlPerValue.Text))
                                            Call .SetText(16, intLoopCounter, rsTmp.GetValue("Tool_cost"))
                                            Call .SetText(3, intLoopCounter, rsTmp.GetValue("Rate") * CDbl(ctlPerValue.Text))
                                            If gblnGSTUnit = False Then
                                                Call .SetText(7, intLoopCounter, rsTmp.GetValue("Excise_TaxID"))
                                            End If

                                            If intLoopCounter > 1 Then
                                                'If Tax Set for first Item is not Equal to Current Item
                                                ' Give alert
                                                If Trim(txtSaleTaxType.Text) <> Trim(rsTmp.GetValue("Sale_TaxId")) Or txtECSSTaxType.Text <> rsTmp.GetValue("TxRT_Rate_No") Then
                                                    MsgBox("Taxes defined for Item Code " & rsSalesDtl.GetValue("Item_Code") & " does not match with other Items selected in the Invoice." & vbCrLf & "Please check the Tax Details.", MsgBoxStyle.Information, ResolveResString(100))
                                                    txtSaleTaxType.Text = rsTmp.GetValue("Sale_TaxId")
                                                    txtSaleTaxType_Validating(txtSaleTaxType, New System.ComponentModel.CancelEventArgs(False))
                                                    txtECSSTaxType.Text = rsTmp.GetValue("TxRT_Rate_No")
                                                    txtECSSTaxType_Validating(txtECSSTaxType, New System.ComponentModel.CancelEventArgs(False))
                                                End If
                                            Else
                                                If gblnGSTUnit = False Then
                                                    txtSaleTaxType.Text = rsTmp.GetValue("Sale_TaxId")
                                                    txtSaleTaxType_Validating(txtSaleTaxType, New System.ComponentModel.CancelEventArgs(False))
                                                    txtECSSTaxType.Text = rsTmp.GetValue("TxRT_Rate_No")
                                                    txtECSSTaxType_Validating(txtECSSTaxType, New System.ComponentModel.CancelEventArgs(False))
                                                End If
                                            End If
                                        End If
                                        rsTmp.ResultSetClose()
                                    Else
                                        Call .SetText(1, intLoopCounter, rsSalesDtl.GetValue("Item_Code"))
                                        Call .SetText(2, intLoopCounter, rsSalesDtl.GetValue("Item_code"))
                                        Call .SetText(3, intLoopCounter, (rsSalesDtl.GetValue("Standard_Rate") * Val(ctlPerValue.Text)))
                                        Call .SetText(3, intLoopCounter, rsSalesDtl.GetValue("Standard_Rate"))
                                    End If
                                Else
                                    Call .SetText(1, intLoopCounter, rsSalesDtl.GetValue("Item_Code"))
                                    If (UCase(CmbInvType.Text) = "TRANSFER INVOICE" And UCase(CmbInvSubType.Text) = "FINISHED GOODS") Or (UCase(CmbInvType.Text) = "INTER-DIVISION" And UCase(CmbInvSubType.Text) = "FINISHED GOODS") Then
                                        Call .SetText(2, intLoopCounter, rsSalesDtl.GetValue("cust_DrgNo"))
                                    Else
                                        Call .SetText(2, intLoopCounter, rsSalesDtl.GetValue("Item_code"))
                                    End If
                                    Call .SetText(3, intLoopCounter, (rsSalesDtl.GetValue("Standard_Rate") * Val(ctlPerValue.Text)))
                                    Call .SetText(17, intLoopCounter, rsSalesDtl.GetValue("Standard_Rate"))
                                End If
                            End If
                            'GST CHANGES    
                            Dim VARITEMCODE As Object
                            VARITEMCODE = Nothing
                            Call .GetText(1, intLoopCounter, VARITEMCODE)
                            If CmbInvType.Text = "REJECTION" Then
                                strSql = "set dateformat 'dmy' SELECT * FROM DBO.UFN_GST_TAXES_REJECTIONINVOICE_DETAILS('" & gstrUNITID & "','" & txtCustCode.Text.Trim & "','" & VARITEMCODE & "','" & GetServerDate() & "')"
                            Else
                                If txtRefNo.Enabled = True Then
                                    strSql = "set dateformat 'dmy' SELECT * FROM DBO.UFN_GST_ITEMWISETAXES('" & gstrUNITID & "','" & txtCustCode.Text.Trim & "','" & VARITEMCODE & "','','')"
                                Else
                                    strSql = "set dateformat 'dmy' SELECT * FROM DBO.UFN_GST_ITEMWISETAXES('" & gstrUNITID & "','" & txtCustCode.Text.Trim & "','" & VARITEMCODE & "','','')"
                                End If

                            End If

                            objSQLConn = SqlConnectionclass.GetConnection()
                            objCommand = New SqlCommand(strSql, objSQLConn)
                            objReader = objCommand.ExecuteReader()

                            If objReader.HasRows = True Then
                                objReader.Read()
                                Call .SetText(30, intLoopCounter, objReader.GetValue(1))
                                Call .SetText(31, intLoopCounter, objReader.GetValue(2))
                                Call .SetText(32, intLoopCounter, objReader.GetValue(3))
                                Call .SetText(33, intLoopCounter, objReader.GetValue(4))
                                Call .SetText(34, intLoopCounter, objReader.GetValue(5))
                                Call .SetText(35, intLoopCounter, objReader.GetValue(6))
                                'GST CHANGE
                            End If
                            objReader = Nothing
                            objSQLConn.Close()
                            objSQLConn = Nothing


                            'GST CHANGES

                    End Select
                End With
                rsSalesDtl.MoveNext()
            Next intLoopCounter
        End If
        If SpChEntry.MaxRows > 5 Then
            SpChEntry.ScrollBars = FPSpreadADO.ScrollBarsConstants.ScrollBarsBoth
        End If
        rsSalesDtl.ResultSetClose()
        rsSalesDtl = Nothing
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function

    Private Function ValidatebeforeSave(ByRef pstrMode As String) As Boolean

        On Error GoTo ErrHandler

        Dim lstrControls As String
        Dim lNo As Integer
        Dim lngcounter As Integer
        Dim lctrFocus As System.Windows.Forms.Control
        Dim rsCess As ClsResultSetDB
        Dim rsdb As ClsResultSetDB
        Dim strItemCode As String
        Dim strCustDrgNo As String
        Dim strSQL As String = ""
        Dim dblrate As Double
        Dim dbltoolcost As Double
        Dim dblquantity As Double
        Dim strexcisetype As String
        Dim dblexciseamt As Double
        Dim strCustDrgNoLists As String = ""
        Dim strhsnsaccodelist As String = ""

        ValidatebeforeSave = True
        lNo = 1
        lstrControls = ResolveResString(10059)

        Select Case UCase(Trim(pstrMode))
            Case "ADD"
                '10736222
                strSQL = "DELETE FROM TMP_CT2_INVOICE_KNOCKOFF where UNIT_CODE='" + gstrUNITID + "' and IP_ADDRESS='" & gstrIpaddressWinSck & "'"
                SqlConnectionclass.ExecuteNonQuery(strSQL)
                '10736222
                'strSQL = "select dbo.UDF_ISCT2INVOICE( '" & gstrUNITID & "','" & txtCustCode.Text.Trim & "','" & CmbInvType.Text.Trim & "','" & CmbInvSubType.Text.Trim & "','" & txtRefNo.Text.Trim & "')"
                'If IsRecordExists(strSQL) = True Then

                '    With SpChEntry
                '        For lngcounter = 1 To SpChEntry.MaxRows
                '            strItemCode = Nothing
                '            strCustDrgNo = Nothing
                '            .Row = lngcounter : .Col = 1 : strItemCode = .Text
                '            .Row = lngcounter : .Col = 2 : strCustDrgNo = .Text
                '            .Row = lngcounter : .Col = 3 : dblrate = .Text
                '            .Row = lngcounter : .Col = 5 : dblquantity = .Text
                '            .Row = lngcounter : .Col = 7 : strexcisetype = .Text
                '            .Row = lngcounter : .Col = 16 : dbltoolcost = .Text

                '            If blnISExciseRoundOff Then
                '                dblexciseamt = System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff))
                '            Else
                '                dblexciseamt = strSalesDtl & CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff)
                '            End If

                '            strSQL = "INSERT INTO TMP_CT2_INVOICE_KNOCKOFF ([UNIT_CODE],[CUST_CODE],[SONO],[AMENDMENT_NO],[TMP_INVOICE_NO],[ITEM_CODE],[CUST_DRG_NO],[CURRENCY_CODE],[QTY],[RATE],[TOOL_COST],[EXCISE_TAX],[EXCISE_AMOUNT],[ECESS_TYPE],[ECESS_AMOUNT],[SECESS_TYPE],[SECESS_AMOUNT],[VALUECONSUMED],[DUTYCONSUMED],[IP_ADDRESS]) "
                '            strSQL = strSQL + " Values('" & gstrUNITID & "','" & txtCustCode.Text.Trim & "','" & txtRefNo.Text.Trim & "','" & txtAmendNo.Text.Trim & "','" & Me.txtChallanNo.Text.Trim & "'," & _
                '            strSQL = strSQL + "'" & strItemCode & "','" & strCustDrgNo & "','" & lblCurrency.Text & "'," & dblquantity & "," & dblrate & "," & dbltoolcost & ""
                '        Next
                '    End With
                'End If

                If mblnCSM_Knockingoff_req = True Then
                    With SpChEntry
                        For lngcounter = 1 To SpChEntry.MaxRows
                            strItemCode = Nothing
                            strCustDrgNo = Nothing
                            .Row = lngcounter : .Col = 1 : strItemCode = .Text
                            .Row = lngcounter : .Col = 2 : strCustDrgNo = .Text
                            rsdb = New ClsResultSetDB
                            Call rsdb.GetResult("SELECT Customer_code,Finish_Item_Code,Item_Code,Cust_Drgno,Item_drgno,Description,Qty,Rate,Active_Flag,VALID_FROM,VALID_TO,GRIN_No,GRIN_Auth_Date FROM CUSTSUPPLIEDITEM_MST WHERE UNIT_CODE='" + gstrUNITID + "' AND  FINISH_ITEM_CODE =   '" & strItemCode.ToString.Trim & "' AND CUST_DRGNO =  '" & strCustDrgNo.ToString.Trim & "'  AND ACTIVE_FLAG = 1 ")
                            If rsdb.RowCount > 0 Then
                                If Validate_CSMRate() = False Then
                                    ValidatebeforeSave = False
                                    rsdb.ResultSetClose()
                                    rsdb = Nothing
                                    Exit Function
                                End If
                            End If
                            rsdb.ResultSetClose()
                            rsdb = Nothing
                        Next
                    End With
                End If

                If (Len(Me.txtLocationCode.Text) = 0) Then
                    lstrControls = lstrControls & vbCrLf & lNo & ". Location Code."
                    lNo = lNo + 1
                    If lctrFocus Is Nothing Then
                        lctrFocus = Me.txtLocationCode
                    End If
                    ValidatebeforeSave = False
                End If
                If Val(lblECSStax_Per.Text) > 0 And gblnGSTUnit = False Then
                    If Len(Trim(txtSECSSTaxType.Text)) = 0 Then
                        lstrControls = lstrControls & vbCrLf & lNo & ". Secondary ECESS"
                        lNo = lNo + 1
                        If lctrFocus Is Nothing Then
                            lctrFocus = txtSECSSTaxType
                        End If
                        ValidatebeforeSave = False
                    End If
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
                If (UCase(Trim(CmbInvType.Text)) = "NORMAL INVOICE") Or (UCase(Trim(CmbInvType.Text)) = "JOBWORK INVOICE") Or (UCase(Trim(CmbInvType.Text)) = "EXPORT INVOICE") Then
                    If (Trim(CmbInvSubType.Text) <> "SCRAP") Then
                        If (Len(Me.txtRefNo.Text) = 0) Then
                            lstrControls = lstrControls & vbCrLf & lNo & ". Reference No.."
                            lNo = lNo + 1
                            If lctrFocus Is Nothing Then
                                lctrFocus = Me.CmdRefNoHelp
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
                If (Len(Me.txtECSSTaxType.Text)) = 0 And gblnGSTUnit = False Then
                    With Me.SpChEntry
                        .Row = 1 : .Col = 7
                        If UCase(.Text) <> "EX0" And Len(.Text) > 0 Then
                            lstrControls = lstrControls & vbCrLf & lNo & ". Cess On ED."
                            lNo = lNo + 1
                            If lctrFocus Is Nothing Then
                                lctrFocus = Me.txtECSSTaxType
                            End If
                            ValidatebeforeSave = False
                        End If
                    End With
                Else
                    If gblnGSTUnit = False Then
                        rsCess = New ClsResultSetDB
                        Call rsCess.GetResult("Select TxRt_Rate_No from Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  TxRt_Rate_No = '" & Trim(UCase(Me.txtECSSTaxType.Text)) & "'")
                        If rsCess.RowCount <= 0 Then
                            lstrControls = lstrControls & vbCrLf & lNo & ". Cess On ED."
                            lNo = lNo + 1
                            If lctrFocus Is Nothing Then
                                lctrFocus = Me.txtECSSTaxType
                            End If
                            ValidatebeforeSave = False
                        End If
                        rsCess.ResultSetClose()
                    End If
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
                If txtCarrServices.Text = "" Then
                    lstrControls = lstrControls & vbCrLf & lNo & ". Carrier Name."
                    lNo = lNo + 1
                    If lctrFocus Is Nothing Then
                        lctrFocus = Me.txtCarrServices
                    End If
                    ValidatebeforeSave = False
                End If
                If txtVehNo.Text = "" Then
                    lstrControls = lstrControls & vbCrLf & lNo & ". Vehicle No."
                    lNo = lNo + 1
                    If lctrFocus Is Nothing Then
                        lctrFocus = Me.txtVehNo
                    End If
                    ValidatebeforeSave = False
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

                If AllowASNTextFileGeneration(Trim(txtCustCode.Text)) = True Then
                    If txtPlantCode.Text = "" Then
                        lstrControls = lstrControls & vbCrLf & lNo & ". Plant Code."
                        lNo = lNo + 1
                        If lctrFocus Is Nothing Then
                            lctrFocus = Me.txtPlantCode
                        End If
                        ValidatebeforeSave = False
                    End If
                End If
                '10869290
                If (UCase(Trim(CmbInvType.Text)) = "SERVICE INVOICE") And gstrUNITID <> "STH" Then
                    If (Len(Me.txtServiceTaxType.Text) = 0) And gblnGSTUnit = False Then
                        lstrControls = lstrControls & vbCrLf & lNo & ". Service Tax Code."
                        lNo = lNo + 1
                        If lctrFocus Is Nothing Then
                            lctrFocus = Me.cmdServiceTaxType
                        End If
                        ValidatebeforeSave = False
                    End If
                End If
                '10869290

                'KKC
                If (Len(Me.txtServiceTaxType.Text) > 0) Then
                    If (Len(Me.txtKKC.Text) = 0) And gblnGSTUnit = False Then
                        lstrControls = lstrControls & vbCrLf & lNo & ". KKC TAX."
                        lNo = lNo + 1
                        If lctrFocus Is Nothing Then
                            lctrFocus = Me.cmdkkccode
                        End If
                        ValidatebeforeSave = False
                    End If
                End If
                'KKC


                'SBC
                If (Len(Me.txtServiceTaxType.Text) > 0) Then
                    If (Len(Me.txtSBC.Text) = 0) And gblnGSTUnit = False Then
                        lstrControls = lstrControls & vbCrLf & lNo & ". SBC TAX."
                        lNo = lNo + 1
                        If lctrFocus Is Nothing Then
                            lctrFocus = Me.cmdSBC
                        End If
                        ValidatebeforeSave = False
                    End If
                End If
                'SBC


                '10808160
                strSQL = "select dbo.UDF_ISEOPINVOICE( '" & gstrUNITID & "','" & txtCustCode.Text.Trim & "','" & CmbInvType.Text.Trim & "','" & CmbInvSubType.Text.Trim & "','" & txtRefNo.Text.Trim & "')"
                If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSQL)) = True Then
                    lngcounter = 0
                    For lngcounter = 1 To SpChEntry.MaxRows
                        With SpChEntry
                            .Col = 27
                            .Row = lngcounter
                            If .Text.Trim.Length = 0 Then
                                .Col = 2
                                .Row = lngcounter
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

                'gst changes
                If gblnGSTUnit = True Then
                    lngcounter = 0
                    For lngcounter = 1 To SpChEntry.MaxRows
                        With SpChEntry
                            .Col = 30
                            .Row = lngcounter
                            If .Text.Trim.Length = 0 Then
                                .Col = 2
                                .Row = lngcounter
                                strhsnsaccodelist = strhsnsaccodelist + .Text.Trim + ","
                            End If
                        End With
                    Next

                    If strhsnsaccodelist.Trim.Length > 0 Then
                        lstrControls = lstrControls & vbCrLf & lNo & ".HSN/SAC CODE can't be blank for below item codes :" & vbCrLf & strhsnsaccodelist
                        lNo = lNo + 1
                        ValidatebeforeSave = False
                    End If
                End If
                    'gst changes
            Case "EDIT"
                '10736222
                strSQL = "DELETE FROM TMP_CT2_INVOICE_KNOCKOFF where UNIT_CODE='" + gstrUNITID + "' and IP_ADDRESS='" & gstrIpaddressWinSck & "'"
                SqlConnectionclass.ExecuteNonQuery(strSQL)
                '10736222
                If mblnCSM_Knockingoff_req = True Then
                    With SpChEntry
                        For lngcounter = 1 To SpChEntry.MaxRows
                            strItemCode = Nothing
                            strCustDrgNo = Nothing
                            .Row = lngcounter : .Col = 1 : strItemCode = .Text
                            .Row = lngcounter : .Col = 2 : strCustDrgNo = .Text
                            rsdb = New ClsResultSetDB
                            Call rsdb.GetResult("SELECT Customer_code,Finish_Item_Code,Item_Code,Cust_Drgno,Item_drgno,Description,Qty,Rate,Active_Flag,VALID_FROM,VALID_TO,GRIN_No,GRIN_Auth_Date FROM CUSTSUPPLIEDITEM_MST WHERE UNIT_CODE='" + gstrUNITID + "' AND  FINISH_ITEM_CODE =   '" & strItemCode.ToString.Trim & "' AND CUST_DRGNO =  '" & strCustDrgNo.ToString.Trim & "'  AND ACTIVE_FLAG = 1 ")
                            If rsdb.RowCount > 0 Then
                                If Validate_CSMRate() = False Then
                                    ValidatebeforeSave = False
                                    rsdb.ResultSetClose()
                                    rsdb = Nothing
                                    Exit Function
                                End If
                            End If
                            rsdb.ResultSetClose()
                            rsdb = Nothing
                        Next
                    End With
                End If


                If (Len(Me.txtECSSTaxType.Text)) = 0 And gblnGSTUnit = False Then
                    With Me.SpChEntry
                        .Row = 1 : .Col = 7
                        If UCase(.Text) <> "EX0" And Len(.Text) > 0 Then
                            lstrControls = lstrControls & vbCrLf & lNo & ". Cess On ED."
                            lNo = lNo + 1
                            If lctrFocus Is Nothing Then
                                lctrFocus = Me.txtECSSTaxType
                            End If
                            ValidatebeforeSave = False
                        End If
                    End With
                Else
                    If gblnGSTUnit = False Then
                        rsCess = New ClsResultSetDB
                        Call rsCess.GetResult("Select TxRt_Rate_No from Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  TxRt_Rate_No = '" & Trim(UCase(Me.txtECSSTaxType.Text)) & "'")
                        If rsCess.RowCount <= 0 Then
                            lstrControls = lstrControls & vbCrLf & lNo & ". Cess On ED."
                            lNo = lNo + 1
                            If lctrFocus Is Nothing Then
                                lctrFocus = Me.txtECSSTaxType
                            End If
                            ValidatebeforeSave = False
                        End If
                        rsCess.ResultSetClose()
                    End If
                End If
                If Val(lblECSStax_Per.Text) > 0 And gblnGSTUnit = False Then
                    If Len(Trim(txtSECSSTaxType.Text)) = 0 Then
                        lstrControls = lstrControls & vbCrLf & lNo & ". Secondary ECESS"
                        lNo = lNo + 1
                        If lctrFocus Is Nothing Then
                            lctrFocus = txtSECSSTaxType
                        End If
                        ValidatebeforeSave = False
                    End If
                End If
                If (Len(Me.txtFreight.Text) = 0) Then
                    txtFreight.Text = "0.00"
                End If
                If txtCarrServices.Text = "" Then
                    lstrControls = lstrControls & vbCrLf & lNo & ". Carrier Name."
                    lNo = lNo + 1
                    If lctrFocus Is Nothing Then
                        lctrFocus = Me.txtCarrServices
                        txtCarrServices.Enabled = True
                    End If
                    ValidatebeforeSave = False
                End If
                If txtVehNo.Text = "" Then
                    lstrControls = lstrControls & vbCrLf & lNo & ". Vehicle No."
                    lNo = lNo + 1
                    If lctrFocus Is Nothing Then
                        lctrFocus = Me.txtVehNo
                        txtVehNo.Enabled = True
                    End If
                    ValidatebeforeSave = False
                End If
                If AllowASNTextFileGeneration(Trim(txtCustCode.Text)) = True Then
                    If txtPlantCode.Text = "" Then
                        lstrControls = lstrControls & vbCrLf & lNo & ". Plant Code."
                        lNo = lNo + 1
                        If lctrFocus Is Nothing Then
                            lctrFocus = Me.txtPlantCode
                        End If
                        ValidatebeforeSave = False
                    End If
                End If

                '10808160
                strSQL = "select dbo.UDF_ISEOPINVOICE( '" & gstrUNITID & "','" & txtCustCode.Text.Trim & "','" & CmbInvType.Text.Trim & "','" & CmbInvSubType.Text.Trim & "','" & txtRefNo.Text.Trim & "')"
                If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSQL)) = True Then
                    lngcounter = 0
                    For lngcounter = 1 To SpChEntry.MaxRows
                        With SpChEntry
                            .Col = 27
                            .Row = lngcounter
                            If .Text.Trim.Length = 0 Then
                                .Col = 2
                                .Row = lngcounter
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
                '10808160

        End Select

        strSQL = "select dbo.UDF_ISDOCKCODE_INVOICE( '" & gstrUNITID & "','" & txtCustCode.Text.Trim & "','" & CmbInvType.Text.Trim.ToUpper & "','" & CmbInvSubType.Text.Trim.ToUpper & "' )"
        If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSQL)) = True Then
            If TxtLRNO.Text = "" Then
                lstrControls = lstrControls & vbCrLf & lNo & ". LR No."
                lNo = lNo + 1
                If lctrFocus Is Nothing Then
                    lctrFocus = Me.TxtLRNO
                    TxtLRNO.Enabled = True
                End If
                ValidatebeforeSave = False
            End If
        End If
        'prashant rajpal changed 
        Dim dblBatchQty As Double
        Dim dblqty As Object
        Dim strCompile As Object
        Dim ItemCode As Object
        Dim Intcounter As Short
        Dim ExciseCode As Object
        Dim strAtnstring As String

        If mblnRejTracking = True And mblnBatchTracking = True Then
            If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                If CmbInvType.Text <> "REJECTION" Then
                    GoTo Complete_Validation
                End If
            ElseIf CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                If strInvType <> "REJ" Then
                    GoTo Complete_Validation
                End If
            End If
            With SpChEntry
                For Intcounter = 1 To .MaxRows
                    strCompile = Nothing
                    .GetText(24, Intcounter, strCompile)
                    dblqty = Nothing
                    .GetText(5, Intcounter, dblqty)
                    ItemCode = Nothing
                    .GetText(1, Intcounter, ItemCode)
                    ExciseCode = Nothing
                    .GetText(7, Intcounter, ExciseCode)
                    If ExciseCode = "" And gblnGSTUnit = False And gstrUNITID <> "STH" Then
                        lstrControls = lstrControls & vbCrLf & lNo & " Excise Code Cannot be empty"
                        lNo = lNo + 1
                        If lctrFocus Is Nothing Then
                            SpChEntry.Row = SpChEntry.ActiveRow : SpChEntry.Col = 7 : SpChEntry.Action = 0 : SpChEntry.Focus()
                            'lctrFocus = Me.txtVehNo
                            'txtVehNo.Enabled = True
                        End If
                        ValidatebeforeSave = False
                    End If
                    dblBatchQty = CalculateDocumentQty(CStr(strCompile))
                    If Val(dblqty) <> Val(dblBatchQty) Then
                        lstrControls = lstrControls & vbCrLf & lNo & " Rejection Document Quantity is Not Equal To Entered Batch Quantity For Item Code :- " & ItemCode
                        lNo = lNo + 1
                        ValidatebeforeSave = False
                    End If
                Next
            End With
        ElseIf mblnRejTracking = True And mblnBatchTracking = False Then
            With SpChEntry
                For Intcounter = 1 To .MaxRows
                    strCompile = Nothing
                    .GetText(24, Intcounter, strCompile)
                    dblqty = Nothing
                    .GetText(5, Intcounter, dblqty)
                    ItemCode = Nothing
                    .GetText(1, Intcounter, ItemCode)
                    ExciseCode = Nothing
                    .GetText(7, Intcounter, ExciseCode)
                    If ExciseCode = "" And gblnGSTUnit = False And gstrUNITID <> "STH" Then
                        lstrControls = lstrControls & vbCrLf & lNo & " Excise Code Cannot be empty"
                        lNo = lNo + 1
                        If lctrFocus Is Nothing Then
                            SpChEntry.Row = SpChEntry.ActiveRow : SpChEntry.Col = 7 : SpChEntry.Action = 0 : SpChEntry.Focus()
                            'lctrFocus = Me.txtVehNo
                            'txtVehNo.Enabled = True
                        End If
                        ValidatebeforeSave = False
                    End If
                Next
            End With
        End If
        If mblnATN_invoicewise = True Then
            For Intcounter = 1 To SpChEntry.MaxRows
                With SpChEntry
                    .Col = 28
                    .Row = Intcounter
                    If .Text = "" Then
                        lstrControls = lstrControls & vbCrLf & lNo & ". ATN Field Can't be empty.."
                        lNo = lNo + 1
                        ValidatebeforeSave = False
                        Exit For
                    Else
                        If Intcounter <> SpChEntry.MaxRows Then
                            strAtnstring = "'" & strAtnstring & .Text & "',"
                        Else
                            strAtnstring = strAtnstring & "'" & .Text & "'"
                        End If
                    End If
                End With
            Next
            If ValidatebeforeSave = True Then
                If CInt(Find_Value("Select count(DISTINCT ASSETGROUP) from VW_FA_ATN WHERE UNIT_CODE='" + gstrUNITID + "' AND  ATNCode in  ( " & strAtnstring & " )")) > 1 Then
                    lstrControls = lstrControls & vbCrLf & lNo & ". More than 1 Asset Group (ATN )not possible .."
                    lNo = lNo + 1
                    ValidatebeforeSave = False
                End If
            End If
        End If


        'atn changes
Complete_Validation:
        If Not ValidatebeforeSave Then
            MsgBox(lstrControls, MsgBoxStyle.Information, ResolveResString(10059))
            If lctrFocus Is Nothing Then
            Else
                lctrFocus.Focus()
            End If
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
        Dim intloopcounter1 As Short
        If mblnBatchTrack = True And Me.CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD And UCase(CmbInvType.Text) <> "JOBWORK INVOICE" And UCase(CmbInvType.Text) <> "REJECTION INVOICE" Then
            If Me.SpChEntry.MaxRows > 0 Then
                ReDim mBatchData(Me.SpChEntry.MaxRows)
            End If
            For intloopcounter1 = 1 To Me.SpChEntry.MaxRows
                ReDim mBatchData(intloopcounter1).Batch_No(0)
                ReDim mBatchData(intloopcounter1).Batch_Date(0)
                ReDim mBatchData(intloopcounter1).Batch_Quantity(0)
            Next
        End If
        With Me.SpChEntry
            Select Case Me.CmdGrpChEnt.Mode
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                    If (UCase(Trim(CmbInvType.Text)) = "NORMAL INVOICE") Or (UCase(Trim(CmbInvType.Text)) = "EXPORT INVOICE") Or (UCase(Trim(CmbInvType.Text)) = "TRANSFER INVOICE") Or (UCase(Trim(CmbInvType.Text)) = "INTER-DIVISION") Then
                        If UCase(Trim(CmbInvSubType.Text)) <> "SCRAP" Then
                            For intRow = 1 To .MaxRows
                                .Row = intRow
                                For intcol = 1 To .MaxCols
                                    .Col = intcol
                                    If intcol = 5 Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                    ElseIf intcol = 12 Or intcol = 13 Or intcol = 16 Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                    ElseIf intcol = 9 Or intcol = 10 Or intcol = 15 Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                                    ElseIf intcol = 21 Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                                    ElseIf intcol = 23 Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                                    Else
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                                    End If
                                Next intcol
                            Next intRow
                            If GetPaletteStatus(CmbInvType.Text, CmbInvSubType.Text, txtCustCode.Text.Trim()) Then
                                PaletteReadOnlyQty()
                            End If
                        Else
                            For intRow = 1 To .MaxRows
                                .Row = intRow
                                For intcol = 1 To .MaxCols
                                    .Col = intcol
                                    If intcol = 5 Or intcol = 12 Or intcol = 13 Or intcol = 16 Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                    ElseIf intcol = 3 Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = 4 : .TypeFloatMin = CDbl("0.0000") : .TypeFloatMax = CDbl("99999999999999.99999")
                                    ElseIf intcol = 15 Or intcol = 9 Or intcol = 10 Or intcol = 7 Then
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
                                If intcol = 5 Or intcol = 12 Or intcol = 13 Or intcol = 16 Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                ElseIf intcol = 3 Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = 4 : .TypeFloatMin = CDbl("0.0000") : .TypeFloatMax = CDbl("99999999999999.99999")
                                ElseIf intcol = 15 Or intcol = 9 Or intcol = 10 Or intcol = 7 Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                                ElseIf intcol = 21 And mblnBatchTrack = True And UCase(CmbInvType.Text) <> "JOBWORK INVOICE" And UCase(CmbInvType.Text) <> "REJECTION" Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                                ElseIf intcol = 25 And mblnRejTracking = True And (strInvType = "REJ" Or UCase(CmbInvType.Text) = "REJECTION") And mblnBatchTracking = True Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton

                                Else
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                                End If
                            Next intcol
                        Next intRow
                    End If
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                    If (UCase(strInvType) = "INV") Or (UCase(strInvType) = "EXP") Then
                        If (UCase(strInvSubType) <> "L") Then
                            For intRow = 1 To .MaxRows
                                .Row = intRow
                                For intcol = 1 To .MaxCols
                                    .Col = intcol
                                    If intcol = 5 Or intcol = 12 Or intcol = 13 Or intcol = 16 Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                    ElseIf intcol = 15 Or intcol = 9 Or intcol = 10 Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                                    ElseIf intcol = 21 Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                                    ElseIf intcol = 23 And strInvSubType = "F" Then
                                        .ColHidden = False : .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                                    Else
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                                    End If
                                Next intcol
                            Next intRow
                            If GetPaletteStatus(UCase(strInvType), UCase(strInvSubType), txtCustCode.Text.Trim()) Then
                                PaletteReadOnlyQty()
                            End If
                        Else
                            For intRow = 1 To .MaxRows
                                .Row = intRow
                                For intcol = 1 To .MaxCols
                                    .Col = intcol
                                    If intcol = 5 Or intcol = 12 Or intcol = 13 Or intcol = 16 Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                    ElseIf intcol = 3 Then
                                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = 4 : .TypeFloatMin = CDbl("0.0000") : .TypeFloatMax = CDbl("99999999999999.99999")
                                    ElseIf intcol = 15 Or intcol = 9 Or intcol = 10 Or intcol = 7 Then
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
                                If intcol = 5 Or intcol = 12 Or intcol = 13 Or intcol = 16 Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                ElseIf intcol = 3 Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = 4 : .TypeFloatMin = CDbl("0.0000") : .TypeFloatMax = CDbl("99999999999999.99999")
                                ElseIf intcol = 15 Or intcol = 9 Or intcol = 10 Or intcol = 7 Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                                ElseIf intcol = 21 Then
                                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
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
        '*****************************************************
        'Created By     -  Nisha
        'Description    -  To Check Schedule Quantity From DailyMktSchedule/MonthlyMktSchedule
        '*****************************************************
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
        Dim intLoopcount As Short
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
        mstrUpdDispatchSql = ""
        For intRwCount = 1 To SpChEntry.MaxRows
            VarDelete = Nothing
            Call SpChEntry.GetText(15, intRwCount, VarDelete)
            If UCase(VarDelete) <> "D" Then
                For intcol = 1 To SpChEntry.MaxCols
                    SpChEntry.Col = intcol
                    If (SpChEntry.Col = 5) Or (SpChEntry.Col = 3) Or (SpChEntry.Col = 13) Or (SpChEntry.Col = 12) Then ''Column Changed By Tapan
                        SpChEntry.Row = intRwCount
                        If (Val(Trim(SpChEntry.Text)) = 0) Then
                            QuantityCheck = True
                            Call ConfirmWindow(10419, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                            SpChEntry.Row = intRwCount : SpChEntry.Col = intcol : SpChEntry.Action = 0 : SpChEntry.Focus()
                            Exit Function
                        End If
                        If (SpChEntry.Col = 13) Then
                            SpChEntry.Row = intRwCount : SpChEntry.Col = 12 : intFromBox = Val(Trim(SpChEntry.Text))
                            SpChEntry.Row = intRwCount : SpChEntry.Col = 13
                            If Val(Trim(SpChEntry.Text)) < intFromBox Then
                                QuantityCheck = True
                                Call ConfirmWindow(10235, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                                SpChEntry.Row = intRwCount : SpChEntry.Col = 13 : SpChEntry.Action = 0 : SpChEntry.Focus()
                                Exit Function
                            End If
                        End If
                    End If
                Next intcol
            End If
        Next intRwCount
        '********************************
        'Validation for Schedule Quantity Start Here
        '********************************
        If ValidateScheduleQuantity() = False Then QuantityCheck = True : Exit Function
        '*******************************
        'Validation for Schedule Quantity End Here
        '********************************
        '****************************************
        'To check Current Balance from Itembal_Mst
        'If Quantity Entered Is Greater Then Cur_Bal In The ItemBal_Mst
        'Then Restrict User To Change The Entered Quantity
        '******************************************
        'To Get Item Code From Spread
        Dim strItCode As String 'To Make Item Code String
        For intRwCount = 1 To Me.SpChEntry.MaxRows
            VarDelete = Nothing
            Call Me.SpChEntry.GetText(15, intRwCount, VarDelete)
            If UCase(VarDelete) <> "D" Then
                varItemCode = Nothing
                Call Me.SpChEntry.GetText(1, intRwCount, varItemCode)
                strItCode = strItCode & "'" & Trim(varItemCode) & "',"
            End If
        Next intRwCount
        strItCode = Mid(strItCode, 1, Len(strItCode) - 1)
        rsSaleConf = New ClsResultSetDB
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT, UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                rsSaleConf.GetResult(" Select Invoice_type,Sub_Category from SalesChallan_Dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  Doc_No=" & txtChallanNo.Text)
                mstrInvoiceType = rsSaleConf.GetValue("Invoice_Type")
                mstrInvSubType = rsSaleConf.GetValue("Sub_Category")
                rsSaleConf.ResultSetClose()
                rsSaleConf = New ClsResultSetDB
                rsSaleConf.GetResult("select Stock_Location From saleconf (nolock) WHERE UNIT_CODE='" + gstrUNITID + "' AND  invoice_type ='" & Trim(mstrInvoiceType) & "' and sub_type ='" & Trim(mstrInvSubType) & "' AND Location_Code ='" & Trim(txtLocationCode.Text) & "' and (fin_start_date <= getdate() and fin_end_date >= getdate())")
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                rsSaleConf.GetResult("select Stock_Location From saleconf (nolock) WHERE UNIT_CODE='" + gstrUNITID + "' AND  Description ='" & Trim(CmbInvType.Text) & "' and sub_type_Description ='" & Trim(CmbInvSubType.Text) & "' AND Location_Code ='" & Trim(txtLocationCode.Text) & "' and (fin_start_date <= getdate() and fin_end_date >= getdate())")
        End Select
        If Len(Trim(rsSaleConf.GetValue("Stock_Location"))) = 0 Then
            MsgBox("Please Define Stock Location in Sales Conf First", MsgBoxStyle.OkOnly, "eMPro")
            QuantityCheck = True
            rsSaleConf.ResultSetClose()
            Exit Function
        End If
        Dim varItemCodeinVeiw As Object
        If mstrInvoiceType <> "SRC" Then
            For intRwCount = 1 To Me.SpChEntry.MaxRows
                varItemCodeinVeiw = Nothing
                Call SpChEntry.GetText(1, intRwCount, varItemCodeinVeiw)
                varDrgNo = Nothing
                Call SpChEntry.GetText(2, intRwCount, varDrgNo)
                VarDelete = Nothing
                Call SpChEntry.GetText(15, intRwCount, VarDelete)
                If UCase(VarDelete) <> "D" Then
                    strItembal = "Select Cur_Bal From ItemBal_Mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  Location_Code ='" & Trim(rsSaleConf.GetValue("Stock_Location")) & "' and item_Code ='" & varItemCodeinVeiw & "'"
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
                    If Not (UCase(CmbInvType.Text) = "REJECTION" Or UCase(strInvType) = "REJ") And AllowASNTextFileGeneration(Trim(txtCustCode.Text)) = True Then
                        Dim strContainer As String
                        Dim strCustDrgNolist As String
                        strContainer = Find_Value("select container from custitem_mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  cust_drgno='" & varDrgNo & "' and item_code='" & varItemCodeinVeiw & "' and account_code='" & Trim(txtCustCode.Text) & "' and active=1")
                        If strContainer.Length() = 0 Then
                            MessageBox.Show("Container is not defined of Customer Part Code : " & varDrgNo, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                            QuantityCheck = True
                            Exit Function
                        End If
                        strCustDrgNolist = strCustDrgNolist & "'" & varDrgNo & "',"
                        If intRwCount = Me.SpChEntry.MaxRows Then
                            strCustDrgNolist = Mid(strCustDrgNolist, 1, strCustDrgNolist.Length() - 1)
                            strContainer = Find_Value("select count(distinct(container)) from custitem_mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  cust_drgno in(" & strCustDrgNolist & ") and account_code='" & Trim(txtCustCode.Text) & "' and active=1")
                            If Val(strContainer) > 1 Then
                                MessageBox.Show("Select the Customer Part Code of same container type.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                                QuantityCheck = True
                                Exit Function
                            End If
                        End If
                    End If
                End If
            Next intRwCount

        End If
        rsSaleConf.ResultSetClose()
        rsSaleConf = Nothing
        '****************************************
        'To check if tool Amortization Check is required
        'then in Invoice if Tool Amortization is there or not
        'to check if this qty is available in Tool Amortization details
        'Added by nisha on 17/02/2004
        '******************************************
        rsSalesParameter = New ClsResultSetDB
        rsSalesParameter.GetResult("Select CheckToolAmortisation from Sales_Parameter WHERE UNIT_CODE='" + gstrUNITID + "'")
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
                    Call SpChEntry.GetText(16, intRwCount, varToolCost)
                    VarDelete = Nothing
                    Call SpChEntry.GetText(15, intRwCount, VarDelete)
                    If UCase(VarDelete) <> "D" Then
                        With mP_Connection
                            .Execute("DELETE FROM tmpBOM WHERE UNIT_CODE='" + gstrUNITID + "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            .Execute("BOMExplosion '" & Trim(varItemCodeinVeiw) & "','" & Trim(varItemCodeinVeiw) & "',1,0,'" + gstrUNITID + "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        End With
                        rsbom = New ClsResultSetDB
                        rsbom.GetResult("select no,slno,item_code,Description,Material,UOM,Quantity,WasteQty,ScrapQty,RunnerQty,GrossWeight,Alternate_Item,FinishedItem from tmpBOM WHERE UNIT_CODE='" + gstrUNITID + "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        If rsbom.GetNoRows > 0 Then
                            irowcount = rsbom.GetNoRows
                            rsbom.MoveFirst()
                            For intRwCount1 = 1 To irowcount
                                strItembal = "select BalanceQty = isnull(a.proj_qty,0) - isnull(a.ClosingValueSMIEL,0),a.Tool_c from Amor_dtl a, tool_mst b "
                                strItembal = strItembal & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  account_code = '" & Trim(txtCustCode.Text) & "'"
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
                                    strItembal = strItembal & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  "
                                    strItembal = strItembal & " Item_code = '" & rsbom.GetValue("item_code") & "' and tool_c = '" & strToolCode & "'"
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
                            '-------------------------------------------------------
                        End If
                        strItembal = "select BalanceQty = isnull(a.proj_qty,0) - isnull(a.ClosingValueSMIEL,0),a.Tool_c from Amor_dtl a,Tool_Mst b"
                        strItembal = strItembal & " WHERE a.UNIT_CODE='" + gstrUNITID + "' AND  account_code = '" & Trim(txtCustCode.Text) & "'"
                        strItembal = strItembal & " and Item_code = '" & varItemCodeinVeiw & "' AND A.UNIT_CODE=B.UNIT_CODE AND a.Tool_c = b.tool_c and a.item_code = b.Product_No order by a.tool_c"
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
                            strItembal = strItembal & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  "
                            strItembal = strItembal & " Item_code = '" & rsbom.GetValue("item_code") & "' and tool_c = '" & strToolCode & "'"
                            rsbom.ResultSetClose()
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
                            strItembal = strItembal & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  account_code = '" & Trim(txtCustCode.Text) & "'"
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
                                strItembal = strItembal & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  "
                                strItembal = strItembal & " Item_code = '" & varItemCodeinVeiw & "'"
                                rsMktSchedule = New ClsResultSetDB
                                rsMktSchedule.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                                rsMktSchedule.MoveFirst()
                                strQuantity = CStr(CDbl(strQuantity) - Val(rsMktSchedule.GetValue("BalanceQty")))
                                'changes Ends Here by nisha on 22 Nov 2004
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
        End If
        rsSalesParameter.ResultSetClose()
        '*******
        '*****************************************
        'to check quantity available in CustAnnex_dtl
        'in case of JobWork Order
        '*****************************************
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
                ctlInsurance.Text = "" : lblCurrencyDes.Text = "" : txtRefNo.Text = ""
                txtServiceTaxType.Text = ""
                CmbInvType.SelectedIndex = -1 : CmbInvSubType.SelectedIndex = -1
                Me.CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                Me.CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
            Case "CHALLAN"
                txtChallanNo.Text = "" : txtCustCode.Text = "" : lblCustCodeDes.Text = "" : lblAddressDes.Text = ""
                txtCarrServices.Text = "" : txtVehNo.Text = ""
                txtFreight.Text = "" : txtSaleTaxType.Text = "" : lblSaltax_Per.Text = "0.00"
                txtSurchargeTaxType.Text = "" : lblSurcharge_Per.Text = "0.00"
                ctlInsurance.Text = "" : lblRGPDes.Text = ""
                txtServiceTaxType.Text = ""
                txtKKC.Text = "" : txtSBC.Text = ""
                CmbInvType.SelectedIndex = -1 : CmbInvSubType.SelectedIndex = -1 : lblCurrencyDes.Text = "" : txtRefNo.Text = ""
                Me.CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                Me.CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
        End Select
        With Me.SpChEntry
            .MaxRows = 1
            .Row = 1 : .Row2 = 1 : .Col = 1 : .Col2 = .MaxCols : .BlockMode = True : .Text = "" : .BlockMode = False
        End With
        strupSalechallan = ""
        strupSaleDtl = ""
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
    Private Sub SelectChallanNoFromSalesChallanDtl(ByVal intExist As Integer)
        '*****************************************************
        'Created By     -  Kapil
        'Description    -  To Select Max.  Challan No. From SalesChallan_Dtl
        '*****************************************************
        On Error GoTo ErrHandler
        Dim strSQL As String
        Dim strUpdateSQL As String
        Dim strChallanNo As String
        Dim rsChallanNo As ClsResultSetDB
        Dim strdateofInvoice As String
        strdateofInvoice = getDateForDB(dtpDateDesc.Text)
        If intExist = 1 Then
            strSQL = "Select Location_Code,Doc_No,Suffix,Transport_Type,Vehicle_No,From_Station,To_Station,Invoice_Date,Account_Code,Cust_Ref,Amendment_No,Bill_Flag,Print_DateTime,Form3,Form3Date,Carriage_Name,Year,Insurance,Frieght_Tax,Invoice_Type,Ref_Doc_No,Cust_Name,Sales_Tax_Amount,Surcharge_Sales_Tax_Amount,Frieght_Amount,Packing_Amount,Sub_Category,SalesTax_Type,SalesTax_FormNo,SalesTax_FormValue,Annex_no,Currency_Code,Nature_of_Contract,OriginStatus,Ctry_Destination_Goods,Delivery_Terms,Payment_Terms,Pre_Carriage_By,Receipt_Precarriage_at,Vessel_Flight_number,Port_Of_Loading,Port_Of_Discharge,Final_destination,Mode_Of_Shipment,Dispatch_mode,Buyer_description_Of_Goods,Invoice_description_of_EPC,Exchange_Date,Buyer_Id,Exchange_Rate,total_quantity,total_amount,TurnoverTax_per,Turnover_amt,other_ref,FIFO_flag,Surcharge_salesTaxType,SalesTax_Per,Surcharge_SalesTax_Per,Ent_dt,Ent_UserId,Upd_dt,Upd_Userid,Print_Flag,Cancel_flag,pervalue,remarks,dataPosted,ftp,Excise_Type,SRVDINO,SRVLocation,ExciseExumpted,LoadingChargeTaxType,LoadingChargeTaxAmount,LoadingChargeTax_Per,ConsigneeContactPerson,ConsigneeAddress1,ConsigneeAddress2,ConsigneeAddress3,ConsigneeECCNo,ConsigneeLST,ServiceInvoiceformatExport,CustBankID,Discount_Type,Discount_Amount,Discount_Per,RejectionPosting,USLOC,SchTime,TCSTax_Type,TCSTax_Per,TCSTaxAmount,To_Location,ECESS_Type,ECESS_Per,ECESS_Amount,FOC_Invoice,From_Location,PrintExciseFormat,FreshCrRecd,SRCESS_Type,SRCESS_Per,SRCESS_Amount,CVDCESS_Type,CVDCESS_Per,CVDCESS_Amount,Excise_Percentage,invoice_time,Permissible_Limit,TurnOverTaxType,TotalInvoiceAmtRoundOff_diff,NRGPNOIncaseOfServiceInvoice,Trans_Parameter_Flag,SDTax_Type,SDTax_Per,SDTax_Amount,InvoiceAgainstMultipleSO,TextFileGenerated,sameunitloading,ServiceTax_Type,ServiceTax_Per,ServiceTax_Amount,Prev_Yr_ExportSales,Permissible_Limit_SmpExport,varGeneralRemarks,SECESS_Type,SECESS_Per,SECESS_Amount,CVDSECESS_Type,CVDSECESS_Per,CVDSECESS_Amount,SRSECESS_Type,SRSECESS_Per,SRSECESS_Amount,postingFlag,CheckSheetNo,MULTIPLESO,ISCHALLAN,ISCONSOLIDATE,Tot_Add_Excise_Amt,Tot_Add_Excise_PER,CONSIGNEE_CODE,Lorry_No,OTL_No,RefChallan,price_bases,LorryNo_date,dataposted_fin,ConsInvString,bond17OpeningBal,barCodeImage,invoicepicking_status,Ecess_TotalDuty_Type,Ecess_TotalDuty_Per,Ecess_TotalDuty_Amount,SEcess_TotalDuty_Type,SEcess_TotalDuty_Per,SEcess_TotalDuty_Amount,ADDVAT_Type,ADDVAT_Per,ADDVAT_Amount from saleschallan_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  Doc_no = '" & txtChallanNo.Text & "' "
            rsChallanNo = New ClsResultSetDB
            rsChallanNo.GetResult(strSQL, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If rsChallanNo.GetNoRows > 0 Then
                rsChallanNo.ResultSetClose()
                rsChallanNo = Nothing
                GoTo Label
            Else
                rsChallanNo.ResultSetClose()
                rsChallanNo = Nothing
                Exit Sub
            End If
        End If
Label:
        strSQL = "Select Current_No From  DocumentType_Mst (nolock) WHERE UNIT_CODE='" + gstrUNITID + "' AND  "
        strSQL = strSQL + " Doc_Type = 9999  AND fin_start_date <= CONVERT(DateTime,'" & strdateofInvoice & "',103) "
        strSQL = strSQL + " And Fin_End_date >= Convert(datetime,'" & strdateofInvoice & "',103)"
        rsChallanNo = New ClsResultSetDB
        rsChallanNo.GetResult(strSQL, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsChallanNo.GetNoRows > 0 Then
            strChallanNo = (rsChallanNo.GetValue("Current_No") + 1).ToString
            strUpdateSQL = "UPDATE DocumentType_Mst with (ROWLOCK) Set Current_No = " & CLng(strChallanNo) & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  "
            strUpdateSQL = strUpdateSQL + " Doc_Type = 9999 AND fin_start_date <= CONVERT(DateTime,'" & strdateofInvoice & "',103) "
            strUpdateSQL = strUpdateSQL + " And Fin_End_date >= Convert(datetime,'" & strdateofInvoice & "',103) "
            mP_Connection.Execute(strUpdateSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            While Len(strChallanNo) < 6
                strChallanNo = "0" + strChallanNo
            End While
            strChallanNo = "99" + strChallanNo
            Me.txtChallanNo.Text = strChallanNo
        Else
            Me.txtChallanNo.Text = "99000001"
        End If
        rsChallanNo.ResultSetClose()
        rsChallanNo = Nothing
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
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
        Dim strSQL As String
        Dim rs_taxes As ClsResultSetDB
        'To Get Data from Cusft_Ord_hdr
        '***************************************
        Select Case CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT, UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                strCustOrdHdr = "Select max(Order_date),SalesTax_Type,"
                strCustOrdHdr = strCustOrdHdr & "Currency_Code,PerValue,ECESS_Code,term_payment from Cust_ord_hdr"
                strCustOrdHdr = strCustOrdHdr & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Account_Code='" & txtCustCode.Text & "' and Cust_Ref ='"
                strCustOrdHdr = strCustOrdHdr & mstrRefNo & "'and Amendment_No ='" & mstrAmmNo & "' Group By salestax_type,currency_code,ECESS_Code,term_payment"
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
                txtECSSTaxType.Text = rsCustOrdHdr.GetValue("ECESS_Code")
                Call txtECSSTaxType_Validating(txtECSSTaxType, New System.ComponentModel.CancelEventArgs(False))
                lblcreditterm.Text = rsCustOrdHdr.GetValue("term_payment")
                Call FN_ECSSH_TAX_Disp()
                rsCustOrdHdr.ResultSetClose()
                rsCustOrdHdr = Nothing
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                '10869290
                'If UCase(Trim(CmbInvType.Text)) = "NORMAL INVOICE" Or UCase(Trim(CmbInvType.Text)) = "JOBWORK INVOICE" Or UCase(Trim(CmbInvType.Text)) = "EXPORT INVOICE" Or UCase(Trim(CmbInvType.Text)) = "TRANSFER INVOICE" Or UCase(Trim(CmbInvType.Text)) = "SERVICE INVOICE" Then
                If UCase(Trim(CmbInvType.Text)) = "NORMAL INVOICE" Or UCase(Trim(CmbInvType.Text)) = "JOBWORK INVOICE" Or UCase(Trim(CmbInvType.Text)) = "EXPORT INVOICE" Or UCase(Trim(CmbInvType.Text)) = "TRANSFER INVOICE" Or UCase(Trim(CmbInvType.Text)) = "SERVICE INVOICE" Or UCase(Trim(CmbInvType.Text)) = "INTER-DIVISION" Then
                    If CBool(UCase(CStr((Trim(CmbInvSubType.Text)) <> "SCRAP"))) Then
                        If Len(Trim(txtRefNo.Text)) Then
                            strCustOrdHdr = "Select max(Order_date),SalesTax_Type,Currency_code,PerValue ,surcharge_code,ECESS_Code,term_payment,SERVICETAX_TYPE ,SBCTAX_TYPE,KKCTAX_TYPE  from Cust_ord_hdr"
                            strCustOrdHdr = strCustOrdHdr & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Account_Code='" & txtCustCode.Text & "' and Cust_Ref in('" & CUSTREFLIST & "') and Authorized_Flag = 1"
                            strCustOrdHdr = strCustOrdHdr & " and active_flag = 'A' Group by salestax_type,currency_code,PerValue,surcharge_code,ECESS_Code,term_payment,SERVICETAX_TYPE,SBCTAX_TYPE,KKCTAX_TYPE  "
                            rsCustOrdHdr = New ClsResultSetDB
                            rsCustOrdHdr.GetResult(strCustOrdHdr, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                            txtSurchargeTaxType.Text = rsCustOrdHdr.GetValue("surcharge_code")
                            Call txtSurchargeTaxType_Validating(txtSurchargeTaxType, New System.ComponentModel.CancelEventArgs(False))
                            txtServiceTaxType.Text = rsCustOrdHdr.GetValue("SERVICETAX_TYPE")
                            lblServiceTax_Per.Text = CStr(GetTaxRate((txtServiceTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " (Tx_TaxeID='SRT' )"))
                            txtSBC.Text = rsCustOrdHdr.GetValue("SBCTAX_TYPE")
                            lblSBC.Text = CStr(GetTaxRate((txtSBC.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " (Tx_TaxeID='sbc' )"))
                            txtKKC.Text = rsCustOrdHdr.GetValue("KKCTAX_TYPE")
                            lblKKC.Text = CStr(GetTaxRate((txtKKC.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " (Tx_TaxeID='KKC' )"))

                            strCurrency = rsCustOrdHdr.GetValue("Currency_code")
                            ctlPerValue.Text = rsCustOrdHdr.GetValue("PerValue")
                            If CBool(UCase(CStr((Trim(CmbInvType.Text)) = "EXPORT INVOICE"))) Or CBool(UCase(CStr((Trim(CmbInvType.Text)) = "SERVICE INVOICE"))) Then
                                lblCurrencyDes.Text = strCurrency
                            End If
                            intDecimalPlace = ToGetDecimalPlaces(strCurrency)
                            If intDecimalPlace < 2 Then
                                intDecimalPlace = 2
                            End If
                            lblcreditterm.Text = rsCustOrdHdr.GetValue("term_payment")
                            ctlInsurance.DecSize = intDecimalPlace : txtFreight.DecSize = intDecimalPlace
                            rsCustOrdHdr.ResultSetClose()
                            rsCustOrdHdr = Nothing
                        End If
                    End If
                End If
                strSQL = "Select isnull(ECESS_Code,'') as ECESS_Code from Cust_ord_hdr"
                strSQL = strSQL & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Account_Code='" & txtCustCode.Text & "' and Cust_Ref in('" & CUSTREFLIST & "') and Authorized_Flag = 1"
                strSQL = strSQL & " and active_flag = 'A' "
                rs_taxes = New ClsResultSetDB
                rs_taxes.GetResult(strSQL, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                If rs_taxes.GetNoRows > 0 Then
                    If Len(Trim(rs_taxes.GetValue("ECESS_Code"))) > 0 Then
                        txtECSSTaxType.Text = rs_taxes.GetValue("ECESS_Code")
                        rs_taxes.ResultSetClose()
                        rs_taxes = Nothing
                    Else
                        '------------------Satvir Handa------------------------
                        'strSQL = "select TxRt_Rate_No,TxRt_Percentage from Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  Tx_TaxeID='ECS' AND DEFAULT_FOR_INVOICE = 1 "
                        strSQL = "select TxRt_Rate_No,TxRt_Percentage from Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  Tx_TaxeID='ECS' AND DEFAULT_FOR_INVOICE = 1 and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
                        '------------------Satvir Handa------------------------
                        rs_taxes = New ClsResultSetDB
                        rs_taxes.GetResult(strSQL, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        If rs_taxes.GetNoRows > 0 And gblnGSTUnit = False Then
                            txtECSSTaxType.Text = rs_taxes.GetValue("TxRt_Rate_No")
                            rs_taxes.ResultSetClose()
                            rs_taxes = Nothing
                        Else
                            txtECSSTaxType.Text = ""
                            rs_taxes.ResultSetClose()
                            rs_taxes = Nothing
                        End If
                    End If
                Else
                    '------------------Satvir Handa------------------------
                    'strSQL = "select TxRt_Rate_No,TxRt_Percentage from Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND Tx_TaxeID='ECS' AND DEFAULT_FOR_INVOICE = 1"
                    strSQL = "select TxRt_Rate_No,TxRt_Percentage from Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND Tx_TaxeID='ECS' AND DEFAULT_FOR_INVOICE = 1 and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
                    '------------------Satvir Handa------------------------
                    rs_taxes = New ClsResultSetDB
                    rs_taxes.GetResult(strSQL, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                    If rs_taxes.GetNoRows > 0 Then
                        txtECSSTaxType.Text = rs_taxes.GetValue("TxRt_Rate_No")
                        rs_taxes.ResultSetClose()
                        rs_taxes = Nothing
                    Else
                        txtECSSTaxType.Text = ""
                        rs_taxes.ResultSetClose()
                        rs_taxes = Nothing
                    End If
                End If
                Call txtECSSTaxType_Validating(txtECSSTaxType, New System.ComponentModel.CancelEventArgs(False))
                '' For Sale Tax Code
                strSQL = "Select isnull(SalesTax_Type,'') as SalesTax_Type from Cust_ord_hdr"
                strSQL = strSQL & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Account_Code='" & txtCustCode.Text & "' and Cust_Ref in('" & CUSTREFLIST & "') and Authorized_Flag = 1"
                strSQL = strSQL & " and active_flag = 'A' "
                rs_taxes = New ClsResultSetDB
                rs_taxes.GetResult(strSQL, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                If rs_taxes.GetNoRows > 0 Then
                    If Len(Trim(rs_taxes.GetValue("SalesTax_Type"))) > 0 Then
                        txtSaleTaxType.Text = rs_taxes.GetValue("SalesTax_Type")
                        rs_taxes.ResultSetClose()
                        rs_taxes = Nothing
                    Else
                        strSQL = "select TxRt_Rate_No,TxRt_Percentage from Gen_TaxRate " &
                                " WHERE UNIT_CODE='" + gstrUNITID + "' AND  (Tx_TaxeID='CST' OR Tx_TaxeID='LST' or Tx_TaxeID='VAT') AND DEFAULT_FOR_INVOICE = 1"
                        rs_taxes = New ClsResultSetDB
                        rs_taxes.GetResult(strSQL, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        If rs_taxes.GetNoRows > 0 Then
                            txtSaleTaxType.Text = rs_taxes.GetValue("TxRt_Rate_No")
                            rs_taxes.ResultSetClose()
                            rs_taxes = Nothing
                        Else
                            txtSaleTaxType.Text = ""
                            rs_taxes.ResultSetClose()
                            rs_taxes = Nothing
                        End If
                    End If
                Else
                    strSQL = "select TxRt_Rate_No,TxRt_Percentage from Gen_TaxRate " &
                            " WHERE UNIT_CODE='" + gstrUNITID + "' AND  (Tx_TaxeID='CST' OR Tx_TaxeID='LST' or Tx_TaxeID='VAT') AND DEFAULT_FOR_INVOICE = 1"
                    rs_taxes = New ClsResultSetDB
                    rs_taxes.GetResult(strSQL, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                    If rs_taxes.GetNoRows > 0 Then
                        txtSaleTaxType.Text = rs_taxes.GetValue("TxRt_Rate_No")
                        rs_taxes.ResultSetClose()
                        rs_taxes = Nothing
                    Else
                        txtSaleTaxType.Text = ""
                        rs_taxes.ResultSetClose()
                        rs_taxes = Nothing
                    End If
                End If
                Call txtSaleTaxType_Validating(txtSaleTaxType, New System.ComponentModel.CancelEventArgs(False))
                'strSQL = "select TxRt_Rate_No,TxRt_Percentage from Gen_TaxRate " & _
                '        " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Tx_TaxeID='ECSSH' and TxRt_Percentage=1 and DEFAULT_FOR_INVOICE = 1 "

                '------------------Satvir Handa------------------------
                strSQL = "select TxRt_Rate_No,TxRt_Percentage from Gen_TaxRate " &
                        " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Tx_TaxeID='ECSSH' and DEFAULT_FOR_INVOICE =1 and ((isnull(deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= deactive_date))"
                '------------------Satvir Handa------------------------
                rs_taxes = New ClsResultSetDB
                rs_taxes.GetResult(strSQL, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                If rs_taxes.GetNoRows > 0 And gblnGSTUnit = False Then
                    txtSECSSTaxType.Text = rs_taxes.GetValue("TxRt_Rate_No")
                    rs_taxes.ResultSetClose()
                    rs_taxes = Nothing
                Else
                    txtSECSSTaxType.Text = ""
                    rs_taxes.ResultSetClose()
                    rs_taxes = Nothing
                End If
                Call txtSECSSTaxType_Validating(txtSECSSTaxType, New System.ComponentModel.CancelEventArgs(False))
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
                .Col = 1 : .TypeMaxEditLen = 17
                .Col = 2 : .TypeMaxEditLen = 30
                .Col = 3 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = pintDecimalSize : .TypeFloatMin = CDbl(strMin) : .TypeFloatMax = CDbl(strMax)
                .Col = 4 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = pintDecimalSize : .TypeFloatMin = CDbl(strMin) : .TypeFloatMax = CDbl(strMax)
                'If CmbInvType.Text = "NORMAL INVOICE" And CmbInvSubType.Text = "FINISHED GOODS" And DataExist("SELECT PCA_CUSTOMER FROM CUSTOMER_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & Trim(txtCustCode.Text) & "' and PCA_CUSTOMER =1 ") = True Then
                '.Col = 5 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Row = .ActiveRow : .Row2 = .ActiveRow : .Col = 5 : .Col2 = 5 : .BlockMode = True : .Lock = True : .BlockMode = False
                'Else
                .Col = 5 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = 2 : .TypeFloatMin = CDbl("0.00") : .TypeFloatMax = CDbl("99999999999999.99")
                'End If

                .Col = 6 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = pintDecimalSize : .TypeFloatMin = CDbl(strMin) : .TypeFloatMax = CDbl(strMax)
                .Col = 9 : .TypeMaxEditLen = 6
                .Col = 10 : .TypeMaxEditLen = 6
                If CmbInvType.Text = "NORMAL INVOICE" Or CmbInvType.Text = "JOBWORK INVOICE" Or CmbInvType.Text = "EXPORT INVOICE" Then
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
                .Col = 11 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = 2 : .TypeFloatMin = CDbl("0.00") : .TypeFloatMax = CDbl("99999999999999.99")
                .Col = 14 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                .Col = 15 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                .Col = 16 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = pintDecimalSize : .TypeFloatMin = CDbl(strMin) : .TypeFloatMax = CDbl(strMax)
                If mblnBatchTrack = True And UCase(CmbInvType.Text) <> "REJECTION" Then
                    .Col = 21 : .ColHidden = False : .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton : .TypeButtonText = "Batch Details"
                End If
                If ((CmbInvType.Text = "NORMAL INVOICE" And CmbInvSubType.Text = "FINISHED GOODS") Or
                    (strInvType = "INV" And strInvSubType = "F")) And mblnCSM_Knockingoff_req Then
                    .Col = 23 : .ColHidden = False : .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton : .TypeButtonText = "Knocking Details"
                Else
                    .Col = 23 : .ColHidden = True
                End If

                If mblnRejTracking = True And mblnBatchTracking = True Then
                    If UCase(CmbInvType.Text) = "REJECTION" Or UCase(strInvType) = "REJ" Then
                        .Col = 22 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                        .Col = 24 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                        .Col = 25 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton : .TypeButtonText = "Doc Detail"
                        .Col = 16 : .Col2 = 16 : .BlockMode = True : .ColHidden = False : .BlockMode = False
                    Else
                        .Col = 25 : .ColHidden = True
                    End If
                Else
                    .Col = 25 : .ColHidden = True
                End If

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
        strupSalechallanUpload = ""
        strupSalechallan = "Delete SalesChallan_Dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  Doc_No =" & Trim(txtChallanNo.Text)
        strupSalechallan = strupSalechallan & " and Location_Code ='" & Trim(txtLocationCode.Text) & "'"
        strupSaleDtl = "Delete Sales_Dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  Doc_No =" & Trim(txtChallanNo.Text)
        strupSaleDtl = strupSaleDtl & " and Location_Code ='" & Trim(txtLocationCode.Text) & "'"
        strupSalechallanUpload = "Delete SalesChallan_Dtl_Upload WHERE UNIT_CODE='" + gstrUNITID + "' AND  Doc_No =" & Trim(txtChallanNo.Text)
        strupSalechallanUpload = strupSalechallanUpload & " and Location_Code ='" & Trim(txtLocationCode.Text) & "'"
        DeleteRecords = True
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Public Function CheckMeasurmentUnit(ByRef strItem As Object, ByRef strQuantity As Object, ByRef intRow As Short) As Boolean
        Dim strMeasure As String
        Dim rsMeasure As ClsResultSetDB
        On Error GoTo ErrHandler
        strMeasure = "select a.Decimal_allowed_flag from Measure_Mst a,Item_Mst b"
        strMeasure = strMeasure & " where A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" + gstrUNITID + "' AND b.cons_Measure_Code=a.Measure_Code and b.Item_Code = '" & strItem & "'"
        rsMeasure = New ClsResultSetDB
        rsMeasure.GetResult(strMeasure, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        If rsMeasure.GetValue("Decimal_allowed_flag") = False Then
            If System.Math.Round(strQuantity, 0) - Val(strQuantity) <> 0 Then
                Call ConfirmWindow(10455, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                CheckMeasurmentUnit = False
                Call SpChEntry.SetText(5, intRow, CShort(strQuantity))
                SpChEntry.Col = 5
                SpChEntry.Row = SpChEntry.ActiveRow
                SpChEntry.Focus()
                rsMeasure.ResultSetClose()
                Exit Function
            Else
                CheckMeasurmentUnit = True
            End If
        Else
            CheckMeasurmentUnit = True
        End If
        rsMeasure.ResultSetClose()
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Public Function ParentQty(ByRef pstrItemCode As String, ByRef pstrfinished As Object) As Double
        On Error GoTo ErrHandler
        Dim strParentQty As String
        Dim rsParentQty As ClsResultSetDB
        strParentQty = "select sum(required_qty + waste_Qty) as TotalQty from Bom_Mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  finished_Product_code ='"
        strParentQty = strParentQty & pstrfinished & "' and rawMaterial_Code ='" & pstrItemCode & "'"
        rsParentQty = New ClsResultSetDB
        rsParentQty.GetResult(strParentQty, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        ParentQty = rsParentQty.GetValue("TotalQty")
        rsParentQty.ResultSetClose()
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
                rsSalesConf.GetResult("Select Stock_Location from SaleConf (nolock) WHERE UNIT_CODE='" + gstrUNITID + "' AND  Description ='" & Trim(pstrInvType) & "' and Sub_type_Description ='" & Trim(pstrInvSubtype) & "' AND Location_Code='" & Trim(txtLocationCode.Text) & "' and (fin_start_date <= getdate() and fin_end_date >= getdate())")
            Case "TYPE"
                rsSalesConf.GetResult("Select Stock_Location from SaleConf (nolock) WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_type ='" & Trim(pstrInvType) & "' and Sub_type ='" & Trim(pstrInvSubtype) & "' AND Location_Code='" & Trim(txtLocationCode.Text) & "' and (fin_start_date <= getdate() and fin_end_date >= getdate())")
        End Select
        If rsSalesConf.GetNoRows > 0 Then
            StockLocation = rsSalesConf.GetValue("Stock_Location")
        End If
        rsSalesConf.ResultSetClose()
        StockLocationSalesConf = StockLocation
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    '    Private Function GetServerDate() As Date
    '        Dim objServerDate As ClsResultSetDB 'Class Object
    '        Dim strsql As String 'Stores the SQL statement
    '        On Error GoTo ErrHandler
    '        'Build the SQL statement
    '        strsql = "SELECT CONVERT(datetime,getdate(),103)"
    '        'Creating the instance
    '        objServerDate = New ClsResultSetDB
    '        With objServerDate
    '            'Open the recordset
    '            Call .GetResult(strsql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
    '            'If we have a record, then getting the financial year else exiting
    '            If .GetNoRows <= 0 Then Exit Function
    '            'Getting the date
    '            'ServerDate = DateValue(Format(.GetValueByNo(0), gstrDateFormat))
    '            GetServerDate = CDate(VB6.Format(DateValue(.GetValueByNo(0)), gstrDateFormat))
    '            'Closing the recordset
    '            .ResultSetClose()
    '        End With
    '        'Releasing the object
    '        objServerDate = Nothing
    '        Exit Function
    'ErrHandler:  'The Error Handling Code Starts here
    '        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
    '        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    '    End Function
    Public Function ExploreBom(ByRef pstrItemCode As String, ByRef pstrFinishedQty As Object, ByRef pstrSPCurrentRow As Object, ByRef pstrFinishedProduct As String) As Boolean
        '*****************************************************
        'Created By     -  Nisha
        'Description    -  to get the values of required items in Sub assambly bom
        'input Variable -  Item Code to Found, reqquantity of Finished Product,row in spread
        '*****************************************************
        Dim strBomMstRaw As String
        Dim rsBomMstRaw As ClsResultSetDB
        Dim rsCustAnnexDtl As ClsResultSetDB
        Dim rsVandorBom As ClsResultSetDB
        Dim rsItemMst As ClsResultSetDB
        Dim intBomMaxRaw As Short
        Dim intCurrentRaw As Short
        Dim dblTotalReqQty As Double
        'Dim strProcessType As String
        Dim strCustAnnexDtl As String
        Dim strRGPQuote As String
        On Error GoTo ErrHandler
        strBomMstRaw = "Select RawMaterial_Code,Required_qty + Waste_qty "
        strBomMstRaw = strBomMstRaw & " As TotalReqQty,Process_Type from Bom_Mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  "
        strBomMstRaw = strBomMstRaw & " item_Code ='" & strBomItem
        strBomMstRaw = strBomMstRaw & "'and finished_product_code ='"
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
                'strProcessType = rsBomMstRaw.GetValue("Process_Type")
                'String for CustAnnex_dtl
                strCustAnnexDtl = "Select Item_Code,Balance_qty = sum(Balance_qty) from CustAnnex_hdr WHERE UNIT_CODE='" + gstrUNITID + "' AND  Customer_code ='"
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
                rsCustAnnexDtl = New ClsResultSetDB
                rsCustAnnexDtl.GetResult(strCustAnnexDtl, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                If rsCustAnnexDtl.GetNoRows >= 1 Then 'if item Found in CustAnnex then replace that item from Parant string
                    rsVandorBom = New ClsResultSetDB
                    rsVandorBom.GetResult("Select RawMaterial_Code from Vendor_bom WHERE UNIT_CODE='" + gstrUNITID + "' AND  Finish_Product_code = '" & pstrFinishedProduct & "'and RawMaterial_code = '" & strBomItem & "' and Vendor_code = '" & txtCustCode.Text & "'")
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
                                    MsgBox("Customer Supplied Materail for Item " & arrItem(inti) & " is " & arrQty(inti) & " .", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, "empower")
                                    SpChEntry.Row = pstrSPCurrentRow
                                    SpChEntry.Col = 5
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
                                MsgBox("Customer Supplied Materail for Item " & arrItem(inti) & " is " & arrQty(inti) & " .", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, "empower")
                                SpChEntry.Row = pstrSPCurrentRow
                                SpChEntry.Col = 5
                                SpChEntry.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                ExploreBom = False
                                Exit Function
                            Else
                                ExploreBom = True
                            End If
                        End If
                    End If
                    rsVandorBom.ResultSetClose()
                Else
                    'If strProcessType = "I" Then
                    rsVandorBom = New ClsResultSetDB
                    rsVandorBom.GetResult("Select RawMaterial_Code from Vendor_bom WHERE UNIT_CODE='" + gstrUNITID + "' AND  Finish_Product_code = '" & pstrFinishedProduct & "'and RawMaterial_code = '" & strBomItem & "' and vendor_code = '" & txtCustCode.Text & "'")
                    If rsVandorBom.GetNoRows > 0 Then
                        MsgBox("Item " & strBomItem & " is not supplied.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, "empower")
                        SpChEntry.Row = pstrSPCurrentRow
                        SpChEntry.Col = 5
                        SpChEntry.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                        rsVandorBom.ResultSetClose()
                        ExploreBom = False
                        Exit Function
                    Else 'if not of Process type I then again Explore
                        rsItemMst = New ClsResultSetDB
                        rsItemMst.GetResult("Select Item_Main_grp from Item_Mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  Item_code = '" & strBomItem & "'")
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
                rsCustAnnexDtl.ResultSetClose()
                rsBomMstRaw.MoveNext()
            Next
        Else
            MsgBox("No BOM Defind for Item (" & pstrItemCode & ") defined in challan", MsgBoxStyle.Information, "empower")
            ExploreBom = False
            Exit Function
        End If
        rsBomMstRaw.ResultSetClose()
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
        'Dim strProcessType As String
        Dim intBomMaxItem As Short
        Dim rsBomMst As ClsResultSetDB
        Dim rsCustAnnexDtl As ClsResultSetDB
        Dim rsVandorBom As ClsResultSetDB
        Dim rsItemMst As ClsResultSetDB
        Dim dblTotalReqQty As Double
        On Error GoTo ErrHandler
        BomCheck = False
        intSpreadRow = SpChEntry.MaxRows
        inti = 0
        Dim intArrCount As Short
        Dim blnItemFoundinArray As Boolean ' to be used to check if item already exist in Array arrItem where we are storing all item we found in Cust annex
        If SpChEntry.MaxRows >= 1 Then
            For intSpCurrentRow = 1 To intSpreadRow
                With SpChEntry
                    VarFinishedItem = Nothing
                    Call .GetText(1, intSpCurrentRow, VarFinishedItem)
                    VarFinishedQty = Nothing
                    Call .GetText(5, intSpCurrentRow, VarFinishedQty)
                End With
                strBomMst = "Select RawMaterial_Code,Process_type,Required_qty + Waste_qty "
                strBomMst = strBomMst & " As TotalReqQty"
                strBomMst = strBomMst & " from Bom_Mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  Finished_Product_code ='"
                strBomMst = strBomMst & VarFinishedItem & "' Order By Bom_Level"
                rsBomMst = New ClsResultSetDB
                rsBomMst.GetResult(strBomMst, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                intBomMaxItem = rsBomMst.GetNoRows
                rsBomMst.MoveFirst()
                If intBomMaxItem > 0 Then ' Item Found in Bom_Mst
                    rsVandorBom = New ClsResultSetDB
                    rsVandorBom.GetResult("Select RawMaterial_Code from Vendor_bom WHERE UNIT_CODE='" + gstrUNITID + "' AND  Finish_Product_code = '" & VarFinishedItem & "' and Vendor_code = '" & txtCustCode.Text & "'")
                    If rsVandorBom.GetNoRows > 0 Then
                        rsVandorBom.ResultSetClose()
                        For intCurrentItem = 1 To intBomMaxItem
                            strBomItem = ""
                            strBomItem = rsBomMst.GetValue("RawMaterial_Code")
                            strCustAnnexDtl = "Select Item_Code,Balance_qty = sum(Balance_qty) from CustAnnex_hdr WHERE UNIT_CODE='" + gstrUNITID + "' AND  Customer_code ='"
                            strCustAnnexDtl = strCustAnnexDtl & Trim(txtCustCode.Text) & "'"
                            If blnFIFO = False Then
                                strCustAnnexDtl = strCustAnnexDtl & " and ref57f4_no in ("
                                '***
                                strRgpsWithQuots = Replace(mstrRGP, "§", "','", 1)
                                strRgpsWithQuots = "'" & strRgpsWithQuots & "'"
                                '***
                                strCustAnnexDtl = strCustAnnexDtl & Trim(strRgpsWithQuots) & ") "
                            End If
                            '****
                            strCustAnnexDtl = strCustAnnexDtl & " and getdate() <= "
                            strCustAnnexDtl = strCustAnnexDtl & " DateAdd(d, 180, ref57f4_date)"
                            strCustAnnexDtl = strCustAnnexDtl & " and Item_code ='" & strBomItem & "' group By Item_code"
                            rsCustAnnexDtl = New ClsResultSetDB
                            rsCustAnnexDtl.GetResult(strCustAnnexDtl)
                            If rsCustAnnexDtl.GetNoRows >= 1 Then 'if item Found in Cust Annex
                                rsVandorBom = New ClsResultSetDB
                                rsVandorBom.GetResult("Select RawMaterial_Code from Vendor_bom WHERE UNIT_CODE='" + gstrUNITID + "' AND  Finish_Product_code = '" & VarFinishedItem & "'and RawMaterial_code = '" & strBomItem & "' and Vendor_code = '" & txtCustCode.Text & "'")
                                If rsVandorBom.GetNoRows > 0 Then
                                    rsVandorBom.ResultSetClose()
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
                                                    MsgBox("Customer Supplied Materail for Item " & arrItem(inti) & "is" & arrQty(inti) & ".", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, "empower")
                                                    SpChEntry.Row = intSpCurrentRow
                                                    SpChEntry.Col = 5
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
                                                MsgBox("Customer Supplied Material for Item " & arrItem(inti) & "is" & arrQty(inti) & ".", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, "empower")
                                                SpChEntry.Row = intSpCurrentRow
                                                SpChEntry.Col = 5
                                                SpChEntry.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                                BomCheck = False
                                                Exit Function
                                            End If
                                        End If
                                    Else ' if inti=0 then to add values
                                        arrItem(inti) = rsCustAnnexDtl.GetValue("Item_code")
                                        arrQty(inti) = rsCustAnnexDtl.GetValue("Balance_qty")
                                        arrReqQty(inti) = dblTotalReqQty * VarFinishedQty
                                        If arrQty(inti) < arrReqQty(inti) Then 'Again Same Check
                                            MsgBox("Customer Supplied Material for Item " & arrItem(inti) & "is" & arrQty(inti) & ".", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, "empower")
                                            SpChEntry.Row = intSpCurrentRow
                                            SpChEntry.Col = 5
                                            SpChEntry.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                            BomCheck = False
                                            Exit Function
                                        End If
                                    End If
                                End If
                                rsVandorBom.ResultSetClose()
                            Else ' if Item Not Found in Cust Annex
                                rsVandorBom = New ClsResultSetDB
                                rsVandorBom.GetResult("Select RawMaterial_Code from Vendor_bom WHERE UNIT_CODE='" + gstrUNITID + "' AND  Finish_Product_code = '" & VarFinishedItem & "'and RawMaterial_code = '" & strBomItem & "' and Vendor_code = '" & txtCustCode.Text & "'")
                                If rsVandorBom.GetNoRows > 0 Then
                                    rsVandorBom.ResultSetClose()
                                    'If strProcessType = "I" Then ' If That Item is has Process Type I in Bom then
                                    MsgBox("Item " & strBomItem & " is not supplied by Customer.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, "empower")
                                    SpChEntry.Row = intSpCurrentRow
                                    SpChEntry.Col = 5
                                    SpChEntry.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                    BomCheck = False
                                    Exit Function
                                Else ' if it'Process type is not I then Explore it Again in BOM_Mst
                                    rsItemMst = New ClsResultSetDB
                                    rsItemMst.GetResult("Select Item_Main_grp from Item_Mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  Item_code = '" & strBomItem & "'")
                                    If (UCase(rsItemMst.GetValue("Item_Main_grp")) = "R") Or (UCase(rsItemMst.GetValue("Item_Main_grp")) = "C") Then
                                        BomCheck = True
                                    Else
                                        VarFinishedQty = VarFinishedQty * rsBomMst.GetValue("TotalReqQty")
                                        If ExploreBom(strBomItem, VarFinishedQty, intSpCurrentRow, CStr(VarFinishedItem)) = False Then
                                            rsVandorBom.ResultSetClose()
                                            BomCheck = False
                                            rsItemMst.ResultSetClose()
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
                        MsgBox("No Customer BOM Defind for Item (" & VarFinishedItem & ") defined in challan", MsgBoxStyle.Information, "empower")
                        BomCheck = False
                        rsVandorBom.ResultSetClose()
                        Exit Function
                    End If
                Else ' if no Item Found from Grid
                    MsgBox("No BOM Defind for Item (" & VarFinishedItem & ") defined in challan", MsgBoxStyle.Information, "empower")
                    BomCheck = False
                    rsBomMst.ResultSetClose()
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
        rscurrency.GetResult("Select Decimal_Place from Currency_Mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  Currency_code ='" & pstrCurrency & "'")
        ToGetDecimalPlaces = Val(rscurrency.GetValue("Decimal_Place"))
        rscurrency.ResultSetClose()
    End Function
    Public Function ToGetCurrencyType() As String
        Dim rsCustOrdHdr As ClsResultSetDB
        Dim strcustHdr As String
        On Error GoTo ErrHandler
        rsCustOrdHdr = New ClsResultSetDB
        strcustHdr = "Select Currency_Code from Cust_ord_hdr"
        strcustHdr = strcustHdr & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Account_Code='" & txtCustCode.Text & "' and Cust_Ref ='"
        strcustHdr = strcustHdr & mstrRefNo & "'and Amendment_No ='" & mstrAmmNo & "'"
        rsCustOrdHdr.GetResult(strcustHdr)
        If rsCustOrdHdr.GetNoRows > 0 Then
            rsCustOrdHdr.MoveFirst()
            ToGetCurrencyType = rsCustOrdHdr.GetValue("Currency_Code")
        End If
        rsCustOrdHdr.ResultSetClose()
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Function SelectDataFromCustOrd_Dtl(ByRef pstrCustCode As String, ByRef pstrInvType As String) As Boolean
        On Error GoTo ErrHandler
        Dim strSelectSql As String 'Declared To Make Select Query
        Dim rsCustOrdDtl As ClsResultSetDB
        SelectDataFromCustOrd_Dtl = False
        If UCase(pstrInvType) = "JOBWORK INVOICE" Then
            strSelectSql = "Select b.Item_Code,b.Cust_DrgNo,b.Amendment_No,b.Cust_Ref from Cust_Ord_hdr a,Cust_Ord_Dtl b"
            strSelectSql = strSelectSql & " where A.UNIT_CODE =B.UNIT_CODE AND A.UNIT_CODE='" + gstrUNITID + "' AND b.Account_Code='" & Trim(pstrCustCode) & "' and b.Active_flag ='A' and "
            strSelectSql = strSelectSql & " a.Account_Code = b.Account_Code and a.Cust_ref =b.Cust_ref and "
            strSelectSql = strSelectSql & " a.Amendment_No = b.amendment_No AND a.Authorized_Flag = 1 and a.PO_type in ('J') "
            strSelectSql = strSelectSql & " and a.Valid_date >='" & VB6.Format(GetServerDate(), "dd/MMM/yyyy") & "' and effect_Date <='" & VB6.Format(GetServerDate(), "dd/MMM/yyyy") & "'"
            strSelectSql = strSelectSql & " AND b.Cust_Ref = '" & Trim(txtRefNo.Text) & "'"
            strSelectSql = strSelectSql & " order by b.Cust_Ref,b.Amendment_No,b.Cust_DrgNo,b.Item_Code "
        ElseIf UCase(pstrInvType) = "EXPORT INVOICE" Then
            strSelectSql = "Select b.Item_Code,b.Cust_DrgNo,b.Amendment_No,b.Cust_Ref from Cust_Ord_hdr a,Cust_Ord_Dtl b"
            strSelectSql = strSelectSql & " where A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" + gstrUNITID + "' AND b.Account_Code='" & Trim(pstrCustCode) & "' and b.Active_flag ='A' and "
            strSelectSql = strSelectSql & " a.Account_Code = b.Account_Code and a.Cust_ref =b.Cust_ref and "
            strSelectSql = strSelectSql & " a.Amendment_No = b.amendment_No AND a.Authorized_Flag = 1 and a.PO_type in ('E') "
            strSelectSql = strSelectSql & " and a.Valid_date >='" & VB6.Format(GetServerDate(), "dd/MMM/yyyy") & "' and effect_date <='" & VB6.Format(GetServerDate(), "dd/MMM/yyyy") & "'"
            strSelectSql = strSelectSql & " AND b.Cust_Ref = '" & Trim(txtRefNo.Text) & "'"
            strSelectSql = strSelectSql & " order by b.Cust_Ref,b.Amendment_No,b.Cust_DrgNo,b.Item_Code "
            '10869290
        ElseIf UCase(pstrInvType) = "SERVICE INVOICE" Then
            strSelectSql = "Select b.Item_Code,b.Cust_DrgNo,b.Amendment_No,b.Cust_Ref from Cust_Ord_hdr a,Cust_Ord_Dtl b"
            strSelectSql = strSelectSql & " where A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" + gstrUNITID + "' AND b.Account_Code='" & Trim(pstrCustCode) & "' and b.Active_flag ='A' and "
            strSelectSql = strSelectSql & " a.Account_Code = b.Account_Code and a.Cust_ref =b.Cust_ref and "
            strSelectSql = strSelectSql & " a.Amendment_No = b.amendment_No AND a.Authorized_Flag = 1 and a.PO_type in ('V') "
            strSelectSql = strSelectSql & " and a.Valid_date >='" & VB6.Format(GetServerDate(), "dd/MMM/yyyy") & "' and effect_date <='" & VB6.Format(GetServerDate(), "dd/MMM/yyyy") & "'"
            strSelectSql = strSelectSql & " AND b.Cust_Ref = '" & Trim(txtRefNo.Text) & "'"
            strSelectSql = strSelectSql & " order by b.Cust_Ref,b.Amendment_No,b.Cust_DrgNo,b.Item_Code "
        ElseIf UCase(pstrInvType) = "REJECTION" Then
            strSelectSql = "select a.Doc_No,a.Item_code,a.Rejected_Quantity from grn_Dtl a,grn_hdr b Where A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" + gstrUNITID + "' "
            strSelectSql = strSelectSql & " and a.Doc_type = b.Doc_type And a.Doc_No = b.Doc_No "
            strSelectSql = strSelectSql & " and a.From_Location = b.From_Location and a.From_Location ='01R1'"
            strSelectSql = strSelectSql & " and a.Rejected_quantity > 0  and b.Vendor_code = '" & pstrCustCode & "' AND A.Doc_No = " & txtRefNo.Text & "  AND ISNULL(b.GRN_Cancelled,0) = 0 order by a.Doc_No"
        Else
            strSelectSql = "Select b.Item_Code,b.Cust_DrgNo,b.Amendment_No,b.Cust_Ref from Cust_Ord_hdr a,Cust_Ord_Dtl b"
            strSelectSql = strSelectSql & " where A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" + gstrUNITID + "' AND  b.Account_Code='" & Trim(pstrCustCode) & "' and b.Active_flag ='A' and "
            strSelectSql = strSelectSql & " a.Account_Code = b.Account_Code and a.Cust_ref =b.Cust_ref and "
            strSelectSql = strSelectSql & " a.Amendment_No = b.amendment_No  AND a.Authorized_Flag = 1 and a.PO_type in ('O','S') "
            strSelectSql = strSelectSql & " and a.Valid_date >='" & VB6.Format(GetServerDate(), "dd/MMM/yyyy") & "' and effect_Date <= '" & VB6.Format(GetServerDate(), "dd/MMM/yyyy") & "'"
            strSelectSql = strSelectSql & " AND b.Cust_Ref = '" & Trim(txtRefNo.Text) & "'"
            strSelectSql = strSelectSql & " order by b.Cust_Ref,b.Amendment_No,b.Cust_DrgNo,b.Item_Code "
        End If
        rsCustOrdDtl = New ClsResultSetDB
        rsCustOrdDtl.GetResult(strSelectSql) ', ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
        If rsCustOrdDtl.GetNoRows > 0 Then '          'if record found
            SelectDataFromCustOrd_Dtl = True 'Return TRUE
        End If
        rsCustOrdDtl.ResultSetClose()
        If CmbInvType.Text.ToUpper = "EXPORT INVOICE" Then
            mstrexportsotype = Find_Value("SELECT EXPORTSOTYPE FROM CUST_ORD_HDR WHERE UNIT_CODE='" + gstrUNITID + "' AND ACCOUNT_CODE='" & txtCustCode.Text & "' AND cust_ref='" & txtRefNo.Text & "' and amendment_no='" & txtAmendNo.Text & "'")
            lblexportsodetails.Text = mstrexportsotype
        Else
            lblexportsodetails.Text = ""
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
        '1st Check if Any Blank Amendment no for Ref No. Exists
        If SelectDataFromTable("Active_Flag", "Cust_ORD_HDR", " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Account_Code ='" & Trim(txtCustCode.Text) & "' AND Cust_Ref = '" & txtRefNo.Text & "' AND Amendment_No = ''") = "O" Then
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
        MaxInvoiceDate = Find_Value("select max(Invoice_date) as invoice_date from saleschallan_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  Bill_Flag=1")
        CurrentDate = GetServerDate()
        If Len(MaxInvoiceDate) = 0 Then
            MaxInvoiceDate = getDateForDB(GetServerDate())
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
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Public Function AddDataTolstRGPs() As Boolean
        Dim rsCustAnnex As ClsResultSetDB
        Dim intLoopCounter As Short
        Dim intMaxCounter As Short
        On Error GoTo ErrHandler
        With lvwRGPs
            .GridLines = True : .Items.Clear() : .Columns.Clear()
            Call .Columns.Insert(0, "", "RGP No(s)", CInt(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(lvwRGPs.Width) / 2)))
            Call .Columns.Insert(1, "", "RGP Date", CInt(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(lvwRGPs.Width) / 2 - 700)))
            rsCustAnnex = New ClsResultSetDB
            rsCustAnnex.GetResult("select distinct ref57f4_No,ref57f4_date from custannex_HDR WHERE UNIT_CODE='" + gstrUNITID + "' AND  customer_Code='" & Trim(txtCustCode.Text) & "' and getdate() < dateadd(d,180,ref57f4_Date) order by ref57f4_Date")
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
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Public Function CheckExchangeRate() As Boolean
        On Error GoTo ErrHandler
        If Trim(lblExchangeRateValue.Text) = "" Then
            MsgBox("Please Define Exchange Rate For this Month in Exchange Master", MsgBoxStyle.Information, "empower")
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
        Dim strsql As String
        Dim varItemCode As Object
        Dim varItemQty As Object
        Dim VarDelete As Object
        Dim dblRejQty As Double
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        Dim varMaxallowedQty As Object

        On Error GoTo ErrHandler
        intMaxLoop = SpChEntry.MaxRows
        ItemQtyCaseRejGrin = False
        For intLoopCounter = 1 To intMaxLoop
            varItemCode = Nothing
            Call SpChEntry.GetText(1, intLoopCounter, varItemCode)
            VarDelete = Nothing
            Call SpChEntry.GetText(15, intLoopCounter, VarDelete)
            varItemQty = Nothing
            Call SpChEntry.GetText(5, intLoopCounter, varItemQty)
            varMaxallowedQty = Nothing
            Call SpChEntry.GetText(22, intLoopCounter, varMaxallowedQty)
            If VarDelete <> "D" Then
                If mblnRejTracking = True Then
                    If Val(varMaxallowedQty) < varItemQty Then
                        MsgBox("Quantity Allowed For Item Code " & varItemCode & " is " & varMaxallowedQty & ", Cannot Enter More then This.", MsgBoxStyle.OkOnly, ResolveResString(100))
                        ItemQtyCaseRejGrin = False
                        Exit Function
                    Else
                        ItemQtyCaseRejGrin = True
                    End If
                Else
                    strsql = "select a.Doc_No,a.Item_code, MaxAllowedQty = ((a.Rejected_Quantity + a.excess_po_quantity) - (isnull(a.Despatch_Quantity,0) + isnull(a.Inspected_Quantity,0) + isnull(a.RGP_Quantity,0)))from grn_Dtl a,grn_hdr b Where A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" + gstrUNITID + "' AND  "
                    strsql = strsql & "a.Doc_type = b.Doc_type And a.Doc_No = b.Doc_No and "
                    strsql = strsql & "a.From_Location = b.From_Location and a.From_Location ='01R1'"
                    strsql = strsql & " and a.Rejected_quantity > 0 and b.Vendor_code = '" & txtCustCode.Text
                    strsql = strsql & "' and a.Doc_No = " & CDbl(txtRefNo.Text) & " and a.Item_code = '" & varItemCode & "' AND ISNULL(b.GRN_Cancelled,0) = 0"
                    rsGrnDtl = New ClsResultSetDB
                    rsGrnDtl.GetResult(strsql)
                    If rsGrnDtl.GetNoRows > 0 Then
                        dblRejQty = rsGrnDtl.GetValue("MaxAllowedQty")
                        If varItemQty > dblRejQty Then
                            MsgBox("Quantity Allowed For This Item is " & dblRejQty & ", cannot Enter More then This.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100))
                            SpChEntry.Row = intLoopCounter : SpChEntry.Col = 22 : SpChEntry.Action = FPSpreadADO.ActionConstants.ActionActiveCell : SpChEntry.Focus()
                            ItemQtyCaseRejGrin = False
                            Exit Function
                        Else
                            ItemQtyCaseRejGrin = True
                        End If
                    Else
                        MsgBox("No Item -" & varItemCode & " available in GRIN No - " & txtRefNo.Text & " Having Rejected Quantity > 0 ")
                        ItemQtyCaseRejGrin = False
                        rsGrnDtl.ResultSetClose()
                        Exit Function
                    End If
                    rsGrnDtl.ResultSetClose()
                End If
            Else
                ItemQtyCaseRejGrin = True
            End If
        Next
        Exit Function
        '    If VarDelete <> "D" Then
        '        strsql = "select a.Doc_No,a.Item_code, MaxAllowedQty = (a.Rejected_Quantity - (isnull(a.Despatch_Quantity,0) + isnull(a.Inspected_Quantity,0) + isnull(a.RGP_Quantity,0)))from grn_Dtl a,grn_hdr b Where A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" + gstrUNITID + "' "
        '        strsql = strsql & " and a.Doc_type = b.Doc_type And a.Doc_No = b.Doc_No and "
        '        strsql = strsql & "a.From_Location = b.From_Location and a.From_Location ='01R1'"
        '        strsql = strsql & "and a.Rejected_quantity > 0 and b.Vendor_code = '" & txtCustCode.Text
        '        strsql = strsql & "' and a.Doc_No = " & CDbl(txtRefNo.Text) & " and a.Item_code = '" & varItemCode & "' AND ISNULL(b.GRN_Cancelled,0) = 0"
        '        rsGrnDtl = New ClsResultSetDB
        '        rsGrnDtl.GetResult(strsql)
        '        If rsGrnDtl.GetNoRows > 0 Then
        '            dblRejQty = rsGrnDtl.GetValue("MaxAllowedQty")
        '            If varItemQty > dblRejQty Then
        '                MsgBox("Quantity Allowed For This Item is " & dblRejQty & ", cannot Enter More then This.")
        '                SpChEntry.Row = intLoopCounter : SpChEntry.Col = 5 : SpChEntry.Action = FPSpreadADO.ActionConstants.ActionActiveCell : SpChEntry.Focus()
        '                ItemQtyCaseRejGrin = False
        '                rsGrnDtl.ResultSetClose()
        '                Exit Function
        '            Else
        '                ItemQtyCaseRejGrin = True
        '            End If
        '        Else
        '            MsgBox("No Item -" & varItemCode & " available in GRIN No - " & txtRefNo.Text & " Having Rejected Quantity >0 ")
        '            ItemQtyCaseRejGrin = False
        '            rsGrnDtl.ResultSetClose()
        '            Exit Function
        '        End If
        '        rsGrnDtl.ResultSetClose()
        '    Else
        '        ItemQtyCaseRejGrin = True
        '    End If
        'Next
        'Exit Function
ErrHandler:  'The Error Handling Code Starts here
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Public Function ScheduleCheckEditMode() As Boolean
        Dim strScheduleSql As String
        Dim varDrgNo As Object
        Dim strQuantity As Double
        Dim rsMktSchedule As ClsResultSetDB
        Dim rsMktSchedule1 As ClsResultSetDB
        Dim varItemQty As Object
        Dim VarDelete As Object
        Dim PresQty As Object
        Dim intRwCount As Short
        Dim varItemCode As Object
        Dim intLoopcount As Short
        Dim strMakeDate As String
        If ((UCase(mstrInvType) = "INV") And (UCase(mstrInvSubType) = "F") Or (UCase(mstrInvSubType) = "T")) Or (UCase(Trim(CmbInvType.Text)) = "JOBWORK INVOICE") Or (UCase(mstrInvType) = "EXP") Then
            strScheduleSql = "Select Quantity=Schedule_Quantity-isnull(Despatch_Qty,0),Cust_DrgNo,Item_Code from DailyMktSchedule WHERE UNIT_CODE='" + gstrUNITID + "' AND  Account_Code='" & Trim(txtCustCode.Text) & "' and "
            strScheduleSql = strScheduleSql & " datepart(yyyy,Trans_Date)='" & Year(ConvertToDate(Trim(lblDateDes.Text))) & "'"
            strScheduleSql = strScheduleSql & " and datepart(mm,Trans_Date)='" & Month(ConvertToDate(Trim(lblDateDes.Text))) & "'"
            strScheduleSql = strScheduleSql & " and datepart(dd,Trans_Date)='" & VB.Day(ConvertToDate(Trim(lblDateDes.Text))) & "'"
            strScheduleSql = strScheduleSql & " and Cust_DrgNo in(" & Trim(mstrItemCode) & ") and Status =1 and Schedule_Flag =1"
            rsMktSchedule.GetResult(strScheduleSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If rsMktSchedule.GetNoRows > 0 Then 'If Record Found
                rsMktSchedule.MoveFirst()
                For intRwCount = 1 To Me.SpChEntry.MaxRows
                    'Select Quantity From The Spread
                    varItemQty = Nothing
                    Call Me.SpChEntry.GetText(5, intRwCount, varItemQty)
                    VarDelete = Nothing
                    Call Me.SpChEntry.GetText(15, intRwCount, VarDelete) ''Column Changed By Tapan
                    strQuantity = rsMktSchedule.GetValue("Quantity")
                    'If Quantity Entered Is Greater Then Schedule Quantity
                    If UCase(VarDelete) <> "D" Then
                        If (Val(varItemQty) - Val(mdblPrevQty(intLoopcount))) > Val(CStr(strQuantity)) Then
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
                            For intLoopcount = 1 To SpChEntry.MaxRows
                                varDrgNo = Nothing
                                Call Me.SpChEntry.GetText(2, intLoopcount, varDrgNo)
                                varItemCode = Nothing
                                Call Me.SpChEntry.GetText(1, intLoopcount, varItemCode)
                                PresQty = Nothing
                                Call Me.SpChEntry.GetText(5, intLoopcount, PresQty)
                                strScheduleSql = "select Despatch_qty  = isnull(Despatch_Qty,0) - (" & Val(mdblPrevQty(intLoopcount - 1)) - Val(PresQty) & "),SChedule_Quantity from DailyMktSchedule "
                                strScheduleSql = strScheduleSql & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Account_Code='" & Trim(txtCustCode.Text) & "' and "
                                strScheduleSql = strScheduleSql & " datepart(yyyy,Trans_Date)='" & Year(ConvertToDate(Trim(lblDateDes.Text))) & "'"
                                strScheduleSql = strScheduleSql & " and datepart(mm,Trans_Date)='" & Month(ConvertToDate(Trim(lblDateDes.Text))) & "'"
                                strScheduleSql = strScheduleSql & " and datepart(dd,Trans_Date)='" & VB.Day(ConvertToDate(Trim(lblDateDes.Text))) & "'"
                                strScheduleSql = strScheduleSql & " and Cust_DrgNo ='" & Trim(varDrgNo) & "'and Item_code = '" & varItemCode & "' and Status =1 and Schedule_flag =1" & vbCrLf
                                rsMktSchedule1 = New ClsResultSetDB
                                rsMktSchedule1.GetResult(strScheduleSql)
                                mstrUpdDispatchSql = Trim(mstrUpdDispatchSql) & "Update DailyMktSchedule set Despatch_qty ="
                                mstrUpdDispatchSql = mstrUpdDispatchSql & "isnull(Despatch_Qty,0) - (" & Val(mdblPrevQty(intLoopcount - 1)) - Val(PresQty) & ")"
                                If Val(rsMktSchedule1.GetValue("Despatch_Qty")) = Val(rsMktSchedule1.GetValue("Schedule_Quantity")) Then
                                    mstrUpdDispatchSql = mstrUpdDispatchSql & ", Schedule_Flag = 0 "
                                End If
                                rsMktSchedule1.ResultSetClose()
                                mstrUpdDispatchSql = mstrUpdDispatchSql & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Account_Code='" & Trim(txtCustCode.Text) & "' and "
                                mstrUpdDispatchSql = mstrUpdDispatchSql & " datepart(yyyy,Trans_Date)='" & Year(ConvertToDate(Trim(lblDateDes.Text))) & "'"
                                mstrUpdDispatchSql = mstrUpdDispatchSql & " and datepart(mm,Trans_Date)='" & Month(ConvertToDate(Trim(lblDateDes.Text))) & "'"
                                mstrUpdDispatchSql = mstrUpdDispatchSql & " and datepart(dd,Trans_Date)='" & VB.Day(ConvertToDate(Trim(lblDateDes.Text))) & "'"
                                mstrUpdDispatchSql = mstrUpdDispatchSql & " and Cust_DrgNo ='" & Trim(varDrgNo) & "'and Item_code = '" & varItemCode & "' and Status =1 and Schedule_flag =1" & vbCrLf
                            Next
                        End If
                    End If
                    rsMktSchedule.MoveNext()
                Next intRwCount
                rsMktSchedule.ResultSetClose()
            ElseIf rsMktSchedule.GetNoRows = 0 Then
                rsMktSchedule.ResultSetClose()
                If Val(CStr(Month(ConvertToDate(lblDateDes.Text)))) < 10 Then
                    strMakeDate = Year(ConvertToDate(lblDateDes.Text)) & "0" & Month(ConvertToDate(lblDateDes.Text))
                Else
                    strMakeDate = Year(ConvertToDate(lblDateDes.Text)) & Month(ConvertToDate(lblDateDes.Text))
                End If
                strScheduleSql = "Select Quantity=Schedule_Qty-isnull(Despatch_Qty,0) from MonthlyMktSchedule WHERE UNIT_CODE='" + gstrUNITID + "' AND  Account_Code='" & Trim(txtCustCode.Text) & "' and "
                strScheduleSql = strScheduleSql & " Year_Month=" & Val(Trim(strMakeDate)) & ""
                strScheduleSql = strScheduleSql & " and Cust_DrgNo in(" & Trim(mstrItemCode) & ") and status =1 and Schedule_flag =1"
                rsMktSchedule = New ClsResultSetDB
                rsMktSchedule.GetResult(strScheduleSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                If rsMktSchedule.GetNoRows > 0 Then
                    rsMktSchedule.MoveFirst()
                    For intRwCount = 1 To Me.SpChEntry.MaxRows
                        Select Case CmdGrpChEnt.Mode
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
                        Call Me.SpChEntry.GetText(12, intRwCount, VarDelete)
                        If UCase(VarDelete) <> "D" Then
                            If Val(varItemQty) > Val(CStr(strQuantity)) Then
                                ScheduleCheckEditMode = False
                                MsgBox("Quantity should not be Greater then Schedule Quantity " & strQuantity, MsgBoxStyle.Information, "empower")
                                With Me.SpChEntry
                                    .Row = intRwCount : .Col = 5 : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                                End With
                                Exit Function
                            Else
                                ScheduleCheckEditMode = False
                                mstrUpdDispatchSql = ""
                                For intLoopcount = 1 To SpChEntry.MaxRows
                                    varDrgNo = Nothing
                                    Call Me.SpChEntry.GetText(2, intLoopcount, varDrgNo)
                                    varItemCode = Nothing
                                    Call Me.SpChEntry.GetText(1, intLoopcount, varItemCode)
                                    PresQty = Nothing
                                    Call Me.SpChEntry.GetText(5, intLoopcount, PresQty)
                                    '**** To Check schedule Quantity
                                    strScheduleSql = "Select Despatch_qty = "
                                    strScheduleSql = strScheduleSql & "isnull(Despatch_Qty,0) - (" & Val(mdblPrevQty(intLoopcount - 1)) - Val(PresQty) & "),Schedule_Qty"
                                    strScheduleSql = strScheduleSql & " From MonthlyMktSchedule "
                                    strScheduleSql = strScheduleSql & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Account_Code='" & Trim(txtCustCode.Text) & "' and "
                                    strScheduleSql = strScheduleSql & " Year_Month=" & Val(Trim(strMakeDate)) & ""
                                    strScheduleSql = strScheduleSql & " and Cust_DrgNo ='" & Trim(varDrgNo) & "'and Item_code = '" & varItemCode & "' and status =1 and Schedule_flag =1" & vbCrLf
                                    rsMktSchedule1 = New ClsResultSetDB
                                    rsMktSchedule1.GetResult(strScheduleSql)
                                    '********
                                    mstrUpdDispatchSql = Trim(mstrUpdDispatchSql) & "Update MonthlyMktSchedule set Despatch_qty ="
                                    mstrUpdDispatchSql = mstrUpdDispatchSql & "isnull(Despatch_Qty,0) - (" & Val(mdblPrevQty(intLoopcount - 1)) - Val(PresQty) & ")"
                                    If rsMktSchedule1.GetValue("Despatch_Qty") = rsMktSchedule1.GetValue("Schedule_Qty") Then
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & ", Schedule_Flag = 0 "
                                    End If
                                    rsMktSchedule1.ResultSetClose()
                                    mstrUpdDispatchSql = mstrUpdDispatchSql & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Account_Code='" & Trim(txtCustCode.Text) & "' and "
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
                    rsMktSchedule.ResultSetClose()
                    Exit Function
                End If
                rsMktSchedule.ResultSetClose()
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
        pstrDate = getDateForDB(pstrDate)
        If IsCustomer = True Then
            strTableSql = "SELECT CExch_MultiFactor FROM Gen_CurExchMaster WHERE UNIT_CODE='" + gstrUNITID + "' AND  CExch_CurrencyTo='" & Trim(pstrCurrencyCode) & "' AND CExch_InOut=1 AND '" & Trim(pstrDate) & "' BETWEEN CExch_DateFrom AND CExch_DateTo"
        Else
            strTableSql = "SELECT CExch_MultiFactor FROM Gen_CurExchMaster WHERE UNIT_CODE='" + gstrUNITID + "' AND  CExch_CurrencyTo='" & Trim(pstrCurrencyCode) & "' AND CExch_InOut=0 AND '" & Trim(pstrDate) & "' BETWEEN CExch_DateFrom AND CExch_DateTo"
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
        '---------------------------------------------------------------------------------------

        On Error GoTo ErrHandler

        Dim ldblTotalBasicValue As Double
        Dim ldblTotalAccessibleValue As Double
        Dim lintLoopCounter As Short
        Dim ldblTempAccessibleVal As Double
        Dim ldblTotalExciseValue As Double
        ''Changed at MSSLED-------------
        Dim ldblTotalCVDValue As Double
        Dim ldblTotalSADValue As Double
        ''Changed at MSSLED-------------
        Dim ldblTotalSaleTaxAmount As Double
        Dim ldblTotalSurchargeTaxAmount As Double
        Dim ldblNetInsurenceValue As Double
        Dim ldblTotalInvoiceValue As Double
        Dim ldblTotalOthersValues As Double
        Dim rsParameterData As ClsResultSetDB
        Dim strParamQuery As String
        Dim blnCSIEX_Inc As Boolean
        Dim strSalesChallan As String
        Dim updateSalesChallan As String
        Dim strSalesDtl As String
        Dim strSalesDtlDelete As String
        Dim rsCustItemMst As ClsResultSetDB
        Dim rsSaleConf As ClsResultSetDB
        Dim rsItemMst As ClsResultSetDB
        Dim lintItemQuantity As Double
        Dim lintItemrate As Double
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
        Dim ldblExciseValueForSaleTax As Double
        Dim ldbltotalpkgvalue As Double
        'changes done by rajesh
        Dim ldblcustsuppexcisevalue As Double
        Dim lblTotalCustMtrlAmount As Double
        Dim varItemQuantity As Object
        Dim varCustsuppmtrl As Object
        Dim varAmor As Object
        Dim ldblTotalAmorValue As Double
        Dim ldblExciseValueForInvoice As Double
        'Code Added by Arshad on 08/07/2004 to add ECSS Tax
        Dim blnECSSTax As Boolean
        Dim intECSSRoundOffDecimal As Short
        Dim ldblTotalECSSTaxAmount As Double
        'Code Ends here
        Dim varAdditionalExciseType As Object
        Dim varAdditionalExcisePer As Object
        Dim dblAdditionalExciseAmount As Double
        Dim ldblTotalSECSSTaxAmount As Double ''Ashutosh
        Dim strbatchdtl As String
        Dim Intcounter As Short
        Dim strBatchNo As String
        Dim dblBatchQty As Double
        Dim strToLocation As String
        Dim strInsertUpdateASNdtl As String
        Dim strSQLREJINV_DTL As String
        Dim dblTCSTaxAmount As Double
        Dim blnTCSTax As Boolean
        Dim intTCSRoundOffDecimal As Short
        Dim rsExternalsalesorder As ClsResultSetDB
        Dim dblAddVATamount As Double
        Dim blnInsIncSTax As Boolean
        Dim intSaleTaxRoundOffDecimal As Short
        Dim strSqlct2qry As String = ""
        Dim dblExcise_Amount As Double
        Dim strSql As String = ""
        Dim blnIsCt2 As Boolean = False
        Dim TmpRs As New ClsResultSetDB
        Dim dblTaxRate As Double
        Dim strModel As String = ""
        Dim ldblTotalServiceTax_Amount As Double
        Dim blnServiceTax_Roundoff As Boolean
        Dim intServiceTaxRoundoff_Decimal As Short
        Dim ldblTotalSBCTax_Amount As Double
        '11 mar 2016
        Dim dblItemToolCost As Double
        '11 mar 2016
        Dim strAtnno As String = ""
        Dim ldblTotalkkcTax_Amount As Double
        Dim intExcise_Roundoff_Decimal As Short
        'GST CHANGES
        Dim intGSTTAXroundoff_decimal As Short
        Dim blnGSTTAXroundoff As Boolean
        Dim STRHSNSACCODE As String
        Dim STRCGSTTAXTYPE As String
        Dim STRSGSTTXRT_TYPE As String
        Dim STRUTGSTTXRT_TYPE As String
        Dim STRIGSTTXRT_TYPE As String
        Dim STRCOMPENSATION_CESS As String
        Dim ldblTotalGSTTAXTES As Double
        Dim strexportsotype As String
        'GST CHANGES
        Dim strPalatteQuery As String = String.Empty
        Dim strpPCAScheduleinvoiceknockoff As String
        Dim strDYNoteNo As String = String.Empty
        Dim STRCONSCODE As String = String.Empty

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
        ldblcustsuppexcisevalue = 0
        lblTotalCustMtrlAmount = 0
        ldblTotalAmorValue = 0
        ldblExciseValueForInvoice = 0
        ldblTotalSECSSTaxAmount = 0
        dblAddVATamount = 0
        ldblTotalServiceTax_Amount = 0
        ldblTotalSBCTax_Amount = 0
        ldblTotalkkcTax_Amount = 0

        SaveData = True
        strToLocation = ReturnCustomerLocation()
        strParamQuery = "SELECT ServiceTax_Roundoff,ServiceTaxRoundoff_Decimal,InsExc_Excise,CustSupp_Inc,EOU_Flag,Basic_Roundoff,SalesTax_Roundoff,Excise_Roundoff,Excise_Roundoff_Decimal,SST_Roundoff,ECESS_Roundoff,ECESSRoundoff_Decimal,SalesTax_Onerupee_Roundoff,TCSTax_Roundoff , TCSTax_Roundoff_decimal , InsInc_SalesTax , SalesTax_Roundoff_decimal,GSTTAX_ROUNDOFF_DECIMAL ,GSTTAX_ROUNDOFF FROM Sales_Parameter WHERE UNIT_CODE='" + gstrUNITID + "'"
        rsParameterData = New ClsResultSetDB
        rsParameterData.GetResult(strParamQuery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsParameterData.GetNoRows > 0 Then
            blnISInsExcisable = rsParameterData.GetValue("InsExc_Excise")
            blnEOUFlag = rsParameterData.GetValue("EOU_Flag")
            blnISBasicRoundOff = rsParameterData.GetValue("Basic_Roundoff")
            blnISExciseRoundOff = rsParameterData.GetValue("Excise_Roundoff")
            intExcise_Roundoff_Decimal = rsParameterData.GetValue("Excise_Roundoff_Decimal")
            blnISSalesTaxRoundOff = rsParameterData.GetValue("SalesTax_Roundoff")
            blnISSurChargeTaxRoundOff = rsParameterData.GetValue("SST_Roundoff")
            blnAddCustMatrl = rsParameterData.GetValue("CustSupp_Inc")
            blnECSSTax = rsParameterData.GetValue("ECESS_Roundoff")
            intECSSRoundOffDecimal = rsParameterData.GetValue("ECESSRoundoff_Decimal")
            BlnSalesTax_Onerupee_Roundoff = rsParameterData.GetValue("SalesTax_Onerupee_Roundoff")
            blnTCSTax = rsParameterData.GetValue("TCSTax_Roundoff")
            intTCSRoundOffDecimal = rsParameterData.GetValue("TCSTax_Roundoff_decimal")
            blnInsIncSTax = rsParameterData.GetValue("InsInc_SalesTax")
            intSaleTaxRoundOffDecimal = rsParameterData.GetValue("SalesTax_Roundoff_decimal")
            blnServiceTax_Roundoff = rsParameterData.GetValue("ServiceTax_Roundoff")
            intServiceTaxRoundoff_Decimal = rsParameterData.GetValue("ServiceTaxRoundoff_Decimal")
            intGSTTAXroundoff_decimal = rsParameterData.GetValue("GSTTAX_ROUNDOFF_DECIMAL")
            blnGSTTAXroundoff = rsParameterData.GetValue("GSTTAX_ROUNDOFF")
        Else
            MsgBox("No data define in Sales_Parameter Table", MsgBoxStyle.Critical, "empower")
            SaveData = False
            rsParameterData.ResultSetClose()
            rsParameterData = Nothing
            Exit Function
        End If
        rsParameterData.ResultSetClose()
        rsParameterData = Nothing
        '**********************************
        'If Not (UCase(CmbInvType.Text) <> "REJECTION invoice") Or mstrInvoiceType <> "REJ ") Then
        If UCase(CmbInvType.Text) <> "REJECTION" Then
            strParamQuery = "SELECT CSIEX_Inc FROM customer_mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  Customer_code = '" & Trim(txtCustCode.Text) & "'"
            rsParameterData = New ClsResultSetDB
            rsParameterData.GetResult(strParamQuery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If rsParameterData.GetNoRows > 0 Then
                blnCSIEX_Inc = rsParameterData.GetValue("CSIEX_Inc")
            Else
                MsgBox("No Data found in Customer Master", MsgBoxStyle.Critical, "empower")
                SaveData = False
                rsParameterData.ResultSetClose()
                rsParameterData = Nothing
                Exit Function
            End If
            rsParameterData.ResultSetClose()
            rsParameterData = Nothing
        End If

        ldblNetInsurenceValue = System.Math.Round(Val(ctlInsurance.Text)) / Val(CStr(SpChEntry.MaxRows))
        dblAdditionalExciseAmount = 0
        ldblTotalGSTTAXTES = 0
        For lintLoopCounter = 1 To SpChEntry.MaxRows

            ldblTotalBasicValue = ldblTotalBasicValue + CalculateBasicValue(lintLoopCounter, blnISBasicRoundOff)
            ldbltotalpkgvalue = ldbltotalpkgvalue + Calculatepkg(lintLoopCounter, 2)
            ldblTempAccessibleVal = CalculateAccessibleValue(lintLoopCounter, ldblNetInsurenceValue, blnISInsExcisable)
            If blnISExciseRoundOff Then
                'ldblTotalExciseValue = System.Math.Round(CalculateExciseValue(lintLoopCounter, ldblTempAccessibleVal, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff,intExcise_Roundoff_Decimal))
                ldblTotalExciseValue = CalculateExciseValue(lintLoopCounter, ldblTempAccessibleVal, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff, intExcise_Roundoff_Decimal)

                ldblTotalCVDValue = System.Math.Round(CalculateExciseValue(lintLoopCounter, ldblTempAccessibleVal, enumExciseType.RETURN_CVD, blnEOUFlag, blnISExciseRoundOff, intExcise_Roundoff_Decimal))
                ldblTotalSADValue = System.Math.Round(CalculateExciseValue(lintLoopCounter, ldblTempAccessibleVal, enumExciseType.RETURN_SAD, blnEOUFlag, blnISExciseRoundOff, intExcise_Roundoff_Decimal))
            Else
                ldblTotalExciseValue = CalculateExciseValue(lintLoopCounter, ldblTempAccessibleVal, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff, intExcise_Roundoff_Decimal)
                ldblTotalCVDValue = System.Math.Round(CalculateExciseValue(lintLoopCounter, ldblTempAccessibleVal, enumExciseType.RETURN_CVD, blnEOUFlag, blnISExciseRoundOff, intExcise_Roundoff_Decimal))
                ldblTotalSADValue = System.Math.Round(CalculateExciseValue(lintLoopCounter, ldblTempAccessibleVal, enumExciseType.RETURN_SAD, blnEOUFlag, blnISExciseRoundOff, intExcise_Roundoff_Decimal))
            End If
            ldblTotalAccessibleValue = ldblTotalAccessibleValue + ldblTempAccessibleVal
            SpChEntry.Row = lintLoopCounter : SpChEntry.Col = 5
            lintItemQuantity = Val(SpChEntry.Text)
            SpChEntry.Row = lintLoopCounter : SpChEntry.Col = 11
            ldblTotalOthersValues = ldblTotalOthersValues + ((Val(SpChEntry.Text) / Val(ctlPerValue.Text)) * lintItemQuantity)
            SpChEntry.Row = lintLoopCounter : SpChEntry.Col = 4
            ldblTotalCustMatrlValue = ldblTotalCustMatrlValue + ((Val(SpChEntry.Text) / Val(ctlPerValue.Text)) * lintItemQuantity)
            'GST CHANGES
            If blnGSTTAXroundoff = True Then
                ldblTotalGSTTAXTES = ldblTotalGSTTAXTES + CalculateGSTtaxes(lintLoopCounter, ldblTempAccessibleVal, "CGST", blnEOUFlag)
                ldblTotalGSTTAXTES = ldblTotalGSTTAXTES + CalculateGSTtaxes(lintLoopCounter, ldblTempAccessibleVal, "SGST", blnEOUFlag)
                ldblTotalGSTTAXTES = ldblTotalGSTTAXTES + CalculateGSTtaxes(lintLoopCounter, ldblTempAccessibleVal, "UTGST", blnEOUFlag)
                ldblTotalGSTTAXTES = ldblTotalGSTTAXTES + CalculateGSTtaxes(lintLoopCounter, ldblTempAccessibleVal, "IGST", blnEOUFlag)
                ldblTotalGSTTAXTES = ldblTotalGSTTAXTES + CalculateGSTtaxes(lintLoopCounter, ldblTempAccessibleVal, "GSTCC", blnEOUFlag)
            Else
                ldblTotalGSTTAXTES = ldblTotalGSTTAXTES + System.Math.Round(CalculateGSTtaxes(lintLoopCounter, ldblTempAccessibleVal, "CGST", blnEOUFlag), intGSTTAXroundoff_decimal)
                ldblTotalGSTTAXTES = ldblTotalGSTTAXTES + System.Math.Round(CalculateGSTtaxes(lintLoopCounter, ldblTempAccessibleVal, "SGST", blnEOUFlag), intGSTTAXroundoff_decimal)
                ldblTotalGSTTAXTES = ldblTotalGSTTAXTES + System.Math.Round(CalculateGSTtaxes(lintLoopCounter, ldblTempAccessibleVal, "UTGST", blnEOUFlag), intGSTTAXroundoff_decimal)
                ldblTotalGSTTAXTES = ldblTotalGSTTAXTES + System.Math.Round(CalculateGSTtaxes(lintLoopCounter, ldblTempAccessibleVal, "IGST", blnEOUFlag), intGSTTAXroundoff_decimal)
                ldblTotalGSTTAXTES = ldblTotalGSTTAXTES + System.Math.Round(CalculateGSTtaxes(lintLoopCounter, ldblTempAccessibleVal, "GSTCC", blnEOUFlag), intGSTTAXroundoff_decimal)

            End If
            'GST CHANGES 

            If blnEOU_FLAG Then
                If blnISExciseRoundOff Then
                    ldblExciseValueForSaleTax = ldblExciseValueForSaleTax + System.Math.Round((ldblTotalExciseValue + ldblTotalCVDValue + ldblTotalSADValue) / 2, 0)
                Else
                    ldblExciseValueForSaleTax = ldblExciseValueForSaleTax + (ldblTotalExciseValue + ldblTotalCVDValue + ldblTotalSADValue) / 2
                End If
            Else
                If blnISExciseRoundOff Then
                    ldblExciseValueForSaleTax = ldblExciseValueForSaleTax + System.Math.Round(ldblTotalExciseValue, 0)
                Else
                    ldblExciseValueForSaleTax = ldblExciseValueForSaleTax + ldblTotalExciseValue
                End If
            End If

            If blnECSSTax Then
                ldblTotalECSSTaxAmount = System.Math.Round(CalculateECSSTaxValue(ldblExciseValueForSaleTax))
                ldblTotalSECSSTaxAmount = System.Math.Round(CalculateSECSSTaxValue(ldblExciseValueForSaleTax))
            Else
                ldblTotalSECSSTaxAmount = System.Math.Round(CalculateSECSSTaxValue(ldblExciseValueForSaleTax), intECSSRoundOffDecimal)
                ldblTotalECSSTaxAmount = System.Math.Round(CalculateECSSTaxValue(ldblExciseValueForSaleTax), intECSSRoundOffDecimal)
            End If
            ldblcustsuppexcisevalue = ldblcustsuppexcisevalue + CalculatecustsuppexciseValue(lintLoopCounter, intExcise_Roundoff_Decimal)


            varAdditionalExciseType = Nothing
            Call SpChEntry.GetText(8, lintLoopCounter, varAdditionalExciseType)

            If Trim(varAdditionalExciseType) <> "" Then
                TmpRs.GetResult("Select TxRt_Percentage from Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  txRt_Rate_No='" & varAdditionalExciseType & "'")
                If TmpRs.GetNoRows > 0 Then
                    dblTaxRate = Val(TmpRs.GetValue("TxRt_Percentage") & "")
                Else
                    dblTaxRate = 0
                End If
                '10727107
                dblAdditionalExciseAmount = dblAdditionalExciseAmount + ((ldblTempAccessibleVal + ldblExciseValueForSaleTax))

            Else
                dblAdditionalExciseAmount = 0
            End If
        Next
        '10727107
        dblAdditionalExciseAmount = System.Math.Round(((dblAdditionalExciseAmount + ldblTotalECSSTaxAmount + ldblTotalSECSSTaxAmount) * dblTaxRate / 100))
        aedamnt.Text = dblAdditionalExciseAmount
        '10727107
        ''Dim TmpRs As New ClsResultSetDB
        'Dim dblTaxRate As Double
        Dim varProductCode As Object
        Dim varCustDrgNo As Object
        Dim rsdb As ClsResultSetDB
        Dim mblnCSMItem As Boolean

        mblnCSMItem = False
        ldblTotalCustMatrlValue = 0

        For lintLoopCounter = 1 To SpChEntry.MaxRows
            varProductCode = Nothing
            Call SpChEntry.GetText(1, lintLoopCounter, varProductCode)
            varCustDrgNo = Nothing
            Call SpChEntry.GetText(2, lintLoopCounter, varCustDrgNo)
            rsdb = New ClsResultSetDB
            Call rsdb.GetResult("SELECT Customer_code,Finish_Item_Code,Item_Code,Cust_Drgno,Item_drgno,Description,Qty,Rate,Active_Flag,VALID_FROM,VALID_TO,GRIN_No,GRIN_Auth_Date FROM CUSTSUPPLIEDITEM_MST WHERE UNIT_CODE='" + gstrUNITID + "' AND  FINISH_ITEM_CODE =   '" & varProductCode.ToString.Trim & "' AND CUST_DRGNO =  '" & varCustDrgNo.ToString.Trim & "'  AND ACTIVE_FLAG = 1 ")
            If rsdb.RowCount > 0 Then
                mblnCSMItem = True
            End If
            rsdb.ResultSetClose()
            rsdb = Nothing
            varCustsuppmtrl = Nothing
            Call SpChEntry.GetText(4, lintLoopCounter, varCustsuppmtrl)
            varItemQuantity = Nothing
            Call SpChEntry.GetText(5, lintLoopCounter, varItemQuantity)
            varAmor = Nothing
            Call SpChEntry.GetText(16, lintLoopCounter, varAmor)

            If UCase(CmbInvType.Text) = "NORMAL INVOICE" And UCase(CmbInvSubType.Text) = "FINISHED GOODS" Then
                If mblnCSM_Knockingoff_req = True And mblnCSMItem = True Then
                    ldblTotalCustMatrlValue = ldblTotalCustMatrlValue + (System.Math.Round(Val(varCustsuppmtrl), 2))
                Else
                    ldblTotalCustMatrlValue = ldblTotalCustMatrlValue + (System.Math.Round(varCustsuppmtrl * varItemQuantity, 2))
                End If
                ldblTotalAmorValue = ldblTotalAmorValue + (System.Math.Round(varAmor * varItemQuantity, 2))
            End If

            'If Trim(varAdditionalExciseType) <> "" Then
            '    'CALCULATE TOTAL ADDITIONAL EXCISE DUTY ON CUSTOMER SUPPLIED MATERIAL AMIT KUMAR
            '    TmpRs.GetResult("Select TxRt_Percentage from Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  txRt_Rate_No='" & varAdditionalExciseType & "'")
            '    If TmpRs.GetNoRows > 0 Then
            '        dblTaxRate = Val(TmpRs.GetValue("TxRt_Percentage") & "")
            '    Else
            '        dblTaxRate = 0
            '    End If
            '    '10727107
            '        'dblAdditionalExciseAmount = dblAdditionalExciseAmount + ((Val(varCustsuppmtrl)) * (dblTaxRate / 100))
            '    dblAdditionalExciseAmount = dblAdditionalExciseAmount + ((ldblTempAccessibleVal + ldblExciseValueForSaleTax))

            'Else
            '    dblAdditionalExciseAmount = 0
            'End If
        Next
        '10727107
        'dblAdditionalExciseAmount = System.Math.Round(((dblAdditionalExciseAmount + ldblTotalECSSTaxAmount + ldblTotalSECSSTaxAmount) * dblTaxRate / 100))
        '10727107

        If blnISSalesTaxRoundOff Then
            ldblTotalSaleTaxAmount = System.Math.Round(CalculateSalesTaxValue(ldblTotalBasicValue, ldblExciseValueForSaleTax + ldblTotalECSSTaxAmount + ldblTotalSECSSTaxAmount + dblAdditionalExciseAmount, ldblTotalCustMatrlValue, ldblTotalAmorValue, ldblcustsuppexcisevalue, ldbltotalpkgvalue, blnCSIEX_Inc))
            ''10706455  
            dblAddVATamount = System.Math.Round(CalculateAdditionalSalesTaxValue(ldblTotalBasicValue, (ldblExciseValueForSaleTax + ldblTotalECSSTaxAmount + ldblTotalSECSSTaxAmount), blnInsIncSTax, Val(ctlInsurance.Text)))
            ''10706455  
        Else
            ldblTotalSaleTaxAmount = System.Math.Round(CalculateSalesTaxValue(ldblTotalBasicValue, ldblExciseValueForSaleTax + ldblTotalECSSTaxAmount + ldblTotalSECSSTaxAmount + dblAdditionalExciseAmount, ldblTotalCustMatrlValue + dblAdditionalExciseAmount, ldblTotalAmorValue, ldblcustsuppexcisevalue, ldbltotalpkgvalue, blnCSIEX_Inc), intSaleTaxRoundOffDecimal)
            ''10706455  
            dblAddVATamount = System.Math.Round(CalculateAdditionalSalesTaxValue(ldblTotalBasicValue, (ldblExciseValueForSaleTax + ldblTotalECSSTaxAmount + ldblTotalSECSSTaxAmount), blnInsIncSTax, Val(ctlInsurance.Text)), intSaleTaxRoundOffDecimal)
            ''10706455  
        End If

        If blnISSurChargeTaxRoundOff Then
            ldblTotalSurchargeTaxAmount = System.Math.Round(CalculateSurchargeTaxValue(ldblTotalSaleTaxAmount))
        Else
            ldblTotalSurchargeTaxAmount = CalculateSurchargeTaxValue(ldblTotalSaleTaxAmount)
        End If

        'ldblTotalSaleTaxAmount = System.Math.Round(ldblTotalSaleTaxAmount, 4)
        ldblTotalSaleTaxAmount = System.Math.Round(ldblTotalSaleTaxAmount, intSaleTaxRoundOffDecimal)
        ldblTotalSurchargeTaxAmount = System.Math.Round(ldblTotalSurchargeTaxAmount, 4)
        '10869290
        If ((mstrInvoiceType = "SRC") Or (CmbInvType.Text = "SERVICE INVOICE")) Then
            If blnServiceTax_Roundoff Then
                ldblTotalServiceTax_Amount = System.Math.Round(CalculateServiceTaxValue(ldblTotalAccessibleValue))
            Else
                ldblTotalServiceTax_Amount = System.Math.Round(CalculateServiceTaxValue(ldblTotalAccessibleValue), intServiceTaxRoundoff_Decimal)
            End If
        End If
        '10869290 end 
        If ((mstrInvoiceType = "JOB") Or (mstrInvoiceType = "SRC")) Then
            If blnServiceTax_Roundoff Then
                ldblTotalSBCTax_Amount = System.Math.Round(CalculateSBCTaxValue(ldblTotalAccessibleValue))
            Else
                ldblTotalSBCTax_Amount = System.Math.Round(CalculateSBCTaxValue(ldblTotalAccessibleValue), intServiceTaxRoundoff_Decimal)
            End If
        End If
        If ((mstrInvoiceType = "JOB") Or (mstrInvoiceType = "SRC")) Then
            If blnServiceTax_Roundoff Then
                ldblTotalkkcTax_Amount = System.Math.Round(CalculatekkcTaxValue(ldblTotalAccessibleValue))
            Else
                ldblTotalkkcTax_Amount = System.Math.Round(CalculatekkcTaxValue(ldblTotalAccessibleValue), intServiceTaxRoundoff_Decimal)
            End If
        End If


        If blnAddCustMatrl Then
            If blnCSIEX_Inc = False Then
                ldblTotalInvoiceValue = ldblTotalBasicValue + System.Math.Round(Val(aedamnt.Text)) + ldblTotalECSSTaxAmount + ldblTotalSECSSTaxAmount + ldblExciseValueForSaleTax + ldblTotalSaleTaxAmount + ldblTotalSurchargeTaxAmount + System.Math.Round(Val(txtFreight.Text)) + System.Math.Round(ldblTotalOthersValues) + System.Math.Round(Val(ctlInsurance.Text)) + System.Math.Round(ldblTotalCustMatrlValue, 2) + ldbltotalpkgvalue + ldblTotalServiceTax_Amount
            Else
                ldblTotalInvoiceValue = ldblTotalBasicValue + +ldblTotalServiceTax_Amount + System.Math.Round(Val(aedamnt.Text)) + ldblExciseValueForSaleTax + ldblTotalECSSTaxAmount + ldblTotalSECSSTaxAmount + ldblTotalSaleTaxAmount + ldblTotalSurchargeTaxAmount + System.Math.Round(Val(txtFreight.Text)) + System.Math.Round(ldblTotalOthersValues) + System.Math.Round(Val(ctlInsurance.Text)) + System.Math.Round(ldblTotalCustMatrlValue, 2) + ldbltotalpkgvalue - ldblcustsuppexcisevalue - System.Math.Round(ldblcustsuppexcisevalue * Val(lblECSStax_Per.Text) / 100, 2) - System.Math.Round(ldblcustsuppexcisevalue * Val(lblSECSStax_Per.Text) / 100, 2)
            End If
        Else
            If blnCSIEX_Inc = False Then
                ldblTotalInvoiceValue = ldblTotalBasicValue + +ldblTotalServiceTax_Amount + System.Math.Round(Val(aedamnt.Text)) + ldblExciseValueForSaleTax + ldblTotalSaleTaxAmount + ldblTotalSurchargeTaxAmount + System.Math.Round(Val(txtFreight.Text)) + System.Math.Round(ldblTotalOthersValues) + System.Math.Round(Val(ctlInsurance.Text), 2) + ldbltotalpkgvalue + ldblTotalECSSTaxAmount + ldblTotalSECSSTaxAmount
            Else
                ldblTotalInvoiceValue = ldblTotalBasicValue + +ldblTotalServiceTax_Amount + System.Math.Round(Val(aedamnt.Text)) + ldblExciseValueForSaleTax + ldblTotalSaleTaxAmount + ldblTotalSurchargeTaxAmount + System.Math.Round(Val(txtFreight.Text)) + System.Math.Round(ldblTotalOthersValues) + System.Math.Round(Val(ctlInsurance.Text), 2) + ldbltotalpkgvalue + ldblTotalECSSTaxAmount + ldblTotalSECSSTaxAmount - ldblcustsuppexcisevalue
            End If
        End If
        '10899126
        ldblTotalInvoiceValue = ldblTotalInvoiceValue + dblAddVATamount


        '10899126
        'ldblTotalInvoiceValue = ldblTotalInvoiceValue + dblTCSTaxAmount + dblAddVATamount
        'ldblTotalInvoiceValue = ldblTotalInvoiceValue + dblTCSTaxAmount + ldblTotalSBCTax_Amount + ldblTotalkkcTax_Amount + ldblTotalGSTTAXTES
        If UCase(CmbInvType.Text) = "EXPORT INVOICE" Then
            If CBool(Find_Value("SELECT SEZ_CUSTOMER FROM CUSTOMER_MST WHERE UNIT_CODE='" + gstrUNITID + "' AND  CUSTOMER_CODE='" & txtCustCode.Text & "'")) Then
                ldblTotalInvoiceValue = ldblTotalInvoiceValue
            Else
                ldblTotalInvoiceValue = ldblTotalInvoiceValue + ldblTotalSBCTax_Amount + ldblTotalkkcTax_Amount + ldblTotalGSTTAXTES
            End If
        Else
            ldblTotalInvoiceValue = ldblTotalInvoiceValue + ldblTotalSBCTax_Amount + ldblTotalkkcTax_Amount + ldblTotalGSTTAXTES
        End If


        '10899126
        If Val(lblTCSTaxPerDes.Text) > 0 Then
            dblTCSTaxAmount = CalculateTCSTax(ldblTotalInvoiceValue, blnTCSTax, Val(lblTCSTaxPerDes.Text))
        End If
        ldblTotalInvoiceValue = ldblTotalInvoiceValue + dblTCSTaxAmount

        mblnCSMItem = False

        Dim strStock_Location As String
        Dim varCustMat As Object

        Select Case Button
            Case "ADD"
                rsSaleConf = New ClsResultSetDB
                rsSaleConf.GetResult("Select Invoice_Type,Sub_Type,Stock_Location from SaleConf (nolock) WHERE UNIT_CODE='" + gstrUNITID + "' AND  Description ='" & Trim(CmbInvType.Text) & "'and Sub_type_Description ='" & Trim(CmbInvSubType.Text) & "' and (fin_start_date <= getdate() and fin_end_date >= getdate())")
                If rsSaleConf.GetNoRows > 0 Then
                    strStock_Location = rsSaleConf.GetValue("Stock_Location")
                Else
                    strStock_Location = ""
                End If
                strSalesChallan = ""
                If UCase(CmbInvType.Text) <> "JOBWORK INVOICE" Then
                    mstrRGP = ""
                End If

                Call SelectChallanNoFromSalesChallanDtl(1)
                If CmbInvType.Text = "NORMAL INVOICE" And CmbInvSubType.Text = "FINISHED GOODS" And DataExist("SELECT PCA_CUSTOMER FROM CUSTOMER_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & Trim(txtCustCode.Text) & "' and PCA_CUSTOMER =1 ") = True Then

                    strpPCAScheduleinvoiceknockoff = "INSERT INTO PCA_SCHEDULE_INVOICE_KNOCKOFF(UNIT_CODE ,INVOICE_NO,ITEM_CODE,PART_CODE,PACKAGE_CODE,"
                    strpPCAScheduleinvoiceknockoff += " MESSAGECALLOFFNO,QUANTITY ,MAXPOSSQTY,ENT_DT,ENT_USERID) "
                    strpPCAScheduleinvoiceknockoff += " SELECT UNIT_CODE,'" & txtChallanNo.Text & "',ITEM_CODE,PART_CODE,PACKAGECODE,MESSAGECALLOFFNO,QUANTITY,MAXPOSSQTY,getdate(),'" & mP_User & "'"
                    strpPCAScheduleinvoiceknockoff += " FROM TMP_PCA_ITEMSELECTION_PACKAGECODE WHERE unit_code='" & gstrUNITID & "' and  IP_ADDRESS='" & gstrIpaddressWinSck & "'"

                End If

                'strexportsotype = ""
                'If UCase(CmbInvType.Text) = "EXPORT INVOICE" And txtRefNo.Text <> "" Then
                '    strexportsotype = Find_Value("Select exportsotype from cust_ord_hdr where UNIT_CODE='" + gstrUNITID + "' AND  account_code='" & Trim(txtCustCode.Text) & "' and cust_ref='" & txtRefNo.Text.Trim & "' and amendment_no='" & txtAmendNo.Text & "' and active_flag='a'")
                'End If

                'strSalesChallan = "INSERT INTO SalesChallan_dtl (SBCTAX_TYPE,SBCTAX_TYPE_Per,SBCTAX_TYPE_Amount,KKCTAX_TYPE,KKCTAX_TYPE_PER,KKCTAX_TYPE_AMOUNT,ServiceTax_Type,ServiceTax_Per,ServiceTax_Amount,Location_Code,From_location,Doc_No,Suffix,Transport_Type,Vehicle_No,"
                strDYNoteNo = txtDeliveryNoteNo.Text.Trim

                If gstrUNITID = "STH" Then
                    STRCONSCODE = Find_Value("Select ISNULL(CONSIGNEE_CODE,'') From CUST_ORD_HDR WHERE UNIT_CODE='" + gstrUNITID + "' AND  ACCOUNT_CODE='" & Trim(txtCustCode.Text) & "' AND CUST_rEF='" & Trim(txtRefNo.Text) & "' AND AMENDMENT_NO='" & Trim(txtAmendNo.Text) & "'")
                End If

                strSalesChallan = "INSERT INTO SalesChallan_dtl (ServiceTax_Type,ServiceTax_Per,ServiceTax_Amount,Location_Code,From_location,Doc_No,DELIVERYNOTENO,Suffix,Transport_Type,Vehicle_No,"
                strSalesChallan = strSalesChallan & "From_Station,To_Station,Account_Code,"
                If gstrUNITID = "STH" Then
                    strSalesChallan = strSalesChallan & "consignee_code,"
                End If

                strSalesChallan = strSalesChallan & "Cust_Ref,Amendment_No,Bill_Flag,Form3,Carriage_Name,"
                strSalesChallan = strSalesChallan & "Year,Insurance,invoice_Type,Ref_Doc_No,"
                strSalesChallan = strSalesChallan & "Cust_Name ,Sales_Tax_Amount , Surcharge_Sales_Tax_Amount,"
                strSalesChallan = strSalesChallan & "Frieght_Amount,Sub_Category,SalesTax_Type,SalesTax_FormNo,SalesTax_FormValue,Annex_no,invoice_Date,Currency_code,Ent_dt,"
                If UCase(CmbInvType.Text) = "EXPORT INVOICE" And txtRefNo.Text <> "" Then
                    strSalesChallan = strSalesChallan & "Ent_UserId,Upd_dt,Upd_UserId,Exchange_Rate,total_amount,Surcharge_salesTaxType,SalesTax_Per,Surcharge_SalesTax_Per,Remarks,PerValue,LorryNo_Date,ECESS_Type, ECESS_Per, ECESS_Amount,SECESS_Type, SECESS_Per, SECESS_Amount, Tot_Add_Excise_Amt,Tot_Add_Excise_Per,payment_terms,TCSTax_Type,TCSTax_Per,TCSTaxAmount, UNIT_CODE,ADDVAT_Type,ADDVAT_Per,ADDVAT_Amount,EXPORTSOTYPE) Values ('" & IIf(txtServiceTaxType.Text.Trim.ToUpper = "UNKNOWN", "", txtServiceTaxType.Text) & "'," & lblServiceTax_Per.Text & " , " & ldblTotalServiceTax_Amount & " , '" & Trim(UCase(txtLocationCode.Text))
                Else
                    strSalesChallan = strSalesChallan & "Ent_UserId,Upd_dt,Upd_UserId,Exchange_Rate,total_amount,Surcharge_salesTaxType,SalesTax_Per,Surcharge_SalesTax_Per,Remarks,PerValue,LorryNo_Date,ECESS_Type, ECESS_Per, ECESS_Amount,SECESS_Type, SECESS_Per, SECESS_Amount, Tot_Add_Excise_Amt,Tot_Add_Excise_Per,payment_terms,TCSTax_Type,TCSTax_Per,TCSTaxAmount, UNIT_CODE,ADDVAT_Type,ADDVAT_Per,ADDVAT_Amount) Values ('" & IIf(txtServiceTaxType.Text.Trim.ToUpper = "UNKNOWN", "", txtServiceTaxType.Text) & "'," & lblServiceTax_Per.Text & " , " & ldblTotalServiceTax_Amount & " , '" & Trim(UCase(txtLocationCode.Text))
                End If

                'strSalesChallan = strSalesChallan & "Ent_UserId,Upd_dt,Upd_UserId,Exchange_Rate,total_amount,Surcharge_salesTaxType,SalesTax_Per,Surcharge_SalesTax_Per,Remarks,PerValue,LorryNo_Date,ECESS_Type, ECESS_Per, ECESS_Amount,SECESS_Type, SECESS_Per, SECESS_Amount, Tot_Add_Excise_Amt,Tot_Add_Excise_Per,payment_terms,TCSTax_Type,TCSTax_Per,TCSTaxAmount, UNIT_CODE,ADDVAT_Type,ADDVAT_Per,ADDVAT_Amount) Values ('" & Trim(txtSBC.Text) & "'," & lblSBC.Text & " , " & ldblTotalSBCTax_Amount & " ,  '" & Trim(txtKKC.Text) & "','" & lblKKC.Text & "'," & ldblTotalkkcTax_Amount & ",'" & Trim(txtServiceTaxType.Text) & "'," & lblServiceTax_Per.Text & " , " & ldblTotalServiceTax_Amount & " , '" & Trim(UCase(txtLocationCode.Text))
                'strSalesChallan = strSalesChallan & "', '" & Trim(strStock_Location) & "','" & Trim(txtChallanNo.Text) & "',''"
                strSalesChallan = strSalesChallan & "', '" & Trim(strStock_Location) & "','" & Trim(txtChallanNo.Text) & "','" & strDYNoteNo & "',''"
                strSalesChallan = strSalesChallan & ",'" & Mid(Trim(CmbTransType.Text), 1, 1) & "', '" & Trim(txtVehNo.Text) & "','"
                strSalesChallan = strSalesChallan & "','','" & Trim(txtCustCode.Text)

                If gstrUNITID = "STH" Then
                    strSalesChallan = strSalesChallan & "','" & STRCONSCODE
                End If

                'strSalesChallan = strSalesChallan & "','" & Trim(txtRefNo.Text) & "','" & Trim(mstrAmmNo) & "','0'"
                If mblnRejTracking = False Then
                    strSalesChallan = strSalesChallan & "','" & Trim(txtRefNo.Text) & "','" & Trim(mstrAmmNo) & "','0',"
                Else
                    If CmbInvType.Text = "REJECTION" Then
                        If chkRejType.Text = "LRN" Then
                            strSalesChallan = strSalesChallan & "','','" & Trim(mstrAmmNo) & "','0',"
                        Else
                            strSalesChallan = strSalesChallan & "','" & Trim(txtRefNo.Text) & "','" & Trim(mstrAmmNo) & "','0',"
                        End If
                    Else
                        strSalesChallan = strSalesChallan & "','" & Trim(txtRefNo.Text) & "','" & Trim(mstrAmmNo) & "','0',"
                    End If
                End If

                strSalesChallan = strSalesChallan & "'','" & Trim(txtCarrServices.Text)
                strSalesChallan = strSalesChallan & "','" & Year(dtpDateDesc.Value).ToString & "',"
                strSalesChallan = strSalesChallan & System.Math.Round(Val(ctlInsurance.Text)) & ",'" & Trim(rsSaleConf.GetValue("Invoice_type")) & "','"
                strSalesChallan = strSalesChallan & Trim(mstrRGP) & "','"
                strSalesChallan = strSalesChallan & Trim(lblCustCodeDes.Text) & "',"
                strSalesChallan = strSalesChallan & Val(CStr(ldblTotalSaleTaxAmount)) & "," & Val(CStr(ldblTotalSurchargeTaxAmount)) & "," & System.Math.Round(Val(txtFreight.Text)) & ",'" & Trim(rsSaleConf.GetValue("Sub_Type")) & "','"
                strSalesChallan = strSalesChallan & Trim(txtSaleTaxType.Text) & "','"
                strSalesChallan = strSalesChallan & "0',0,'0','"
                strSalesChallan = strSalesChallan & getDateForDB(dtpDateDesc.Value) & "','" & lblCurrencyDes.Text & "',getdate(),'" & mP_User & "',  getdate() ,'" & mP_User & "','"
                strSalesChallan = strSalesChallan & Val(lblExchangeRateValue.Text) & "'," & ldblTotalInvoiceValue & ",'" & Trim(txtSurchargeTaxType.Text) & "'," & Val(lblSaltax_Per.Text) & "," & Val(lblSurcharge_Per.Text) & ",'" & Trim(txtRemarks.Text) & "'," & ctlPerValue.Text & ",'" & Trim(TxtLRNO.Text) & "'"
                strSalesChallan = strSalesChallan & ",'" & Trim(txtECSSTaxType.Text) & "'," & Val(lblECSStax_Per.Text)
                strSalesChallan = strSalesChallan & "," & ldblTotalECSSTaxAmount & ",'" & Trim(txtSECSSTaxType.Text) & "'," & Val(lblSECSStax_Per.Text) & " ," & ldblTotalSECSSTaxAmount & ", " & dblAdditionalExciseAmount & "," & dblTaxRate & ",'" & lblcreditterm.Text.Trim() & "'"
                strSalesChallan = strSalesChallan & ",'" & Trim(txtTCSTaxCode.Text) & "'," & Val(lblTCSTaxPerDes.Text)
                '''10706455  
                If UCase(CmbInvType.Text) = "EXPORT INVOICE" And txtRefNo.Text <> "" Then
                    strSalesChallan = strSalesChallan & "," & dblTCSTaxAmount & ",'" + gstrUNITID + "', '" & Trim(txtAddVAT.Text) & "', " & Val(lblAddVAT.Text) & ", " & dblAddVATamount & ",'" & mstrexportsotype & "')"
                Else
                    strSalesChallan = strSalesChallan & "," & dblTCSTaxAmount & ",'" + gstrUNITID + "', '" & Trim(txtAddVAT.Text) & "', " & Val(lblAddVAT.Text) & ", " & dblAddVATamount & " )"
                End If

                ''10706455  
                rsSaleConf.ResultSetClose()
                rsSaleConf = Nothing
                strSalesDtl = ""

                With SpChEntry
                    For lintLoopCounter = 1 To .MaxRows
                        dblExcise_Amount = 0

                        .Row = lintLoopCounter
                        .Col = 1
                        lstrItemCode = Trim(.Text)
                        .Col = 2
                        lstrItemDrgno = Trim(.Text)
                        rsdb = New ClsResultSetDB
                        Call rsdb.GetResult("SELECT Customer_code,Finish_Item_Code,Item_Code,Cust_Drgno,Item_drgno,Description,Qty,Rate,Active_Flag,VALID_FROM,VALID_TO,GRIN_No,GRIN_Auth_Date FROM CUSTSUPPLIEDITEM_MST WHERE UNIT_CODE='" + gstrUNITID + "' AND  FINISH_ITEM_CODE =   '" & lstrItemCode.ToString.Trim & "' AND CUST_DRGNO =  '" & lstrItemDrgno.ToString.Trim & "'  AND ACTIVE_FLAG = 1 ")
                        If rsdb.RowCount > 0 Then
                            mblnCSMItem = True
                        End If
                        rsdb.ResultSetClose()
                        rsdb = Nothing
                        .Col = 3
                        ldblItemRate = Val(.Text) / Val(ctlPerValue.Text)

                        '10808160
                        .Col = 27
                        strModel = Trim(.Text)
                        '10808160

                        .Col = 4
                        ldblItemCustMtrl = Val(.Text) / Val(ctlPerValue.Text)
                        .Col = 5
                        lintItemQuantity = Val(.Text)
                        .Col = 6
                        ldblItemPacking = Val(.Text)
                        .Col = 7
                        lstrItemExciseCode = Trim(.Text)
                        .Col = 8
                        varAdditionalExciseType = .Text
                        'CALCULATE TOTAL ADDITIONAL EXCISE DUTY ON CUSTOMER SUPPLIED MATERIAL AMIT KUMAR
                        If Trim(varAdditionalExciseType) <> "" Then
                            TmpRs = New ClsResultSetDB
                            TmpRs.GetResult("Select TxRt_Percentage from Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  txRt_Rate_No='" & varAdditionalExciseType & "'")
                            If TmpRs.GetNoRows > 0 Then
                                dblTaxRate = Val(TmpRs.GetValue("TxRt_Percentage") & "")
                            Else
                                dblTaxRate = 0
                            End If
                            'If CmbInvType.Text = "NORMAL INVOICE" And CmbInvSubType.Text = "FINISHED GOODS" And mblnCSM_Knockingoff_req = True And mblnCSMItem = True Then
                            '    dblAdditionalExciseAmount = (ldblItemCustMtrl * (dblTaxRate / 100))
                            'Else
                            '    dblAdditionalExciseAmount = ((ldblItemCustMtrl * lintItemQuantity) * (dblTaxRate / 100))
                            'End If
                            TmpRs.ResultSetClose()
                            TmpRs = Nothing
                        Else
                            dblAdditionalExciseAmount = 0
                        End If
                        'CALCULATE TOTAL ADDITIONAL EXCISE DUTY ON CUSTOMER SUPPLIED MATERIAL AMIT KUMAR
                        .Col = 9
                        lstrItemCVDCode = Trim(.Text)
                        .Col = 10
                        lstrItemSADCode = Trim(.Text)
                        .Col = 11
                        ldblItemOthers = Val(.Text) / Val(ctlPerValue.Text)
                        .Col = 12
                        ldblItemFromBox = Val(.Text)
                        .Col = 13
                        ldblItemToBox = Val(.Text)
                        .Col = 15
                        lstrItemDelete = Trim(.Text)
                        .Col = 16
                        ldblItemToolCost = Val(.Text) / Val(ctlPerValue.Text)
                        '11 mar 2016
                        .Col = 16
                        dblItemToolCost = Val(.Text)
                        '11 mar 2016
                        .Col = 28
                        strAtnno = .Text

                        'GST CHANGES
                        STRHSNSACCODE = ""
                        .Col = 30
                        STRHSNSACCODE = Trim(.Text)

                        STRCGSTTAXTYPE = ""
                        .Col = 31
                        STRCGSTTAXTYPE = Trim(.Text)

                        STRSGSTTXRT_TYPE = ""
                        .Col = 32
                        STRSGSTTXRT_TYPE = Trim(.Text)

                        STRUTGSTTXRT_TYPE = ""
                        .Col = 33
                        STRUTGSTTXRT_TYPE = Trim(.Text)

                        STRIGSTTXRT_TYPE = ""
                        .Col = 34
                        STRIGSTTXRT_TYPE = Trim(.Text)

                        STRCOMPENSATION_CESS = ""
                        .Col = 35
                        STRCOMPENSATION_CESS = Trim(.Text)

                        'GST CHANGES

                        rsCustItemMst = New ClsResultSetDB
                        rsItemMst = New ClsResultSetDB
                        rsExternalsalesorder = New ClsResultSetDB

                        rsItemMst.GetResult("SELECT Description FROM Item_Mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  Item_Code ='" & Trim(lstrItemCode) & "'", ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                        rsCustItemMst.GetResult("SELECT Drg_desc FROM CustItem_Mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  Account_code ='" & Trim(txtCustCode.Text) & "'and Cust_DrgNo='" & lstrItemDrgno & "'and Item_code ='" & lstrItemCode & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        rsExternalsalesorder.GetResult("SELECT external_salesorder_no FROM cust_ord_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  Account_code ='" & Trim(txtCustCode.Text) & "'and Cust_DrgNo='" & lstrItemDrgno & "'and Item_code ='" & lstrItemCode & "'and active_flag='A' and cust_ref='" & Me.txtRefNo.Text & "' ", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        If UCase(Trim(lstrItemDelete)) <> "D" Then
                            strSalesDtl = Trim(strSalesDtl) & "INSERT INTO sales_Dtl(EOP_MODEL,Location_Code,Doc_No,Suffix,Item_Code,Sales_Quantity,"
                            strSalesDtl = strSalesDtl & "From_Box,To_Box,Rate,Sales_Tax,Excise_Tax,Packing,Others,Cust_Mtrl,"
                            strSalesDtl = strSalesDtl & "Year,Cust_Item_Code,Cust_Item_Desc,Tool_Cost,Measure_Code,Excise_type,SalesTax_type,CVD_type,SAD_type,Basic_Amount,Accessible_amount,CVD_Amount,SVD_amount,"
                            strSalesDtl = strSalesDtl & "Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,Excise_per,CVD_per,SVD_per,CustMtrl_Amount,ToolCost_Amount,PerValue,pkg_amount,csiexcise_amount,ADD_EXCISE_TYPE,ADD_EXCISE_PER,ADD_EXCISE_AMOUNT,UNIT_CODE,EXTERNAL_SALESORDER_NO,LINE_NO,HSNSACCODE,CGSTTXRT_TYPE,CGST_PERCENT,CGST_AMT,SGSTTXRT_TYPE,SGST_PERCENT,SGST_AMT,UTGSTTXRT_TYPE,UTGST_PERCENT,UTGST_AMT,IGSTTXRT_TYPE,IGST_PERCENT,IGST_AMT,COMPENSATION_CESS_TYPE,COMPENSATION_CESS_PERCENT,COMPENSATION_CESS_AMT ) values ('" & strModel & "','" & Trim(UCase(txtLocationCode.Text)) & "','"
                            strSalesDtl = strSalesDtl & Trim(txtChallanNo.Text) & "','','" & Trim(lstrItemCode) & "','" & Val(CStr(lintItemQuantity)) & "','"
                            strSalesDtl = strSalesDtl & Val(CStr(ldblItemFromBox)) & "','" & Val(CStr(ldblItemToBox)) & "'," & Val(CStr(ldblItemRate)) & "," & Trim(lblSaltax_Per.Text) & ","
                            TempAccessibleVal = CalculateAccessibleValue(lintLoopCounter, ldblNetInsurenceValue, blnISInsExcisable)
                            If blnISExciseRoundOff Then
                                '10736222
                                dblExcise_Amount = System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff, intExcise_Roundoff_Decimal))
                                strSalesDtl = strSalesDtl & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff, intExcise_Roundoff_Decimal))
                            Else
                                '10736222
                                dblExcise_Amount = CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff, intExcise_Roundoff_Decimal)
                                strSalesDtl = strSalesDtl & CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff, intExcise_Roundoff_Decimal)
                            End If

                            If mblnCSM_Knockingoff_req = True And mblnCSMItem = True Then
                                strSalesDtl = strSalesDtl & "," & Val(CStr(ldblItemPacking)) & "," & Val(CStr(ldblItemOthers)) & "," & Val(CStr(ldblItemCustMtrl)) / Val(CStr(lintItemQuantity)) & ",'"
                            Else
                                strSalesDtl = strSalesDtl & "," & Val(CStr(ldblItemPacking)) & "," & Val(CStr(ldblItemOthers)) & "," & Val(CStr(ldblItemCustMtrl)) & ",'"
                            End If
                            strSalesDtl = strSalesDtl & Year(dtpDateDesc.Value).ToString & "','" & Trim(lstrItemDrgno) & "','" & IIf((Len(Trim(rsCustItemMst.GetValue("Drg_Desc"))) <= 0 Or Trim(CStr(rsCustItemMst.GetValue("Drg_Desc") = "Unknown"))), Trim(rsItemMst.GetValue("Description")), Trim(rsCustItemMst.GetValue("Drg_Desc"))) & "',"
                            If UCase(CmbInvType.Text) = "NORMAL INVOICE" Or UCase(CmbInvType.Text) = "EXPORT INVOICE" Then
                                If UCase(CmbInvSubType.Text) <> "SCRAP" Then
                                    '11 mar 2016
                                    'strSalesDtl = strSalesDtl & mdblToolCost(lintLoopCounter - 1) & ",'','"
                                    strSalesDtl = strSalesDtl & dblItemToolCost & ",'','"
                                    '11 mar 2016
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
                                strSalesDtl = strSalesDtl & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_CVD, blnEOUFlag, blnISExciseRoundOff, intExcise_Roundoff_Decimal))
                                strSalesDtl = strSalesDtl & "," & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_SAD, blnEOUFlag, blnISExciseRoundOff, intExcise_Roundoff_Decimal))
                            Else
                                strSalesDtl = strSalesDtl & (CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_CVD, blnEOUFlag, blnISExciseRoundOff, intExcise_Roundoff_Decimal))
                                strSalesDtl = strSalesDtl & "," & (CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_SAD, blnEOUFlag, blnISExciseRoundOff, intExcise_Roundoff_Decimal))
                            End If

                            strSalesDtl = strSalesDtl & ",GetDate(),'"
                            strSalesDtl = strSalesDtl & Trim(mP_User) & "', GetDate(),'" & Trim(mP_User) & "'," & GetTaxRate(lstrItemExciseCode, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='EXC'") & "," & GetTaxRate(lstrItemCVDCode, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='CVD'") & "," & GetTaxRate(lstrItemSADCode, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='SAD'") & ","
                            If CmbInvType.Text = "NORMAL INVOICE" And CmbInvSubType.Text = "FINISHED GOODS" And mblnCSM_Knockingoff_req Then
                                strSalesDtl = strSalesDtl & System.Math.Round(Val(CStr(ldblItemCustMtrl)), 2) & ","
                            Else
                                strSalesDtl = strSalesDtl & System.Math.Round(Val(CStr(lintItemQuantity * ldblItemCustMtrl)), 2) & ","
                            End If

                            strSalesDtl = strSalesDtl & System.Math.Round(Val(CStr(lintItemQuantity * ldblItemToolCost)), 2) & "," & ctlPerValue.Text & ""
                            strSalesDtl = strSalesDtl & "," & (Calculatepkg(lintLoopCounter, 2))

                            Dim mblnAllowCustomerSpecificReport_COMP As Boolean = SqlConnectionclass.ExecuteScalar("Select isnull(AllowCustomerSpecificReport_COMP,0) from customer_mst (Nolock) where customer_code='" & Trim(txtCustCode.Text) & "' and unit_code='" + gstrUNITID + "'")
                            'changes done on 02 dec 2024 by prashant rajpal -EXTERNAL INTERNAL SALES ORDER SHOULD COME IN RAW MATERIAL INVOICE TOO
                            Dim mblnAllowCustomerSpecificReport_RAW As Boolean = SqlConnectionclass.ExecuteScalar("Select isnull(AllowCustomerSpecificReport_RAW ,0) from customer_mst (Nolock) where customer_code='" & Trim(txtCustCode.Text) & "' and unit_code='" + gstrUNITID + "'")

                            If (CmbInvType.Text = "NORMAL INVOICE" And CmbInvSubType.Text = "FINISHED GOODS") Or (CmbInvType.Text = "NORMAL INVOICE" And CmbInvSubType.Text = "COMPONENTS" And mblnAllowCustomerSpecificReport_COMP = True) And DataExist("SELECT SOUPLD_LINE_LEVEL_SALESORDER FROM CUSTOMER_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & Trim(txtCustCode.Text) & "'") = True Then
                                strSalesDtl = strSalesDtl & "," & CalculatecustsuppexciseValue(lintLoopCounter, intExcise_Roundoff_Decimal) & ",'" & Trim(varAdditionalExciseType) & "'," & dblTaxRate & "," & 0 & ",'" + gstrUNITID + "','" & rsExternalsalesorder.GetValue("External_salesorder_NO").ToString & "'," & lintLoopCounter
                            ElseIf (CmbInvType.Text = "NORMAL INVOICE" And CmbInvSubType.Text = "FINISHED GOODS") Or (CmbInvType.Text = "NORMAL INVOICE" And CmbInvSubType.Text = "RAW MATERIAL" And mblnAllowCustomerSpecificReport_RAW = True) And DataExist("SELECT SOUPLD_LINE_LEVEL_SALESORDER FROM CUSTOMER_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & Trim(txtCustCode.Text) & "'") = True Then '' CONDITION ADDED ON 02 DEC 2024
                                strSalesDtl = strSalesDtl & "," & CalculatecustsuppexciseValue(lintLoopCounter, intExcise_Roundoff_Decimal) & ",'" & Trim(varAdditionalExciseType) & "'," & dblTaxRate & "," & 0 & ",'" + gstrUNITID + "','" & rsExternalsalesorder.GetValue("External_salesorder_NO").ToString & "'," & lintLoopCounter
                            Else
                                strSalesDtl = strSalesDtl & "," & CalculatecustsuppexciseValue(lintLoopCounter, intExcise_Roundoff_Decimal) & ",'" & Trim(varAdditionalExciseType) & "'," & dblTaxRate & "," & 0 & ",'" + gstrUNITID + "',''," & lintLoopCounter
                            End If
                            'changes done on 02 dec 2024 by prashant rajpal 
                            'GST CHANGES
                            If gblnGSTUnit = True Then
                                If blnGSTTAXroundoff = True Then
                                    strSalesDtl = strSalesDtl & ",'" & STRHSNSACCODE & "','" & STRCGSTTAXTYPE & "','" & GetTaxRate(STRCGSTTAXTYPE, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='CGST'") & "'," & System.Math.Round(CalculateGSTtaxes(lintLoopCounter, TempAccessibleVal, "CGST", blnEOUFlag)) & ",'" & STRSGSTTXRT_TYPE & "','" & GetTaxRate(STRSGSTTXRT_TYPE, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='SGST'") & "'," & System.Math.Round(CalculateGSTtaxes(lintLoopCounter, TempAccessibleVal, "SGST", blnEOUFlag)) & ",'" & STRUTGSTTXRT_TYPE & "','" & GetTaxRate(STRUTGSTTXRT_TYPE, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='UTGST'") & "'," & System.Math.Round(CalculateGSTtaxes(lintLoopCounter, TempAccessibleVal, "UTGST", blnEOUFlag)) & ",'" & STRIGSTTXRT_TYPE & "','" & GetTaxRate(STRIGSTTXRT_TYPE, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='IGST'") & "'," & System.Math.Round(CalculateGSTtaxes(lintLoopCounter, TempAccessibleVal, "IGST", blnEOUFlag)) & ",'" & STRCOMPENSATION_CESS & "','" & GetTaxRate(STRCOMPENSATION_CESS, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='COMPENSATION CESS'") & "'," & System.Math.Round(CalculateGSTtaxes(lintLoopCounter, TempAccessibleVal, "GSTCC", blnEOUFlag)) & ")"
                                Else
                                    strSalesDtl = strSalesDtl & ",'" & STRHSNSACCODE & "','" & STRCGSTTAXTYPE & "','" & GetTaxRate(STRCGSTTAXTYPE, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='CGST'") & "'," & System.Math.Round(CalculateGSTtaxes(lintLoopCounter, TempAccessibleVal, "CGST", blnEOUFlag), intGSTTAXroundoff_decimal) & ",'" & STRSGSTTXRT_TYPE & "','" & GetTaxRate(STRSGSTTXRT_TYPE, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='SGST'") & "'," & System.Math.Round(CalculateGSTtaxes(lintLoopCounter, TempAccessibleVal, "SGST", blnEOUFlag), intGSTTAXroundoff_decimal) & ",'" & STRUTGSTTXRT_TYPE & "','" & GetTaxRate(STRUTGSTTXRT_TYPE, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='UTGST'") & "'," & System.Math.Round(CalculateGSTtaxes(lintLoopCounter, TempAccessibleVal, "UTGST", blnEOUFlag), intGSTTAXroundoff_decimal) & ",'" & STRIGSTTXRT_TYPE & "','" & GetTaxRate(STRIGSTTXRT_TYPE, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='IGST'") & "'," & System.Math.Round(CalculateGSTtaxes(lintLoopCounter, TempAccessibleVal, "IGST", blnEOUFlag), intGSTTAXroundoff_decimal) & ",'" & STRCOMPENSATION_CESS & "','" & GetTaxRate(STRCOMPENSATION_CESS, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='COMPENSATION CESS'") & "'," & System.Math.Round(CalculateGSTtaxes(lintLoopCounter, TempAccessibleVal, "GSTCC", blnEOUFlag), intGSTTAXroundoff_decimal) & ")"
                                End If

                            Else
                                strSalesDtl = strSalesDtl & ",'','',0,0,'',0,0,'',0,0,'',0,0,'',0,0"
                                strSalesDtl = strSalesDtl & ")" & vbCrLf
                            End If
                            'GST CHANGES

                            If mblnBatchTrack = True And UCase(CmbInvType.Text) <> "JOBWORK INVOICE" And UCase(CmbInvType.Text) <> "REJECTION" Then
                                For Intcounter = 1 To UBound(mBatchData(lintLoopCounter).Batch_No)
                                    strBatchNo = mBatchData(lintLoopCounter).Batch_No(Intcounter)
                                    dblBatchQty = mBatchData(lintLoopCounter).Batch_Quantity(Intcounter)
                                    strbatchdtl = strbatchdtl & " Insert into ItemBatch_dtl ("
                                    strbatchdtl = strbatchdtl & " Doc_Type,Doc_No,Doc_Date,From_Location,To_Location,Item_Code,Cust_Item_Code,Serial_No,"
                                    strbatchdtl = strbatchdtl & " Batch_No,Batch_Date,Batch_Qty,Batch_Accepted_Qty,Batch_Rejected_Qty,"
                                    strbatchdtl = strbatchdtl & " Batch_ReInspected_Qty,Ent_Userid,Ent_Dt,Upd_Userid,Upd_Dt,UNIT_CODE)"
                                    strbatchdtl = strbatchdtl & " Values (9999,'" & Trim(txtChallanNo.Text) & "','" & getDateForDB(GetServerDate()) & "','" & mstrLocationCode & "','" & strToLocation & "','" & lstrItemCode & "','" & lstrItemDrgno & "',1,'" & strBatchNo & "','" & getDateForDB(GetServerDate()) & "'," & dblBatchQty & "," & dblBatchQty & ",0,0,'" & mP_User & "',getdate(),'" & mP_User & "',getdate(),'" + gstrUNITID + "') "
                                Next
                            End If

                            If AllowASNTextFileGeneration(Trim(txtCustCode.Text)) = True Then
                                strInsertUpdateASNdtl = Trim(strInsertUpdateASNdtl) & "INSERT INTO MKT_ASN_INVDTL(Doc_no,Cust_PlantCode,ARL_Code,ASN_Status,Cust_Part_Code,Cummulative_Qty,UNIT_CODE)values('"
                                strInsertUpdateASNdtl = strInsertUpdateASNdtl & Trim(txtChallanNo.Text) & "','" & Trim(txtPlantCode.Text) & "','" & Trim(txtActualReceivingLoc.Text) & "',0,'"
                                strInsertUpdateASNdtl = strInsertUpdateASNdtl & Trim(lstrItemDrgno) & "','" & Val(CStr(lintItemQuantity)) & "','" + gstrUNITID + "')" & vbCrLf
                                If CheckExistanceOfFieldData(Trim(lstrItemDrgno), "cust_part_code", "MKT_ASN_CUMFIG", "(cust_part_code='" & Trim(lstrItemDrgno) & "' and cust_PlantCode='" & Trim(txtPlantCode.Text) & "')") = False Then
                                    strInsertUpdateASNdtl = strInsertUpdateASNdtl & "INSERT INTO MKT_ASN_CUMFIG(Cust_Part_Code,Cust_PlantCode,Cummulative_Qty,UNIT_CODE) VALUES('" & Trim(lstrItemDrgno) & "','" & Trim(txtPlantCode.Text) & "',0,'" + gstrUNITID + "')" & vbCrLf
                                End If
                            End If

                            '10736222
                            strSql = "select dbo.UDF_ISCT2INVOICE( '" & gstrUNITID & "','" & txtCustCode.Text.Trim & "','" & CmbInvType.Text.Trim & "','" & CmbInvSubType.Text.Trim & "','" & txtRefNo.Text.Trim & "')"
                            If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSql)) = True Then
                                blnIsCt2 = True
                                strSqlct2qry = "INSERT INTO TMP_CT2_INVOICE_KNOCKOFF ([UNIT_CODE],[CUST_CODE],[SONO],[AMENDMENT_NO],[TMP_INVOICE_NO],[ITEM_CODE],[CUST_DRG_NO],[CURRENCY_CODE],[QTY],[RATE],[TOOL_COST],[EXCISE_TAX],[EXCISE_AMOUNT],[ECESS_TYPE],[SECESS_TYPE],[IP_ADDRESS]) "
                                strSqlct2qry = strSqlct2qry + " Values('" & gstrUNITID & "','" & txtCustCode.Text.Trim & "','" & txtRefNo.Text.Trim & "','" & txtAmendNo.Text.Trim & "','" & Me.txtChallanNo.Text.Trim & "',"
                                '11 mar 2016
                                'strSqlct2qry = strSqlct2qry + "'" & lstrItemCode.Trim & "','" & lstrItemDrgno.Trim & "','" & lblCurrencyDes.Text.Trim & "'," & Val(CStr(lintItemQuantity)) & "," & Val(CStr(ldblItemRate)) & "," & Val(mdblToolCost(lintLoopCounter - 1)) & ",'" & lstrItemExciseCode.Trim & "'," & dblExcise_Amount & ",'" & txtECSSTaxType.Text.Trim & "','" & txtSECSSTaxType.Text.Trim & "','" & gstrIpaddressWinSck & "' ) "
                                strSqlct2qry = strSqlct2qry + "'" & lstrItemCode.Trim & "','" & lstrItemDrgno.Trim & "','" & lblCurrencyDes.Text.Trim & "'," & Val(CStr(lintItemQuantity)) & "," & Val(CStr(ldblItemRate)) & "," & dblItemToolCost & ",'" & lstrItemExciseCode.Trim & "'," & dblExcise_Amount & ",'" & txtECSSTaxType.Text.Trim & "','" & txtSECSSTaxType.Text.Trim & "','" & gstrIpaddressWinSck & "' ) "
                                '11 mar 2016
                                SqlConnectionclass.ExecuteNonQuery(strSqlct2qry)
                            End If
                        End If
                        rsCustItemMst.ResultSetClose()
                        rsItemMst.ResultSetClose()
                        mblnCSMItem = False

                        If mblnRejTracking = True And CmbInvType.Text = "REJECTION" Then
                            strSQLREJINV_DTL = strSQLREJINV_DTL & MakeSQL_REJINVTRACKING(lintLoopCounter)
                        End If

                    Next
                End With
                'Rajesh
            Case "EDIT"
                If DataExist("SELECT PCA_CUSTOMER FROM CUSTOMER_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & Trim(txtCustCode.Text) & "' and PCA_CUSTOMER =1 ") = True Then
                    strpPCAScheduleinvoiceknockoff = " DELETE FROM PCA_SCHEDULE_INVOICE_KNOCKOFF WHERE UNIT_CODE='" & gstrUNITID & "' AND INVOICE_NO='" & txtChallanNo.Text & "' "

                    strpPCAScheduleinvoiceknockoff += "INSERT INTO PCA_SCHEDULE_INVOICE_KNOCKOFF(UNIT_CODE ,INVOICE_NO,ITEM_CODE,PART_CODE,PACKAGE_CODE,"
                    strpPCAScheduleinvoiceknockoff += " MESSAGECALLOFFNO,QUANTITY ,MAXPOSSQTY,ENT_DT,ENT_USERID) "
                    strpPCAScheduleinvoiceknockoff += " SELECT UNIT_CODE,'" & txtChallanNo.Text & "',ITEM_CODE,PART_CODE,PACKAGECODE,MESSAGECALLOFFNO,QUANTITY,MAXPOSSQTY,getdate(),'" & mP_User & "'"
                    strpPCAScheduleinvoiceknockoff += " FROM TMP_PCA_ITEMSELECTION_PACKAGECODE WHERE unit_code='" & gstrUNITID & "' and  IP_ADDRESS='" & gstrIpaddressWinSck & "'"

                End If

                strSalesChallan = ""
                strSalesChallan = "UPDATE SalesChallan_Dtl SET Insurance = " & System.Math.Round(Val(ctlInsurance.Text))
                strSalesChallan = strSalesChallan & ",DeliveryNoteNo ='" & txtDeliveryNoteNo.Text.Trim & "'"
                If blnISSalesTaxRoundOff Then
                    strSalesChallan = strSalesChallan & ",Sales_Tax_Amount =" & System.Math.Round(Val(CStr(ldblTotalSaleTaxAmount)))
                    strSalesChallan = strSalesChallan & ",ADDVAT_Type='" & txtAddVAT.Text.Trim() & "',AddVat_Per=" & Val(lblAddVAT.Text) & ",ADDVAT_Amount =" & System.Math.Round(Val(CStr(dblAddVATamount)))
                Else
                    strSalesChallan = strSalesChallan & ",Sales_Tax_Amount =" & Val(CStr(ldblTotalSaleTaxAmount))
                    strSalesChallan = strSalesChallan & ",ADDVAT_Type='" & txtAddVAT.Text.Trim() & "',AddVat_Per=" & Val(lblAddVAT.Text) & ",ADDVAT_Amount =" & Val(CStr(dblAddVATamount))
                End If
                If blnISSurChargeTaxRoundOff Then
                    strSalesChallan = strSalesChallan & ",Surcharge_Sales_Tax_Amount =" & System.Math.Round(Val(CStr(ldblTotalSurchargeTaxAmount)))
                Else
                    strSalesChallan = strSalesChallan & ",Surcharge_Sales_Tax_Amount =" & Val(CStr(ldblTotalSurchargeTaxAmount))
                End If
                strSalesChallan = strSalesChallan & ",Frieght_Amount=" & System.Math.Round(Val(txtFreight.Text))
                strSalesChallan = strSalesChallan & ",SalesTax_Type='" & Trim(txtSaleTaxType.Text) & "'"
                strSalesChallan = strSalesChallan & ",total_amount=" & ldblTotalInvoiceValue
                strSalesChallan = strSalesChallan & ",Surcharge_salesTaxType='" & Trim(txtSurchargeTaxType.Text) & "'"
                strSalesChallan = strSalesChallan & ",SalesTax_Per=" & Val(lblSaltax_Per.Text)
                strSalesChallan = strSalesChallan & ",Surcharge_SalesTax_Per=" & Val(lblSurcharge_Per.Text)
                strSalesChallan = strSalesChallan & ",Remarks= '" & Trim(txtRemarks.Text) & "'"
                strSalesChallan = strSalesChallan & ",ECESS_Type = '" & txtECSSTaxType.Text & "'"
                strSalesChallan = strSalesChallan & ",ECESS_Per = " & Val(lblECSStax_Per.Text)
                strSalesChallan = strSalesChallan & ",ECESS_Amount = " & ldblTotalECSSTaxAmount
                strSalesChallan = strSalesChallan & ",SECESS_Type = '" & txtSECSSTaxType.Text & "'"
                strSalesChallan = strSalesChallan & ",SECESS_Per = " & Val(lblSECSStax_Per.Text)
                strSalesChallan = strSalesChallan & ",SECESS_Amount = " & ldblTotalSECSSTaxAmount
                strSalesChallan = strSalesChallan & ",Tot_Add_Excise_Amt = " & dblAdditionalExciseAmount
                strSalesChallan = strSalesChallan & ",Tot_Add_Excise_Per = " & dblTaxRate
                strSalesChallan = strSalesChallan & ",TCSTax_Type = '" & txtTCSTaxCode.Text & "'"
                strSalesChallan = strSalesChallan & ",TCSTax_Per = " & Val(lblTCSTaxPerDes.Text)
                strSalesChallan = strSalesChallan & ",TCSTaxAmount = " & dblTCSTaxAmount
                strSalesChallan = strSalesChallan & ",LorryNo_Date= '" & Trim(TxtLRNO.Text) & "'"
                strSalesChallan = strSalesChallan & ",Carriage_Name= '" & Trim(txtCarrServices.Text) & "'"
                strSalesChallan = strSalesChallan & ",Vehicle_No= '" & Trim(txtVehNo.Text) & "'"

                If ((mstrInvoiceType = "JOB") Or (mstrInvoiceType = "SRC")) Then
                    strSalesChallan = strSalesChallan & ",ServiceTax_Type =  '" & Trim(txtServiceTaxType.Text) & "'"
                    strSalesChallan = strSalesChallan & ",ServiceTax_Per = " & lblServiceTax_Per.Text
                    strSalesChallan = strSalesChallan & ",ServiceTax_Amount = " & ldblTotalServiceTax_Amount
                    strSalesChallan = strSalesChallan & ",SBCTAX_TYPE =  '" & Trim(txtSBC.Text) & "'"
                    strSalesChallan = strSalesChallan & ",SBCTAX_TYPE_Per = " & lblSBC.Text
                    strSalesChallan = strSalesChallan & ",SBCTAX_TYPE_Amount = " & ldblTotalSBCTax_Amount
                    strSalesChallan = strSalesChallan & ",KKCTAX_TYPE =  '" & Trim(txtKKC.Text) & "'"
                    strSalesChallan = strSalesChallan & ",KKCTAX_TYPE_Per = " & lblKKC.Text
                    strSalesChallan = strSalesChallan & ",KKCTAX_TYPE_Amount = " & ldblTotalkkcTax_Amount
                End If

                strSalesChallan = strSalesChallan & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Location_Code ='" & Trim(UCase(txtLocationCode.Text)) & "'"
                strSalesChallan = strSalesChallan & " and Doc_No ='" & Val(txtChallanNo.Text) & "'"
                strSalesDtl = ""
                strSalesDtlDelete = ""
                With SpChEntry
                    For lintLoopCounter = 1 To .MaxRows
                        .Row = lintLoopCounter
                        .Col = 1
                        lstrItemCode = Trim(.Text)
                        .Col = 4
                        varCustMat = Val(.Text) / Val(ctlPerValue.Text)
                        .Col = 5
                        lintItemQuantity = Val(.Text)
                        '10808160
                        .Col = 27
                        strModel = Trim(.Text)
                        '10808160
                        .Col = 2
                        lstrItemDrgno = Trim(.Text)
                        rsdb = New ClsResultSetDB
                        Call rsdb.GetResult("SELECT Customer_code,Finish_Item_Code,Item_Code,Cust_Drgno,Item_drgno,Description,Qty,Rate,Active_Flag,VALID_FROM,VALID_TO,GRIN_No,GRIN_Auth_Date FROM CUSTSUPPLIEDITEM_MST WHERE UNIT_CODE='" + gstrUNITID + "' AND  FINISH_ITEM_CODE =   '" & lstrItemCode.ToString.Trim & "' AND CUST_DRGNO =  '" & lstrItemDrgno.ToString.Trim & "'  AND ACTIVE_FLAG = 1 ")
                        If rsdb.RowCount > 0 Then
                            mblnCSMItem = True
                        End If
                        rsdb.ResultSetClose()
                        rsdb = Nothing
                        'in case of edit if rate is changed not update in database
                        .Col = 3
                        lintItemrate = Val(.Text)
                        .Col = 15
                        lstrItemDelete = Trim(.Text)
                        .Col = 7
                        lstrItemExciseCode = Trim(.Text)
                        '.Col = 8
                        'varAdditionalExciseType = .Text
                        'CALCULATE TOTAL ADDITIONAL EXCISE DUTY ON CUSTOMER SUPPLIED MATERIAL AMIT KUMAR
                        'If Trim(varAdditionalExciseType) <> "" Then
                        '    TmpRs = New ClsResultSetDB
                        '    TmpRs.GetResult("Select TxRt_Percentage from Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  txRt_Rate_No='" & varAdditionalExciseType & "'")
                        '    If TmpRs.GetNoRows > 0 Then
                        '        dblTaxRate = Val(TmpRs.GetValue("TxRt_Percentage") & "")
                        '    Else
                        '        dblTaxRate = 0
                        '    End If
                        '    dblAdditionalExciseAmount = dblAdditionalExciseAmount + ((ldblTempAccessibleVal + ldblExciseValueForSaleTax))
                        '    TmpRs = Nothing
                        'Else
                        '    dblAdditionalExciseAmount = 0
                        'End If
                        ''10727107
                        'dblAdditionalExciseAmount = System.Math.Round(((dblAdditionalExciseAmount + ldblTotalECSSTaxAmount + ldblTotalSECSSTaxAmount) * dblTaxRate / 100))
                        'aedamnt.Text = dblAdditionalExciseAmount
                        ''10727107
                        'CALCULATE TOTAL ADDITIONAL EXCISE DUTY ON CUSTOMER SUPPLIED MATERIAL AMIT KUMAR
                        .Col = 9
                        lstrItemCVDCode = Trim(.Text)
                        .Col = 10
                        lstrItemSADCode = Trim(.Text)
                        .Col = 12
                        ldblItemFromBox = Val(.Text)
                        .Col = 13
                        ldblItemToBox = Val(.Text)
                        '11 mar 2016
                        .Col = 16
                        dblItemToolCost = Val(.Text)
                        '11 mar 2016
                        .Col = 16
                        ldblItemToolCost = Val(.Text) / Val(ctlPerValue.Text)
                        If UCase(lstrItemDelete) <> "D" Then
                            strSalesDtl = Trim(strSalesDtl) & "UPDATE Sales_dtl SET EOP_MODEL='" & strModel & "', Sales_Quantity ='" & Val(CStr(lintItemQuantity)) & "',Sales_Tax =" & Trim(lblSaltax_Per.Text) & ","
                            If mblnCSM_Knockingoff_req = True And mblnCSMItem = True Then
                                strSalesDtl = Trim(strSalesDtl) & "rate = " & Val(CStr(lintItemrate)) & ", Cust_Mtrl= " & Val(CStr(varCustMat)) / Val(CStr(lintItemQuantity)) & ", CustMtrl_Amount= " & Val(CStr(varCustMat)) & ",ToolCost_Amount=" & Val(CStr(lintItemQuantity * ldblItemToolCost))
                            Else
                                strSalesDtl = Trim(strSalesDtl) & "rate = " & Val(CStr(lintItemrate)) & ", Cust_Mtrl= " & Val(CStr(varCustMat)) & ", CustMtrl_Amount= " & Val(CStr(varCustMat)) & ",ToolCost_Amount=" & Val(CStr(lintItemQuantity * ldblItemToolCost))
                            End If

                            TempAccessibleVal = CalculateAccessibleValue(lintLoopCounter, ldblNetInsurenceValue, blnISInsExcisable)
                            If blnISExciseRoundOff Then
                                '10736222
                                dblExcise_Amount = System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff, intExcise_Roundoff_Decimal))
                                strSalesDtl = Trim(strSalesDtl) & ",Excise_Tax=" & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff, intExcise_Roundoff_Decimal))
                            Else
                                '10736222
                                dblExcise_Amount = CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff, intExcise_Roundoff_Decimal)
                                strSalesDtl = Trim(strSalesDtl) & ",Excise_Tax=" & CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_EXCISE, blnEOUFlag, blnISExciseRoundOff, intExcise_Roundoff_Decimal)
                            End If

                            strSalesDtl = Trim(strSalesDtl) & ",Excise_type='" & lstrItemExciseCode & "',SalesTax_type='" & Trim(txtSaleTaxType.Text) & "'"
                            strSalesDtl = Trim(strSalesDtl) & ",CVD_type='" & Trim(lstrItemCVDCode) & "',SAD_type='" & Trim(lstrItemSADCode) & "',Basic_Amount=" & CalculateBasicValue(lintLoopCounter, blnISBasicRoundOff)
                            strSalesDtl = Trim(strSalesDtl) & ",Accessible_amount=" & Val(CStr(TempAccessibleVal))
                            If blnISExciseRoundOff Then
                                strSalesDtl = Trim(strSalesDtl) & ",CVD_Amount=" & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_CVD, blnEOUFlag, blnISExciseRoundOff, intExcise_Roundoff_Decimal)) & ",SVD_amount=" & System.Math.Round(CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_SAD, blnEOUFlag, blnISExciseRoundOff, intExcise_Roundoff_Decimal))
                            Else
                                strSalesDtl = Trim(strSalesDtl) & ",CVD_Amount=" & CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_CVD, blnEOUFlag, blnISExciseRoundOff, intExcise_Roundoff_Decimal) & ",SVD_amount=" & CalculateExciseValue(lintLoopCounter, TempAccessibleVal, enumExciseType.RETURN_SAD, blnEOUFlag, blnISExciseRoundOff, intExcise_Roundoff_Decimal)
                            End If
                            strSalesDtl = Trim(strSalesDtl) & ",pkg_Amount=" & (Calculatepkg(lintLoopCounter, 2))
                            strSalesDtl = Trim(strSalesDtl) & ",ADD_EXCISE_TYPE='" & Trim(varAdditionalExciseType) & "'"
                            strSalesDtl = Trim(strSalesDtl) & ",ADD_EXCISE_PER=" & dblTaxRate
                            strSalesDtl = Trim(strSalesDtl) & ",ADD_EXCISE_AMOUNT=" & dblAdditionalExciseAmount
                            strSalesDtl = Trim(strSalesDtl) & ",csiexcise_Amount=" & CalculatecustsuppexciseValue(lintLoopCounter, intExcise_Roundoff_Decimal)
                            strSalesDtl = Trim(strSalesDtl) & ",Excise_per=" & GetTaxRate(lstrItemExciseCode, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='EXC'")
                            strSalesDtl = Trim(strSalesDtl) & ",CVD_per=" & GetTaxRate(lstrItemCVDCode, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='CVD'")
                            strSalesDtl = Trim(strSalesDtl) & ",SVD_per=" & GetTaxRate(lstrItemSADCode, "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " Tx_TaxeID='SAD'")
                            strSalesDtl = Trim(strSalesDtl) & ",Tool_Cost =" & ldblItemToolCost & ",From_box = " & ldblItemFromBox & ", To_box = " & ldblItemToBox

                            'GST CHANGES
                            If blnGSTTAXroundoff = True Then
                                strSalesDtl = Trim(strSalesDtl) & ",CGST_AMT=" & CalculateGSTtaxes(lintLoopCounter, TempAccessibleVal, "CGST", blnEOUFlag)
                                strSalesDtl = Trim(strSalesDtl) & ",SGST_AMT=" & CalculateGSTtaxes(lintLoopCounter, TempAccessibleVal, "SGST", blnEOUFlag)
                                strSalesDtl = Trim(strSalesDtl) & ",UTGST_AMT=" & CalculateGSTtaxes(lintLoopCounter, TempAccessibleVal, "UTGST", blnEOUFlag)
                                strSalesDtl = Trim(strSalesDtl) & ",IGST_AMT=" & CalculateGSTtaxes(lintLoopCounter, TempAccessibleVal, "IGST", blnEOUFlag)
                                strSalesDtl = Trim(strSalesDtl) & ",COMPENSATION_CESS_AMT=" & CalculateGSTtaxes(lintLoopCounter, TempAccessibleVal, "GSTCC", blnEOUFlag)
                            Else
                                strSalesDtl = Trim(strSalesDtl) & ",CGST_AMT=" & System.Math.Round(CalculateGSTtaxes(lintLoopCounter, TempAccessibleVal, "CGST", blnEOUFlag), intGSTTAXroundoff_decimal)
                                strSalesDtl = Trim(strSalesDtl) & ",SGST_AMT=" & System.Math.Round(CalculateGSTtaxes(lintLoopCounter, TempAccessibleVal, "SGST", blnEOUFlag), intGSTTAXroundoff_decimal)
                                strSalesDtl = Trim(strSalesDtl) & ",UTGST_AMT=" & System.Math.Round(CalculateGSTtaxes(lintLoopCounter, TempAccessibleVal, "UTGST", blnEOUFlag), intGSTTAXroundoff_decimal)
                                strSalesDtl = Trim(strSalesDtl) & ",IGST_AMT=" & System.Math.Round(CalculateGSTtaxes(lintLoopCounter, TempAccessibleVal, "IGST", blnEOUFlag), intGSTTAXroundoff_decimal)
                                strSalesDtl = Trim(strSalesDtl) & ",COMPENSATION_CESS_AMT=" & System.Math.Round(CalculateGSTtaxes(lintLoopCounter, TempAccessibleVal, "GSTCC", blnEOUFlag), intGSTTAXroundoff_decimal)
                            End If

                            'GST CHANGES


                            strSalesDtl = Trim(strSalesDtl) & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Location_Code ='" & Trim(UCase(txtLocationCode.Text)) & "'"
                            strSalesDtl = Trim(strSalesDtl) & " and Doc_No =" & Val(txtChallanNo.Text) & " and Cust_Item_Code='"
                            strSalesDtl = Trim(strSalesDtl) & Trim(lstrItemDrgno) & "'" & vbCrLf

                            '10736222
                            strSql = "select dbo.UDF_ISCT2INVOICE( '" & gstrUNITID & "','" & txtCustCode.Text.Trim & "','" & CmbInvType.Text.Trim & "','" & CmbInvSubType.Text.Trim & "','" & txtRefNo.Text.Trim & "')"
                            If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSql)) = True Then
                                blnIsCt2 = True
                                strSqlct2qry = "INSERT INTO TMP_CT2_INVOICE_KNOCKOFF ([UNIT_CODE],[CUST_CODE],[SONO],[AMENDMENT_NO],[TMP_INVOICE_NO],[ITEM_CODE],[CUST_DRG_NO],[CURRENCY_CODE],[QTY],[RATE],[TOOL_COST],[EXCISE_TAX],[EXCISE_AMOUNT],[ECESS_TYPE],[SECESS_TYPE],[IP_ADDRESS]) "
                                strSqlct2qry = strSqlct2qry + " Values('" & gstrUNITID & "','" & txtCustCode.Text.Trim & "','" & txtRefNo.Text.Trim & "','" & txtAmendNo.Text.Trim & "','" & Me.txtChallanNo.Text.Trim & "',"
                                '11 mar 2016
                                'strSqlct2qry = strSqlct2qry + "'" & lstrItemCode.Trim & "','" & lstrItemDrgno.Trim & "','" & lblCurrencyDes.Text.Trim & "'," & Val(CStr(lintItemQuantity)) & "," & Val(CStr(lintItemrate)) & "," & Val(mdblToolCost(lintLoopCounter - 1)) & ",'" & lstrItemExciseCode.Trim & "'," & dblExcise_Amount & ",'" & txtECSSTaxType.Text.Trim & "','" & txtSECSSTaxType.Text.Trim & "','" & gstrIpaddressWinSck & "' ) "
                                strSqlct2qry = strSqlct2qry + "'" & lstrItemCode.Trim & "','" & lstrItemDrgno.Trim & "','" & lblCurrencyDes.Text.Trim & "'," & Val(CStr(lintItemQuantity)) & "," & Val(CStr(lintItemrate)) & "," & dblItemToolCost & ",'" & lstrItemExciseCode.Trim & "'," & dblExcise_Amount & ",'" & txtECSSTaxType.Text.Trim & "','" & txtSECSSTaxType.Text.Trim & "','" & gstrIpaddressWinSck & "' ) "
                                '11 mar 2016
                                SqlConnectionclass.ExecuteNonQuery(strSqlct2qry)
                            End If
                        Else
                            strSalesDtlDelete = Trim(strSalesDtlDelete) & "DELETE Sales_dtl "
                            strSalesDtlDelete = Trim(strSalesDtlDelete) & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Location_Code ='" & Trim(UCase(txtLocationCode.Text)) & "'"
                            strSalesDtlDelete = Trim(strSalesDtlDelete) & " and Doc_No =" & Val(txtChallanNo.Text) & " and Cust_Item_Code='"
                            strSalesDtlDelete = Trim(strSalesDtlDelete) & Trim(lstrItemDrgno) & "'" & vbCrLf
                        End If

                        If mblnBatchTrack = True And UCase(CmbInvType.Text) <> "JOBWORK INVOICE" And UCase(CmbInvType.Text) <> "REJECTION INVOICE" Then
                            If UCase(lstrItemDelete) <> "D" Then
                                strbatchdtl = strbatchdtl & " Delete from ItemBatch_Dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  Doc_no = " & Trim(Me.txtChallanNo.Text) & " and From_Location = '" & mstrLocationCode & "' and cust_Item_Code  = '" & lstrItemDrgno & "' and Doc_Type = 9999 "
                                For Intcounter = 1 To UBound(mBatchData(lintLoopCounter).Batch_No)
                                    strBatchNo = mBatchData(lintLoopCounter).Batch_No(Intcounter)
                                    dblBatchQty = mBatchData(lintLoopCounter).Batch_Quantity(Intcounter)
                                    strbatchdtl = strbatchdtl & " Insert into ItemBatch_dtl ("
                                    strbatchdtl = strbatchdtl & " Doc_Type,Doc_No,Doc_Date,From_Location,To_Location,Item_Code,Cust_Item_Code,Serial_No,"
                                    strbatchdtl = strbatchdtl & " Batch_No,Batch_Date,Batch_Qty,Batch_Accepted_Qty,Batch_Rejected_Qty,"
                                    strbatchdtl = strbatchdtl & " Batch_ReInspected_Qty,Ent_Userid,Ent_Dt,Upd_Userid,Upd_Dt,UNIT_CODE)"
                                    strbatchdtl = strbatchdtl & " Values (9999,'" & Trim(txtChallanNo.Text) & "','" & getDateForDB(GetServerDate()) & "','" & mstrLocationCode & "','" & strToLocation & "','" & lstrItemCode & "','" & lstrItemDrgno & "',1,'" & strBatchNo & "','" & getDateForDB(GetServerDate()) & "'," & dblBatchQty & "," & dblBatchQty & ",0,0,'" & mP_User & "','" & getDateForDB(GetServerDate()) & "','" & mP_User & "','" & getDateForDB(GetServerDate()) & "','" + gstrUNITID + "') "
                                Next
                            Else
                                strbatchdtl = strbatchdtl & " Delete from ItemBatch_Dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  Doc_no = " & Trim(Me.txtChallanNo.Text) & " and From_Location = '" & mstrLocationCode & "' and cust_Item_Code  = '" & lstrItemDrgno & "' and Doc_Type = 9999 "
                            End If
                        End If

                        If mblnRejTracking = True And mstrInvType = "REJ" Then
                            strSQLREJINV_DTL = strSQLREJINV_DTL & MakeSQL_REJINVTRACKING(lintLoopCounter)
                        End If

                        If AllowASNTextFileGeneration(Trim(txtCustCode.Text)) = True Then
                            If UCase(lstrItemDelete) <> "D" Then
                                strInsertUpdateASNdtl = Trim(strInsertUpdateASNdtl) & "UPDATE MKT_ASN_INVDTL SET Cummulative_Qty=" & Val(CStr(lintItemQuantity)) & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  DOC_NO='" & Val(txtChallanNo.Text) & "' and Cust_part_Code='" & Trim(lstrItemDrgno) & "'" & vbCrLf
                            Else
                                strInsertUpdateASNdtl = Trim(strInsertUpdateASNdtl) & "DELETE FROM MKT_ASN_INVDTL WHERE UNIT_CODE='" + gstrUNITID + "' AND  DOC_NO='" & Val(txtChallanNo.Text) & "' and Cust_part_Code='" & Trim(lstrItemDrgno) & "'" & vbCrLf
                            End If
                        End If
                        mblnCSMItem = False
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

        Dim strtime As String = GetServerDateTime()
        Dim RSCNT As ADODB.Recordset

        If GetPaletteStatus(IIf(Button = "EDIT", strInvType, CmbInvType.Text), IIf(Button = "EDIT", strInvSubType, CmbInvSubType.Text), txtCustCode.Text.Trim) Then
            If dtPaletteItemQty IsNot Nothing AndAlso dtPaletteItemQty.Rows.Count > 0 Then
                strPalatteQuery = String.Empty
                If Button = "EDIT" Then
                    strPalatteQuery += "DELETE FROM INVOICE_PALETTE_DTL WHERE UNIT_CODE='" & gstrUNITID & "' AND TEMP_INVOICE_NO=" & txtChallanNo.Text & ""
                End If
                For i As Integer = 0 To dtPaletteItemQty.Rows.Count - 1
                    strPalatteQuery += "INSERT INTO INVOICE_PALETTE_DTL(TEMP_INVOICE_NO,ITEM_CODE,PALETTE_LABEL," &
                                        "QTY,UNIT_CODE,ENT_DT,ENT_USERID,UPD_DT,UPD_USERID) " &
                                        "VALUES (" & txtChallanNo.Text & ",'" & Convert.ToString(dtPaletteItemQty.Rows(i)("ITEM_CODE")) & "'" &
                                        ",'" & Convert.ToString(dtPaletteItemQty.Rows(i)("PALETTE_LABEL")) & "'" &
                                        "," & Convert.ToInt32(dtPaletteItemQty.Rows(i)("QTY")) & "" &
                                        ",'" & gstrUNITID & "',GETDATE(),'" & mP_User & "',GETDATE(),'" & mP_User & "')"

                Next
            End If
        End If

        With mP_Connection
            .BeginTrans()

            .Execute(strSalesChallan, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            If Len(Trim(strupSalechallan)) > 0 Then
                .Execute(strupSalechallan, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End If

            '10736222
            If Len(Trim(strpPCAScheduleinvoiceknockoff)) > 0 Then
                .Execute(strpPCAScheduleinvoiceknockoff, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End If

            .Execute(strSalesDtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            If Len(Trim(mstrUpdDispatchSql)) > 0 Then
                .Execute(mstrUpdDispatchSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End If

            If Len(Trim(strPalatteQuery)) > 0 Then
                .Execute(strPalatteQuery, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End If

            If Not mblnCSM_Knockingoff_req Then
                If Button = "EDIT" Then
                    mP_Connection.Execute("DELETE FROM CSM_INVOICE_DTL WHERE UNIT_CODE='" + gstrUNITID + "' AND  DOC_NO = " & txtChallanNo.Text & " AND INVOICE_LOCK = 0", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                End If
                With SpChEntry
                    Dim varitemcode, varqty As Object
                    Dim objCn As New ADODB.Command
                    objCn.ActiveConnection = mP_Connection
                    objCn.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                    objCn.CommandText = "USP_INSERT_CSM_INVOICE_DTL"
                    objCn.CommandTimeout = 0
                    For lintLoopCounter = 1 To .MaxRows
                        varitemcode = Nothing
                        Call .GetText(1, lintLoopCounter, varitemcode)
                        varqty = Nothing
                        Call .GetText(5, lintLoopCounter, varqty)
                        objCn.Parameters.Append(objCn.CreateParameter("@UNITCODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                        objCn.Parameters.Append(objCn.CreateParameter("@CUST_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, txtCustCode.Text))
                        objCn.Parameters.Append(objCn.CreateParameter("@FG_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 25, varitemcode))
                        objCn.Parameters.Append(objCn.CreateParameter("@QTY", ADODB.DataTypeEnum.adBigInt, ADODB.ParameterDirectionEnum.adParamInput, , varqty))
                        objCn.Parameters.Append(objCn.CreateParameter("@DOC_NO", ADODB.DataTypeEnum.adBigInt, ADODB.ParameterDirectionEnum.adParamInput, , txtChallanNo.Text.Trim))
                        objCn.Parameters.Append(objCn.CreateParameter("@USER_ID", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 12, mP_User))
                        objCn.Parameters.Append(objCn.CreateParameter("@RETURN", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInputOutput, , 0))
                        objCn.Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        If objCn.Parameters(objCn.Parameters.Count - 1).Value = 1 Then
                            MsgBox("Unable To Save CSM Details.", MsgBoxStyle.Information, ResolveResString(100))
                            objCn = Nothing
                            mP_Connection.RollbackTrans()
                            SaveData = False
                            Exit Function
                        End If
                        objCn.Parameters.Delete(6)
                        objCn.Parameters.Delete(5)
                        objCn.Parameters.Delete(4)
                        objCn.Parameters.Delete(3)
                        objCn.Parameters.Delete(2)
                        objCn.Parameters.Delete(1)
                        objCn.Parameters.Delete(0)
                    Next lintLoopCounter
                    objCn = Nothing
                End With
            End If

            If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                If mblnRejTracking = True And CmbInvType.Text = "REJECTION" Then
                    If Len(Trim(strSQLREJINV_DTL)) > 0 Then
                        .Execute(strSQLREJINV_DTL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End If
                End If
            ElseIf CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                If mblnRejTracking = True And strInvType = "REJ" Then
                    If Len(Trim(strSQLREJINV_DTL)) > 0 Then
                        .Execute(strSQLREJINV_DTL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End If
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
            If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                If Len(Trim(strSalesDtlDelete)) > 0 Then
                    .Execute(strSalesDtlDelete, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                End If
            End If

            If Len(Trim(strbatchdtl)) > 0 Then
                .Execute(strbatchdtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End If

            If Len(strInsertUpdateASNdtl) > 0 Then
                .Execute(strInsertUpdateASNdtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End If

            If Trim(CmbInvType.Text) = "NORMAL INVOICE" And MULTIPLESO > 1 Then
                .Execute("update " + frmMKTTRN0020NEW.strTmpTable + " set doc_no = " & Val(txtChallanNo.Text) & "", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                RSCNT = .Execute("select count(*) from " + frmMKTTRN0020NEW.strTmpTable)
                If RSCNT.Fields(0).Value > 1 Then
                    .Execute("insert into Emp_InvoiceSOLinkage select DISTINCT Doc_No,Cust_Ref,Amendment_No,'" & gstrUNITID & "' from " + frmMKTTRN0020NEW.strTmpTable, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    .Execute("update saleschallan_dtl set MultipleSO = 1,Cust_Ref = '',Amendment_No='' WHERE UNIT_CODE='" + gstrUNITID + "' AND  doc_no = " & Val(txtChallanNo.Text) & "", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                End If
                .Execute("Drop table " + frmMKTTRN0020NEW.strTmpTable, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End If
            'Added for Issue ID 22303 Starts
            'issue id 10192547
            'If InvAgstBarCode() = True And CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT And mstrFGDomestic = "1" Then
            If InvAgstBarCode() = True And CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT And (mstrFGDomestic = "1" Or mstrFGDomestic = "2") Then
                If BarCodeTracking(Trim(txtChallanNo.Text), "EDIT") = True Then
                    mP_Connection.Execute(mstrupdateBarBondedStockFlag, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                End If
            End If

            'ADDED BY VINOD FOR GLOBAL TOOL INVOICE
            If UCase(CmbInvType.Text) = "NORMAL INVOICE" Or UCase(CmbInvType.Text) = "TRANSFER INVOICE" Or UCase(CmbInvType.Text) = "INTER-DIVISION" Or UCase(CmbInvType.Text) = "SERVICE INVOICE" Then '101254587 
                If UCase(CmbInvSubType.Text) = "ASSETS" Or UCase(CmbInvSubType.Text) = "RAW MATERIAL" Or UCase(CmbInvSubType.Text) = "INPUTS" Or UCase(CmbInvSubType.Text) = "SERVICE" Then '101254587
                    If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                        If UpdateGlobalToolInvoice(Val(txtChallanNo.Text), "A", Me.txtCustCode.Text.Trim) = False Then
                            Exit Function
                        End If
                    ElseIf CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                        If UpdateGlobalToolInvoice(Val(txtChallanNo.Text), "E", Me.txtCustCode.Text.Trim) = False Then
                            Exit Function
                        End If
                    End If

                End If
            End If
            'END OF CHANGES

            ' Changes start to update Cust_ref in case of PDS_TOYOTA_CUSTOMER is ON .issue comes report blank--Priti
            If UCase(CmbInvType.Text) <> "REJECTION" Then
                If ((CBool(Find_Value("SELECT ISNULL(MULTIPLE_SO_PDS_TOYOTA,0)as MULTIPLE_SO_PDS_TOYOTA FROM SALES_PARAMETER where unit_code = '" & gstrUNITID & "'")) = True)) Then
                    If CBool(Find_Value("SELECT ISNULL(PDS_TOYOTA_CUSTOMER,0)as PDS_TOYOTA_CUSTOMER  FROM CUSTOMER_MST WHERE UNIT_CODE = '" & gstrUNITID & "'AND CUSTOMER_CODE ='" & txtCustCode.Text.Trim & "'")) = True Then
                        strSql = "update sales_dtl  set Cust_ref='" & txtRefNo.Text & "',amendment_no='" & txtAmendNo.Text & "' where doc_no ='" & Val(txtChallanNo.Text) & "' "
                        mP_Connection.Execute(strSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End If

                End If
            End If
            'END OF CHANGES

            .CommitTrans()
            Call Logging_Starting_End_Time("Invoice Entry ", strtime, "Saved", txtChallanNo.Text)
        End With
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        mP_Connection.RollbackTrans()
        SaveData = False
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
        '    Dim lintQty As Long
        Dim lintQty As Double
        Dim rsdb As ClsResultSetDB
        Dim stritemCode As String
        Dim strCustDrgNo As String
        Dim mblnCSMItem As Boolean
        mblnCSMItem = False
        On Error GoTo ErrHandler
        With SpChEntry
            .Row = pintRowNo : .Col = 1 : stritemCode = .Text
            .Row = pintRowNo : .Col = 2 : strCustDrgNo = .Text
            rsdb = New ClsResultSetDB
            Call rsdb.GetResult("SELECT Customer_code,Finish_Item_Code,Item_Code,Cust_Drgno,Item_drgno,Description,Qty,Rate,Active_Flag,VALID_FROM,VALID_TO,GRIN_No,GRIN_Auth_Date FROM CUSTSUPPLIEDITEM_MST WHERE UNIT_CODE='" + gstrUNITID + "' AND  FINISH_ITEM_CODE = '" & stritemCode.ToString.Trim & "' AND CUST_DRGNO =  '" & strCustDrgNo.ToString.Trim & "'  AND ACTIVE_FLAG = 1 ")
            If rsdb.RowCount > 0 Then
                mblnCSMItem = True
            End If
            rsdb.ResultSetClose()
            rsdb = Nothing
            .Row = pintRowNo
            .Col = 3
            ldblRate = Val(.Text) / Val(ctlPerValue.Text)
            .Col = 6
            ldblPkg_Per = Val(.Text)
            .Col = 5
            lintQty = Val(.Text)
            .Col = 4
            ldblCustMat = Val(.Text) / Val(ctlPerValue.Text)
            .Col = 16
            ldblToolCost = Val(.Text) / Val(ctlPerValue.Text)
            If pblnISInsAdd = True Then
                If mblnCSM_Knockingoff_req = True And mblnCSMItem = True Then
                    CalculateAccessibleValue = System.Math.Round((ldblRate + ldblToolCost + pdblInsurenceValue + ldblPkg_Per) * lintQty, 2) + ldblCustMat
                Else
                    CalculateAccessibleValue = System.Math.Round((ldblRate + ldblCustMat + ldblToolCost + pdblInsurenceValue + ldblPkg_Per) * lintQty, 2)
                End If
            Else
                If mblnCSM_Knockingoff_req = True And mblnCSMItem = True Then
                    CalculateAccessibleValue = System.Math.Round((ldblRate + ldblToolCost + ldblPkg_Per) * lintQty, 2) + ldblCustMat
                Else
                    CalculateAccessibleValue = System.Math.Round((ldblRate + ldblCustMat + ldblToolCost + ldblPkg_Per) * lintQty, 2)
                End If
            End If
        End With
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function CalculateExciseValue(ByVal pintRowNo As Short, ByVal pdblAccessibleValue As Double, ByVal penumTaxType As enumExciseType, ByRef pblnEOU_FLAG As Boolean, ByRef blnExciseFlag As Boolean, ByVal pintexciseroundoff As Short) As Double
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
        Dim strsql As String
        On Error GoTo ErrHandler
        ldblTempTotalExcise = 0
        ldblTempTotalCVD = 0
        ldblTempTotalSAD = 0
        With SpChEntry
            .Row = pintRowNo
            .Col = 7
            rsGetTaxRate = New ClsResultSetDB
            strTableSql = "SELECT TxRt_Percentage FROM Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  TxRt_Rate_No='" & Trim(.Text) & "'"
            rsGetTaxRate.GetResult(strTableSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If rsGetTaxRate.GetNoRows > 0 Then
                ldblTaxRate = rsGetTaxRate.GetValue("TxRt_Percentage")
            Else
                ldblTaxRate = 0
            End If
            rsGetTaxRate.ResultSetClose()
            If pblnEOU_FLAG Then
                .Col = 9
                strTableSql = "SELECT TxRt_Percentage FROM Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  TxRt_Rate_No='" & Trim(.Text) & "'"
                rsGetTaxRate = New ClsResultSetDB
                rsGetTaxRate.GetResult(strTableSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                If rsGetTaxRate.GetNoRows > 0 Then
                    ldblCVDRate = rsGetTaxRate.GetValue("TxRt_Percentage")
                Else
                    ldblCVDRate = 0
                End If
                rsGetTaxRate.ResultSetClose()
                .Col = 10
                strTableSql = "SELECT TxRt_Percentage FROM Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  TxRt_Rate_No='" & Trim(.Text) & "'"
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
                End If
                ldblTempTotalCVD = (((ldblTempTotalExcise + pdblAccessibleValue) * ldblCVDRate) / 100)
                If blnExciseFlag = True Then
                    ldblTempTotalCVD = System.Math.Round(ldblTempTotalCVD, 0)
                End If
                ldblTempTotalSAD = (((ldblTempTotalCVD + ldblTempTotalExcise + pdblAccessibleValue) * ldblSADRate) / 100)
                If blnExciseFlag = True Then
                    ldblTempTotalSAD = System.Math.Round(ldblTempTotalSAD, 0)
                End If
                If penumTaxType = enumExciseType.RETURN_EXCISE Then
                    CalculateExciseValue = (ldblTempTotalExcise)
                ElseIf penumTaxType = enumExciseType.RETURN_CVD Then
                    CalculateExciseValue = ldblTempTotalCVD
                Else
                    CalculateExciseValue = ldblTempTotalSAD
                End If
            Else
                If penumTaxType = enumExciseType.RETURN_EXCISE Then
                    CalculateExciseValue = ((pdblAccessibleValue * ldblTaxRate) / 100)
                ElseIf penumTaxType = enumExciseType.RETURN_CVD Then
                    CalculateExciseValue = 0
                Else
                    CalculateExciseValue = 0
                End If
            End If
        End With
        If blnExciseFlag = True Then
            strsql = "select dbo.UFN_ROUNDOFF_DECIMAL(" & CalculateExciseValue & " )"
            CalculateExciseValue = SqlConnectionclass.ExecuteScalar(strsql)
        Else
            CalculateExciseValue = System.Math.Round(CalculateExciseValue, pintexciseroundoff)
        End If

        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function CalculateSalesTaxValue(ByVal pdblTotalBasicValue As Double, ByVal pdblTotalExciseValue As Double, ByRef pdblTotalCustMtrl As Double, ByRef pdblTotalamort As Double, ByRef pdblcustsuppexcisevalue As Double, ByRef pdblTotalpackagevalue As Double, ByRef PblnblnCSIEX_Inc As Boolean) As Double
        '---------------------------------------------------------------------------------------
        'Name       :   CalculateSalesTaxValue
        'Type       :   Function
        'Author     :   Tapan Jain
        'Arguments  :
        'Return     :
        'Purpose    :
        '---------------------------------------------------------------------------------------
        Dim rsCustomer As New ClsResultSetDB
        Dim rsChallanEntry As ClsResultSetDB
        Dim strInvoiceType As String
        On Error GoTo ErrHandler
        If PblnblnCSIEX_Inc = True Then
            'pdblTotalExciseValue = pdblTotalExciseValue - pdblcustsuppexcisevalue
        End If
        If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            strInvoiceType = UCase(Trim(CmbInvType.Text))
        ElseIf CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
            rsChallanEntry = New ClsResultSetDB
            rsChallanEntry.GetResult("Select a.Description,a.Sub_Type_Description from SaleConf a (nolock) ,SalesChallan_Dtl b where A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" + gstrUNITID + "' AND Doc_No = " & txtChallanNo.Text & " and a.Invoice_Type = b.Invoice_type and a.Sub_type = b.Sub_Category and a.Location_code = b.Location_code and (fin_start_date <= getdate() and fin_end_date >= getdate())")
            strInvoiceType = UCase(rsChallanEntry.GetValue("Description"))
            rsChallanEntry.ResultSetClose()
        End If
        If UCase(Trim(strInvoiceType)) <> "REJECTION" Then
            rsCustomer.GetResult("Select CST_EVAL,CST_AMT from Customer_Mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  Customer_code = '" & Trim(txtCustCode.Text) & "'")
            If rsCustomer.GetNoRows > 0 Then
                If rsCustomer.GetValue("CST_EVAL") = True And rsCustomer.GetValue("CST_AMT") = True Then
                    CalculateSalesTaxValue = ((pdblTotalBasicValue + pdblTotalExciseValue + pdblTotalCustMtrl + pdblTotalamort + pdblTotalpackagevalue) * Val(lblSaltax_Per.Text)) / 100
                End If
                If rsCustomer.GetValue("CST_EVAL") = True And rsCustomer.GetValue("CST_AMT") = False Then
                    CalculateSalesTaxValue = ((pdblTotalBasicValue + pdblTotalExciseValue + pdblTotalCustMtrl + pdblTotalpackagevalue) * Val(lblSaltax_Per.Text)) / 100
                End If
                If rsCustomer.GetValue("CST_EVAL") = False And rsCustomer.GetValue("CST_AMT") = True Then
                    CalculateSalesTaxValue = ((pdblTotalBasicValue + pdblTotalExciseValue + pdblTotalamort + pdblTotalpackagevalue - pdblTotalamort - pdblcustsuppexcisevalue) * Val(lblSaltax_Per.Text)) / 100
                End If
                If rsCustomer.GetValue("CST_EVAL") <> True And rsCustomer.GetValue("CST_AMT") <> True Then
                    CalculateSalesTaxValue = ((pdblTotalBasicValue + pdblTotalExciseValue + pdblTotalpackagevalue) * Val(lblSaltax_Per.Text)) / 100
                End If
            Else
                CalculateSalesTaxValue = ((pdblTotalBasicValue + pdblTotalExciseValue + pdblTotalpackagevalue) * Val(lblSaltax_Per.Text)) / 100
            End If
            rsCustomer.ResultSetClose()
        Else
            CalculateSalesTaxValue = ((pdblTotalBasicValue + pdblTotalExciseValue + pdblTotalpackagevalue) * Val(lblSaltax_Per.Text)) / 100
        End If
        If CalculateSalesTaxValue > 0 And CalculateSalesTaxValue < 1 And BlnSalesTax_Onerupee_Roundoff = True Then
            CalculateSalesTaxValue = 1
        End If
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
        Dim strsql As String
        Dim lclsGetTariffCode As ClsResultSetDB
        PrepareQueryForShowingExcise = ""
        If pblnTarrifCodeReq = True Then
            strsql = "SELECT Tariff_code FROM Item_Mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  Item_Code ='" & pstrItemCode & "'"
            lclsGetTariffCode = New ClsResultSetDB
            Call lclsGetTariffCode.GetResult(strsql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If lclsGetTariffCode.GetNoRows > 0 Then
                strsql = "SELECT Excise_duty FROM Tax_Tariff_Mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  Tariff_SubHead='" & lclsGetTariffCode.GetValue("Tariff_code") & "'"
                lclsGetTariffCode.ResultSetClose()
                lclsGetTariffCode = New ClsResultSetDB
                Call lclsGetTariffCode.GetResult(strsql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                If lclsGetTariffCode.GetNoRows > 0 Then
                    PrepareQueryForShowingExcise = " AND TxRt_Rate_No='" & lclsGetTariffCode.GetValue("Excise_duty") & "'"
                End If
            End If
            lclsGetTariffCode.ResultSetClose()
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
        Dim strsql As String
        Dim rsObj As New ADODB.Recordset
        On Error GoTo ErrHandler
        strsql = ""
        strsql = "SELECT " & pstrFieldName & " FROM Sales_Parameter WHERE UNIT_CODE='" + gstrUNITID + "'"
        If rsObj.State = 1 Then rsObj.Close()
        rsObj.Open(strsql, mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
        If rsObj.EOF Or rsObj.BOF Then
            MsgBox("No Data define in Sales_Parameter Table", MsgBoxStyle.Critical, "empower")
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
        On Error GoTo ErrHandler
        ldblTotalDispatchQuantity = 0
        ldblTotalScheduleQuantity = 0
        pstrDate = getDateForDB(pstrDate)
        mP_Connection.Execute("SET DATEFORMAT 'mdy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        If pstrMode = "ADD" Then
            strScheduleSql = "Select isnull(Schedule_Quantity,0) Schedule_Quantity,isnull(Despatch_Qty,0) Despatch_Qty from DailyMktSchedule WHERE UNIT_CODE='" + gstrUNITID + "' AND  Account_Code='" & pstrAccountCode & "' and "
            strScheduleSql = strScheduleSql & " datepart(yyyy,Trans_Date)='" & Year(CDate(pstrDate)) & "'"
            strScheduleSql = strScheduleSql & " and datepart(mm,Trans_Date)='" & Month(CDate(pstrDate)) & "'"
            strScheduleSql = strScheduleSql & " and Trans_Date <='" & pstrDate & "'"
            strScheduleSql = strScheduleSql & " and Cust_DrgNo = '" & pstrCustomerDrawNo & "' AND ITEM_CODE  = '" & pstrItemCode & "' and Status =1  ORDER BY Trans_Date DESC"
            If objRsForSchedule.State = 1 Then objRsForSchedule.Close()
            objRsForSchedule.CursorLocation = ADODB.CursorLocationEnum.adUseClient
            objRsForSchedule.Open(strScheduleSql, mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
            If objRsForSchedule.EOF Or objRsForSchedule.BOF Then
                GetTotalDispatchQuantityFromDailySchedule = -1
                mP_Connection.Execute("SET DATEFORMAT 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                objRsForSchedule.Close()
                Exit Function
            Else
                objRsForSchedule.MoveFirst()
                For lintLoopCounter = 1 To objRsForSchedule.RecordCount
                    ldblTotalScheduleQuantity = ldblTotalScheduleQuantity + Val(objRsForSchedule.Fields("Schedule_Quantity").Value)
                    ldblTotalDispatchQuantity = ldblTotalDispatchQuantity + Val(objRsForSchedule.Fields("Despatch_Qty").Value)
                    objRsForSchedule.MoveNext()
                Next
                GetTotalDispatchQuantityFromDailySchedule = Val(CStr(ldblTotalScheduleQuantity - ldblTotalDispatchQuantity))
                mP_Connection.Execute("SET DATEFORMAT 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                objRsForSchedule.Close()
                Exit Function
            End If
        Else
            strScheduleSql = "Select Schedule_Quantity,Despatch_Qty from DailyMktSchedule WHERE UNIT_CODE='" + gstrUNITID + "' AND  Account_Code='" & pstrAccountCode & "' and "
            strScheduleSql = strScheduleSql & " datepart(yyyy,Trans_Date)='" & Year(CDate(pstrDate)) & "'"
            strScheduleSql = strScheduleSql & " and datepart(mm,Trans_Date)='" & Month(CDate(pstrDate)) & "'"
            strScheduleSql = strScheduleSql & " and Trans_Date <='" & pstrDate & "'"
            strScheduleSql = strScheduleSql & " and Cust_DrgNo = '" & pstrCustomerDrawNo & "' AND ITEM_CODE  = '" & pstrItemCode & "' and Status =1  ORDER BY Trans_Date DESC" '''and Schedule_Flag =1   ( Now Not Consider)
            If objRsForSchedule.State = 1 Then objRsForSchedule.Close()
            objRsForSchedule.CursorLocation = ADODB.CursorLocationEnum.adUseClient
            objRsForSchedule.Open(strScheduleSql, mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
            If objRsForSchedule.EOF Or objRsForSchedule.BOF Then
                GetTotalDispatchQuantityFromDailySchedule = -1
                mP_Connection.Execute("SET DATEFORMAT 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                objRsForSchedule.Close()
                Exit Function
            Else
                objRsForSchedule.MoveFirst()
                For lintLoopCounter = 1 To objRsForSchedule.RecordCount
                    ldblTotalScheduleQuantity = ldblTotalScheduleQuantity + Val(objRsForSchedule.Fields("Schedule_Quantity").Value)
                    ldblTotalDispatchQuantity = ldblTotalDispatchQuantity + Val(objRsForSchedule.Fields("Despatch_Qty").Value)
                    objRsForSchedule.MoveNext()
                Next
                GetTotalDispatchQuantityFromDailySchedule = Val(CStr(ldblTotalScheduleQuantity - ldblTotalDispatchQuantity)) + Val(CStr(pdblPrevQty))
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
            strScheduleSql = "Select Schedule_Qty,isnull(Despatch_Qty,0) AS Despatch_qty  from MonthlyMktSchedule WHERE UNIT_CODE='" + gstrUNITID + "' AND  Account_Code='" & pstrAccountCode & "' and "
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
            strScheduleSql = "Select Schedule_Qty,isnull(Despatch_Qty,0)AS Despatch_qty from MonthlyMktSchedule WHERE UNIT_CODE='" + gstrUNITID + "' AND  Account_Code='" & pstrAccountCode & "' and "
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
    Public Function CheckcustorddtlQty(ByRef pstrMode As String, ByRef pstrItemCode As String, ByRef pstrDrgno As String, ByRef pdblQty As Double) As Boolean
        Dim rsCustOrdDtl As New ClsResultSetDB
        Dim rssaledtl As ClsResultSetDB
        Dim dblSaleQuantity As Double
        Dim strCustOrdDtl As String
        On Error GoTo ErrHandler
        If MULTIPLESO > 1 Then
            strCustOrdDtl = "Select a.openso,balance_Qty = a.order_qty - a.Despatch_qty from Cust_ord_dtl A , " + frmMKTTRN0020NEW.strTmpTable + " B where  A.UNIT_CODE ='" & gstrUNITID & "' AND "
            strCustOrdDtl = strCustOrdDtl & "a.Account_code ='" & txtCustCode.Text & "'" & " and a.Item_code ='"
            strCustOrdDtl = strCustOrdDtl & pstrItemCode & "' and a.cust_drgNo ='" & pstrDrgno
            strCustOrdDtl = strCustOrdDtl & "' and a.Authorized_flag = 1 and a.Active_Flag = 'A' and a.cust_ref = b.cust_ref"
            strCustOrdDtl = strCustOrdDtl & " and a.Amendment_no = b.Amendment_no"
        Else
            strCustOrdDtl = "Select openso,balance_Qty = order_qty - Despatch_qty from Cust_ord_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  "
            strCustOrdDtl = strCustOrdDtl & "Account_code ='" & txtCustCode.Text & "'" & " and Item_code ='"
            strCustOrdDtl = strCustOrdDtl & pstrItemCode & "' and cust_drgNo ='" & pstrDrgno
            strCustOrdDtl = strCustOrdDtl & "' and Authorized_flag = 1 and Active_Flag = 'A' and cust_ref = '" & txtRefNo.Text
            strCustOrdDtl = strCustOrdDtl & "' and Amendment_no = '" & txtAmendNo.Text & "'"
        End If
        rsCustOrdDtl.GetResult(strCustOrdDtl)
        If rsCustOrdDtl.GetValue("OpenSO") = True Then
            CheckcustorddtlQty = True
        Else
            Select Case pstrMode
                Case "ADD"
                    If Val(rsCustOrdDtl.GetValue("Balance_Qty")) < pdblQty Then
                        MsgBox("Balance Quantity available in SO for Customer Part code [ " & pstrDrgno & "] is " & Val(rsCustOrdDtl.GetValue("Balance_Qty")) & ".", MsgBoxStyle.Information, "empower")
                        CheckcustorddtlQty = False
                    Else
                        CheckcustorddtlQty = True
                    End If
                Case "EDIT"
                    rssaledtl = New ClsResultSetDB
                    rssaledtl.GetResult("Select Sales_Quantity from Sales_Dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  doc_no = " & txtChallanNo.Text & " and item_code = '" & pstrItemCode & "' and cust_ITem_code = '" & pstrDrgno & "'")
                    dblSaleQuantity = rssaledtl.GetValue("Sales_Quantity")
                    rssaledtl.ResultSetClose()
                    If (rsCustOrdDtl.GetValue("Balance_Qty")) < pdblQty Then
                        MsgBox("Balance Quantity available in SO for Customer Part code [ " & pstrDrgno & "] is " & rsCustOrdDtl.GetValue("Balance_Qty") & ".", MsgBoxStyle.Information, "empower")
                        CheckcustorddtlQty = False
                    Else
                        CheckcustorddtlQty = True
                    End If
            End Select
        End If
        rsCustOrdDtl.ResultSetClose()
        rsCustOrdDtl = Nothing
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function Calculatepkg(ByVal pintRowNo As Short, ByVal blnRoundoff As Boolean) As Double
        '---------------------------------------------------------------------------------------
        'Name       :   Calculatepkg
        'Type       :   Function
        'Author     :   Tapan Jain
        'Arguments  :
        'Return     :
        'Purpose    :
        '---------------------------------------------------------------------------------------
        Dim ldblPkg_Per As Double
        Dim lintQty As Double
        On Error GoTo ErrHandler
        With SpChEntry
            .Row = pintRowNo
            .Col = 3
            .Col = 6
            ldblPkg_Per = Val(.Text)
            .Col = 5
            lintQty = Val(.Text)
            Calculatepkg = System.Math.Round(ldblPkg_Per * lintQty, 2)
        End With
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function CalculatecustsuppexciseValue(ByVal pintRowNo As Short, ByVal pintexciseroundoff As Short) As Double
        Dim ldblcustexcvalue As Double
        Dim ldblcustqty As Double
        Dim ldblCustMat As Double
        Dim ldblToolCost As Double
        Dim ldblTaxRate As Double
        Dim ldblPkg_Per As Double
        Dim lintQty As Double
        Dim StrItemCode As String
        Dim strCustDrgNo As String
        Dim rsGetCustRate As ClsResultSetDB
        Dim rsGetTaxRate As ClsResultSetDB
        Dim strTableSql As String
        Dim intLoopCounter As Integer
        Dim rsdb As ClsResultSetDB
        Dim mblnCSMItem As Boolean
        mblnCSMItem = False
        rsGetCustRate = New ClsResultSetDB
        ldblcustexcvalue = 0
        Dim RSCUSTSUPPED As ADODB.Recordset
        Dim i As Short
        On Error GoTo ErrHandler
        With SpChEntry
            .Row = pintRowNo
            .Col = 1
            StrItemCode = Trim(.Text)
            .Col = 2
            strCustDrgNo = Trim(.Text)
            rsdb = New ClsResultSetDB
            Call rsdb.GetResult("SELECT Customer_code,Finish_Item_Code,Item_Code,Cust_Drgno,Item_drgno,Description,Qty,Rate,Active_Flag,VALID_FROM,VALID_TO,GRIN_No,GRIN_Auth_Date FROM CUSTSUPPLIEDITEM_MST WHERE UNIT_CODE='" + gstrUNITID + "' AND  FINISH_ITEM_CODE = '" & StrItemCode.ToString.Trim & "' AND CUST_DRGNO =  '" & strCustDrgNo.ToString.Trim & "'  AND ACTIVE_FLAG = 1 ")
            If rsdb.RowCount > 0 Then
                mblnCSMItem = True
            End If
            rsdb.ResultSetClose()
            rsdb = Nothing
            .Col = 6
            ldblPkg_Per = Val(.Text)
            .Col = 5
            lintQty = Val(.Text)
            .Col = 4
            ldblCustMat = Val(.Text) / Val(ctlPerValue.Text)
            .Col = 7
            rsGetTaxRate = New ClsResultSetDB
            strTableSql = "SELECT TxRt_Percentage FROM Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  TxRt_Rate_No='" & Trim(.Text) & "'"
            rsGetTaxRate.GetResult(strTableSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If rsGetTaxRate.GetNoRows > 0 Then
                ldblTaxRate = rsGetTaxRate.GetValue("TxRt_Percentage")
            Else
                ldblTaxRate = 0
            End If
            rsGetTaxRate.ResultSetClose()
            If ldblCustMat > 0 Then
                If mblnCSM_Knockingoff_req = True And mblnCSMItem = True Then
                    ldblcustexcvalue = ldblCustMat * (ldblTaxRate / 100)
                    ldblcustexcvalue = ldblcustexcvalue + System.Math.Round(ldblcustexcvalue * Val(lblECSStax_Per.Text) / 100, 2) + System.Math.Round(ldblcustexcvalue * Val(lblSECSStax_Per.Text) / 100, 2)
                Else
                    ldblcustexcvalue = lintQty * ldblCustMat * (ldblTaxRate / 100)
                End If
                'CalculatecustsuppexciseValue = System.Math.Round(ldblcustexcvalue, 0)
                CalculatecustsuppexciseValue = System.Math.Round(ldblcustexcvalue, pintexciseroundoff)
            End If
        End With
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function CalculateECSSTaxValue(ByVal pdblTotalExciseValue As Double) As Double
        On Error GoTo ErrHandler
        CalculateECSSTaxValue = (pdblTotalExciseValue * Val(lblECSStax_Per.Text) / 100)
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub txtECSSTaxType_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtECSSTaxType.TextChanged
        If Len(txtECSSTaxType.Text) = 0 Then
            lblECSStax_Per.Text = "0.00"
        End If
    End Sub
    Private Sub txtECSSTaxType_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtECSSTaxType.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        If txtTCSTaxCode.Enabled Then
                            txtTCSTaxCode.Focus()
                        End If
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
            '------------------Satvir Handa------------------------
            If CheckExistanceOfFieldData((txtECSSTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", " (Tx_TaxeID='ECS') and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))") Then
                '------------------Satvir Handa------------------------
                lblECSStax_Per.Text = CStr(GetTaxRate((txtECSSTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " (Tx_TaxeID='ECS')"))
            Else
                MsgBox("Invalid Cess on ED, Press F1 for help.", MsgBoxStyle.Information, ResolveResString(100))
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
        Dim strHelp As String
        On Error GoTo ErrHandler
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                If Len(txtECSSTaxType.Text) = 0 Then 'To check if There is No Text Then Show All Help
                    'strHelp = ShowList(1, (txtECSSTaxType.MaxLength), "", "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND (Tx_TaxeID='ECS')")
                    '------------------Satvir Handa------------------------
                    strHelp = ShowList(1, (txtECSSTaxType.MaxLength), "", "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND (Tx_TaxeID='ECS')  and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))")
                    '------------------Satvir Handa------------------------
                    If strHelp = "-1" Then 'If No Record Exists In The Table
                        Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Exit Sub
                    Else
                        txtECSSTaxType.Text = strHelp
                    End If
                Else
                    '------------------Satvir Handa------------------------
                    strHelp = ShowList(1, (txtECSSTaxType.MaxLength), txtECSSTaxType.Text, "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND (Tx_TaxeID='ECS') and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))")
                    '------------------Satvir Handa------------------------
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
    Public Function ValidateScheduleQuantity() As Boolean
        '------------------------------------------------------------------
        'Name       :   Validate For Schedule Quantity
        'Type       :   Function
        'Author     :   Sourabh Khatri
        'Arguments  :
        'Return     :   True : If validation is Successfully Completed
        'Purpose    :
        '------------------------------------------------------------------
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
        '*********************************************************
        'Validation For Schedule Start From Here
        '*********************************************************
        ValidateScheduleQuantity = True
        If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            strInvoiceType = UCase(Trim(CmbInvType.Text))
            strInvoiceSubType = UCase(Trim(CmbInvSubType.Text))
        ElseIf CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
            rsChallanEntry = New ClsResultSetDB
            rsChallanEntry.GetResult("Select a.Description,a.Sub_Type_Description from SaleConf a (nolock) ,SalesChallan_Dtl b where A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" + gstrUNITID + "' AND  Doc_No = " & txtChallanNo.Text & " and a.Invoice_Type = b.Invoice_type and a.Sub_type = b.Sub_Category and a.Location_code = b.Location_code and (fin_start_date <= getdate() and fin_end_date >= getdate())")
            strInvoiceType = UCase(rsChallanEntry.GetValue("Description"))
            strInvoiceSubType = UCase(rsChallanEntry.GetValue("sub_type_Description"))
            rsChallanEntry.ResultSetClose()
        End If
        Dim strMakeDate As String
        '10778579 changes done by Abhinav
        'If ((UCase(Trim(strInvoiceType)) = "NORMAL INVOICE") And (UCase(Trim(strInvoiceSubType)) = "FINISHED GOODS" Or (UCase(Trim(strInvoiceSubType)) = "TRADING GOODS"))) Or (UCase(Trim(strInvoiceType)) = "JOBWORK INVOICE") Or (UCase(Trim(strInvoiceType)) = "EXPORT INVOICE") Then
        If ((UCase(Trim(strInvoiceType)) = "NORMAL INVOICE") And (UCase(Trim(strInvoiceSubType)) = "FINISHED GOODS" Or (UCase(Trim(strInvoiceSubType)) = "TRADING GOODS"))) Or (UCase(Trim(strInvoiceType)) = "JOBWORK INVOICE") Or (UCase(Trim(strInvoiceType)) = "EXPORT INVOICE") Or (UCase(Trim(strInvoiceType)) = "TRANSFER INVOICE") And (UCase(Trim(strInvoiceSubType)) = "FINISHED GOODS") Or (UCase(Trim(strInvoiceType)) = "INTER-DIVISION") And (UCase(Trim(strInvoiceSubType)) = "FINISHED GOODS") Then
            rsChallanEntry = New ClsResultSetDB
            Call rsChallanEntry.GetResult("Select DSWiseTracking From Sales_parameter WHERE UNIT_CODE='" + gstrUNITID + "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
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
                Call SpChEntry.GetText(15, intRwCount, VarDelete)
                If UCase(VarDelete) <> "D" Then
                    If CheckMeasurmentUnit(varItemCode, varItemQty, intRwCount) = False Then
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
                    ldblNetDispatchQty = GetTotalDispatchQuantityFromDailySchedule(Trim(txtCustCode.Text), Trim(varDrgNo), Trim(varItemCode), Trim(lblDateDes.Text), "ADD", 0)
                ElseIf CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT And UCase(VarDelete) <> "D" Then
                    If Len(Trim(varDrgNo)) > 0 Then
                        If UCase(VarDelete) = "A" Then
                            ldblNetDispatchQty = GetTotalDispatchQuantityFromDailySchedule(Trim(txtCustCode.Text), Trim(varDrgNo), Trim(varItemCode), Trim(lblDateDes.Text), "EDIT", 0)
                        Else
                            ldblNetDispatchQty = GetTotalDispatchQuantityFromDailySchedule(Trim(txtCustCode.Text), Trim(varDrgNo), Trim(varItemCode), Trim(lblDateDes.Text), "EDIT", mdblPrevQty(intRwCount - 1))
                        End If
                    End If
                End If
                If ldblNetDispatchQty <> -1 And UCase(VarDelete) <> "D" Then
                    If Len(Trim(varDrgNo)) > 0 Then
                        If Val(varItemQty) > Val(CStr(ldblNetDispatchQty)) Then
                            ValidateScheduleQuantity = False
                            MsgBox("Quantity should not be Greater than Schedule Quantity " & CStr(ldblNetDispatchQty) & " For Item Code " & varItemCode, MsgBoxStyle.Information, "eMPro")
                            With Me.SpChEntry
                                .Row = intRwCount : .Col = 5 : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                            End With
                            Exit Function
                        Else
                            ValidateScheduleQuantity = True
                            If blnDSTracking = False Then
                                If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD And Trim(UCase(VarDelete)) <> "D" Then
                                    rsMktDailySchedule = New ClsResultSetDB
                                    rsMktDailySchedule.GetResult("Select Schedule_quantity from DailyMktSchedule WHERE UNIT_CODE='" + gstrUNITID + "' AND Account_Code='" & Trim(txtCustCode.Text) & "' and  datepart(yyyy,Trans_Date)='" & Year(ConvertToDate(Trim(lblDateDes.Text))) & "' and datepart(mm,Trans_Date)='" & Month(ConvertToDate(Trim(lblDateDes.Text))) & "' and datepart(dd,Trans_Date)='" & VB.Day(ConvertToDate(Trim(lblDateDes.Text))) & "' and Cust_DrgNo ='" & Trim(varDrgNo) & "'and Item_code = '" & varItemCode & "' and Status =1 ")
                                    If rsMktDailySchedule.GetNoRows > 0 Then
                                        mstrUpdDispatchSql = Trim(mstrUpdDispatchSql) & "Update DailyMktSchedule set Despatch_qty ="
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & "isnull(Despatch_Qty,0) + (" & Val(varItemQty) & ")"
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Account_Code='" & Trim(txtCustCode.Text) & "' and "
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & " datepart(yyyy,Trans_Date)='" & Year(ConvertToDate(Trim(lblDateDes.Text))) & "'"
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & " and datepart(mm,Trans_Date)='" & Month(ConvertToDate(Trim(lblDateDes.Text))) & "'"
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & " and datepart(dd,Trans_Date)='" & VB.Day(ConvertToDate(Trim(lblDateDes.Text))) & "'"
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & " and Cust_DrgNo ='" & Trim(varDrgNo) & "'and Item_code = '" & varItemCode & "' and Status =1 " & vbCrLf
                                    Else
                                        mstrUpdDispatchSql = Trim(mstrUpdDispatchSql) & " Insert into dailymktschedule "
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & "(Account_Code,Trans_date,cust_drgno,"
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & "Schedule_Flag,Item_Code,Schedule_Quantity,Despatch_qty,"
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & "Status,Ent_UserId,Upd_UserId,Ent_dt,Upd_dt,"
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & "RevisionNo,UNIT_CODE) values ('" & Trim(txtCustCode.Text) & "',"
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & "'" & getDateForDB(dtpDateDesc.Value) & "', '" & Trim(varDrgNo)
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & "',1,'" & varItemCode & "',0," & Val(varItemQty) & ",1,'" & mP_User & "',"
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & "'" & mP_User & "',getdate(),getdate(),0,'" + gstrUNITID + "')" & vbCrLf
                                    End If
                                    rsMktDailySchedule.ResultSetClose()
                                ElseIf CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                                    If Trim(UCase(VarDelete)) <> "D" Then
                                        mstrUpdDispatchSql = Trim(mstrUpdDispatchSql) & "Update DailyMktSchedule set Despatch_qty ="
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & "isnull(Despatch_Qty,0) + (" & Val(varItemQty) & ") - (" & mdblPrevQty(intRwCount - 1) & ")"
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Account_Code='" & Trim(txtCustCode.Text) & "' and "
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & " datepart(yyyy,Trans_Date)='" & Year(ConvertToDate(Trim(lblDateDes.Text))) & "'"
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & " and datepart(mm,Trans_Date)='" & Month(ConvertToDate(Trim(lblDateDes.Text))) & "'"
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & " and datepart(dd,Trans_Date)='" & VB.Day(ConvertToDate(Trim(lblDateDes.Text))) & "'"
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & " and Cust_DrgNo ='" & Trim(varDrgNo) & "'and Item_code = '" & varItemCode & "' and Status =1 " & vbCrLf
                                    Else
                                        mstrUpdDispatchSql = Trim(mstrUpdDispatchSql) & "Update DailyMktSchedule set Despatch_qty ="
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & "isnull(Despatch_Qty,0)  - (" & mdblPrevQty(intRwCount - 1) & ")"
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Account_Code='" & Trim(txtCustCode.Text) & "' and "
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
                    If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD And UCase(VarDelete) <> "D" Then
                        ldblNetDispatchQty = GetTotalDispatchQuantityFromMonthlySchedule(Trim(txtCustCode.Text), Trim(varDrgNo), Trim(varItemCode), Trim(lblDateDes.Text), "ADD", 0)
                    ElseIf CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT And UCase(VarDelete) <> "D" Then
                        If Len(Trim(varDrgNo)) > 0 Then
                            If UCase(VarDelete) = "A" Then
                                ldblNetDispatchQty = GetTotalDispatchQuantityFromMonthlySchedule(Trim(txtCustCode.Text), Trim(varDrgNo), Trim(varItemCode), Trim(lblDateDes.Text), "EDIT", 0)
                            Else
                                ldblNetDispatchQty = GetTotalDispatchQuantityFromMonthlySchedule(Trim(txtCustCode.Text), Trim(varDrgNo), Trim(varItemCode), Trim(lblDateDes.Text), "EDIT", mdblPrevQty(intRwCount - 1))
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
                                If Val(CStr(Month(ConvertToDate(lblDateDes.Text)))) < 11 Then
                                    strMakeDate = Year(ConvertToDate(lblDateDes.Text)) & "0" & Month(ConvertToDate(lblDateDes.Text))
                                Else
                                    strMakeDate = Year(ConvertToDate(lblDateDes.Text)) & Month(ConvertToDate(lblDateDes.Text))
                                End If
                                If blnDSTracking = False Then
                                    If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD And UCase(VarDelete) <> "D" Then
                                        mstrUpdDispatchSql = Trim(mstrUpdDispatchSql) & "Update MonthlyMktSchedule set Despatch_qty ="
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & "isnull(Despatch_Qty,0) + (" & Val(varItemQty) & ")"
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Account_Code='" & Trim(txtCustCode.Text) & "' and "
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & " Year_Month=" & Val(Trim(strMakeDate)) & ""
                                        mstrUpdDispatchSql = mstrUpdDispatchSql & " and Cust_DrgNo ='" & Trim(varDrgNo) & "'and Item_code = '" & varItemCode & "' and status =1 " & vbCrLf
                                    ElseIf CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                                        If VarDelete = "A" Then
                                            mstrUpdDispatchSql = Trim(mstrUpdDispatchSql) & "Update MonthlyMktSchedule set Despatch_qty ="
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & "isnull(Despatch_Qty,0) + (" & Val(varItemQty) & ") "
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Account_Code='" & Trim(txtCustCode.Text) & "' and "
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & " Year_Month=" & Val(Trim(strMakeDate)) & ""
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & " and Cust_DrgNo ='" & Trim(varDrgNo) & "'and Item_code = '" & varItemCode & "' and status =1 " & vbCrLf
                                        ElseIf VarDelete = "D" Then
                                            mstrUpdDispatchSql = Trim(mstrUpdDispatchSql) & "Update MonthlyMktSchedule set Despatch_qty ="
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & "isnull(Despatch_Qty,0)  - (" & mdblPrevQty(intRwCount - 1) & ") "
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Account_Code='" & Trim(txtCustCode.Text) & "' and "
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & " Year_Month=" & Val(Trim(strMakeDate)) & ""
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & " and Cust_DrgNo ='" & Trim(varDrgNo) & "'and Item_code = '" & varItemCode & "' and status =1 " & vbCrLf
                                        Else
                                            mstrUpdDispatchSql = Trim(mstrUpdDispatchSql) & "Update MonthlyMktSchedule set Despatch_qty ="
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & "isnull(Despatch_Qty,0) + (" & Val(varItemQty) & ") - (" & mdblPrevQty(intRwCount - 1) & ") "
                                            mstrUpdDispatchSql = mstrUpdDispatchSql & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Account_Code='" & Trim(txtCustCode.Text) & "' and "
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
        Else 'To Check Decimal places for all type of invoices
            For intRwCount = 1 To SpChEntry.MaxRows
                varItemCode = Nothing
                Call SpChEntry.GetText(1, intRwCount, varItemCode)
                varItemQty = Nothing
                Call SpChEntry.GetText(5, intRwCount, varItemQty)
                VarDelete = Nothing
                Call SpChEntry.GetText(15, intRwCount, VarDelete)
                '****Delete Flag Check
                If UCase(VarDelete) <> "D" Then
                    If CheckMeasurmentUnit(varItemCode, varItemQty, intRwCount) = False Then
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
    Private Function CalculateSECSSTaxValue(ByVal pdblTotalExciseValue As Double) As Double
        '-----------------------------------------------------------------------------------
        'Created By      : Ashutosh Verma
        'Issue ID        :
        'Creation Date   : 06 Mar 2007
        'Function        : To Calculate New Tax SEcess
        '-----------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        CalculateSECSSTaxValue = (pdblTotalExciseValue * Val(lblSECSStax_Per.Text) / 100)
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function InvAgstBarCode() As Boolean
        '----------------------------------------------------------------------------
        'Author         :   Manoj Kr. Vaish
        'Function       :   Get the BarCodefor Invoice from sales_parameter
        'Comments       :   Date: 04 Feb 2007 ,Issue Id: 22303
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim strQry As String
        Dim rs As ClsResultSetDB
        InvAgstBarCode = False
        strQry = "Select isnull(BarCodeTrackingInInvoice,0) as BarCodeTrackingInInvoice from sales_parameter WHERE UNIT_CODE='" + gstrUNITID + "'"
        rs = New ClsResultSetDB
        If rs.GetResult(strQry) = False Then GoTo ErrHandler
        If rs.GetValue("BarCodeTrackingInInvoice") = "True" Then
            rs.ResultSetClose()
            strQry = "Select isnull(a.BarcodeTrackingAllowed,0) as BarcodeTrackingAllowed"
            strQry = strQry & " from SaleConf a (nolock) ,SalesChallan_Dtl b where A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" + gstrUNITID + "' AND Doc_No ='" & txtChallanNo.Text & "'"
            strQry = strQry & " and a.Invoice_Type = b.Invoice_type and a.Sub_type = b.Sub_Category and"
            strQry = strQry & " a.Location_Code = b.Location_Code And (Fin_Start_Date <= getDate() And Fin_End_Date >= getDate())"
            rs = New ClsResultSetDB
            Call rs.GetResult(strQry, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If rs.GetNoRows > 0 Then
                If rs.GetValue("BarcodeTrackingAllowed") = "True" Then
                    InvAgstBarCode = True
                Else
                    InvAgstBarCode = False
                End If
            End If
            rs.ResultSetClose()
        Else
            rs.ResultSetClose()
        End If
        Exit Function
ErrHandler:
        rs = Nothing
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Public Function BarCodeTracking(ByVal pstrInvNo As String, ByVal pstrMode As String) As Boolean
        '----------------------------------------------------------------------------
        'Author         :   Manoj Kr. Vaish
        'Argument       :   Invoice Numbers.
        'Return Value   :   True or False
        'Function       :   Update Bar_BondedStock while invoice editing,deleting & Locking
        'Comments       :   Date: 04 Feb 2007 ,Issue Id: 22303
        'Revised By     :   Manoj Kr Vaish
        'Revision Date  :   12 Feb 2009 Issue ID : eMpro-20090209-27201
        'History        :   Functionality of Raw Material Invoice through Bar Code
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim rsGetQty As ClsResultSetDB
        Dim rsGetBondedQty As ClsResultSetDB
        Dim strsql As String
        Dim blnQuantitymatch As Boolean
        BarCodeTracking = False
        mblnQuantityCheck = True
        Select Case pstrMode
            Case "DELETE"
                If mstrInvsubTypeDesc = "RAW MATERIAL" Or mstrInvsubTypeDesc = "INPUTS" Then
                    mstrupdateBarBondedStockFlag = "Delete from Bar_Invoice_Issue WHERE UNIT_CODE='" + gstrUNITID + "' AND  Issue_Misno='" & Trim(pstrInvNo) & "' and Invoice_Status IS NULL"
                Else
                    mstrupdateBarBondedStockFlag = "Update bar_BondedStock_Dtl set Status_Flag='D' WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_no='" & Trim(pstrInvNo) & "' and Status_Flag<>'L'"
                End If
                BarCodeTracking = True
            Case "EDIT"
                If mstrInvsubTypeDesc = "RAW MATERIAL" Or mstrInvsubTypeDesc = "INPUTS" Then
                    mstrupdateBarBondedStockFlag = "Delete from Bar_Invoice_Issue WHERE UNIT_CODE='" + gstrUNITID + "' AND  Issue_Misno='" & Trim(pstrInvNo) & "' and Invoice_Status IS NULL"
                Else
                    mstrupdateBarBondedStockFlag = "Update bar_BondedStock_Dtl set Status_Flag='E' WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_no='" & Trim(pstrInvNo) & "' and Status_Flag<>'L'"
                End If
                BarCodeTracking = True
            Case "LOCK"
                If mstrInvsubTypeDesc = "RAW MATERIAL" Or mstrInvsubTypeDesc = "INPUTS" Then
                    '**************************Check Barcode Tracking Flag for RM from Item Master********************************
                    strsql = "select A.item_code,B.Itm_itemalias,isnull(Sales_Quantity,0) as Sales_Qty"
                    strsql = strsql & " from sales_dtl A Inner join Item_mst B on A.item_code=B.item_code AND A.UNIT_CODE=B.UNIT_CODE "
                    strsql = strsql & " where A.UNIT_CODE='" + gstrUNITID + "' AND B.Barcode_tracking=1 and A.doc_no='" & Trim(pstrInvNo) & "'"
                    rsGetQty.GetResult(strsql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                    If rsGetQty.GetNoRows > 0 Then
                        rsGetQty.MoveFirst()
                        Do While Not rsGetQty.EOFRecord
                            strsql = "select Isnull(Sum(Convert(numeric(16,4),Issue_Qty)),0) as Issue_Qty from Bar_Invoice_Issue "
                            strsql = strsql & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Issue_misno='" & Trim(pstrInvNo) & "' and invoice_status is null and substring(Issue_partBarCode,1,8)='" & rsGetQty.GetValue("Itm_itemalias") & "'"
                            rsGetBondedQty.GetResult(strsql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                            If rsGetBondedQty.GetNoRows > 0 Then
                                If rsGetQty.GetValue("Sales_Qty") = rsGetBondedQty.GetValue("Issue_Qty") Then
                                    blnQuantitymatch = True
                                Else
                                    MsgBox("Issued Quantity is less than Invoice Quantity.", vbInformation, ResolveResString(100))
                                    mblnQuantityCheck = False
                                    Exit Function
                                End If
                                rsGetQty.MoveNext()
                            Else
                                MsgBox("No Items are issued for this invoice.", vbInformation, ResolveResString(100))
                                mblnQuantityCheck = False
                                Exit Function
                            End If
                        Loop
                    End If
                    '******************************Insert into bar_Issue and update Bar_crossReference**********************************
                    mstrupdateBarBondedStockQty = ""
                    mstrupdateBarBondedStockFlag = ""
                    If blnQuantitymatch = True Then
                        mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & "Insert into Bar_Issue(Issue_No,Issue_UserId,Issue_Unit,Issue_CurrDtTm,Issue_PTId,Issue_UId,"
                        mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & "Issue_PartBarcode,Issue_MISNo,Issue_LocNo,Issue_Qty,Issue_Date,Issue_Time,Doc_Type,Location_code,UNIT_CODE)"
                        mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & "Select Issue_No,Issue_UserId,Issue_Unit,Issue_CurrDtTm,Issue_PTId,Issue_UId,Issue_PartBarcode,"
                        mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & "Issue_MISNo,Issue_LocNo,Issue_Qty,Issue_Date,Issue_Time,Doc_Type,Location_code,UNIT_CODE "
                        mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & "from Bar_Invoice_Issue WHERE UNIT_CODE='" + gstrUNITID + "' AND  Issue_MISNo='" & Trim(pstrInvNo) & "' and Invoice_status is null" & vbCrLf
                        strsql = "select A.CRef_PacketNo,isnull(sum(A.CRef_BalQty),0)as BarQuantity,Isnull(sum(Convert(numeric(16,4),Issue_Qty)),0)as SalesQuantity "
                        strsql = strsql & "from Bar_CrossReference A,Bar_Invoice_Issue B WHERE A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" + gstrUNITID + "' AND A.CRef_PacketNo=substring(B.Issue_PartbarCode,9,len(CRef_PacketNo)) and "
                        strsql = strsql & "A.CRef_PartCode=substring(B.Issue_partBarCode,1,8) and A.CRef_Stage='B' and B.Issue_MISNO='" & Trim(pstrInvNo) & "' and Invoice_status is null group by A.CRef_PacketNo"
                        rsGetQty.GetResult(strsql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        If rsGetQty.GetNoRows > 0 Then
                            rsGetQty.MoveFirst()
                            Do While Not rsGetQty.EOFRecord
                                If rsGetQty.GetValue("BarQuantity") - rsGetQty.GetValue("SalesQuantity") > 0 Then
                                    mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & "update A set A.CRef_BalQty=A.CRef_BalQty-" & rsGetQty.GetValue("SalesQuantity") & ""
                                    mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & " from  Bar_CrossReference A,Bar_Invoice_Issue B"
                                    mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & " Where A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" + gstrUNITID + "' AND A.CRef_PartCode=substring(B.Issue_partBarCode,1,8) and A.CRef_Stage='B'"
                                    mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & " and B.Issue_MISNO='" & Trim(pstrInvNo) & "' and A.CRef_PacketNo='" & rsGetQty.GetValue("CRef_PacketNo") & "'" & vbCrLf
                                ElseIf rsGetQty.GetValue("BarQuantity") - rsGetQty.GetValue("SalesQuantity") = 0 Then
                                    mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & "update A set A.CRef_BalQty=A.CRef_BalQty-" & rsGetQty.GetValue("SalesQuantity") & ",A.CRef_Stage='I'"
                                    mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & " from Bar_CrossReference A,Bar_Invoice_Issue B"
                                    mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & " Where A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" + gstrUNITID + "' AND A.CRef_PartCode=substring(B.Issue_partBarCode,1,8) and A.CRef_Stage='B'"
                                    mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & " and B.Issue_MISNO='" & Trim(pstrInvNo) & "' and A.CRef_PacketNo='" & rsGetQty.GetValue("CRef_PacketNo") & "'" & vbCrLf
                                ElseIf rsGetQty.GetValue("BarQuantity") - rsGetQty.GetValue("SalesQuantity") < 0 Then
                                    MsgBox("Quantity is not available in this Packet [" & rsGetQty.GetValue("CRef_PacketNo") & "] against issued quantity.", vbInformation, ResolveResString(100))
                                    mblnQuantityCheck = False
                                    Exit Function
                                End If
                                rsGetQty.MoveNext()
                            Loop
                            mstrupdateBarBondedStockFlag = mstrupdateBarBondedStockFlag & "Update Bar_Invoice_Issue Set Issue_Misno='" & mInvNo & "',Invoice_Status=1 WHERE UNIT_CODE='" + gstrUNITID + "' AND  Issue_MisNo='" & Trim(pstrInvNo) & "'" & vbCrLf
                            mstrupdateBarBondedStockFlag = mstrupdateBarBondedStockFlag & "Update Bar_Issue Set Issue_Misno='" & mInvNo & "' WHERE UNIT_CODE='" + gstrUNITID + "' AND  Issue_MisNo='" & Trim(pstrInvNo) & "'" & vbCrLf
                            BarCodeTracking = True
                        End If
                    Else
                        BarCodeTracking = False
                    End If
                    rsGetBondedQty = Nothing
                    rsGetQty = Nothing
                Else
                    '**************************Check Picked Quantity Against Invocie Quantity********************************
                    strsql = "select A.item_code,B.Itm_itemalias,isnull(Sales_Quantity,0) as Sales_Qty"
                    strsql = strsql & " from sales_dtl A Inner join Item_mst B on A.item_code=B.item_code AND A.UNIT_CODE=B.UNIT_CODE "
                    strsql = strsql & " WHERE A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" + gstrUNITID + "' AND  B.Barcode_tracking=1 and A.doc_no='" & Trim(pstrInvNo) & "'"
                    rsGetQty = New ClsResultSetDB
                    rsGetQty.GetResult(strsql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                    If rsGetQty.GetNoRows > 0 Then
                        rsGetQty.MoveFirst()
                        Do While Not rsGetQty.EOFRecord
                            strsql = "select isnull(sum(Quantity),0) as BondedStock_Qty from bar_BondedStock_Dtl "
                            strsql = strsql & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  invoice_no='" & Trim(pstrInvNo) & "' and Status_Flag='W' and item_alias='" & rsGetQty.GetValue("Itm_itemalias") & "'"
                            rsGetBondedQty = New ClsResultSetDB
                            rsGetBondedQty.GetResult(strsql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                            If rsGetQty.GetValue("Sales_Qty") = rsGetBondedQty.GetValue("BondedStock_Qty") Then
                                blnQuantitymatch = True
                            Else
                                MsgBox("Picked Quantity is less than Invoice Quantity.", MsgBoxStyle.Information, ResolveResString(100))
                                mblnQuantityCheck = False
                                rsGetQty.ResultSetClose()
                                Exit Function
                            End If
                            rsGetQty.MoveNext()
                        Loop
                    End If
                    rsGetQty.ResultSetClose()
                    '******************************Update bar Bonded Stock**********************************
                    mstrupdateBarBondedStockQty = ""
                    If blnQuantitymatch = True Then
                        strsql = "select B.Box_label,isnull(sum(A.Quantity),0)as BarQuantity,isnull(sum(B.Quantity),0)as SalesQuantity from Bar_BondedStock A,Bar_BondedStock_Dtl B WHERE A.UNIT_CODE='" + gstrUNITID + "' AND  "
                        strsql = strsql & "A.Box_Label=B.Box_label and A.Status='B' and B.Status_Flag='W' and B.Invoice_No='" & Trim(pstrInvNo) & "' Group By B.Box_label"
                        rsGetQty = New ClsResultSetDB
                        rsGetQty.GetResult(strsql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        If rsGetQty.GetNoRows > 0 Then
                            rsGetQty.MoveFirst()
                            Do While Not rsGetQty.EOFRecord
                                If rsGetQty.GetValue("BarQuantity") - rsGetQty.GetValue("SalesQuantity") > 0 Then
                                    mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & "update A set A.Quantity=A.Quantity-" & rsGetQty.GetValue("SalesQuantity") & ""
                                    mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & " from bar_BondedStock A,bar_BondedStock_Dtl B"
                                    mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & " Where A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" + gstrUNITID + "' AND A.Box_label = B.Box_label and A.Status='B' and B.Status_Flag='W'"
                                    mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & " and B.Invoice_no='" & Trim(pstrInvNo) & "' and A.Box_label='" & rsGetQty.GetValue("Box_label") & "'" & vbCrLf
                                ElseIf rsGetQty.GetValue("BarQuantity") - rsGetQty.GetValue("SalesQuantity") = 0 Then
                                    mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & "update A set A.Quantity=A.Quantity-" & rsGetQty.GetValue("SalesQuantity") & ",A.Status='I'"
                                    mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & " from bar_BondedStock A,bar_BondedStock_Dtl B"
                                    mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & " Where A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" + gstrUNITID + "' AND A.Box_label = B.Box_label and A.Status='B' and B.Status_Flag='W'"
                                    mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & " and B.Invoice_no='" & Trim(pstrInvNo) & "' and A.Box_label='" & rsGetQty.GetValue("Box_label") & "'" & vbCrLf
                                ElseIf rsGetQty.GetValue("BarQuantity") - rsGetQty.GetValue("SalesQuantity") < 0 Then
                                    MsgBox("Quantity is not available in this Box [" & rsGetQty.GetValue("Box_label") & "] against picked quantity.", MsgBoxStyle.Information, ResolveResString(100))
                                    mblnQuantityCheck = False
                                    rsGetQty.ResultSetClose()
                                    Exit Function
                                End If
                                rsGetQty.MoveNext()
                            Loop
                            mstrupdateBarBondedStockFlag = "Update bar_BondedStock_Dtl set Status_Flag='L',Invoice_no='" & Trim(CStr(mInvNo)) & "' WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_no='" & Trim(pstrInvNo) & "' and Status_Flag='W'"
                            BarCodeTracking = True
                        End If
                        rsGetQty.ResultSetClose()
                    Else
                        BarCodeTracking = False
                    End If
                End If
        End Select
        Exit Function
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
        Dim rs As New ADODB.Recordset
        rs = New ADODB.Recordset
        rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        rs.Open(strField, mP_Connection, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText)
        If rs.RecordCount > 0 Then
            If IsDBNull(rs.Fields(0).Value) = False Then
                Find_Value = rs.Fields(0).Value
            Else
                Find_Value = ""
            End If
        Else
            Find_Value = ""
        End If
        rs.Close()
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub CmdGrpChEnt_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.UCbtndgrp.ButtonClickEventArgs) Handles CmdGrpChEnt.ButtonClick
        On Error GoTo ErrHandler
        Dim strSalesChallan As String
        Dim updateSalesChallan As String
        Dim strSalesDtl As String
        Dim strSalesDtlDelete As String
        Dim strCurrency As String
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
        Dim VarDelete As Object
        Dim PresQty As Object
        Dim rsSalesChallandtl As ClsResultSetDB
        Dim rsInvoiceType As ClsResultSetDB
        Dim intLoop As Short
        Dim strMakeDate As String
        Dim strBatch As String
        Dim strDeleteASNdtl As String
        Dim strRejInvdtl As String
        Dim strSql As String
        Dim blnIsCt2 As Boolean = False
        Dim strRemoveInvFromLoadingSlip As String
        Dim intLoopCounter As Short

        Select Case e.Button
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD
                Call EnableControls(True, Me, True)
                'added by priti on 16 march 2020 to add vehicle help box
                If mblnAllowTransporterfromMaster Then
                    txtVehNo.Enabled = True
                    cmdVehicleCodeHelp.Enabled = True
                    cmdVehicleCodeHelp.Visible = True
                Else
                    txtVehNo.Enabled = True
                    cmdVehicleCodeHelp.Visible = False
                    cmdVehicleCodeHelp.Enabled = False
                End If
                lblSaltax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                lblSurcharge_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                Call SelectChallanNoFromSalesChallanDtl(0)
                txtChallanNo.Enabled = False : txtChallanNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                CmdChallanNo.Enabled = False : txtChallanNo.Enabled = False
                txtRefNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : CmdRefNoHelp.Enabled = False
                lblLocCodeDes.Text = "" : lblCustCodeDes.Text = ""
                Me.SpChEntry.Enabled = True
                aedamnt.Text = 0
                aedamnt.Enabled = False : aedamnt.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                Cmbtrninvtype.Enabled = False : Cmbtrninvtype.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                If blnEOU_FLAG = False Then
                    'CmbInvType.SelectedIndex = 3
                    'If CmbInvSubType.Items.Count > 2 Then
                    '    CmbInvSubType.SelectedIndex = 2
                    'End If

                    'CmbTransType.SelectedIndex = 0
                    For intLoopCounter = 0 To CmbInvType.Items.Count - 1 'Selecting Normal Invoice as default
                        If UCase(Trim(ObsoleteManagement.GetItemString(CmbInvType, intLoopCounter))) = "NORMAL INVOICE" Then
                            CmbInvType.SelectedIndex = intLoopCounter
                            Exit For
                        End If
                    Next
                    For intLoopCounter = 0 To CmbInvSubType.Items.Count - 1 'Selecting Finished Goods as default
                        If UCase(Trim(ObsoleteManagement.GetItemString(CmbInvSubType, intLoopCounter))) = "FINISHED GOODS" Then
                            CmbInvSubType.SelectedIndex = intLoopCounter
                            Exit For
                        End If
                    Next
                    CmbTransType.SelectedIndex = 0

                Else
                    CmbInvType.SelectedIndex = 1 : CmbInvSubType.SelectedIndex = 2 : CmbTransType.SelectedIndex = 0
                End If
                With Me.SpChEntry
                    .MaxRows = 1
                    .Row = 1 : .Row2 = 1 : .Col = 1 : .Col2 = .MaxCols : .BlockMode = True : .Text = "" : .Lock = False : .BlockMode = False
                End With
                'If UCase(Trim(CmbInvType.Text)) = "NORMAL INVOICE" And UCase(Trim(CmbInvSubType.Text)) = "SCRAP" And gblnGSTUnit = False Then
                If gblnGSTUnit = True Then
                    txtTCSTaxCode.Enabled = True : txtTCSTaxCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : cmdHelpTCSTax.Enabled = True
                Else
                    txtTCSTaxCode.Enabled = False : txtTCSTaxCode.Text = "" : txtTCSTaxCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdHelpTCSTax.Enabled = False : lblTCSTaxPerDes.Text = "0.00"
                End If
                If UCase(Trim(CmbInvType.Text)) <> "NORMAL INVOICE" Then
                    txtRefNo.Enabled = False
                    txtRefNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    CmdRefNoHelp.Enabled = False
                Else
                    txtRefNo.Enabled = True
                    txtRefNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    CmdRefNoHelp.Enabled = True
                End If
                'In Add Mode Enable Combo Of Invoice Type and Inv. Sub type
                CmbInvType.Visible = True : CmbInvSubType.Visible = True
                lblInvSubType.Visible = True : lblInvType.Visible = True

                With dtpDateDesc
                    .Format = DateTimePickerFormat.Custom
                    .CustomFormat = gstrDateFormat
                    .Value = GetServerDate()
                    .Visible = False 'Don't Show DatePicker
                End With
                'Get Server Date
                lblDateDes.Text = dtpDateDesc.Text

                'Set The Column Length in Spread
                Call SetMaxLengthInSpread(0)
                'Set Cell Type In Spread
                Call ChangeCellTypeStaticText()
                lblRGPDes.Text = ""
                If CmbInvType.Enabled Then CmbInvType.Focus()
                If gblnGSTUnit = False Then
                    txtECSSTaxType.Enabled = True
                    lblECSStax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    CmdECSSTaxType.Enabled = True
                    txtSECSSTaxType.Enabled = True
                    lblSECSStax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    Command1.Enabled = True
                    lblCurrencyDes.Text = ""
                Else
                    txtECSSTaxType.Enabled = False
                    lblECSStax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    CmdECSSTaxType.Enabled = False
                    txtSECSSTaxType.Enabled = False
                    lblSECSStax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    Command1.Enabled = False
                    txtKKC.Enabled = False
                    txtKKC.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    txtSBC.Enabled = False
                    txtSBC.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    cmdSBC.Enabled = False
                    cmdkkccode.Enabled = False
                    lblCurrencyDes.Text = ""
                    txtServiceTaxType.Enabled = False
                    txtServiceTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    cmdServiceTaxType.Enabled = False
                    txtAddVAT.Enabled = False
                    txtAddVAT.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    CmdAddVat.Enabled = False
                End If
                'Call checktcsvalue(CmbInvType.Text, CmbInvSubType.Text)
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT
                Call EnableControls(False, Me)
                'added by priti on 16 march 2020 to add vehicle help box
                If mblnAllowTransporterfromMaster Then
                    txtVehNo.Enabled = True
                    cmdVehicleCodeHelp.Enabled = True
                    cmdVehicleCodeHelp.Visible = True
                Else
                    txtVehNo.Enabled = True
                    cmdVehicleCodeHelp.Visible = False
                    cmdVehicleCodeHelp.Enabled = False
                End If

                PaletteActiveInActive()
                txtDeliveryNoteNo.Enabled = True : txtDeliveryNoteNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                rsSalesChallandtl = New ClsResultSetDB
                rsSalesChallandtl.GetResult("select Invoice_type,Sub_Category from Saleschallan_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  doc_no = " & txtChallanNo.Text, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                If rsSalesChallandtl.GetValue("Invoice_type") <> "JOB" Then
                    ctlInsurance.Enabled = True
                    ctlInsurance.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                End If
                If (UCase(Trim(rsSalesChallandtl.GetValue("Invoice_type"))) = "INV") And (UCase(Trim(rsSalesChallandtl.GetValue("Sub_Category"))) = "L") Then
                    If gblnGSTUnit = False Then
                        txtTCSTaxCode.Enabled = True : txtTCSTaxCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : cmdHelpTCSTax.Enabled = True
                    End If
                Else
                    'txtTCSTaxCode.Enabled = False : txtTCSTaxCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdHelpTCSTax.Enabled = False
                    txtTCSTaxCode.Enabled = True : txtTCSTaxCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : cmdHelpTCSTax.Enabled = True
                End If

                If UCase(rsSalesChallandtl.GetValue("Invoice_type")) = "INV" Or UCase(CStr(rsSalesChallandtl.GetValue("Invoice_type") = "REJ")) Or UCase(CStr(rsSalesChallandtl.GetValue("Invoice_type") = "EXP")) Then
                    If gblnGSTUnit = False Then
                        txtSaleTaxType.Enabled = True : txtSaleTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        CmdSaleTaxType.Enabled = True
                        lblSaltax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        lblSurcharge_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    Else
                        txtSaleTaxType.Enabled = False : txtSaleTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        CmdSaleTaxType.Enabled = False
                        lblSaltax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        lblSurcharge_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    End If

                    If CBool(UCase(CStr(rsSalesChallandtl.GetValue("Invoice_type") = "REJ"))) Then
                    End If
                    ctlInsurance.Enabled = True
                    ctlInsurance.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                End If
                rsSalesChallandtl.ResultSetClose()
                txtFreight.Enabled = True
                txtFreight.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                If gblnGSTUnit = False Then
                    txtSurchargeTaxType.Enabled = True : txtSurchargeTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    ''10706455  
                    txtAddVAT.Enabled = True : txtAddVAT.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    ''10706455  
                    txtRemarks.Enabled = True : txtRemarks.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    cmdSurchargeTaxCode.Enabled = True
                    CmdAddVat.Enabled = True
                End If

                SpChEntry.Enabled = True
                'SpChEntry.Row = 1 : SpChEntry.Row2 = SpChEntry.MaxRows : SpChEntry.Col = 0 : SpChEntry.Col2 = SpChEntry.MaxCols
                'SpChEntry.BlockMode = True : SpChEntry.Lock = False : SpChEntry.BlockMode = False
                SpChEntry.Row = 1 : SpChEntry.Row2 = SpChEntry.MaxRows : SpChEntry.Col = 0 : SpChEntry.Col2 = 2
                SpChEntry.BlockMode = True : SpChEntry.Lock = False : SpChEntry.BlockMode = False
                SpChEntry.Row = 1 : SpChEntry.Row2 = SpChEntry.MaxRows : SpChEntry.Col = 4 : SpChEntry.Col2 = SpChEntry.MaxCols
                SpChEntry.BlockMode = True : SpChEntry.Lock = False : SpChEntry.BlockMode = False
                Call SetMaxLengthInSpread(0)
                Call ChangeCellTypeStaticText()
                If DataExist("SELECT PCA_CUSTOMER FROM CUSTOMER_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & Trim(txtCustCode.Text) & "' and PCA_CUSTOMER =1 ") = True Then
                    SpChEntry.Row = 1 : SpChEntry.Row2 = SpChEntry.MaxRows : SpChEntry.Col = 5 : SpChEntry.Col2 = 5 : SpChEntry.BlockMode = True : SpChEntry.Lock = True : SpChEntry.BlockMode = False
                End If
                If GetPaletteStatus(UCase(CmbInvType.Text), UCase(CmbInvSubType.Text), txtCustCode.Text.Trim()) Then
                    PaletteReadOnlyQty()
                End If
                ReDim mdblPrevQty(SpChEntry.MaxRows - 1) ' To get value of Quantity in Array for updation in despatch
                For intLoop = 1 To SpChEntry.MaxRows
                    Call SpChEntry.GetText(5, intLoop, mdblPrevQty(intLoop - 1))
                Next
                If ctlInsurance.Enabled Then ctlInsurance.Focus()
                If gblnGSTUnit = False Then
                    txtECSSTaxType.Enabled = True
                    txtECSSTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    lblECSStax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    CmdECSSTaxType.Enabled = True
                    Command1.Enabled = True
                    txtSECSSTaxType.Enabled = True
                    txtSECSSTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    lblSECSStax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                Else
                    txtECSSTaxType.Enabled = False
                    txtECSSTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    lblECSStax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    CmdECSSTaxType.Enabled = False
                    Command1.Enabled = False
                    txtSECSSTaxType.Enabled = False
                    txtSECSSTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    lblSECSStax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                End If

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

                        'VINOD'
                        If UCase(CmbInvType.Text) = "NORMAL INVOICE" Or UCase(CmbInvType.Text) = "TRANSFER INVOICE" Then
                            If UCase(CmbInvSubType.Text) = "ASSETS" Then
                                If ValidateGlobalToolItemQty() = False Then
                                    Exit Sub
                                End If
                            End If
                        End If
                        '
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
                        If GetPaletteStatus(CmbInvType.Text, CmbInvSubType.Text, txtCustCode.Text.Trim()) Then
                            If Not ValidatePalette() Then Exit Sub
                        End If
                        If Not SaveData("ADD") Then Exit Sub
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
                        rsInvoiceType = New ClsResultSetDB
                        rsInvoiceType.GetResult("select Invoice_type from Saleschallan_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  doc_no = " & txtChallanNo.Text)
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
                        If GetPaletteStatus(strInvType, strInvSubType, txtCustCode.Text.Trim()) Then
                            If Not ValidatePalette() Then Exit Sub
                        End If
                        If Not SaveData("EDIT") Then Exit Sub
                End Select
                Call ConfirmWindow(10049, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                Me.CmdGrpChEnt.Revert()
                Me.CmdGrpChEnt.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                gblnCancelUnload = False : gblnFormAddEdit = False
                Call EnableControls(False, Me)
                txtDeliveryNoteNo.Enabled = True : txtDeliveryNoteNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                lblSaltax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                lblSurcharge_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                SpChEntry.Enabled = True
                SpChEntry.Row = 1 : SpChEntry.Row2 = SpChEntry.MaxRows : SpChEntry.Col = 0 : SpChEntry.Col2 = SpChEntry.MaxCols
                SpChEntry.BlockMode = True : SpChEntry.Lock = True : SpChEntry.BlockMode = False
                '****In View Mode Disable Combo Of Invoice Type and Inv. Sub type
                CmbInvType.Visible = False : CmbInvSubType.Visible = False
                lblInvSubType.Visible = False : lblInvType.Visible = False
                txtLocationCode.Enabled = True : txtLocationCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                txtChallanNo.Enabled = True : txtChallanNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                CmdLocCodeHelp.Enabled = True : CmdChallanNo.Enabled = True 'Me.SpChEntry.Enabled = False
                lblDateDes.Text = dtpDateDesc.Text
                dtpDateDesc.Visible = False
                If txtLocationCode.Enabled Then txtLocationCode.Focus()
                txtECSSTaxType.Enabled = False
                txtECSSTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                lblECSStax_Per.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                CmdECSSTaxType.Enabled = False
                CreatePaletteDataTable()
                mP_Connection.Execute("if exists(select name from sysobjects where name = '" + frmMKTTRN0020NEW.strTmpTable + "') drop table " + frmMKTTRN0020NEW.strTmpTable, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL
                Call frmMKTTRN0009_SOUTH_KeyPress(Me, New System.Windows.Forms.KeyPressEventArgs(Chr(System.Windows.Forms.Keys.Escape)))
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE
                If DataExist("select doc_no from saleschallan_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  Doc_No =" & Trim(txtChallanNo.Text) & " and bill_flag =1 ") Then
                    MsgBox("Locked Invoice cannot be deleted ", MsgBoxStyle.OkOnly, "eMPro")
                    Exit Sub
                Else
                    If ConfirmWindow(10054, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                        mstrUpdDispatchSql = ""
                        If Val(CStr(Month(ConvertToDate(lblDateDes.Text)))) < 10 Then
                            strMakeDate = Year(ConvertToDate(lblDateDes.Text)) & "0" & Month(ConvertToDate(lblDateDes.Text))
                        Else
                            strMakeDate = Year(ConvertToDate(lblDateDes.Text)) & Month(ConvertToDate(lblDateDes.Text))
                        End If
                        For intLoopcount = 1 To SpChEntry.MaxRows
                            varDrgNo = Nothing
                            Call Me.SpChEntry.GetText(2, intLoopcount, varDrgNo)
                            varItemCode = Nothing
                            Call Me.SpChEntry.GetText(1, intLoopcount, varItemCode)
                            PresQty = Nothing
                            Call Me.SpChEntry.GetText(5, intLoopcount, PresQty)
                            mstrUpdDispatchSql = Trim(mstrUpdDispatchSql) & "Update DailyMktSchedule set Despatch_qty ="
                            mstrUpdDispatchSql = mstrUpdDispatchSql & "isnull(Despatch_Qty,0) -  " & Val(PresQty) & ",Schedule_flag =1 "
                            mstrUpdDispatchSql = mstrUpdDispatchSql & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Account_Code='" & Trim(txtCustCode.Text) & "' and "
                            mstrUpdDispatchSql = mstrUpdDispatchSql & " datepart(yyyy,Trans_Date)='" & Year(ConvertToDate(Trim(lblDateDes.Text))) & "'"
                            mstrUpdDispatchSql = mstrUpdDispatchSql & " and datepart(mm,Trans_Date)='" & Month(ConvertToDate(Trim(lblDateDes.Text))) & "'"
                            mstrUpdDispatchSql = mstrUpdDispatchSql & " and datepart(dd,Trans_Date)='" & VB.Day(ConvertToDate(Trim(lblDateDes.Text))) & "'"
                            mstrUpdDispatchSql = mstrUpdDispatchSql & " and Cust_DrgNo ='" & Trim(varDrgNo) & "' and Item_code = '" & varItemCode & "' and Status =1" & vbCrLf
                            mstrUpdDispatchSql = Trim(mstrUpdDispatchSql) & " Update MonthlyMktSchedule set Despatch_qty ="
                            mstrUpdDispatchSql = mstrUpdDispatchSql & "isnull(Despatch_Qty,0)  - " & Val(PresQty) & ",Schedule_flag =1 "
                            mstrUpdDispatchSql = mstrUpdDispatchSql & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Account_Code='" & Trim(txtCustCode.Text) & "' and "
                            mstrUpdDispatchSql = mstrUpdDispatchSql & " Year_Month=" & Val(Trim(strMakeDate)) & ""
                            mstrUpdDispatchSql = mstrUpdDispatchSql & " and Cust_DrgNo ='" & Trim(varDrgNo) & "' and Item_code = '" & varItemCode & "' and Status =1 " & vbCrLf
                        Next
                        Call DeleteRecords()
                        If mblnBatchTrack = True And UCase(CmbInvType.Text) <> "JOBWORK INVOICE" And UCase(CmbInvType.Text) <> "REJECTION INVOICE" Then
                            strBatch = " Delete from ItemBatch_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  Doc_no = " & Trim(Me.txtChallanNo.Text) & " and  from_location = '" & mstrLocationCode & "'"
                        End If
                        If mblnRejTracking = True Then
                            strRejInvdtl = " Delete from MKT_INVREJ_DTL WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_no = " & Trim(Me.txtChallanNo.Text) & " "
                        End If
                        mP_Connection.BeginTrans()
                        'issue id 10192547
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
                            MsgBox("Unable To delete  CT2 Invoice Knock Off Details.", MsgBoxStyle.Information, ResolveResString(100))
                            objCmd = Nothing
                            mP_Connection.RollbackTrans()
                            Exit Sub
                        End If
                        objCmd = Nothing
                        '10736222

                        'If InvAgstBarCode() = True And mstrFGDomestic = "1" Then
                        If InvAgstBarCode() = True And (mstrFGDomestic = "1" Or mstrFGDomestic = "2") Then
                            'issue id 10192547
                            If BarCodeTracking(Trim(txtChallanNo.Text), "DELETE") = True Then
                                mP_Connection.Execute(mstrupdateBarBondedStockFlag, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            End If
                        End If
                        If AllowASNTextFileGeneration(Trim(txtCustCode.Text)) = True Then
                            strDeleteASNdtl = "DELETE FROM MKT_ASN_INVDTL WHERE UNIT_CODE='" + gstrUNITID + "' AND  DOC_NO='" & Val(txtChallanNo.Text) & " '"
                        End If
                        mblnCSM_Knockingoff_req = DataExist("SELECT TOP 1 1 FROM CUSTOMER_MST WHERE CSM_FLAG = 1 AND CUSTOMER_CODE=( select account_code from saleschallan_dtl where unit_code='" & gstrUNITID & "' and doc_no=" & Val(txtChallanNo.Text) & ") and Unit_code='" & gstrUNITID & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))")
                        mblnSinglelinelevelso = DataExist("SELECT TOP 1 1 FROM CUSTOMER_MST WHERE SOUPLD_LINE_LEVEL_SALESORDER = 1 AND CUSTOMER_CODE=( select account_code from saleschallan_dtl where unit_code='" & gstrUNITID & "' and doc_no=" & Val(txtChallanNo.Text) & ") and Unit_code='" & gstrUNITID & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))")
                        If mblnCSM_Knockingoff_req Then
                            mP_Connection.Execute("IF EXISTS(SELECT TOP 1 1 FROM CSM_KNOCKOFF_DTL WHERE UNIT_CODE='" + gstrUNITID + "' AND  INV_NO = " & txtChallanNo.Text & ") DELETE FROM CSM_KNOCKOFF_DTL WHERE UNIT_CODE='" + gstrUNITID + "' AND  INV_NO = " & txtChallanNo.Text, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        End If
                        mP_Connection.Execute(strupSaleDtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        mP_Connection.Execute(strupSalechallan, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        If gstrUNITID <> "STH" Then
                            mP_Connection.Execute(strupSalechallanUpload, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        End If

                        mP_Connection.Execute(mstrUpdDispatchSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            mP_Connection.Execute("DELETE from PCA_SCHEDULE_INVOICE_KNOCKOFF WHERE UNIT_CODE='" + gstrUNITID + "' AND  INVOICE_NO= '" & txtChallanNo.Text & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            If Len(Trim(strBatch)) > 0 Then
                                mP_Connection.Execute(strBatch, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            End If
                            If Len(Trim(strRejInvdtl)) > 0 Then
                                mP_Connection.Execute(strRejInvdtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            End If
                            If Len(strDeleteASNdtl) > 0 Then
                                mP_Connection.Execute(strDeleteASNdtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            End If
                            mP_Connection.Execute("Delete Emp_InvoiceSOLinkage WHERE UNIT_CODE='" + gstrUNITID + "' AND  Doc_No =" & Trim(txtChallanNo.Text) & "", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            'ADDED BY VINOD FOR GLOBAL TOOL CHANGES'
                            UpdateGlobalToolInvoice(Trim(txtChallanNo.Text), "D", Me.txtCustCode.Text.Trim)
                            'END OF CHANGES

                            'Code Added By Shubhra To Remove Cancelled Invoice from ILVS
                            'Begin
                            strRemoveInvFromLoadingSlip = "Update Loadingslip set InvoiceNo = NULL, ACT_INV_NO = NULL" &
                            " where Unit_Code = '" & gstrUNITID & "' and InvoiceNo = " & Val(txtChallanNo.Text.Trim) & ""
                            mP_Connection.Execute(strRemoveInvFromLoadingSlip, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        'End

                        If GetPaletteStatus(UCase(strInvType), UCase(strInvSubType), txtCustCode.Text.Trim()) Then
                            Dim strPalatte As String = "DELETE FROM INVOICE_PALETTE_DTL WHERE UNIT_CODE='" & gstrUNITID & "' AND TEMP_INVOICE_NO=" & Val(txtChallanNo.Text.Trim) & ""
                            mP_Connection.Execute(strPalatte, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        End If

                        If gstrUNITID = "STH" Then
                            Dim strshippingdtl As String =
                                    "UPDATE JD " &
                                    "SET SHI_EMPRO_STATUS = 0 " &
                                    "FROM JUMP_TO_EMPRO_SHIPPING JD " &
                                    "WHERE SHI_SITE_CODE = '" & gstrUNITID & "' " &
                                    "AND EXISTS ( " &
                                    "    SELECT 1 " &
                                    "    FROM DELIVERY_MKT_ACKN_HISTORY DM " &
                                    "    WHERE DM.PORECEIVEDNO = JD.SHI_REFERENCE_NUMBER " &
                                    "    AND DM.DELIVERYNUMBER = JD.SHI_DELIVERY_NUMBER " &
                                    "    AND DM.DOC_NO = " & Val(txtChallanNo.Text.Trim) &
                                    ")"
                            mP_Connection.Execute(strshippingdtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        End If

                        mP_Connection.CommitTrans()
                            Call EnableControls(False, Me, True)
                            txtLocationCode.Enabled = True
                            txtLocationCode.BackColor = System.Drawing.Color.White
                            CmdLocCodeHelp.Enabled = True
                            txtChallanNo.Enabled = True
                            txtChallanNo.BackColor = System.Drawing.Color.White
                            CmdChallanNo.Enabled = True
                            CreatePaletteDataTable()
                            Exit Sub
                        Else
                            Exit Sub
                    End If
                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                If mblnCSM_Knockingoff_req And Len(txtChallanNo.Text.Trim) > 0 Then
                    mP_Connection.Execute("IF NOT EXISTS(SELECT TOP 1 1 FROM CSM_KNOCKOFF_DTL CS INNER JOIN SALES_DTL SC on (SC.UNIT_CODE=CS.UNIT_CODE AND SC.DOC_NO=CS.INV_NO ) AND SC.UNIT_CODE='" + gstrUNITID + "' AND  CS.INV_NO = " & txtChallanNo.Text & ") DELETE FROM CSM_KNOCKOFF_DTL WHERE UNIT_CODE='" + gstrUNITID + "' AND  INV_NO = " & txtChallanNo.Text, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                End If

                Me.Close()
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
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
        lblDateDes.Text = dtpDateDesc.Text
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
    Private Sub txtFreight_KeyPress(ByVal Sender As Object, ByVal e As UCActXCtl.UCctlFloat.KeyPressEventArgs) Handles txtFreight.KeyPress
        On Error GoTo ErrHandler
        Select Case e.KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        If (CmbInvType.Text = "SAMPLE INVOICE") Or (CmbInvType.Text = "TRANSFER INVOICE") Or (CmbInvType.Text = "INTER-DIVISION") Or (CmbInvType.Text = "JOBWORK INVOICE") Then
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
    Private Sub ctlPerValue_KeyPress(ByVal Sender As Object, ByVal e As UCActXCtl.UCctlFloat.KeyPressEventArgs) Handles ctlPerValue.KeyPress
        On Error GoTo ErrHandler
        Select Case e.KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        With Me.SpChEntry
                            txtRemarks.Focus()
                        End With
                End Select
            Case 39, 34, 96
                e.KeyAscii = 0
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub SpChEntry_KeyDownEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyDownEvent) Handles SpChEntry.KeyDownEvent
        Dim strHelp As String
        Dim strCondition As String
        Dim StrItemCode As String
        Dim strpartcode As String
        Dim ADD_MODE As String
        Dim EDIT_MODE As String
        Dim dblqty As Double
        If CmdGrpChEnt.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            If e.keyCode = System.Windows.Forms.Keys.F1 And SpChEntry.ActiveCol = 9 And gblnGSTUnit = False Then
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
                        strHelp = ShowList(1, 6, "", "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND Tx_TaxeID='CVD'")
                        If strHelp = "-1" Then 'If No Record Exists In The Table
                            Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                            Exit Sub
                        Else
                            .Text = strHelp
                        End If
                    End If
                End With
            ElseIf e.keyCode = System.Windows.Forms.Keys.F1 And SpChEntry.ActiveCol = 10 And gblnGSTUnit = False Then
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
                'pras'
                '23 aug 2022
            ElseIf e.keyCode = System.Windows.Forms.Keys.F1 And SpChEntry.ActiveCol <= 5 And gblnGSTUnit = True And DataExist("SELECT PCA_CUSTOMER FROM CUSTOMER_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & Trim(txtCustCode.Text) & "' and PCA_CUSTOMER =1 ") = True Then
                With SpChEntry
                    .Row = .ActiveRow
                    .Col = 2
                    strpartcode = .Text.Trim
                    frmPCAPackage.PartCode = strpartcode
                End With
                If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                    frmPCAPackage.MODE = "ADD_MODE"
                End If
                If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                    frmPCAPackage.MODE = "EDIT_MODE"
                    frmPCAPackage.challan_no = txtChallanNo.Text
                End If
                frmPCAPackage.ShowDialog()
                With SpChEntry
                    If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD And CmbInvType.Text = "NORMAL INVOICE" And CmbInvSubType.Text = "FINISHED GOODS" And DataExist("SELECT PCA_CUSTOMER FROM CUSTOMER_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & Trim(txtCustCode.Text) & "' and PCA_CUSTOMER =1 ") = True Then
                        .Row = .ActiveRow
                        .Col = 1
                        StrItemCode = Trim(.Text)
                        dblqty = CDbl(Find_Value("Select SUM(QUANTITY) QUANTITY from TMP_PCA_ITEMSELECTION_PACKAGECODE WHERE UNIT_CODE='" + gstrUNITID + "' AND  IP_ADDRESS='" & gstrIpaddressWinSck & "' and Part_code='" & strpartcode & "'"))
                        Call .SetText(5, .ActiveRow, dblqty)
                    End If
                    If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT And DataExist("SELECT PCA_CUSTOMER FROM CUSTOMER_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & Trim(txtCustCode.Text) & "' and PCA_CUSTOMER =1 ") = True Then
                        .Row = .ActiveRow
                        .Col = 1
                        StrItemCode = Trim(.Text)
                        dblqty = CDbl(Find_Value("Select SUM(QUANTITY) QUANTITY from TMP_PCA_ITEMSELECTION_PACKAGECODE WHERE UNIT_CODE='" + gstrUNITID + "' AND  IP_ADDRESS='" & gstrIpaddressWinSck & "' and Part_code='" & strpartcode & "'"))
                        Call .SetText(5, .ActiveRow, dblqty)
                    End If
                End With
                'With SpChEntry
                '    .Row = .ActiveRow
                '    .Col = .ActiveCol

                '    .Col = 1
                '    StrItemCode = Trim(.Text)


                '    .Col = 2
                '    strpartcode = Trim(.Text)

                '    If Len(Trim(.Text)) = 0 Then 'To check if There is No Text Then Show All Help
                '        strHelp = ShowList(1, 6, "", "PACKAGECODE", "QUANTITY", "TMP_PCA_ITEMSELECTION_PACKAGECODE", "AND IP_ADDRESS='" & gstrIpaddressWinSck & "' and PART_CODE='" & strpartcode & "' and item_code='" & StrItemCode & "'")
                '        If strHelp = "-1" Then 'If No Record Exists In The Table
                '            Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                '            Exit Sub
                '        Else
                '            .Text = strHelp
                '        End If
                '    Else
                '        'To Display All Possible Help Starting With Text in TextField
                '        strHelp = ShowList(1, 6, "", "PACKAGECODE", "QUANTITY", "TMP_PCA_ITEMSELECTION_PACKAGECODE", "AND IP_ADDRESS='" & gstrIpaddressWinSck & "' and PART_CODE='" & strpartcode & "' and item_code='" & StrItemCode & "'")
                '        If strHelp = "-1" Then 'If No Record Exists In The Table
                '            Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                '            Exit Sub
                '        Else
                '            .Text = strHelp
                '        End If
                '    End If
                'End With
                ''pras'
                '23 aug 2022
            ElseIf e.keyCode = System.Windows.Forms.Keys.F1 And SpChEntry.ActiveCol = 8 And gblnGSTUnit = False Then
                    With SpChEntry
                        .Row = .ActiveRow
                        .Col = .ActiveCol
                        If Len(Trim(.Text)) = 0 Then 'To check if There is No Text Then Show All Help
                            strHelp = ShowList(1, 6, "", "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND Tx_TaxeID='AED'")
                            If strHelp = "-1" Then 'If No Record Exists In The Table
                                Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                                Exit Sub
                            Else
                                .Text = strHelp
                            End If
                        Else
                            'To Display All Possible Help Starting With Text in TextField
                            strHelp = ShowList(1, 6, Trim(.Text), "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND Tx_TaxeID='AED'")
                            If strHelp = "-1" Then 'If No Record Exists In The Table
                                Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                                Exit Sub
                            Else
                                .Text = strHelp
                            End If
                        End If
                    End With

                    'pras

                    '10808160
            ElseIf e.keyCode = System.Windows.Forms.Keys.F1 And SpChEntry.ActiveCol = 27 Then
                    With SpChEntry
                        .Row = .ActiveRow
                        .Col = 1
                        StrItemCode = Trim(.Text)

                        .Row = .ActiveRow
                        .Col = 2
                        strpartcode = Trim(.Text)

                        .Col = .ActiveCol
                        .Row = .ActiveRow

                        If Len(Trim(.Text)) = 0 Then 'To check if There is No Text Then Show All Help
                            strHelp = ShowList(1, 6, "", "MODEL_CODE", "ENDDATE", "BUDGETITEM_MST ", "  AND CUST_DRGNO='" & strpartcode & "' AND ITEM_CODE='" & StrItemCode & "' AND ENDDATE >= '" & GetServerDateNew().ToString("dd MMM yyyy") & "'")
                            If strHelp = "-1" Then 'If No Record Exists In The Table
                                Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                                Exit Sub
                            Else
                                .Text = strHelp
                            End If
                        Else
                            'To Display All Possible Help Starting With Text in TextField
                            strHelp = ShowList(1, 6, "", "MODEL_CODE", "ENDDATE", "BUDGETITEM_MST ", " AND CUST_DRGNO='" & strpartcode & "' AND ITEM_CODE='" & StrItemCode & "' AND ENDDATE >= '" & GetServerDateNew().ToString("dd MMM yyyy") & "'")
                            If strHelp = "-1" Then 'If No Record Exists In The Table
                                Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                                Exit Sub
                            Else
                                .Text = strHelp
                            End If
                        End If
                    End With
                    '10808160
                    'atn changes
            ElseIf e.keyCode = System.Windows.Forms.Keys.F1 And SpChEntry.ActiveCol = 28 Then
                    With SpChEntry
                        .Row = .ActiveRow
                        .Col = 28
                        StrItemCode = Trim(.Text)
                        .Col = .ActiveCol
                        strCondition = "and companycode='" + mstrcurrentATNCode + "' AND  transferedTo='" & mstrcustcompcode & "' AND NOT EXISTS ("
                        strCondition += " SELECT TOP  1 1 FROM SALES_DTL WHERE UNIT_CODE =VW_FA_ATN.UNIT_CODE AND ATNNO = VW_FA_ATN.ATNCODE  )"
                        If Len(Trim(.Text)) = 0 Then 'To check if There is No Text Then Show All Help
                            strHelp = ShowList(1, 40, "", "ATNCODE", "AssetCode", "VW_FA_ATN", strCondition)
                            If strHelp = "-1" Then 'If No Record Exists In The Table
                                Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                                Exit Sub
                            Else
                                If checkforDuplicateATNNO(strHelp, .ActiveRow, .MaxRows) = False Then
                                    .Text = strHelp
                                End If

                            End If
                        Else
                            'To Display All Possible Help Starting With Text in TextField
                            strHelp = ShowList(1, 40, "", "ATNCODE", "AssetCode", "VW_FA_ATN", strCondition)
                            If strHelp = "-1" Then 'If No Record Exists In The Table
                                Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                                Exit Sub
                            Else
                                .Text = strHelp
                            End If
                        End If
                    End With

                    'atn changes
            ElseIf e.keyCode = System.Windows.Forms.Keys.F1 And SpChEntry.ActiveCol = 7 And gblnGSTUnit = False Then
                    With SpChEntry
                        .Row = .ActiveRow
                        .Col = 1
                        StrItemCode = Trim(.Text)
                        .Col = .ActiveCol
                        strCondition = "AND Tx_TaxeID='EXC' " & PrepareQueryForShowingExcise(False, StrItemCode)
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
    End Sub
    Private Sub SpChEntry_KeyPressEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyPressEvent) Handles SpChEntry.KeyPressEvent
        On Error GoTo ErrHandler
        Select Case e.keyAscii
            Case 39, 34, 96, 45
                e.keyAscii = 0
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub SpChEntry_KeyUpEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyUpEvent) Handles SpChEntry.KeyUpEvent
        Dim intRow As Short
        Dim intDelete As Short
        Dim intLoopcount As Short
        Dim intMaxLoop As Short
        Dim VarDelete As Object
        'GST CHANGES
        If gblnGSTUnit = True And SpChEntry.ActiveCol = 7 Then
            For intLoopcount = 1 To SpChEntry.MaxRows
                VarDelete = Nothing
                Call SpChEntry.GetText(7, intLoopcount, VarDelete)
                If VarDelete <> "" Then
                    MsgBox("Excise Tax not allowed", MsgBoxStyle.Critical, ResolveResString(100))
                    Call SpChEntry.SetText(7, intLoopcount, "")
                    Exit For
                End If
            Next
        End If
        'GST CHANGES

        If ((e.shift = 2) And (e.keyCode = System.Windows.Forms.Keys.D)) Then
            If CmdGrpChEnt.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                With SpChEntry
                    If .MaxRows > 1 Then
                        intRow = .ActiveRow : intMaxLoop = SpChEntry.MaxRows
                        For intLoopcount = 1 To intMaxLoop
                            If intLoopcount <> intRow Then
                                VarDelete = Nothing
                                Call .GetText(15, intLoopcount, VarDelete)
                                If UCase(VarDelete) = "D" Then
                                    intDelete = intDelete + 1
                                End If
                            End If
                        Next
                        If (intMaxLoop - intDelete) > 1 Then
                            Call .SetText(15, intRow, "D")
                            .Row = .ActiveRow : .Row2 = .ActiveRow : .BlockMode = True : .RowHidden = True : .BlockMode = False
                        End If
                    End If
                End With
            End If
        End If
    End Sub
    Private Sub SpChEntry_LeaveCell(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_LeaveCellEvent) Handles SpChEntry.LeaveCell
        Dim lstrReturnVal As String
        Dim strWhereClause As String
        Dim varItemCode As Object
        Dim varDrgNo As Object
        Dim rsCustPartDes As ClsResultSetDB
        lstrReturnVal = ""
        With SpChEntry
            varItemCode = Nothing
            Call .GetText(1, e.newRow, varItemCode)
            varDrgNo = Nothing
            Call .GetText(2, e.newRow, varDrgNo)
            rsCustPartDes = New ClsResultSetDB
            rsCustPartDes.GetResult("Select Drg_Desc from CustItem_Mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  account_code = '" & Trim(txtCustCode.Text) & "' and item_code = '" & Trim(varItemCode) & "' and cust_drgno = '" & Trim(varDrgNo) & "'")
            If rsCustPartDes.GetNoRows > 0 Then
                lblCustPartDesc.Text = rsCustPartDes.GetValue("Drg_Desc")
            Else
                lblCustPartDesc.Text = ""
            End If
            rsCustPartDes.ResultSetClose()
        End With
        'Call DisplayDetailsInSpread("INR")
        'Changes ends here
        If CmdGrpChEnt.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            With SpChEntry
                If e.col = 9 Then
                    .Col = 9
                    .Row = .ActiveRow
                    If Trim(.Text) <> "" Then
                        strWhereClause = " WHERE UNIT_CODE='" + gstrUNITID + "' AND  TxRt_Rate_No='" & Trim(.Text) & "' AND Tx_TaxeID='CVD'"
                        lstrReturnVal = SelectDataFromTable("TxRt_Rate_No", "Gen_TaxRate", strWhereClause)
                        If Len(lstrReturnVal) = 0 Then
                            .Text = ""
                            MsgBox("Invalid Tax Code", MsgBoxStyle.Critical, "empower")
                        End If
                    End If
                ElseIf e.col = 10 Then
                    .Col = 10
                    .Row = .ActiveRow
                    If Trim(.Text) <> "" Then
                        strWhereClause = " WHERE UNIT_CODE='" + gstrUNITID + "' AND  TxRt_Rate_No='" & Trim(.Text) & "' AND Tx_TaxeID='SAD'"
                        lstrReturnVal = SelectDataFromTable("TxRt_Rate_No", "Gen_TaxRate", strWhereClause)
                        If Len(lstrReturnVal) = 0 Then
                            .Text = ""
                            MsgBox("Invalid Tax Code", MsgBoxStyle.Critical, "empower")
                        End If
                    End If
                End If
            End With
        End If
    End Sub
    Private Sub lblCurrencyDes_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblCurrencyDes.TextChanged
        If Trim(lblCurrencyDes.Text) <> "" Then
            If Trim(lblCurrencyDes.Text) = Trim(gstrCURRENCYCODE) Then
                lblExchangeRateValue.Text = CStr(1.0#)
            Else
                If UCase(Trim(mstrInvoiceType)) = "INV" Or UCase(Trim(mstrInvoiceType)) = "SMP" Or UCase(Trim(mstrInvoiceType)) = "TRF" Or UCase(Trim(mstrInvoiceType)) = "JOB" Or UCase(Trim(mstrInvoiceType)) = "EXP" Or UCase(Trim(mstrInvoiceType)) = "SRC" Then
                    lblExchangeRateValue.Text = CStr(GetExchangeRate(lblCurrencyDes.Text, dtpDateDesc.Text, True))
                Else
                    lblExchangeRateValue.Text = CStr(GetExchangeRate(lblCurrencyDes.Text, dtpDateDesc.Text, False))
                End If
                If Val(Trim(lblExchangeRateValue.Text)) = 1 Then
                    MsgBox("Exchange Rate for " & Trim(lblCurrencyDes.Text) & " is not defined on " & dtpDateDesc.Text, MsgBoxStyle.Information, "empower")
                    lblExchangeRateValue.Text = ""
                End If
            End If
        Else
            lblExchangeRateValue.Text = ""
        End If
    End Sub
    Private Sub ctlPerValue_TextChanged(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlPerValue.TextChanged
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
                    .Row = 0 : .Col = 3
                    .Text = "Rate (Per " & Val(ctlPerValue.Text) & ")"
                    .Row = 0 : .Col = 4
                    .Text = "Cust Supp Mat. (Per " & Val(ctlPerValue.Text) & ")"
                    .Row = 0 : .Col = 16
                    .Text = "Tool Cost (Per " & Val(ctlPerValue.Text) & ")"
                    .Row = 0 : .Col = 11
                    .Text = "Others (Per " & Val(ctlPerValue.Text) & ")"
                    With SpChEntry
                        intMaxLoop = .MaxRows
                        For intLoopCounter = 1 To intMaxLoop
                            varDrgNo = Nothing
                            Call .GetText(2, intLoopCounter, varDrgNo)
                            varItemCode = Nothing
                            Call .GetText(1, intLoopCounter, varItemCode)
                            If (Len(Trim(CStr(varDrgNo))) > 0) And (Len(Trim(CStr(varItemCode))) > 0) Then
                                varRate = Nothing
                                Call .GetText(17, intLoopCounter, varRate)
                                varCustMtrl = Nothing
                                Call .GetText(18, intLoopCounter, varCustMtrl)
                                varToolCost = Nothing
                                Call .GetText(20, intLoopCounter, varToolCost)
                                varOthers = Nothing
                                Call .GetText(19, intLoopCounter, varOthers)
                                Call .SetText(3, intLoopCounter, varRate * CDbl(ctlPerValue.Text))
                                Call .SetText(4, intLoopCounter, Val(varCustMtrl) * CDbl(ctlPerValue.Text))
                                Call .SetText(16, intLoopCounter, Val(varToolCost) * CDbl(ctlPerValue.Text))
                                Call .SetText(11, intLoopCounter, Val(varOthers) * CDbl(ctlPerValue.Text))
                            End If
                        Next
                    End With
                Else
                    .Row = 0 : .Col = 3 : .Text = "Rate (Per Unit)"
                    .Row = 0 : .Col = 4 : .Text = "Cust Supp Mat. (Per Unit)"
                    .Row = 0 : .Col = 16 : .Text = "Tool Cost (Per Unit)"
                    .Row = 0 : .Col = 11 : .Text = "Others (Per Unit)"
                    With SpChEntry
                        intMaxLoop = .MaxRows
                        For intLoopCounter = 1 To intMaxLoop
                            varDrgNo = Nothing
                            Call .GetText(2, intLoopCounter, varDrgNo)
                            varItemCode = Nothing
                            Call .GetText(1, intLoopCounter, varItemCode)
                            If (Len(Trim(CStr(varDrgNo))) > 0) And (Len(Trim(CStr(varItemCode))) > 0) Then
                                varRate = Nothing
                                Call .GetText(17, intLoopCounter, varRate)
                                varCustMtrl = Nothing
                                Call .GetText(18, intLoopCounter, varCustMtrl)
                                varToolCost = Nothing
                                Call .GetText(20, intLoopCounter, varToolCost)
                                varOthers = Nothing
                                Call .GetText(19, intLoopCounter, varOthers)
                                Call .SetText(3, intLoopCounter, varRate)
                                Call .SetText(4, intLoopCounter, Val(varCustMtrl))
                                Call .SetText(16, intLoopCounter, Val(varToolCost))
                                Call .SetText(11, intLoopCounter, Val(varOthers))
                            End If
                        Next
                    End With
                End If
            End With
        End With
    End Sub
    Public Sub CheckBatchTrackingAllowed(ByVal pInvType As String, ByVal pInvSubType As String)
        '-----------------------------------------------------------------------------------
        'Created By      : Manoj Kr.Vaish
        'Issue ID        : eMpro-20090209-27201
        'Creation Date   : 11 Feb 2009
        'Procedure       : To Check BatchTrackingAllowed for Any Invoice Type
        '-----------------------------------------------------------------------------------
        Dim rsCheckSo As ClsResultSetDB
        Dim strSql As String
        On Error GoTo ErrHandler
        rsCheckSo = New ClsResultSetDB
        ' strSql = "select isnull(BatchTrackingAllowed,0) as BatchTrackingAllowed from saleconf WHERE UNIT_CODE='" + gstrUNITID + "' AND  description='" & CmbInvType.Text & "' and sub_type_description='" & CmbInvSubType.Text & "' and (fin_start_date <= getdate() and fin_end_date >= getdate())"
        strSql = "select isnull(BatchTrackingAllowed,0) as BatchTrackingAllowed from saleconf (nolock) WHERE UNIT_CODE='" + gstrUNITID + "' AND  description='" & Trim(pInvType) & "' and sub_type_description='" & Trim(pInvSubType) & "' and (fin_start_date <= getdate() and fin_end_date >= getdate())"
        rsCheckSo.GetResult(strSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsCheckSo.GetNoRows > 0 Then
            mblnBatchTrack = rsCheckSo.GetValue("BatchTrackingAllowed")
            mblnBatchTracking = rsCheckSo.GetValue("BatchTrackingAllowed")
            With SpChEntry
                If mblnBatchTrack = True And UCase(CmbInvType.Text) <> "JOBWORK INVOICE" And UCase(CmbInvType.Text) <> "REJECTION" Then
                    .Row = 0 : .Col = 21 : .ColHidden = False : .BlockMode = False
                Else
                    .Row = 0 : .Col = 21 : .ColHidden = True : .BlockMode = False
                End If
            End With
        Else
            With SpChEntry
                If mblnBatchTrack = True And UCase(CmbInvType.Text) <> "JOBWORK INVOICE" And UCase(CmbInvType.Text) <> "REJECTION INVOICE" Then
                    .Row = 0 : .Col = 21 : .ColHidden = True : .BlockMode = False
                End If
            End With
        End If
        rsCheckSo.ResultSetClose()
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
        '------------------------------------------------------------------------------------------
    End Sub
    Private Sub ctlFormHeader1_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlFormHeader1.Click
        On Error GoTo ErrHandler
        Call ShowHelp("HLPMKTTRN0009.HTM") 'HLPMKTTRN0005.htm
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Public Function ReturnCustomerLocation() As String
        On Error GoTo Errorhandler
        Dim rsObject As New ClsResultSetDB
        Call rsObject.GetResult("Select Cust_Location=isnull(Cust_Location,'') from Customer_mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  Customer_Code = '" & Trim(Me.txtCustCode.Text) & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsObject.RowCount > 0 Then
            ReturnCustomerLocation = rsObject.GetValue("Cust_Location")
        Else
            ReturnCustomerLocation = ""
        End If
        rsObject.ResultSetClose()
        Exit Function
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function AllowASNTextFileGeneration(ByVal pstraccountcode As String) As Boolean
        'Revised By     : Manoj Kr. Vaish
        'Revised On     : 13 May 2009
        'Arguments      : Account Code
        'Return Value   : True/False
        'Issue ID       : eMpro-20090513-31282
        'Reason         : Check ASNTextFileGeneration from Customer Master
        '--------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim strQry As String
        Dim Rs As ClsResultSetDB
        AllowASNTextFileGeneration = False
        If (UCase(Trim(CmbInvType.Text)) = "NORMAL INVOICE" And UCase(Trim(CmbInvSubType.Text)) = "FINISHED GOODS") Then
            strQry = "Select isnull(AllowASNTextGeneration,0) as AllowASNTextGeneration from customer_mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  Customer_Code='" & Trim(pstraccountcode) & "'"
            Rs = New ClsResultSetDB
            If Rs.GetResult(strQry) = False Then GoTo ErrHandler
            If Rs.GetValue("AllowASNTextGeneration") = "True" Then
                AllowASNTextFileGeneration = True
            Else
                AllowASNTextFileGeneration = False
            End If
            Rs.ResultSetClose()
            Rs = Nothing
        End If
        '10126648 STARTS
        If (UCase(Trim(CmbInvType.Text)) = "REJECTION" And UCase(Trim(CmbInvSubType.Text)) = "REJECTION") Then
            strQry = "Select isnull(AllowASNTextGeneration,0) as AllowASNTextGeneration from Vendor_mst  where Vendor_code='" & Trim(pstraccountcode) & "' and Unit_code='" & gstrUNITID & "'"

            Rs = New ClsResultSetDB
            If Rs.GetResult(strQry) = False Then GoTo ErrHandler

            If Rs.GetValue("AllowASNTextGeneration") = "True" Then
                AllowASNTextFileGeneration = True
            Else
                AllowASNTextFileGeneration = False
            End If

            Rs.ResultSetClose()
            Rs = Nothing
        End If
        '10126648: End
        Exit Function
ErrHandler:
        Rs = Nothing
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function Validate_CSMRate() As Boolean
        On Error GoTo ErrHandler
        Dim strQry As String
        Validate_CSMRate = False
        If CmbInvType.Text = "NORMAL INVOICE" And CmbInvSubType.Text = "FINISHED GOODS" And mblnCSM_Knockingoff_req = True Then
            strQry = "SELECT TOP 1 1 FROM CSM_KNOCKOFF_DTL WHERE UNIT_CODE='" + gstrUNITID + "' AND  INV_NO = " & txtChallanNo.Text
            If DataExist(strQry) = False Then
                MsgBox("Customer Supplied Material (CSM) Stock Not Available." & vbCrLf & _
                       "Invoice Can't Be Saved.", MsgBoxStyle.Information, ResolveResString(100))
                Exit Function
            Else
                Validate_CSMRate = True
            End If
            strQry = "SELECT TOP 1 1 FROM CSM_KNOCKOFF_DTL_EXCEPTION WHERE UNIT_CODE='" + gstrUNITID + "' AND  INV_NO = " & txtChallanNo.Text
            If DataExist(strQry) = True Then
                MsgBox("Unable To Knock Off Customer Supplied Material (CSM) Item." & vbCrLf & _
                       "Invoice Can't Be Saved.", MsgBoxStyle.Information, ResolveResString(100))
                Validate_CSMRate = False
            Else
                Validate_CSMRate = True
            End If
        Else
            Validate_CSMRate = True
        End If
        Exit Function
ErrHandler:
        'Rs = Nothing
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub SetControlsforASNDetails(ByVal pstraccountcode As String)
        'Revised By     : Manoj Kr. Vaish
        'Revised On     : 13 May 2009
        'Arguments      : Account Code
        'Issue ID       : eMpro-20090513-31282
        'Reason         : Set controls to capture additional ASN details
        '--------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        CmbTransType.Items.Clear()
        If AllowASNTextFileGeneration(pstraccountcode) = True Then
            txtPlantCode.Enabled = True
            txtPlantCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            txtActualReceivingLoc.Enabled = True
            txtActualReceivingLoc.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            'Add Transport Mode in combo box according to the FORD
            CmbTransType.Items.Add("M - Motor")
            CmbTransType.Items.Add("A - Air")
            CmbTransType.Items.Add("W - Inland Waterway")
            CmbTransType.Items.Add("H - Customer Pickup")
            CmbTransType.Items.Add("R - Rail")
            CmbTransType.Items.Add("S - Ocean")
            CmbTransType.Items.Add("O - Containerized Ocean")
            CmbTransType.Items.Add("C - Consolidation")
            CmbTransType.Items.Add("U - UPS")
            CmbTransType.Items.Add("E - Expedited Truck")
            CmbTransType.SelectedIndex = 0
            Call SelectDescriptionForField("Plant_Code", "Customer_Code", "Customer_Mst", txtPlantCode, (txtCustCode.Text))
            If txtPlantCode.Text.Length > 0 Then txtPlantCode.Enabled = False
        Else
            txtPlantCode.Text = String.Empty
            txtActualReceivingLoc.Text = String.Empty
            txtPlantCode.Enabled = False
            txtPlantCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            txtActualReceivingLoc.Enabled = False
            txtActualReceivingLoc.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            Call AddTransPortTypeToCombo()      'Add default transport types
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtPlantCode_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPlantCode.KeyPress
        'Revised By     : Manoj Kr. Vaish
        'Revised On     : 13 May 2009
        'Issue ID       : eMpro-20090513-31282
        'Reason         : Validate only for Plant Code
        '--------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim Keyascii As Short = Asc(e.KeyChar)
        Select Case Keyascii
            Case System.Windows.Forms.Keys.Return
                If txtActualReceivingLoc.Enabled = True Then txtActualReceivingLoc.Focus()
            Case 39, 34, 96
                Keyascii = 0
        End Select
        If Keyascii = 0 Then
            e.Handled = True
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    
    Private Sub Cmbtrninvtype_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Cmbtrninvtype.SelectedIndexChanged
        Select Case (Cmbtrninvtype.Text)
            Case "LOCAL"
                aedamnt.Enabled = False : aedamnt.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            Case "IMPORTED"
                aedamnt.Enabled = True : aedamnt.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        End Select
    End Sub

    Private Sub frmMKTTRN0009_SOUTH_LocationChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.LocationChanged

    End Sub

    Private Sub chkRejType_CheckStateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkRejType.CheckStateChanged
        On Error GoTo Errorhandler
        If chkRejType.CheckState = 0 Then
            chkRejType.Text = "GRN"
        Else
            chkRejType.Text = "LRN"
        End If
        txtRefNo.Text = ""
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Function GetDocumentDetail(ByRef strItem_code As String) As String
        On Error GoTo ErrHandler
        Dim strCompileString As String
        Dim vardoc_no As Object
        Dim varBatch_No As Object
        Dim varBatch_Date As Object
        Dim varBatchReq As Object
        Dim varQty As Object
        Dim intRow As Short
        Dim rsTmp As New ClsResultSetDB
        Dim strSql As String
        strSql = "Select REF_DOC_NO, Batch_No, Quantity from MKT_INVREJ_DTL WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_No=" & Trim(txtChallanNo.Text) & " and Item_code='" & strItem_code & "'"
        rsTmp.GetResult(strSql)
        strCompileString = ""
        If rsTmp.RowCount > 0 Then
            Do While Not rsTmp.EOFRecord
                strCompileString = strCompileString & rsTmp.GetValue("Ref_Doc_no") & "§" & rsTmp.GetValue("Batch_No") & "§§" & rsTmp.GetValue("Quantity") & "¶"
                rsTmp.MoveNext()
            Loop
        End If
        rsTmp.ResultSetClose()
        GetDocumentDetail = strCompileString
        Exit Function
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

    Sub AddMaxAllowedQuanity(ByRef intRow As Short)
        On Error GoTo ErrHandler
        Dim dblqty As Double
        Dim dblStock As Double
        Dim varItemCode As Object
        Dim varMaxQuanity As Object
        Dim strsaledtl As String
        varItemCode = Nothing
        SpChEntry.GetText(1, intRow, varItemCode)
        If Find_Value("Select top 1 REJ_TYPE from MKT_INVREJ_DTL  WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_No=" & txtChallanNo.Text) = "1" Then
            strsaledtl = "select MaxAllowedQty = SUM( ((a.Rejected_Quantity + a.excess_po_quantity) - (isnull(a.Despatch_Quantity,0) + isnull(a.Inspected_Quantity,0) + isnull(a.RGP_Quantity,0)))) from grn_Dtl a, grn_hdr b Where "
            strsaledtl = strsaledtl & " a.Doc_type = b.Doc_type and a.unit_code=b.unit_code and a.unit_code='" + gstrUNITID + "' And a.Doc_No = b.Doc_No and "
            strsaledtl = strsaledtl & " a.From_Location = b.From_Location and a.From_Location ='01R1'"
            strsaledtl = strsaledtl & " and a.Rejected_quantity > 0 and b.Vendor_code = '" & txtCustCode.Text
            strsaledtl = strsaledtl & "' and a.Doc_No in (" & Trim(txtRefNo.Text) & ") and a.Item_code = '" & varItemCode & "' AND ISNULL(b.GRN_Cancelled,0) = 0"
            strsaledtl = strsaledtl & " Group by a.Item_code "
            dblqty = Val(Find_Value(strsaledtl))
        ElseIf Find_Value("Select top 1 REJ_TYPE from MKT_INVREJ_DTL WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_No=" & txtChallanNo.Text) = "2" Then
            strsaledtl = "Select   MaxAllowedQty = Sum(rejected_Quantity) from LRN_HDR as a " & " Inner Join LRN_DTL as b on a.doc_No=b.doc_no and a.unit_code=b.unit_code and a.unit_code='" + gstrUNITID + "' and a.Doc_Type=b.doc_Type and a.from_Location=b.from_location " & " Where Authorized_Code Is Not Null " & " and a.Doc_No IN (" & Trim(txtRefNo.Text) & ") and ITem_code = '" & varItemCode & "'" & " Group by B.Item_Code "
            dblqty = Val(Find_Value(strsaledtl))
        End If
        dblStock = Val(Find_Value("Select Cur_Bal From ItemBal_Mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  Location_code='01J1' and Item_code='" & varItemCode & "'"))
        If dblStock < dblqty Then
            varMaxQuanity = dblStock
        Else
            varMaxQuanity = dblqty
        End If
        SpChEntry.SetText(22, intRow, varMaxQuanity)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Function MakeSQL_REJINVTRACKING(ByRef intRow As Short) As String
        On Error GoTo ErrHandler
        Dim strSql As String
        Dim varCompileData As Object
        Dim strallbatches() As String
        Dim strBatch As Object
        Dim strBatchDetail() As String
        Dim strBatDetail As Object
        Dim strLocation_Code As String
        Dim intRejType As Short
        Dim strItem_code As Object
        Dim strcustpartcode As Object
        Dim strDoc_No As String
        Dim strbatch_no As String
        Dim dblQuantity As String
        Dim strInvNo As String
        strLocation_Code = Trim(txtLocationCode.Text)
        strInvNo = txtChallanNo.Text
        Dim varQty As Object
        varQty = Nothing
        SpChEntry.GetText(5, intRow, varQty)
        varCompileData = Nothing
        SpChEntry.GetText(24, intRow, varCompileData)
        strItem_code = Nothing
        SpChEntry.GetText(1, intRow, strItem_code)
        strcustpartcode = Nothing
        SpChEntry.GetText(2, intRow, strcustpartcode)
        If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            '1-GRN;2-LRN
            If chkRejType.Text = "GRN" Then
                intRejType = 1
            Else
                intRejType = 2
            End If
            strSql = ""

        ElseIf CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
            intRejType = Val(Find_Value("Select REJ_TYPE from MKT_INVREJ_DTL WHERE UNIT_CODE='" + gstrUNITID + "' AND  INvoice_No=" & Trim(txtChallanNo.Text)))

            strSql = "Delete From MKT_INVREJ_DTL WHERE UNIT_CODE='" + gstrUNITID + "' AND  Item_code='" & strItem_code & "' and Invoice_No='" & Trim(txtChallanNo.Text) & "'" & vbCrLf
        End If

        If mblnBatchTracking = True Then
            If Len(Trim(varCompileData)) <> 0 Then

                strallbatches = Split(varCompileData, "¶")
                For Each strBatch In strallbatches

                    If Len(Trim(strBatch)) <> 0 Then

                        strBatchDetail = Split(strBatch, "§")
                        strDoc_No = strBatchDetail(0)
                        strbatch_no = strBatchDetail(1)
                        dblQuantity = CStr(CDbl(strBatchDetail(3)))


                        strSql = Trim(strSql) & "Insert Into  MKT_INVREJ_DTL" & " (Location_code, Invoice_No, REJ_Type, Item_Code, Cust_Part_Code, Ref_Doc_No, Batch_No, Quantity, Cancel_Flag, END_DT, END_USERID,UPD_DT, UPD_USERID,UNIT_CODE) " & " values ('" & strLocation_Code & "'," & strInvNo & "," & intRejType & ", '" & strItem_code & "', '" & strcustpartcode & "', " & strDoc_No & ",'" & strbatch_no & "', " & dblQuantity & ", 0, Getdate(), '" & mP_User & "',Getdate(), '" & mP_User & "','" + gstrUNITID + "') "
                    End If
                Next strBatch
            End If
        Else
            strSql = Trim(strSql) & "Insert Into  MKT_INVREJ_DTL" & " (Location_code, Invoice_No, REJ_Type, Item_Code, Cust_Part_Code, Ref_Doc_No, Batch_No, Quantity, Cancel_Flag, END_DT, END_USERID,UPD_DT, UPD_USERID,UNIT_CODE) " & " values ('" & strLocation_Code & "'," & strInvNo & "," & intRejType & ", '" & strItem_code & "', '" & strcustpartcode & "', " & Trim(txtRefNo.Text) & ",'', " & CDbl(varQty) & ", 0, Getdate(), '" & mP_User & "',Getdate(), '" & mP_User & "','" + gstrUNITID + "') "
        End If
        MakeSQL_REJINVTRACKING = strSql
        Exit Function ' This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

    Public Function CalculateDocumentQty(ByRef strCompileString As String) As Double
        On Error GoTo ErrHandler
        Dim strallbatches() As String
        Dim strBatch As Object
        Dim strBatchDetail() As String
        Dim strBatDetail As Object
        Dim dblQuantity As Double
        dblQuantity = 0
        If Len(Trim(strCompileString)) <> 0 Then
            strallbatches = Split(strCompileString, "¶")
            For Each strBatch In strallbatches

                If Len(Trim(strBatch)) <> 0 Then

                    strBatchDetail = Split(strBatch, "§")
                    dblQuantity = dblQuantity + Val(strBatchDetail(3))
                End If
            Next strBatch
        End If
        CalculateDocumentQty = dblQuantity
        Exit Function
ErrHandler:
        gobjError.RaiseError(Err.Number, err.Source, Err.Description, mp_connection)
    End Function
    Private Sub SplitBatchColumnsRejection(ByRef pstrBatchRecords() As String, ByVal mActiveRowNoforArrindex As Short)
        '-------------------------------------------------------------------------------------------
        'Created By     -   Sourabh Khatri
        'Description    -   Split Batch Details Information Received from Batch Details Form Column Wise
        '-------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim strBatchCols() As String
        Dim Intcounter As Short
        If UBound(mBatchDataRejction) < mActiveRowNoforArrindex Then 'If New Row is Added In Issue GRID
            ReDim Preserve mBatchDataRejction(mActiveRowNoforArrindex)
        End If
        For Intcounter = 0 To UBound(pstrBatchRecords) - 1
            strBatchCols = Split(pstrBatchRecords(Intcounter), "§")
            ReDim Preserve mBatchDataRejction(Me.SpChEntry.ActiveRow).Document_No(Intcounter + 1)
            mBatchDataRejction(Me.SpChEntry.ActiveRow).Document_No(Intcounter + 1) = strBatchCols(0)
            ReDim Preserve mBatchDataRejction(Me.SpChEntry.ActiveRow).Batch_No(Intcounter + 1)
            mBatchDataRejction(Me.SpChEntry.ActiveRow).Batch_No(Intcounter + 1) = strBatchCols(1)
            ReDim Preserve mBatchDataRejction(Me.SpChEntry.ActiveRow).Batch_Date(Intcounter + 1)
            mBatchDataRejction(Me.SpChEntry.ActiveRow).Batch_Date(Intcounter + 1) = CDate(strBatchCols(2))
            ReDim Preserve mBatchDataRejction(Me.SpChEntry.ActiveRow).Batch_Quantity(Intcounter + 1)
            mBatchDataRejction(Me.SpChEntry.ActiveRow).Batch_Quantity(Intcounter + 1) = CDbl(strBatchCols(3))

        Next
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub txtTCSTaxCode_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtTCSTaxCode.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpChEnt.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        
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

    Private Sub txtTCSTaxCode_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTCSTaxCode.KeyUp
        Dim KeyCode As Short = e.KeyCode
        Dim Shift As Short = e.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = 112 Then
            If cmdHelpTCSTax.Enabled Then Call cmdHelpTCSTax_Click(cmdHelpTCSTax, New System.EventArgs())
        End If
        Exit Sub

ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred

    End Sub

    Private Sub txtTCSTaxCode_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTCSTaxCode.TextChanged
        If Len(txtTCSTaxCode.Text) = 0 Then
            lblTCSTaxPerDes.Text = "0.00"
        End If
    End Sub

    Private Sub txtTCSTaxCode_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtTCSTaxCode.Validating
        Dim Cancel As Boolean = e.Cancel
        Dim strInvoiceType As String
        Dim rsChallanEntry As ClsResultSetDB
        On Error GoTo ErrHandler
        If Len(txtTCSTaxCode.Text) > 0 Then
            If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                strInvoiceType = UCase(Trim(CmbInvType.Text))
            ElseIf (CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT) Or (CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW) Then
                rsChallanEntry = New ClsResultSetDB
                rsChallanEntry.GetResult("Select a.Description,a.Sub_Type_Description from SaleConf a (nolock) ,SalesChallan_Dtl b where A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" + gstrUNITID + "' AND Doc_No = " & txtChallanNo.Text & " and a.Invoice_Type = b.Invoice_type and a.Sub_type = b.Sub_Category and a.Location_code = b.Location_code and (fin_start_date <= getdate() and fin_end_date >= getdate())")
                strInvoiceType = UCase(rsChallanEntry.GetValue("Description"))
                rsChallanEntry.ResultSetClose()
            End If
            If CheckExistanceOfFieldData((txtTCSTaxCode.Text), "TxRt_Rate_No", "Gen_TaxRate", " (Tx_TaxeID='TCS')  and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))") Then
                lblTCSTaxPerDes.Text = CStr(GetTaxRate((txtTCSTaxCode.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " (Tx_TaxeID='TCS')"))
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
        e.Cancel = Cancel
    End Sub

    Private Sub cmdHelpTCSTax_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdHelpTCSTax.Click
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
                    salechallan = "select b.Description, b.Sub_type_Description from SalesChallan_dtl a,saleconf b (nolock)  where doc_no = " & Trim(txtChallanNo.Text)
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
    Public Function CalculateTCSTax(ByRef pdblTotalValue As Double, ByRef pblnTCSRoundOFF As Boolean, ByRef pintTCSPer As Double) As Double
        Dim dblTCSTax As Double
        Dim strsql As String

        If pblnTCSRoundOFF = True Then
            dblTCSTax = System.Math.Round((pdblTotalValue * pintTCSPer) / 100, 0)
        Else
            dblTCSTax = System.Math.Round((pdblTotalValue * pintTCSPer) / 100, 2)
        End If
        'CalculateTCSTax = dblTCSTax
        strSql = "select dbo.UFN_HIGHERVALUE_TCSROUNDING(" & dblTCSTax & ",'" & gstrUNITID & "'  )"
        CalculateTCSTax = SqlConnectionclass.ExecuteScalar(strSql)

    End Function
#Region "Global Tool Invoice Changes"
    'ADDED BY VINOD FOR GLOBAL TOOL CHANGES
    Private Function ValidateGlobalToolItemQty() As Boolean
        Dim intRow As Integer
        Dim strItemCode As String
        Dim rs As New ClsResultSetDB
        Dim strQry As String
        Try
            With Me.SpChEntry
                For intRow = 1 To .MaxRows
                    .Row = intRow
                    .Col = 1
                    strItemCode = .Text.Trim
                    rs = New ClsResultSetDB
                    strQry = "SELECT TOP 1 1 FROM VW_GBL_TOOL_INV_ITEM WHERE ITEM_CODE='" & strItemCode & "' AND UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & txtCustCode.Text.Trim & "'"
                    rs.GetResult(strQry)
                    If Not rs.EOFRecord Then
                        .Row = intRow
                        .Col = 5
                        If Val(.Text) > 1 Then
                            MsgBox("Invoice Qty. of Global Tool Item [" & strItemCode & "] can not be more than 1.", MsgBoxStyle.Information, ResolveResString(100))
                            rs.ResultSetClose()
                            rs = Nothing
                            Return False
                        End If
                    End If
                    rs.ResultSetClose()
                Next
            End With
            Return True
        Catch ex As Exception
            Throw ex
        Finally
            rs = Nothing
        End Try
    End Function

    Private Function UpdateGlobalToolInvoice(ByVal intInvNo As Integer, ByVal strMode As String, ByVal strCustomerMode As String) As Boolean
        Dim Cmd As New ADODB.Command
        Try
            With Cmd
                .CommandText = "USP_UPDATE_GLOBAL_TOOL_INVOICENO"
                .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                .CommandTimeout = 0
                .let_ActiveConnection(mP_Connection)
                .Parameters.Append(.CreateParameter("@UNIT_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                .Parameters.Append(.CreateParameter("@INVOICE_NO", ADODB.DataTypeEnum.adBigInt, ADODB.ParameterDirectionEnum.adParamInput, , intInvNo))
                .Parameters.Append(.CreateParameter("@CUSTOMER_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, strCustomerMode))
                .Parameters.Append(.CreateParameter("@MODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 1, strMode))
                .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End With
            Return True
        Catch ex As Exception
            Throw ex
        Finally
            Cmd = Nothing
        End Try
    End Function

#End Region

    Private Sub lblTCSPer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblTCSPer.Click

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
        If Len(txtAddVAT.Text) > 0 Then
            If CheckExistanceOfFieldData((txtAddVAT.Text), "TxRt_Rate_No", "Gen_TaxRate", " (Tx_TaxeID='ADVAT' OR Tx_TaxeID='ADCST')  and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))") Then
                lblAddVAT.Text = CStr(GetTaxRate((txtAddVAT.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " (Tx_TaxeID='ADVAT' OR Tx_TaxeID='ADCST')"))
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

    Private Sub CmdAddVat_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmdAddVat.Click
        '10706455
        Dim strHelp As String
        On Error GoTo ErrHandler
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                If Len(txtAddVAT.Text) = 0 Then
                    strHelp = ShowList(1, (txtAddVAT.MaxLength), "", "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND (Tx_TaxeID in('ADVAT','ADCST'))")
                    If strHelp = "-1" Then
                        Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Exit Sub
                    Else
                        txtAddVAT.Text = strHelp
                    End If
                Else
                    strHelp = ShowList(1, (txtAddVAT.MaxLength), txtAddVAT.Text, "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND (Tx_TaxeID in('ADVAT','ADCST') )")
                    If strHelp = "-1" Then
                        Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Exit Sub
                    Else
                        txtAddVAT.Text = strHelp
                    End If
                End If
                Call txtAddVAT_Validating(txtAddVAT, New System.ComponentModel.CancelEventArgs(False))
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub

    Private Function CalculateAdditionalSalesTaxValue(ByVal pdblTotalBasicValue As Double, ByVal pdblTotalExciseValue As Double, ByRef pblnIncStax As Boolean, ByRef pdblInsurance As Double) As Double
        On Error GoTo ErrHandler
        Dim dbldiscountamount As Double
            dbldiscountamount = 0
        If pblnIncStax = True Then
                CalculateAdditionalSalesTaxValue = ((pdblTotalBasicValue + pdblTotalExciseValue + pdblInsurance - dbldiscountamount) * Val(lblAddVAT.Text)) / 100
        Else
                CalculateAdditionalSalesTaxValue = ((pdblTotalBasicValue + pdblTotalExciseValue - dbldiscountamount) * Val(lblAddVAT.Text)) / 100
        End If

        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

    Private Sub ctlFormHeader1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ctlFormHeader1.Load

    End Sub

    Private Sub TxtLRNO_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TxtLRNO.Validating
        Dim strSQL As String
        If TxtLRNO.Text.Trim = "" Then Exit Sub
        Try
            '10856126
            strSQL = "select dbo.UDF_ISDOCKCODE_INVOICE( '" & gstrUNITID & "','" & txtCustCode.Text.Trim & "','" & CmbInvType.Text.Trim.ToUpper & "','" & CmbInvSubType.Text.Trim.ToUpper & "' )"
            If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSQL)) = True Then
                strSQL = "select dbo.UDF_ISLORRYNO_DOCKCODE( '" & gstrUNITID & "','" & txtCustCode.Text.Trim & "','" & TxtLRNO.Text.Trim & "' )"
                If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSQL)) = True Then
                    MsgBox("LR NO is already exists, , Enter some other text.", MsgBoxStyle.Information, ResolveResString(100))
                    TxtLRNO.Text = ""
                    If TxtLRNO.Enabled Then TxtLRNO.Focus()
                End If
            End If
            '10856126
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
        End Try
    End Sub

    Private Sub cmdServiceTaxType_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdServiceTaxType.Click
        '10706455
        Dim strHelp As String
        On Error GoTo ErrHandler
        Select Case Me.CmdGrpChEnt.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                If Len(txtServiceTaxType.Text) = 0 Then
                    strHelp = ShowList(1, (txtServiceTaxType.MaxLength), "", "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND (Tx_TaxeID in('SRT'))")
                    If strHelp = "-1" Then
                        Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Exit Sub
                    Else
                        txtServiceTaxType.Text = strHelp
                    End If
                Else
                    strHelp = ShowList(1, (txtServiceTaxType.MaxLength), txtAddVAT.Text, "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND (Tx_TaxeID in('SRT') )")
                    If strHelp = "-1" Then
                        Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Exit Sub
                    Else
                        txtServiceTaxType.Text = strHelp
                    End If
                End If
                Call txtServiceTaxType_Validating(txtServiceTaxType, New System.ComponentModel.CancelEventArgs(False))
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub

    Private Sub txtServiceTaxType_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtServiceTaxType.Validating
        Dim Cancel As Boolean = e.Cancel
        Dim strInvoiceType As String
        Dim rsChallanEntry As ClsResultSetDB
        On Error GoTo ErrHandler
        If Len(txtServiceTaxType.Text) > 0 Then
            If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                strInvoiceType = UCase(Trim(CmbInvType.Text))
            ElseIf (CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT) Or (CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW) Then
                rsChallanEntry = New ClsResultSetDB
                rsChallanEntry.GetResult("Select a.Description,a.Sub_Type_Description from SaleConf a (nolock) ,SalesChallan_Dtl b where A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" + gstrUNITID + "' AND Doc_No = " & txtChallanNo.Text & " and a.Invoice_Type = b.Invoice_type and a.Sub_type = b.Sub_Category and a.Location_code = b.Location_code and (fin_start_date <= getdate() and fin_end_date >= getdate())")
                strInvoiceType = UCase(rsChallanEntry.GetValue("Description"))
                rsChallanEntry.ResultSetClose()
            End If
            '10869290
            If UCase(Trim(strInvoiceType)) = "JOBWORK INVOICE" Or (UCase(Trim(strInvoiceType)) = "SERVICE INVOICE") Then
                If CheckExistanceOfFieldData((txtServiceTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", " (Tx_TaxeID='SRT')  and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))") Then
                    lblServiceTax_Per.Text = CStr(GetTaxRate((txtServiceTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " (Tx_TaxeID='SRT' )"))
                Else
                    MsgBox("This Service tax type is not correct. Please press F1 for help", MsgBoxStyle.Information, ResolveResString(100))
                    Cancel = True
                    txtServiceTaxType.Text = ""
                    lblServiceTax_Per.Text = ""
                    If txtServiceTaxType.Enabled Then txtServiceTaxType.Focus()
                End If
            End If
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Function CalculateServiceTaxValue(ByVal pdblTotalAccessValue As Double) As Object
        On Error GoTo ErrHandler

        CalculateServiceTaxValue = ((pdblTotalAccessValue) * Val(lblServiceTax_Per.Text)) / 100
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function CalculateSBCTaxValue(ByVal pdblTotalAccessValue As Double) As Object

        On Error GoTo ErrHandler
        CalculateSBCTaxValue = ((pdblTotalAccessValue) * Val(lblSBC.Text)) / 100
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function CalculatekkcTaxValue(ByVal pdblTotalAccessValue As Double) As Object

        On Error GoTo ErrHandler
        CalculatekkcTaxValue = ((pdblTotalAccessValue) * Val(lblKKC.Text)) / 100
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub txtServiceTaxType_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtServiceTaxType.TextChanged
        On Error GoTo ErrHandler
        If Len(txtServiceTaxType.Text) = 0 Then
            lblServiceTax_Per.Text = "0.00"
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtServiceTaxType_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtServiceTaxType.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F1 Then
            Call cmdServiceTaxType_Click(cmdServiceTaxType, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub

    Private Sub cmdSBC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSBC.Click
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
                    salechallan = "select b.Description, b.Sub_type_Description from SalesChallan_dtl a,saleconf b (nolock)  where doc_no = " & Trim(txtChallanNo.Text)
                    salechallan = salechallan & " and a.Location_code = b.Location_code and a.unit_code=b.unit_code and a.unit_code='" + gstrUNITID + "' and a.Invoice_type = b.invoice_type and a.sub_category = b.Sub_type"
                    rssalechallan.GetResult(salechallan, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                    If rssalechallan.GetNoRows > 0 Then
                        rssalechallan.MoveFirst()
                        strInvoiceType = rssalechallan.GetValue("Description")
                    End If
                    rssalechallan.ResultSetClose()
                Else
                    strInvoiceType = CmbInvType.Text
                End If
                If Len(Trim(Me.txtSBC.Text)) = 0 Then 'To check if There is No Text Then Show All Help
                    If ((UCase(strInvoiceType) = "JOBWORK INVOICE") Or (UCase(strInvoiceType) = "SERVICE INVOICE")) Then
                        strHelp = ShowList(1, (txtSBC.MaxLength), txtSBC.Text, "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND (Tx_TaxeID='SBC')  and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))")
                    End If
                    If strHelp = "-1" Then 'If No Record Exists In The Table
                        Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Exit Sub
                    Else
                        txtSBC.Text = strHelp
                        lblSBC.Text = CStr(GetTaxRate((txtSBC.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " (Tx_TaxeID='SBC')"))
                    End If
                Else
                    'To Display All Possible Help Starting With Text in TextField
                    If UCase(strInvoiceType) = "JOBWORK INVOICE" Then
                        strHelp = ShowList(1, (txtSBC.MaxLength), txtSBC.Text, "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND (Tx_TaxeID='SBC')  and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))")
                    End If
                    If strHelp = "-1" Then 'If No Record Exists In The Table
                        Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Exit Sub
                    Else
                        txtSBC.Text = Trim(strHelp)
                    End If
                End If
                Call txtSBC_Validating(txtSBC, New System.ComponentModel.CancelEventArgs(False))
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub

    End Sub
    Private Sub txtSBC_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtSBC.Validating
        Dim Cancel As Boolean = e.Cancel
        Dim strInvoiceType As String
        Dim rsChallanEntry As ClsResultSetDB
        On Error GoTo ErrHandler
        If Len(txtSBC.Text) > 0 Then
            If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                strInvoiceType = UCase(Trim(CmbInvType.Text))
            ElseIf (CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT) Or (CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW) Then
                rsChallanEntry = New ClsResultSetDB
                rsChallanEntry.GetResult("Select a.Description,a.Sub_Type_Description from SaleConf a (nolock) ,SalesChallan_Dtl b where A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" + gstrUNITID + "' AND Doc_No = " & txtChallanNo.Text & " and a.Invoice_Type = b.Invoice_type and a.Sub_type = b.Sub_Category and a.Location_code = b.Location_code and (fin_start_date <= getdate() and fin_end_date >= getdate())")
                strInvoiceType = UCase(rsChallanEntry.GetValue("Description"))
                rsChallanEntry.ResultSetClose()
            End If
            If UCase(Trim(strInvoiceType)) = "JOBWORK INVOICE" Or (UCase(Trim(strInvoiceType)) = "SERVICE INVOICE") Then
                If CheckExistanceOfFieldData((txtSBC.Text), "TxRt_Rate_No", "Gen_TaxRate", " (Tx_TaxeID='SBC')  and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))") Then
                Else
                    MsgBox("This SBC tax type is not correct. Please press F1 for help", MsgBoxStyle.Information, ResolveResString(100))
                    Cancel = True
                    txtSBC.Text = ""
                    If txtSBC.Enabled Then txtSBC.Focus()
                End If
            End If
        End If
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        e.Cancel = Cancel
    End Sub
    Public Function checkforDuplicateATNNO(ByRef pstrATNNO As String, ByVal pintRow As Short, ByVal pintMaxRow As Short) As Boolean
        'ATN CHANGES 
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        Dim varATNNo As Object

        On Error GoTo ErrHandler

        intMaxLoop = pintMaxRow
        checkforDuplicateATNNO = False

        For intLoopCounter = 1 To intMaxLoop
            With SpChEntry
                If intLoopCounter <> pintRow Then
                    varATNNo = Nothing
                    Call .GetText(28, intLoopCounter, varATNNo)
                    If (Trim(varATNNo) = Trim(pstrATNNO)) Then
                        checkforDuplicateATNNO = True
                        MsgBox("ASN No. you have entered already exist in grid", MsgBoxStyle.Information, "eMPro")
                        Exit For
                    End If
                End If
            End With
        Next

        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

    Public Sub CheckATNRequired(ByVal pInvType As String, ByVal pInvSubType As String)

        Dim rsCheckATN As ClsResultSetDB
        Dim strSql As String
        On Error GoTo ErrHandler

        rsCheckATN = New ClsResultSetDB
        strSql = "select ATN_ENABLED from saleconf (nolock) where unit_code='" & gstrUNITID & "' and description='" & Trim(pInvType) & "' and sub_type_description='" & Trim(pInvSubType) & "' and (fin_start_date <= getdate() and fin_end_date >= getdate())"
        rsCheckATN.GetResult(strSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

        If rsCheckATN.GetNoRows > 0 Then
            If rsCheckATN.GetValue("ATN_ENABLED") = True Then
                mblnATN_invoicewise = True

            Else
                mblnATN_invoicewise = False
            End If
        End If
        rsCheckATN.ResultSetClose()

        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
        '------------------------------------------------------------------------------------------
    End Sub

    Private Sub cmdkkccode_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdkkccode.Click
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
                    salechallan = "select b.Description, b.Sub_type_Description from SalesChallan_dtl a,saleconf b (nolock)  where doc_no = " & Trim(txtChallanNo.Text)
                    salechallan = salechallan & " and a.Location_code = b.Location_code and a.unit_code=b.unit_code and a.unit_code='" + gstrUNITID + "' and a.Invoice_type = b.invoice_type and a.sub_category = b.Sub_type"
                    rssalechallan.GetResult(salechallan, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                    If rssalechallan.GetNoRows > 0 Then
                        rssalechallan.MoveFirst()
                        strInvoiceType = rssalechallan.GetValue("Description")
                    End If
                    rssalechallan.ResultSetClose()
                Else
                    strInvoiceType = CmbInvType.Text
                End If
                If Len(Trim(Me.txtKKC.Text)) = 0 Then 'To check if There is No Text Then Show All Help
                    If ((UCase(strInvoiceType) = "JOBWORK INVOICE") Or (UCase(strInvoiceType) = "SERVICE INVOICE")) Then
                        strHelp = ShowList(1, (txtKKC.MaxLength), txtKKC.Text, "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND (Tx_TaxeID='KKC')  and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))")
                    End If
                    If strHelp = "-1" Then 'If No Record Exists In The Table
                        Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Exit Sub
                    Else
                        txtKKC.Text = strHelp
                    End If
                Else
                    'To Display All Possible Help Starting With Text in TextField
                    If UCase(strInvoiceType) = "JOBWORK INVOICE" Then
                        strHelp = ShowList(1, (txtKKC.MaxLength), txtKKC.Text, "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND (Tx_TaxeID='KKC')  and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))")
                    End If
                    If strHelp = "-1" Then 'If No Record Exists In The Table
                        Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Exit Sub
                    Else
                        txtKKC.Text = Trim(strHelp)
                    End If
                End If
                Call txtKKC_Validating(txtKKC, New System.ComponentModel.CancelEventArgs(False))
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub

    Private Sub txtKKC_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtKKC.KeyDown
        Dim KeyCode As Short = e.KeyCode
        Dim Shift As Short = e.KeyData \ &H10000

        On Error GoTo ErrHandler

        If KeyCode = System.Windows.Forms.Keys.F1 Then
            Call cmdkkccode_Click(cmdkkccode, New System.EventArgs())
        End If
        Exit Sub

ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub

    Private Sub txtKKC_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtKKC.Validating
        Dim Cancel As Boolean = e.Cancel
        Dim strInvoiceType As String
        Dim rsChallanEntry As ClsResultSetDB
        On Error GoTo ErrHandler
        If Len(txtKKC.Text) > 0 Then
            If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                strInvoiceType = UCase(Trim(CmbInvType.Text))
            ElseIf (CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT) Or (CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW) Then
                rsChallanEntry = New ClsResultSetDB
                rsChallanEntry.GetResult("Select a.Description,a.Sub_Type_Description from SaleConf a (nolock) ,SalesChallan_Dtl b where A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" + gstrUNITID + "' AND Doc_No = " & txtChallanNo.Text & " and a.Invoice_Type = b.Invoice_type and a.Sub_type = b.Sub_Category and a.Location_code = b.Location_code and (fin_start_date <= getdate() and fin_end_date >= getdate())")
                strInvoiceType = UCase(rsChallanEntry.GetValue("Description"))
                rsChallanEntry.ResultSetClose()
            End If
            If UCase(Trim(strInvoiceType)) = "JOBWORK INVOICE" Or (UCase(Trim(strInvoiceType)) = "SERVICE INVOICE") Then
                If CheckExistanceOfFieldData((txtKKC.Text), "TxRt_Rate_No", "Gen_TaxRate", " (Tx_TaxeID='KKC')  and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))") Then
                    lblKKC.Text = CStr(GetTaxRate((txtKKC.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " (Tx_TaxeID='KKC' )"))
                Else
                    MsgBox("This KKC tax type is not correct. Please press F1 for help", MsgBoxStyle.Information, ResolveResString(100))
                    Cancel = True
                    txtKKC.Text = ""
                    lblKKC.Text = ""
                    If txtKKC.Enabled Then txtKKC.Focus()
                End If
            End If
        End If
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
EventExitSub:
        e.Cancel = Cancel
    End Sub
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

        With SpChEntry

            .Row = pintRowNo

            If UCase(pTaxType) = "CGST" Then
                .Col = 31
            ElseIf UCase(pTaxType) = "SGST" Then
                .Col = 32
            ElseIf UCase(pTaxType) = "UTGST" Then
                .Col = 33
            ElseIf UCase(pTaxType) = "IGST" Then
                .Col = 34
            Else
                .Col = 35
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


        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    '101375632 
    Private Sub btnPalette_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPalette.Click
        Dim items As New DataTable
        Dim dtTemp As New DataTable

        Try
            items.Columns.Add("ITEM_CODE", GetType(System.String))
            items.Columns.Add("UNIT_CODE", GetType(System.String))
            Dim dr As DataRow
            For i As Integer = 1 To SpChEntry.MaxRows
                SpChEntry.Row = i
                SpChEntry.Col = 15
                If Convert.ToString(SpChEntry.Text) = "D" Then Continue For
                dr = items.NewRow()
                SpChEntry.Row = i
                SpChEntry.Col = 1
                dr("ITEM_CODE") = Convert.ToString(SpChEntry.Text)
                dr("UNIT_CODE") = gstrUnitId
                items.Rows.Add(dr)
            Next
            Dim objPaletteHelp As New frmPaletteHelp(CmdGrpChEnt.Mode, txtCustCode.Text.Trim(), Val(txtChallanNo.Text.Trim()), items)
            If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                If dtPaletteItemQty Is Nothing OrElse dtPaletteItemQty.Rows.Count = 0 Then
                    Dim strSql As String = " SELECT PALETTE_LABEL,ITEM_CODE,QTY FROM INVOICE_PALETTE_DTL WHERE UNIT_CODE='" & gstrUnitId & "' AND TEMP_INVOICE_NO=" & Val(txtChallanNo.Text) & ""
                    dtTemp = SqlConnectionclass.GetDataTable(strSql)
                    If dtTemp IsNot Nothing AndAlso dtTemp.Rows.Count > 0 Then
                        Dim drow As DataRow
                        For Each drTemp As DataRow In dtTemp.Rows
                            drow = dtPaletteItemQty.NewRow()
                            drow("PALETTE_LABEL") = Convert.ToString(drTemp("PALETTE_LABEL"))
                            drow("ITEM_CODE") = Convert.ToString(drTemp("ITEM_CODE"))
                            drow("QTY") = Convert.ToInt32(drTemp("QTY"))
                            dtPaletteItemQty.Rows.Add(drow)
                        Next
                    End If
                End If
            End If

            objPaletteHelp.SetPalette = dtPaletteItemQty
            objPaletteHelp.ShowDialog()
            dtPaletteItemQty = objPaletteHelp.GetPalette
            If dtPaletteItemQty IsNot Nothing AndAlso dtPaletteItemQty.Rows.Count > 0 Then
                Dim result = (From item In dtPaletteItemQty.AsEnumerable() Group By ITEM_CODE = item.Field(Of String)("ITEM_CODE") Into g = Group Select New With {Key ITEM_CODE, .QTY = g.Sum(Function(r) r.Field(Of Integer)("QTY"))}).OrderBy(Function(tkey) tkey.ITEM_CODE).ToList()
                If result IsNot Nothing AndAlso result.Count > 0 Then
                    For i As Integer = 0 To result.Count - 1
                        For j As Integer = 1 To SpChEntry.MaxRows
                            SpChEntry.Row = j
                            SpChEntry.Col = 1
                            If Convert.ToString(result(i).ITEM_CODE) = Convert.ToString(SpChEntry.Text) Then
                                SpChEntry.Row = j
                                SpChEntry.Col = 5
                                SpChEntry.Text = result(i).QTY
                                Exit For
                            End If
                        Next
                    Next
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        Finally
            items.Dispose()
            dtTemp.Dispose()
        End Try
    End Sub
    Private Sub PaletteActiveInActive()
        CreatePaletteDataTable()
        If SpChEntry.MaxRows > 0 Then
            If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                btnPalette.Enabled = GetPaletteStatus(CmbInvType.Text, CmbInvSubType.Text, txtCustCode.Text.Trim())
            Else
                btnPalette.Enabled = GetPaletteStatus(strInvType, strInvSubType, txtCustCode.Text.Trim())
            End If
        Else
            btnPalette.Enabled = False
        End If
    End Sub
    Private Sub CreatePaletteDataTable()
        dtPaletteItemQty = New DataTable()
        dtPaletteItemQty.Columns.Add("PALETTE_LABEL", GetType(System.String))
        dtPaletteItemQty.Columns.Add("ITEM_CODE", GetType(System.String))
        dtPaletteItemQty.Columns.Add("QTY", GetType(System.Int32))
    End Sub
    Private Sub PaletteReadOnlyQty()
        If SpChEntry.MaxRows > 0 Then
            With SpChEntry
                For i As Integer = 1 To .MaxRows
                    .Row = i
                    .Col = 5
                    .Lock = True
                Next
            End With
        End If
    End Sub
    Public Sub checktcsvalue(ByVal pInvType As String, ByVal pInvSubType As String)
        Dim rsTCSReq As ClsResultSetDB
        Try
            rsTCSReq = New ClsResultSetDB
            rsTCSReq.GetResult("Select isnull(REQD_TCS,0) as REQD_TCS , TCSTXRT_TYPE from saleConf (nolock) Where UNIT_CODE='" + gstrUNITID + "' AND description ='" & Trim(pInvType) & "' and Sub_Type_Description='" & Trim(pInvSubType) & "' and  (fin_start_date <= getdate() and fin_end_date >= getdate())")
            If rsTCSReq.GetValue("REQD_TCS") = True Then
                txtTCSTaxCode.Enabled = True : txtTCSTaxCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : cmdHelpTCSTax.Enabled = True : txtTCSTaxCode.Text = rsTCSReq.GetValue("TCSTXRT_TYPE").ToString
                If CheckExistanceOfFieldData((txtTCSTaxCode.Text), "TxRt_Rate_No", "Gen_TaxRate", " (Tx_TaxeID='TCS')  and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))") Then
                    lblTCSTaxPerDes.Text = CStr(GetTaxRate((txtTCSTaxCode.Text), "TxRt_Rate_No", "Gen_TaxRate", "TxRt_Percentage", " (Tx_TaxeID='TCS')"))
                End If
            Else
                If (UCase(Trim(pInvType) = "NORMAL INVOICE") And (UCase(Trim(pInvSubType)) = "SCRAP")) Then
                    If gblnGSTUnit = False Then
                        txtTCSTaxCode.Enabled = True : txtTCSTaxCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : cmdHelpTCSTax.Enabled = True
                    End If
                Else
                    txtTCSTaxCode.Enabled = False : txtTCSTaxCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdHelpTCSTax.Enabled = False : txtTCSTaxCode.Text = ""
                End If

            End If
            rsTCSReq.ResultSetClose()

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Function ValidatePalette() As Boolean
        Dim dtValidateData As New DataTable
        Dim result As Boolean = True
        Try
            If dtPaletteItemQty IsNot Nothing AndAlso dtPaletteItemQty.Rows.Count > 0 Then
                Dim qry = (From item In dtPaletteItemQty.AsEnumerable() Group By ITEM_CODE = item.Field(Of String)("ITEM_CODE") Into g = Group Select New With {Key ITEM_CODE, .QTY = g.Sum(Function(r) r.Field(Of Integer)("QTY"))}).OrderBy(Function(tkey) tkey.ITEM_CODE).ToList()
                If qry IsNot Nothing AndAlso qry.Count > 0 Then
                    For i As Integer = 0 To qry.Count - 1
                        For j As Integer = 1 To SpChEntry.MaxRows
                            SpChEntry.Row = j
                            SpChEntry.Col = 1
                            If Convert.ToString(qry(i).ITEM_CODE) = Convert.ToString(SpChEntry.Text) Then
                                SpChEntry.Row = j
                                SpChEntry.Col = 5
                                If Val(SpChEntry.Text) <> Val(qry(i).QTY) Then
                                    MsgBox("Selected Palette QTY. is not equal to Invoice Item Qty. for Item Code : " & Convert.ToString(qry(i).ITEM_CODE), MsgBoxStyle.Information, "eMPro")
                                    result = False
                                    Exit For
                                End If
                            End If
                        Next
                        If Not result Then
                            Exit For
                        End If
                    Next
                End If
                If result Then
                    Dim strSql As String = String.Empty
                    Dim strPalette As New System.Text.StringBuilder("")
                    Dim strItems As New System.Text.StringBuilder("")

                    For i As Integer = 0 To dtPaletteItemQty.Rows.Count - 1
                        strPalette.Append("'" & Convert.ToString(dtPaletteItemQty.Rows(i)("PALETTE_LABEL")) & "',")
                        strItems.Append("'" & Convert.ToString(dtPaletteItemQty.Rows(i)("ITEM_CODE")) & "',")
                    Next
                    If CmdGrpChEnt.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                        strSql = "SELECT PALETTE_LABEL,ITEM_CODE FROM INVOICE_PALETTE_DTL WHERE UNIT_CODE='" & gstrUnitId & "' AND ITEM_CODE IN (" & strItems.ToString().TrimEnd(",") & ") AND PALETTE_LABEL IN (" & strPalette.ToString().TrimEnd(",") & ") AND TEMP_INVOICE_NO<> " & Val(txtChallanNo.Text.Trim()) & ""
                    Else
                        strSql = "SELECT PALETTE_LABEL,ITEM_CODE FROM INVOICE_PALETTE_DTL WHERE UNIT_CODE='" & gstrUnitId & "' AND ITEM_CODE IN (" & strItems.ToString().TrimEnd(",") & ") AND PALETTE_LABEL IN (" & strPalette.ToString().TrimEnd(",") & ")"
                    End If

                    dtValidateData = SqlConnectionclass.GetDataTable(strSql)
                    If dtValidateData IsNot Nothing AndAlso dtValidateData.Rows.Count > 0 Then
                        MsgBox("Selected Palettes have already used in another invoice.")
                        result = False
                    End If
                End If
            Else
                MsgBox("Please select Palette", MsgBoxStyle.Information, "eMPro")
                result = False
            End If
        Catch ex As Exception
            RaiseException(ex)
            result = False
        Finally
            dtValidateData.Dispose()
        End Try
        Return result
    End Function

    Private Sub cmdVehicleCodeHelp_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdVehicleCodeHelp.Click
        'select transporter_code as [Transporter Code],Transporter_name as [Transporter Name],Vehicle_Type as [Vehicle Type],vehicle_no as [Vehicle No] from vehicle_mst where active=1
        Dim strSql As String = ""
        Dim strVehicle As String = ""
        Dim varRetVal As Object
        On Error GoTo ErrHandler
        With txtVehNo
            If Len(.Text) = 0 Then

                'varRetVal = ShowList(1, (txtCustCode.MaxLength), "", "Customer_Code", "Cust_Name", "Customer_Mst", "  and Group_Customer=1 and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))")

                varRetVal = ShowList(1, .MaxLength, "", "vehicle_no", "Transporter_name", "vehicle_mst", " and active=1 ", "Help", "", "", 0, "transporter_code")
                If varRetVal = "-1" Then
                    Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                    .Text = ""
                Else
                    .Text = varRetVal

                End If
            Else
                varRetVal = ShowList(1, .MaxLength, , "vehicle_no", "Transporter_name", "vehicle_mst", " and active=1 ", "Help", "", "", 0, "transporter_code")
                If varRetVal = "-1" Then
                    Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                    .Text = ""
                Else
                    .Text = varRetVal

                End If
            End If
            .Focus()
        End With

        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
End Class
