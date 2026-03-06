Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.IO
Imports System.Data.SqlClient
Imports System.Drawing.Imaging
Imports System.Runtime.InteropServices
Friend Class frmMKTTRN0008_Hilex
    Inherits System.Windows.Forms.Form
    '===================================================================================
    ' (c) MothersonSumi INfotech & Design Ltd. All rights reserverd.
    ' File Name         :   FRMMKTTRN0008.frm
    ' Function          :   Used to Print & View Invoice deatails
    ' Created By        :   Nisha Rai
    ' Created On        :   09 May, 2001
    ' Revision History  :   Nisha Rai
    '21/09/2001 MARKED CHECKED BY BCs changed on version 3
    '03/10/2001 MARKED CHECKED BY BCs  FOR JOBWORK INVOICE changed on version 4
    '09/10/2001 changed on version 6 to make changes in case of checking from Daily/Monthly
    'Schedule having Status =1
    '09/01/2002 changed of Smiel Chennei to add CVD_PER,SVD_Per,Insurance
    '22/01/2002 changed for addSalesTax account_Code commented checkedout form no = 4013
    '28/01/2002 changed in case of Transfer invoice to allow to update in Received_dtl
    '15/01/2002 CHANGED FOR DOCUMENT NO. ON FORM NO. 4067
    '10/04/2002 100% EOU Changes & Delivery address Required Yes/No
    '19/04/2002 changed for opennin balance updation
    '24/04/2002 Round off account data
    '30/04/2001 Mod Function is not working Changes in Assigning value to Array
    'Reprinting of invoice
    '08/05/2002 SCRAP invoice Changes
    '27/05/02 Three Copies of invoice printing
    '29/05/02 Stock check & RG23 PLA no entry from front end
    '02/07/2002 Jobwork invoice Selected List
    '08/072002 *** to add one more insertion in Round off Account in save account Data
    '          *** to check if round off account is defined or not
    '10/07/2002 changed for Printed no of Copies stored in Salesconf
    '18/07/2002 changed to add export invoice option in case of domestic invoices
    '23/07/2002 changed to add Grin Linking in Rejection Invoice
    '07/08/2002 changed for Jobwork invoice to check Customer supplied from Vendor Bom
    'changed by nisha 0n 11/10/2002
    'changed by nisha on 22/03/2003 for Finance Rollover 24/03/200327/03/2003
    'changed by nisha on 04/04/2003
    'changed By nisha on 13/05/2003 for summit issues
    'changes done by nisha on 10/07/2003 for AIM Issues11/07/2003
    'changes done by nisha on 29/08/2003 for DATE Excise Calculation
    'Changed by nisha for 04/09/2003 for Rejection Posting 17/09/2003
    'changes Done By Nisha for Discount Posting on 19/09/2003
    'chnages done by nisha for Discount Posting in Rejection as well 23/09/2003
    'changes Done Nisha on 11/11/2003 for Excise Priority
    'Changes Done by Nisha To check if Tool Cost Deduction will be done or Not on 16/02/2004
    'Changes Done by Nisha To check if TCS Tax will be done or Not on 26/02/2004
    'Changes Done by Nisha To check if ECS Tax will be done or Not on 08/07/2004
    'Changes Done by Rajani Kant Tool Cost deducation on BOM of Finished Item on 23/08/2004
    'Changes made for Schedule updations for DS wise knocking off : By JS on 08/09/2004
    'Changes made for Schedule updations for DS wise knocking off : By NR on 11/09/2004(DSTracking-10623)
    '===================================================================================
    '-----------------------------------------------------------------------------------
    'Changes by Arshad on 20/09/2004
    'Ecess on Sale Tax Added and posting in A/c for same is going from this form
    '-----------------------------------------------------------------------------------
    'Changes by Arshad on 30/09/2004
    'by nisha for non eou sereis updat
    'Changes Done by Nisha for Tool Amortization to Link With Tool Mst on 06/10/2004-2
    'Changes Done by Nisha for removing the check of Finished item in amor_dtl on 20/10/2004
    'Changes Done by Nisha for removing the check of tool code in Finished Items in amor_dtl on 24/11/2004
    '--------------------------------------------------------
    'Changes Done By Sourabh For Add Functionality of Batch
    'Date :- 30 Nov 2004
    '-----------------------------------------------------------------------------------
    'Changes Done By Arshad Ali to Lock Service invoice Against NRGP
    'Date :- 18 Feb 2005
    ' ----------------------------------------------------------------------------------
    ' Changes Done By Sandeep
    ' Lock Rejection Invoice Detail Table MKT_INVREJ_DTL
    ' Date :- 30 March 2005
    ' ----------------------------------------------------------------------------------
    ' Changes Done By Sandeep
    ' Update the Invoice no after prining in Form Detail Table
    ' Date :- 1 Apr 2005
    ' ----------------------------------------------------------------------------------
    ' Changes Done By Nisha
    ' To Include the Sate Development Tax
    ' Date :- 03 May 2005
    ' ----------------------------------------------------------------------------------
    ' Changes Done By Nisha
    ' Condition added by nisha on 13 May 2005 for rejection invoice Round off Correctin
    ' Date :- 13 May 2005
    ' ----------------------------------------------------------------------------------
    ' Code add by sourabh
    ' For multi unit invoice (Smiel / Sumit)
    ' Date :- 05 Sep 2005
    ' ----------------------------------------------------------------------------------
    'Revision  By       : Ashutosh Verma,issue id:15591
    'Revision On        : 19-09-2005
    'History            : Save Loading Charges parameter value to Saleschallan_dtl w.r.t value in Sales_Parameter.SameUnitLoading field.
    '=======================================================================================
    'Revision  By       : Sourabh Khatri,issue id:PRJ-2004-04-003-16261
    'Revision On        : 22/Nov/2005
    'History            : To Add functionality of invoice against bar code
    '=======================================================================================
    'Revision  By       : Ashutosh Verma,issue id:16685
    'Revision On        : 26-12-2005
    'History            : Mark Unlocked as TEMPORARY INVOICE (for SUNVAC).
    '=======================================================================================
    'Revision  By       : Ashutosh Verma,issue id:17610
    'Revision On        : 02-06-2006
    'History            : Posting for Service tax in accounts.
    '=======================================================================================
    'Revision  By       : Ashutosh Verma,issue id:18099
    'Revision On        : 13-06-2006
    'History            : Check for Wrong stock updation.
    '=======================================================================================
    'Revised By         : Davinder Singh
    'Revision Date      : 20-06-2006
    'Issue Id           : 18103
    'Revision History   : To print the invoice on the multiple pages by reading the No. of
    '                     records from the SaleConf table
    '=======================================================================================
    '***********************************************************************************
    'Revised By         : Manoj Kr. Vaish
    'Issue ID           : 20052
    'Revision Date      : 04 June 2007
    'History            : Posting of SHEccess on Service Tax for JOB WORK INVOICE
    '***********************************************************************************
    'Revised By      : Manoj Kr. Vaish
    'Issue ID        : 19992
    'Revision Date   : 27 June 2007
    'History         : To add the functionality of Multiple SO for Export Invoice.
    '***********************************************************************************
    'Revised By      : Manoj Kr. Vaish
    'Issue ID        : 21105
    'Revision Date   : 09 Oct 2007
    'History         : To add Bar Code functionality for Mate manesar
    '***********************************************************************************
    'Revised By      : Manoj Kr. Vaish
    'Issue ID        : 21473
    'Revision Date   : 31 Oct 2007
    'History         : To Correct Wrong knocking off of Customer Supplied material
    '                : during job work invoice entry
    '***********************************************************************************
    'Revised By      : Manoj Kr. Vaish
    'Issue ID        : 21551
    'Revision Date   : 20-Nov-2007
    'History         : Add New Tax VAT with Sale Tax help
    '***********************************************************************************
    'Revised By      : Manoj Kr. Vaish
    'Issue ID        : 21840
    'Revision Date   : 18-DEC-2007
    'History         : Wrong schedule knockoff while making E-Nagare Invoice Entry.
    '***********************************************************************************
    'Revised By      : Manoj Kr.Vaish
    'Issue ID        : 22035
    'Revision Date   : 02-JAN-2008
    'History         : Boolean varible was not getting update when any error comes in updatemktschedule function
    '***********************************************************************************
    'Revised By      : Manoj Kr.Vaish
    'Issue ID        : 22207
    'Revision Date   : 17-JAN-2008
    'History         : Invoice Can't save with number 0
    '***********************************************************************************
    'Revised By      : Manoj Kr. Vaish
    'Issue ID        : 22286
    'Revision Date   : 06 Feb 2008
    'History         : While reprinting the Job Work Invoice 'Subscript out of range 'error message is coming.
    '                  Also it check the BOM again and accordingly knocks off the same quantity.
    '***********************************************************************************
    '***********************************************************************************
    'Revised By      : Manoj Kr. Vaish
    'Issue ID        : 22598
    'Revision Date   : 05 Mar 2008
    'History         : ECSS schedule should not be knocked off if the AllowExcessSchedule Flag is off
    '                : in Customer master for customer
    '***********************************************************************************
    'Revised By      : Manoj Kr. Vaish
    'Issue ID        : eMpro-20080430-18033
    'Revision Date   : 27 Apr 2008
    'History         : To add Calculation of Ecess & SHEccess on Total Duty (EOU Unit)
    '                  And New tax head for the calculation of CVD Excise,Ecess & SEcess
    '***********************************************************************************
    'Revised By      : Manoj Kr.Vaish
    'Issue ID        : eMpro-20080508-18500
    'Revision Date   : 09 May 2008
    'History         : To Allow SAD Tax for Transfer Invoice in Mate Noida
    '***********************************************************************************
    '***********************************************************************************
    'Revised By      : Manoj Kr. Vaish
    'Issue ID        : eMpro -20080516 - 18915
    'Revision Date   : 15-MAY-2008
    'History         : Wrong schedule knockoff while making MSSL Invoice against DS Schedule
    '***********************************************************************************
    'Revised By      : Manoj Kr. Vaish
    'Issue ID        : eMpro-20080805-20745
    'Revision Date   : 06-Aug-2008
    'History         : Assign dsn name to prj_InvoicePrinting Dll and rectification of .Net conversion
    '***********************************************************************************
    'Revised By      : Manoj Kr.Vaish
    'Issue ID        : eMpro-20080930-22159
    'Revision Date   : 30 Sep 2008
    'History         : BatchWise Tracking of Invoices Made from 01M1 Location including BarCode Tracking
    '                  Knocking Off Daily Marketing Schedule on DayWise
    '***********************************************************************************
    'Revised By      : Manoj Kr.Vaish
    'Issue ID        : eMpro-20081023-22907
    'Revision Date   : 23 Oct 2008
    'History         : While making Job Work Invoice BOM was not exploring further
    '                : when Semi Finsihed Product is using for Finished Product.
    '***********************************************************************************
    'Revised By      : Manoj Kr.Vaish
    'Issue ID        : eMpro-20090216-27468
    'Revision Date   : 17 Feb 2009
    'History         : Empro .Net Issues Rectification.
    '***********************************************************************************
    'Revised By      : Manoj Kr.Vaish
    'Issue ID        : eMpro-20090226-27911
    'Revision Date   : 26 Feb 2009
    'History         : Master Record Insertion Failure while locking the Invoice -Mate Manesar
    '***********************************************************************************
    'Revised By      : Manoj Kr. Vaish
    'Revised On      : 15 Apr 2009
    'Issue ID        : eMpro-20090415-30143
    'Reason          : ASN File Printing for Mahindra & Mahindra-Mate Pune
    '***********************************************************************************
    'Revised By      : Manoj Vaish
    'Revision On     : 14 Jan 2009
    'Issue ID        : eMpro-20090112-25902
    'History         : Some Issues in LRN/GRN Rejection Invoice
    '***********************************************************************************
    'Revised By      : Manoj Vaish
    'Revision On     : 19 May 2009
    'Issue ID        : eMpro-20090519-31544
    'History         : Conversion from string to Double is not valid Error is coming while locking the invoice
    '***********************************************************************************
    'Revised By      : Manoj Vaish
    'Revision On     : 01 Jun 2009
    'Issue ID        : eMpro-20090601-31918
    'History         : Posting of new additional VAT tax
    '***********************************************************************************
    'Revised By      : Manoj Vaish
    'Revision On     : 01 Jun 2009
    'Issue ID        : eMpro-20090610-32326
    'History         : Posting of new additional CST tax
    '***********************************************************************************
    'Revised By      : Manoj Kr. Vaish
    'Issue ID        : eMpro-20090611-32362
    'Revision Date   : 17 Jun 2009
    'History         : Export Invoice Schdule Knocking Off (RAN No. Wise)---HILEX
    '****************************************************************************************
    'Revised By      : Manoj Kr. Vaish
    'Issue ID        : eMpro-20090713-33572
    'Revision Date   : 13 Jul 2009
    'History         : Invoices values was not showing in Sales Ledger Report due to not posted in Finance table
    '****************************************************************************************
    'Revised By      : Manoj Kr. Vaish
    'Issue ID        : eMpro-20090723-34088
    'Revision Date   : 24 Jul 2009
    'History         : Addition of New Fields in Toyota CSV File-Hilex
    '****************************************************************************************
    'Revised By      : SIDDHARTH RANJAN
    'Issue ID        : eMpro-20090919-36542
    'Revision Date   : 19 SEP 2009
    'History         : Tool Amortisation cost not updateding in Invoice Locking
    '****************************************************************************************
    'Revised By      : SIDDHARTH RANJAN
    'Revision On     : 11 NOV 2009
    'Issue ID        : eMpro-20091113-38843
    'History         : ADD NEW INVOICE TYPE & SUB INVOICE TYPE ("CSM INVOICE")
    '                   ONLY TAXES WILL BE DEBITED TO CUSTOMER
    '***********************************************************************************
    'Revised By      : SIDDHARTH RANJAN
    'Revision On     : 24 NOV 2009
    'Issue ID        : eMpro-20091124-39248
    'History         : Report will not be print untill lock
    '***********************************************************************************
    'Revised By      : PRASHANT RAJPAL
    'Revision On     : 05 APR 2011
    'Issue ID        : 1084018
    'History         : CC CODE CHANGES ADDED IN TRANSFER ,SAMPLE AND REJECTION INVOICE ( CONFIGURABLE FUNCTIONALITY )
    '***********************************************************************************
    'Revised By      : PRASHANT RAJPAL
    'Revision On     : 28 JUNE 2011
    'Issue ID        : 10109115 
    'History         : Change addded for calling the Sub report 
    '***********************************************************************************
    '***********************************************************************************
    'Revised By      : PRASHANT RAJPAL
    'Revision On     : 28 JUNE 2011
    'Issue ID        : 10127115
    'History         : during invoice Generation , Prefix not appeared , now resolved
    '***********************************************************************************
    'Revised By      : Prashant Dhingra
    'Revision On     : 22 Sep 2011
    'Issue ID        : 10140220 
    'History         : Change for Auto Invoice Generation
    '***********************************************************************************
    'Revised By      : PRASHANT RAJPAL
    'Revision On     : 10 Oct  2011
    'Issue ID        : 10146492
    'History         : Invoicing is too slow for Mate Pune Changes 
    '***********************************************************************************
    'Revised By      : PRASHANT RAJPAL
    'Revision On     : 20 Oct  2011
    'Issue ID        : 10150806 
    'History         : change for shell batch file execution
    '***********************************************************************************
    'Revised By      : PRASHANT RAJPAL
    'Revision On     : 12 Nov   2011
    'Issue ID        : 10158952 
    'History         : change for Printing intead of cmd using process 
    '***********************************************************************************
    'Revised By      : Rajeev Gupta
    'Revision On     : 01 Dec 2011
    'Issue ID        : 10166568 
    'History         : Change in Invoice Printing due to CITRIX environment , File Location is configurable.
    '***********************************************************************************
    'Revised By      : PRASHANT RAJPAL
    'Issue ID        : 10170787
    'Revision Date   : 20 DEC 2011
    'History         : Time tracking start and End Time
    '***********************************************************************************
    '***********************************************************************************
    'Revised By      : PRASHANT RAJPAL
    'Issue ID        : 10192547 
    'Revision Date   : 08 FEB 2012
    'History         : Changes in Invoice Entry FOR barcode process (At Main Store )
    '***********************************************************************************
    'Revised By      : PRASHANT RAJPAL
    'Issue ID        : 10196453 
    'Revision Date   : 22 FEB 2012
    'History         : Changes in Invoice locking Auto Pune changes
    '***********************************************************************************
    'Revised By      : PRASHANT RAJPAL
    'Issue ID        : 10237233
    'Revision Date   : 15 june 2012
    'History         : Changes for ASN  Generation for RSA and Vacuform
    '***********************************************************************************
    'Revised By      : PRASHANT RAJPAL
    'Issue ID        : 10278955 
    'Revision Date   : 21 Sep 2012
    'History         : Changes for ASN  
    '***********************************************************************************
    'Revised By      : PRASHANT RAJPAL
    'Issue ID        : 10293155 
    'Revision Date   : 08 OCt 2012
    'History         : Changes for New SMIEl Database name 
    '***********************************************************************************
    'Revised By      : PRASHANT RAJPAL
    'Issue ID        : 10309680 
    'Revision Date   : 28 Nov 2012
    'History         : Changes for HILEX - Auto transfer invoice locking changes
    '***********************************************************************************
    'Revised By      : PRASHANT RAJPAL
    'Issue ID        : 10309680
    'Revision Date   : 10-dec- 2012
    'History         : Changes for HILEX - SKIP FIFO FLAG ENABLED FOR BARCODE TRACKING ON ITEMS 
    '***********************************************************************************
    'Revised By      : PRASHANT RAJPAL
    'Issue ID        : 10341052 
    'Revision Date   : 01-FEB 2013-18 FEB 2013
    'History         : Changes for MSSL AND MAE MAPPING 
    '***********************************************************************************
    'Revised By      : VINOD SINGH
    'Issue ID        : 10349154 
    'Revision Date   : 06 MARCH 2013
    'History         : SKIP_FIFO FLAG REMOVED FORM BARCODETRACKING FUNCTION
    '***********************************************************************************
    'Revised By      : PRASHANT RAJPAL
    'Issue ID        : 10364243 
    'Revision Date   : 19-apr-2013-22 apr 2013
    'History         : Changes for TKML Bacode Implementation 
    '***********************************************************************************
    'REVISED BY      : VINOD SINGH
    'ISSUE ID        : 10433205 
    'REVISION DATE   : 06 AUG 2013
    'HISTORY         : TRANSFER INVOICE KNOCKING OFF DETAILS NOT UPDAING IN BAR_ISSUE 
    '                  IN CASE OF PARTIAL PACKET ISSUE
    '***********************************************************************************
    'REVISED BY     :  VINOD SINGH
    'REVISED ON     :  06 MAY 2015
    'ISSUE ID       :  10804443 - MULTI LOCATION IN BARCODE - HILEX 
    '***********************************************************************************
    'REVISED BY     :  PRASHANT RAJPAL
    'REVISED ON     :  21-SEP2-15 TO 23-SEP-215
    'ISSUE ID       :  10902255 - A4 IN HILEX

    'REVISED BY     -  ASHISH SHARMA    
    'REVISED ON     -  23 JUN 2017
    'PURPOSE        -  101188073 — GST CHANGES
    '***********************************************************************************
    'REVISED BY     -  ASHISH SHARMA
    'REVISED ON     -  20 AUG 2020
    'PURPOSE        -  102027599 - IRN CHANGES
    '***********************************************************************************
    Public gobjDB As New ClsResultSetDB_Invoice
    Dim mStrCustMst As String
    Dim mresult As ClsResultSetDB_Invoice
    Dim mintFormIndex As Short
    Dim salesconf As String
    Dim msubTotal, mInvNo, mExDuty, mBasicAmt, mOtherAmt, TempInvNo As Double
    Dim mFrAmt, mGrTotal, mStAmt, mCustmtrl As Double
    Dim mDoc_No As Short
    Dim mAccount_Code, mInvType, mSubCat, mlocation As String
    Dim mstrAnnex As String
    'Dim str57f4Date As String      'used in BomCheck() insertupdateAnnex()
    Dim arrQty() As Double 'used in BomCheck() insertupdateAnnex()
    Dim arrItem() As String 'used in BomCheck() insertupdateAnnex()
    Dim arrReqQty() As Double
    Dim arrCustAnnex(0, 0) As Object
    Dim ref57f4 As String 'used in BomCheck() insertupdateAnnex()
    Dim dblFinishedQty As Double 'To get Qty of Finished Item from Spread
    Dim strCustCode As String 'used in BomCheck() insertupdateAnnex()
    Dim StrItemCode As String 'used in BomCheck() insertupdateAnnex()
    Dim inti As Short 'To Change Array Size used in BomCheck() insertupdateAnnex()
    Dim strsaledetails As String
    Dim strupdateGrinhdr As String
    Dim strupdateitbalmst As String
    Dim strSelectItmbalmst As String

    Dim strupdatecustodtdtl As String
    Dim strUpdateAmorDtl As String
    Dim strupdateamordtlbom As String
    Dim mCust_Ref, mAmendment_No As String
    Dim saleschallan As String
    Dim ValidRecord As Boolean
    Dim updatestockflag, updatePOflag As Boolean
    Dim strStockLocation As String
    Dim mAmortization As Double
    Dim mblnEOUUnit As Boolean
    Dim mAssessableValue As Double
    Dim mOpeeningBalance As Double
    Dim mblnCustSupp As Boolean
    Dim strBomItem As String 'For Latest Item To Explore
    Dim blnFIFOFlag As Boolean
    Dim rsBomMst As ClsResultSetDB_Invoice
    Dim mstrMasterString As String 'To store master string for passing to Dr Cr COM
    Dim mstrDetailString As String 'To store detail string for passing to Dr Cr COM
    Dim mstrPurposeCode As String 'To store the Purpose Code which will be used for the fetching of GL and SL
    Dim mblnAddCustomerMaterial As Boolean 'To decide whether to add customer material in basic or not
    Dim mblnSameSeries As Boolean 'To store the flag whether the selected invoice will have same series as others
    Dim mstrReportFilename As String 'To store the report filename
    Dim mblnInsuranceFlag As Boolean 'To store insurance flag
    Dim mblnpostinfin As Boolean
    ''Added By Tapan On 8-Mar-2K3
    Dim mblnExciseRoundOFFFlag As Boolean
    ''Addition Ends
    Dim mSaleConfNo As Double
    Dim mstrExcisePriorityUpdationString As String
    Dim objInvoicePrint As prj_InvoicePrinting.clsInvoicePrinting 'Added By Arshad on 23/04/2004 for Dos Based Printing
    '''Dim objInvoicePrint As New clsInvoicePrinting
    Dim intNoCopies As Short
    Dim mblnServiceInvoiceWithoutSO As Boolean
    Dim mstrGrinQtyUpdate As String
    Dim mstrInvRejSQL As String
    Dim mblnJobWkFormulation As String
    Dim mblnInvoiceAgainstBarCode As Boolean
    'Added for Issue ID 19992 Starts
    Dim mblnMultipleSOAllowed As Boolean
    Dim mblnSORequired As Boolean
    Dim mstrCreditTermId As String
    'Added for Issue ID 19992 Ends
    'Added for Issue ID 20918 Starts
    Dim mblnInvoiceMTLSharjah As Boolean
    Dim mexchange_rate As Double
    'Added for Issue ID 20918 Ends
    'Added for Issue ID 21105 Starts
    Dim mstrupdateBarBondedStockFlag As String
    Dim mstrupdateBarBondedStockQty As String
    Dim mblnQuantityCheck As Boolean
    'Added for Issue ID 21105 Ends
    'Added for Issue ID 21473 Starts
    Dim mblnConversion As Boolean
    'Added for Issue ID 21473 Ends
    'Adde for BarCode Issue Starts
    Dim mstrFGDomestic As String
    Dim mstrError As String
    'Adde for BarCode Issue Ends
    'Added for Issue ID eMpro-20090415-30143 Starts
    Dim mblnASNExist As Boolean
    'Added for Issue ID eMpro-20090415-30143 Ends
    Dim mblnCCFlag As Boolean
    Dim strCitrix_Inv_Pronting_Loc As String
    Dim strGrnDespatchQuantity_TradingInvoice As String = String.Empty 'AMIT RANA mrigendra modification
    Dim strSALESTRADINGGRINDTL_TradingInvoice As String = String.Empty 'AMIT RANA mrigendra modification
    Dim mstrupdateASNdtl As String
    Dim mstrupdateASNCumFig As String
    Dim mblnDuplicateASNExist As Boolean
    Dim mstrins As String   ''mrigendra modification
    Dim ArrEdiFiles As New ArrayList  ''mrigendra modification
    Dim mblnTrading As Boolean = False  ''mrigendra modification
    Dim mblnDiscount_invoicewise As Boolean  ''mrigendra modification
    Dim mblnDiscountFunctionality As Boolean  ''mrigendra modification
    Dim mblncustomerlevel_A4report_functionlity As Boolean  ''mrigendra modification
    Dim mblnA4reports_invoicewise As Boolean  ''mrigendra modification
    Dim intNoCopies_A4reports As Short
    Dim mblncustomerspecificreport As Boolean
    Dim mblnCSMspecificreport As Boolean = False
    Dim mblnskipdacinvoicebincheck As Boolean
    Dim mblnftsitem As Boolean = False
    Dim mblnftsbarcodeitem As Boolean = False
    Dim mblnftsenabled As Boolean = False
    Dim strFTSstocklocation As String
    Dim strsaleconfLocation As String
    Dim mblnEwaybill_Print As Boolean = False
    Dim mblnEWAY_BILL_STARTDATE As String
    Dim mdblewaymaximumvalue As Double
    Dim mblnTOYOTA_MULTIPLESO_ONEPDS_STDATE As String
    Dim mstrREJ_INVOICE_NEWINVREPORT_STARTDATE As String
    Dim blnlinelevelcustomer As Boolean = False


    Private Sub cmbInvType_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbInvType.TextChanged
        On Error GoTo ErrHandler
        Call cmbInvType_Validating(cmbInvType, New System.ComponentModel.CancelEventArgs(False))
        Exit Sub
ErrHandler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub cmdClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClose.Click
        On Error GoTo ErrHandler
        FraInvoicePreview.Visible = False
        'UPGRADE_NOTE: Object objInvoicePrint may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        objInvoicePrint = Nothing
        Exit Sub
ErrHandler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, err.Source, Err.Description, mp_connection)
    End Sub


    Private Sub cmdPrint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdPrint.Click
        '=======================================================================================
        'Revision  By       : Ashutosh Verma,issue id:16685
        'Revision On        : 26-12-2005
        'History            : Mark Unlocked as TEMPORARY INVOICE (for SUNVAC).
        '=======================================================================================

        On Error GoTo ErrHandler
        Dim intCount As Short
        Dim varTemp As Object
        Dim strFileName As String
        Dim rsComp As New ADODB.Recordset


        If rsComp.State = ADODB.ObjectStateEnum.adStateOpen Then rsComp.Close()
        rsComp.Open("Select Company_Code from sales_parameter WHERE UNIT_CODE='" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        If rsComp.RecordCount > 0 Then
            If UCase(Trim(rsComp.Fields("Company_Code").Value)) = "SUN" Then

            End If
        End If
        'Change for Issue ID eMpro-20090226-27911 Starts
        If objInvoicePrint Is Nothing Then
            'shalini
            'strFileName = "C:\InvoicePrint.txt"
            strFileName = strCitrix_Inv_Pronting_Loc & "InvoicePrint.txt"
        Else
            If Len(objInvoicePrint.FileName) > 0 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object objInvoicePrint.FileName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                strFileName = objInvoicePrint.FileName
            End If
            objInvoicePrint = Nothing
        End If
        'Change for Issue ID eMpro-20090226-27911 Ends

        If intNoCopies = 0 Then intNoCopies = 1

TypeFileNotFoundCreateRetry:
        For intCount = 1 To intNoCopies
            'shalini
            'varTemp = Shell("C:\TypeToPrn.bat " & strFileName, AppWinStyle.Hide)
            '10150806 
            'varTemp = Shell(gStrDriveLocation & "TypeToPrn.bat " & strFileName, AppWinStyle.Hide)
            'issue id 10158952
            '            Write_In_Log_File(GetServerDateTime() & " : Command Exe Invoice Print Start Time")
            varTemp = Shell("cmd.exe /c " & strCitrix_Inv_Pronting_Loc & "TypeToPrn.bat " & strFileName, AppWinStyle.Hide)
            '            Write_In_Log_File(GetServerDateTime() & " : Command Exe Invoice Print End Time")
            '10150806 end 
            'Sleep(5000)

            'Call PrintViaBatchFile(gStrDriveLocation & "TypeToPrn.bat ", strFileName)
            'Exit For
            'issue id 10158952 end 

            If UCase(Trim(rsComp.Fields("Company_Code").Value)) <> "SUN" Then
                'shalini
                'varTemp = Shell("c:\TypeToPrn.bat C:\PageFeed.txt", AppWinStyle.Hide)
                '10150806 
                'varTemp = Shell(gStrDriveLocation & "TypeToPrn.bat " & gStrDriveLocation & "PageFeed.txt", AppWinStyle.Hide)
                '                Write_In_Log_File(GetServerDateTime() & " : Command Exe PageFeed Start Time")
                varTemp = Shell("cmd.exe /c " & strCitrix_Inv_Pronting_Loc & "TypeToPrn.bat " & strCitrix_Inv_Pronting_Loc & "PageFeed.txt", AppWinStyle.Hide)
                '                Write_In_Log_File(GetServerDateTime() & " : Command Exe PageFeed End Time")
                '10150806 end 
            Else
                'shalini
                'varTemp = Shell("c:\TypeToPrn.bat C:\BarCodePageFeed.txt", AppWinStyle.Hide)
                '10150806 
                'varTemp = Shell(gStrDriveLocation & "TypeToPrn.bat " & gStrDriveLocation & "BarCodePageFeed.txt", AppWinStyle.Hide)
                'Write_In_Log_File(GetServerDateTime() & " : Command Exe BarCodePageFeed Start Time")
                varTemp = Shell("cmd.exe /c " & strCitrix_Inv_Pronting_Loc & "TypeToPrn.bat " & strCitrix_Inv_Pronting_Loc & "BarCodePageFeed.txt", AppWinStyle.Hide)
                'Write_In_Log_File(GetServerDateTime() & " : Command Exe BarCodePageFeed End Time")
                '10150806 end 
            End If
            '''***** Changes by Ashutosh on 24-12-2005 end here.
        Next
        Exit Sub
ErrHandler:
        If Err.Number = 53 Then
            'Open App.Path & "\" & "TypeToPrn.bat" For Append As #1

            'shalini
            'FileOpen(1, "C:\TypeToPrn.bat", OpenMode.Append)
            FileOpen(1, strCitrix_Inv_Pronting_Loc & "TypeToPrn.bat", OpenMode.Append)
            PrintLine(1, "Type %1> prn") '& Printer.Port
            FileClose(1)
            GoTo TypeFileNotFoundCreateRetry
        End If
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub


    '    Private Function Write_In_Log_File(ByVal Log_Line As String) As Boolean
    '        On Error GoTo Errorhandler
    '        Call CheckCreateFolders(strCitrix_Inv_Pronting_Loc & "LOG")
    '        Call Append_Line(strCitrix_Inv_Pronting_Loc & "LOG" & "\" & "InvoicePrintingLog.txt", Log_Line)
    '        Exit Function
    'Errorhandler:
    '        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    '    End Function
    Private Function Append_Line(ByVal pstrPath As String, ByVal Log_Line As String) As Object

        On Error GoTo Errorhandler
        Dim FSO As Scripting.FileSystemObject
        Dim objWriter As New System.IO.StreamWriter(pstrPath, True)

        FSO = New Scripting.FileSystemObject
        If FSO.FileExists(pstrPath) = True Then
        Else
            FSO.CreateTextFile(pstrPath)
        End If
        objWriter.WriteLine(Log_Line)
        objWriter.Close()

        FSO = Nothing
        Exit Function
Errorhandler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function CheckCreateFolders(ByVal pstrPath As String) As Object
        ''---------------------------------------------------------------------------
        '' Function     : Check the existence of folder 
        ''---------------------------------------------------------------------------
        On Error GoTo Errorhandler
        Dim FSO As Scripting.FileSystemObject

        FSO = New Scripting.FileSystemObject
        If FSO.FolderExists(pstrPath) = True Then
            FSO = Nothing
            Exit Function
        Else
            FSO.CreateFolder(pstrPath)
        End If
        FSO = Nothing
        Exit Function
Errorhandler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

    Private Sub cmdUnitCodeList_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdUnitCodeList.Click
        On Error GoTo ErrHandler
        Call ShowCode_Desc("SELECT Unt_CodeID,unt_unitname FROM Gen_UnitMaster WHERE Unt_CodeID='" + gstrUNITID + "' AND Unt_Status=1", txtUnitCode)
        Exit Sub
ErrHandler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub Ctlinvoice_Change(ByVal Sender As System.Object, ByVal e As System.EventArgs) Handles Ctlinvoice.Change
        chkAcceffects.CheckState = System.Windows.Forms.CheckState.Unchecked
    End Sub
    Private Sub dtpRemoval_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSComCtl2.DDTPickerEvents_KeyDownEvent) Handles dtpRemoval.KeyDownEvent
        Select Case eventArgs.keyCode
            Case 39, 34, 96
                eventArgs.keyCode = 0
            Case 13
                'Cmdinvoice.SetFocus
                '********Added By Tapan on 20-Aug-2K2******
                'chkLockPrintingFlag.Enabled = True
                'chkLockPrintingFlag.Focus()
                '*********Addition Ends**************
        End Select

    End Sub

    'UPGRADE_WARNING: Event txtUnitCode.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub txtUnitCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUnitCode.TextChanged
        On Error GoTo ErrHandler
        If Trim(txtUnitCode.Text) = "" Then
            cmbInvType.Enabled = False
            'cmbInvType.BackColor = glngCOLOR_DISABLED
            cmbInvType.SelectedIndex = -1
            CmbCategory.Enabled = False
            'CmbCategory.BackColor = glngCOLOR_DISABLED
            CmbCategory.SelectedIndex = -1
            chkAcceffects.CheckState = System.Windows.Forms.CheckState.Unchecked
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtUnitCode_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUnitCode.Enter
        On Error GoTo ErrHandler

        'Selecting the text on focus
        With txtUnitCode
            .SelectionStart = 0 : .SelectionLength = Len(Trim(.Text))
        End With

        Exit Sub 'This is to avoid the execution of the error handler

ErrHandler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub txtUnitCode_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtUnitCode.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler

        'If Ctrl/Alt/Shift is also pressed
        If Shift <> 0 Then Exit Sub
        'Show the help form when user pressed F1
        If KeyCode = System.Windows.Forms.Keys.F1 Then Call cmdUnitCodeList_Click(cmdUnitCodeList, New System.EventArgs())
        Exit Sub 'This is to avoid the execution of the error handler
        'UPGRADE_ISSUE: Constant vbEnter was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
        If KeyCode = Keys.Enter Then System.Windows.Forms.SendKeys.Send("{TAB}")
ErrHandler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub txtUnitCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtUnitCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then
            System.Windows.Forms.SendKeys.Send("{TAB}")
            'Supressing ¬ ¤ ¦ » characters since these are being used as string delimiters
        ElseIf KeyAscii = 187 Or KeyAscii = 166 Or KeyAscii = 164 Or KeyAscii = 172 Or KeyAscii = 39 Or KeyAscii = 34 Or KeyAscii = 96 Then
            KeyAscii = 0
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub TxtUnitCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtUnitCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        Dim strUnitDesc As String
        Dim mobjGLTrans As prj_GLTransactions.cls_GLTransactions
        'Populate the details
        If Trim(txtUnitCode.Text) = "" Then GoTo EventExitSub
        ''mobjGLTrans = New prj_GLTransactions.cls_GLTransactions  mrigendra modification
        mobjGLTrans = New prj_GLTransactions.cls_GLTransactions(gstrUNITID, GetServerDate)
        strUnitDesc = mobjGLTrans.GetUnit(Trim(txtUnitCode.Text), ConnectionString:=gstrCONNECTIONSTRING)
        mobjGLTrans = Nothing

        If CheckString(strUnitDesc) <> "Y" Then
            MsgBox(CheckString(strUnitDesc), MsgBoxStyle.Critical, "eMPro")
            txtUnitCode.Text = ""
            cmbInvType.Enabled = True
            'cmbInvType.BackColor = glngCOLOR_DISABLED
            cmbInvType.SelectedIndex = -1
            CmbCategory.Enabled = True
            'CmbCategory.BackColor = glngCOLOR_DISABLED
            CmbCategory.SelectedIndex = -1
            Cancel = True
        Else
            If mblnEOUUnit = True Then
                'changes done by nish to add service type of invoice
                Call selectDataFromSaleConf(Trim(txtUnitCode.Text), cmbInvType, "Description", "'INV','SMP','TRF','REJ','JOB','SRC','CSM'", "datediff(dd,getdate(),fin_start_date)<=0  and datediff(dd,fin_end_date,getdate())<=0")
            Else
                'changes done by nish to add service type of invoice
                Call selectDataFromSaleConf(Trim(txtUnitCode.Text), cmbInvType, "Description", "'INV','SMP','TRF','REJ','JOB','EXP','SRC','CSM'", "datediff(dd,getdate(),fin_start_date)<=0  and datediff(dd,fin_end_date,getdate())<=0")
            End If
            cmbInvType.Enabled = True
            cmbInvType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            CmbCategory.Enabled = True
            CmbCategory.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            Call cmbInvType_Validating(cmbInvType, New System.ComponentModel.CancelEventArgs(False))
        End If

        GoTo EventExitSub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    '********Added By Tapan On 20-Aug-2K2*************
    Private Sub chkLockPrintingFlag_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkLockPrintingFlag.Enter
        shpLock.Visible = True
    End Sub
    Private Sub chkLockPrintingFlag_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles chkLockPrintingFlag.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then
            Cmdinvoice.Focus()
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub chkLockPrintingFlag_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkLockPrintingFlag.Leave
        shpLock.Visible = False
    End Sub
    '********Addition Ends*************************
    'UPGRADE_WARNING: Event CmbCategory.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub CmbCategory_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmbCategory.SelectedIndexChanged
        On Error GoTo Err_Handler
        If Len(Trim(CmbCategory.Text)) = 0 Or Trim(CmbCategory.Text) = "-None-" Or Len(Trim(CmbCategory.Text)) > 0 Then
            lblcategory.Text = ""
            Ctlinvoice.Text = ""
            'cmdHelp(2).Enabled = False
        End If
        Exit Sub
Err_Handler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub CmbCategory_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles CmbCategory.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo Err_Handler
        If KeyAscii = 13 Then
            Call CmbCategory_Validating(CmbCategory, New System.ComponentModel.CancelEventArgs(False))
            If Ctlinvoice.Enabled = False Then
                Cmdinvoice.Focus()
            Else
                Ctlinvoice.Focus()
            End If
        End If
        GoTo EventExitSub
Err_Handler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub CmbCategory_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles CmbCategory.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim rsSalesConf As New ADODB.Recordset
        On Error GoTo Err_Handler
        If Len(Trim(CmbCategory.Text)) = 0 Or (Trim(CmbCategory.Text) = "-None-") Then
            If Trim(CmbCategory.Text) = "-None-" Then
                Ctlinvoice.Enabled = False
                frachkRequired.Enabled = False
                Ctlinvoice.Text = ""
                Ctlinvoice.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                cmdHelp(2).Enabled = False
                GoTo EventExitSub
            End If
        End If
        'CHANGES DONE BY AMIT BHATNAGAR ON 26.09.2002****************************************
        If Not (Len(CmbCategory.Text) <= 0) Then 'Checking if Item Field is not Blank
            If UCase(lbldescription.Text) = "SMP" And mblnpostinfin = True Then
                If rsSalesConf.State = ADODB.ObjectStateEnum.adStateOpen Then rsSalesConf.Close()
                rsSalesConf.Open("SELECT * FROM fin_GlobalGl WHERE UNIT_CODE='" + gstrUNITID + "' AND  gbl_prpsCode='Sample_Expences'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
                If rsSalesConf.EOF Then
                    MsgBox("Please define Sample Expences Account in Global Gl Definition", MsgBoxStyle.Information, "eMPro")
                End If
            End If
            If rsSalesConf.State = ADODB.ObjectStateEnum.adStateOpen Then rsSalesConf.Close()
            rsSalesConf.CursorLocation = ADODB.CursorLocationEnum.adUseClient
            'CHANGES DONE BY NISHA ON 21/03/2003 FOR FINANCIAL ROLLOVER
            rsSalesConf.Open("SELECT * FROM SaleConf WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_Type='" & lbldescription.Text & "' AND Sub_Type_description ='" & CmbCategory.Text & "' AND Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,getdate(),fin_start_date)<=0  and datediff(dd,fin_end_date,getdate())<=0 ", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
            'CHANGES ENDS HERE ON 21/03/2003
            If Not rsSalesConf.EOF Then
                mblnEwaybill_Print = rsSalesConf.Fields("EWAY_BILL_FUNCTIONALITY").Value

                If mblnEwaybill_Print = True Then
                    chkprintreprint.Enabled = True
                    chkprintreprint.Checked = True
                Else
                    chkprintreprint.Enabled = False
                    chkprintreprint.Checked = False
                End If

                mstrPurposeCode = Trim(IIf(IsDBNull(rsSalesConf.Fields("inv_GLD_prpsCode").Value), "", rsSalesConf.Fields("inv_GLD_prpsCode").Value))
                mblnSameSeries = rsSalesConf.Fields("Single_Series").Value

                mstrReportFilename = Trim(IIf(IsDBNull(rsSalesConf.Fields("Report_filename").Value), "", rsSalesConf.Fields("Report_filename").Value))
                If mstrPurposeCode = "" Then
                    MsgBox("Please select a Purpose Code in Sales Configuration", MsgBoxStyle.Information, "eMPro")
                    Me.CmbCategory.SelectedIndex = 0
                    Me.lblcategory.Text = ""
                    Me.cmbInvType.SelectedIndex = 3
                    Me.lbldescription.Text = ""
                    Me.cmbInvType.Focus()
                    mstrPurposeCode = ""
                    GoTo EventExitSub
                End If
            Else
                MsgBox("No record found in Sales Configuration for the selected Location, Invoice Type and Sub-Category", MsgBoxStyle.Information, "eMPro")
                Me.CmbCategory.SelectedIndex = 0
                Me.lblcategory.Text = ""
                Me.cmbInvType.SelectedIndex = 3
                Me.lbldescription.Text = ""
                Me.cmbInvType.Focus()
                mstrPurposeCode = ""
                GoTo EventExitSub
            End If
            'mresult.GetResult ("Select sub_type,Sub_Type_Description,Stock_Location,updateStock_Flag  from SaleConf where Invoice_type = '" & Trim(Me.lblDescription.Caption) & "' and sub_Type_Description = '" & Trim(Me.CmbCategory.Text) & "'")
            'CHANGES DONE BY NISHA ON 21/03/2003 FOR FINANCE ROLLOVER
            mresult = New ClsResultSetDB_Invoice
            mresult.GetResult("Select sub_type,Sub_Type_Description,Stock_Location,updateStock_Flag  from SaleConf where UNIT_CODE='" + gstrUNITID + "' AND  Invoice_type = '" & Trim(Me.lbldescription.Text) & "' and sub_Type_Description = '" & Trim(Me.CmbCategory.Text) & "' and Location_Code ='" & Trim(txtUnitCode.Text) & "' and datediff(dd,GETDATE(),fin_start_date)<=0  and datediff(dd,fin_end_date,GETDATE())<=0")
            'CHANGES ENDS HERE 21/03/2003
            'CHANGES DONE BY AMIT BHATNAGAR ON 26.09.2002 ENDS HERE*****************************
            If (mresult.GetNoRows = 0) Then
                Call ConfirmWindow(10002, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                'Cancel = True
                Me.CmbCategory.SelectedIndex = 0
                Me.lblcategory.Text = ""
                Ctlinvoice.Text = ""
                Ctlinvoice.Enabled = False
                frachkRequired.Enabled = False
                Ctlinvoice.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                cmdHelp(2).Enabled = False
                mresult.ResultSetClose()
                Me.CmbCategory.Focus()
                GoTo EventExitSub
            Else

                If (CBool(Trim(mresult.GetValue("updateStock_Flag"))) = True) Then

                    If (Len(Trim(mresult.GetValue("Stock_location"))) > 0) Then

                        lblcategory.Text = mresult.GetValue("Sub_Type")

                        mresult.ResultSetClose()

                        'changes done by Nisha on 04/09/2003 for taking no effacts in accounts
                        If Trim(UCase(CmbCategory.Text)) = "REJECTION" Then
                            ' lblaccEffects.Visible = True
                            '''Changes done By Ashutosh on 05 Jun 2007, Issue Id:19934
                            If RejInvOptionalPostingFlag() = True Then
                                chkAcceffects.Enabled = True
                            Else
                                chkAcceffects.Enabled = False
                            End If
                            '''Changes for Issue Id;19934 end here.
                        Else
                            ' lblaccEffects.Visible = False
                            chkAcceffects.Enabled = False
                            shpacceffects.Visible = False
                        End If
                        'changes ends here on 04/09/2003
                        Ctlinvoice.Enabled = True
                        Ctlinvoice.BackColor = System.Drawing.Color.White
                        frachkRequired.Enabled = True
                        optYes(0).Enabled = True
                        optYes(1).Enabled = True
                        cmdHelp(2).Enabled = True
                        Me.Ctlinvoice.Focus()
                    Else
                        Call ConfirmWindow(10439, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                        'Cancel = True
                        Me.CmbCategory.SelectedIndex = 0
                        Me.lblcategory.Text = ""
                        Me.cmbInvType.SelectedIndex = 3
                        Me.lbldescription.Text = ""
                        mresult.ResultSetClose()
                        Me.cmbInvType.Focus()
                        GoTo EventExitSub
                    End If
                Else

                    lblcategory.Text = mresult.GetValue("Sub_Type")
                    mresult.ResultSetClose()
                    Ctlinvoice.Enabled = True
                    Ctlinvoice.BackColor = System.Drawing.Color.White
                    frachkRequired.Enabled = True
                    optYes(0).Enabled = True
                    optYes(1).Enabled = True
                    cmdHelp(2).Enabled = True
                    Me.Ctlinvoice.Focus()
                End If
            End If

        End If
        GoTo EventExitSub
Err_Handler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    'UPGRADE_WARNING: Event CmbInvType.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub CmbInvType_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbInvType.SelectedIndexChanged
        On Error GoTo Err_Handler
        If Len(Trim(cmbInvType.Text)) = 0 Then
            lbldescription.Text = ""
            CmbCategory.SelectedIndex = -1

        End If
        If Len(cmbInvType.Text) > 0 Or cmbInvType.Text = "-None-" Then 'Checking if Item Field is not Blank
            'Me.cmbInvType.ListIndex = 0
            Me.lbldescription.Text = ""
            Ctlinvoice.Text = ""
            cmdHelp(2).Enabled = False
            CmbCategory.Enabled = True
            lblcategory.Text = ""
            CmbCategory.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            Ctlinvoice.Enabled = True
            frachkRequired.Enabled = False
            Ctlinvoice.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            Call cmbInvType_Validating(cmbInvType, New System.ComponentModel.CancelEventArgs(False))
            Exit Sub
        End If
        Exit Sub
Err_Handler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmbInvType_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles cmbInvType.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo Err_Handler
        If KeyAscii = 13 Then
            Call cmbInvType_Validating(cmbInvType, New System.ComponentModel.CancelEventArgs(False))
            If CmbCategory.Enabled = False Then
                Cmdinvoice.Focus()
            Else
                CmbCategory.Focus()
            End If
        End If
        GoTo EventExitSub
Err_Handler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub cmbInvType_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles cmbInvType.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo Err_Handler
        If (Len(cmbInvType.Text) = 0) Or cmbInvType.Text = "-None-" Then
            If cmbInvType.Text = "-None-" Then
                CmbCategory.Enabled = False
                CmbCategory.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                Ctlinvoice.Enabled = False
                frachkRequired.Enabled = False
                Ctlinvoice.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                Ctlinvoice.Text = ""
                cmdHelp(2).Enabled = False
            End If
            GoTo EventExitSub
        End If
        If Len(cmbInvType.Text) > 0 Then 'Checking if Item Field is not Blank
            'changed by nisha on 18/07/2002 for export invoice
            If mblnEOUUnit = True Then
                'CHANGED BY NISHA ON 21/03/2003 FOR FINANCE ROLLOVER
                mresult = New ClsResultSetDB_Invoice
                mresult.GetResult("Select distinct(Invoice_type),Description from SaleConf where UNIT_CODE='" + gstrUNITID + "' AND  Invoice_Type in('INV','SMP','REJ','TRF','JOB','SRC','CSM')and Description = '" & cmbInvType.Text & "' and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,GETDATE(),fin_start_date)<=0  and datediff(dd,fin_end_date,GETDATE())<=0")
            Else
                mresult = New ClsResultSetDB_Invoice
                mresult.GetResult("Select distinct(Invoice_type),Description from SaleConf where UNIT_CODE='" + gstrUNITID + "' AND  Invoice_Type in('INV','SMP','REJ','TRF','JOB','EXP','SRC','CSM')and Description = '" & cmbInvType.Text & "' and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,GETDATE(),fin_start_date)<=0  and datediff(dd,fin_end_date,GETDATE())<=0")
                'CHANGES ENDS HERE 21/03/2003
            End If
            If mresult.GetNoRows = 0 Then
                mresult.ResultSetClose()
                Call ConfirmWindow(10002, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_EXCLAMATION)
                Me.cmbInvType.SelectedIndex = 0
                Me.lbldescription.Text = ""
                Cancel = True
                GoTo EventExitSub
            Else

                lbldescription.Text = mresult.GetValue("Invoice_type")
                mresult.ResultSetClose()
                CmbCategory.Enabled = True
                CmbCategory.BackColor = System.Drawing.Color.White
                'changes done by Nisha on 04/09/2003 for taking no effacts in accounts
                If Trim(UCase(cmbInvType.Text)) = "REJECTION" Then
                    '  lblaccEffects.Visible = True
                    '''Changes done by Ashutosh on 12 jun 2007 ,Issue Id:19934
                    If RejInvOptionalPostingFlag() = True Then
                        chkAcceffects.Enabled = True
                    Else
                        chkAcceffects.Enabled = False
                    End If
                    '''Changes for Issue Id:19934 end here.
                Else
                    ' lblaccEffects.Visible = False
                    chkAcceffects.Enabled = False
                    shpacceffects.Visible = False
                End If
                'changes ends here on 04/09/2003
                'CHANGED BY NISHA ON 21/03/2003 FOR FINANCE ROLLOVER
                Call selectDataFromSaleConf(Trim(txtUnitCode.Text), CmbCategory, "Sub_Type_Description", "'" & Trim(lbldescription.Text) & "'", " datediff(dd,getdate(),fin_start_date)<=0  and datediff(dd,fin_end_date,getdate())<=0")
                'CHANGES ENDS HERE 21/03/2003
                Call CmbCategory_Validating(CmbCategory, New System.ComponentModel.CancelEventArgs(False))
            End If

            Me.CmbCategory.Focus()
        End If
        GoTo EventExitSub
Err_Handler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
        Dim Index As Short = cmdHelp.GetIndex(eventSender)

        Dim varHelp As Object
        Dim strQry As String

        On Error GoTo Err_Handler
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        Select Case Index
            Case 2
                If optInvYes(0).Checked = True Then
                    strQry = "SELECT convert(bigint,Doc_no) as doc_no,Invoice_Type,cust_name FROM Saleschallan_dtl SC WHERE UNIT_CODE='" + gstrUNITID + "' AND  invoice_type='" & Me.lbldescription.Text & "' and sub_category='" & Me.lblcategory.Text & "' and Doc_No >99000000 and bill_flag = 0 and Location_Code='" & Trim(txtUnitCode.Text) & "' "
                    varHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQry)
                Else
                    If mblnEwaybill_Print = True Then
                        If chkprintreprint.Checked = True Then
                            'strQry = "set dateformat 'dmy' SELECT distinct convert(varchar,Doc_no) as doc_no,Invoice_Type,cust_name  ,'' as EWAY_BILL_NO  FROM  Saleschallan_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND invoice_type='" & Trim(Me.lbldescription.Text) & "' and sub_category='" & Trim(Me.lblcategory.Text) & "' and Doc_No < 99000000 and bill_flag = '1' and cancel_flag = '0' and Location_Code='" & Trim(txtUnitCode.Text) & "'"
                            'strQry += " AND (( INVOICE_DATE < '" & mblnEWAY_BILL_STARTDATE & " '))"
                            'strQry += " UNION  SELECT distinct CONVERT(VARCHAR,DOC_NO) AS DOC_NO,INVOICE_TYPE,CUST_NAME,EWAY_BILL_NO  FROM  SALESCHALLAN_DTL WHERE UNIT_CODE='" + gstrUNITID + "' AND INVOICE_TYPE='" & Trim(Me.lbldescription.Text) & "' AND SUB_CATEGORY='" & Trim(Me.lblcategory.Text) & "' AND DOC_NO < 99000000 AND BILL_FLAG = '1' AND CANCEL_FLAG = '0' AND LOCATION_CODE='" & Trim(txtUnitCode.Text) & "'"
                            'strQry += " AND TOTAL_AMOUNT <= " & mdblewaymaximumvalue
                            'strQry += " UNION  SELECT distinct CONVERT(VARCHAR,DOC_NO) AS DOC_NO,INVOICE_TYPE,CUST_NAME,EWAY_BILL_NO  FROM  SALESCHALLAN_DTL WHERE UNIT_CODE='" + gstrUNITID + "' AND INVOICE_TYPE='" & Trim(Me.lbldescription.Text) & "' AND SUB_CATEGORY='" & Trim(Me.lblcategory.Text) & "' AND DOC_NO < 99000000 AND BILL_FLAG = '1' AND CANCEL_FLAG = '0' AND LOCATION_CODE='" & Trim(txtUnitCode.Text) & "'"
                            'strQry += " AND TOTAL_AMOUNT > " & mdblewaymaximumvalue & " AND  INVOICE_DATE >= '" & mblnEWAY_BILL_STARTDATE & " ' "
                            'strQry += " AND ISNULL(EWAY_BILL_NO,'')<>'' AND  INVOICE_DATE >= '" & mblnEWAY_BILL_STARTDATE & " '"
                            'strQry += " AND  EXISTS ( SELECT TOP 1 1 FROM FIRSTTIME_INVOICEPRINTING (NOLOCK) WHERE UNIT_CODE =SALESCHALLAN_DTL.UNIT_CODE AND FIRSTTIME_INVOICEPRINTING.DOC_NO=SALESCHALLAN_DTL.DOC_NO )"

                            '102027599
                            strQry = "SELECT DOC_NO,INVOICE_TYPE,CUST_NAME,EWAY_BILL_NO FROM DBO.UDF_GET_LOCKED_INVOICES_FOR_INVOICE_PRINTING('" & gstrUnitId & "','" & Trim(txtUnitCode.Text) & "',1,'" & Me.lbldescription.Text & "','" & Me.lblcategory.Text & "') ORDER BY DOC_NO "
                        Else
                            'strQry = "set dateformat 'dmy' "
                            'strQry += " SELECT CONVERT(VARCHAR,DOC_NO) AS DOC_NO,INVOICE_TYPE,CUST_NAME ,EWAY_BILL_NO FROM  SALESCHALLAN_DTL WHERE UNIT_CODE='" + gstrUNITID + "' AND INVOICE_TYPE='" & Trim(Me.lbldescription.Text) & "' AND SUB_CATEGORY='" & Trim(Me.lblcategory.Text) & "' AND DOC_NO < 99000000 AND BILL_FLAG = '1' AND CANCEL_FLAG = '0' AND LOCATION_CODE='" & Trim(txtUnitCode.Text) & "'"
                            'strQry += " AND TOTAL_AMOUNT > " & mdblewaymaximumvalue & " AND  INVOICE_DATE >= '" & mblnEWAY_BILL_STARTDATE & " ' "
                            'strQry += " AND ISNULL(EWAY_BILL_NO,'')<>'' AND  INVOICE_DATE >= '" & mblnEWAY_BILL_STARTDATE & " ' "
                            'strQry += "AND  NOT EXISTS ( SELECT TOP 1 1 FROM FIRSTTIME_INVOICEPRINTING (NOLOCK) WHERE UNIT_CODE =SALESCHALLAN_DTL.UNIT_CODE AND FIRSTTIME_INVOICEPRINTING.DOC_NO=SALESCHALLAN_DTL.DOC_NO ) "
                            'strQry += " order by SALESCHALLAN_DTL.DOC_NO " ' added by priti on 13022020 to order by docno

                            '102027599
                            strQry = "SELECT DOC_NO,INVOICE_TYPE,CUST_NAME,EWAY_BILL_NO FROM DBO.UDF_GET_LOCKED_INVOICES_FOR_INVOICE_PRINTING('" & gstrUnitId & "','" & Trim(txtUnitCode.Text) & "',0,'" & Me.lbldescription.Text & "','" & Me.lblcategory.Text & "') ORDER BY DOC_NO "
                        End If
                    Else
                        strQry = "SELECT convert(varchar,Doc_no) as doc_no,Invoice_Type,cust_name  FROM  Saleschallan_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND invoice_type='" & Trim(Me.lbldescription.Text) & "' and sub_category='" & Trim(Me.lblcategory.Text) & "' and Doc_No < 99000000  and bill_flag=1 and cancel_flag = '0' and Location_Code='" & Trim(txtUnitCode.Text) & "'"
                    End If
                    varHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQry)
                End If
                Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                If UBound(varHelp) <> -1 Then

                    If varHelp(0) <> "0" Then
                        Ctlinvoice.Text = Trim(varHelp(0))
                        Ctlinvoice.Focus()
                    Else
                        MsgBox("No Record Available", MsgBoxStyle.Information, "empower")
                    End If
                End If
        End Select
        'SATISH KESHAERWANI CHANGE
        If Len(Ctlinvoice.Text.Trim) > 0 Then
            If DataExist("Select TOP 1 1 FROM SALESCHALLAN_DTL SC WHERE DOC_NO=" & Ctlinvoice.Text.Trim & " AND EXISTS(SELECT  TOP 1 1  from customer_mst C WHERE C.CUSTOMER_CODE=SC.ACCOUNT_CODE AND  CUSTOMERTYPE_GROUP ='TOYOTA' )") = True Then

                llbtkmlPrintingFlag.Enabled = True
                chktkmlbarcode.Enabled = True
                chktkmlbarcode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                chktkmlbarcode.Checked = False
            Else
                llbtkmlPrintingFlag.Enabled = False
                chktkmlbarcode.Enabled = False
                chktkmlbarcode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                chktkmlbarcode.Checked = False
            End If
        End If
        'SATISH KESHAERWANI CHANGE
        'CHANGED BY PRASHANT RAJPAL ON 15 JULY 2019
        If Len(Ctlinvoice.Text.Trim) > 0 Then
            If DataExist("SELECT TOP 1 1 FROM SALESCHALLAN_DTL SC WHERE UNIT_CODE='" & gstrUNITID & "' AND DOC_NO=" & Ctlinvoice.Text.Trim & " AND EXISTS(SELECT  TOP 1 1  FROM CUSTOMER_MST C WHERE C.UNIT_CODE=SC.UNIT_CODE AND C.CUSTOMER_CODE=SC.ACCOUNT_CODE AND C.UNIT_CODE='" + gstrUNITID + "' AND ALLOWBARCODEPRINTING =1 )") = True Then
                If optInvYes(0).Checked = False Then
                    ChkQrbarcodereprint.Enabled = True
                    ChkQrbarcodereprint.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    ChkQrbarcodereprint.Checked = False
                End If
            Else
                If optInvYes(0).Checked = True Then
                    ChkQrbarcodereprint.Enabled = False
                    ChkQrbarcodereprint.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    ChkQrbarcodereprint.Checked = False
                End If
            End If
        End If
        'CHANGED ENDED BY PRASHANT RAJPAL

        Exit Sub
        Exit Sub 'This is to avoid the execution of the error handler
Err_Handler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub CtlInvoice_KeyPress(ByVal Sender As System.Object, ByVal e As CtlGeneral.KeyPressEventArgs) Handles Ctlinvoice.KeyPress
        Dim KeyAscii As Short = e.KeyAscii
        On Error GoTo Err_Handler
        If KeyAscii = 13 Then
            If Len(Trim(Me.Ctlinvoice.Text)) = 0 Then
                Me.Ctlinvoice.Focus()
            ElseIf Len(Trim(Me.cmbInvType.Text)) > 0 Then
                Call CtlInvoice_Validating(Ctlinvoice, New System.ComponentModel.CancelEventArgs(False))
                'Me.Cmdinvoice.SetFocus
                txtPLA.Enabled = True : txtPLA.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                dtpRemoval.Enabled = True
                dtpRemovalTime.Enabled = True
                ' ctlflCopies.Enabled = True: ctlflCopies.BackColor = glngCOLOR_ENABLED
                txtPLA.Focus()
                Exit Sub
            End If
        End If
        'SATISH KESHAERWANI CHANGE
        If Me.Ctlinvoice.Text.Trim.Length > 0 Then
            If DataExist("Select TOP 1 1 FROM SALESCHALLAN_DTL SC WHERE DOC_NO=" & Me.Ctlinvoice.Text.Trim & " AND EXISTS(SELECT  TOP 1 1  from customer_mst C WHERE C.CUSTOMER_CODE=SC.ACCOUNT_CODE AND CUSTOMERTYPE_GROUP='TOYOTA' )") = True Then
                llblLockPrintingFlag.Enabled = True
                chktkmlbarcode.Enabled = True

            Else
                llblLockPrintingFlag.Enabled = False
                chktkmlbarcode.Enabled = False

            End If
        End If
        'SATISH KESHAERWANI CHANGE

        If DataExist("Select TOP 1 1 FROM SALESCHALLAN_DTL SC WHERE UNIT_CODE='" & gstrUNITID & "' AND DOC_NO=" & mInvNo & " AND EXISTS(SELECT  TOP 1 1  from customer_mst C WHERE C.UNIT_CODE=SC.UNIT_CODE AND C.CUSTOMER_CODE=SC.ACCOUNT_CODE AND C.UNIT_CODE='" + gstrUNITID + "' and allowbarcodeprinting =1 )") = True Then
            If optInvYes(0).Checked = False Then
                ChkQrbarcodereprint.Enabled = True
            End If
        Else
            If optInvYes(0).Checked = True Then
                ChkQrbarcodereprint.Enabled = False
            End If
        End If

        If Ctlinvoice.Text.Length() > 8 Then
            KeyAscii = 0
        ElseIf (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
            KeyAscii = KeyAscii
        Else
            KeyAscii = 0
        End If
        DirectCast(Sender, CtlGeneral).KeyPressKeyascii = KeyAscii
        Exit Sub
Err_Handler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub CtlInvoice_KeyUp(ByVal Sender As System.Object, ByVal e As CtlGeneral.KeyUpEventArgs) Handles Ctlinvoice.KeyUp
        Dim KeyCode As Short = e.KeyCode
        Dim Shift As Short = e.Shift
        On Error GoTo Err_Handler
        If KeyCode = 112 Then
            Call cmdHelp_Click(cmdHelp.Item(2), New System.EventArgs())
        End If

        Exit Sub
Err_Handler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub CtlInvoice_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles Ctlinvoice.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim strSql As String
        'Added for Issue ID eMpro-20090415-30143 Starts
        Dim strAccountCode As String
        'Added for Issue ID eMpro-20090415-30143 Ends

        Dim rsInvoiceType As ClsResultSetDB_Invoice
        On Error GoTo Err_Handler
        If Len(Ctlinvoice.Text) = 0 Then GoTo EventExitSub
        If mblnEOUUnit = True Then
            mStrCustMst = "Select Doc_No,Invoice_type from SalesChallan_Dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_Type <> 'EXP' and"
        Else
            mStrCustMst = "Select Doc_No,Invoice_type from SalesChallan_Dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  "
        End If

        'If mblnEwaybill_Print = True And optInvYes(0).Checked = False Then
        '    If chkprintreprint.Checked = False Then
        '        mStrCustMst += " TOTAL_AMOUNT > " & mdblewaymaximumvalue & " AND  INVOICE_DATE >= '" & mblnEWAY_BILL_STARTDATE & " ' "
        '        mStrCustMst += " AND ISNULL(EWAY_BILL_NO,'')<>'' AND  INVOICE_DATE >= '" & mblnEWAY_BILL_STARTDATE & " ' "
        '        mStrCustMst += "AND NOT EXISTS ( SELECT TOP 1 1 FROM FIRSTTIME_INVOICEPRINTING (NOLOCK) WHERE UNIT_CODE =SALESCHALLAN_DTL.UNIT_CODE AND FIRSTTIME_INVOICEPRINTING.DOC_NO=SALESCHALLAN_DTL.DOC_NO ) and "
        '        mStrCustMst = mStrCustMst & " Location_Code='" & Trim(txtUnitCode.Text) & "'" & " and invoice_type='" & Me.lbldescription.Text & "' and sub_category='" & Me.lblcategory.Text & "' and Doc_No ="
        '    Else
        '        mStrCustMst += " Location_Code='" & Trim(txtUnitCode.Text) & "'" & " and invoice_type='" & Me.lbldescription.Text & "' and sub_category='" & Me.lblcategory.Text & "' and Doc_No ='" & Ctlinvoice.Text & "'"
        '        mStrCustMst += "  and (( INVOICE_DATE < '" & mblnEWAY_BILL_STARTDATE & " '))"
        '        mStrCustMst += " UNION  SELECT CONVERT(VARCHAR,DOC_NO) AS DOC_NO,INVOICE_TYPE FROM  SALESCHALLAN_DTL WHERE UNIT_CODE='" + gstrUNITID + "' AND INVOICE_TYPE='" & Trim(Me.lbldescription.Text) & "' AND SUB_CATEGORY='" & Trim(Me.lblcategory.Text) & "' AND DOC_NO < 99000000 AND BILL_FLAG = '1' AND CANCEL_FLAG = '0' AND LOCATION_CODE='" & Trim(txtUnitCode.Text) & "'"
        '        mStrCustMst += " AND TOTAL_AMOUNT <= " & mdblewaymaximumvalue & " and Doc_No ='" & Ctlinvoice.Text & "'"
        '        mStrCustMst += " UNION  SELECT CONVERT(VARCHAR,DOC_NO) AS DOC_NO,INVOICE_TYPE FROM  SALESCHALLAN_DTL WHERE UNIT_CODE='" + gstrUNITID + "' AND INVOICE_TYPE='" & Trim(Me.lbldescription.Text) & "' AND SUB_CATEGORY='" & Trim(Me.lblcategory.Text) & "' AND DOC_NO < 99000000 AND BILL_FLAG = '1' AND CANCEL_FLAG = '0' AND LOCATION_CODE='" & Trim(txtUnitCode.Text) & "'"
        '        mStrCustMst += " AND TOTAL_AMOUNT > " & mdblewaymaximumvalue & " AND  INVOICE_DATE >= '" & mblnEWAY_BILL_STARTDATE & " ' "
        '        mStrCustMst += " AND ISNULL(EWAY_BILL_NO,'')<>'' AND  INVOICE_DATE >= '" & mblnEWAY_BILL_STARTDATE & " '"
        '        mStrCustMst += " AND  EXISTS ( SELECT TOP 1 1 FROM FIRSTTIME_INVOICEPRINTING (NOLOCK) WHERE UNIT_CODE =SALESCHALLAN_DTL.UNIT_CODE AND FIRSTTIME_INVOICEPRINTING.DOC_NO=SALESCHALLAN_DTL.DOC_NO ) and  Doc_No ="
        '    End If
        'Else
        '    If optInvYes(0).Checked = True Then
        '        mStrCustMst = mStrCustMst & " Location_Code='" & Trim(txtUnitCode.Text) & "'" & " and bill_flag =0 and invoice_type='" & Me.lbldescription.Text & "' and sub_category='" & Me.lblcategory.Text & "' and Doc_No ="
        '    Else
        '        mStrCustMst = mStrCustMst & " Location_Code='" & Trim(txtUnitCode.Text) & "'" & " and bill_flag =1 and invoice_type='" & Me.lbldescription.Text & "' and sub_category='" & Me.lblcategory.Text & "' and Doc_No ="
        '    End If
        'End If

        If optInvYes(0).Checked = True Then
            mStrCustMst = mStrCustMst & " Location_Code='" & Trim(txtUnitCode.Text) & "'" & " and bill_flag =0 and invoice_type='" & Me.lbldescription.Text & "' and sub_category='" & Me.lblcategory.Text & "' and Doc_No ="
        Else
            mStrCustMst = mStrCustMst & " Location_Code='" & Trim(txtUnitCode.Text) & "'" & " and bill_flag =1 and invoice_type='" & Me.lbldescription.Text & "' and sub_category='" & Me.lblcategory.Text & "' and Doc_No ="
        End If

        '102027599
        If mblnEwaybill_Print = True AndAlso optInvYes(0).Checked = False Then
            If mblnEOUUnit = True Then
                If chkprintreprint.Checked = True Then
                    mStrCustMst = "Select S.Doc_No,S.Invoice_type from SalesChallan_Dtl S where S.UNIT_CODE = '" & gstrUnitId & "' and S.Invoice_Type <> 'EXP' "
                    mStrCustMst += " AND S.EWAY_IRN_REQUIRED='N' "
                    mStrCustMst += " AND S.Location_Code='" & Trim(txtUnitCode.Text) & "'" & " and S.bill_flag =1 and S.CANCEL_FLAG = 0 and S.invoice_type='" & Me.lbldescription.Text & "' and S.sub_category='" & Me.lblcategory.Text & "' and S.Doc_No = " & Ctlinvoice.Text & " "
                    mStrCustMst += " UNION Select S.Doc_No,S.Invoice_type from SalesChallan_Dtl S LEFT JOIN SALESCHALLAN_DTL_IRN I ON I.UNIT_CODE=S.UNIT_CODE AND I.DOC_NO=S.DOC_NO where  S.UNIT_CODE = '" & gstrUnitId & "' and S.Invoice_Type <> 'EXP' "
                    mStrCustMst += " AND ((S.EWAY_IRN_REQUIRED='E' AND ISNULL(S.EWAY_BILL_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='I' AND ISNULL(I.IRN_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='B' AND ISNULL(S.EWAY_BILL_NO,'')<>'' AND ISNULL(I.IRN_NO,'')<>'')) "
                    mStrCustMst += " AND S.Location_Code='" & Trim(txtUnitCode.Text) & "'" & " and S.bill_flag =1 and S.CANCEL_FLAG = 0 and S.invoice_type='" & Me.lbldescription.Text & "' and S.sub_category='" & Me.lblcategory.Text & "' "
                    mStrCustMst += " AND EXISTS (SELECT TOP 1 1 FROM FIRSTTIME_INVOICEPRINTING F (NOLOCK) WHERE F.UNIT_CODE =S.UNIT_CODE AND F.DOC_NO=S.DOC_NO) and S.Doc_No = "
                Else
                    mStrCustMst = " Select S.Doc_No,S.Invoice_type from SalesChallan_Dtl S LEFT JOIN SALESCHALLAN_DTL_IRN I ON I.UNIT_CODE=S.UNIT_CODE AND I.DOC_NO=S.DOC_NO where  S.UNIT_CODE = '" & gstrUnitId & "' and S.Invoice_Type <> 'EXP' "
                    mStrCustMst += " AND ((S.EWAY_IRN_REQUIRED='E' AND ISNULL(S.EWAY_BILL_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='I' AND ISNULL(I.IRN_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='B' AND ISNULL(S.EWAY_BILL_NO,'')<>'' AND ISNULL(I.IRN_NO,'')<>'')) "
                    mStrCustMst += " AND S.Location_Code='" & Trim(txtUnitCode.Text) & "'" & " and S.bill_flag =1 and S.CANCEL_FLAG = 0 and S.invoice_type='" & Me.lbldescription.Text & "' and S.sub_category='" & Me.lblcategory.Text & "' "
                    mStrCustMst += " AND NOT EXISTS (SELECT TOP 1 1 FROM FIRSTTIME_INVOICEPRINTING F (NOLOCK) WHERE F.UNIT_CODE =S.UNIT_CODE AND F.DOC_NO=S.DOC_NO) and S.Doc_No = "
                End If
            Else
                If chkprintreprint.Checked = True Then
                    mStrCustMst = "Select S.Doc_No,S.Invoice_type from SalesChallan_Dtl S where S.UNIT_CODE = '" & gstrUnitId & "' "
                    mStrCustMst += " AND S.EWAY_IRN_REQUIRED='N' "
                    mStrCustMst += " AND S.Location_Code='" & Trim(txtUnitCode.Text) & "'" & " and S.bill_flag =1 and S.CANCEL_FLAG = 0 and S.invoice_type='" & Me.lbldescription.Text & "' and S.sub_category='" & Me.lblcategory.Text & "' and S.Doc_No = " & Ctlinvoice.Text & " "
                    mStrCustMst += " UNION Select S.Doc_No,S.Invoice_type from SalesChallan_Dtl S LEFT JOIN SALESCHALLAN_DTL_IRN I ON I.UNIT_CODE=S.UNIT_CODE AND I.DOC_NO=S.DOC_NO where  S.UNIT_CODE = '" & gstrUnitId & "' "
                    mStrCustMst += " AND ((S.EWAY_IRN_REQUIRED='E' AND ISNULL(S.EWAY_BILL_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='I' AND ISNULL(I.IRN_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='B' AND ISNULL(S.EWAY_BILL_NO,'')<>'' AND ISNULL(I.IRN_NO,'')<>'')) "
                    mStrCustMst += " AND S.Location_Code='" & Trim(txtUnitCode.Text) & "'" & " and S.bill_flag =1 and S.CANCEL_FLAG = 0 and S.invoice_type='" & Me.lbldescription.Text & "' and S.sub_category='" & Me.lblcategory.Text & "' "
                    mStrCustMst += " AND EXISTS (SELECT TOP 1 1 FROM FIRSTTIME_INVOICEPRINTING F (NOLOCK) WHERE F.UNIT_CODE =S.UNIT_CODE AND F.DOC_NO=S.DOC_NO) and S.Doc_No = "
                Else
                    mStrCustMst = " Select S.Doc_No,S.Invoice_type from SalesChallan_Dtl S LEFT JOIN SALESCHALLAN_DTL_IRN I ON I.UNIT_CODE=S.UNIT_CODE AND I.DOC_NO=S.DOC_NO where  S.UNIT_CODE = '" & gstrUnitId & "' "
                    mStrCustMst += " AND ((S.EWAY_IRN_REQUIRED='E' AND ISNULL(S.EWAY_BILL_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='I' AND ISNULL(I.IRN_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='B' AND ISNULL(S.EWAY_BILL_NO,'')<>'' AND ISNULL(I.IRN_NO,'')<>'')) "
                    mStrCustMst += " AND S.Location_Code='" & Trim(txtUnitCode.Text) & "'" & " and S.bill_flag =1 and S.CANCEL_FLAG = 0 and S.invoice_type='" & Me.lbldescription.Text & "' and S.sub_category='" & Me.lblcategory.Text & "' "
                    mStrCustMst += " AND NOT EXISTS (SELECT TOP 1 1 FROM FIRSTTIME_INVOICEPRINTING F (NOLOCK) WHERE F.UNIT_CODE =S.UNIT_CODE AND F.DOC_NO=S.DOC_NO) and S.Doc_No = "
                End If
            End If
        End If

        strSql = mStrCustMst & Ctlinvoice.Text
        Me.Ctlinvoice.ExistRecQry = mStrCustMst

        mresult = New ClsResultSetDB_Invoice
        mresult.GetResult(strSql)
        mresult.ResultSetClose()

        'SATISH KESHAERWANI CHANGE
        If Len(Ctlinvoice.Text.Trim) > 0 Then
            If DataExist("Select TOP 1 1 FROM SALESCHALLAN_DTL SC WHERE DOC_NO=" & Ctlinvoice.Text.Trim & " AND EXISTS(SELECT  TOP 1 1  from customer_mst C WHERE C.CUSTOMER_CODE=SC.ACCOUNT_CODE AND  CUSTOMERTYPE_GROUP ='TOYOTA' )") = True Then

                llbtkmlPrintingFlag.Enabled = True
                chktkmlbarcode.Enabled = True
                chktkmlbarcode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                chktkmlbarcode.Checked = False


            Else
                llbtkmlPrintingFlag.Enabled = False
                chktkmlbarcode.Enabled = False
                chktkmlbarcode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                chktkmlbarcode.Checked = False

            End If
        End If
        'SATISH KESHAERWANI CHANGE
        If Len(Ctlinvoice.Text.Trim) > 0 Then
            If DataExist("Select TOP 1 1 FROM SALESCHALLAN_DTL SC WHERE DOC_NO=" & Ctlinvoice.Text.Trim & " AND EXISTS(SELECT  TOP 1 1  from customer_mst C WHERE C.CUSTOMER_CODE=SC.ACCOUNT_CODE AND  ALLOWBARCODEPRINTING=1 )") = True Then
                If optInvYes(0).Checked = False Then
                    ChkQrbarcodereprint.Enabled = True
                Else
                    ChkQrbarcodereprint.Enabled = False
                End If
            End If
        End If
        'SATISH KESHAERWANI CHANGE

        If Len(Ctlinvoice.Text) > 0 Then 'Checking if Item Field is not Blank
            If Ctlinvoice.ExistsRec = True Then 'Checking if the Record Exists
                Me.Cmdinvoice.Focus()
                chkPDFExport.Enabled = True
                cmdpdfpath.Enabled = True
                'Added for Issue ID eMpro-20090415-30143 Starts
                strAccountCode = Find_Value("select account_code from saleschallan_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  doc_no='" & Trim(Me.Ctlinvoice.Text) & "'")
                If Me.lbldescription.Text = "REJ" Then
                    blnlinelevelcustomer = False
                Else
                    blnlinelevelcustomer = Find_Value("SELECT SOUPLD_LINE_LEVEL_SALESORDER FROM CUSTOMER_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & strAccountCode & "'")
                End If
                If optInvYes(0).Checked = False Then
                    'Added by priti on 16 Jan 2025 to skip first time invoice printing in case of PDF generation
                    If UCase(cmbInvType.Text) <> "REJECTION" Then
                        Dim strPDFPath As String = SqlConnectionclass.ExecuteScalar("Select isnull(Invoice_PDFCOPYPATH_PRINT,'') from Customer_mst where unit_code='" & gstrUNITID & "'  and customer_code='" & strAccountCode & "'")
                        txtPDFpath.Text = strPDFPath
                        If Len(strPDFPath.ToString) > 0 Then
                            chkPDFExport.Checked = True
                        Else
                            chkPDFExport.Checked = False
                        End If
                    End If
                    'End by priti on 16 Jan 2025 to skip first time invoice printing in case of PDF generation
                End If

                If optInvYes(0).Checked = True And AllowASNPrinting(strAccountCode) = True Then
                    Me.txtASNNumber.Visible = True
                    Me.txtASNNumber.Enabled = True
                    Me.txtASNNumber.Text = ""
                    Me.txtASNNumber.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    Me.lblASNNumber.Visible = True
                    Me.txtASNNumber.Text = CheckASNExist(Me.Ctlinvoice.Text)        'Get Saved ASN Number
                    Me.txtASNNumber.Focus()
                    Me.dtpASNDatetime.Visible = True
                    Me.lblASNDateTime.Visible = True
                Else
                    Me.txtASNNumber.Visible = False
                    Me.txtASNNumber.Enabled = False
                    Me.txtASNNumber.Text = ""
                    Me.txtASNNumber.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    Me.lblASNNumber.Visible = False
                    Me.dtpASNDatetime.Visible = False
                    Me.lblASNDateTime.Visible = False
                End If
                'Added for Issue ID eMpro-20090415-30143 Ends
            Else
                Cancel = True
                chkPDFExport.Checked = False
                'Added for Issue ID eMpro-20090415-30143 Starts
                If optInvYes(1).Checked = True Then
                    Me.txtASNNumber.Visible = False
                    Me.txtASNNumber.Text = ""
                    Me.txtASNNumber.Enabled = False
                    Me.txtASNNumber.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    Me.lblASNNumber.Visible = False
                    Me.lblASNDateTime.Visible = False
                    Me.dtpASNDatetime.Visible = False
                End If
                'Added for Issue ID eMpro-20090415-30143 Ends
                Ctlinvoice.Text = ""
                Ctlinvoice.Focus()

                GoTo EventExitSub
            End If
        End If
        'changes done by Nisha on 04/09/2003 for taking no effacts in accounts
        If Trim(UCase(cmbInvType.Text)) = "REJECTION" Then
            'Commented for Issue ID eMpro-20090112-25902 Starts
            'If Trim(UCase(CmbCategory.Text)) = "REJECTION" Then
            '    strSql = "Select Cust_ref from salesChallan_dtl where doc_no = " & Ctlinvoice.Text
            '    rsInvoiceType = New ClsResultSetDB_Invoice
            '    rsInvoiceType.GetResult(strSql)
            '    'UPGRADE_WARNING: Couldn't resolve default property of object rsInvoiceType.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            '    If Len(Trim(rsInvoiceType.GetValue("Cust_ref"))) > 0 Then
            '        ' lblaccEffects.Visible = True
            '        '''Changes for Issue Id:19934, on 12 jun 2007
            '        If RejInvOptionalPostingFlag() = True Then
            '            chkAcceffects.Enabled = True
            '        Else
            '            chkAcceffects.Enabled = False
            '        End If
            '    Else
            '        ' lblaccEffects.Visible = False
            '        chkAcceffects.Enabled = False
            '        shpacceffects.Visible = False
            '    End If
            '    rsInvoiceType.ResultSetClose()
            '    'UPGRADE_NOTE: Object rsInvoiceType may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            '    '''rsInvoiceType = Nothing
            'End If
            'Commented for Issue ID eMpro-20090112-25902 Ends
        End If
        'changes ends here on 04/09/2003
        GoTo EventExitSub
Err_Handler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.Cancel = Cancel

        '        Dim Cancel As Boolean = eventArgs.Cancel
        '        Dim strSql As String
        '        Dim strfinalsql As String
        '        'Added for Issue ID eMpro-20090415-30143 Starts
        '        Dim strAccountCode As String
        '        'Added for Issue ID eMpro-20090415-30143 Ends

        '        Dim rsInvoiceType As ClsResultSetDB_Invoice
        '        On Error GoTo Err_Handler
        '        If Len(Ctlinvoice.Text) = 0 Then GoTo EventExitSub
        '        If mblnEOUUnit = True Then
        '            mStrCustMst = "Select Doc_No,Invoice_type from SalesChallan_Dtl SC WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_Type <> 'EXP' and"
        '        Else
        '            mStrCustMst = "Select Doc_No,Invoice_type from SalesChallan_Dtl SC WHERE UNIT_CODE='" + gstrUNITID + "' AND  "
        '        End If
        '        If optInvYes(0).Checked = True Then
        '            mStrCustMst = mStrCustMst & " Location_Code='" & Trim(txtUnitCode.Text) & "'" & " and bill_flag =0 and invoice_type='" & Me.lbldescription.Text & "' and sub_category='" & Me.lblcategory.Text & "' and Doc_No ="
        '        Else
        '            mStrCustMst = mStrCustMst & " Location_Code='" & Trim(txtUnitCode.Text) & "'" & " and bill_flag =1 and invoice_type='" & Me.lbldescription.Text & "' and sub_category='" & Me.lblcategory.Text & "' and Doc_No ="
        '        End If
        '        mStrCustMst += Ctlinvoice.Text
        '        If optInvYes(0).Checked = True Then
        '            strSql = " AND 1 ="
        '            strSql += "CASE WHEN FTS_BARCODE =0 OR FTS_LOCATION='01P3'  THEN 1 Else (SELECT TOP 1 1  FROM SALES_DTL SD ,(SELECT SUM(LABEL_QTY)LABELQTY ,ITEMCODE ,UNIT_CODE ,TEMP_INV_NO AS DOC_NO FROM FTS_LABEL_ISSUE "
        '            strSql += " GROUP BY ITEMCODE ,UNIT_CODE ,TEMP_INV_NO)A WHERE(SD.UNIT_CODE = A.UNIT_CODE)"
        '            strSql += " AND SD.DOC_NO =A.DOC_NO  AND SD.ITEM_CODE =A.ITEMCODE AND SD.UNIT_CODE =SC.UNIT_CODE AND SD.DOC_NO =SC.DOC_NO AND SD.SALES_QUANTITY =A.LABELQTY ) End "
        '        Else
        '            strSql = " AND 1 =1 "
        '        End If

        '        strfinalsql = mStrCustMst + strSql
        '        Me.Ctlinvoice.ExistRecQry = strfinalsql

        '        mresult = New ClsResultSetDB_Invoice
        '        mresult.GetResult(strfinalsql)
        '        If mresult.GetNoRows = 0 Then
        '            mresult.ResultSetClose()
        '            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
        '            Ctlinvoice.Text = ""
        '            Exit Sub
        '        Else
        '            mresult.ResultSetClose()
        '        End If


        '        'SATISH KESHAERWANI CHANGE
        '        If Len(Ctlinvoice.Text.Trim) > 0 Then
        '            If DataExist("Select TOP 1 1 FROM SALESCHALLAN_DTL SC WHERE DOC_NO=" & Ctlinvoice.Text.Trim & " AND EXISTS(SELECT  TOP 1 1  from customer_mst C WHERE C.CUSTOMER_CODE=SC.ACCOUNT_CODE AND  CUSTOMERTYPE_GROUP ='TOYOTA' )") = True Then

        '                llbtkmlPrintingFlag.Enabled = True
        '                chktkmlbarcode.Enabled = True
        '                chktkmlbarcode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        '                chktkmlbarcode.Checked = False
        '            Else
        '                llbtkmlPrintingFlag.Enabled = False
        '                chktkmlbarcode.Enabled = False
        '                chktkmlbarcode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        '                chktkmlbarcode.Checked = False
        '            End If
        '        End If
        '        'SATISH KESHAERWANI CHANGE

        '        If Len(Ctlinvoice.Text) > 0 Then 'Checking if Item Field is not Blank
        '            'If Ctlinvoice.ExistsRec = True Then 'Checking if the Record Exists
        '            Me.Cmdinvoice.Focus()
        '            'Added for Issue ID eMpro-20090415-30143 Starts
        '            strAccountCode = Find_Value("select account_code from saleschallan_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  doc_no='" & Trim(Me.Ctlinvoice.Text) & "'")
        '            If optInvYes(0).Checked = True And AllowASNPrinting(strAccountCode) = True Then
        '                Me.txtASNNumber.Visible = True
        '                Me.txtASNNumber.Enabled = True
        '                Me.txtASNNumber.Text = ""
        '                Me.txtASNNumber.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        '                Me.lblASNNumber.Visible = True
        '                Me.txtASNNumber.Text = CheckASNExist(Me.Ctlinvoice.Text)        'Get Saved ASN Number
        '                Me.txtASNNumber.Focus()
        '                Me.dtpASNDatetime.Visible = True
        '                Me.lblASNDateTime.Visible = True
        '            Else
        '                Me.txtASNNumber.Visible = False
        '                Me.txtASNNumber.Enabled = False
        '                Me.txtASNNumber.Text = ""
        '                Me.txtASNNumber.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        '                Me.lblASNNumber.Visible = False
        '                Me.dtpASNDatetime.Visible = False
        '                Me.lblASNDateTime.Visible = False
        '            End If
        '            'Added for Issue ID eMpro-20090415-30143 Ends
        '        Else
        '            Cancel = True
        '            'Added for Issue ID eMpro-20090415-30143 Starts
        '            If optInvYes(1).Checked = True Then
        '                Me.txtASNNumber.Visible = False
        '                Me.txtASNNumber.Text = ""
        '                Me.txtASNNumber.Enabled = False
        '                Me.txtASNNumber.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        '                Me.lblASNNumber.Visible = False
        '                Me.lblASNDateTime.Visible = False
        '                Me.dtpASNDatetime.Visible = False
        '            End If
        '            'Added for Issue ID eMpro-20090415-30143 Ends
        '            Ctlinvoice.Text = ""
        '            Ctlinvoice.Focus()

        '            GoTo EventExitSub
        '            'End If
        '        End If
        '        'changes done by Nisha on 04/09/2003 for taking no effacts in accounts
        '        If Trim(UCase(cmbInvType.Text)) = "REJECTION" Then
        '            'Commented for Issue ID eMpro-20090112-25902 Starts
        '            'If Trim(UCase(CmbCategory.Text)) = "REJECTION" Then
        '            '    strSql = "Select Cust_ref from salesChallan_dtl where doc_no = " & Ctlinvoice.Text
        '            '    rsInvoiceType = New ClsResultSetDB_Invoice
        '            '    rsInvoiceType.GetResult(strSql)
        '            '    'UPGRADE_WARNING: Couldn't resolve default property of object rsInvoiceType.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '            '    If Len(Trim(rsInvoiceType.GetValue("Cust_ref"))) > 0 Then
        '            '        ' lblaccEffects.Visible = True
        '            '        '''Changes for Issue Id:19934, on 12 jun 2007
        '            '        If RejInvOptionalPostingFlag() = True Then
        '            '            chkAcceffects.Enabled = True
        '            '        Else
        '            '            chkAcceffects.Enabled = False
        '            '        End If
        '            '    Else
        '            '        ' lblaccEffects.Visible = False
        '            '        chkAcceffects.Enabled = False
        '            '        shpacceffects.Visible = False
        '            '    End If
        '            '    rsInvoiceType.ResultSetClose()
        '            '    'UPGRADE_NOTE: Object rsInvoiceType may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '            '    '''rsInvoiceType = Nothing
        '            'End If
        '            'Commented for Issue ID eMpro-20090112-25902 Ends
        '        End If
        '        'changes ends here on 04/09/2003
        '        GoTo EventExitSub
        'Err_Handler:
        '        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
        'EventExitSub:
        '        eventArgs.Cancel = Cancel
    End Sub
    'UPGRADE_WARNING: Form event frmMKTTRN0008.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
    Private Sub frmMKTTRN0008_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo Err_Handler
        mdifrmMain.CheckFormName = mintFormIndex
        frmModules.NodeFontBold(Tag) = True
        Exit Sub
Err_Handler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    'UPGRADE_WARNING: Form event frmMKTTRN0008.Deactivate has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
    Private Sub frmMKTTRN0008_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        On Error GoTo Err_Handler
        frmModules.NodeFontBold(Tag) = False
        Exit Sub
Err_Handler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    'UPGRADE_NOTE: Form_Initialize was upgraded to Form_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Private Sub Form_Initialize_Renamed()
        On Error GoTo Err_Handler
        'CHANGES DONE BY AMIT BHATNAGAR ON 25/09/2002*******************************************
        'gobjDB.GetResult "Select EOU_Flag from Company_Mst"
        'CHANGES DONE BY AMIT BHATNAGAR ON 25/09/2002 ENDS HERE*********************************
        gobjDB = New ClsResultSetDB_Invoice
        gobjDB.GetResult("SELECT isnull(EWAY_INV_MAXRANGE,0)EWAY_INV_MAXRANGE ,EWAY_BILL_STARTDATE,EOU_Flag, CustSupp_Inc,InsExc_Excise,postinfin,Excise_RoundOFF ,TOYOTA_MULTIPLESO_ONEPDS_STDATE FROM sales_parameter WHERE UNIT_CODE='" + gstrUNITID + "'")
        'UPGRADE_WARNING: Couldn't resolve default property of object gobjDB.GetValue(EOU_Flag). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If gobjDB.GetValue("EOU_Flag") = True Then
            mStrCustMst = "Select Doc_No,Invoice_type from SalesChallan_Dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_Type <> 'EXP' and Location_Code='" & Trim(txtUnitCode.Text) & "'"
            mblnEOUUnit = True
        Else
            mStrCustMst = "Select Doc_No,Invoice_type from SalesChallan_Dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  Location_Code='" & Trim(txtUnitCode.Text) & "'"
            mblnEOUUnit = False
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object gobjDB.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        mblnAddCustomerMaterial = gobjDB.GetValue("CustSupp_Inc")
        'UPGRADE_WARNING: Couldn't resolve default property of object gobjDB.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        mblnInsuranceFlag = gobjDB.GetValue("InsExc_Excise")
        'UPGRADE_WARNING: Couldn't resolve default property of object gobjDB.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        mblnpostinfin = gobjDB.GetValue("postinfin")
        'UPGRADE_WARNING: Couldn't resolve default property of object gobjDB.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        mblnExciseRoundOFFFlag = gobjDB.GetValue("Excise_RoundOFF")
        mblnEWAY_BILL_STARTDATE = gobjDB.GetValue("EWAY_BILL_STARTDATE")
        mdblewaymaximumvalue = gobjDB.GetValue("EWAY_INV_MAXRANGE")
        mblnTOYOTA_MULTIPLESO_ONEPDS_STDATE = gobjDB.GetValue("TOYOTA_MULTIPLESO_ONEPDS_STDATE")

        gobjDB.ResultSetClose()
        gobjDB = Nothing
        Exit Sub
Err_Handler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub frmMKTTRN0008_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0 Then Call ctlFormHeader1_Click(ctlFormHeader1, New System.EventArgs()) : Exit Sub
    End Sub
    Private Sub frmMKTTRN0008_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo Err_Handler
        Dim objFinconfRecordset As New ADODB.Recordset
        Dim rsSalesParameter As New ClsResultSetDB_Invoice
        mintFormIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader1.Tag)
        Call FillLabelFromResFile(Me) 'To Fill label description from Resource file
        Call FitToClient(Me, fraInvoice, ctlFormHeader1, Cmdinvoice) 'To fit the form in the MDI
        Call EnableControls(False, Me) 'To Disable controls
        optInvYes(0).Enabled = True : optInvYes(1).Enabled = True
        'cmbInvType.BackColor = glngCOLOR_ENABLED
        'CmbCategory.BackColor = glngCOLOR_ENABLED
        gblnCancelUnload = False
        txtUnitCode.Enabled = True
        txtUnitCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        cmdUnitCodeList.Enabled = True
        'cmbInvType.Enabled = True
        'CmbCategory.Enabled = True
        lbldescription.Visible = False
        lblcategory.Visible = False
        cmdHelp(2).Image = My.Resources.ico111.ToBitmap
        optYes(1).Checked = True
        optInvYes(0).Checked = True
        'dtpRemoval.Format = MSComCtl2.FormatConstants.dtpCustom
        'dtpRemoval.CustomFormat = gstrDateFormat
        'dtpRemoval.Value = GetServerDate()
        'changed by nisha on 18/07/2002 for Export invoice
        'gobjDB.GetResult "Select EOU_Flag from Company_Mst"
        'Me.chkLockPrintingFlag.Enabled = True
        'Code Added By Arshad Ali 18/02/2005
        mblnServiceInvoiceWithoutSO = CBool(Find_Value("SELECT ServiceInvoiceWithoutSO FROM SALES_PARAMETER WHERE UNIT_CODE='" + gstrUNITID + "'"))
        mblnJobWkFormulation = Find_Value("Select JobWorkOnBaseofFormulation from sales_parameter WHERE UNIT_CODE='" + gstrUNITID + "'")
        mblnInvoiceAgainstBarCode = CBool(Find_Value("Select isnull(InvoiceAgainstBarCode,0) as InvoiceAgainstBarCode From Sales_Parameter WHERE UNIT_CODE='" + gstrUNITID + "'"))
        'Added for Issue ID 20918 Starts
        mblnInvoiceMTLSharjah = CBool(Find_Value("Select isnull(InvoiceForMTLSharjah,0) as InvoiceForMTLSharjah from sales_parameter WHERE UNIT_CODE='" + gstrUNITID + "'"))
        'Added for Issue ID 20918 Ends
        'Added for Barcode Issue Starts
        mstrFGDomestic = Find_Value("Select FG_DOMESTIC from BarCode_config_mst WHERE UNIT_CODE='" + gstrUNITID + "'")
        '10902255
        mblnCSMspecificreport = CBool(Find_Value("SELECT CSMspecificreport FROM SALES_PARAMETER WHERE UNIT_CODE='" + gstrUNITID + "'"))
        '10902255
        mblnskipdacinvoicebincheck = CBool(Find_Value("SELECT SKIPDACINVOICEBINCHECK  FROM ProductionConf WHERE UNIT_CODE='" + gstrUNITID + "'"))
        mstrREJ_INVOICE_NEWINVREPORT_STARTDATE = CDate(Find_Value("SELECT REJ_INVOICE_NEWINVREPORT_STARTDATE  FROM SALES_PARAMETER WHERE UNIT_CODE='" + gstrUNITID + "'"))
        'Added for Barcode Issue Ends
        Call Form_Initialize_Renamed()
        mblnCCFlag = False
        If objFinconfRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objFinconfRecordset.Close() : objFinconfRecordset = Nothing
        objFinconfRecordset.Open("SELECT * FROM FIN_CONF WHERE UNIT='" + gstrUNITID + "' AND FUNCTIONALITY='CC_GLOBALFLAG' AND ACTIVE=1", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
        If Not objFinconfRecordset.EOF Then
            mblnCCFlag = True
        Else
            mblnCCFlag = False
        End If
        If objFinconfRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objFinconfRecordset.Close() : objFinconfRecordset = Nothing

        rsSalesParameter.GetResult("SELECT CITRIX_INV_PRONTING_LOC FROM SALES_PARAMETER where Unit_Code='" & gstrUNITID & "' ")
        If rsSalesParameter.GetNoRows > 0 Then
            strCitrix_Inv_Pronting_Loc = rsSalesParameter.GetValue("CITRIX_INV_PRONTING_LOC")
        End If
        rsSalesParameter.ResultSetClose()
        rsSalesParameter = Nothing

        Me.dtpASNDatetime.Enabled = True
        chkprintreprint.Enabled = False
        chkprintreprint.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        'dtpASNDatetime.Value = GetServerDate()

        '102027599
        If CBool(Find_Value("SELECT ISNULL(MAX(CAST(EWAY_BILL_FUNCTIONALITY AS INT)),0) EWAY_BILL_FUNCTIONALITY FROM SALECONF (NOLOCK) WHERE  UNIT_CODE = '" & gstrUnitId & "' AND DATEDIFF(DD,GETDATE(),FIN_START_DATE)<=0  AND DATEDIFF(DD,FIN_END_DATE,GETDATE())<=0 ")) Then
            btnExceptionInvoices.Enabled = True
        Else
            btnExceptionInvoices.Enabled = False
        End If
        txtPDFpath.Text = String.Empty
        Exit Sub

Err_Handler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub selectDataFromSaleConf(ByRef LocationCode As String, ByRef combo As System.Windows.Forms.ComboBox, ByRef feild As String, ByRef invoicetype As String, ByRef pstrCondition As String)
        Dim strSql As String
        Dim rsSaleConf As ClsResultSetDB_Invoice
        Dim intRowCount As Short
        Dim intLoopCount As Short
        On Error GoTo Err_Handler
        strSql = "select Distinct(" & feild & ") from Saleconf WHERE UNIT_CODE='" + gstrUNITID + "' AND  Location_Code='" & LocationCode & "' and Invoice_Type in(" & invoicetype & ") and " & pstrCondition
        rsSaleConf = New ClsResultSetDB_Invoice
        rsSaleConf.GetResult(strSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsSaleConf.GetNoRows > 0 Then
            combo.Items.Clear()
            intRowCount = rsSaleConf.GetNoRows
            VB6.SetItemString(combo, 0, "-None-")
            rsSaleConf.MoveFirst()
            For intLoopCount = 1 To intRowCount
                'UPGRADE_WARNING: Couldn't resolve default property of object rsSaleConf.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                VB6.SetItemString(combo, intLoopCount, rsSaleConf.GetValue(feild))
                rsSaleConf.MoveNext()
            Next intLoopCount
            rsSaleConf.ResultSetClose()
            'UPGRADE_NOTE: Object rsSaleConf may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            rsSaleConf = Nothing
        End If
        Exit Sub
Err_Handler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0008_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error GoTo Err_Handler
        'REFRESH
        'Removing the form name from list
        mdifrmMain.RemoveFormNameFromWindowList = mintFormIndex
        'Setting the corresponding node's tag
        frmModules.NodeFontBold(Tag) = False
        'Closing the recordset
        'Releasing the form reference
        'UPGRADE_NOTE: Object frmMKTTRN0008 may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        Me.Dispose()
        Exit Sub
Err_Handler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Sub ValuetoVariables()
        Dim strSql As String
        Dim rsSalesChallan As ClsResultSetDB_Invoice
        Dim strInvoiceDate As String
        Dim strCustomerCode As String

        On Error GoTo Err_Handler
        'CODE ADDED BY NISHA ON 21/03/2003 FOR FINANCIAL ROLLOVER
        strSql = "select Account_Code,INVOICE_DATE,Exchange_rate from Saleschallan_Dtl WHERE UNIT_CODE='" + gstrUnitId + "' AND  Doc_No =" & Me.Ctlinvoice.Text & "  and Location_Code='" & Trim(txtUnitCode.Text) & "'"
        rsSalesChallan = New ClsResultSetDB_Invoice
        rsSalesChallan.GetResult(strSql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        'UPGRADE_WARNING: Couldn't resolve default property of object rsSalesChallan.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        strInvoiceDate = VB6.Format(rsSalesChallan.GetValue("Invoice_Date"), gstrDateFormat)
        'Added for Issue ID 20918 Starts
        'UPGRADE_WARNING: Couldn't resolve default property of object rsSalesChallan.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        mexchange_rate = IIf(rsSalesChallan.GetValue("Exchange_rate") = "", 1, rsSalesChallan.GetValue("Exchange_rate"))
        strCustomerCode = rsSalesChallan.GetValue("Account_Code")
        rsSalesChallan.ResultSetClose()
        'Added for Issue ID 20918 Ends
        mInvType = Me.lbldescription.Text
        mSubCat = Me.lblcategory.Text
        mInvNo = Val(GenerateInvoiceNo(mInvType, mSubCat, strInvoiceDate, strCustomerCode))
        strSql = " Select Asseccable= isNull(SUM(Accessible_amount),0) from sales_dtl "
        strSql = strSql & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Doc_No =" & Me.Ctlinvoice.Text & " and Location_Code='" & Trim(txtUnitCode.Text) & "'"

        mresult = New ClsResultSetDB_Invoice
        mresult.GetResult(strSql)

        mAssessableValue = mresult.GetValue("Asseccable")
        mresult.ResultSetClose()
        Exit Sub
Err_Handler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Sub updatesalesconfandsaleschallan()
        Dim strSql As String
        Dim rsSalesChallan As ClsResultSetDB_Invoice
        Dim dblInvoiceAmt As Double
        Dim strInvoiceDate As String
        On Error GoTo Err_Handler
        strSql = "select *  from Saleschallan_dtl where  UNIT_CODE='" + gstrUNITID + "' AND   Doc_No = " & Me.Ctlinvoice.Text
        strSql = strSql & " and Invoice_type = '" & mInvType & "'  and  sub_category =  '" & mSubCat & "' and Location_Code='" & Trim(txtUnitCode.Text) & "'"
        rsSalesChallan = New ClsResultSetDB_Invoice
        rsSalesChallan.GetResult(strSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsSalesChallan.GetNoRows > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object rsSalesChallan.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            mAccount_Code = rsSalesChallan.GetValue("Account_Code")
            'UPGRADE_WARNING: Couldn't resolve default property of object rsSalesChallan.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            mCust_Ref = rsSalesChallan.GetValue("Cust_ref")
            'UPGRADE_WARNING: Couldn't resolve default property of object rsSalesChallan.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            mAmendment_No = rsSalesChallan.GetValue("Amendment_No")
            'UPGRADE_WARNING: Couldn't resolve default property of object rsSalesChallan.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            dblInvoiceAmt = Val(rsSalesChallan.GetValue("total_amount"))
            'UPGRADE_WARNING: Couldn't resolve default property of object rsSalesChallan.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            strInvoiceDate = VB6.Format(rsSalesChallan.GetValue("Invoice_Date"), gstrDateFormat)
        End If
        rsSalesChallan.ResultSetClose()
        'UPGRADE_NOTE: Object rsSalesChallan may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        rsSalesChallan = Nothing
        If mblnEOUUnit = True Then
            'Changes Done By Nisha on 30/09/2004
            If UCase(lbldescription.Text) <> "EXP" Then
                If Not mblnSameSeries Then
                    If IsGSTINSAME(mAccount_Code) And lbldescription.Text.ToUpper = "TRF" Then
                        salesconf = "update saleconf set CURRENT_NO_TRF_SAMEGSTIN = " & mSaleConfNo & ", OpenningBal = openningBal - " & mAssessableValue & " WHERE UNIT_CODE='" + gstrUnitId + "' AND   Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
                    Else
                        salesconf = "update saleconf set current_No = " & mSaleConfNo & ", OpenningBal = openningBal - " & mAssessableValue & " WHERE UNIT_CODE='" + gstrUnitId + "' AND   Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
                    End If
                Else
                    If IsGSTINSAME(mAccount_Code) And lbldescription.Text.ToUpper = "TRF" Then
                        salesconf = "update saleconf set CURRENT_NO_TRF_SAMEGSTIN = " & mSaleConfNo & " where  UNIT_CODE='" + gstrUnitId + "' AND  Single_Series = 1  and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & strInvoiceDate & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & strInvoiceDate & "')<=0" & vbCrLf
                        salesconf = salesconf & "update saleconf set OpenningBal = openningBal - " & mAssessableValue & " WHERE UNIT_CODE='" + gstrUnitId + "' AND  Invoice_type <> 'EXP' and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
                    Else
                        salesconf = "update saleconf set current_No = " & mSaleConfNo & " where  UNIT_CODE='" + gstrUnitId + "' AND  Single_Series = 1  and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & strInvoiceDate & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & strInvoiceDate & "')<=0" & vbCrLf
                        salesconf = salesconf & "update saleconf set OpenningBal = openningBal - " & mAssessableValue & " WHERE UNIT_CODE='" + gstrUnitId + "' AND  Invoice_type <> 'EXP' and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
                    End If

                End If
            Else
                If Not mblnSameSeries Then
                    salesconf = "update saleconf set current_No = " & mSaleConfNo & " WHERE UNIT_CODE='" + gstrUnitId + "' AND  Invoice_type = 'EXP' and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
                Else
                    salesconf = "update saleconf set current_No = " & mSaleConfNo & " WHERE UNIT_CODE='" + gstrUnitId + "' AND  Single_Series =1 and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
                End If
                'Changes Ends here
            End If
        Else
            If Not mblnSameSeries Then
                If IsGSTINSAME(mAccount_Code) And lbldescription.Text.ToUpper = "TRF" Then
                    salesconf = "update saleconf set CURRENT_NO_TRF_SAMEGSTIN = " & mSaleConfNo & " where  UNIT_CODE='" + gstrUnitId + "' AND  Invoice_type = '" & Me.lbldescription.Text & "' and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & strInvoiceDate & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & strInvoiceDate & "')<=0"
                Else
                    salesconf = "update saleconf set current_No = " & mSaleConfNo & " where  UNIT_CODE='" + gstrUnitId + "' AND  Invoice_type = '" & Me.lbldescription.Text & "' and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & strInvoiceDate & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & strInvoiceDate & "')<=0"
                End If
            Else
                If IsGSTINSAME(mAccount_Code) And lbldescription.Text.ToUpper = "TRF" Then
                    salesconf = "update saleconf set CURRENT_NO_TRF_SAMEGSTIN = " & mSaleConfNo & " WHERE UNIT_CODE='" + gstrUnitId + "' AND  Single_Series = 1 and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
                Else
                    salesconf = "update saleconf set current_No = " & mSaleConfNo & " WHERE UNIT_CODE='" + gstrUnitId + "' AND  Single_Series = 1 and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
                End If
            End If
        End If

        'saleschallan = "UPDATE SalesChallan_Dtl SET doc_no=" & mInvNo & ", total_amount = " & Round(dblInvoiceAmt, 0) & ", Bill_Flag=1,print_flag = 1 WHERE Doc_No=" & Ctlinvoice.Text & " and Location_Code='" & Trim(TxtUnitcode.Text) & "' " & vbCrLf
        '''Changes doen By Ashutosh on 12 Jun 2007 Issue Id:19934, Save posting flag in invoice.
        Dim intInvoicePostingFlag As Short
        If InvoicePostingFlag() = True Then
            intInvoicePostingFlag = 1
        Else
            intInvoicePostingFlag = 0
        End If
        saleschallan = "UPDATE SalesChallan_Dtl SET doc_no=" & mInvNo & ", Bill_Flag=1,print_flag = 1 , postingFlag=" & intInvoicePostingFlag & ",Payment_Terms='" & mstrCreditTermId & "',Upd_dt=getdate(),Upd_Userid='" & mP_User & "' WHERE  UNIT_CODE='" + gstrUnitId + "' AND  Doc_No=" & Ctlinvoice.Text & " and Location_Code='" & Trim(txtUnitCode.Text) & "' " & vbCrLf

        saleschallan = saleschallan & "UPDATE Sales_Dtl SET doc_no=" & mInvNo & " ,Upd_dt=getdate(),Upd_Userid='" & mP_User & "' WHERE  UNIT_CODE='" + gstrUnitId + "' AND  Doc_No=" & Ctlinvoice.Text & " and Location_Code='" & Trim(txtUnitCode.Text) & "'" & vbCrLf
        '''Changes for Issue Id:19934 end here.
        If UCase(cmbInvType.Text) = "SERVICE INVOICE" And mblnServiceInvoiceWithoutSO Then
            saleschallan = saleschallan & " UPDATE RGP_HDR SET InvoiceNo=" & mInvNo & " WHERE UNIT_CODE='" + gstrUnitId + "' AND  InvoiceNo=" & Ctlinvoice.Text & ""
        End If
        Exit Sub
Err_Handler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Function ValidSelection() As Boolean
        Dim blnInvalidData As Boolean
        Dim strErrMsg As String
        Dim ctlBlank As System.Windows.Forms.Control
        Dim lNo As Integer
        On Error GoTo Err_Handler
        ValidRecord = False
        lNo = 1
        'Call ChangeMousePointer(Screen, vbHourglass)
        'Checking if all details have been entered
        strErrMsg = ResolveResString(10059) & vbCrLf & vbCrLf
        If Len(Trim(txtUnitCode.Text)) = 0 Then
            blnInvalidData = True
            strErrMsg = strErrMsg & vbCrLf & lNo & "." & "Location Code"
            lNo = lNo + 1
            If ctlBlank Is Nothing Then ctlBlank = txtUnitCode
        End If
        If Len(Trim(cmbInvType.Text)) = 0 Or cmbInvType.Text = "-None-" Then
            blnInvalidData = True
            strErrMsg = strErrMsg & vbCrLf & lNo & "." & ResolveResString(60371)
            lNo = lNo + 1
            If ctlBlank Is Nothing Then ctlBlank = cmbInvType
        End If
        If Len(Trim(CmbCategory.Text)) = 0 Or CmbCategory.Text = "-None-" Then
            blnInvalidData = True
            strErrMsg = strErrMsg & vbCrLf & lNo & "." & ResolveResString(60372)
            lNo = lNo + 1
            If ctlBlank Is Nothing Then ctlBlank = CmbCategory
        End If
        If Len(Trim(Ctlinvoice.Text)) = 0 Then
            blnInvalidData = True
            strErrMsg = strErrMsg & vbCrLf & lNo & "." & ResolveResString(60373)
            lNo = lNo + 1
            If ctlBlank Is Nothing Then ctlBlank = Ctlinvoice
        End If
        strErrMsg = VB.Left(strErrMsg, Len(strErrMsg) - 1)
        strErrMsg = strErrMsg & "."
        lNo = lNo + 1
        If blnInvalidData = True Then
            gblnCancelUnload = True
            'Call ChangeMousePointer(Screen, vbDefault)
            Call MsgBox(strErrMsg, MsgBoxStyle.Information, "Error")
            ctlBlank.Focus()
            Exit Function
        End If
        ValidRecord = True
        ValidSelection = True
        Exit Function
Err_Handler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

    Public Sub RefreshForm()
        On Error GoTo ErrHandler
        Call EnableControls(False, Me, True)
        optInvYes(0).Enabled = True : optInvYes(1).Enabled = True : optInvYes(0).Checked = True
        Me.cmbInvType.Enabled = True : Me.cmbInvType.BackColor = System.Drawing.Color.White : Me.cmbInvType.Focus()
        'Me.chkLockPrintingFlag.Enabled = True
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub

    Public Function InvoiceGeneration(ByRef RdAddSold As ReportDocument, ByRef RepPath As String, ByRef Frm As Object) As Boolean

        'Dim RdAddSold As ReportDocument
        'Dim RepPath As String
        'Dim Frm As Object = Nothing
        Dim rsCompMst As ClsResultSetDB_Invoice
        Dim rsGrnHdr As ClsResultSetDB_Invoice
        Dim rsSalesConf As ClsResultSetDB_Invoice
        Dim rsSalesInvoiceDate As ClsResultSetDB_Invoice
        Dim Phone, Range, RegNo, EccNo, Address, Invoice_Rule As String
        Dim CST, PLA, Fax, EMail, UPST, Division As String
        Dim Commissionerate As String
        Dim strSql As String
        Dim strCompMst, DeliveredAdd As String
        Dim strGRNDate As String
        Dim strVendorInvNo As String
        Dim strVendorInvDate As String
        Dim strCustRefForGrn As String
        Dim strSuffix As String
        Dim STRCUSTOMERCODE As String

        Dim strQry As String ''Don
        'Added for Issue ID 20918 Starts
        Dim ExpCode As String
        Dim strInvoiceDate As String
        Dim dblExistingInvNo As Double
        Dim strsql1 As String
        'Added for Issue ID 20918 Ends
        'Added for Issue Id 21551 Starts
        Dim TinNo As String
        Dim blnPrintTinNo As Boolean
        'Added for Issue Id 21551 Ends
        Dim oCmd As ADODB.Command
        Dim strIPAddress As String
        Dim blnprintsuffix As Boolean
        Dim invoiceLength As Short = 6
        Dim strStartingPDSMultipleQry As String
        'Dim DeliveredAdd As String
        Dim dblTCStaxAmt As Double
        Dim strAccountCode As String
        On Error GoTo Err_Handler

        'Changed for Issue ID eMpro(-20090611 - 32362) Starts

        'tcs changes
        dblTCStaxAmt = Trim(Find_Value("SELECT SUM(ISNULL(TCSTAXAMOUNT,0)) AS TCSTAXAMOUNT FROM SALES_DTL WHERE UNIT_CODE='" + gstrUNITID + "'AND DOC_NO='" & Trim(Ctlinvoice.Text) & "'"))

        If DataExist("SELECT TOP 1 1 FROM saleschallan_Dtl WHERE UNIT_CODE = '" & gstrUNITID & "' AND DOC_NO='" & Trim(Ctlinvoice.Text) & "' AND ISNULL(TCSTAXAMOUNT,0) <>" & dblTCStaxAmt) Then
            oCmd = New ADODB.Command
            With oCmd
                .ActiveConnection = mP_Connection
                .CommandTimeout = 0
                .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                .CommandText = "USP_INVOICE_DISTRIBUTE_TCS_AMT_ITEM_WISE"
                .Parameters.Append(.CreateParameter("@UNITCODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                .Parameters.Append(.CreateParameter("@DOC_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, Trim(Ctlinvoice.Text)))
                .Parameters.Append(.CreateParameter("@ERRCODE", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInputOutput))

                .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End With
            If oCmd.Parameters("@ERRCODE").Value <> 0 Then
                MsgBox("Error encountered while Calculating TCS Item level .Please try Again.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, My.Resources.resEmpower.STR100)
                oCmd = Nothing
                Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                Exit Function
            End If
            oCmd = Nothing

        End If

        'tcs changes

        strIPAddress = gstrIpaddressWinSck
        STRCUSTOMERCODE = Trim(Find_Value("SELECT ACCOUNT_CODE FROM SALESCHALLAN_DTL WHERE UNIT_CODE='" + gstrUNITID + "'AND DOC_NO='" & Trim(Ctlinvoice.Text) & "'"))
        If IsGSTINSAME(STRCUSTOMERCODE) = True And UCase(Trim(cmbInvType.Text)) = "TRANSFER INVOICE" And System.IO.File.Exists(My.Application.Info.DirectoryPath & "\Reports\Delivery_Challan_Hilex_GST_A4REPORTS.rpt") And mblncustomerspecificreport = False Then
            oCmd = New ADODB.Command
            With oCmd
                .ActiveConnection = mP_Connection
                .CommandTimeout = 0
                .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                .CommandText = "PRC_INVOICEPRINTING_HILEX"
                .Parameters.Append(.CreateParameter("@UNITCODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                .Parameters.Append(.CreateParameter("@LOC_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 8, Trim(txtUnitCode.Text)))
                .Parameters.Append(.CreateParameter("@DOC_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 8, Trim(Ctlinvoice.Text)))
                .Parameters.Append(.CreateParameter("@INV_TYPE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 3, Trim(Me.lbldescription.Text)))
                .Parameters.Append(.CreateParameter("@INV_SUBTYPE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Me.lblcategory.Text)))
                .Parameters.Append(.CreateParameter("@IP_ADDRESS", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 16, strIPAddress.Trim()))
                .Parameters.Append(.CreateParameter("@ERRCODE", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInputOutput))

                .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End With
            If oCmd.Parameters("@ERRCODE").Value <> 0 Then
                MsgBox("Error encountered while generating data for report.Please try Again.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, My.Resources.resEmpower.STR100)
                oCmd = Nothing
                Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                Exit Function
            End If
            oCmd = Nothing
            strSql = "{TMP_INVOICEPRINT.IP_ADDRESS}='" & strIPAddress.Trim() & "'  and {TMP_INVOICEPRINT.Unit_Code}='" & gstrUNITID & "'"
        Else
            'Changed for Issue ID eMpro-20090112-25902 Starts
            If UCase(Trim(GetPlantName)) = "HILEX" And (((UCase(Trim(cmbInvType.Text))) = "NORMAL INVOICE") Or (UCase((Trim(cmbInvType.Text))) = "JOBWORK INVOICE") Or (UCase((Trim(cmbInvType.Text))) = "REJECTION") Or (UCase((Trim(cmbInvType.Text))) = "SAMPLE INVOICE") Or (UCase((Trim(cmbInvType.Text))) = "TRANSFER INVOICE") Or (UCase((Trim(cmbInvType.Text))) = "SERVICE INVOICE")) And mblncustomerspecificreport = False Then
                If (UCase((Trim(CmbCategory.Text))) = "REJECTION") Or (UCase((Trim(CmbCategory.Text))) = "INPUTS") Or (UCase((Trim(CmbCategory.Text))) = "MISC") Or (UCase((Trim(CmbCategory.Text))) = "ASSETS") Or (UCase((Trim(CmbCategory.Text))) = "SUB ASSEMBLY") Or (UCase((Trim(CmbCategory.Text))) = "COMPONENTS") Or ((UCase((CmbCategory.Text))) = "FINISHED GOODS") Or ((UCase((CmbCategory.Text))) = "TOOLS & DIES") Or (UCase((Trim(CmbCategory.Text))) = "RAW MATERIAL") Or (UCase((Trim(CmbCategory.Text))) = "SERVICE") Or (UCase((Trim(CmbCategory.Text))) = "SCRAP") Then
                    oCmd = New ADODB.Command
                    With oCmd
                        .ActiveConnection = mP_Connection
                        .CommandTimeout = 0
                        .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                        .CommandText = "PRC_INVOICEPRINTING_HILEX"
                        .Parameters.Append(.CreateParameter("@UNITCODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                        .Parameters.Append(.CreateParameter("@LOC_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 8, Trim(txtUnitCode.Text)))
                        .Parameters.Append(.CreateParameter("@DOC_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 8, Trim(Ctlinvoice.Text)))
                        .Parameters.Append(.CreateParameter("@INV_TYPE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 3, Trim(Me.lbldescription.Text)))
                        .Parameters.Append(.CreateParameter("@INV_SUBTYPE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Me.lblcategory.Text)))
                        .Parameters.Append(.CreateParameter("@IP_ADDRESS", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 16, strIPAddress.Trim()))
                        .Parameters.Append(.CreateParameter("@ERRCODE", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInputOutput))

                        .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)


                    End With

                    If oCmd.Parameters("@ERRCODE").Value <> 0 Then
                        MsgBox("Error encountered while generating data for report.Please try Again.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, My.Resources.resEmpower.STR100)
                        oCmd = Nothing
                        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                        Exit Function
                    End If
                    oCmd = Nothing

                    strSql = "{TMP_INVOICEPRINT.IP_ADDRESS}='" & strIPAddress.Trim() & "' and {TMP_INVOICEPRINT.Unit_Code}='" & gstrUNITID & "'"
                Else
                    strSql = "{SalesChallan_Dtl.Location_Code}='" & Trim(txtUnitCode.Text) & "'  and {SalesChallan_Dtl.Unit_Code}='" & gstrUNITID & "'  and {SalesChallan_Dtl.Doc_No} =" & Trim(Ctlinvoice.Text) & " and {SalesChallan_Dtl.Invoice_Type}"
                    strSql = strSql & " = '" & Trim(Me.lbldescription.Text) & "'  and {SalesChallan_Dtl.Sub_Category} = '" & Trim(Me.lblcategory.Text) & "'"
                End If


            Else           '****CHANGES FOR SMIEL BY SIDDHARTH ON 05 OCT 2010******

                If (UCase(Trim(GetPlantName)) = "SMIEL") Or (UCase(Trim(GetPlantName)) = "SUMIT") Then
                    oCmd = New ADODB.Command
                    With oCmd
                        .ActiveConnection = mP_Connection
                        .CommandTimeout = 0
                        .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                        .CommandText = " PRC_INVOICEPRINTING_HILEX"
                        .Parameters.Append(.CreateParameter("@UNITCODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                        .Parameters.Append(.CreateParameter("@LOC_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 8, Trim(txtUnitCode.Text)))
                        .Parameters.Append(.CreateParameter("@DOC_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 8, Trim(Ctlinvoice.Text)))
                        .Parameters.Append(.CreateParameter("@INV_TYPE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 3, Trim(Me.lbldescription.Text)))
                        .Parameters.Append(.CreateParameter("@INV_SUBTYPE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Me.lblcategory.Text)))
                        .Parameters.Append(.CreateParameter("@IP_ADDRESS", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 16, strIPAddress.Trim()))
                        .Parameters.Append(.CreateParameter("@ERRCODE", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInputOutput))
                        .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End With


                    If oCmd.Parameters("@ERRCODE").Value <> 0 Then
                        MsgBox("Error encountered while generating data for report.Please try Again.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, My.Resources.resEmpower.STR100)
                        oCmd = Nothing
                        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                        Exit Function
                    End If
                    oCmd = Nothing
                    strSql = "{TMP_INVOICEPRINT_SMIEL.IP_ADDRESS}='" & strIPAddress.Trim() & "' AND {TMP_INVOICEPRINT_SMIEL.Unit_code}='" & gstrUNITID & "'"
                Else
                    'PRASHANT CHANGED ON 10 OCT 2011 for issue id :10146492
                    If UCase(Trim(GetPlantName)) = "MATE" And (((UCase(Trim(cmbInvType.Text))) = "NORMAL INVOICE") And (UCase((Trim(CmbCategory.Text))) = "FINISHED GOODS")) Then
                        oCmd = New ADODB.Command
                        With oCmd
                            .ActiveConnection = mP_Connection
                            .CommandTimeout = 0
                            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                            .CommandText = "PRC_INVOICEPRINTING_MATE"
                            .Parameters.Append(.CreateParameter("@UnitCode", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                            .Parameters.Append(.CreateParameter("@LOC_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 8, Trim(txtUnitCode.Text)))
                            .Parameters.Append(.CreateParameter("@DOC_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 8, Trim(Ctlinvoice.Text)))
                            .Parameters.Append(.CreateParameter("@INV_TYPE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 3, Trim(Me.lbldescription.Text)))
                            .Parameters.Append(.CreateParameter("@INV_SUBTYPE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Me.lblcategory.Text)))
                            .Parameters.Append(.CreateParameter("@IP_ADDRESS", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 16, strIPAddress.Trim()))
                            .Parameters.Append(.CreateParameter("@ERRCODE", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInputOutput))
                            .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        End With
                        If oCmd.Parameters("@ERRCODE").Value <> 0 Then
                            MsgBox("Error encountered while generating data for report.Please try Again.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, My.Resources.resEmpower.STR100)
                            oCmd = Nothing
                            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                            Exit Function
                        End If
                        oCmd = Nothing
                        strSql = "{TMP_INVOICEPRINT.IP_ADDRESS}='" & strIPAddress.Trim() & "'  and {TMP_INVOICEPRINT.Unit_Code}='" & gstrUNITID & "'"
                    Else
                        'added by priti on 16 Feb 2021 for TATA QR code
                        Dim strPrintMethod = SqlConnectionclass.ExecuteScalar("SELECT isnull(PRINT_METHOD,'') FROM CUSTOMER_MST C WHERE C.UNIT_CODE='" & gstrUNITID & "' AND C.CUSTOMER_CODE='" & STRCUSTOMERCODE & "'")
                        If mblncustomerspecificreport = True And strPrintMethod = "TATA" Then
                            oCmd = New ADODB.Command
                            With oCmd
                                .ActiveConnection = mP_Connection
                                .CommandTimeout = 0
                                .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                                .CommandText = "PRC_INVOICEPRINTING_HILEX"
                                .Parameters.Append(.CreateParameter("@UNITCODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                                .Parameters.Append(.CreateParameter("@LOC_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 8, Trim(txtUnitCode.Text)))
                                .Parameters.Append(.CreateParameter("@DOC_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 8, Trim(Ctlinvoice.Text)))
                                .Parameters.Append(.CreateParameter("@INV_TYPE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 3, Trim(Me.lbldescription.Text)))
                                .Parameters.Append(.CreateParameter("@INV_SUBTYPE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Me.lblcategory.Text)))
                                .Parameters.Append(.CreateParameter("@IP_ADDRESS", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 16, strIPAddress.Trim()))
                                .Parameters.Append(.CreateParameter("@ERRCODE", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInputOutput))
                                .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            End With

                            If oCmd.Parameters("@ERRCODE").Value <> 0 Then
                                MsgBox("Error encountered while generating data for report.Please try Again.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, My.Resources.resEmpower.STR100)
                                oCmd = Nothing
                                Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                                Exit Function
                            End If
                            oCmd = Nothing

                            strSql = "{TMP_INVOICEPRINT.IP_ADDRESS}='" & strIPAddress.Trim() & "' and {TMP_INVOICEPRINT.Unit_Code}='" & gstrUNITID & "'"
                        Else
                            strSql = "{SalesChallan_Dtl.Location_Code}='" & Trim(txtUnitCode.Text) & "'  and {SalesChallan_Dtl.Unit_code}='" & gstrUNITID & "' and {SalesChallan_Dtl.Doc_No} =" & Trim(Ctlinvoice.Text) & " and {SalesChallan_Dtl.Invoice_Type}"
                            strSql = strSql & " = '" & Trim(Me.lbldescription.Text) & "'  and {SalesChallan_Dtl.Sub_Category} = '" & Trim(Me.lblcategory.Text) & "'"
                        End If
                    End If


                    'PRASHANT CHANGED ENDED ON 10 OCT 2011 for issue id :10146492

                End If
                '****END OF CHANGES FOR SMIEL BY SIDDHARTH******
                'strSql = "{SalesChallan_Dtl.Location_Code}='" & Trim(txtUnitCode.Text) & "' and {SalesChallan_Dtl.Doc_No} =" & Trim(Ctlinvoice.Text) & " and {SalesChallan_Dtl.Invoice_Type}"
                'strSql = strSql & " = '" & Trim(Me.lbldescription.Text) & "'  and {SalesChallan_Dtl.Sub_Category} = '" & Trim(Me.lblcategory.Text) & "'"

            End If
        End If

        'Changed for Issue ID eMpro-20090112-25902 Ends

        'Changed for Issue ID eMpro(-20090611 - 32362) Ends



        strCompMst = "Select * from Company_Mst where  UNIT_CODE='" + gstrUNITID + "'"
        rsCompMst = New ClsResultSetDB_Invoice
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
            'Added for Issue ID 20918 Starts

            ExpCode = rsCompMst.GetValue("Exporter_Code")
            'Added for Issue ID 20918 Ends
            'Added for Issue Id 21551 Starts

            TinNo = rsCompMst.GetValue("Tin_no")
            'Added for Issue Id 21551 Ends
            'strRoundAc = rsCompMst.GetValue("Rounding_ac")
        End If
        rsCompMst.ResultSetClose()
        'COMMENTED BY AMIT BHATNAGAR ON 14.09.2002 SINCE ROUND OFFF ACCOUNT WILL NOW COME FROM GLOBAL GL
        ''    '**** to Check if round off account is added OR not
        ''    If Len(Trim(strRoundAc)) = 0 Then
        ''        MsgBox "First Define Round off Account in Company Master", vbInformation, "eMPro"
        ''        InvoiceGeneration = False
        ''        Exit Function
        ''    End If
        ''    '****
        '****************************************************************************************
        'If mbilledFlag = False Then
        Dim rsgrin As New ADODB.Recordset
        Dim strsqlstring As String
        Dim dblGRINQty As Double
        Dim dblAvailableQty As Double
        If optInvYes(1).Checked = False Then

            Call InitializeValues()
            Call ValuetoVariables()

            If mblnEOUUnit = True Then
                If lbldescription.Text <> "EXP" Then
                    If mOpeeningBalance < mAssessableValue Then
                        MsgBox("Opening Balance is Less then Invoice Assessable Value", MsgBoxStyle.Information, "eMPro")
                        InvoiceGeneration = False
                        Exit Function
                    End If
                End If
            End If

            If mblnpostinfin = True Then
                If Not CreateStringForAccounts() Then
                    InvoiceGeneration = False
                    Exit Function
                End If
            End If

            Call updatesalesconfandsaleschallan()

            'If UCase(cmbInvType.Text) <> "SERVICE INVOICE" And Not mblnServiceInvoiceWithoutSO Then
            If Not (UCase(cmbInvType.Text) = "SERVICE INVOICE" And mblnServiceInvoiceWithoutSO) Then
                Call UpdateinSale_Dtl()
            Else
                'Update GRIN Despatch Quantity in case of Service Invoice

                mstrGrinQtyUpdate = ""
                strsqlstring = "SELECT NRGP_NO, GRIN_NO, ITEM_CODE, ITEM_QTY FROM NRGP_GRIN_Dtl WHERE UNIT_CODE='" + gstrUNITID + "'  AND  "
                strsqlstring = strsqlstring & " NRGP_NO IN (SELECT NRGPNoInCaseOfServiceInvoice FROM SALESCHALLAN_DTL WHERE UNIT_CODE='" + gstrUNITID + "' AND  DOC_NO=" & Trim(Ctlinvoice.Text) & " and Location_code='" & txtUnitCode.Text & "')"
                If rsgrin.State = ADODB.ObjectStateEnum.adStateOpen Then
                    rsgrin.Close()
                    'UPGRADE_NOTE: Object rsgrin may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                    rsgrin = Nothing
                End If
                rsgrin.Open(strsqlstring, mP_Connection)

                While Not rsgrin.EOF
                    dblGRINQty = rsgrin.Fields("Item_Qty").Value
                    mstrGrinQtyUpdate = mstrGrinQtyUpdate & " UPDATE GRN_DTL"
                    mstrGrinQtyUpdate = mstrGrinQtyUpdate & " SET DESPATCH_QUANTITY = DESPATCH_QUANTITY + " & dblGRINQty
                    mstrGrinQtyUpdate = mstrGrinQtyUpdate & " FROM GRN_HDR H INNER JOIN GRN_DTL D ON H.DOC_TYPE = D.DOC_TYPE AND H.DOC_NO = D.DOC_NO AND H.FROM_LOCATION = D.FROM_LOCATION "
                    mstrGrinQtyUpdate = mstrGrinQtyUpdate & " WHERE H.UNIT_CODE=D.UNIT_CODE AND H.UNIT_CODE='" + gstrUNITID + "' AND H.DOC_CATEGORY='Z' AND H.QA_AUTHORIZED_CODE IS NOT NULL"
                    mstrGrinQtyUpdate = mstrGrinQtyUpdate & " AND H.DOC_NO='" & rsgrin.Fields("Grin_No").Value & "'"
                    mstrGrinQtyUpdate = mstrGrinQtyUpdate & " AND D.ITEM_CODE='" & rsgrin.Fields("ITEM_CODE").Value & "'" & vbCrLf
                    rsgrin.MoveNext()
                End While

                If rsgrin.State = ADODB.ObjectStateEnum.adStateOpen Then
                    rsgrin.Close()
                    'UPGRADE_NOTE: Object rsgrin may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                    rsgrin = Nothing
                End If
                'Ends here
            End If

            If UCase(lbldescription.Text) = "REJ" Then
                If Len(Trim(mCust_Ref)) > 0 Then
                    Call UpdateGrnHdr(Val(mCust_Ref), mInvNo)
                End If
            End If

            'Added for Issue ID 22286 Starts
            If UCase(lbldescription.Text) = "JOB" Then
                If GetBOMCheckFlagValue("BomCheck_Flag") Then
                    mP_Connection.Execute("DELETE FROM tempCustAnnex WHERE UNIT_CODE='" + gstrUNITID + "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords) ' to delete all the records from table before inserting new one for selected invoice
                    If BomCheck() = False Then
                        InvoiceGeneration = False
                        Exit Function
                    End If
                ElseIf CBool(mblnJobWkFormulation) = True Then
                    If CheckforCustSupplyMaterial(False) = False Then
                        InvoiceGeneration = False
                        Exit Function
                    End If
                End If
            End If
            'Added for Issue ID 22286 Ends
        End If
        'Commented for Issue ID 22286 Starts
        '    'changes one by nisha on 13/05/2003 add condition of bomcheck in sales parameter
        '    If UCase(lbldescription.Caption) = "JOB" Then
        '    'changes ends here 13/05/2003
        '        If GetBOMCheckFlagValue("BomCheck_Flag") Then
        '            mP_Connection.Execute "Truncate Table tempCustAnnex" ' to delete all the records from table before inserting new one for selected invoice
        '            If BomCheck = False Then
        '                InvoiceGeneration = False
        '                Exit Function
        '            End If
        '        ElseIf mblnJobWkFormulation = True Then
        '            If CheckforCustSupplyMaterial(False) = False Then
        '                InvoiceGeneration = False
        '                Exit Function
        '            End If
        '        End If
        '    End If
        'Commented for Issue ID 22286 Ends
        'End If

        Address = gstr_WRK_ADDRESS1 & gstr_WRK_ADDRESS2

        '*******************To Calculate Value of Delivery Address in Case of Delivery Address requird
        '*******************To Calculate Value of consignee address on Parameter basis
        rsCompMst = New ClsResultSetDB_Invoice
        rsCompMst.GetResult("Select ConsigneeDetails from Sales_parameter WHERE UNIT_CODE='" + gstrUNITID + "'")

        'UPGRADE_WARNING: Couldn't resolve default property of object rsCompMst.GetValue(ConsigneeDetails). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If rsCompMst.GetValue("ConsigneeDetails") = False Then
            rsCompMst.ResultSetClose()
            rsCompMst = New ClsResultSetDB_Invoice
            rsCompMst.GetResult("Select a.* ,b.account_code,b.invoice_date from Customer_Mst a, saleschallan_dtl b WHERE A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" + gstrUNITID + "' AND a.Customer_code = b.Account_code and b.Doc_No = " & Ctlinvoice.Text & " and b.Location_Code='" & Trim(txtUnitCode.Text) & "'")
            If rsCompMst.GetNoRows > 0 Then
                strAccountCode = rsCompMst.GetValue("Account_code").ToString.Trim
                strInvoiceDate = rsCompMst.GetValue("invoice_date").ToString
                DeliveredAdd = Trim(rsCompMst.GetValue("Ship_address1"))
                If Len(Trim(DeliveredAdd)) Then

                    DeliveredAdd = Trim(DeliveredAdd) & "," & Trim(rsCompMst.GetValue("Ship_address2"))
                Else

                    DeliveredAdd = Trim(rsCompMst.GetValue("Ship_address2"))
                End If
            End If
        Else
            rsCompMst.ResultSetClose()
            rsCompMst = New ClsResultSetDB_Invoice
            rsCompMst.GetResult("Select ConsigneeAddress1,ConsigneeAddress2,ConsigneeAddress3 from Saleschallan_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  Doc_No = " & Ctlinvoice.Text & " and Location_Code='" & Trim(txtUnitCode.Text) & "'")
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
            '**************Changes Ends Here
        End If

        rsCompMst.ResultSetClose()
        If CBool(Find_Value("select TextPrinting from sales_parameter WHERE UNIT_CODE='" + gstrUNITID + "'")) Then
            InvoiceGeneration = True
            Exit Function
        Else
            If mstrReportFilename = "" Then
                MsgBox("No Report filename selected for the invoice. Invoice cannot be printed", MsgBoxStyle.Information, "eMPro")
                Exit Function
            End If
        End If
        '11 AUG 2017
        'If optInvYes(0).Checked = False And Me.Ctlinvoice.Text <> "" Then
        If Me.Ctlinvoice.Text <> "" Then '' Changed by priti on 20 Nov 2024
            strInvoiceDate = Find_Value("select INVOICE_DATE from Saleschallan_Dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  Doc_No =" & Me.Ctlinvoice.Text & "  and Location_Code='" & Trim(txtUnitCode.Text) & "'")
            If strInvoiceDate <> "" Then
                strInvoiceDate = VB6.Format(strInvoiceDate, gstrDateFormat)
                If mblncustomerspecificreport = True Then
                    RepPath = My.Application.Info.DirectoryPath & "\Reports\" & mstrReportFilename & ".rpt"
                Else
                    mstrReportFilename = Find_Value("SELECT report_filename FROM SaleConf WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_Type='" & lbldescription.Text & "' AND Sub_Type_description ='" & CmbCategory.Text & "' AND Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & strInvoiceDate & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & strInvoiceDate & "')<=0")
                End If
            End If

        End If
        '11 AUG 2017

        RdAddSold = Frm.GetReportDocument()
        If UCase(lbldescription.Text) <> "REJ" Then
            mblnA4reports_invoicewise = Find_Value("SELECT AllowA4Reports FROM SaleConf WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_Type='" & lbldescription.Text & "' AND Sub_Type_description ='" & CmbCategory.Text & "' AND Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & strInvoiceDate & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & strInvoiceDate & "')<=0")
        End If

        '10902255
        strStartingPDSMultipleQry = "set dateformat 'dmy' SELECT top 1 1  FROM  TMP_INVOICEPRINT WHERE UNIT_CODE='" + gstrUNITID + "'"
        strStartingPDSMultipleQry += " and ip_address='" & gstrIpaddressWinSck & "'and doc_no=" & Ctlinvoice.Text & "  AND (( INVOICE_DATE >= '" & mblnTOYOTA_MULTIPLESO_ONEPDS_STDATE & "'))"

        If AllowA4Reports(strAccountCode) = True Then
            mblncustomerlevel_A4report_functionlity = True
        End If
        Dim blnrejinv_fullvalue As Boolean



        If IsGSTINSAME(STRCUSTOMERCODE) = True And UCase(Trim(cmbInvType.Text)) = "TRANSFER INVOICE" And System.IO.File.Exists(My.Application.Info.DirectoryPath & "\Reports\Delivery_Challan_Hilex_GST_A4REPORTS.rpt") Then
            RepPath = My.Application.Info.DirectoryPath & "\Reports\Delivery_Challan_Hilex_GST_A4REPORTS.rpt"
        Else
            If DataExist("SELECT TOP 1 1 FROM TMP_INVOICEPRINT WHERE UNIT_CODE='" + gstrUNITID + "' AND IP_ADDRESS='" & gstrIpaddressWinSck & "' AND CUST_MTRL >0 ") And mblnCSMspecificreport = True Then
                If mblncustomerlevel_A4report_functionlity = True And mblnA4reports_invoicewise = True Then
                    RepPath = My.Application.Info.DirectoryPath & "\Reports\" & mstrReportFilename & "_CSM_A4reports.rpt"
                Else
                    RepPath = My.Application.Info.DirectoryPath & "\Reports\" & mstrReportFilename & "_CSM.rpt"
                End If
                '10902255
            Else
                'If DataExist(strStartingPDSMultipleQry) = True And ((CBool(Find_Value("SELECT ISNULL(PDS_TOYOTA_CUSTOMER,0)as PDS_TOYOTA_CUSTOMER  FROM CUSTOMER_MST WHERE UNIT_CODE = '" & gstrUNITID & "'AND CUSTOMER_CODE ='" & STRCUSTOMERCODE & "'")) = True)) Then
                If DataExist(strStartingPDSMultipleQry) = True And DataExist("SELECT TOP 1 1 FROM CUSTOMER_MST WHERE UNIT_CODE = '" & gstrUNITID & "'AND CUSTOMER_CODE ='" & STRCUSTOMERCODE & "'AND  PDS_TOYOTA_CUSTOMER=1 ") = True Then
                    If mblncustomerlevel_A4report_functionlity = True And mblnA4reports_invoicewise = True Then
                        RepPath = My.Application.Info.DirectoryPath & "\Reports\" & mstrReportFilename & "_MultipleSO_A4reports.rpt"
                    End If
                ElseIf mblncustomerlevel_A4report_functionlity = True And mblnA4reports_invoicewise = True Then
                    RepPath = My.Application.Info.DirectoryPath & "\Reports\" & mstrReportFilename & "_A4reports.rpt"
                ElseIf mblnA4reports_invoicewise = True And UCase(cmbInvType.Text) = "REJECTION" Then
                    RepPath = My.Application.Info.DirectoryPath & "\Reports\" & mstrReportFilename & "_A4reports.rpt"
                Else
                    RepPath = My.Application.Info.DirectoryPath & "\Reports\" & mstrReportFilename & ".rpt"
                End If
            End If
        End If

        If DataExist("SELECT TOP 1 1 FROM SALES_PARAMETER WHERE REJINV_POSTING_WITH_FULLVALUE=1  and UNIT_CODE = '" & gstrUNITID & "'") Then
            blnrejinv_fullvalue = True
        Else
            blnrejinv_fullvalue = False
        End If
        If UCase(cmbInvType.Text) = "REJECTION" And blnrejinv_fullvalue = True And strInvoiceDate >= CDate(mstrREJ_INVOICE_NEWINVREPORT_STARTDATE) Then
            RepPath = My.Application.Info.DirectoryPath & "\Reports\rptRejectionInvoiceGST_HILEX.rpt"
        End If

        RdAddSold.Load(RepPath)

        '' With rptinvoice
        ''  .Reset()
        ''  .DiscardSavedData = True
        '.Connect = mp_connection.ConnectionString
        '' .Connect = gstrREPORTCONNECT
        '' .WindowShowPrintSetupBtn = True
        ''   .WindowShowCloseBtn = Truer
        ''  .WindowShowCancelBtn = True
        ''  .WindowShowPrintBtn = True
        '' .WindowShowExportBtn = True
        '' .WindowShowSearchBtn = True
        ''  .WindowState = Crystal.WindowStateConstants.crptMaximized
        'Changed for Issue ID 20918 Starts
        If mblnInvoiceMTLSharjah = True Then
            RdAddSold.DataDefinition.FormulaFields("Comp_name").Text = "'" + gstrCOMPANY + "'"
            RdAddSold.DataDefinition.FormulaFields("comp_add").Text = "'" + Address + "'"
            RdAddSold.DataDefinition.FormulaFields("exchangerate").Text = "'" + CStr(mexchange_rate) + "'"
            RdAddSold.DataDefinition.FormulaFields("ExpCode").Text = "'" + ExpCode + "'"
            If optInvYes(0).Checked = True Then
                RdAddSold.DataDefinition.FormulaFields("CurrentNo").Text = "'" + CStr(mSaleConfNo) + "'"
            Else

                strsql1 = "select * from Saleschallan_Dtl where  UNIT_CODE='" + gstrUNITID + "' AND  Doc_No =" & Me.Ctlinvoice.Text & "  and Location_Code='" & Trim(txtUnitCode.Text) & "'"
                rsSalesInvoiceDate = New ClsResultSetDB_Invoice
                rsSalesInvoiceDate.GetResult(strsql1, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                'UPGRADE_WARNING: Couldn't resolve default property of object rsSalesInvoiceDate.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                strInvoiceDate = VB6.Format(rsSalesInvoiceDate.GetValue("Invoice_Date"), gstrDateFormat)
                rsSalesInvoiceDate.ResultSetClose()

                rsSalesConf = New ClsResultSetDB_Invoice
                rsSalesConf.GetResult("Select Suffix from SaleConf WHERE UNIT_CODE='" + gstrUNITID + "' AND  Description ='" & cmbInvType.Text & "' AND Location_Code ='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0")
                'UPGRADE_WARNING: Couldn't resolve default property of object rsSalesConf.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                strSuffix = rsSalesConf.GetValue("Suffix")
                rsSalesConf.ResultSetClose()
                '101188073
                If gblnGSTUnit Then
                    If Len(Ctlinvoice.Text) > invoiceLength - 1 Then
                        dblExistingInvNo = Val(Ctlinvoice.Text.Substring(Len(Ctlinvoice.Text) - invoiceLength))
                    Else
                        dblExistingInvNo = Val(Ctlinvoice.Text)
                    End If
                Else
                    If Len(Trim(strSuffix)) > 0 Then
                        If Val(strSuffix) > 0 Then
                            dblExistingInvNo = Val(Mid(CStr(Ctlinvoice.Text), Len(Trim(strSuffix)) + 1))
                        Else
                            dblExistingInvNo = Val(Ctlinvoice.Text)
                        End If
                    Else
                        dblExistingInvNo = Val(Ctlinvoice.Text)
                    End If
                End If
                '101188073
                RdAddSold.DataDefinition.FormulaFields("CurrentNo").Text = "'" + dblExistingInvNo + "'"
            End If
        Else
            If UCase(cmbInvType.Text) <> "JOBWORK INVOICE" Then
                RdAddSold.DataDefinition.FormulaFields("Category").Text = "'" + Me.lblcategory.Text + "'"
            End If

            RdAddSold.DataDefinition.FormulaFields("Registration").Text = "'" + RegNo + "'"
            RdAddSold.DataDefinition.FormulaFields("ECC").Text = "'" + EccNo + "'"
            RdAddSold.DataDefinition.FormulaFields("Range").Text = "'" + Range + "'"
            RdAddSold.DataDefinition.FormulaFields("CompanyName").Text = "'" + gstrCOMPANY + "'"
            RdAddSold.DataDefinition.FormulaFields("CompanyAddress").Text = "'" + Address + "'"
            RdAddSold.DataDefinition.FormulaFields("Phone").Text = "'" + Phone + "'"
            RdAddSold.DataDefinition.FormulaFields("Fax").Text = "'" + Fax + "'"

            If UCase(cmbInvType.Text) <> "JOBWORK INVOICE" Then
                RdAddSold.DataDefinition.FormulaFields("EMail").Text = "'" + EMail + "'"
            End If
            RdAddSold.DataDefinition.FormulaFields("PLA").Text = "'" + PLA + "'"
            RdAddSold.DataDefinition.FormulaFields("UPST").Text = "'" + UPST + "'"
            RdAddSold.DataDefinition.FormulaFields("CST").Text = "'" + CST + "'"
            RdAddSold.DataDefinition.FormulaFields("Division").Text = "'" + Division + "'"
            RdAddSold.DataDefinition.FormulaFields("commissionerate").Text = "'" + Commissionerate + "'"
            RdAddSold.DataDefinition.FormulaFields("InvoiceRule").Text = "'" + Invoice_Rule + "'"
            RdAddSold.DataDefinition.FormulaFields("EOUFlag").Text = "'" + CStr(mblnEOUUnit) + "'"

            If optYes(0).Checked = True Then
                RdAddSold.DataDefinition.FormulaFields("DeliveredAt").Text = "' Delivered At '"
                RdAddSold.DataDefinition.FormulaFields("Address2").Text = "'" + DeliveredAdd + "'"
            Else
                RdAddSold.DataDefinition.FormulaFields("DeliveredAt").Text = "''"
                RdAddSold.DataDefinition.FormulaFields("Address2").Text = "''"
                ''.set_Formulas(18, "Address2=''") 'to pass blanck Address in this case will overwrite this Formula written in Crystal Report for else case
            End If

            'rptinvoice.Formulas(16) = "EOUFlag='" & blnEOUFlag & "'"
            RdAddSold.DataDefinition.FormulaFields("PLADuty").Text = "'" + Trim(txtPLA.Text) + "'"
            RdAddSold.DataDefinition.FormulaFields("InsuranceFlag").Text = CStr(mblnInsuranceFlag)
            RdAddSold.DataDefinition.FormulaFields("StringYear").Text = "'" + Year(GetServerDate).ToString + "'"
            RdAddSold.DataDefinition.FormulaFields("DateOfRemoval").Text = "'" + dtpRemoval.Text + "'"

            If optInvYes(0).Checked = True Then
                ' 10127115 starts
                '    .set_Formulas(27, "InvoiceNo='" & mSaleConfNo & "'")
                RdAddSold.DataDefinition.FormulaFields("InvoiceNo").Text = mInvNo
                ' 10127115 end 
            Else
                strsql1 = "select * from Saleschallan_Dtl where  UNIT_CODE='" + gstrUNITID + "' AND  Doc_No =" & Me.Ctlinvoice.Text & "  and Location_Code='" & Trim(txtUnitCode.Text) & "'"
                rsSalesInvoiceDate = New ClsResultSetDB_Invoice
                rsSalesInvoiceDate.GetResult(strsql1, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                'UPGRADE_WARNING: Couldn't resolve default property of object rsSalesInvoiceDate.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                strInvoiceDate = VB6.Format(rsSalesInvoiceDate.GetValue("Invoice_Date"), gstrDateFormat)
                rsSalesInvoiceDate.ResultSetClose()


                rsSalesConf = New ClsResultSetDB_Invoice
                rsSalesConf.GetResult("Select Suffix from SaleConf WHERE UNIT_CODE='" + gstrUNITID + "' AND  Description ='" & cmbInvType.Text & "' AND Location_Code ='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0")
                'UPGRADE_WARNING: Couldn't resolve default property of object rsSalesConf.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                strSuffix = rsSalesConf.GetValue("Suffix")
                rsSalesConf.ResultSetClose()

                blnprintsuffix = CBool(Find_Value("select isnull(printsuffix,0)as printsuffix from sales_parameter WHERE UNIT_CODE='" + gstrUNITID + "'"))

                If blnprintsuffix = False And Len(Trim(strSuffix)) Then
                    'If Len(Trim(strSuffix)) > 0 Then
                    If Val(strSuffix) > 0 Then
                        dblExistingInvNo = Val(Mid(CStr(Ctlinvoice.Text), Len(Trim(strSuffix)) + 1))
                    Else
                        dblExistingInvNo = Val(Ctlinvoice.Text)
                    End If
                Else
                    dblExistingInvNo = Val(Ctlinvoice.Text)
                End If
                '101188073
                If gblnGSTUnit Then
                    dblExistingInvNo = Val(Ctlinvoice.Text)
                End If
                '101188073
                RdAddSold.DataDefinition.FormulaFields("InvoiceNo").Text = dblExistingInvNo
            End If

            ''---- Added by Davinder on 20-06-2006(issue ID:-18103) to provide the provision of printing the invoice on multiple pages by reading the records per page from the saleconf
            strQry = "Select isnull(SC.RecordsPerPage,7) as RecordsPerPage,isnull(SP.MoreThan7ItemInInvoice,0) as MoreThan7ItemInInvoice"
            strQry = strQry & " From saleschallan_dtl SCD inner join SaleConf SC on SCD.Invoice_Type = SC.Invoice_Type And SCD.Sub_Category = SC.Sub_Type AND SCD.UNIT_CODE = SC.UNIT_CODE"
            strQry = strQry & " inner join Sales_parameter SP on Not (isnull(SP.maruti_ac,'') = SCD.Account_code or isnull(SP.maruti_ac1,'') = SCD.Account_code) AND SP.UNIT_CODE = SCD.UNIT_CODE"
            strQry = strQry & " WHERE   SCD.UNIT_CODE='" + gstrUNITID + "' AND datediff(dd,SCD.Invoice_Date,SC.fin_start_date)<=0 And datediff(dd,SC.fin_end_date,SCD.Invoice_Date)<=0 And SCD.doc_no=" & Ctlinvoice.Text

            rsCompMst = New ClsResultSetDB_Invoice
            Call rsCompMst.GetResult(strQry, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If rsCompMst.RowCount > 0 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object rsCompMst.GetValue(MoreThan7ItemInInvoice). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If rsCompMst.GetValue("MoreThan7ItemInInvoice") = "True" Then

                    RdAddSold.DataDefinition.FormulaFields("RecPerPage").Text = "'" + rsCompMst.GetValue("RecordsPerPage") + "'"
                End If
            End If
            rsCompMst.ResultSetClose()
            ''---- Changes by Davinder end's here

            'Added for Issue Id 21551 Starts
            blnPrintTinNo = CBool(Find_Value("Select isnull(PrintTinNO,0) as PrintTinNO from sales_parameter WHERE UNIT_CODE='" + gstrUNITID + "'"))
            If blnPrintTinNo = True Then
                RdAddSold.DataDefinition.FormulaFields("TinNo").Text = "'" + TinNo + "'"
            End If
            'Added for Issue Id 21551 Ends
            '*********************** added by Nisha on 21/02/2003
            'to fetch value grin date Vend invoice no. and vend invoice date in case of Rejection invoice
            If UCase(cmbInvType.Text) = "REJECTION" Then
                rsGrnHdr = New ClsResultSetDB_Invoice
                strGRNDate = "" : strVendorInvDate = "" : strVendorInvNo = "" : strCustRefForGrn = ""
                rsGrnHdr.GetResult("Select Cust_ref from salesChallan_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  doc_No = " & Ctlinvoice.Text)
                If rsGrnHdr.GetNoRows > 0 Then
                    rsGrnHdr.MoveFirst()
                    'UPGRADE_WARNING: Couldn't resolve default property of object rsGrnHdr.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    strCustRefForGrn = rsGrnHdr.GetValue("Cust_ref")
                End If
                rsGrnHdr.ResultSetClose()
                If Len(Trim(strCustRefForGrn)) > 0 Then
                    rsGrnHdr = New ClsResultSetDB_Invoice
                    rsGrnHdr.GetResult("select grn_date,Invoice_no,Invoice_date from grn_hdr WHERE UNIT_CODE='" + gstrUNITID + "' AND  From_Location ='01R1' and doc_No = " & strCustRefForGrn)
                    If rsGrnHdr.GetNoRows > 0 Then
                        rsGrnHdr.MoveFirst()

                        strGRNDate = Convert.ToDateTime(rsGrnHdr.GetValue("grn_date")).ToString(gstrDateFormat)
                        strVendorInvDate = Convert.ToDateTime(rsGrnHdr.GetValue("invoice_date")).ToString(gstrDateFormat) 'VB.Format(rsGrnHdr.GetValue("invoice_date"), "dd/mm/yyyy")
                        'UPGRADE_WARNING: Couldn't resolve default property of object rsGrnHdr.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        strVendorInvNo = rsGrnHdr.GetValue("Invoice_No")
                    End If
                    rsGrnHdr.ResultSetClose()
                End If
                RdAddSold.DataDefinition.FormulaFields("GrinDate").Text = "'" + strGRNDate + "'"
                RdAddSold.DataDefinition.FormulaFields("GrinInvoiceNo").Text = "'" + strVendorInvNo + "'"
                RdAddSold.DataDefinition.FormulaFields("GrinInvoiceDate").Text = "'" + strVendorInvDate + "'"
                'UPGRADE_NOTE: Object rsGrnHdr may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                '''rsGrnHdr = Nothing
            End If
        End If
        'Changed for Issue ID 20918 Ends

        'AMIT RANA 
        If cmbInvType.Text.Trim.ToUpper = "NORMAL INVOICE" And CmbCategory.Text.Trim.ToUpper = "TRADING GOODS" Then
            RdAddSold.DataDefinition.RecordSelectionFormula = "{TRADING_TMP_INVOICE_PRINT_HDR.UNIT_CODE}='" + gstrUNITID + "' And {TRADING_TMP_INVOICE_PRINT_HDR.IPADDRESS}='" + gstrIpaddressWinSck + "'"
        Else
            RdAddSold.DataDefinition.RecordSelectionFormula = strSql
        End If
        'AMIT RANA

        If CBool(Find_Value("select TextPrinting from sales_parameter where  UNIT_CODE= '" & gstrUNITID & "' ")) Then

        Else
            If mstrReportFilename = "" Then
                MsgBox("No Report filename selected for the invoice. Invoice cannot be printed", MsgBoxStyle.Information, "eMPro")
                Exit Function
            End If
        End If
        ''  .ReportFileName = My.Application.Info.DirectoryPath & "\Reports\" & mstrReportFilename & ".rpt"
        ''.SelectionFormula = strSql
        ''End With
        '''rsCompMst.ResultSetClose()
        'UPGRADE_NOTE: Object rsCompMst may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        '''	rsCompMst = Nothing
        InvoiceGeneration = True
        Exit Function
Err_Handler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

    Public Sub InitializeValues()
        On Error GoTo ErrHandler
        mExDuty = 0 : mInvNo = 0 : mBasicAmt = 0 : msubTotal = 0 : mOtherAmt = 0 : mGrTotal = 0 : mStAmt = 0 : mFrAmt = 0
        mDoc_No = 0 : mCustmtrl = 0 : mAmortization = 0 : mstrAnnex = "" : strupdateGrinhdr = "" : mblnCustSupp = False

        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Public Function ParentQty(ByRef pstrItemCode As String, ByRef pstrfinished As Object) As Double
        On Error GoTo ErrHandler
        Dim strParentQty As String
        Dim rsParentQty As ClsResultSetDB_Invoice
        'Added for Issue ID 21473 Starts
        Dim dblTotalQty As Double
        Dim rsconvertQty As ClsResultSetDB_Invoice
        Dim strPurUOM As String
        Dim strConsUOM As String
        Dim dblconversionfactor As Double
        mblnConversion = True
        'Added for Issue ID 21473 Ends

        strParentQty = "select sum(Gross_Weight) as TotalQty from Bom_Mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  finished_Product_code ='"
        'UPGRADE_WARNING: Couldn't resolve default property of object pstrfinished. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        strParentQty = strParentQty & pstrfinished & "' and rawMaterial_Code ='" & pstrItemCode & "'"
        rsParentQty = New ClsResultSetDB_Invoice
        rsParentQty.GetResult(strParentQty, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        'Added for Issue ID 21473 Starts
        If rsParentQty.GetNoRows > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object rsParentQty.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            dblTotalQty = rsParentQty.GetValue("TotalQty")
        End If
        rsParentQty.ResultSetClose()

        If dblTotalQty > 0 Then
            rsconvertQty = New ClsResultSetDB_Invoice
            strParentQty = "select pur_measure_code,cons_measure_code,isnull(conversion_Ratio,0) as conversion_factor from item_mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  item_code='" & Trim(pstrItemCode) & "'"
            rsconvertQty.GetResult(strParentQty, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            If rsconvertQty.GetNoRows > 0 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object rsconvertQty.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                strPurUOM = rsconvertQty.GetValue("pur_measure_code")
                'UPGRADE_WARNING: Couldn't resolve default property of object rsconvertQty.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                strConsUOM = rsconvertQty.GetValue("cons_measure_code")
                'UPGRADE_WARNING: Couldn't resolve default property of object rsconvertQty.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                dblconversionfactor = rsconvertQty.GetValue("conversion_factor")
                If dblconversionfactor > 0 Then
                    If StrComp(strPurUOM, strConsUOM, CompareMethod.Text) <> 0 And (Len(strPurUOM) > 0 And Len(strConsUOM) > 0) Then
                        ParentQty = dblTotalQty / dblconversionfactor
                    Else
                        ParentQty = dblTotalQty
                    End If
                Else
                    ParentQty = dblTotalQty
                End If
            End If
            'UPGRADE_NOTE: Object rsconvertQty may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            rsconvertQty = Nothing
            'UPGRADE_NOTE: Object rsParentQty may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            ''rsParentQty = Nothing
        End If
        'Added for Issue ID 21473 Ends

        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Public Function InsertUpdateAnnex(ByRef parrCustAnnex As Object, ByRef pstrFinishedItem As Object, ByRef intMaxCount As Short) As Object
        Dim intLoopCount As Short
        Dim intLoopcount1 As Short
        Dim intMaxLoop As Short
        Dim intPosChar As Short
        Dim intCharCount As Short
        Dim strRef57F4 As String
        Dim strannex As String
        Dim str57f4Date As String
        Dim rsCustAnnex As ClsResultSetDB_Invoice
        Dim rsVandBom As ClsResultSetDB_Invoice
        Dim dblbalanceqty As Double
        Dim blnValue As Boolean

        On Error GoTo ErrHandler
        '****
        For intLoopCount = 0 To intMaxCount
            'UPGRADE_WARNING: Couldn't resolve default property of object parrCustAnnex(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object pstrFinishedItem. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            rsVandBom = New ClsResultSetDB_Invoice
            rsVandBom.GetResult("Select RawMaterial_Code from Vendor_bom WHERE UNIT_CODE='" + gstrUNITID + "' AND  Finish_Product_code = '" & pstrFinishedItem & "' and Vendor_code = '" & strCustCode & "' and rawMaterial_code ='" & parrCustAnnex(0, intLoopCount) & "'")
            If rsVandBom.GetNoRows > 0 Then
                strRef57F4 = Replace(ref57f4, "§", "','")
                strRef57F4 = "'" & strRef57F4 & "'"

                strannex = "Select Balance_qty,Ref57f4_No,ref57f4_Date from CustAnnex_HDR "
                'UPGRADE_WARNING: Couldn't resolve default property of object parrCustAnnex(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                strannex = strannex & " WHERE UNIT_CODE='" + gstrUNITID + "' AND Item_code ='" & parrCustAnnex(0, intLoopCount) & "' and Customer_code ='"
                strannex = strannex & strCustCode & "'"

                If blnFIFOFlag = False Then
                    strannex = strannex & " and Ref57f4_No in (" & strRef57F4 & ") "
                End If
                'Changes for Issue Id 21473 Starts
                'strannex = strannex & " order by ref57f4_Date"
                strannex = strannex & " and Balance_qty>0 order by ref57f4_Date"
                'Changes for Issue Id 21473 Ends

                rsCustAnnex = New ClsResultSetDB_Invoice
                rsCustAnnex.GetResult(strannex)
                intMaxLoop = rsCustAnnex.GetNoRows
                rsCustAnnex.MoveFirst()
                blnValue = True
                For intLoopcount1 = 1 To intMaxLoop
                    If blnValue = True Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object rsCustAnnex.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        strRef57F4 = rsCustAnnex.GetValue("Ref57f4_No")
                        'UPGRADE_WARNING: Couldn't resolve default property of object rsCustAnnex.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        dblbalanceqty = rsCustAnnex.GetValue("Balance_Qty")
                        'UPGRADE_WARNING: Couldn't resolve default property of object rsCustAnnex.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        str57f4Date = rsCustAnnex.GetValue("ref57f4_Date")
                        mstrAnnex = Trim(mstrAnnex) & " Update CustAnnex_HDR "
                        'UPGRADE_WARNING: Couldn't resolve default property of object parrCustAnnex(1, intLoopCount). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If dblbalanceqty < parrCustAnnex(1, intLoopCount) Then
                            mstrAnnex = Trim(mstrAnnex) & " Set Balance_Qty = 0 "
                        Else
                            'UPGRADE_WARNING: Couldn't resolve default property of object parrCustAnnex(1, intLoopCount). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            mstrAnnex = Trim(mstrAnnex) & " Set Balance_Qty = Balance_Qty - " & parrCustAnnex(1, intLoopCount)
                        End If
                        'UPGRADE_WARNING: Couldn't resolve default property of object parrCustAnnex(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        mstrAnnex = mstrAnnex & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Item_code ='" & parrCustAnnex(0, intLoopCount) & "' and Customer_code ='"
                        mstrAnnex = mstrAnnex & strCustCode & "' and Ref57f4_No ='" & strRef57F4 & "' "

                        mstrAnnex = mstrAnnex & "Insert into CustAnnex_dtl (Doc_Ty,"
                        mstrAnnex = mstrAnnex & "Invoice_No,Invoice_Date,ref57f4_Date,Ref57f4_No,"
                        mstrAnnex = mstrAnnex & "Item_Code,Quantity,"
                        mstrAnnex = mstrAnnex & "Customer_Code,"
                        mstrAnnex = mstrAnnex & "Location_Code,Product_Code,Ent_Userid,Ent_dt,"
                        mstrAnnex = mstrAnnex & "Upd_Userid,Upd_dt,UNIT_CODE) values ('O'," & mInvNo & ",GetDate(),'" & str57f4Date & "','"
                        mstrAnnex = mstrAnnex & strRef57F4 & "','" & parrCustAnnex(0, intLoopCount) & "',"
                        'Added for Issue ID 21473 Starts
                        'UPGRADE_WARNING: Couldn't resolve default property of object parrCustAnnex(1, intLoopCount). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If dblbalanceqty < parrCustAnnex(1, intLoopCount) Then
                            mstrAnnex = Trim(mstrAnnex) & dblbalanceqty & ","
                        Else
                            'UPGRADE_WARNING: Couldn't resolve default property of object parrCustAnnex(1, intLoopCount). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            mstrAnnex = Trim(mstrAnnex) & parrCustAnnex(1, intLoopCount) & ","
                        End If
                        'Added for Issue ID 21473 Starts
                        ' jul
                        mstrAnnex = mstrAnnex & "'" & strCustCode & "',"
                        'UPGRADE_WARNING: Couldn't resolve default property of object pstrFinishedItem. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        mstrAnnex = mstrAnnex & "'" & Trim(txtUnitCode.Text) & "','" & pstrFinishedItem & "','" & mP_User & "',GETDATE(),'"
                        mstrAnnex = mstrAnnex & mP_User & "',GETDATE(),'" + gstrUNITID + "')"
                        'UPGRADE_WARNING: Couldn't resolve default property of object parrCustAnnex(1, intLoopCount). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If dblbalanceqty < parrCustAnnex(1, intLoopCount) Then
                            'UPGRADE_WARNING: Couldn't resolve default property of object pstrFinishedItem. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object parrCustAnnex(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            mP_Connection.Execute(" insert into tempCustAnnex (Ref57f4_No,Annex_No,ref57f4_date,Item_code,Quantity,Balance_qty,finishedItem,UNIT_CODE) values ('" & strRef57F4 & "',0,'" & str57f4Date & "','" & parrCustAnnex(0, intLoopCount) & "'," & dblbalanceqty & ",0,'" & pstrFinishedItem & "','" + gstrUNITID + "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            'Uncomment code for Issue ID 21473 Starts
                            'UPGRADE_WARNING: Couldn't resolve default property of object parrCustAnnex(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            parrCustAnnex(1, intLoopCount) = parrCustAnnex(1, intLoopCount) - dblbalanceqty
                            'Uncomment code for Issue ID 21473 Ends
                        Else
                            'UPGRADE_WARNING: Couldn't resolve default property of object pstrFinishedItem. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object parrCustAnnex(1, intLoopCount). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'UPGRADE_WARNING: Couldn't resolve default property of object parrCustAnnex(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            mP_Connection.Execute(" insert into tempCustAnnex (Ref57f4_No,Annex_No,ref57f4_date,Item_code,Quantity,Balance_qty,finishedItem,UNIT_CODE) values ('" & strRef57F4 & "',0,'" & str57f4Date & "','" & parrCustAnnex(0, intLoopCount) & "'," & parrCustAnnex(1, intLoopCount) & "," & dblbalanceqty - parrCustAnnex(1, intLoopCount) & ",'" & pstrFinishedItem & "','" + gstrUNITID + "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            blnValue = False
                            'Uncomment code for Issue ID 21473 Starts
                            'UPGRADE_WARNING: Couldn't resolve default property of object parrCustAnnex(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            parrCustAnnex(1, intLoopCount) = 0
                            'Uncomment code for Issue ID 21473 Ends
                        End If
                        rsCustAnnex.MoveNext()
                    Else
                        Exit For
                    End If
                Next
                rsCustAnnex.ResultSetClose()
            End If
            rsVandBom.ResultSetClose()

        Next


        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Public Function BomCheck() As Boolean
        '*****************************************************
        'Created By     -  Nisha
        'Description    -  to get the values of required items in Sub assambly bom
        'input Variable -  Item Code to Found, reqquantity of Finished Product,row in spread
        '*****************************************************
        Dim intChallanMax As Short
        Dim intSpCurrentRow As Short
        Dim intCurrentItem As Short
        Dim VarFinishedItem As Object
        Dim strRef57F4 As String
        Dim strBomMst As String
        Dim strCustAnnexDtl As String
        'Dim strProcessType As String
        Dim intBomMaxItem As Short
        Dim rsCustAnnexDtl As ClsResultSetDB_Invoice
        Dim rsSalesChallan As ClsResultSetDB_Invoice
        Dim rsVandorBom As ClsResultSetDB_Invoice
        Dim rsItemMst As ClsResultSetDB_Invoice
        Dim dblTotalReqQty As Double
        Dim strchallan As String
        Dim intAnnexMaxCount As Short
        On Error GoTo ErrHandler
        BomCheck = False
        'intSpreadRow = SpChEntry.MaxRows
        rsSalesChallan = New ClsResultSetDB_Invoice
        rsVandorBom = New ClsResultSetDB_Invoice
        rsItemMst = New ClsResultSetDB_Invoice
        inti = 0
        intAnnexMaxCount = 0
        ReDim arrCustAnnex(3, intAnnexMaxCount)
        strchallan = " select a.Account_code,a.ref_Doc_No,a.Fifo_Flag,b.Item_Code,b.Sales_Quantity from "
        strchallan = strchallan & "salesChallan_dtl a,Sales_dtl b where A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" + gstrUNITID + "' AND a.Doc_No = " & Ctlinvoice.Text
        strchallan = strchallan & " and a.Doc_No = b.Doc_no"
        'Loop for Spread
        rsSalesChallan.GetResult(strchallan, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        intChallanMax = rsSalesChallan.GetNoRows
        rsSalesChallan.MoveFirst()
        Dim intArrCount As Short
        Dim blnItemFoundinArray As Boolean ' to be used to check if item already exist in Array arrItem where we are storing all item we found in Cust annex
        If intChallanMax >= 1 Then
            For intSpCurrentRow = 1 To intChallanMax
                'UPGRADE_WARNING: Couldn't resolve default property of object rsSalesChallan.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object VarFinishedItem. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                VarFinishedItem = rsSalesChallan.GetValue("Item_Code")
                'UPGRADE_WARNING: Couldn't resolve default property of object rsSalesChallan.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                strCustCode = rsSalesChallan.GetValue("Account_code")
                'UPGRADE_WARNING: Couldn't resolve default property of object rsSalesChallan.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                dblFinishedQty = rsSalesChallan.GetValue("Sales_quantity")
                'jul
                'UPGRADE_WARNING: Couldn't resolve default property of object rsSalesChallan.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ref57f4 = rsSalesChallan.GetValue("ref_doc_no")
                strRef57F4 = Replace(ref57f4, "§", "','", 1)
                strRef57F4 = "'" & strRef57F4 & "'"
                'UPGRADE_WARNING: Couldn't resolve default property of object rsSalesChallan.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                blnFIFOFlag = rsSalesChallan.GetValue("FIFO_Flag")
                'Changed for Issue ID 21473 Starts
                'strBomMst = "Select RawMaterial_Code,Process_type,Required_qty + Waste_qty "
                strBomMst = "Select RawMaterial_Code,Process_type,Gross_Weight"
                'Changed for Issue ID 21473 Ends
                strBomMst = strBomMst & " As TotalReqQty"
                strBomMst = strBomMst & " from Bom_Mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  Finished_Product_code ='"
                'UPGRADE_WARNING: Couldn't resolve default property of object VarFinishedItem. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                strBomMst = strBomMst & VarFinishedItem & "' Order By Bom_Level"
                rsBomMst = New ClsResultSetDB_Invoice
                rsBomMst.GetResult(strBomMst, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                intBomMaxItem = rsBomMst.GetNoRows
                rsBomMst.MoveFirst()
                If intBomMaxItem > 0 Then ' Item Found in Bom_Mst
                    'UPGRADE_WARNING: Couldn't resolve default property of object VarFinishedItem. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    rsVandorBom.GetResult("Select RawMaterial_Code from Vendor_bom WHERE UNIT_CODE='" + gstrUNITID + "' AND  Finish_Product_code = '" & VarFinishedItem & "' and Vendor_code = '" & strCustCode & "'")
                    If rsVandorBom.GetNoRows > 0 Then
                        'Loop for Parent Items of Items at First lavel
                        For intCurrentItem = 1 To intBomMaxItem
                            strBomItem = ""
                            'UPGRADE_WARNING: Couldn't resolve default property of object rsBomMst.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            strBomItem = rsBomMst.GetValue("RawMaterial_Code")
                            'strProcessType = rsBomMst.GetValue("Process_type")
                            'String for CustAnnex_dtl
                            strCustAnnexDtl = "Select Item_Code,Balance_qty = sum(Balance_qty) from CustAnnex_hdr WHERE UNIT_CODE='" + gstrUNITID + "' AND  Customer_code ='"
                            strCustAnnexDtl = strCustAnnexDtl & Trim(strCustCode) & "'"
                            If blnFIFOFlag = False Then
                                strCustAnnexDtl = strCustAnnexDtl & " and ref57f4_no in("
                                strCustAnnexDtl = strCustAnnexDtl & Trim(strRef57F4) & ")"
                            End If
                            strCustAnnexDtl = strCustAnnexDtl & " and getdate() <= "
                            strCustAnnexDtl = strCustAnnexDtl & " DateAdd(d, 180, ref57f4_date)"
                            strCustAnnexDtl = strCustAnnexDtl & " and Item_code ='" & strBomItem & "' group by Item_code"
                            rsCustAnnexDtl = New ClsResultSetDB_Invoice
                            rsCustAnnexDtl.GetResult(strCustAnnexDtl)
                            If rsCustAnnexDtl.GetNoRows >= 1 Then 'if item Found in Cust Annex
                                'UPGRADE_WARNING: Couldn't resolve default property of object VarFinishedItem. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                rsVandorBom.GetResult("Select RawMaterial_Code from Vendor_bom WHERE UNIT_CODE='" + gstrUNITID + "' AND  Finish_Product_code = '" & VarFinishedItem & "'and RawMaterial_code = '" & strBomItem & "' and Vendor_Code = '" & strCustCode & "'")
                                If rsVandorBom.GetNoRows > 0 Then
                                    'To Remove  that item from string will be used later for checking in case any item is not supplied
                                    '                   strParent = Replace(strParent, Chr(34) & strBomItem & Chr(34), Chr(34) & "Found" & Chr(34), 1, 1)
                                    rsCustAnnexDtl.MoveFirst()
                                    ReDim Preserve arrItem(inti)
                                    ReDim Preserve arrQty(inti)
                                    ReDim Preserve arrReqQty(inti)
                                    dblTotalReqQty = ParentQty(strBomItem, VarFinishedItem)
                                    'Added for Issue ID 21473 Starts
                                    If mblnConversion = False Then
                                        Exit Function
                                    End If
                                    'Added for Issue ID 21473 Ends
                                    If inti > 0 Then
                                        blnItemFoundinArray = False
                                        For intArrCount = 0 To UBound(arrItem) - 1
                                            'if item already exist in array then to sumup required Quantity
                                            'UPGRADE_WARNING: Couldn't resolve default property of object rsCustAnnexDtl.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                            If UCase(Trim(arrItem(intArrCount))) = UCase(rsCustAnnexDtl.GetValue("Item_code")) Then
                                                ' if item already exist in arritem then will sum up its requied Quantity in arrreqQty() and mark blnFoundinarray as true will be used later
                                                blnItemFoundinArray = True
                                                arrReqQty(intArrCount) = arrReqQty(intArrCount) + (dblTotalReqQty * dblFinishedQty)
                                                If arrQty(intArrCount) < arrReqQty(intArrCount) Then 'in case if sum up is less then Quantity suplied in cust annex
                                                    'Changed for Issue ID 22286 Starts
                                                    'MsgBox "Customer Supplied Materail for Item " & arrItem(inti, 0) & "is" & arrQty(inti, 1) & ".", vbOKOnly, ResolveResString(100)
                                                    MsgBox("Customer Supplied Materail for Item " & arrItem(inti) & " is " & arrQty(inti) & ".", MsgBoxStyle.OkOnly, ResolveResString(100))
                                                    'Changed for Issue ID 22286 Ends
                                                    Cmdinvoice.Focus()
                                                    BomCheck = False
                                                    Exit Function
                                                Else
                                                    Call ToGetIteminAcustannex(arrCustAnnex, strBomItem, intAnnexMaxCount, dblTotalReqQty * dblFinishedQty)
                                                End If
                                            End If
                                        Next
                                        If blnItemFoundinArray = False Then
                                            'in case item not found in arrItem with help of blnItemFoundinarray = false then will add new value to Arrays
                                            'UPGRADE_WARNING: Couldn't resolve default property of object rsCustAnnexDtl.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                            arrItem(inti) = rsCustAnnexDtl.GetValue("Item_code")
                                            'UPGRADE_WARNING: Couldn't resolve default property of object rsCustAnnexDtl.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                            arrQty(inti) = rsCustAnnexDtl.GetValue("Balance_qty")
                                            'str57f4Date = rsCustAnnexDtl.GetValue("REF57F4_DATE")
                                            arrReqQty(inti) = dblTotalReqQty * dblFinishedQty
                                            If arrQty(inti) < arrReqQty(inti) Then 'again  check for Quantity requird as compare to supplied in CustAnnex
                                                'Changed for Issue ID 22286 Starts
                                                'MsgBox "Customer Supplied Materail for Item " & arrItem(inti, 0) & "is" & arrQty(inti, 1) & ".", vbOKOnly, ResolveResString(100)
                                                MsgBox("Customer Supplied Materail for Item " & arrItem(inti) & " is " & arrQty(inti) & ".", MsgBoxStyle.OkOnly, ResolveResString(100))
                                                'Changed for Issue ID 22286 Ends
                                                Cmdinvoice.Focus()
                                                BomCheck = False
                                                Exit Function
                                            Else
                                                '*********** for Adding Values in CustAnnex Array
                                                'UPGRADE_WARNING: Couldn't resolve default property of object arrCustAnnex(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                                If Len(Trim(arrCustAnnex(0, intAnnexMaxCount))) > 0 Then
                                                    intAnnexMaxCount = intAnnexMaxCount + 1
                                                    ReDim Preserve arrCustAnnex(3, intAnnexMaxCount)
                                                End If
                                                'UPGRADE_WARNING: Couldn't resolve default property of object rsCustAnnexDtl.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                                'UPGRADE_WARNING: Couldn't resolve default property of object arrCustAnnex(0, intAnnexMaxCount). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                                arrCustAnnex(0, intAnnexMaxCount) = rsCustAnnexDtl.GetValue("Item_code")
                                                'UPGRADE_WARNING: Couldn't resolve default property of object arrCustAnnex(1, intAnnexMaxCount). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                                arrCustAnnex(1, intAnnexMaxCount) = dblTotalReqQty * dblFinishedQty
                                                'UPGRADE_WARNING: Couldn't resolve default property of object rsCustAnnexDtl.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                                'UPGRADE_WARNING: Couldn't resolve default property of object arrCustAnnex(2, intAnnexMaxCount). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                                arrCustAnnex(2, intAnnexMaxCount) = (rsCustAnnexDtl.GetValue("Balance_qty") - (dblTotalReqQty * dblFinishedQty))
                                                '************
                                            End If
                                        End If
                                    Else ' if inti=0 then to add values
                                        'UPGRADE_WARNING: Couldn't resolve default property of object rsCustAnnexDtl.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                        arrItem(inti) = rsCustAnnexDtl.GetValue("Item_code")
                                        'UPGRADE_WARNING: Couldn't resolve default property of object rsCustAnnexDtl.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                        arrQty(inti) = rsCustAnnexDtl.GetValue("Balance_qty")
                                        'str57f4Date = rsCustAnnexDtl.GetValue("REF57F4_DATE")
                                        arrReqQty(inti) = dblTotalReqQty * dblFinishedQty
                                        If arrQty(inti) < arrReqQty(inti) Then 'Again Same Check
                                            'Changed for Issue ID 22286 Starts
                                            'MsgBox "Customer Supplied Materail for Item " & arrItem(inti, 0) & "is" & arrQty(inti, 1) & ".", vbOKOnly, ResolveResString(100)
                                            MsgBox("Customer Supplied Materail for Item " & arrItem(inti) & " is " & arrQty(inti) & ".", MsgBoxStyle.OkOnly, ResolveResString(100))
                                            'Changed for Issue ID 22286 Ends
                                            Cmdinvoice.Focus()
                                            BomCheck = False
                                            Exit Function
                                        Else
                                            '***********for Adding Values in CustAnnex Array
                                            'UPGRADE_WARNING: Couldn't resolve default property of object arrCustAnnex(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                            If Len(Trim(arrCustAnnex(0, intAnnexMaxCount))) > 0 Then
                                                intAnnexMaxCount = intAnnexMaxCount + 1
                                                ReDim Preserve arrCustAnnex(3, intAnnexMaxCount)
                                            End If
                                            'UPGRADE_WARNING: Couldn't resolve default property of object rsCustAnnexDtl.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                            'UPGRADE_WARNING: Couldn't resolve default property of object arrCustAnnex(0, intAnnexMaxCount). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                            arrCustAnnex(0, intAnnexMaxCount) = rsCustAnnexDtl.GetValue("Item_code")
                                            'UPGRADE_WARNING: Couldn't resolve default property of object arrCustAnnex(1, intAnnexMaxCount). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                            arrCustAnnex(1, intAnnexMaxCount) = dblTotalReqQty * dblFinishedQty
                                            'UPGRADE_WARNING: Couldn't resolve default property of object rsCustAnnexDtl.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                            'UPGRADE_WARNING: Couldn't resolve default property of object arrCustAnnex(2, intAnnexMaxCount). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                            arrCustAnnex(2, intAnnexMaxCount) = (rsCustAnnexDtl.GetValue("Balance_qty") - (dblTotalReqQty * dblFinishedQty))
                                            '************
                                        End If
                                    End If
                                End If
                            Else ' if Item Not Found in Cust Annex
                                rsVandorBom = New ClsResultSetDB_Invoice
                                'UPGRADE_WARNING: Couldn't resolve default property of object VarFinishedItem. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                rsVandorBom.GetResult("Select RawMaterial_Code from Vendor_bom WHERE UNIT_CODE='" + gstrUNITID + "' AND  Finish_Product_code = '" & VarFinishedItem & "'and RawMaterial_code = '" & strBomItem & "' and Vendor_Code = '" & strCustCode & "'")
                                If rsVandorBom.GetNoRows > 0 Then
                                    'If strProcessType = "I" Then ' If That Item is has Process Type I in Bom then
                                    MsgBox("Item " & strBomItem & " is not supplied.", MsgBoxStyle.OkOnly, "eMPro")
                                    Cmdinvoice.Focus()
                                    BomCheck = False
                                    Exit Function
                                Else ' if it'Process type is not I then Explore it Again in BOM_Mst
                                    rsItemMst.GetResult("Select Item_Main_grp from Item_Mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  Item_code = '" & strBomItem & "'")
                                    ''''                        If (UCase(rsItemMst.GetValue("Item_Main_grp")) = "R") Or (UCase(rsItemMst.GetValue("Item_Main_grp")) = "C") Then
                                    'UPGRADE_WARNING: Couldn't resolve default property of object rsItemMst.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                    'Changed for Issue ID eMpro-20081023-22907 Starts
                                    'If (UCase(rsItemMst.GetValue("Item_Main_grp")) <> "F") Or (UCase(rsItemMst.GetValue("Item_Main_grp")) <> "S") Then
                                    '    BomCheck = True
                                    'Else
                                    '    'UPGRADE_WARNING: Couldn't resolve default property of object rsBomMst.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                    '    dblFinishedQty = dblFinishedQty * rsBomMst.GetValue("TotalReqQty")
                                    '    'UPGRADE_WARNING: Couldn't resolve default property of object VarFinishedItem. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                    '    If ExploreBom(strBomItem, dblFinishedQty, intSpCurrentRow, strCustCode, ref57f4, intAnnexMaxCount, CStr(VarFinishedItem)) = False Then
                                    '        BomCheck = False
                                    '        Exit Function
                                    '    End If
                                    'End If

                                    If (UCase(rsItemMst.GetValue("Item_Main_grp")) = "F") Or (UCase(rsItemMst.GetValue("Item_Main_grp")) = "S") Then
                                        dblFinishedQty = dblFinishedQty * rsBomMst.GetValue("TotalReqQty")
                                        If ExploreBom(strBomItem, dblFinishedQty, intSpCurrentRow, strCustCode, ref57f4, intAnnexMaxCount, CStr(VarFinishedItem)) = False Then
                                            BomCheck = False
                                            Exit Function
                                        End If
                                    Else
                                        BomCheck = True
                                    End If
                                    'Changed for Issue ID eMpro-20081023-22907 Ends
                                End If
                            End If
                            rsBomMst.MoveNext()
                            inti = inti + 1
                        Next
                        'intSpCurrentRow = intSpCurrentRow + 1 'for next spread item
                        rsSalesChallan.MoveNext()
                    Else
                        'UPGRADE_WARNING: Couldn't resolve default property of object VarFinishedItem. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        MsgBox("No BOM Defind for the Item (" & VarFinishedItem & ") defined in challan", MsgBoxStyle.Information, "eMPro")
                        BomCheck = False
                        Exit Function
                    End If
                Else
                    'UPGRADE_WARNING: Couldn't resolve default property of object VarFinishedItem. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    MsgBox("No Customer BOM Defind for Item (" & VarFinishedItem & ") defined in challan", MsgBoxStyle.Information, "eMPro")
                    BomCheck = False
                    rsBomMst.ResultSetClose()
                    Exit Function
                End If
                rsBomMst.ResultSetClose()

                'jul
                Call InsertUpdateAnnex(arrCustAnnex, VarFinishedItem, intAnnexMaxCount)
                inti = 0
                intAnnexMaxCount = 0
                ReDim arrCustAnnex(3, intAnnexMaxCount)
                ReDim arrItem(inti)
                ReDim arrQty(inti)
                ReDim arrReqQty(inti)
            Next
        End If
        BomCheck = True
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    ''''''#######
    '''''''Public Function BomCheck() As Boolean
    ''''''''*****************************************************
    ''''''''Created By     -  Nisha
    ''''''''Description    -  to get the values of required items in Sub assambly bom
    ''''''''input Variable -  Item Code to Found, reqquantity of Finished Product,row in spread
    ''''''''*****************************************************
    '''''''Dim intChallanMax As Integer
    '''''''Dim intSpCurrentRow As Integer
    '''''''Dim intCurrentItem As Integer
    '''''''Dim VarFinishedItem As Variant
    '''''''Dim strRef57F4 As String
    '''''''Dim strBomMst As String
    '''''''Dim strCustAnnexDtl As String
    ''''''''Dim strProcessType As String
    '''''''Dim intBomMaxItem As Integer
    '''''''Dim rsCustAnnexDtl As ClsResultSetDB_Invoice
    '''''''Dim rsSalesChallan As ClsResultSetDB_Invoice
    '''''''Dim rsVandorBom As ClsResultSetDB_Invoice
    '''''''Dim rsItemMst As ClsResultSetDB_Invoice
    '''''''Dim dblTotalReqQty As Double
    '''''''Dim strchallan As String
    '''''''Dim intAnnexMaxCount As Integer
    '''''''On Error GoTo ErrHandler
    '''''''BomCheck = False
    ''''''''intSpreadRow = SpChEntry.MaxRows
    '''''''Set rsSalesChallan = New ClsResultSetDB_Invoice
    '''''''Set rsVandorBom = New ClsResultSetDB_Invoice
    '''''''Set rsItemMst = New ClsResultSetDB_Invoice
    '''''''inti = 0
    '''''''intAnnexMaxCount = 0
    '''''''ReDim arrCustAnnex(3, intAnnexMaxCount)
    '''''''strchallan = " select a.Account_code,a.ref_Doc_No,a.Fifo_Flag,b.Item_Code,b.Sales_Quantity from "
    '''''''strchallan = strchallan & "salesChallan_dtl a,Sales_dtl b where a.Doc_No = " & Ctlinvoice.Text
    '''''''strchallan = strchallan & " and a.Location_Code = b.Location_Code and a.Doc_No = b.Doc_no and b.Location_Code='" & Trim(txtUnitCode.Text) & "'"
    ''''''''Loop for Spread
    '''''''rsSalesChallan.GetResult strchallan, adOpenKeyset, adLockReadOnly
    '''''''intChallanMax = rsSalesChallan.GetNoRows
    '''''''If intChallanMax >= 1 Then
    '''''''    For intSpCurrentRow = 1 To intChallanMax
    '''''''        VarFinishedItem = rsSalesChallan.GetValue("Item_Code")
    '''''''        strCustCode = rsSalesChallan.GetValue("Account_code")
    '''''''        dblFinishedQty = rsSalesChallan.GetValue("Sales_quantity")
    '''''''        'jul
    '''''''        ref57f4 = rsSalesChallan.GetValue("ref_doc_no")
    '''''''        strRef57F4 = Replace(ref57f4, "§", "','", 1)
    '''''''        strRef57F4 = "'" & strRef57F4 & "'"
    '''''''        blnFIFOFlag = rsSalesChallan.GetValue("FIFO_Flag")
    '''''''        strBomMst = "Select RawMaterial_Code,Process_type,Required_qty + Waste_qty "
    '''''''        strBomMst = strBomMst & " As TotalReqQty"
    '''''''        strBomMst = strBomMst & " from Bom_Mst where Finished_Product_code ='"
    '''''''        strBomMst = strBomMst & VarFinishedItem & "' Order By Bom_Level"
    '''''''        Set rsBomMst = New ClsResultSetDB_Invoice
    '''''''        rsBomMst.GetResult strBomMst, adOpenKeyset, adLockReadOnly
    '''''''        intBomMaxItem = rsBomMst.GetNoRows
    '''''''        rsBomMst.MoveFirst
    '''''''        If intBomMaxItem > 0 Then ' Item Found in Bom_Mst
    '''''''            rsVandorBom.GetResult "Select RawMaterial_Code from Vendor_bom where Finish_Product_code = '" & VarFinishedItem & "' and Vendor_code = '" & strCustCode & "'"
    '''''''            If rsVandorBom.GetNoRows > 0 Then
    '''''''            'Loop for Parent Items of Items at First lavel
    '''''''                For intCurrentItem = 1 To intBomMaxItem
    '''''''                    strBomItem = ""
    '''''''                    strBomItem = rsBomMst.GetValue("RawMaterial_Code")
    '''''''                    'strProcessType = rsBomMst.GetValue("Process_type")
    '''''''                    'String for CustAnnex_dtl
    '''''''                    strCustAnnexDtl = "Select Item_Code,Balance_qty = sum(Balance_qty) from CustAnnex_hdr where Customer_code ='"
    '''''''                    strCustAnnexDtl = strCustAnnexDtl & Trim(strCustCode) & "'"
    '''''''                    If blnFIFOFlag = False Then
    '''''''                        strCustAnnexDtl = strCustAnnexDtl & " and ref57f4_no in("
    '''''''                        strCustAnnexDtl = strCustAnnexDtl & Trim(strRef57F4) & ")"
    '''''''                    End If
    '''''''                    strCustAnnexDtl = strCustAnnexDtl & " and getdate() <= "
    '''''''                    strCustAnnexDtl = strCustAnnexDtl & " DateAdd(d, 180, ref57f4_date)"
    '''''''                    strCustAnnexDtl = strCustAnnexDtl & " and Item_code ='" & strBomItem & "' group by Item_code"
    '''''''                    Set rsCustAnnexDtl = New ClsResultSetDB_Invoice
    '''''''                    rsCustAnnexDtl.GetResult strCustAnnexDtl
    '''''''                    If rsCustAnnexDtl.GetNoRows >= 1 Then 'if item Found in Cust Annex
    '''''''                        rsVandorBom.GetResult "Select RawMaterial_Code from Vendor_bom where Finish_Product_code = '" & VarFinishedItem & "'and RawMaterial_code = '" & strBomItem & "'"
    '''''''                        If rsVandorBom.GetNoRows > 0 Then
    '''''''                        'To Remove  that item from string will be used later for checking in case any item is not supplied
    '''''''        '                   strParent = Replace(strParent, Chr(34) & strBomItem & Chr(34), Chr(34) & "Found" & Chr(34), 1, 1)
    '''''''                            rsCustAnnexDtl.MoveFirst
    '''''''                            ReDim Preserve arrItem(inti)
    '''''''                            ReDim Preserve arrQty(inti)
    '''''''                            ReDim Preserve arrReqQty(inti)
    '''''''                            dblTotalReqQty = ParentQty(strBomItem, VarFinishedItem)
    '''''''                            If inti > 0 Then
    '''''''                            Dim intArrCount As Integer
    '''''''                            Dim blnItemFoundinArray As Boolean ' to be used to check if item already exist in Array arrItem where we are storing all item we found in Cust annex
    '''''''                                blnItemFoundinArray = False
    '''''''                                For intArrCount = 0 To UBound(arrItem) - 1
    '''''''                                'if item already exist in array then to sumup required Quantity
    '''''''                                    If UCase(Trim(arrItem(intArrCount))) = UCase(rsCustAnnexDtl.GetValue("Item_code")) Then
    '''''''                                    ' if item already exist in arritem then will sum up its requied Quantity in arrreqQty() and mark blnFoundinarray as true will be used later
    '''''''                                        blnItemFoundinArray = True
    '''''''                                        arrReqQty(intArrCount) = arrReqQty(intArrCount) + (dblTotalReqQty * dblFinishedQty)
    '''''''                                        If arrQty(intArrCount) < arrReqQty(intArrCount) Then 'in case if sum up is less then Quantity suplied in cust annex
    '''''''                                            MsgBox "Customer Supplied Materail for Item " & arrItem(inti, 0) & "is" & arrQty(inti, 1) & ".", vbOKOnly, "eMPro"
    '''''''                                            Cmdinvoice.SetFocus
    '''''''                                            BomCheck = False
    '''''''                                            Exit Function
    '''''''                                        Else
    '''''''                                            Call ToGetIteminAcustannex(arrCustAnnex, strBomItem, intAnnexMaxCount, (dblTotalReqQty * dblFinishedQty))
    '''''''                                        End If
    '''''''                                    End If
    '''''''                                Next
    '''''''                                If blnItemFoundinArray = False Then
    '''''''                                'in case item not found in arrItem with help of blnItemFoundinarray = false then will add new value to Arrays
    '''''''                                    arrItem(inti) = rsCustAnnexDtl.GetValue("Item_code")
    '''''''                                    arrQty(inti) = rsCustAnnexDtl.GetValue("Balance_qty")
    '''''''                                    'str57f4Date = rsCustAnnexDtl.GetValue("REF57F4_DATE")
    '''''''                                    arrReqQty(inti) = dblTotalReqQty * dblFinishedQty
    '''''''                                    If arrQty(inti) < arrReqQty(inti) Then 'again  check for Quantity requird as compare to supplied in CustAnnex
    '''''''                                        MsgBox "Customer Supplied Materail for Item " & arrItem(inti, 0) & "is" & arrQty(inti, 1) & ".", vbOKOnly, "eMPro"
    '''''''                                         Cmdinvoice.SetFocus
    '''''''                                        BomCheck = False
    '''''''                                        Exit Function
    '''''''                                    Else
    '''''''                                    '*********** for Adding Values in CustAnnex Array
    '''''''                                        If Len(Trim(arrCustAnnex(0, intAnnexMaxCount))) > 0 Then
    '''''''                                            intAnnexMaxCount = intAnnexMaxCount + 1
    '''''''                                            ReDim Preserve arrCustAnnex(3, intAnnexMaxCount)
    '''''''                                        End If
    '''''''                                        arrCustAnnex(0, intAnnexMaxCount) = rsCustAnnexDtl.GetValue("Item_code")
    '''''''                                        arrCustAnnex(1, intAnnexMaxCount) = dblTotalReqQty * dblFinishedQty
    '''''''                                        arrCustAnnex(2, intAnnexMaxCount) = (rsCustAnnexDtl.GetValue("Balance_qty") - (dblTotalReqQty * dblFinishedQty))
    '''''''                                    '************
    '''''''                                    End If
    '''''''                                End If
    '''''''                            Else ' if inti=0 then to add values
    '''''''                                arrItem(inti) = rsCustAnnexDtl.GetValue("Item_code")
    '''''''                                arrQty(inti) = rsCustAnnexDtl.GetValue("Balance_qty")
    '''''''                                'str57f4Date = rsCustAnnexDtl.GetValue("REF57F4_DATE")
    '''''''                                arrReqQty(inti) = dblTotalReqQty * dblFinishedQty
    '''''''                                If arrQty(inti) < arrReqQty(inti) Then 'Again Same Check
    '''''''                                    MsgBox "Customer Supplied Materail for Item " & arrItem(inti, 0) & "is" & arrQty(inti, 1) & ".", vbOKOnly, "eMPro"
    '''''''                                    Cmdinvoice.SetFocus
    '''''''                                    BomCheck = False
    '''''''                                    Exit Function
    '''''''                                Else
    '''''''                                '***********for Adding Values in CustAnnex Array
    '''''''                                    If Len(Trim(arrCustAnnex(0, intAnnexMaxCount))) > 0 Then
    '''''''                                        intAnnexMaxCount = intAnnexMaxCount + 1
    '''''''                                        ReDim Preserve arrCustAnnex(3, intAnnexMaxCount)
    '''''''                                    End If
    '''''''                                    arrCustAnnex(0, intAnnexMaxCount) = rsCustAnnexDtl.GetValue("Item_code")
    '''''''                                    arrCustAnnex(1, intAnnexMaxCount) = dblTotalReqQty * dblFinishedQty
    '''''''                                    arrCustAnnex(2, intAnnexMaxCount) = (rsCustAnnexDtl.GetValue("Balance_qty") - (dblTotalReqQty * dblFinishedQty))
    '''''''                                '************
    '''''''                                End If
    '''''''                            End If
    '''''''                        End If
    '''''''                    Else ' if Item Not Found in Cust Annex
    '''''''                        Set rsVandorBom = New ClsResultSetDB_Invoice
    '''''''                        rsVandorBom.GetResult "Select RawMaterial_Code from Vendor_bom where Finish_Product_code = '" & VarFinishedItem & "'and RawMaterial_code = '" & strBomItem & "'"
    '''''''                        If rsVandorBom.GetNoRows > 0 Then
    '''''''                        'If strProcessType = "I" Then ' If That Item is has Process Type I in Bom then
    '''''''                            MsgBox "Item " & strBomItem & " is not supplied.", vbOKOnly, "eMPro"
    '''''''                            Cmdinvoice.SetFocus
    '''''''                            BomCheck = False
    '''''''                            Exit Function
    '''''''                        Else ' if it'Process type is not I then Explore it Again in BOM_Mst
    '''''''                        'jul
    '''''''                            '####
    '''''''                            rsItemMst.GetResult "Select Item_Main_grp from Item_Mst Where Item_code = '" & strBomItem & "'"
    '''''''                            If (UCase(rsItemMst.GetValue("Item_Main_grp")) = "R") Or (UCase(rsItemMst.GetValue("Item_Main_grp")) = "C") Then
    '''''''                                BomCheck = True
    '''''''                            Else
    ''''''''                                VarFinishedQty = VarFinishedQty * rsBomMst.GetValue("TotalReqQty")
    ''''''''                                If ExploreBom(strBomItem, VarFinishedQty, intSpCurrentRow, CStr(VarFinishedItem)) = False Then
    ''''''''                                    BomCheck = False
    ''''''''                                    Exit Function
    ''''''''                                End If
    '''''''                                dblFinishedQty = dblFinishedQty * rsBomMst.GetValue("TotalReqQty")
    '''''''                                If ExploreBom(strBomItem, dblFinishedQty, intSpCurrentRow, strCustCode, ref57f4, intAnnexMaxCount, CStr(VarFinishedItem)) = False Then
    '''''''                                    BomCheck = False
    '''''''                                    Exit Function
    '''''''                                End If
    '''''''                            End If
    '''''''                        End If
    '''''''                    End If
    '''''''                    rsBomMst.MoveNext
    '''''''                    inti = inti + 1
    '''''''                Next
    '''''''              '  intSpCurrentRow = intSpCurrentRow + 1 'for next spread item
    '''''''            Else
    '''''''                MsgBox "No BOM Defind for the Item (" & VarFinishedItem & ") defined in challan", vbInformation, "eMPro"
    '''''''                BomCheck = False
    '''''''                Exit Function
    '''''''            End If
    '''''''        Else
    '''''''            MsgBox "No Customer BOM Defind for Item (" & VarFinishedItem & ") defined in challan", vbInformation, "eMPro"
    '''''''            BomCheck = False
    '''''''            Exit Function
    '''''''        End If
    '''''''    'jul
    '''''''        Call InsertUpdateAnnex(arrCustAnnex, VarFinishedItem, intAnnexMaxCount)
    '''''''    Next
    '''''''End If
    '''''''BomCheck = True
    '''''''Exit Function
    '''''''ErrHandler:                             'The Error Handling Code Starts here
    '''''''    Call gobjError.RAISEERROR_INVOICE(Err.number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    '''''''End Function
    ''''''#######
    Public Function ExploreBom(ByRef pstrItemCode As String, ByRef pstrFinishedQty As Object, ByRef pstrSPCurrentRow As Object, ByRef pstrCustCode As String, ByRef pstrRef As String, ByRef pintAnnexMaxCount As Short, ByRef pstrFinishedProduct As String) As Boolean
        '*****************************************************
        'Created By     -  Nisha
        'Description    -  to get the values of required items in Sub assambly bom
        'input Variable -  Item Code to Found, reqquantity of Finished Product,row in spread
        '*****************************************************
        Dim strBomMstRaw As String
        Dim rsBomMstRaw As ClsResultSetDB_Invoice
        Dim rsCustAnnexDtl As ClsResultSetDB_Invoice
        Dim rsVandorBom As ClsResultSetDB_Invoice
        Dim rsItemMst As ClsResultSetDB_Invoice
        Dim intBomMaxRaw As Short
        Dim intCurrentRaw As Short
        Dim dblTotalReqQty As Double
        'Dim strProcessType As String
        Dim strCustAnnexDtl As String
        Dim VarFinishedItem As Object
        Dim strref As String
        On Error GoTo ErrHandler
        rsBomMstRaw = New ClsResultSetDB_Invoice
        rsCustAnnexDtl = New ClsResultSetDB_Invoice
        rsItemMst = New ClsResultSetDB_Invoice

        'Added for Issue ID eMpro-20081023-22907 Starts
        VarFinishedItem = pstrItemCode
        'Added for Issue ID eMpro-20081023-22907 Ends

        'Changed for Issue ID 21473 Starts
        'strBomMstRaw = "Select RawMaterial_Code,Required_qty + Waste_qty "
        strBomMstRaw = "Select RawMaterial_Code,Gross_Weight"
        'Changed for Issue ID 21473 Ends
        strBomMstRaw = strBomMstRaw & " As TotalReqQty,Process_Type from Bom_Mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  "
        strBomMstRaw = strBomMstRaw & " item_Code ='" & strBomItem
        strBomMstRaw = strBomMstRaw & "'and finished_product_code ='"
        strBomMstRaw = strBomMstRaw & pstrItemCode & "'"
        rsBomMstRaw.GetResult(strBomMstRaw, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        Dim intArrCount As Short
        Dim blnArrItemFound As Boolean
        If rsBomMstRaw.GetNoRows > 0 Then ' If Item Found in Bom Mst
            intBomMaxRaw = rsBomMstRaw.GetNoRows
            rsBomMstRaw.MoveFirst()
            For intCurrentRaw = 1 To intBomMaxRaw
                'UPGRADE_WARNING: Couldn't resolve default property of object rsBomMstRaw.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                strBomItem = rsBomMstRaw.GetValue("RawMaterial_code")
                'UPGRADE_WARNING: Couldn't resolve default property of object rsBomMstRaw.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                dblTotalReqQty = rsBomMstRaw.GetValue("TotalReqQty")

                'strProcessType = rsBomMstRaw.GetValue("Process_Type")
                'String for CustAnnex_dtl

                'Changed for Issue ID eMpro-20081023-22907 Starts
                'strCustAnnexDtl = "Select Item_Code,Balance_qty,REF57F4_DATE from CustAnnex_hdr where Customer_code ='"
                'strCustAnnexDtl = strCustAnnexDtl & Trim(pstrCustCode) & "'"
                'If blnFIFOFlag = False Then
                '	strref = Replace(pstrRef, "§", "','", 1)
                '	strref = "'" & strref & "'"
                '	strCustAnnexDtl = strCustAnnexDtl & " and ref57f4_no IN ("
                '	strCustAnnexDtl = strCustAnnexDtl & Trim(strref) & ")"
                'End If
                'strCustAnnexDtl = strCustAnnexDtl & "  and getdate() <= "
                'strCustAnnexDtl = strCustAnnexDtl & " DateAdd(d, 180, ref57f4_date)"
                'strCustAnnexDtl = strCustAnnexDtl & " and Item_code ='" & strBomItem & "'"
                strCustAnnexDtl = "Select Item_Code,Balance_qty = sum(Balance_qty) from CustAnnex_hdr WHERE UNIT_CODE='" + gstrUNITID + "' AND  Customer_code ='"
                strCustAnnexDtl = strCustAnnexDtl & Trim(strCustCode) & "'"
                If blnFIFOFlag = False Then
                    strCustAnnexDtl = strCustAnnexDtl & " and ref57f4_no in("
                    strCustAnnexDtl = strCustAnnexDtl & Trim(strref) & ")"
                End If
                strCustAnnexDtl = strCustAnnexDtl & " and getdate() <= "
                strCustAnnexDtl = strCustAnnexDtl & " DateAdd(d, 180, ref57f4_date)"
                strCustAnnexDtl = strCustAnnexDtl & " and Item_code ='" & strBomItem & "' group by Item_code"
                'Changed for Issue ID eMpro-20081023-22907 Ends

                rsCustAnnexDtl.GetResult(strCustAnnexDtl, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                If rsCustAnnexDtl.GetNoRows >= 1 Then 'if item Found in CustAnnex then replace that item from Parant string
                    rsVandorBom = New ClsResultSetDB_Invoice
                    rsVandorBom.GetResult("Select RawMaterial_Code from Vendor_bom WHERE UNIT_CODE='" + gstrUNITID + "' AND  Finish_Product_code = '" & pstrFinishedProduct & "'and RawMaterial_code = '" & strBomItem & "' and Vendor_code = '" & pstrCustCode & "'")
                    If rsVandorBom.GetNoRows > 0 Then
                        rsCustAnnexDtl.MoveFirst()
                        inti = inti + 1
                        ReDim Preserve arrItem(inti)
                        ReDim Preserve arrQty(inti)
                        ReDim Preserve arrReqQty(inti)

                        'Added for Issue ID 21473 Starts
                        dblTotalReqQty = ParentQty(strBomItem, VarFinishedItem)
                        If mblnConversion = False Then
                            Exit Function
                        End If
                        'Added for Issue ID 21473 Ends
                        blnArrItemFound = False
                        For intArrCount = 0 To UBound(arrItem) - 1 'to check if ITem Already there in ArrItem Array
                            'UPGRADE_WARNING: Couldn't resolve default property of object rsCustAnnexDtl.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            If UCase(Trim(arrItem(intArrCount))) = UCase(Trim(rsCustAnnexDtl.GetValue("Item_code"))) Then
                                ' if found then sum up Requird Quantity in array arrReqQty and assign value true to blnArrITemFound
                                blnArrItemFound = True
                                'UPGRADE_WARNING: Couldn't resolve default property of object pstrFinishedQty. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                arrReqQty(intArrCount) = arrReqQty(intArrCount) + (dblTotalReqQty * pstrFinishedQty)
                                If arrQty(intArrCount) < arrReqQty(intArrCount) Then ' to Check with Quantity supplieded in Cust Annex
                                    MsgBox("Customer Supplied Materail for Item " & arrItem(inti) & " is " & arrQty(inti) & " .", MsgBoxStyle.OkOnly, "eMPro")
                                    Cmdinvoice.Focus()
                                    ExploreBom = False
                                    Exit Function
                                Else
                                    'UPGRADE_WARNING: Couldn't resolve default property of object pstrFinishedQty. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                    Call ToGetIteminAcustannex(arrCustAnnex, strBomItem, pintAnnexMaxCount, dblTotalReqQty * pstrFinishedQty)
                                    ExploreBom = True
                                    Exit For
                                End If
                                blnArrItemFound = False
                            End If
                        Next
                        If blnArrItemFound = False Then ' if item not found
                            inti = inti + 1
                            ReDim Preserve arrItem(inti)
                            ReDim Preserve arrQty(inti)
                            ReDim Preserve arrReqQty(inti)
                            'UPGRADE_WARNING: Couldn't resolve default property of object rsCustAnnexDtl.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            arrItem(inti) = rsCustAnnexDtl.GetValue("Item_code")
                            'UPGRADE_WARNING: Couldn't resolve default property of object rsCustAnnexDtl.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            arrQty(inti) = rsCustAnnexDtl.GetValue("Balance_qty")
                            'jul
                            '                str57f4Date = rsCustAnnexDtl.GetValue("REF57F4_DATE")
                            'UPGRADE_WARNING: Couldn't resolve default property of object pstrFinishedQty. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            arrReqQty(inti) = dblTotalReqQty * pstrFinishedQty
                            If arrQty(inti) < arrReqQty(inti) Then
                                MsgBox("Customer Supplied Materail for Item " & arrItem(inti) & " is " & arrQty(inti) & " .", MsgBoxStyle.OkOnly, "eMPro")
                                Cmdinvoice.Focus()
                                ExploreBom = False
                                Exit Function
                            Else
                                '***********for Adding Values in CustAnnex Array
                                'UPGRADE_WARNING: Couldn't resolve default property of object arrCustAnnex(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                If Len(Trim(arrCustAnnex(0, pintAnnexMaxCount))) > 0 Then
                                    pintAnnexMaxCount = pintAnnexMaxCount + 1
                                    ReDim Preserve arrCustAnnex(3, pintAnnexMaxCount)
                                End If
                                'UPGRADE_WARNING: Couldn't resolve default property of object rsCustAnnexDtl.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                'UPGRADE_WARNING: Couldn't resolve default property of object arrCustAnnex(0, pintAnnexMaxCount). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                arrCustAnnex(0, pintAnnexMaxCount) = rsCustAnnexDtl.GetValue("Item_code")
                                'UPGRADE_WARNING: Couldn't resolve default property of object pstrFinishedQty. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                'UPGRADE_WARNING: Couldn't resolve default property of object arrCustAnnex(1, pintAnnexMaxCount). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                arrCustAnnex(1, pintAnnexMaxCount) = (dblTotalReqQty * pstrFinishedQty)
                                'UPGRADE_WARNING: Couldn't resolve default property of object pstrFinishedQty. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                'UPGRADE_WARNING: Couldn't resolve default property of object rsCustAnnexDtl.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                'UPGRADE_WARNING: Couldn't resolve default property of object arrCustAnnex(2, pintAnnexMaxCount). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                arrCustAnnex(2, pintAnnexMaxCount) = (rsCustAnnexDtl.GetValue("Balance_qty") - (dblTotalReqQty * pstrFinishedQty))
                                '*********************
                                ExploreBom = True
                            End If
                        Else
                            '                    arrCustAnnex(0, pintAnnexMaxCount) = rsCustAnnexDtl.GetValue("Item_code")
                            '                    arrCustAnnex(1, pintAnnexMaxCount) = (dblTotalReqQty * pstrFinishedQty)
                            '                    arrCustAnnex(2, pintAnnexMaxCount) = (rsCustAnnexDtl.GetValue("Balance_qty") - (dblTotalReqQty * pstrFinishedQty))
                        End If
                    End If
                Else
                    '            If strProcessType = "I" Then
                    'Added for Issue ID eMpro-20081023-22907 Starts
                    rsVandorBom = New ClsResultSetDB_Invoice
                    'Added for Issue ID eMpro-20081023-22907 Ends

                    rsVandorBom.GetResult("Select RawMaterial_Code from Vendor_bom WHERE UNIT_CODE='" + gstrUNITID + "' AND  Finish_Product_code = '" & pstrItemCode & "'and RawMaterial_code = '" & strBomItem & "' and Vendor_code ='" & pstrCustCode & "'")
                    If rsVandorBom.GetNoRows > 0 Then
                        MsgBox("Item " & strBomItem & " is not supplied.", MsgBoxStyle.OkOnly, "eMPro")
                        Cmdinvoice.Focus()
                        ExploreBom = False
                        Exit Function
                    Else 'if not of Process type I then again Explore
                        rsItemMst.GetResult("Select Item_Main_grp from Item_Mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  Item_code = '" & strBomItem & "'")
                        ''''            If (UCase(rsItemMst.GetValue("Item_Main_grp")) = "R") Or (UCase(rsItemMst.GetValue("Item_Main_grp")) = "C") Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object rsItemMst.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'Changed for Issue ID eMpro-20081023-22907 Starts
                        'If (UCase(rsItemMst.GetValue("Item_Main_grp")) <> "F") Or (UCase(rsItemMst.GetValue("Item_Main_grp")) <> "S") Then
                        '    ExploreBom = True
                        'Else
                        '    'UPGRADE_WARNING: Couldn't resolve default property of object pstrFinishedQty. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        '    pstrFinishedQty = pstrFinishedQty * dblTotalReqQty
                        '    Call ExploreBom(strBomItem, pstrFinishedQty, pstrSPCurrentRow, pstrCustCode, pstrRef, pintAnnexMaxCount, pstrFinishedProduct)
                        'End If
                        If (UCase(rsItemMst.GetValue("Item_Main_grp")) = "F") Or (UCase(rsItemMst.GetValue("Item_Main_grp")) = "S") Then
                            pstrFinishedQty = pstrFinishedQty * dblTotalReqQty
                            Call ExploreBom(strBomItem, pstrFinishedQty, pstrSPCurrentRow, pstrCustCode, pstrRef, pintAnnexMaxCount, pstrFinishedProduct)
                        Else
                            ExploreBom = True
                        End If
                        'Changed for Issue ID eMpro-20081023-22907 Ends
                    End If
                End If
                rsBomMstRaw.MoveNext()
            Next
        Else
            MsgBox("No BOM Defind for Item (" & strBomItem & ") defined in challan", MsgBoxStyle.Information, "eMPro")
            ExploreBom = False
            Exit Function
        End If
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        'if Child Item not found in CustAnnex_dtl
    End Function
    ''''''''Public Function ExploreBom(pstrItemCode As String, pstrFinishedQty, pstrSPCurrentRow, pstrCustCode As String, pstrRef As String, pintAnnexMaxCount As Integer, pstrFinishedProduct As String) As Boolean
    '''''''''*****************************************************
    '''''''''Created By     -  Nisha
    '''''''''Description    -  to get the values of required items in Sub assambly bom
    '''''''''input Variable -  Item Code to Found, reqquantity of Finished Product,row in spread
    '''''''''*****************************************************
    ''''''''Dim strBomMstRaw As String
    ''''''''Dim rsBomMstRaw As ClsResultSetDB_Invoice
    ''''''''Dim rsCustAnnexDtl As ClsResultSetDB_Invoice
    ''''''''Dim rsVandorBom As ClsResultSetDB_Invoice
    ''''''''Dim rsItemMst As ClsResultSetDB_Invoice
    ''''''''Dim intBomMaxRaw As Integer
    ''''''''Dim intCurrentRaw As Integer
    ''''''''Dim dblTotalReqQty As Double
    '''''''''Dim strProcessType As String
    ''''''''Dim strCustAnnexDtl As String
    ''''''''Dim strRef As String
    ''''''''On Error GoTo ErrHandler
    ''''''''    Set rsBomMstRaw = New ClsResultSetDB_Invoice
    ''''''''     Set rsCustAnnexDtl = New ClsResultSetDB_Invoice
    ''''''''     Set rsItemMst = New ClsResultSetDB_Invoice
    ''''''''    strBomMstRaw = "Select RawMaterial_Code,Required_qty + Waste_qty "
    ''''''''    strBomMstRaw = strBomMstRaw & " As TotalReqQty,Process_Type from Bom_Mst where "
    ''''''''    strBomMstRaw = strBomMstRaw & " item_Code ='" & strBomItem
    ''''''''    strBomMstRaw = strBomMstRaw & "'and finished_product_code ='"
    ''''''''    strBomMstRaw = strBomMstRaw & pstrItemCode & "'"
    ''''''''    rsBomMstRaw.GetResult strBomMstRaw, adOpenKeyset, adLockReadOnly
    ''''''''If rsBomMstRaw.GetNoRows > 0 Then ' If Item Found in Bom Mst
    ''''''''    intBomMaxRaw = rsBomMstRaw.GetNoRows
    ''''''''    rsBomMstRaw.MoveFirst
    ''''''''    For intCurrentRaw = 1 To intBomMaxRaw
    ''''''''        strBomItem = rsBomMstRaw.GetValue("RawMaterial_code")
    ''''''''        dblTotalReqQty = rsBomMstRaw.GetValue("TotalReqQty")
    ''''''''        'strProcessType = rsBomMstRaw.GetValue("Process_Type")
    ''''''''        'String for CustAnnex_dtl
    ''''''''        strCustAnnexDtl = "Select Item_Code,Balance_qty,REF57F4_DATE from CustAnnex_hdr where Customer_code ='"
    ''''''''        strCustAnnexDtl = strCustAnnexDtl & Trim(pstrCustCode) & "'"
    ''''''''        If blnFIFOFlag = False Then
    ''''''''            strRef = Replace(pstrRef, "§", "','", 1)
    ''''''''            strRef = "'" & strRef & "'"
    ''''''''            strCustAnnexDtl = strCustAnnexDtl & " and ref57f4_no IN ("
    ''''''''            strCustAnnexDtl = strCustAnnexDtl & Trim(strRef) & ")"
    ''''''''        End If
    ''''''''        strCustAnnexDtl = strCustAnnexDtl & "  and getdate() <= "
    ''''''''        strCustAnnexDtl = strCustAnnexDtl & " DateAdd(d, 180, ref57f4_date)"
    ''''''''        strCustAnnexDtl = strCustAnnexDtl & " and Item_code ='" & strBomItem & "'"
    ''''''''        rsCustAnnexDtl.GetResult strCustAnnexDtl, adOpenKeyset, adLockReadOnly
    ''''''''        If rsCustAnnexDtl.GetNoRows >= 1 Then 'if item Found in CustAnnex then replace that item from Parant string
    ''''''''        Set rsVandorBom = New ClsResultSetDB_Invoice
    ''''''''            rsVandorBom.GetResult "Select RawMaterial_Code from Vendor_bom where Finish_Product_code = '" & pstrFinishedProduct & "'and RawMaterial_code = '" & strBomItem & "'"
    ''''''''            If rsVandorBom.GetNoRows > 0 Then
    ''''''''                Dim intArrCount As Integer
    ''''''''                Dim blnArrItemFound As Boolean
    ''''''''                rsCustAnnexDtl.MoveFirst
    ''''''''                inti = inti + 1
    ''''''''                ReDim Preserve arrItem(inti)
    ''''''''                ReDim Preserve arrQty(inti)
    ''''''''                ReDim Preserve arrReqQty(inti)
    ''''''''                blnArrItemFound = False
    ''''''''                For intArrCount = 0 To UBound(arrItem) - 1 'to check if ITem Already there in ArrItem Array
    ''''''''                    If UCase(Trim(arrItem(intArrCount))) = UCase(Trim(rsCustAnnexDtl.GetValue("Item_code"))) Then
    ''''''''                    ' if found then sum up Requird Quantity in array arrReqQty and assign value true to blnArrITemFound
    ''''''''                        blnArrItemFound = True
    ''''''''                        arrReqQty(intArrCount) = arrReqQty(intArrCount) + (dblTotalReqQty * pstrFinishedQty)
    ''''''''                        If arrQty(intArrCount) < arrReqQty(intArrCount) Then ' to Check with Quantity supplieded in Cust Annex
    ''''''''                            MsgBox "Customer Supplied Materail for Item " & arrItem(inti) & " is " & arrQty(inti) & " .", vbOKOnly, "eMPro"
    ''''''''                            Cmdinvoice.SetFocus
    ''''''''                            ExploreBom = False
    ''''''''                            Exit Function
    ''''''''                        Else
    ''''''''                            Call ToGetIteminAcustannex(arrCustAnnex, strBomItem, pintAnnexMaxCount, (dblTotalReqQty * pstrFinishedQty))
    ''''''''                            ExploreBom = True
    ''''''''                            Exit For
    ''''''''                        End If
    ''''''''                        blnArrItemFound = False
    ''''''''                    End If
    ''''''''                Next
    ''''''''                If blnArrItemFound = False Then ' if item not found
    ''''''''                    inti = inti + 1
    ''''''''                    ReDim Preserve arrItem(inti)
    ''''''''                    ReDim Preserve arrQty(inti)
    ''''''''                    ReDim Preserve arrReqQty(inti)
    ''''''''                    arrItem(inti) = rsCustAnnexDtl.GetValue("Item_code")
    ''''''''                    arrQty(inti) = rsCustAnnexDtl.GetValue("Balance_qty")
    ''''''''                    'jul
    ''''''''    '                str57f4Date = rsCustAnnexDtl.GetValue("REF57F4_DATE")
    ''''''''                    arrReqQty(inti) = dblTotalReqQty * pstrFinishedQty
    ''''''''                    If arrQty(inti) < arrReqQty(inti) Then
    ''''''''                        MsgBox "Customer Supplied Materail for Item " & arrItem(inti) & " is " & arrQty(inti) & " .", vbOKOnly, "eMPro"
    ''''''''                        Cmdinvoice.SetFocus
    ''''''''                        ExploreBom = False
    ''''''''                        Exit Function
    ''''''''                    Else
    ''''''''                    '***********for Adding Values in CustAnnex Array
    ''''''''                        If Len(Trim(arrCustAnnex(0, pintAnnexMaxCount))) > 0 Then
    ''''''''                            pintAnnexMaxCount = pintAnnexMaxCount + 1
    ''''''''                            ReDim Preserve arrCustAnnex(3, pintAnnexMaxCount)
    ''''''''                        End If
    ''''''''                        arrCustAnnex(0, pintAnnexMaxCount) = rsCustAnnexDtl.GetValue("Item_code")
    ''''''''                        arrCustAnnex(1, pintAnnexMaxCount) = (dblTotalReqQty * pstrFinishedQty)
    ''''''''                        arrCustAnnex(2, pintAnnexMaxCount) = (rsCustAnnexDtl.GetValue("Balance_qty") - (dblTotalReqQty * pstrFinishedQty))
    ''''''''                    '*********************
    ''''''''                        ExploreBom = True
    ''''''''                    End If
    ''''''''                End If
    ''''''''            End If
    ''''''''        Else
    '''''''''            If strProcessType = "I" Then
    ''''''''
    ''''''''            rsVandorBom.GetResult "Select RawMaterial_Code from Vendor_bom where Finish_Product_code = '" & pstrItemCode & "'and RawMaterial_code = '" & strBomItem & "'"
    ''''''''            If rsVandorBom.GetNoRows > 0 Then
    ''''''''                MsgBox "Item " & strBomItem & " is not supplied.", vbOKOnly, "eMPro"
    ''''''''                Cmdinvoice.SetFocus
    ''''''''                ExploreBom = False
    ''''''''                Exit Function
    ''''''''            Else 'if not of Process type I then again Explore
    ''''''''            '####
    ''''''''                rsItemMst.GetResult "Select Item_Main_grp from Item_Mst Where Item_code = '" & strBomItem & "'"
    ''''''''                If (UCase(rsItemMst.GetValue("Item_Main_grp")) = "R") Or (UCase(rsItemMst.GetValue("Item_Main_grp")) = "C") Then
    ''''''''                    ExploreBom = True
    ''''''''                Else
    ''''''''                    pstrFinishedQty = pstrFinishedQty * dblTotalReqQty
    ''''''''                    Call ExploreBom(strBomItem, pstrFinishedQty, pstrSPCurrentRow, pstrCustCode, pstrRef, pintAnnexMaxCount, pstrFinishedProduct)
    ''''''''                End If
    ''''''''            '####
    ''''''''            End If
    ''''''''        End If
    ''''''''        rsBomMstRaw.MoveNext
    ''''''''    Next
    ''''''''Else
    ''''''''    MsgBox "No BOM Defind for Item (" & strBomItem & ") defined in challan", vbInformation, "eMPro"
    ''''''''    ExploreBom = False
    ''''''''    Exit Function
    ''''''''End If
    ''''''''Exit Function
    ''''''''ErrHandler:                             'The Error Handling Code Starts here
    ''''''''    Call gobjError.RAISEERROR_INVOICE(Err.number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    ''''''''        'if Child Item not found in CustAnnex_dtl
    ''''''''End Function
    Public Function ToGetIteminAcustannex(ByRef pvarArray(,) As Object, ByRef pstrItemCode As Object, ByRef pintArrMaxCount As Short, ByRef pdblReqQuantity As Double) As Object
        Dim intLoopCounter As Short
        On Error GoTo ErrHandler
        For intLoopCounter = 0 To pintArrMaxCount
            'UPGRADE_WARNING: Couldn't resolve default property of object pstrItemCode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object pvarArray(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If UCase(Trim(pvarArray(0, intLoopCounter))) = UCase(Trim(pstrItemCode)) Then
                'UPGRADE_WARNING: Couldn't resolve default property of object pvarArray(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object pvarArray(1, intLoopCounter). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                pvarArray(1, intLoopCounter) = pvarArray(1, intLoopCounter) + pdblReqQuantity
                'UPGRADE_WARNING: Couldn't resolve default property of object pvarArray(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object pvarArray(2, intLoopCounter). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                pvarArray(2, intLoopCounter) = pvarArray(2, intLoopCounter) - pdblReqQuantity
                'UPGRADE_WARNING: Couldn't resolve default property of object ToGetIteminAcustannex. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ToGetIteminAcustannex = True
            Else
                'UPGRADE_WARNING: Couldn't resolve default property of object ToGetIteminAcustannex. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ToGetIteminAcustannex = False
            End If
        Next
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    ''''''#######
    '''''''Public Function ToGetIteminAcustannex(pvarArray() As Variant, pstrItemCode, pintArrMaxCount As Integer, pdblReqQuantity As Double)
    '''''''Dim intLoopCounter As Integer
    '''''''On Error GoTo ErrHandler
    '''''''For intLoopCounter = 0 To pintArrMaxCount - 1
    '''''''    If UCase(Trim(pvarArray(0, intLoopCounter))) = UCase(Trim(pstrItemCode)) Then
    '''''''        pvarArray(1, intLoopCounter) = pvarArray(1, intLoopCounter) + pdblReqQuantity
    '''''''        pvarArray(2, intLoopCounter) = pvarArray(2, intLoopCounter) - pdblReqQuantity
    '''''''        ToGetIteminAcustannex = True
    '''''''    Else
    '''''''        ToGetIteminAcustannex = False
    '''''''    End If
    '''''''Next
    ''''''' Exit Function
    '''''''ErrHandler:                             'The Error Handling Code Starts here
    '''''''    ChangeMousePointer obj_Screen, Me, vbDefault
    '''''''    Call gobjError.RAISEERROR_INVOICE(Err.number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    '''''''End Function
    'UPGRADE_WARNING: Event optInvYes.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub optInvYes_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optInvYes.CheckedChanged
        If eventSender.Checked Then
            Dim Index As Short = optInvYes.GetIndex(eventSender)
            Ctlinvoice.Text = ""
        End If
    End Sub
    Private Sub optInvYes_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles optInvYes.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Dim Index As Short = optInvYes.GetIndex(eventSender)
        On Error GoTo Err_Handler
        If KeyAscii = 13 Then
            Call cmbInvType.Focus()
        End If
        GoTo EventExitSub
Err_Handler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)

EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtPLA_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtPLA.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Select Case KeyAscii
            Case 39, 34, 96
                KeyAscii = 0
            Case 13
                'Cmdinvoice.SetFocus
                '********Added By Tapan on 20-Aug-2K2******
                'chkLockPrintingFlag.Enabled = True
                'chkLockPrintingFlag.SetFocus
                '*********Addition Ends**************
                '********Added By Nisha on 23-Jan-2K3******
                dtpRemoval.Enabled = True
                dtpRemovalTime.Enabled = True
                dtpRemoval.Focus()
                '*********Addition Ends**************
        End Select
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Public Function CheckDataFromGrin(ByRef pdblDocNo As Double, ByRef pstrCustCode As String) As Boolean
        Dim rsGrnDtl As ClsResultSetDB_Invoice
        Dim rsSalesDtl As ClsResultSetDB_Invoice
        Dim strSql As String
        Dim StrItemCode As String
        Dim dblItemQty As Double
        Dim dblRejQty As Double
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        On Error GoTo ErrHandler
        rsGrnDtl = New ClsResultSetDB_Invoice
        rsSalesDtl = New ClsResultSetDB_Invoice
        rsSalesDtl.GetResult("Select Item_Code,Sales_Quantity from Sales_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  doc_No =" & Ctlinvoice.Text & " and Location_Code='" & Trim(txtUnitCode.Text) & "'")
        intMaxLoop = rsSalesDtl.GetNoRows : rsSalesDtl.MoveFirst()
        CheckDataFromGrin = False
        'strSQL = "Select * from Grn_hdr where Doc_No = " & pdblDocNo & " and ref_INvoice_Flag =0"
        'rsGrnDtl.GetResult strSQL
        'If rsGrnDtl.GetNoRows > 0 Then
        '    CheckDataFromGrin = True
        'Else
        '    CheckDataFromGrin = False
        '    MsgBox "There is already One invoice Generated For Grin No :" & pdblDocNo & " ,Cannot Generate More then One.", vbInformation, "eMPro"
        '    Exit Function
        'End If
        For intLoopCounter = 1 To intMaxLoop
            'UPGRADE_WARNING: Couldn't resolve default property of object rsSalesDtl.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            StrItemCode = rsSalesDtl.GetValue("Item_code")
            'UPGRADE_WARNING: Couldn't resolve default property of object rsSalesDtl.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            dblItemQty = rsSalesDtl.GetValue("Sales_quantity")
            strSql = "select a.Doc_No,a.Item_code,a.Rejected_Quantity,"
            strSql = strSql & "Despatch_Quantity = isnull(a.Despatch_Quantity,0),"
            strSql = strSql & " Inspected_Quantity = isnull(Inspected_Quantity,0),"
            strSql = strSql & "RGP_Quantity = isnull(RGP_Quantity,0) from grn_Dtl a,grn_hdr b Where A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" + gstrUNITID + "' and "
            strSql = strSql & "a.Doc_type = b.Doc_type And a.Doc_No = b.Doc_No and "
            strSql = strSql & "a.From_Location = b.From_Location and a.From_Location ='01R1'"
            strSql = strSql & "and a.Rejected_quantity > 0 and b.Vendor_code = '" & pstrCustCode
            strSql = strSql & "' and a.Doc_No = " & pdblDocNo & " and a.Item_code = '" & StrItemCode & "'"
            rsGrnDtl.GetResult(strSql)
            'UPGRADE_WARNING: Couldn't resolve default property of object rsGrnDtl.GetValue(RGP_Quantity). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object rsGrnDtl.GetValue(Inspected_Quantity). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object rsGrnDtl.GetValue(Despatch_Quantity). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object rsGrnDtl.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            dblRejQty = rsGrnDtl.GetValue("Rejected_Quantity") - rsGrnDtl.GetValue("Despatch_Quantity") - rsGrnDtl.GetValue("Inspected_Quantity") - rsGrnDtl.GetValue("RGP_Quantity")
            If rsGrnDtl.GetNoRows > 0 Then
                If dblItemQty > (dblRejQty) Then
                    MsgBox("Max. Quantity Allowed For Item " & StrItemCode & " is " & dblRejQty & ", Quantity Entered in Invoice is : " & dblItemQty)
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
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Public Function UpdateGrnHdr(ByRef pdblGrinNo As Double, ByRef pdblinvoiceNo As Double) As Object
        Dim rsSalesDtl As ClsResultSetDB_Invoice
        Dim intMaxLoop As Short
        Dim StrItemCode As String
        Dim dblqty As Double
        Dim intLoopCount As Short
        rsSalesDtl = New ClsResultSetDB_Invoice
        rsSalesDtl.GetResult("select * from sales_dtl where  UNIT_CODE= '" & gstrUNITID & "' and Doc_No = " & Ctlinvoice.Text & " and Location_Code='" & Trim(txtUnitCode.Text) & "'")
        If rsSalesDtl.GetNoRows > 0 Then
            intMaxLoop = rsSalesDtl.GetNoRows
            rsSalesDtl.MoveFirst()
            strupdateGrinhdr = ""
            For intLoopCount = 1 To intMaxLoop
                'UPGRADE_WARNING: Couldn't resolve default property of object rsSalesDtl.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                StrItemCode = rsSalesDtl.GetValue("ITem_code")
                'UPGRADE_WARNING: Couldn't resolve default property of object rsSalesDtl.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                dblqty = rsSalesDtl.GetValue("Sales_Quantity")
                If Len(Trim(strupdateGrinhdr)) = 0 Then
                    strupdateGrinhdr = "Update Grn_Dtl Set Despatch_Quantity = isnull(Despatch_Quantity,0) +" & dblqty
                    strupdateGrinhdr = strupdateGrinhdr & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  ITem_Code = '" & StrItemCode & "' and Doc_No = " & pdblGrinNo
                Else
                    strupdateGrinhdr = strupdateGrinhdr & vbCrLf & "Update Grn_Dtl Set Despatch_Quantity = isnull(Despatch_Quantity,0) + " & dblqty
                    strupdateGrinhdr = strupdateGrinhdr & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  ITem_Code = '" & StrItemCode & "' and Doc_No = " & pdblGrinNo
                End If
                rsSalesDtl.MoveNext()
            Next
        Else
            MsgBox("No Items Available in Invoice " & Ctlinvoice.Text)
        End If
    End Function
    Private Function GetTaxGlSl(ByVal TaxType As String) As String
        Dim objRecordSet As New ADODB.Recordset
        On Error GoTo ErrHandler
        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close()

        ' Added by priti on 15 Oct 2020 to add rejection tax type APRC 
        If UCase(Trim(cmbInvType.Text)) = "REJECTION" And DataExist("SELECT  TOP 1 1 FROM MKT_INVREJ_DTL WHERE  invoice_no=" & Ctlinvoice.Text & "  and rej_Type=2 and unit_code='" & gstrUNITID & "'") = True Then
            Dim strTaxType = SqlConnectionclass.ExecuteScalar("select RejectionTaxType from sales_parameter where UNIT_CODE = '" & gstrUNITID & "'")
            objRecordSet.Open("SELECT tx_glCode, tx_slCode FROM fin_TaxGlRel WHERE UNIT_CODE='" + gstrUNITID + "' AND  tx_rowType = '" & strTaxType & "' AND tx_taxId ='" & TaxType & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
        Else
            objRecordSet.Open("SELECT tx_glCode, tx_slCode FROM fin_TaxGlRel WHERE UNIT_CODE='" + gstrUNITID + "' AND  tx_rowType = 'ARTAX' AND tx_taxId ='" & TaxType & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
        End If
        'objRecordSet.Open("SELECT tx_glCode, tx_slCode FROM fin_TaxGlRel WHERE UNIT_CODE='" + gstrUNITID + "' AND  tx_rowType = 'ARTAX' AND tx_taxId ='" & TaxType & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)

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
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        GetTaxGlSl = "N"
        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
            objRecordSet.Close()

            objRecordSet = Nothing
        End If
    End Function
    Private Function CreateStringForAccounts() As Boolean
        '-----------------------------------------------------------------------------------
        'Revision  By       : Ashutosh Verma,issue id:15591
        'Revision On        : 19-09-2005
        'History            : Save Loading Charges parameter value to Saleschallan_dtl w.r.t value in Sales_Parameter.SameUnitLoading field.
        '----------------------------------------------------------------------------------
        'Revision  By       : Ashutosh Verma,issue id:17610
        'Revision On        : 02-06-2006
        'History            : Posting for Service tax in accounts.
        '=======================================================================================
        'Revised By      : Manoj Kr. Vaish
        'Issue ID        : 19992
        'Revision Date   : 27 June 2007
        'History         : Group same Item_Code for diiferent SO.in Multiple SO Export Invoice for posting
        '                : Fetch credit term from sales_dtl for saving in ar_docmaster
        '***************************************************************************************
        'Revised By      : Manoj Kr. Vaish
        'Issue ID        : eMpro-20080430-18033
        'Revision Date   : 27 Apr 2008
        'History         : Posting under New tax head for the calculation of CVD Excise,Ecess & SEcess
        '***********************************************************************************
        'Revised By      : Manoj Kr.Vaish
        'Issue ID        : eMpro-20080508-18500
        'Revision Date   : 09 May 2008
        'History         : Posting of SAD Tax for Transfer Invoice in Mate Noida
        '***********************************************************************************
        'Revised By      : Manoj Kr.Vaish
        'Issue ID        : eMpro(-20090601 - 31918)
        'Revision Date   : 01 Jun 2008
        'History         : Posting of Additional VAT Tax
        '***************************************************************************
        'Revised By      : Manoj Vaish
        'Revision On     : 01 Jun 2009
        'Issue ID        : eMpro-20090610-32326
        'History         : Posting of new additional CST tax
        '***********************************************************************************


        Dim objRecordSet As New ADODB.Recordset
        Dim objTmpRecordset As New ADODB.Recordset
        Dim strRetVal As String
        Dim strInvoiceNo As String
        Dim strInvoiceDate As String
        Dim strCurrencyCode As String
        Dim dblInvoiceAmt As Double
        Dim dblInvoiceAmtRoundOff_diff As Double
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
        Dim rsFULLExciseAmount As ClsResultSetDB_Invoice
        Dim dblFullExciseAmount As Double
        Dim blnMsgBox As Boolean
        Dim RejectionRoundoff As Double
        Dim RejectionTotalAmount As Double
        rsFULLExciseAmount = New ClsResultSetDB_Invoice

        '''***** Added by Ashutosh on 20-09-2005.
        Dim dblTotalLoadingcharges As Double
        Dim rsSalesDtl As ClsResultSetDB_Invoice
        Dim intNumberOfItems As Short
        Dim dblLoadingChargePerItem As Double
        Dim dblTempLoadChargesPerItem As Double
        Dim i As Short
        ''cc code DECLARATION changes by prashant rajpal
        Dim strTaxCCCode As String
        Dim blnrejinv_fullvalue As Boolean
        'cc code changes end 

        '''***** End here

        mstrExcisePriorityUpdationString = ""
        RejectionRoundoff = 0
        RejectionTotalAmount = 0
        blnMsgBox = False


        On Error GoTo ErrHandler


        objRecordSet.Open("SELECT * FROM  saleschallan_dtl WHERE  UNIT_CODE= '" & gstrUNITID & "' and Doc_No='" & Trim(Ctlinvoice.Text) & "' and Location_Code='" & Trim(txtUnitCode.Text) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
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
        'strInvoiceDate = CStr(CDate(VB6.Format(objRecordSet.Fields("Invoice_Date").Value, "dd/mm/yyyy")))
        If DataExist("SELECT TOP 1 1 FROM SALES_PARAMETER WHERE INVOICE_LOCKING_ENTRY_SAMEDATE=1  and UNIT_CODE = '" & gstrUNITID & "'") Then
            strInvoiceDate = VB6.Format(GetServerDateTime(), "dd-MMM-yyyy")
        Else
            strInvoiceDate = VB6.Format(objRecordSet.Fields("Invoice_Date").Value, "dd-MMM-yyyy")
        End If

        If DataExist("SELECT TOP 1 1 FROM SALES_PARAMETER WHERE REJINV_POSTING_WITH_FULLVALUE=1  and UNIT_CODE = '" & gstrUNITID & "'") Then
            blnrejinv_fullvalue = True
        Else
            blnrejinv_fullvalue = False
        End If

        strCurrencyCode = Trim(IIf(IsDBNull(objRecordSet.Fields("Currency_Code").Value), "", objRecordSet.Fields("Currency_Code").Value))


        dblInvoiceAmt = IIf(IsDBNull(objRecordSet.Fields("total_amount").Value), 0, objRecordSet.Fields("total_amount").Value)

        dblInvoiceAmtRoundOff_diff = IIf(IsDBNull(objRecordSet.Fields("TotalInvoiceAmtRoundOff_diff").Value), 0, objRecordSet.Fields("TotalInvoiceAmtRoundOff_diff").Value)

        dblExchangeRate = IIf(IsDBNull(objRecordSet.Fields("Exchange_Rate").Value), 1, objRecordSet.Fields("Exchange_Rate").Value)

        dblTCStaxAmt = IIf(IsDBNull(objRecordSet.Fields("TCSTaxAmount").Value), 1, objRecordSet.Fields("TCSTaxAmount").Value)
        strCustCode = Trim(objRecordSet.Fields("Account_Code").Value)

        strCustRef = Trim(IIf(IsDBNull(objRecordSet.Fields("cust_ref").Value), "", objRecordSet.Fields("cust_ref").Value))
        blnExciseExumpted = objRecordSet.Fields("ExciseExumpted").Value
        '''***** Added by Ashutosh on 20-09-2005

        dblTotalLoadingcharges = IIf(IsDBNull(objRecordSet.Fields("loadingChargeTaxAmount").Value), 0, objRecordSet.Fields("loadingChargeTaxAmount").Value)
        '''***** End here
        'Added for Issue ID 19992 Starts

        strCreditTermsID = Trim(IIf(IsDBNull(objRecordSet.Fields("payment_terms").Value), "", objRecordSet.Fields("payment_terms").Value))
        mstrCreditTermId = strCreditTermsID
        'Added for Issue ID 19992 End

        ''Added by priti for blank currency issue in transfer invoice issue on 11 Mar 2025
        If Len(strCurrencyCode) = 0 And UCase(Trim(cmbInvType.Text)) = "TRANSFER INVOICE" Then
            strCurrencyCode = gstrCURRENCYCODE
            dblExchangeRate = CStr(1.0#)
            SqlConnectionclass.ExecuteNonQuery("Update SalesChallan_Dtl set currency_code='" & strCurrencyCode & "' , exchange_rate=" & dblExchangeRate & " where Doc_no=" & Ctlinvoice.Text & " And unit_code ='" & gstrUNITID & "'")

        End If
        ''End by priti for blank currency issue in transfer invoice issue  on 11 Mar 2025


        If UCase(lbldescription.Text) <> "SMP" Then 'if invoice type is not sample sales then
            'Retreiving the customer gl, sl and credit term id
            If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
            If UCase(Trim(cmbInvType.Text)) = "REJECTION" Then
                objTmpRecordset.Open("SELECT ISNULL(SUM(Basic_Amount),0) AS Basic_Amt FROM sales_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  doc_no =" & Ctlinvoice.Text, mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
                If Not objTmpRecordset.EOF Then
                    dblBasicAmount = Val(objTmpRecordset.Fields("Basic_Amt").Value)
                End If
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
                'Change done by Amit Bhatnagar on 05/05/2003
                If (UCase(Trim(cmbInvType.Text)) = "REJECTION" And strCustRef <> "") And blnrejinv_fullvalue = False Then 'In case of non line rejections Basic posting is not done
                    dblInvoiceAmt = dblInvoiceAmt - dblBasicAmount
                    dblBasicAmount = 0
                End If

                objTmpRecordset.Open("SELECT GL_AccountID, Ven_slCode, CrTrm_Termid FROM Pur_VendorMaster WHERE UNIT_CODE='" + gstrUNITID + "' AND  Prty_PartyID='" & strCustCode & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
            Else
                objTmpRecordset.Open("SELECT Cst_ArCode, Cst_slCode, Cst_CreditTerm FROM Sal_CustomerMaster WHERE UNIT_CODE='" + gstrUNITID + "' AND  Prty_PartyID='" & strCustCode & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
            End If
            If objTmpRecordset.EOF Then
                If UCase(Trim(cmbInvType.Text)) = "REJECTION" Then
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
            If UCase(Trim(cmbInvType.Text)) = "REJECTION" Then

                strCustomerGL = Trim(IIf(IsDBNull(objTmpRecordset.Fields("GL_AccountID").Value), "", objTmpRecordset.Fields("GL_AccountID").Value))

                strCustomerSL = Trim(IIf(IsDBNull(objTmpRecordset.Fields("Ven_slCode").Value), "", objTmpRecordset.Fields("Ven_slCode").Value))

                strCreditTermsID = Trim(IIf(IsDBNull(objTmpRecordset.Fields("CrTrm_Termid").Value), "", objTmpRecordset.Fields("CrTrm_Termid").Value))
            Else

                strCustomerGL = Trim(IIf(IsDBNull(objTmpRecordset.Fields("Cst_ArCode").Value), "", objTmpRecordset.Fields("Cst_ArCode").Value))

                strCustomerSL = Trim(IIf(IsDBNull(objTmpRecordset.Fields("Cst_slCode").Value), "", objTmpRecordset.Fields("Cst_slCode").Value))
                'Changed for Issue ID 19992 Starts
                If strCreditTermsID = "" Then

                    strCreditTermsID = Trim(IIf(IsDBNull(objTmpRecordset.Fields("Cst_CreditTerm").Value), "", objTmpRecordset.Fields("Cst_CreditTerm").Value))
                    mstrCreditTermId = strCreditTermsID
                End If
                'Changed for Issue ID 19992 Ends
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

            Dim objCreditTerms As prj_CreditTerm.clsCR_Term_Resolver
            objCreditTerms = New prj_CreditTerm.clsCR_Term_Resolver
            strRetVal = objCreditTerms.RetCR_Term_Dates("", "INV", strCreditTermsID, strInvoiceDate, gstrUNITID, "", "", gstrCONNECTIONSTRING)
            objCreditTerms = Nothing

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
        'code Added by Arshad to round off Total invoice amount according to parameter
        Dim rsSalesParameter As New ADODB.Recordset
        Dim blnTotalInvoiceAmountRoundOff As Boolean
        Dim intTotalInvoiceAmountRoundOff As Short
        If rsSalesParameter.State = ADODB.ObjectStateEnum.adStateOpen Then rsSalesParameter.Close()
        rsSalesParameter.Open("SELECT TotalInvoiceAmount_RoundOff, TotalInvoiceAmountRoundOff_Decimal FROM SALES_PARAMETER WHERE UNIT_CODE='" + gstrUNITID + "'", mP_Connection)
        If Not rsSalesParameter.EOF Then
            blnTotalInvoiceAmountRoundOff = rsSalesParameter.Fields("TotalInvoiceAmount_RoundOff").Value
            intTotalInvoiceAmountRoundOff = rsSalesParameter.Fields("TotalInvoiceAmountRoundOff_Decimal").Value
        End If
        If rsSalesParameter.State = ADODB.ObjectStateEnum.adStateOpen Then
            rsSalesParameter.Close()

            rsSalesParameter = Nothing
        End If
        'Ends here

        If UCase(Trim(cmbInvType.Text)) <> "REJECTION" Then
            mstrMasterString = "I»" & strInvoiceNo & "»Dr»»" & strInvoiceDate & "»»»»»SAL»I»" & strInvoiceNo & "»" & strInvoiceDate & "»"
            If UCase(lbldescription.Text) <> "SMP" Then
                mstrMasterString = mstrMasterString & Trim(strCustCode) & "»" & gstrUNITID & "»" & strCurrencyCode & "»"
            Else
                mstrMasterString = mstrMasterString & "»" & gstrUNITID & "»" & strCurrencyCode & "»"
            End If
            '        mstrMasterString = mstrMasterString & Round(dblInvoiceAmt, 0) & "»" & Round(dblInvoiceAmt * dblExchangeRate, 0) & "»" & _
            ''                            dblExchangeRate & "»" & strCreditTermsID & "»" & strBasicDueDate & "»" & _
            ''                            strPaymentDueDate & "»" & strExpectedDueDate & "»" & strCustomerGL & "»" & _
            ''                            strCustomerSL & "»" & mP_User & "»getdate()»»"
            'IF Condition Added by Arshad to round off Total invoice amount according to parameter
            If blnTotalInvoiceAmountRoundOff Then
                mstrMasterString = mstrMasterString & System.Math.Round(dblInvoiceAmt, 0) & "»" & System.Math.Round(dblInvoiceAmt * dblExchangeRate, 0) & "»" & dblExchangeRate & "»" & strCreditTermsID & "»" & strBasicDueDate & "»" & strPaymentDueDate & "»" & strExpectedDueDate & "»" & strCustomerGL & "»" & strCustomerSL & "»" & mP_User & "»getdate()»»"

            Else
                mstrMasterString = mstrMasterString & System.Math.Round(dblInvoiceAmt, intTotalInvoiceAmountRoundOff) & "»" & System.Math.Round(dblInvoiceAmt * dblExchangeRate, intTotalInvoiceAmountRoundOff) & "»" & dblExchangeRate & "»" & strCreditTermsID & "»" & strBasicDueDate & "»" & strPaymentDueDate & "»" & strExpectedDueDate & "»" & strCustomerGL & "»" & strCustomerSL & "»" & mP_User & "»getdate()»»"

            End If

            'IF Condition Ends here
        Else
            '        mstrMasterString = "M»»" & Format(Date, "dd/mm/yyyy") & "»0»»" & gstrUNITID & "»" & Trim(strCustCode) & "»" & _
            ''                            strInvoiceNo & "»" & strInvoiceDate & "»" & strBasicDueDate & "»" & strPaymentDueDate & "»" & strExpectedDueDate & "»" & strCurrencyCode & "»" & _
            ''                            dblExchangeRate & "»" & Round(dblInvoiceAmt) & "»0»»»Rej. Inv. " & strInvoiceNo & "»" & strCustomerGL & "»" & strCustomerSL & "»DR»" & strCustomerGL & "»" & strCustomerSL & _
            ''                            "»»" & gstrCURRENCYCODE & "»" & mP_User & "»getdate()»0»AP»»»»0»»¦"
            'IF Condition Added by Arshad to round off Total invoice amount according to parameter
            If blnTotalInvoiceAmountRoundOff Then
                mstrMasterString = "M»»" & VB6.Format(GetServerDate(), "dd-MMM-yyyy") & "»0»»" & gstrUNITID & "»" & Trim(strCustCode) & "»" & strInvoiceNo & "»" & strInvoiceDate & "»" & strBasicDueDate & "»" & strPaymentDueDate & "»" & strExpectedDueDate & "»" & strCurrencyCode & "»" & dblExchangeRate & "»" & System.Math.Round(dblInvoiceAmt) & "»0»»»Rej. Inv. " & strInvoiceNo & "»" & strCustomerGL & "»" & strCustomerSL & "»DR»" & strCustomerGL & "»" & strCustomerSL & "»»" & gstrCURRENCYCODE & "»" & mP_User & "»getdate()»0»AP»»»»0»»¦"
                RejectionTotalAmount = System.Math.Round(dblInvoiceAmt)
            Else
                mstrMasterString = "M»»" & VB6.Format(GetServerDate(), "dd-MMM-yyyy") & "»0»»" & gstrUNITID & "»" & Trim(strCustCode) & "»" & strInvoiceNo & "»" & strInvoiceDate & "»" & strBasicDueDate & "»" & strPaymentDueDate & "»" & strExpectedDueDate & "»" & strCurrencyCode & "»" & dblExchangeRate & "»" & System.Math.Round(dblInvoiceAmt, intTotalInvoiceAmountRoundOff) & "»0»»»Rej. Inv. " & strInvoiceNo & "»" & strCustomerGL & "»" & strCustomerSL & "»DR»" & strCustomerGL & "»" & strCustomerSL & "»»" & gstrCURRENCYCODE & "»" & mP_User & "»getdate()»0»AP»»»»0»»¦"
                RejectionTotalAmount = System.Math.Round(dblInvoiceAmt, intTotalInvoiceAmountRoundOff)
            End If
            'IF Condition Ends here
        End If

        iCtr = 1

        'CST/LST/SRT Posting
        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()

        If Trim(IIf(IsDBNull(objRecordSet.Fields("SalesTax_Type").Value), "", objRecordSet.Fields("SalesTax_Type").Value)) <> "" Then

            objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("SalesTax_Type").Value), "", objRecordSet.Fields("SalesTax_Type").Value)) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
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
                    If UCase(Trim(cmbInvType.Text)) <> "REJECTION" Then
                        mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»" & strTaxType & "»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                    Else
                        mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & dblTaxAmt & "»»CST/LST/VAT for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                        RejectionRoundoff = RejectionRoundoff + dblTaxAmt
                    End If
                    iCtr = iCtr + 1
                End If
            End If

        End If

        'State Development TAX
        'Code Added by Nisha on 03 May 2005 for SDT Tax
        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()

        If Trim(IIf(IsDBNull(objRecordSet.Fields("SDTax_Type").Value), "", objRecordSet.Fields("SDTax_Type").Value)) <> "" Then

            objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("SalesTax_Type").Value), "", objRecordSet.Fields("SDTax_Type").Value)) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
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


            If strTaxType = "SDT" Then

                dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("SDTax_Amount").Value), 0, objRecordSet.Fields("SDTax_Amount").Value)
                dblBaseCurrencyAmount = dblTaxAmt

                dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("SDTax_Per").Value), 0, objRecordSet.Fields("SDTax_Per").Value)
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
                    If UCase(Trim(cmbInvType.Text)) <> "REJECTION" Then
                        mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»" & strTaxType & "»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                    Else
                        mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & dblTaxAmt & "»»SDT for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                        RejectionRoundoff = RejectionRoundoff + dblTaxAmt
                    End If
                    iCtr = iCtr + 1
                End If
            End If
            'Code Ends Here by nisha On 03 May 2005
        End If


        'Added for Issue ID eMpro(-20090601 - 31918) Starts----Additional VAT Tax
        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()

        If Trim(IIf(IsDBNull(objRecordSet.Fields("ADDVAT_Type").Value), "", objRecordSet.Fields("ADDVAT_Type").Value)) <> "" Then

            objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("ADDVAT_Type").Value), "", objRecordSet.Fields("ADDVAT_Type").Value)) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
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
            If strTaxType = "ADVAT" Or strTaxType = "ADCST" Then

                dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("ADDVAT_Amount").Value), 0, objRecordSet.Fields("ADDVAT_Amount").Value)
                dblBaseCurrencyAmount = dblTaxAmt

                dblTaxRate = Find_Value("select ADDVAT_Per from saleschallan_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  doc_no='" & Ctlinvoice.Text & "'")
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
                    If UCase(Trim(cmbInvType.Text)) <> "REJECTION" Then
                        mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»" & strTaxType & "»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                    Else
                        mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & dblTaxAmt & "»»ADDVAT for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                        RejectionRoundoff = RejectionRoundoff + dblTaxAmt
                    End If
                    iCtr = iCtr + 1
                End If
            End If

        End If
        'Added for Issue ID eMpro(-20090601 - 31918) Ends

        'Changes Done By nisha on 08/07/2004 forECESS Details
        'ECS Posting
        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()

        If Trim(IIf(IsDBNull(objRecordSet.Fields("ECESS_Type").Value), "", objRecordSet.Fields("ECESS_Type").Value)) <> "" Then

            objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("ECESS_Type").Value), "", objRecordSet.Fields("ECESS_Type").Value)) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
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
                    If UCase(Trim(cmbInvType.Text)) <> "REJECTION" Then
                        mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»" & strTaxType & "»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                    Else
                        mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & dblTaxAmt & "»»ECS for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                        RejectionRoundoff = RejectionRoundoff + dblTaxAmt
                    End If
                    iCtr = iCtr + 1
                End If
            End If
        End If

        ''---- S.ECS Posting (Added by Davinder for ESCESS Posting)
        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()

        If Trim(IIf(IsDBNull(objRecordSet.Fields("SECESS_Type").Value), "", objRecordSet.Fields("SECESS_Type").Value)) <> "" Then

            objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("SECESS_Type").Value), "", objRecordSet.Fields("SECESS_Type").Value)) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
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
                    If UCase(Trim(cmbInvType.Text)) <> "REJECTION" Then
                        mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»" & strTaxType & "»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                    Else
                        mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & dblTaxAmt & "»»ECSSH for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                        RejectionRoundoff = RejectionRoundoff + dblTaxAmt
                    End If
                    iCtr = iCtr + 1
                End If
            End If
        End If

        'Changes Ends Here 08/07/2004

        'Added for Issue ID eMpro-20080430-18033 Starts
        If mblnEOUUnit = True Then

            'Posting of ECS on CVD
            If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
            If Trim(IIf(IsDBNull(objRecordSet.Fields("CVDCESS_Type").Value), "", objRecordSet.Fields("CVDCESS_Type").Value)) <> "" Then
                objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("CVDCESS_Type").Value), "", objRecordSet.Fields("CVDCESS_Type").Value)) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
                If Not objTmpRecordset.EOF Then
                    strTaxType = Trim(UCase$(objTmpRecordset.Fields("Tx_TaxeID").Value))
                Else
                    MsgBox("Tax type not found", vbInformation, ResolveResString(100))
                    CreateStringForAccounts = False
                    If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close() : objRecordSet = Nothing
                    If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close() : objTmpRecordset = Nothing
                    Exit Function
                End If
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
                If strTaxType = "ECSCV" Then
                    dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("CVDCESS_Amount").Value), 0, objRecordSet.Fields("CVDCESS_Amount").Value)
                    dblBaseCurrencyAmount = dblTaxAmt
                    dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("CVDCESS_Per").Value), 0, objRecordSet.Fields("CVDCESS_Per").Value)
                    If dblBaseCurrencyAmount > 0 Then
                        'initializing the tax gl and sl here
                        strRetVal = GetTaxGlSl(strTaxType)
                        If strRetVal = "N" Then
                            MsgBox("GL for ARTAX is not defined for " & strTaxType, vbInformation, "eMPro")
                            CreateStringForAccounts = False
                            If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close() : objRecordSet = Nothing
                            Exit Function
                        End If
                        varTmp = Split(strRetVal, "»")
                        strTaxGL = varTmp(0)
                        strTaxSL = varTmp(1)
                        If UCase$(Trim(cmbInvType.Text)) <> "REJECTION" Then
                            mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" &
                                    iCtr & "»TAX»" & strTaxType & "»0»" &
                                    "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                        Else
                            mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" &
                                               dblTaxAmt & "»»ECSCV for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                            RejectionRoundoff = RejectionRoundoff + dblTaxAmt
                        End If
                        iCtr = iCtr + 1
                    End If
                End If
            End If

            ''---- Posting of S.ECS on CVD
            If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
            If Trim(IIf(IsDBNull(objRecordSet.Fields("CVDSECESS_Type").Value), "", objRecordSet.Fields("CVDSECESS_Type").Value)) <> "" Then
                objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("CVDSECESS_Type").Value), "", objRecordSet.Fields("CVDSECESS_Type").Value)) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
                If Not objTmpRecordset.EOF Then
                    strTaxType = Trim(UCase$(objTmpRecordset.Fields("Tx_TaxeID").Value))
                Else
                    MsgBox("Tax type not found", vbInformation, ResolveResString(100))
                    CreateStringForAccounts = False
                    If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close() : objRecordSet = Nothing
                    If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close() : objTmpRecordset = Nothing
                    Exit Function
                End If
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
                If strTaxType = "HCSCV" Then
                    dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("CVDSECESS_Amount").Value), 0, objRecordSet.Fields("CVDSECESS_Amount").Value)
                    dblBaseCurrencyAmount = dblTaxAmt
                    dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("CVDSECESS_Per").Value), 0, objRecordSet.Fields("CVDSECESS_Per").Value)
                    If dblBaseCurrencyAmount > 0 Then
                        'initializing the tax gl and sl here
                        strRetVal = GetTaxGlSl(strTaxType)
                        If strRetVal = "N" Then
                            MsgBox("GL for ARTAX is not defined for " & strTaxType, vbInformation, ResolveResString(100))
                            CreateStringForAccounts = False
                            If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close() : objRecordSet = Nothing
                            Exit Function
                        End If
                        varTmp = Split(strRetVal, "»")
                        strTaxGL = varTmp(0)
                        strTaxSL = varTmp(1)
                        If UCase$(Trim(cmbInvType.Text)) <> "REJECTION" Then
                            mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" &
                                    iCtr & "»TAX»" & strTaxType & "»0»" &
                                    "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                        Else
                            mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" &
                                               dblTaxAmt & "»»HCSCV for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                            RejectionRoundoff = RejectionRoundoff + dblTaxAmt
                        End If
                        iCtr = iCtr + 1
                    End If
                End If
            End If

            'Posting of Additional ECS on Total Duty
            If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
            If Trim(IIf(IsDBNull(objRecordSet.Fields("Ecess_TotalDuty_Type").Value), "", objRecordSet.Fields("Ecess_TotalDuty_Type").Value)) <> "" Then
                objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("Ecess_TotalDuty_Type").Value), "", objRecordSet.Fields("Ecess_TotalDuty_Type").Value)) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
                If Not objTmpRecordset.EOF Then
                    strTaxType = Trim(UCase$(objTmpRecordset.Fields("Tx_TaxeID").Value))
                Else
                    MsgBox("Tax type not found", vbInformation, ResolveResString(100))
                    CreateStringForAccounts = False
                    If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close() : objRecordSet = Nothing
                    If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close() : objTmpRecordset = Nothing
                    Exit Function
                End If
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
                If strTaxType = "ADECS" Then
                    dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("Ecess_TotalDuty_Amount").Value), 0, objRecordSet.Fields("Ecess_TotalDuty_Amount").Value)
                    dblBaseCurrencyAmount = dblTaxAmt
                    dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("Ecess_TotalDuty_Per").Value), 0, objRecordSet.Fields("Ecess_TotalDuty_Per").Value)
                    If dblBaseCurrencyAmount > 0 Then
                        'initializing the tax gl and sl here
                        strRetVal = GetTaxGlSl(strTaxType)
                        If strRetVal = "N" Then
                            MsgBox("GL for ARTAX is not defined for " & strTaxType, vbInformation, ResolveResString(100))
                            CreateStringForAccounts = False
                            If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close() : objRecordSet = Nothing
                            Exit Function
                        End If
                        varTmp = Split(strRetVal, "»")
                        strTaxGL = varTmp(0)
                        strTaxSL = varTmp(1)
                        If UCase$(Trim(cmbInvType.Text)) <> "REJECTION" Then
                            mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" &
                                    iCtr & "»TAX»" & strTaxType & "»0»" &
                                    "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                        Else
                            mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" &
                                               dblTaxAmt & "»»ADECS for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                            RejectionRoundoff = RejectionRoundoff + dblTaxAmt
                        End If
                        iCtr = iCtr + 1
                    End If
                End If
            End If

            ''---- Posting of Additionla S.ECS on Total Duty
            If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
            If Trim(IIf(IsDBNull(objRecordSet.Fields("SEcess_TotalDuty_Type").Value), "", objRecordSet.Fields("SEcess_TotalDuty_Type").Value)) <> "" Then
                objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("SEcess_TotalDuty_Type").Value), "", objRecordSet.Fields("SEcess_TotalDuty_Type").Value)) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
                If Not objTmpRecordset.EOF Then
                    strTaxType = Trim(UCase$(objTmpRecordset.Fields("Tx_TaxeID").Value))
                Else
                    MsgBox("Tax type not found", vbInformation, ResolveResString(100))
                    CreateStringForAccounts = False
                    If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close() : objRecordSet = Nothing
                    If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close() : objTmpRecordset = Nothing
                    Exit Function
                End If
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
                If strTaxType = "ADHCS" Then
                    dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("SEcess_TotalDuty_Amount").Value), 0, objRecordSet.Fields("SEcess_TotalDuty_Amount").Value)
                    dblBaseCurrencyAmount = dblTaxAmt
                    dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("SEcess_TotalDuty_Per").Value), 0, objRecordSet.Fields("SEcess_TotalDuty_Per").Value)
                    If dblBaseCurrencyAmount > 0 Then
                        'initializing the tax gl and sl here
                        strRetVal = GetTaxGlSl(strTaxType)
                        If strRetVal = "N" Then
                            MsgBox("GL for ARTAX is not defined for " & strTaxType, vbInformation, ResolveResString(100))
                            CreateStringForAccounts = False
                            If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close() : objRecordSet = Nothing
                            Exit Function
                        End If
                        varTmp = Split(strRetVal, "»")
                        strTaxGL = varTmp(0)
                        strTaxSL = varTmp(1)
                        If UCase$(Trim(cmbInvType.Text)) <> "REJECTION" Then
                            mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" &
                                    iCtr & "»TAX»" & strTaxType & "»0»" &
                                    "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                        Else
                            mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" &
                                               dblTaxAmt & "»»ADHCS for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                            RejectionRoundoff = RejectionRoundoff + dblTaxAmt
                        End If
                        iCtr = iCtr + 1
                    End If
                End If
            End If
        End If
        'Added for Issue ID eMpro-20080430-18033 Ends

        'Changes Done By Arshad Ali on 01/02/2005 for Turn Over Tax
        'Turn Over Tax Posting
        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()

        If Trim(IIf(IsDBNull(objRecordSet.Fields("TurnOverTaxType").Value), "", objRecordSet.Fields("TurnOverTaxType").Value)) <> "" Then

            objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("TurnOverTaxType").Value), "", objRecordSet.Fields("TurnOverTaxType").Value)) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
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
            If strTaxType = "TOVT" Then

                dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("turnOver_amt").Value), 0, objRecordSet.Fields("turnOver_amt").Value)
                dblBaseCurrencyAmount = dblTaxAmt

                dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("TurnOverTax_per").Value), 0, objRecordSet.Fields("TurnOverTax_per").Value)
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
                    If UCase(Trim(cmbInvType.Text)) <> "REJECTION" Then
                        mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»" & strTaxType & "»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                    Else
                        mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & dblTaxAmt & "»»TOVT for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                        RejectionRoundoff = RejectionRoundoff + dblTaxAmt
                    End If
                    iCtr = iCtr + 1
                End If
            End If
        End If
        'Changes Ends Here 01/02/2005


        '''***** Changes done By Ashutosh on 02-06-2006,Issue Id :17610
        'Service Tax Posting
        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()

        If Trim(IIf(IsDBNull(objRecordSet.Fields("ServiceTax_Type").Value), "", objRecordSet.Fields("ServiceTax_Type").Value)) <> "" Then

            objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("ServiceTax_Type").Value), "", objRecordSet.Fields("ServiceTax_Type").Value)) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
            If Not objTmpRecordset.EOF Then
                strTaxType = Trim(UCase(objTmpRecordset.Fields("Tx_TaxeID").Value))
            Else
                MsgBox("SRT Tax type not found", MsgBoxStyle.Information, "eMPro")
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
            If strTaxType = "SRT" Then

                dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("ServiceTax_Amount").Value), 0, objRecordSet.Fields("ServiceTax_Amount").Value)
                dblBaseCurrencyAmount = dblTaxAmt

                dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("ServiceTax_Per").Value), 0, objRecordSet.Fields("ServiceTax_Per").Value)
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

                    If UCase(Trim(cmbInvType.Text)) <> "REJ" Then
                        mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»" & strTaxType & "»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                    Else
                        mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & dblTaxAmt & "»»SRTAX for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                    End If
                    iCtr = iCtr + 1
                End If
            End If
        End If
        '''***** Changes for issue id:17610 end here.


        'Changes Done By Arshad on 20/09/2004 for ECESS ODetails
        'ECS on Sale Tax Posting
        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()

        If Trim(IIf(IsDBNull(objRecordSet.Fields("SRCESS_Type").Value), "", objRecordSet.Fields("SRCESS_Type").Value)) <> "" Then

            objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("SRCESS_Type").Value), "", objRecordSet.Fields("SRCESS_Type").Value)) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
            If Not objTmpRecordset.EOF Then
                strTaxType = Trim(UCase(objTmpRecordset.Fields("Tx_TaxeID").Value))
            Else
                MsgBox("ECESS on Sale Tax type not found", MsgBoxStyle.Information, "eMPro")
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
            If strTaxType = "ECSR" Then

                dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("SRCESS_Amount").Value), 0, objRecordSet.Fields("SRCESS_Amount").Value)
                dblBaseCurrencyAmount = dblTaxAmt

                dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("SRCESS_Per").Value), 0, objRecordSet.Fields("SRCESS_Per").Value)
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
                    If UCase(Trim(cmbInvType.Text)) <> "REJECTION" Then
                        mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»" & strTaxType & "»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                    Else
                        mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & dblTaxAmt & "»»ECS for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                        RejectionRoundoff = RejectionRoundoff + dblTaxAmt
                    End If
                    iCtr = iCtr + 1
                End If
            End If
        End If
        'Changes Ends Here 20/09/2004


        'Added for Issue ID 20052(Posting of SECESS on Service tax for Job Work Invoice) Start
        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()

        If Trim(IIf(IsDBNull(objRecordSet.Fields("SRSECESS_Type").Value), "", objRecordSet.Fields("SRSECESS_Type").Value)) <> "" Then

            objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("SRSECESS_Type").Value), "", objRecordSet.Fields("SRSECESS_Type").Value)) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
            If Not objTmpRecordset.EOF Then
                strTaxType = Trim(UCase(objTmpRecordset.Fields("Tx_TaxeID").Value))
            Else
                MsgBox("HECSR Tax type not found", MsgBoxStyle.Information, ResolveResString(100))
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
            If strTaxType = "HECSR" Then

                dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("SRSECESS_Amount").Value), 0, objRecordSet.Fields("SRSECESS_Amount").Value)
                dblBaseCurrencyAmount = dblTaxAmt

                dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("SRSECESS_Per").Value), 0, objRecordSet.Fields("SRSECESS_Per").Value)
                If dblBaseCurrencyAmount > 0 Then
                    'initializing the tax gl and sl here
                    strRetVal = GetTaxGlSl(strTaxType)
                    If strRetVal = "N" Then
                        MsgBox("GL for ARTAX is not defined for " & strTaxType, MsgBoxStyle.Information, ResolveResString(100))
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
                    If UCase(Trim(cmbInvType.Text)) <> "REJ" Then
                        mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»" & strTaxType & "»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                    Else
                        mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & dblTaxAmt & "»»SHECS for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                    End If
                    iCtr = iCtr + 1
                End If
            End If

        End If
        'Added for Issue ID 20052(Posting of SECESS on Service tax for Job Work Invoice) End

        'SST Posting

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

            If UCase(Trim(cmbInvType.Text)) <> "REJECTION" Then
                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»SST»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
            Else
                mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & dblTaxAmt & "»»Surcharge for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                RejectionRoundoff = RejectionRoundoff + dblTaxAmt
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
            If UCase(Trim(cmbInvType.Text)) <> "REJECTION" Then
                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»INS»0»" & "»»0»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
            Else
                mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & dblTaxAmt & "»»Insurance for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                RejectionRoundoff = RejectionRoundoff + dblTaxAmt
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
            If UCase(Trim(cmbInvType.Text)) <> "REJECTION" Then
                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»FRT»0»" & "»»0»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
            Else
                mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & dblTaxAmt & "»»Freight for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                RejectionRoundoff = RejectionRoundoff + dblTaxAmt
            End If
            iCtr = iCtr + 1
        End If
        '******************Discount Posting code added by nisha on 18/09/2003

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
            strTaxCCCode = ""
            If (((UCase(Trim(cmbInvType.Text))) = "TRANSFER INVOICE") Or ((UCase(Trim(cmbInvType.Text))) = "REJECTION") Or ((UCase(Trim(cmbInvType.Text))) = "SAMPLE INVOICE")) Then
                If mblnCCFlag = True Then
                    If VerifyGLCCMappingFlag(strTaxGL) Then
                        strTaxCCCode = GetCommonCCCode("Discount_Interest")
                        If strTaxCCCode = "N" Then
                            CreateStringForAccounts = False
                            If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                                objRecordSet.Close()

                                objRecordSet = Nothing
                            End If
                            Exit Function
                        End If
                    End If
                End If
            End If
            If System.Math.Abs(dblBaseCurrencyAmount) > 0 Then
                If UCase(Trim(cmbInvType.Text)) <> "REJECTION" Then
                    'mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»»TAX»0»" & "»»0»" & strTaxGL & "»" & strTaxSL & "»" & System.Math.Abs(dblBaseCurrencyAmount) & "»" & "Dr»»»»»»0»0»0»0»0" & "¦"
                    mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»»TAX»0»" & "»»0»" & strTaxGL & "»" & strTaxSL & "»" & System.Math.Abs(dblBaseCurrencyAmount) & "»" & "Dr»»" & strTaxCCCode & "»»»»0»0»0»0»0" & "¦"
                Else
                    mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»DR»" & System.Math.Abs(dblBaseCurrencyAmount) & "»Discount amount for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                    RejectionRoundoff = RejectionRoundoff + System.Math.Abs(dblBaseCurrencyAmount)
                End If
            End If
            iCtr = iCtr + 1
        End If
        '********************************** changes ends here by nisha on 18/09/2003
        '******************TCS Tax Posting code added by nisha on 26/02/2004
        'If (UCase(Trim(lbldescription.Text)) = "INV") And (UCase(Trim(lblcategory.Text)) = "L") Then

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
        'End If
        '********************************** changes ends here by nisha on 26/02/2004

        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close()
        'Changed for Issue ID 19992 Starts
        If mblnMultipleSOAllowed = False Then
            objRecordSet.Open("SELECT sales_dtl.*, item_mst.GlGrp_code FROM sales_dtl, item_mst WHERE  sales_dtl.UNIT_CODE='" + gstrUNITID + "' AND SALES_DTL.UNIT_CODE=ITEM_MST.UNIT_CODE AND sales_dtl.Doc_No='" & Trim(Ctlinvoice.Text) & "' and sales_dtl.Item_Code=item_mst.Item_Code and sales_dtl.Location_Code='" & Trim(txtUnitCode.Text) & "'")
        Else
            objRecordSet.Open("SELECT isnull(sum(a.basic_amount),0) as Basic_Amount,isnull(sum(a.CustMtrl_Amount),0) as CustMtrl_Amount, isnull(sum(a.Excise_tax),0) as Excise_tax,isnull(sum(ItemPacking_Amount),0)  as ItemPacking_Amount,isnull(sum(others),0) as Others,a.item_code, b.GlGrp_code," & "isnull(sum(packing),0) as packing FROM sales_dtl a, item_mst b WHERE A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" + gstrUNITID + "' AND a.Doc_No='" & Trim(Ctlinvoice.Text) & "' and a.Item_Code=b.Item_Code and a.Location_Code='" & Trim(txtUnitCode.Text) & "'" & " group by a.item_code,b.GlGrp_code")
        End If
        'Changed for Issue ID 19992 Ends

        If objRecordSet.EOF Then
            MsgBox("Item details not found.", MsgBoxStyle.Information, "eMPro")
            CreateStringForAccounts = False
            If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                objRecordSet.Close()

                objRecordSet = Nothing
            End If
            Exit Function
        End If
        '''***** Code added by Ashutosh on 20-09-2005, Issue id:15591, to count number of items in invoice.
        rsSalesDtl = New ClsResultSetDB_Invoice
        'Changed for Issue ID 19992 Starts
        If mblnMultipleSOAllowed = False Then
            rsSalesDtl.GetResult("select doc_no from sales_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  Doc_No='" & Trim(Ctlinvoice.Text) & "' and Location_Code='" & Trim(txtUnitCode.Text) & "'")
        Else
            rsSalesDtl.GetResult("select item_code from sales_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  Doc_No='" & Trim(Ctlinvoice.Text) & "' and Location_Code='" & Trim(txtUnitCode.Text) & "' group by item_code")
        End If
        'Changed for Issue ID 19992 Ends
        intNumberOfItems = rsSalesDtl.GetNoRows
        dblLoadingChargePerItem = System.Math.Round(dblTotalLoadingcharges / intNumberOfItems, 2)
        'UPGRADE_NOTE: Object rsSalesDtl may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        rsSalesDtl = Nothing
        '''***** Changes on 20-09-2005 end here.

        Dim blnFOC As Boolean
        Dim dblPacking_per As Double
        While Not objRecordSet.EOF
            i = i + 1

            strGlGroupId = Trim(IIf(IsDBNull(objRecordSet.Fields("GlGrp_code").Value), "", objRecordSet.Fields("GlGrp_code").Value))

            'Basic Amount Posting

            'Added by arshad
            'blnFOC = IIf(Find_Value("select foc_invoice from salesChallan_dtl where Location_Code='" & Trim(txtLocationCode.Text) & "' and doc_no='" & Trim(txtChallanNo.Text) & "'") = "true", True, False)
            blnFOC = CBool(Find_Value("select foc_invoice from salesChallan_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  Location_Code='" & Trim(txtUnitCode.Text) & "' and doc_no='" & Trim(Ctlinvoice.Text) & "'"))
            If (UCase(Trim(cmbInvType.Text)) = "SAMPLE INVOICE" Or UCase(Trim(cmbInvType.Text)) = "CSM INVOICE") And blnFOC Then
                'skip posting of basic if invoice is FOC Sample invoice
            ElseIf (UCase(Trim(cmbInvType.Text)) = "REJECTION" And (strCustRef = "" Or blnrejinv_fullvalue = True)) Or UCase(Trim(cmbInvType.Text)) <> "REJECTION" Then  'In case of non line rejections Basic posting is not done
                'ends here

                '''***** Changes done by Ashutosh on 20-09-2005, Issue id:15591
                '''dblBasicAmount = IIf(IsNull(objRecordSet.Fields("Basic_Amount").value), 0, objRecordSet.Fields("Basic_Amount").value)
                If i = intNumberOfItems Then

                    dblBasicAmount = IIf(IsDBNull(objRecordSet.Fields("Basic_Amount").Value), 0, objRecordSet.Fields("Basic_Amount").Value) + (dblTotalLoadingcharges - dblTempLoadChargesPerItem)
                Else

                    dblBasicAmount = IIf(IsDBNull(objRecordSet.Fields("Basic_Amount").Value), 0, objRecordSet.Fields("Basic_Amount").Value) + dblLoadingChargePerItem
                    dblTempLoadChargesPerItem = dblTempLoadChargesPerItem + dblLoadingChargePerItem

                End If
                '''***** Changes done on 20-09-2005 end here.



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
                    objTmpRecordset.Open("SELECT * FROM invcc_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_Type='" & lbldescription.Text & "' AND Sub_Type = '" & lblcategory.Text & "' AND Location_Code ='" & Trim(txtUnitCode.Text) & "' AND ccM_cc_Percentage > 0", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
                    If Not objTmpRecordset.EOF Then
                        While Not objTmpRecordset.EOF
                            dblCCShare = (dblBaseCurrencyAmount / 100) * objTmpRecordset.Fields("ccM_cc_Percentage").Value
                            If UCase(Trim(cmbInvType.Text)) <> "REJECTION" Then
                                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»ITM»SAL»" & iCtr & "»" & Trim(objRecordSet.Fields("item_code").Value) & "»" & strGlGroupId & "»0»" & strItemGL & "»" & strItemSL & "»" & dblCCShare & "»Cr»»" & Trim(objTmpRecordset.Fields("ccM_ccCode").Value) & "»»»»0»0»0»0»0¦"
                            Else
                                mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strItemGL & "»" & strItemSL & "»" & Trim(objTmpRecordset.Fields("ccM_ccCode").Value) & "»»»CR»" & dblCCShare & "»»Basic for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                            End If
                            objTmpRecordset.MoveNext()
                            iCtr = iCtr + 1
                        End While
                    Else
                        'ISSUE ID : 1084018
                        strTaxCCCode = ""
                        If (((UCase(Trim(cmbInvType.Text))) = "TRANSFER INVOICE") Or ((UCase(Trim(cmbInvType.Text))) = "REJECTION")) Then
                            If mblnCCFlag = True Then
                                If VerifyGLCCMappingFlag(strItemGL) Then
                                    strTaxCCCode = GetCommonCCCode("Sales")
                                    If strTaxCCCode = "N" Then
                                        CreateStringForAccounts = False
                                        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                                            objRecordSet.Close()

                                            objRecordSet = Nothing
                                        End If
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If
                        If (((UCase(Trim(cmbInvType.Text))) = "SAMPLE INVOICE")) Then
                            If mblnCCFlag = True Then
                                If VerifyGLCCMappingFlag(strItemGL) Then
                                    strTaxCCCode = GetCommonCCCode("Sales")
                                    If strTaxCCCode = "N" Then
                                        CreateStringForAccounts = False
                                        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                                            objRecordSet.Close()

                                            objRecordSet = Nothing
                                        End If
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If
                        'ISSUE ID : 1084018 DONE
                        If (((UCase(Trim(cmbInvType.Text))) = "SERVICE INVOICE")) Then
                            If mblnCCFlag = True Then
                                strTaxCCCode = Find_Value("SELECT serviceinvoice_costcenter FROM SALES_DTL WHERE UNIT_CODE='" + gstrUNITID + "'and doc_no='" & Trim(Ctlinvoice.Text) & "' and  item_code='" & Trim(objRecordSet.Fields("item_code").Value) & "'")
                            End If
                        End If

                        If UCase(Trim(cmbInvType.Text)) <> "REJECTION" Then
                            mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»ITM»SAL»" & iCtr & "»" & Trim(objRecordSet.Fields("item_code").Value) & "»" & strGlGroupId & "»0»" & strItemGL & "»" & strItemSL & "»" & dblBaseCurrencyAmount & "»Cr»»" & strTaxCCCode & "»»»»0»0»0»0»0" & "¦"
                            'mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»ITM»SAL»" & iCtr & "»" & Trim(objRecordSet.Fields("item_code").Value) & "»" & strGlGroupId & "»0»" & strItemGL & "»" & strItemSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                        Else
                            'mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strItemGL & "»" & strItemSL & "»»»»CR»" & dblBaseCurrencyAmount & "»»Basic for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                            mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strItemGL & "»" & strItemSL & "»" & strTaxCCCode & "»»»CR»" & dblBaseCurrencyAmount & "»»Basic for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                        End If
                        iCtr = iCtr + 1
                    End If

                End If
            End If


            '*******************************************************************************
            'EXC Duty Posting
            '*******************************************************************************
            'IF Condition added by nisha for Excise Exumption on 10/07/2003
            If blnExciseExumpted = False Then
                If mblnEOUUnit = False Then

                    dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("Excise_Tax").Value), 0, objRecordSet.Fields("Excise_Tax").Value)
                Else
                    'code changed by nisha on 30/08/2003 for DTA Invoices

                    dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("TotalExciseAmount").Value), 0, objRecordSet.Fields("TotalExciseAmount").Value)
                    'changes Ends Here
                End If
                ''---------Added By Tapan On 8-Mar-2K3--------------------------
                If mblnExciseRoundOFFFlag Then dblTaxAmt = System.Math.Round(dblTaxAmt, 0)
                ''---------Addition Ends--------------------------
                dblBaseCurrencyAmount = dblTaxAmt
                If dblBaseCurrencyAmount > 0 Then
                    'initializing the tax gl and sl here
                    rsFULLExciseAmount.GetResult("Select Sum(isnull(TotalExciseAmount,0)) as TotalExciseAmount from Sales_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  Doc_no =" & Ctlinvoice.Text)
                    'UPGRADE_WARNING: Couldn't resolve default property of object rsFULLExciseAmount.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    dblFullExciseAmount = Val(rsFULLExciseAmount.GetValue("TotalExciseAmount"))
                    If CheckExcPriority() = 0 Then
                        If blnMsgBox = False Then
                            If MsgBox("No Excise Priority is Defined Would like to Post in ARTax ?", MsgBoxStyle.YesNo, "eMPro") = MsgBoxResult.Yes Then
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
                    If UCase(Trim(cmbInvType.Text)) <> "REJECTION" Then
                        mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»EXC»0»" & Trim(objRecordSet.Fields("item_code").Value) & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                    Else
                        mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & dblTaxAmt & "»»Excise for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                        RejectionRoundoff = RejectionRoundoff + dblTaxAmt
                    End If
                    iCtr = iCtr + 1
                End If
            End If
            'Changes Ends Here 10/07/2003


            '''***** Changes done By Ashutosh on 28-07-2006,Issue Id:18350
            '*******************************************************************************
            'Packing Value Posting
            '*******************************************************************************

            dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("ItemPacking_Amount").Value), 0, objRecordSet.Fields("ItemPacking_Amount").Value)

            dblPacking_per = IIf(IsDBNull(objRecordSet.Fields("Packing").Value), 0, objRecordSet.Fields("Packing").Value)

            dblBaseCurrencyAmount = dblTaxAmt

            '---------------Changes made by JS on 23/08/2004--------------------------------

            If dblBaseCurrencyAmount > 0 Then
                'Changed for Issue ID 20918 Starts
                strRetVal = GetTaxGlSl("PKT")
                If strRetVal = "N" Then
                    MsgBox("GL for ARTAX is not defined for PACKING", MsgBoxStyle.Information, "eMPro")
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
                'Changed for Issue ID 20918 Ends
                If UCase(Trim(cmbInvType.Text)) <> "REJECTION" Then
                    mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»PKT»0»" & Trim(objRecordSet.Fields("item_code").Value) & "»»" & dblPacking_per & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                Else
                    mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & dblTaxAmt & "»»Packing Charges for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                    RejectionRoundoff = RejectionRoundoff + dblTaxAmt
                End If
                iCtr = iCtr + 1
            End If

            'Added for Issue ID eMpro-20080430-18033 Starts
            If mblnEOUUnit = True Then
                'Posting of CVD Excise

                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
                If Trim(IIf(IsDBNull(objRecordSet.Fields("CVD_type").Value), "", objRecordSet.Fields("CVD_type").Value)) <> "" Then
                    objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("CVD_type").Value), "", objRecordSet.Fields("CVD_type").Value)) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
                    If Not objTmpRecordset.EOF Then
                        strTaxType = Trim(UCase$(objTmpRecordset.Fields("Tx_TaxeID").Value))
                    Else
                        MsgBox("Tax type not found", vbInformation, ResolveResString(100))
                        CreateStringForAccounts = False
                        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close() : objRecordSet = Nothing
                        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close() : objTmpRecordset = Nothing
                        Exit Function
                    End If
                    If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
                    If strTaxType = "CVD" Then
                        dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("CVD_Amount").Value), 0, objRecordSet.Fields("CVD_Amount").Value)
                        dblBaseCurrencyAmount = dblTaxAmt
                        dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("CVD_per").Value), 0, objRecordSet.Fields("CVD_per").Value)
                        If dblBaseCurrencyAmount > 0 Then
                            'initializing the tax gl and sl here
                            strRetVal = GetTaxGlSl(strTaxType)
                            If strRetVal = "N" Then
                                MsgBox("GL for ARTAX is not defined for " & strTaxType, vbInformation, ResolveResString(100))
                                CreateStringForAccounts = False
                                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close() : objRecordSet = Nothing
                                Exit Function
                            End If
                            varTmp = Split(strRetVal, "»")
                            strTaxGL = varTmp(0)
                            strTaxSL = varTmp(1)
                            If UCase$(Trim(cmbInvType.Text)) <> "REJECTION" Then
                                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" &
                                iCtr & "»TAX»CVD»0»" & Trim(objRecordSet.Fields("item_code").Value) &
                               "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                            Else
                                mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" &
                                         dblTaxAmt & "»»CVD for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                                RejectionRoundoff = RejectionRoundoff + dblTaxAmt
                            End If
                            iCtr = iCtr + 1
                        End If
                    End If
                End If


            End If
            'Added for Issue ID eMpro-20080430-18033 Ends

            'Added for Issue ID eMpro-20080508-18500 Starts
            If mblnEOUUnit = False And UCase$(Trim(cmbInvType.Text)) = "TRANSFER INVOICE" Then
                'Posting of SAD Tax

                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
                If Trim(IIf(IsDBNull(objRecordSet.Fields("SAD_type").Value), "", objRecordSet.Fields("SAD_type").Value)) <> "" Then
                    objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("SAD_type").Value), "", objRecordSet.Fields("SAD_type").Value)) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText) 'adOpenForwardOnly, adLockReadOnly, adCmdText)
                    If Not objTmpRecordset.EOF Then
                        strTaxType = Trim(UCase$(objTmpRecordset.Fields("Tx_TaxeID").Value))
                    Else
                        MsgBox("Tax type not found", vbInformation, ResolveResString(100))
                        CreateStringForAccounts = False
                        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close() : objRecordSet = Nothing
                        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close() : objTmpRecordset = Nothing
                        Exit Function
                    End If
                    If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
                    If strTaxType = "SAD" Then
                        dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("SVD_Amount").Value), 0, objRecordSet.Fields("SVD_Amount").Value)
                        dblBaseCurrencyAmount = dblTaxAmt
                        dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("SVD_per").Value), 0, objRecordSet.Fields("SVD_per").Value)
                        If dblBaseCurrencyAmount > 0 Then
                            'initializing the tax gl and sl here
                            strRetVal = GetTaxGlSl(strTaxType)
                            If strRetVal = "N" Then
                                MsgBox("GL for ARTAX is not defined for " & strTaxType, vbInformation, ResolveResString(100))
                                CreateStringForAccounts = False
                                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close() : objRecordSet = Nothing
                                Exit Function
                            End If
                            varTmp = Split(strRetVal, "»")
                            strTaxGL = varTmp(0)
                            strTaxSL = varTmp(1)
                            If UCase$(Trim(cmbInvType.Text)) <> "REJECTION" Then
                                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" &
                                iCtr & "»TAX»SAD»0»" & Trim(objRecordSet.Fields("item_code").Value) &
                               "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                            Else
                                mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" &
                                         dblTaxAmt & "»»SVD for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                                RejectionRoundoff = RejectionRoundoff + dblTaxAmt
                            End If
                            iCtr = iCtr + 1
                        End If
                    End If
                End If


            End If
            'Added for Issue ID eMpro-20080508-18500 Ends

            'Others Posting

            dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("Others").Value), 0, objRecordSet.Fields("Others").Value)
            dblBaseCurrencyAmount = dblTaxAmt

            If dblBaseCurrencyAmount > 0 Then
                'Changed for Issue ID 20918 Starts
                'If amount is zero then GLSL check is not require
                '---------------Changes made by JS on 23/08/2004--------------------------------
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
                '---------------Changes made by JS on 23/08/2004--------------------------------
                'Changed for Issue ID 20918 Ends

                If UCase(Trim(cmbInvType.Text)) <> "REJECTION" Then
                    mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»OTH»0»" & Trim(objRecordSet.Fields("item_code").Value) & "»»0»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                Else
                    mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & dblTaxAmt & "»»Other Charges for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                    RejectionRoundoff = RejectionRoundoff + dblTaxAmt
                End If
                iCtr = iCtr + 1
            End If
            '101188073 Start
            If gblnGSTUnit Then
                'CGST
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
                If Trim(IIf(IsDBNull(objRecordSet.Fields("CGSTTXRT_TYPE").Value), "", objRecordSet.Fields("CGSTTXRT_TYPE").Value)) <> "" Then
                    objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("CGSTTXRT_TYPE").Value), "", objRecordSet.Fields("CGSTTXRT_TYPE").Value)) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
                    If Not objTmpRecordset.EOF Then
                        strTaxType = Trim(UCase$(objTmpRecordset.Fields("Tx_TaxeID").Value))
                    Else
                        MsgBox("Tax type not found", vbInformation, ResolveResString(100))
                        CreateStringForAccounts = False
                        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close() : objRecordSet = Nothing
                        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close() : objTmpRecordset = Nothing
                        Exit Function
                    End If
                    If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
                    If strTaxType = "CGST" Then
                        dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("CGST_AMT").Value), 0, objRecordSet.Fields("CGST_AMT").Value)
                        dblBaseCurrencyAmount = dblTaxAmt
                        dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("CGST_PERCENT").Value), 0, objRecordSet.Fields("CGST_PERCENT").Value)
                        If dblBaseCurrencyAmount > 0 Then
                            strRetVal = GetTaxGlSl(strTaxType)
                            If strRetVal = "N" Then
                                MsgBox("GL for ARTAX is not defined for " & strTaxType, vbInformation, ResolveResString(100))
                                CreateStringForAccounts = False
                                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close() : objRecordSet = Nothing
                                Exit Function
                            End If
                            varTmp = Split(strRetVal, "»")
                            strTaxGL = varTmp(0)
                            strTaxSL = varTmp(1)
                            If UCase$(Trim(cmbInvType.Text)) <> "REJECTION" Then
                                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" &
                                iCtr & "»TAX»CGST»0»" & Trim(objRecordSet.Fields("item_code").Value) &
                               "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                            Else
                                mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" &
                                         dblTaxAmt & "»»CGST for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                                RejectionRoundoff = RejectionRoundoff + dblTaxAmt
                            End If
                            iCtr = iCtr + 1
                        End If
                    End If
                End If
                'SGST
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
                If Trim(IIf(IsDBNull(objRecordSet.Fields("SGSTTXRT_TYPE").Value), "", objRecordSet.Fields("SGSTTXRT_TYPE").Value)) <> "" Then
                    objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("SGSTTXRT_TYPE").Value), "", objRecordSet.Fields("SGSTTXRT_TYPE").Value)) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
                    If Not objTmpRecordset.EOF Then
                        strTaxType = Trim(UCase$(objTmpRecordset.Fields("Tx_TaxeID").Value))
                    Else
                        MsgBox("Tax type not found", vbInformation, ResolveResString(100))
                        CreateStringForAccounts = False
                        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close() : objRecordSet = Nothing
                        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close() : objTmpRecordset = Nothing
                        Exit Function
                    End If
                    If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
                    If strTaxType = "SGST" Then
                        dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("SGST_AMT").Value), 0, objRecordSet.Fields("SGST_AMT").Value)
                        dblBaseCurrencyAmount = dblTaxAmt
                        dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("SGST_PERCENT").Value), 0, objRecordSet.Fields("SGST_PERCENT").Value)
                        If dblBaseCurrencyAmount > 0 Then
                            strRetVal = GetTaxGlSl(strTaxType)
                            If strRetVal = "N" Then
                                MsgBox("GL for ARTAX is not defined for " & strTaxType, vbInformation, ResolveResString(100))
                                CreateStringForAccounts = False
                                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close() : objRecordSet = Nothing
                                Exit Function
                            End If
                            varTmp = Split(strRetVal, "»")
                            strTaxGL = varTmp(0)
                            strTaxSL = varTmp(1)
                            If UCase$(Trim(cmbInvType.Text)) <> "REJECTION" Then
                                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" &
                                iCtr & "»TAX»SGST»0»" & Trim(objRecordSet.Fields("item_code").Value) &
                               "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                            Else
                                mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" &
                                         dblTaxAmt & "»»SGST for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                                RejectionRoundoff = RejectionRoundoff + dblTaxAmt
                            End If
                            iCtr = iCtr + 1
                        End If
                    End If
                End If
                'IGST
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
                If Trim(IIf(IsDBNull(objRecordSet.Fields("IGSTTXRT_TYPE").Value), "", objRecordSet.Fields("IGSTTXRT_TYPE").Value)) <> "" Then
                    objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("IGSTTXRT_TYPE").Value), "", objRecordSet.Fields("IGSTTXRT_TYPE").Value)) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
                    If Not objTmpRecordset.EOF Then
                        strTaxType = Trim(UCase$(objTmpRecordset.Fields("Tx_TaxeID").Value))
                    Else
                        MsgBox("Tax type not found", vbInformation, ResolveResString(100))
                        CreateStringForAccounts = False
                        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close() : objRecordSet = Nothing
                        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close() : objTmpRecordset = Nothing
                        Exit Function
                    End If
                    If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
                    If strTaxType = "IGST" Then
                        dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("IGST_AMT").Value), 0, objRecordSet.Fields("IGST_AMT").Value)
                        dblBaseCurrencyAmount = dblTaxAmt
                        dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("IGST_PERCENT").Value), 0, objRecordSet.Fields("IGST_PERCENT").Value)
                        If dblBaseCurrencyAmount > 0 Then
                            strRetVal = GetTaxGlSl(strTaxType)
                            If strRetVal = "N" Then
                                MsgBox("GL for ARTAX is not defined for " & strTaxType, vbInformation, ResolveResString(100))
                                CreateStringForAccounts = False
                                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close() : objRecordSet = Nothing
                                Exit Function
                            End If
                            varTmp = Split(strRetVal, "»")
                            strTaxGL = varTmp(0)
                            strTaxSL = varTmp(1)
                            If UCase$(Trim(cmbInvType.Text)) <> "REJECTION" Then
                                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" &
                                iCtr & "»TAX»IGST»0»" & Trim(objRecordSet.Fields("item_code").Value) &
                               "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                            Else
                                mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" &
                                         dblTaxAmt & "»»IGST for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                                RejectionRoundoff = RejectionRoundoff + dblTaxAmt
                            End If
                            iCtr = iCtr + 1
                        End If
                    End If
                End If
                'UTGST
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
                If Trim(IIf(IsDBNull(objRecordSet.Fields("UTGSTTXRT_TYPE").Value), "", objRecordSet.Fields("UTGSTTXRT_TYPE").Value)) <> "" Then
                    objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("UTGSTTXRT_TYPE").Value), "", objRecordSet.Fields("UTGSTTXRT_TYPE").Value)) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
                    If Not objTmpRecordset.EOF Then
                        strTaxType = Trim(UCase$(objTmpRecordset.Fields("Tx_TaxeID").Value))
                    Else
                        MsgBox("Tax type not found", vbInformation, ResolveResString(100))
                        CreateStringForAccounts = False
                        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close() : objRecordSet = Nothing
                        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close() : objTmpRecordset = Nothing
                        Exit Function
                    End If
                    If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
                    If strTaxType = "UTGST" Then
                        dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("UTGST_AMT").Value), 0, objRecordSet.Fields("UTGST_AMT").Value)
                        dblBaseCurrencyAmount = dblTaxAmt
                        dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("UTGST_PERCENT").Value), 0, objRecordSet.Fields("UTGST_PERCENT").Value)
                        If dblBaseCurrencyAmount > 0 Then
                            strRetVal = GetTaxGlSl(strTaxType)
                            If strRetVal = "N" Then
                                MsgBox("GL for ARTAX is not defined for " & strTaxType, vbInformation, ResolveResString(100))
                                CreateStringForAccounts = False
                                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close() : objRecordSet = Nothing
                                Exit Function
                            End If
                            varTmp = Split(strRetVal, "»")
                            strTaxGL = varTmp(0)
                            strTaxSL = varTmp(1)
                            If UCase$(Trim(cmbInvType.Text)) <> "REJECTION" Then
                                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" &
                                iCtr & "»TAX»UTGST»0»" & Trim(objRecordSet.Fields("item_code").Value) &
                               "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                            Else
                                mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" &
                                         dblTaxAmt & "»»UTGST for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                                RejectionRoundoff = RejectionRoundoff + dblTaxAmt
                            End If
                            iCtr = iCtr + 1
                        End If
                    End If
                End If
                'CCESS
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
                If Trim(IIf(IsDBNull(objRecordSet.Fields("COMPENSATION_CESS_TYPE").Value), "", objRecordSet.Fields("COMPENSATION_CESS_TYPE").Value)) <> "" Then
                    objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("COMPENSATION_CESS_TYPE").Value), "", objRecordSet.Fields("COMPENSATION_CESS_TYPE").Value)) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
                    If Not objTmpRecordset.EOF Then
                        strTaxType = Trim(UCase$(objTmpRecordset.Fields("Tx_TaxeID").Value))
                    Else
                        MsgBox("Tax type not found", vbInformation, ResolveResString(100))
                        CreateStringForAccounts = False
                        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close() : objRecordSet = Nothing
                        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close() : objTmpRecordset = Nothing
                        Exit Function
                    End If
                    If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
                    If strTaxType = "GSTEC" Then
                        dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("COMPENSATION_CESS_AMT").Value), 0, objRecordSet.Fields("COMPENSATION_CESS_AMT").Value)
                        dblBaseCurrencyAmount = dblTaxAmt
                        dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("COMPENSATION_CESS_PERCENT").Value), 0, objRecordSet.Fields("COMPENSATION_CESS_PERCENT").Value)
                        If dblBaseCurrencyAmount > 0 Then
                            strRetVal = GetTaxGlSl(strTaxType)
                            If strRetVal = "N" Then
                                MsgBox("GL for ARTAX is not defined for " & strTaxType, vbInformation, ResolveResString(100))
                                CreateStringForAccounts = False
                                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close() : objRecordSet = Nothing
                                Exit Function
                            End If
                            varTmp = Split(strRetVal, "»")
                            strTaxGL = varTmp(0)
                            strTaxSL = varTmp(1)
                            If UCase$(Trim(cmbInvType.Text)) <> "REJECTION" Then
                                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" &
                                iCtr & "»TAX»CCESS»0»" & Trim(objRecordSet.Fields("item_code").Value) &
                               "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                            Else
                                mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" &
                                         dblTaxAmt & "»»CCESS for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                                RejectionRoundoff = RejectionRoundoff + dblTaxAmt
                            End If
                            iCtr = iCtr + 1
                        End If
                    End If
                End If
            End If
            '101188073 End
            objRecordSet.MoveNext()
        End While

        'Posting of rounded off amount

        'dblBaseCurrencyAmount = dblInvoiceAmt - Round(dblInvoiceAmt, 0)
        dblBaseCurrencyAmount = dblInvoiceAmtRoundOff_diff
        'Ends here
        dblBaseCurrencyAmount = System.Math.Round(dblBaseCurrencyAmount, 4)
        If System.Math.Abs(dblBaseCurrencyAmount) > 0 Then
            'Changed for Issue ID 20918 Starts
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
            'prashant rajpal cc code changes
            strTaxCCCode = ""
            If (((UCase(Trim(cmbInvType.Text))) = "TRANSFER INVOICE") Or ((UCase(Trim(cmbInvType.Text))) = "REJECTION") Or ((UCase(Trim(cmbInvType.Text))) = "SAMPLE INVOICE")) Then
                If mblnCCFlag = True Then
                    If VerifyGLCCMappingFlag(strTaxGL) Then
                        strTaxCCCode = GetCommonCCCode("Rounded_Amt")
                        If strTaxCCCode = "N" Then
                            CreateStringForAccounts = False
                            If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                                objRecordSet.Close()

                                objRecordSet = Nothing
                            End If
                            Exit Function
                        End If
                    End If
                End If
            End If

            'prashant rajpal cc code changes ended

            'Changed for Issue ID 20918 Ends
            If UCase(Trim(cmbInvType.Text)) <> "REJECTION" Then
                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»»RND»0»" & "»»0»" & strItemGL & "»" & strItemSL & "»" & System.Math.Abs(dblBaseCurrencyAmount) & "»"
                If dblBaseCurrencyAmount < 0 Then
                    'mstrDetailString = mstrDetailString & "Cr»»»»»»0»0»0»0»0" & "¦"
                    mstrDetailString = mstrDetailString & "Cr»»" & strTaxCCCode & "»»»»0»0»0»0»0" & "¦"
                Else
                    ' mstrDetailString = mstrDetailString & "Dr»»»»»»0»0»0»0»0" & "¦"
                    mstrDetailString = mstrDetailString & "Dr»»" & strTaxCCCode & "»»»»0»0»0»0»0" & "¦"
                End If
            Else
                If Len(Trim(strCustRef)) > 0 Then

                    If blnTotalInvoiceAmountRoundOff = True Then
                        RejectionRoundoff = System.Math.Round(RejectionRoundoff - RejectionTotalAmount, Len(CStr(RejectionRoundoff)) - IIf(InStr(1, CStr(RejectionRoundoff), ".") > 0, InStr(1, CStr(RejectionRoundoff), "."), Len(CStr(RejectionRoundoff))))
                    Else
                        RejectionRoundoff = System.Math.Round(RejectionRoundoff - RejectionTotalAmount, intTotalInvoiceAmountRoundOff)
                    End If
                    If RejectionRoundoff < 0 Then
                        'mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strItemGL & "»" & strItemSL & "»»»»CR»" & System.Math.Abs(RejectionRoundoff) & "»Rounding off amount for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                        mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strItemGL & "»" & strItemSL & "»" & strTaxCCCode & "»»»CR»" & System.Math.Abs(RejectionRoundoff) & "»Rounding off amount for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                    Else
                        'mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strItemGL & "»" & strItemSL & "»»»»DR»" & System.Math.Abs(RejectionRoundoff) & "»Rounding off amount for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                        mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strItemGL & "»" & strItemSL & "»" & strTaxCCCode & "»»»DR»" & System.Math.Abs(RejectionRoundoff) & "»Rounding off amount for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                    End If
                Else
                    If dblBaseCurrencyAmount < 0 Then
                        ' Condition added by nisha on 13 May 2005 for rejection invoice Round off Correctin
                        'mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strItemGL & "»" & strItemSL & "»»»»CR»" & System.Math.Abs(dblBaseCurrencyAmount) & "»Rounding off amount for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                        mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strItemGL & "»" & strItemSL & "»" & strTaxCCCode & "»»»CR»" & System.Math.Abs(dblBaseCurrencyAmount) & "»Rounding off amount for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                    Else
                        'mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strItemGL & "»" & strItemSL & "»»»»DR»" & System.Math.Abs(dblBaseCurrencyAmount) & "»Rounding off amount for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                        mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strItemGL & "»" & strItemSL & "»" & strTaxCCCode & "»»»DR»" & System.Math.Abs(dblBaseCurrencyAmount) & "»Rounding off amount for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                    End If
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
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        CreateStringForAccounts = False
        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
            objRecordSet.Close()

            objRecordSet = Nothing
        End If
    End Function

    Private Function GetItemGLSL(ByVal InventoryGlGroup As String, ByVal PurposeCode As String) As String

        Dim objRecordSet As New ADODB.Recordset
        Dim strGL As String
        Dim strSL As String


        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close()
        objRecordSet.Open("SELECT invGld_glcode, invGld_slcode FROM fin_InvGLGrpDtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  invGld_prpsCode = '" & PurposeCode & "' AND invGld_invGLGrpId = '" & InventoryGlGroup & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
        If objRecordSet.EOF Then
            objRecordSet.Close()
            objRecordSet.Open("SELECT gbl_glCode, gbl_slCode FROM fin_globalGL WHERE UNIT_CODE='" + gstrUNITID + "' AND  gbl_prpsCode = '" & PurposeCode & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
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
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        GetItemGLSL = "N"
        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
            objRecordSet.Close()

            objRecordSet = Nothing
        End If
    End Function
    Public Sub UpdateinSale_Dtl()
        Dim rssaledtl As ClsResultSetDB_Invoice
        Dim rsSaleConf As ClsResultSetDB_Invoice
        Dim rsSalesChallan As ClsResultSetDB_Invoice
        Dim strSql As String
        Dim strInvoiceDate As String
        Dim strStockLocCode As String
        Dim rsSalesParameter As New ClsResultSetDB_Invoice
        Dim intRow, intLoopCount As Short
        Dim mItem_Code, mCust_Item_Code As String
        Dim mSales_Quantity As Double
        Dim mToolCost As Double
        Dim blnCheckToolCost As Boolean
        Dim strAccountCode As String
        Dim rsbom As New ClsResultSetDB_Invoice
        Dim irowcount As Short
        Dim intRwCount1 As Short
        Dim strItembal As String
        Dim strQuantity As String
        Dim varItemQty1 As String
        Dim rsMktSchedule As New ClsResultSetDB_Invoice
        Dim strToolCode As String
        strupdateitbalmst = ""
        strupdatecustodtdtl = ""
        strUpdateAmorDtl = ""
        strupdateamordtlbom = ""
        On Error GoTo Err_Handler
        'CODE ADDED BY NISHA ON 21/03/2003 FOR FINANCIAL ROLLOVER
        strSql = "select * from Saleschallan_Dtl where UNIT_CODE='" + gstrUNITID + "' AND  Doc_No =" & Me.Ctlinvoice.Text & "  and Location_Code='" & Trim(txtUnitCode.Text) & "'"
        rsSalesChallan = New ClsResultSetDB_Invoice
        rsSalesChallan.GetResult(strSql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        'UPGRADE_WARNING: Couldn't resolve default property of object rsSalesChallan.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        strInvoiceDate = VB6.Format(rsSalesChallan.GetValue("Invoice_Date"), gstrDateFormat)
        'UPGRADE_WARNING: Couldn't resolve default property of object rsSalesChallan.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        strAccountCode = rsSalesChallan.GetValue("Account_code")
        rsSaleConf = New ClsResultSetDB_Invoice
        rsSaleConf.GetResult("Select Stock_Location from saleconf WHERE UNIT_CODE='" + gstrUNITID + "' AND  Description = '" & Me.cmbInvType.Text & "' and Sub_Type_Description ='" & Me.CmbCategory.Text & "' and Location_Code='" & Trim(txtUnitCode.Text) & "'and datediff(dd,'" & getDateForDB(strInvoiceDate) & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0")
        'UPGRADE_WARNING: Couldn't resolve default property of object rsSaleConf.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        strStockLocCode = rsSaleConf.GetValue("Stock_Location")
        '' added by priti for transfer rejection invoice
        Dim isRejectionInvoice As Boolean = Find_Value("Select isnull(is_Rejection_invoice,0) from saleschallan_dtl  WHERE UNIT_CODE='" + gstrUNITID + "' and doc_no='" & Ctlinvoice.Text & "'")
        If isRejectionInvoice Then
            strStockLocation = Find_Value("Select Transfer_RejectionLoc from sales_parameter WHERE UNIT_CODE='" + gstrUNITID + "'")
            strsaleconfLocation = strStockLocation
        End If
        '' end by priti 
        strSql = "Select * from sales_Dtl where UNIT_CODE='" + gstrUNITID + "' AND  Doc_No = " & Me.Ctlinvoice.Text & " and Location_Code='" & Trim(txtUnitCode.Text) & "'"

        rsSalesParameter.GetResult("Select CheckToolAmortisation from Sales_Parameter WHERE UNIT_CODE='" + gstrUNITID + "'")
        If rsSalesParameter.GetNoRows > 0 Then
            rsSalesParameter.MoveFirst()
            'UPGRADE_WARNING: Couldn't resolve default property of object rsSalesParameter.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If Len(Trim(rsSalesParameter.GetValue("CheckToolAmortisation"))) = 0 Then
                MsgBox("First define Check Tool Amortisation in Sales Parameter", MsgBoxStyle.Information, "eMPro")
                Exit Sub
            End If
            'UPGRADE_WARNING: Couldn't resolve default property of object rsSalesParameter.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            blnCheckToolCost = rsSalesParameter.GetValue("CheckToolAmortisation")
        End If
        rssaledtl = New ClsResultSetDB_Invoice
        rssaledtl.GetResult(strSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rssaledtl.GetNoRows > 0 Then
            intRow = rssaledtl.GetNoRows
            rssaledtl.MoveFirst()
            For intLoopCount = 1 To intRow
                If Not rssaledtl.EOFRecord Then

                    'UPGRADE_WARNING: Couldn't resolve default property of object rssaledtl.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    mItem_Code = rssaledtl.GetValue("Item_Code")
                    'UPGRADE_WARNING: Couldn't resolve default property of object rssaledtl.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    mCust_Item_Code = rssaledtl.GetValue("Cust_Item_Code")
                    'UPGRADE_WARNING: Couldn't resolve default property of object rssaledtl.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    mSales_Quantity = IIf(rssaledtl.GetValue("Sales_Quantity") = "", 0, rssaledtl.GetValue("Sales_Quantity"))
                    'UPGRADE_WARNING: Couldn't resolve default property of object rssaledtl.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    mToolCost = Val(rssaledtl.GetValue("toolCost_amount"))
                    'Added for Issue ID 19992 Starts
                    'UPGRADE_WARNING: Couldn't resolve default property of object rssaledtl.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    mCust_Ref = rssaledtl.GetValue("Cust_ref")
                    'UPGRADE_WARNING: Couldn't resolve default property of object rssaledtl.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    mAmendment_No = rssaledtl.GetValue("Amendment_no")
                    'Added for Issue ID 19992 Starts
                    strSelectItmbalmst = Trim(strSelectItmbalmst) & "Select cur_bal from Itembal_mst "
                    strSelectItmbalmst = strSelectItmbalmst & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Location_code = '" & strStockLocation
                    strSelectItmbalmst = strSelectItmbalmst & "' and item_code = '" & mItem_Code & "'»"

                    strupdateitbalmst = Trim(strupdateitbalmst) & "Update Itembal_mst set cur_bal= cur_bal-"
                    strupdateitbalmst = strupdateitbalmst & mSales_Quantity & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Location_code = '" & strsaleconfLocation
                    strupdateitbalmst = strupdateitbalmst & "' and item_code = '" & mItem_Code & "'»"

                    strupdatecustodtdtl = Trim(strupdatecustodtdtl) & "Update Cust_ord_dtl set Despatch_Qty = Despatch_Qty + "
                    strupdatecustodtdtl = strupdatecustodtdtl & mSales_Quantity & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Account_code ='"
                    strupdatecustodtdtl = strupdatecustodtdtl & mAccount_Code & "'and Cust_DrgNo = '"
                    strupdatecustodtdtl = strupdatecustodtdtl & mCust_Item_Code & "' and Cust_ref = '" & mCust_Ref
                    strupdatecustodtdtl = strupdatecustodtdtl & "'and amendment_no = '" & mAmendment_No & "' and active_Flag ='A'"
                    '***********To check if Tool Cost Deduction will be done or Not on 16/02/2004
                    If blnCheckToolCost = True Then
                        strItembal = "select BalanceQty = isnull(a.proj_qty,0) - isnull(a.ClosingValueSMIEL,0),a.Tool_C from Amor_dtl a,Tool_Mst b"
                        strItembal = strItembal & " WHERE a.unit_code = b.unit_code and  a.UNIT_CODE='" + gstrUNITID + "' AND  account_code = '" & strAccountCode & "'"
                        strItembal = strItembal & " and Item_code = '" & mItem_Code & "' and a.Tool_c = b.tool_c and a.Item_code = b.Product_No order by a.tool_c"
                        rsMktSchedule.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        If rsMktSchedule.GetNoRows > 0 Then
                            rsMktSchedule.MoveFirst()
                            'UPGRADE_WARNING: Couldn't resolve default property of object rsMktSchedule.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            strQuantity = CStr(Val(rsMktSchedule.GetValue("BalanceQty")))
                            'Changes Done By nisha on 22 Nov
                            'UPGRADE_WARNING: Couldn't resolve default property of object rsMktSchedule.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            strToolCode = rsMktSchedule.GetValue("tool_c")
                            strItembal = "select BalanceQty = sum(isnull(usedProjQty,0)) from Amor_dtl a "
                            strItembal = strItembal & " WHERE a.UNIT_CODE='" + gstrUNITID + "' AND  Item_code = '" & mItem_Code & "' and a.Tool_c = '" & strToolCode & "'"
                            rsMktSchedule.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                            'UPGRADE_WARNING: Couldn't resolve default property of object rsMktSchedule.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            strQuantity = CStr(Val(strQuantity) - Val(rsMktSchedule.GetValue("BalanceQty")))
                            'changes ends Here by nisha on 22 Nov
                            If Val(CStr(mSales_Quantity)) > Val(strQuantity) Then
                                If Val(strQuantity) = 0 Then
                                    MsgBox("No Balance Available for Item (" & mItem_Code & ") and customer Part Code (" & mCust_Item_Code & ") For Amortisation Calculations. ", MsgBoxStyle.OkOnly, "eMPro")
                                Else
                                    MsgBox("Quantity should not be Greater then available Balance Quantity for Amortisarion " & strQuantity, MsgBoxStyle.OkOnly, "eMPro")

                                End If
                                Exit Sub
                            Else
                                'Changes Done By nisha on 22 Nov added Item Clouse as well in Where Condition
                                strUpdateAmorDtl = Trim(strUpdateAmorDtl) & " Update Amor_dtl set usedProjQty = "
                                strUpdateAmorDtl = Trim(strUpdateAmorDtl) & " isnull(usedProjQty,0) + " & mSales_Quantity
                                strUpdateAmorDtl = Trim(strUpdateAmorDtl) & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  account_code = '" & strAccountCode
                                strUpdateAmorDtl = Trim(strUpdateAmorDtl) & "' and tool_c = '" & strToolCode & "'"
                                strUpdateAmorDtl = Trim(strUpdateAmorDtl) & " and item_code = '" & mItem_Code & "'"
                                'Changes ends Here by nisha 22 Nov
                            End If
                        Else
                            'Commented By nisha on 20 oct 2004 for removing the check of finished good in Amor_dtl Table
                            '                        MsgBox "No Record Available in Tool Amortisation Master for Item (" & mItem_Code & ") and customer Part Code (" & mCust_Item_Code & ") For Amortisation Calculations. ", vbOKOnly, "eMPro"
                            '                        Exit Sub
                            'Changes Done By nisha on 22 Nov
                            strItembal = "select BalanceQty = isnull(proj_qty,0) - isnull(ClosingValueSMIEL,0) from Amor_dtl "
                            strItembal = strItembal & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  account_code = '" & strAccountCode & "'"
                            strItembal = strItembal & " and Item_code = '" & mItem_Code & "' "
                            rsMktSchedule.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                            If rsMktSchedule.GetNoRows > 0 Then
                                rsMktSchedule.MoveFirst()
                                'UPGRADE_WARNING: Couldn't resolve default property of object rsMktSchedule.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                strQuantity = CStr(Val(rsMktSchedule.GetValue("BalanceQty")))
                                strItembal = "select BalanceQty = sum(isnull(usedProjQty,0)) from Amor_dtl "
                                strItembal = strItembal & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Item_code = '" & mItem_Code & "'"
                                rsMktSchedule.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                                'UPGRADE_WARNING: Couldn't resolve default property of object rsMktSchedule.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                strQuantity = CStr(Val(strQuantity) - Val(rsMktSchedule.GetValue("BalanceQty")))

                                If Val(CStr(mSales_Quantity)) > Val(strQuantity) Then
                                    If Val(strQuantity) = 0 Then
                                        MsgBox("No Balance Available for Item (" & mItem_Code & ") and customer Part Code (" & mCust_Item_Code & ") For Amortisation Calculations. ", MsgBoxStyle.OkOnly, "eMPro")
                                    Else
                                        MsgBox("Quantity should not be Greater then available Balance Quantity for Amortisarion " & strQuantity, MsgBoxStyle.OkOnly, "eMPro")

                                    End If
                                    Exit Sub
                                Else
                                    strUpdateAmorDtl = Trim(strUpdateAmorDtl) & " Update Amor_dtl set usedProjQty = "
                                    strUpdateAmorDtl = Trim(strUpdateAmorDtl) & " isnull(usedProjQty,0) + " & mSales_Quantity
                                    strUpdateAmorDtl = Trim(strUpdateAmorDtl) & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  account_code = '" & strAccountCode & "'"
                                    strUpdateAmorDtl = Trim(strUpdateAmorDtl) & " and item_code = '" & mItem_Code & "'"
                                End If
                                'Changes ends here by nisha on 22 Nov
                            End If
                        End If
                        '************Add Rajani Kant 19/08/2004
                        With mP_Connection
                            .Execute("DELETE FROM TMPBOM WHERE UNIT_CODE='" + gstrUNITID + "' AND IP_Address='" & gstrIpaddressWinSck & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            'Added By ekta uniyal on 6 Mar 2014
                            '.Execute("BOMExplosion '" & Trim(mItem_Code) & "','" & Trim(mItem_Code) & "',1,0,0,1,'" & gstrIpaddressWinSck & "','" + gstrUNITID + "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            .Execute("BOMEXPLOSION_Hilex '" & Trim(mItem_Code) & "','" & Trim(mItem_Code) & "',1,0,0,1,'" & gstrIpaddressWinSck & "','" + gstrUNITID + "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            'End Here
                        End With
                        rsbom.GetResult("select * from tmpBOM where UNIT_CODE='" + gstrUNITID + "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        If rsbom.GetNoRows > 0 Then
                            irowcount = rsbom.GetNoRows
                            rsbom.MoveFirst()
                            For intRwCount1 = 1 To irowcount
                                strItembal = "select BalanceQty = isnull(a.proj_qty,0) - isnull(a.ClosingValueSMIEL,0),a.tool_C from Amor_dtl a, tool_mst b "
                                strItembal = strItembal & " where  A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" + gstrUNITID + "' AND account_code = '" & Trim(strAccountCode) & "'"
                                'UPGRADE_WARNING: Couldn't resolve default property of object rsbom.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                strItembal = strItembal & " and Item_code = '" & rsbom.GetValue("item_code") & "' and a.Tool_c = b.Tool_c and a.ITem_code = b.Product_no order by a.tool_c"
                                rsMktSchedule.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                                If rsMktSchedule.GetNoRows > 0 Then
                                    rsMktSchedule.MoveFirst()
                                    'UPGRADE_WARNING: Couldn't resolve default property of object rsMktSchedule.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                    strQuantity = CStr(Val(rsMktSchedule.GetValue("BalanceQty")))
                                    'UPGRADE_WARNING: Couldn't resolve default property of object rsMktSchedule.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                    strToolCode = rsMktSchedule.GetValue("tool_c")
                                    'UPGRADE_WARNING: Couldn't resolve default property of object rsbom.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                    varItemQty1 = CStr(mSales_Quantity * Val(rsbom.GetValue("grossweight")))
                                    strItembal = "select BalanceQty = sum(isnull(usedProjQty,0)) from Amor_dtl a "
                                    strItembal = strItembal & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  account_code = '" & Trim(strAccountCode) & "'"
                                    'UPGRADE_WARNING: Couldn't resolve default property of object rsbom.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                    strItembal = strItembal & " and Item_code = '" & rsbom.GetValue("item_code") & "' and a.Tool_c = '" & strToolCode & "'"
                                    'UPGRADE_WARNING: Couldn't resolve default property of object rsMktSchedule.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                    strQuantity = CStr(Val(strQuantity) - Val(rsMktSchedule.GetValue("BalanceQty")))
                                    rsMktSchedule.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                                    If Val(varItemQty1) > Val(strQuantity) Then
                                        If Val(strQuantity) = 0 Then
                                            'UPGRADE_WARNING: Couldn't resolve default property of object rsbom.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                            MsgBox("No Balance Available for Item (" & rsbom.GetValue("item_code") & ") and customer Part Code (" & mCust_Item_Code & ") For Amortisation Calculations. ", MsgBoxStyle.OkOnly, "eMPro")
                                        Else
                                            'UPGRADE_WARNING: Couldn't resolve default property of object rsbom.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                            MsgBox("Quantity should not be Greater then available Balance Quantity for Amortisarion of this Item (" & rsbom.GetValue("item_code") & ")" & mSales_Quantity, MsgBoxStyle.OkOnly, "eMPro")
                                        End If
                                        Exit Sub
                                    Else
                                        strupdateamordtlbom = Trim(strupdateamordtlbom) & " Update Amor_dtl set usedProjQty = "
                                        'UPGRADE_WARNING: Couldn't resolve default property of object rsbom.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                        strupdateamordtlbom = Trim(strupdateamordtlbom) & " isnull(usedProjQty,0) + " & mSales_Quantity * Val(rsbom.GetValue("grossweight"))
                                        strupdateamordtlbom = Trim(strupdateamordtlbom) & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  account_code = '" & strAccountCode
                                        'UPGRADE_WARNING: Couldn't resolve default property of object rsbom.GetValue(item_code). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                        strupdateamordtlbom = Trim(strupdateamordtlbom) & "' and Item_code = '" & rsbom.GetValue("item_code")
                                        strupdateamordtlbom = Trim(strupdateamordtlbom) & "' and tool_c = '" & strToolCode & "'"
                                    End If
                                End If
                                rsbom.MoveNext()
                            Next
                        End If
                    End If
                    '**************
                    rssaledtl.MoveNext()
                End If
            Next
        End If
        rssaledtl.ResultSetClose()
        'UPGRADE_NOTE: Object rssaledtl may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        rssaledtl = Nothing
        rsSaleConf.ResultSetClose()
        'UPGRADE_NOTE: Object rsSaleConf may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        rsSaleConf = Nothing
        Exit Sub
Err_Handler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub ShowCode_Desc(ByVal pstrQuery As String, ByRef pctlCode As System.Windows.Forms.TextBox, Optional ByRef pctlDesc As System.Windows.Forms.Label = Nothing)
        '--------------------------------------------------------------------------------------
        'Name       :   ShowCode_Desc
        'Type       :   Sub
        'Author     :   tapanjain
        'Arguments  :   Query(string),Code(Text Box),Description(Label)
        'Return     :   None
        'Purpose    :   Show Code and Description window and set focus on code
        '---------------------------------------------------------------------------------------
        Dim varHelp As Object
        On Error GoTo ErrHandler
        With ctlHelp
            .CreateDSN(gstrCONNECTIONSERVER, gstrCONNECTIONDESCRIPTION, gstrCONNECTIONDSN, gstrCONNECTIONDATABASE)
            .ConnectAsUser = gstrCONNECTIONUSER
            .ConnectThroughDSN = gstrCONNECTIONDSN
            .ConnectWithPWD = gstrCONNECTIONPASSWORD
        End With
        'Changing the mouse pointer
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)

        varHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, pstrQuery)
        'Changing the mouse pointer
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        If UBound(varHelp) <> -1 Then

            If varHelp(0) <> "0" Then

                pctlCode.Text = Trim(varHelp(0))
                If Not (pctlDesc Is Nothing) Then

                    pctlDesc.Text = Trim(varHelp(1))
                End If
                pctlCode.Focus()
            Else
                MsgBox("No Record Available", MsgBoxStyle.Information, "eMPro")
            End If
        End If
        Exit Sub 'This is to avoid the execution of the error handler

ErrHandler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)

    End Sub
    Public Function GenerateInvoiceNo(ByVal pstrInvoiceType As String, ByRef pstrInvoiceSubType As String, ByVal pstrRequiredDate As String, ByVal strCustomerCode As String) As String
        On Error GoTo ErrHandler
        Dim clsInstEMPDBDbase As New EMPDataBase.EMPDB(gstrUnitId)
        Dim strCheckDOcNo As String 'Gets the Doc Number from Back End
        Dim strTempSeries As String 'Find the Numeric series in Doc No
        Dim strSuffix As String 'Generate a NEW Series
        Dim strZeroSuffix As String
        Dim strFin_Start_Date As String
        Dim strFin_End_Date As String
        Dim strSql As String 'String SQL Query
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        Dim objRs As New ClsResultSetDB_Invoice
        Dim strUPDATESql As String 'String SQL Query

        If Len(Trim(pstrInvoiceType)) > 0 Then 'For Dated Docs
            If CBool(Find_Value("select SALECONF_LOCKING_OPER_DT from sales_parameter (nolock) WHERE UNIT_CODE='" + gstrUNITID + "'")) = True Then

                If DataExist("SELECT TOP 1 1 FROM SALECONF (nolock) WHERE UNIT_CODE='" + gstrUNITID + "' AND single_series=1 AND Invoice_Type ='" & pstrInvoiceType & "' " &
                         " and sub_type='" & pstrInvoiceSubType & "' and '" & getDateForDB(pstrRequiredDate) & "' between fin_start_date and fin_end_date") Then

                    strUPDATESql = " update saleConf set OPER_DT=getdate() where unit_code='" & gstrUNITID & "' and Invoice_Type ='" & pstrInvoiceType & "' "
                    strUPDATESql = strUPDATESql & " and sub_type='" & pstrInvoiceSubType & "' and single_series=1 "
                    strUPDATESql = strUPDATESql & "and '" & getDateForDB(pstrRequiredDate) & "' between fin_start_date and fin_end_date"
                Else
                    strUPDATESql = " update saleConf set OPER_DT=getdate() where unit_code='" & gstrUNITID & "' and Invoice_Type ='" & pstrInvoiceType & "' "
                    strUPDATESql = strUPDATESql & " and sub_type='" & pstrInvoiceSubType & "' and single_series=0 "
                    strUPDATESql = strUPDATESql & "and '" & getDateForDB(pstrRequiredDate) & "' between fin_start_date and fin_end_date"
                End If

            End If
            If Len(strUPDATESql) <> 0 Then
                mP_Connection.Execute(strUPDATESql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End If

            strSql = "Select Current_No,Suffix,Fin_start_date,Fin_end_Date,ISNULL(CURRENT_NO_TRF_SAMEGSTIN,0) CURRENT_NO_TRF From saleConf WHERE UNIT_CODE='" + gstrUNITID + "' AND  "
            strSql = strSql & "Invoice_Type ='" & pstrInvoiceType & "' and  sub_type='" & pstrInvoiceSubType & "' AND Location_Code ='" & Trim(txtUnitCode.Text) &
                    "' and '" & getDateForDB(pstrRequiredDate) & "' between fin_start_date and fin_end_date"

            'With clsInstEMPDBDbase.CConnection
            '    .OpenConnection(gstrDSNName, gstrDatabaseName)
            '    .ExecuteSQL("Set Dateformat 'dmy'")
            'End With
            'clsInstEMPDBDbase.CRecordset.OpenRecordset(strSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
            'If clsInstEMPDBDbase.CRecordset.Recordcount > 0 Then
            '    'Get Last Doc No Saved
            '    strCheckDOcNo = CStr(clsInstEMPDBDbase.CRecordset.GetFieldValue("Current_No", EMPDataBase.ADODataType.ADONumeric, EMPDataBase.ADOCustomFormat.CustomZeroDecimal))
            '    strSuffix = CStr(clsInstEMPDBDbase.CRecordset.GetFieldValue("suffix", EMPDataBase.ADODataType.ADONumeric, EMPDataBase.ADOCustomFormat.CustomZeroDecimal))
            '    strFin_Start_Date = CStr(clsInstEMPDBDbase.CRecordset.GetFieldValue("Fin_Start_Date", EMPDataBase.ADODataType.ADODate, EMPDataBase.ADOCustomFormat.CustomDate))
            '    strFin_End_Date = CStr(clsInstEMPDBDbase.CRecordset.GetFieldValue("Fin_End_Date", EMPDataBase.ADODataType.ADODate, EMPDataBase.ADOCustomFormat.CustomDate))
            'Else
            '    'No Records Found
            '    Err.Raise(vbObjectError + 20008, "[GenerateDocNo]", "Incorrect Parameters Passed Invoice Number cannot be Generated.")
            'End If
            'clsInstEMPDBDbase.CRecordset.CloseRecordset() 'Close Recordset

            objRs.GetResult(strSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If objRs.GetNoRows > 0 Then
                If pstrInvoiceType.ToUpper() = "TRF" Then
                    If IsGSTINSAME(strCustomerCode) Then
                        strCheckDOcNo = objRs.GetValue("CURRENT_NO_TRF")
                    Else
                        strCheckDOcNo = objRs.GetValue("Current_No")
                    End If
                Else
                    strCheckDOcNo = objRs.GetValue("Current_No")
                End If

                strSuffix = objRs.GetValue("suffix")
                strFin_Start_Date = objRs.GetValue("Fin_Start_Date")
                strFin_End_Date = objRs.GetValue("Fin_End_Date")
            Else
                Err.Raise(vbObjectError + 20008, "[GenerateDocNo]", "Incorrect Parameters Passed Invoice Number cannot be Generated.")
            End If
            objRs.ResultSetClose()
            objRs = Nothing
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
            '101188073 Start
            If gblnGSTUnit Then
                If Len(GSTUnitPrefixCode) > 0 Then
                    strTempSeries = GSTUnitPrefixCode & strTempSeries
                End If
            End If
            '101188073 End
            'UpDate Back New Number
            GenerateInvoiceNo = strTempSeries
        End If
        Exit Function
ErrHandler:
        'Dim clsErrorInst As New EMPDataBase.EMPDB
        'clsErrorInst.CError.RaiseError(20008, "[frmexptrn0006]", "[GenerateInvoiceNo]", "", "No. Not Generated For DocType = " & pstrInvoiceType & " due to [ " & Err.Description & " ].", My.Application.Info.DirectoryPath, gstrDSNName, gstrDatabaseName)
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
        objRs = Nothing
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

        Dim strSql As String
        Dim rsObj As New ADODB.Recordset

        On Error GoTo ErrHandler

        strSql = ""
        strSql = "SELECT " & pstrFieldName & " FROM Sales_Parameter WHERE UNIT_CODE='" + gstrUNITID + "'"

        If rsObj.State = 1 Then rsObj.Close()
        rsObj.Open(strSql, mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
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
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
        GetBOMCheckFlagValue = False
    End Function
    Public Function CheckExcPriority() As Boolean
        Dim strSql As String
        Dim strTaxGL As String
        Dim strTaxSL As String
        Dim rsTaxPriority As ClsResultSetDB_Invoice
        rsTaxPriority = New ClsResultSetDB_Invoice
        strSql = "Select * from Tax_PriorityMst where UNIT_CODE='" + gstrUNITID + "' "
        rsTaxPriority.GetResult(strSql)
        If rsTaxPriority.GetNoRows > 0 Then
            rsTaxPriority.MoveFirst()
            CheckExcPriority = True
            'UPGRADE_WARNING: Couldn't resolve default property of object rsTaxPriority.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If Len(Trim(rsTaxPriority.GetValue("VarExPriority1"))) = 0 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object rsTaxPriority.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If Len(Trim(rsTaxPriority.GetValue("VarExPriority2"))) = 0 Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object rsTaxPriority.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
        Dim strSql As String
        Dim strBalance As String
        Dim strExcGL As String
        Dim strExcSL As String
        Dim StrData(2) As String
        Dim strExcType As String
        Dim rsExGLSLCode As ClsResultSetDB_Invoice
        Dim rsCheckBalance As ClsResultSetDB_Invoice
        rsExGLSLCode = New ClsResultSetDB_Invoice
        rsCheckBalance = New ClsResultSetDB_Invoice
        strSql = "Select VarExPriority1,VarExGL1,VarExSL1,VarExPriority2,VarExGL2,VarExSL2,VarExPriority3,VarExGL3,VarExSL3 from Tax_PriorityMst WHERE UNIT_CODE='" + gstrUNITID + "'"
        rsExGLSLCode.GetResult(strSql)
        rsExGLSLCode.MoveFirst()
        Select Case pintPriority
            Case 1
                'UPGRADE_WARNING: Couldn't resolve default property of object rsExGLSLCode.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                strExcGL = Trim(rsExGLSLCode.GetValue("VarExGL1"))
                'UPGRADE_WARNING: Couldn't resolve default property of object rsExGLSLCode.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                strExcSL = Trim(rsExGLSLCode.GetValue("VarExSL1"))
                'UPGRADE_WARNING: Couldn't resolve default property of object rsExGLSLCode.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                strExcType = Trim(rsExGLSLCode.GetValue("VarExPriority1"))
                If Len(Trim(strExcGL)) > 0 Then
                    If Len(Trim(strExcSL)) > 0 Then
                        '********To check about in case data is found on first Priority
                        strBalance = "Select isnull(sum(br_amount),0) as br_amount From fin_balRel where br_UntCodeID = '"
                        strBalance = strBalance & Trim(txtUnitCode.Text) & "' and br_slCode = '" & strExcSL & "'"
                        strBalance = strBalance & " and br_glCode = '" & strExcGL & "'"
                        rsCheckBalance.GetResult(strBalance)
                        If rsCheckBalance.GetNoRows > 0 Then
                            rsCheckBalance.MoveFirst()
                            'UPGRADE_WARNING: Couldn't resolve default property of object rsCheckBalance.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
                'UPGRADE_WARNING: Couldn't resolve default property of object rsExGLSLCode.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                strExcGL = Trim(rsExGLSLCode.GetValue("VarExGL2"))
                'UPGRADE_WARNING: Couldn't resolve default property of object rsExGLSLCode.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                strExcSL = Trim(rsExGLSLCode.GetValue("VarExSL2"))
                'UPGRADE_WARNING: Couldn't resolve default property of object rsExGLSLCode.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                strExcType = Trim(rsExGLSLCode.GetValue("VarExPriority2"))
                If Len(Trim(strExcGL)) > 0 Then
                    If Len(Trim(strExcSL)) > 0 Then
                        '********To check about in case data is found on first Priority
                        strBalance = "Select ISNULL(sum(br_amount),0) as br_amount From fin_balRel where br_UntCodeID = '"
                        strBalance = strBalance & Trim(txtUnitCode.Text) & "' and br_slCode = '" & strExcSL & "'"
                        strBalance = strBalance & " and br_glCode = '" & strExcGL & "'"
                        rsCheckBalance.GetResult(strBalance)
                        If rsCheckBalance.GetNoRows > 0 Then
                            rsCheckBalance.MoveFirst()
                            'UPGRADE_WARNING: Couldn't resolve default property of object rsCheckBalance.GetValue(br_amount). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
                    Else
                        ReturnGLSLAccExcPriority = VB6.CopyArray(StrData)
                    End If
                Else
                    ReturnGLSLAccExcPriority = VB6.CopyArray(StrData)
                End If
            Case 3
                'UPGRADE_WARNING: Couldn't resolve default property of object rsExGLSLCode.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                strExcGL = Trim(rsExGLSLCode.GetValue("VarExGL3"))
                'UPGRADE_WARNING: Couldn't resolve default property of object rsExGLSLCode.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                strExcSL = Trim(rsExGLSLCode.GetValue("VarExSL3"))
                'UPGRADE_WARNING: Couldn't resolve default property of object rsExGLSLCode.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                strExcType = Trim(rsExGLSLCode.GetValue("VarExPriority3"))
                If Len(Trim(strExcGL)) > 0 Then
                    If Len(Trim(strExcSL)) > 0 Then
                        '********To check about in case data is found on first Priority
                        strBalance = "Select sum(isnull(br_amount,0)) as br_amount From fin_balRel where br_UntCodeID = '"
                        strBalance = strBalance & Trim(txtUnitCode.Text) & "' and br_slCode = '" & strExcSL & "'"
                        strBalance = strBalance & " and br_glCode = '" & strExcGL & "'"
                        rsCheckBalance.GetResult(strBalance)
                        If rsCheckBalance.GetNoRows > 0 Then
                            rsCheckBalance.MoveFirst()
                            'UPGRADE_WARNING: Couldn't resolve default property of object rsCheckBalance.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

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
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

    Sub printBarCode(ByVal pstrFileName As String)
        Dim varTemp As Object
        Dim strString As String
        'strString = App.Path + "\pdf-dot.bat BarCode.txt 4 2 2 1"
        'shalini
        'strString = "C:\pdf-dot.bat BarCode.txt 4 2 2 1"
        'issue id 10150806 
        'strString = gStrDriveLocation & "pdf-dot.bat BarCode.txt 4 2 2 1"
        strString = "cmd.exe /c " & strCitrix_Inv_Pronting_Loc & "pdf-dot.bat BarCode.txt 4 2 2 1"
        'issue id 10150806 end 

        varTemp = Shell(strString)
        Exit Sub
ErrHandler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Public Function UpdateMktSchedule() As Boolean

        '*******************************************************************
        'Revised By     : Manoj Vaish
        'Return Value   : TRUE  - If Successfull
        '                 FALSE - If error Occured during processing
        'Revised Date   : 15 May 2008 Issue ID :eMpro -20080516 - 18915
        'Revised History: Wrong Schedule is getting updated throug this code
        '                 Now store procedule will knock off the schedule
        '*******************************************************************

        On Error GoTo errHandler
        Dim blnDSTracking As Boolean
        Dim Com As ADODB.Command
        Dim rsGetData As ClsResultSetDB_Invoice
        Dim strMSG As String
        Dim strQuery As String
        Dim intRowCount As Integer
        Dim intTotalRows As Integer
        Dim straccountcode As String

        intTotalRows = 0
        UpdateMktSchedule = True
        'Changed for Issue ID eMpro-20080805-20745 Starts
        blnDSTracking = CBool(Find_Value("SELECT isnull(DSWiseTracking,0)as DSWiseTracking FROM sales_parameter WHERE UNIT_CODE='" + gstrUNITID + "'"))
        'Changed for Issue ID eMpro-20080805-20745 Ends

        ' 'Changed for Issue ID eMpro-20090611-32362 Starts
        straccountcode = Find_Value("select account_code from saleschallan_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  doc_no='" & mInvNo & "'")
        If GetPlantName() = "HILEX" And blnDSTracking = True Then
            If AllowASNPrinting(straccountcode) = True And (UCase(Trim(cmbInvType.Text)) = "EXPORT INVOICE") Then
                rsGetData = New ClsResultSetDB_Invoice
                strQuery = "SELECT a.item_code,a.sales_quantity,b.account_code,b.invoice_date,a.Cust_item_Code FROM sales_dtl a,saleschallan_dtl b WHERE A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" + gstrUNITID + "' AND a.doc_no=b.doc_no and a.location_Code = b.LOcation_Code AND a.doc_no='" & mInvNo & "' "

                rsGetData.GetResult(strQuery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
                If rsGetData.GetNoRows > 0 Then
                    intTotalRows = rsGetData.GetNoRows
                    rsGetData.MoveFirst()
                    For intRowCount = 1 To intTotalRows
                        Com = New ADODB.Command
                        With Com
                            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                            .CommandText = "MKT_DSSCHEDULE_KNOCKOFF_HILEX"
                            .Parameters.Append(.CreateParameter("@UNITCODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                            .Parameters.Append(.CreateParameter("@LOCATION_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 3, Trim$(txtUnitCode.Text)))
                            .Parameters.Append(.CreateParameter("@DOC_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 8, mInvNo))
                            .Parameters.Append(.CreateParameter("@ACCOUNT_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 8, Trim(rsGetData.GetValue("Account_code"))))
                            .Parameters.Append(.CreateParameter("@ITEM_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 30, Trim(rsGetData.GetValue("Item_Code"))))
                            .Parameters.Append(.CreateParameter("@CUSTITEM_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 30, Trim(rsGetData.GetValue("Cust_item_Code"))))
                            .Parameters.Append(.CreateParameter("@REQ_QTY", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, ADODB.DataTypeEnum.adCurrency, Val(rsGetData.GetValue("sales_quantity"))))
                            .Parameters.Append(.CreateParameter("@DATE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 11, getDateForDB(Trim(rsGetData.GetValue("INvoice_date")))))
                            .Parameters.Append(.CreateParameter("@USERID", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 4, gstrUserIDSelected))
                            .Parameters.Append(.CreateParameter("@MSG", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamOutput, 500))
                            .Parameters.Append(.CreateParameter("@ERR", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamOutput, 100))

                            .ActiveConnection = mP_Connection
                            .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

                            If Len(.Parameters(9).Value) > 0 Then
                                MsgBox(.Parameters(9).Value, vbInformation + vbOKOnly, ResolveResString(100))
                                UpdateMktSchedule = False
                                Com = Nothing
                                Exit Function
                            End If

                            If Len(.Parameters(8).Value) > 0 Then
                                MsgBox(.Parameters(8).Value, vbInformation + vbOKOnly, ResolveResString(100))
                                Com = Nothing
                                UpdateMktSchedule = False
                                Exit Function
                            End If

                        End With
                        rsGetData.MoveNext()
                        Com = Nothing
                    Next

                End If
                rsGetData.ResultSetClose()
                rsGetData = Nothing
                Com = Nothing
            End If
        ElseIf blnDSTracking = True Then

            rsGetData = New ClsResultSetDB_Invoice


            strQuery = "SELECT a.item_code,a.sales_quantity,b.account_code,b.invoice_date,a.Cust_item_Code FROM sales_dtl a,saleschallan_dtl b WHERE A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" + gstrUNITID + "' AND a.doc_no=b.doc_no and a.location_Code = b.LOcation_Code AND a.doc_no='" & mInvNo & "' "

            rsGetData.GetResult(strQuery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
            If rsGetData.GetNoRows > 0 Then
                intTotalRows = rsGetData.GetNoRows
                rsGetData.MoveFirst()
                For intRowCount = 1 To intTotalRows
                    Com = New ADODB.Command
                    With Com
                        .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                        .CommandText = "MKT_DSSCHEDULE_KNOCKOFF"
                        .Parameters.Append(.CreateParameter("@UNITCODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                        .Parameters.Append(.CreateParameter("@LOCATION_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 3, Trim$(txtUnitCode.Text)))
                        .Parameters.Append(.CreateParameter("@DOC_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 8, mInvNo))
                        .Parameters.Append(.CreateParameter("@ACCOUNT_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 8, Trim(rsGetData.GetValue("Account_code"))))
                        .Parameters.Append(.CreateParameter("@ITEM_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 30, Trim(rsGetData.GetValue("Item_Code"))))
                        .Parameters.Append(.CreateParameter("@CUSTITEM_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 30, Trim(rsGetData.GetValue("Cust_item_Code"))))
                        .Parameters.Append(.CreateParameter("@REQ_QTY", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, ADODB.DataTypeEnum.adCurrency, Val(rsGetData.GetValue("sales_quantity"))))
                        .Parameters.Append(.CreateParameter("@DATE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 11, getDateForDB(Trim(rsGetData.GetValue("INvoice_date")))))
                        .Parameters.Append(.CreateParameter("@USERID", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 4, gstrUserIDSelected))
                        .Parameters.Append(.CreateParameter("@MSG", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamOutput, 500))
                        .Parameters.Append(.CreateParameter("@ERR", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamOutput, 100))

                        .ActiveConnection = mP_Connection
                        .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

                        If Len(.Parameters(9).Value) > 0 Then
                            MsgBox(.Parameters(9).Value, vbInformation + vbOKOnly, ResolveResString(100))
                            UpdateMktSchedule = False
                            Com = Nothing
                            Exit Function
                        End If

                        If Len(.Parameters(8).Value) > 0 Then
                            MsgBox(.Parameters(8).Value, vbInformation + vbOKOnly, ResolveResString(100))
                            Com = Nothing
                            UpdateMktSchedule = False
                            Exit Function
                        End If

                    End With
                    rsGetData.MoveNext()
                    Com = Nothing
                Next

            End If
            rsGetData.ResultSetClose()
            rsGetData = Nothing
            Com = Nothing

        End If
        Exit Function
errHandler:
        UpdateMktSchedule = False
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)

    End Function

    Function CheckforCustSupplyMaterial(ByRef blnUpdateStock As Boolean) As Boolean

        On Error GoTo ErrHandler
        Dim objCust As New cls_JWkInv_Reconcile
        Dim strItemCodes As String
        Dim strQuanity As String
        Dim strSql As String
        Dim StrInvDate As String
        Dim strCustCode As String
        Dim strRGPNO As String

        Dim rsTmp As New ClsResultSetDB_Invoice

        strSql = "Select ref_doc_no, a.Invoice_Date, a.Doc_No, Account_code, Item_code, Sales_Quantity  from SalesChallan_DTL as a Inner Join Sales_DTL as b on a.doc_No=b.Doc_No and a.Location_code=B.Location_code AND A.UNIT_CODE=B.UNIT_CODE where A.UNIT_CODE='" + gstrUNITID + "' AND  Doc_No='" & Trim(Ctlinvoice.Text) & "'"

        rsTmp.GetResult(strSql)
        Do While Not rsTmp.EOFRecord

            If Len(Trim(strItemCodes)) = 0 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object rsTmp.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                strItemCodes = rsTmp.GetValue("Item_Code")
                'UPGRADE_WARNING: Couldn't resolve default property of object rsTmp.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                strQuanity = rsTmp.GetValue("Sales_Quantity")
            Else
                'UPGRADE_WARNING: Couldn't resolve default property of object rsTmp.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                strItemCodes = strItemCodes & "," & rsTmp.GetValue("Item_Code")
                'UPGRADE_WARNING: Couldn't resolve default property of object rsTmp.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                strQuanity = strQuanity & "," & rsTmp.GetValue("Sales_Quantity")
            End If

            'UPGRADE_WARNING: Couldn't resolve default property of object rsTmp.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            StrInvDate = rsTmp.GetValue("Invoice_date")
            'UPGRADE_WARNING: Couldn't resolve default property of object rsTmp.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            strCustCode = rsTmp.GetValue("Account_code")
            'UPGRADE_WARNING: Couldn't resolve default property of object rsTmp.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            strRGPNO = rsTmp.GetValue("REF_DOC_NO")

            rsTmp.MoveNext()
        Loop

        objCust.ConnectionString = gstrCONNECTIONSTRING
        objCust.OpenConnection()

        If Len(Trim(strItemCodes)) = 0 Then
            CheckforCustSupplyMaterial = False
            MsgBox("No Item found.", MsgBoxStyle.Information, "eMPro")
            Exit Function
        End If

        If objCust.IsStockExist_Items(strItemCodes, strQuanity, strCustCode, Trim(strRGPNO)) = False Then
            CheckforCustSupplyMaterial = False
        Else
            CheckforCustSupplyMaterial = True

            If blnUpdateStock = True Then

                objCust.Invoice_No = Trim(Ctlinvoice.Text)
                objCust.Invoice_Date = getDateForDB(VB6.Format(StrInvDate, gstrDateFormat))
                objCust.Invoice_Type = "o"

                objCust.UserId = mP_User

                If objCust.Update_Customer_Stocks = True Then
                    CheckforCustSupplyMaterial = True
                Else
                    CheckforCustSupplyMaterial = False
                End If

            End If

        End If

        rsTmp.ResultSetClose()
        'UPGRADE_NOTE: Object rsTmp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        rsTmp = Nothing

        Exit Function
ErrHandler:
        gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

    Private Function GetNextWorkingDay(ByVal pstrDate As String) As String
        '----------------------------------------------------------------------------
        'Author         :   Ashutosh Verma
        'Argument       :   Invoice Date
        'Return Value   :   Next working day from Invoice date.
        'Function       :   Return Next working day from Invoice date.
        'Comments       :   created on 17-11-2005,Issue id:16240
        '----------------------------------------------------------------------------

        Dim rsCalendarDate As New ADODB.Recordset
        Dim strCalDate As String
        Dim strQuery As String

        On Error GoTo ErrHandler

        strQuery = "select dt from calendar_mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  dt > '" & getDateForDB(pstrDate) & "' and work_flg<>1 order by dt"
        If rsCalendarDate.State = ADODB.ObjectStateEnum.adStateOpen Then rsCalendarDate.Close()
        rsCalendarDate.Open(strQuery, mP_Connection, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockBatchOptimistic)

        If rsCalendarDate.EOF Or rsCalendarDate.BOF Or IsDBNull(rsCalendarDate.Fields("dt").Value) Then
            MsgBox("Date in Calendar Master not defined !", MsgBoxStyle.Information, "eMPro")
            GetNextWorkingDay = CStr(-1)
            'mP_Connection.Execute "SET DATEFORMAT 'dmy'"
            rsCalendarDate.Close()
            Exit Function
        Else
            rsCalendarDate.MoveFirst()
            GetNextWorkingDay = VB6.Format(rsCalendarDate.Fields("DT").Value, "dd/mmm/yyyy")
        End If
        rsCalendarDate.Close()

        Exit Function
ErrHandler:
        gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)

    End Function


    Private Function InvoicePostingFlag() As Boolean
        '-----------------------------------------------------------------------------------
        'Created By      : Ashutosh Verma
        'Issue ID        : 19934
        'Creation Date   : 11 May 2007
        'Function        : Check Posting flag from Sales Parameter
        '-----------------------------------------------------------------------------------

        On Error GoTo ErrHandler
        Dim rsInvPost As ClsResultSetDB_Invoice
        rsInvPost = New ClsResultSetDB_Invoice
        rsInvPost.GetResult("select isnull(postinfin,0) as postinfin from sales_parameter WHERE UNIT_CODE='" + gstrUNITID + "'")
        If rsInvPost.GetNoRows > 0 Then
            InvoicePostingFlag = rsInvPost.GetValue("postinfin")
        Else
            InvoicePostingFlag = False
        End If
        rsInvPost.ResultSetClose()

        Exit Function
ErrHandler:
        InvoicePostingFlag = False
        Call gobjError.RAISEERROR_INVOICE(Err.Number, err.Source, Err.Description, mp_connection)
    End Function

    Private Function VerifyInvPostingFlag() As Boolean
        '-----------------------------------------------------------------------------------
        'Created By      : Ashutosh Verma
        'Issue ID        : 19934
        'Creation Date   : 11 May 2007
        'Function        : Check Verify Invoice Posting flag from Sales Parameter
        '-----------------------------------------------------------------------------------

        On Error GoTo ErrHandler
        Dim rsVerifyInvPostFlag As ClsResultSetDB_Invoice
        rsVerifyInvPostFlag = New ClsResultSetDB_Invoice

        rsVerifyInvPostFlag.GetResult("select isnull(VerifyInvPosting,0) as VerifyInvPosting from sales_parameter WHERE UNIT_CODE='" + gstrUNITID + "'")
        If rsVerifyInvPostFlag.GetNoRows > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object rsVerifyInvPostFlag.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            VerifyInvPostingFlag = rsVerifyInvPostFlag.GetValue("VerifyInvPosting")
        Else
            VerifyInvPostingFlag = False
        End If
        rsVerifyInvPostFlag.ResultSetClose()

        Exit Function
ErrHandler:
        VerifyInvPostingFlag = False
        Call gobjError.RAISEERROR_INVOICE(Err.Number, err.Source, Err.Description, mp_connection)
    End Function

    Private Function RejInvOptionalPostingFlag() As Boolean
        '-----------------------------------------------------------------------------------
        'Created By      : Ashutosh Verma
        'Issue ID        : 19934
        'Creation Date   : 05 Jun 2007
        'Function        : Check Rejection Invoice Optional Posting flag from Sales Parameter
        '-----------------------------------------------------------------------------------

        On Error GoTo ErrHandler
        Dim rsRejInvPostFlag As ClsResultSetDB_Invoice
        rsRejInvPostFlag = New ClsResultSetDB_Invoice

        rsRejInvPostFlag.GetResult("select isnull(RejInvOptionalPostingFlag,0) as RejInvOptionalPostingFlag from sales_parameter WHERE UNIT_CODE='" + gstrUNITID + "'")
        If rsRejInvPostFlag.GetNoRows > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object rsRejInvPostFlag.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            RejInvOptionalPostingFlag = rsRejInvPostFlag.GetValue("RejInvOptionalPostingFlag")
        Else
            RejInvOptionalPostingFlag = False
        End If
        rsRejInvPostFlag.ResultSetClose()

        Exit Function
ErrHandler:
        RejInvOptionalPostingFlag = False
        Call gobjError.RAISEERROR_INVOICE(Err.Number, err.Source, Err.Description, mp_connection)
    End Function
    Public Sub CheckMultipleSOAllowed(ByVal pInvType As String, ByVal pInvSubType As String)
        '-----------------------------------------------------------------------------------
        'Created By      : Manoj Kr.Vaish
        'Issue ID        : 19992
        'Creation Date   : 27 JUNE 2007
        'Procedure       : To Check MultipleSOAllowed for Any Invoice Type
        '-----------------------------------------------------------------------------------

        Dim rsCheckSo As ClsResultSetDB_Invoice
        Dim strSql As String

        On Error GoTo ErrHandler

        rsCheckSo = New ClsResultSetDB_Invoice
        strSql = "select isnull(sorequired,0) as SORequired,isnull(MultipleSOAllowed,0) as MultipleSOAllowed from saleconf WHERE UNIT_CODE='" + gstrUNITID + "' AND  description='" & Trim(pInvType) & "' and sub_type_description='" & Trim(pInvSubType) & "' and (fin_start_date <= getdate() and fin_end_date >= getdate())"
        rsCheckSo.GetResult(strSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsCheckSo.GetNoRows > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object rsCheckSo.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            mblnMultipleSOAllowed = rsCheckSo.GetValue("MultipleSOAllowed")
            'UPGRADE_WARNING: Couldn't resolve default property of object rsCheckSo.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            mblnSORequired = rsCheckSo.GetValue("SORequired")
        End If
        rsCheckSo.ResultSetClose()
        'UPGRADE_NOTE: Object rsCheckSo may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        rsCheckSo = Nothing
        Exit Sub
ErrHandler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, err.Source, Err.Description, mp_connection)
        Exit Sub
        '------------------------------------------------------------------------------------------
    End Sub
    Public Function CheckBalanceForPrinting() As Object
        '-------------------------------------------------------------------------------------------
        ' Author        : Manoj Kr. Vaish
        ' Arguments     : NIL
        ' Return Value  : 'Error'  - If error occured during processing
        '                 Msg if Balance doesn't exist for Item(s)
        ' Function      : To Check Current Balance and Pending SO Quantity
        ' Datetime      : 27 JUNE 2007
        ' Issue ID      : 19992
        '----------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler

        Dim strSalesDtl As String
        Dim rsSalesDtl As ClsResultSetDB_Invoice
        Dim Com As ADODB.Command
        Dim strMSG As String
        Dim intCtr As Short
        Dim intRowCount As Short

        rsSalesDtl = New ClsResultSetDB_Invoice

        strSalesDtl = "Select a.Sales_Quantity,a.Item_code,a.Cust_Item_Code,a.cust_ref,a.amendment_no,a.doc_no,b.account_code from sales_Dtl a ,saleschallan_dtl b where A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" + gstrUNITID + "' AND a.Doc_No = " & Trim(Ctlinvoice.Text) & " and a.Location_Code='" & Trim(txtUnitCode.Text) & "' and a.doc_no=b.doc_no"
        rsSalesDtl.GetResult(strSalesDtl)
        If rsSalesDtl.GetNoRows > 0 Then
            intRowCount = rsSalesDtl.GetNoRows
            rsSalesDtl.MoveFirst()
            For intCtr = 1 To intRowCount
                Com = New ADODB.Command
                Com.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                Com.CommandTimeout = 0
                Com.CommandText = "CHECK_BALANCEQTY"
                Com.Parameters.Append(Com.CreateParameter("@UNITCODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                Com.Parameters.Append(Com.CreateParameter("@DOC_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 16, rsSalesDtl.GetValue("doc_no")))
                Com.Parameters.Append(Com.CreateParameter("@ITEMCODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 25, rsSalesDtl.GetValue("item_code")))
                Com.Parameters.Append(Com.CreateParameter("@ACCOUNTCODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 16, rsSalesDtl.GetValue("account_code")))
                Com.Parameters.Append(Com.CreateParameter("@CUSTDRGNO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 25, rsSalesDtl.GetValue("cust_item_code")))
                Com.Parameters.Append(Com.CreateParameter("@SALE_QTY", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, ADODB.DataTypeEnum.adCurrency, rsSalesDtl.GetValue("sales_quantity")))
                Com.Parameters.Append(Com.CreateParameter("@CUST_REF", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 25, rsSalesDtl.GetValue("cust_ref")))
                Com.Parameters.Append(Com.CreateParameter("@AMENDMENT_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 25, rsSalesDtl.GetValue("amendment_no")))
                Com.Parameters.Append(Com.CreateParameter("@MSG", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamOutput, 500))
                Com.Parameters.Append(Com.CreateParameter("@ERR", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamOutput, 100))

                Com.let_ActiveConnection(mp_connection)
                Com.Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

                If Len(Com.Parameters(8).Value) > 0 Then
                    MsgBox(Com.Parameters(8).Value, MsgBoxStyle.Information + MsgBoxStyle.OKOnly, ResolveResString(100))
                    'UPGRADE_WARNING: Couldn't resolve default property of object CheckBalanceForPrinting. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    CheckBalanceForPrinting = "Error"
                    'UPGRADE_NOTE: Object Com may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                    Com = Nothing
                    Exit Function
                End If

                If Len(Com.Parameters(7).Value) > 0 Then
                    strMSG = strMSG & Com.Parameters(7).Value
                End If

                'UPGRADE_NOTE: Object Com may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                Com = Nothing
                rsSalesDtl.MoveNext()
            Next intCtr
            'UPGRADE_WARNING: Couldn't resolve default property of object CheckBalanceForPrinting. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            CheckBalanceForPrinting = strMSG
        End If
        Exit Function
ErrHandler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, err.Source, Err.Description, mp_connection)

    End Function
    Private Function ToyotaTextFile(ByRef pstrInvoice As String) As String
        '----------------------------------------------------------------------------
        'Author         :   Ashutosh Verma
        'Argument       :   Invoice Numbers.
        'Return Value   :   Message string with value False or True.
        'Function       :   Generate text file for Toyota Barcode.
        'Comments       :   Date: 20 Aug 2007 ,Issue Id: 20876
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim intCount As Short
        Dim strLocation As String
        Dim strFileName As String
        Dim intLineNo As Short
        Dim strSql As String
        Dim strRecord As String
        Dim rsInvoice As New ClsResultSetDB_Invoice
        Dim rsEx_Cess As ClsResultSetDB_Invoice
        Dim rsSalesDtl As ClsResultSetDB_Invoice
        Dim strInvoice_item As String
        Dim dblTotalExciseAmt As Double
        Dim FSO As New Scripting.FileSystemObject
        Dim strCheckSheetNo As String

        strLocation = Trim(Find_Value("SELECT ISNULL(ToyotaTextFileLocation,'') FROM SALES_PARAMETER WHERE UNIT_CODE='" + gstrUNITID + "'"))
        If Len(strLocation) = 0 Then
            ToyotaTextFile = "FALSE|Default location not defined in sales_parameter."
            Exit Function
        Else
            If Not FSO.FolderExists(strLocation) Then
                FSO.CreateFolder(strLocation)
            End If

            If Mid(Trim(strLocation), Len(Trim(strLocation))) <> "\" Then
                strLocation = strLocation & "\"
            End If
            strFileName = "BC" & pstrInvoice & VB6.Format(Now, "ddMMyy") & ".csv"
            strFileName = strLocation & strFileName

            On Error Resume Next
            Kill(strLocation & "*.csv")
            FileClose(1)
            On Error GoTo ErrHandler

            FileOpen(1, strFileName, OpenMode.Append)
        End If

        'Changed for Issue ID eMpro-20090723-34088 Starts
        If Len(pstrInvoice) > 0 Then
            strSql = ""
            strSql = "Select a.Doc_No as Inv_No,a.CheckSheetNo,a.Total_Amount as Invoice_Amount,a.Sales_Tax_Amount,"
            strSql = strSql & " a.Invoice_date,isnull(sum(b.basic_Amount),0) as Basic_Amount,isnull(a.ECESS_Amount,0) as ECESS_Amount,"
            strSql = strSql & " isnull(a.SECESS_Amount,0) as SECESS_Amount,isnull(sum(b.cvd_amount),0)as cvd_amount From Saleschallan_dtl a"
            strSql = strSql & " Inner Join Sales_dtl b on a.doc_no=b.doc_no AND A.UNIT_CODE=B.UNIT_CODE "
            strSql = strSql & " Where A.UNIT_CODE='" + gstrUNITID + "' AND a.Doc_No='" & pstrInvoice & "'"
            strSql = strSql & " Group by a.Doc_No,a.Invoice_date,a.CheckSheetNo,a.Total_Amount,a.Sales_Tax_Amount,a.ECESS_Amount,a.SECESS_Amount"

            rsInvoice.GetResult(strSql)

            If rsInvoice.GetNoRows > 0 Then
                '''For intCount = 1 To rsInvoice.GetNoRows
                rsInvoice.MoveFirst()
                rsEx_Cess = New ClsResultSetDB_Invoice
                strInvoice_item = " select sum(excise_tax) as Total_Excise from sales_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  doc_no= '" & rsInvoice.GetValue("Inv_no") & "' "
                rsEx_Cess.GetResult(strInvoice_item)
                'Changed for Issue ID eMpro-20090723-34088 Starts
                'dblTotalExciseAmt = Val(rsEx_Cess.GetValue("Total_Excise")) + Val(rsInvoice.GetValue("Ecess_Amount")) + Val(rsInvoice.GetValue("SEcess_Amount"))
                dblTotalExciseAmt = Val(rsEx_Cess.GetValue("Total_Excise"))
                'Changed for Issue ID eMpro-20090723-34088 Ends
                rsEx_Cess = Nothing

                strRecord = ""

                strRecord = rsInvoice.GetValue("Inv_no")
                strRecord = strRecord & "," & VB6.Format(rsInvoice.GetValue("Invoice_Date"), gstrDateFormat)
                strRecord = strRecord & "," & IIf(IsDBNull(rsInvoice.GetValue("Basic_Amount")), 0, rsInvoice.GetValue("Basic_Amount"))
                strRecord = strRecord & "," & dblTotalExciseAmt
                strRecord = strRecord & "," & IIf(IsDBNull(rsInvoice.GetValue("cvd_amount")), 0, rsInvoice.GetValue("cvd_amount"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsInvoice.GetValue("Ecess_Amount")), 0, rsInvoice.GetValue("Ecess_Amount"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsInvoice.GetValue("SEcess_Amount")), 0, rsInvoice.GetValue("SEcess_Amount"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsInvoice.GetValue("Sales_Tax_Amount")), 0, rsInvoice.GetValue("Sales_Tax_Amount"))
                strCheckSheetNo = IIf(IsDBNull(rsInvoice.GetValue("CheckSheetNo")), 0, rsInvoice.GetValue("CheckSheetNo"))
                'Changed for Issue ID eMpro-20090723-34088 Ends

                PrintLine(1, strRecord) : intLineNo = intLineNo + 1

                rsSalesDtl = New ClsResultSetDB_Invoice

                strSql = " Select Distinct Cust_Item_Code,Sales_Quantity"
                strSql = strSql & " From Sales_Dtl Where Doc_No ='" & pstrInvoice & "'"
                rsSalesDtl.GetResult(strSql)

                While Not rsSalesDtl.EOFRecord
                    strRecord = ""
                    strRecord = strCheckSheetNo
                    strRecord = strRecord & "," & IIf(IsDBNull(rsSalesDtl.GetValue("Cust_Item_Code")), "", Trim(rsSalesDtl.GetValue("Cust_Item_Code")))
                    strRecord = strRecord & "," & IIf(IsDBNull(rsSalesDtl.GetValue("sales_quantity")), 0, rsSalesDtl.GetValue("sales_quantity"))
                    PrintLine(1, strRecord) : intLineNo = intLineNo + 1
                    rsSalesDtl.MoveNext()
                End While
            Else
                ToyotaTextFile = "FALSE|No Invoice Records found to generate the File."
                FileClose(1)
                Kill(strFileName)
                Exit Function
            End If

        Else
            ToyotaTextFile = "FALSE| File Not Generated."
        End If
        rsInvoice.ResultSetClose()
        rsSalesDtl.ResultSetClose()
        'Changed for Issue ID eMpro-20090723-34088 Ends

        FileClose(1)
        ToyotaTextFile = "TRUE| Toyota File Generated Successfully."
        Exit Function
ErrHandler:
        FileClose(1)
        Kill(strFileName)
        Call gobjError.RAISEERROR_INVOICE(Err.Number, err.Source, Err.Description, mp_connection)

    End Function
    Private Function InvAgstBarCode() As Boolean
        '----------------------------------------------------------------------------
        'Author         :   Manoj Kr. Vaish
        'Function       :   Get the BarCodefor Invoice from sales_parameter
        'Comments       :   Date: 19 Sep 2007 ,Issue Id: 21105
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler

        Dim strQry As String
        Dim Rs As ClsResultSetDB_Invoice
        InvAgstBarCode = False
        strQry = "Select isnull(BarCodeTrackingInInvoice,0) as BarCodeTrackingInInvoice from sales_parameter WHERE UNIT_CODE='" + gstrUNITID + "'"
        Rs = New ClsResultSetDB_Invoice
        If Rs.GetResult(strQry) = False Then GoTo ErrHandler
        'UPGRADE_WARNING: Couldn't resolve default property of object Rs.GetValue(BarCodeTrackingInInvoice). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If Rs.GetValue("BarCodeTrackingInInvoice") = "True" Then
            strQry = "Select isnull(a.BarcodeTrackingAllowed,0) as BarcodeTrackingAllowed"
            strQry = strQry & " from SaleConf a,SalesChallan_Dtl b where  A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" + gstrUNITID + "' AND Doc_No ='" & Trim(Ctlinvoice.Text) & "'"
            strQry = strQry & " and a.Invoice_Type = b.Invoice_type and a.Sub_type = b.Sub_Category and"
            strQry = strQry & " a.Location_Code = b.Location_Code And (Fin_Start_Date <= getDate() And Fin_End_Date >= getDate())"
            Call Rs.GetResult(strQry, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If Rs.GetNoRows > 0 Then
                'UPGRADE_WARNING: Couldn't resolve default property of object Rs.GetValue(BarcodeTrackingAllowed). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If Rs.GetValue("BarcodeTrackingAllowed") = "True" Then
                    InvAgstBarCode = True
                Else
                    InvAgstBarCode = False
                End If
            End If
        End If
        Rs.ResultSetClose()
        'UPGRADE_NOTE: Object Rs may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        Rs = Nothing
        Exit Function
ErrHandler:
        'UPGRADE_NOTE: Object Rs may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        Rs = Nothing
        Call gobjError.RAISEERROR_INVOICE(Err.Number, err.Source, Err.Description, mp_connection)
    End Function
    Public Function BarCodeTracking(ByVal pstrInvNo As String, ByVal pstrMode As String) As Boolean
        '----------------------------------------------------------------------------
        'Author         :   Manoj Kr. Vaish
        'Argument       :   Invoice Numbers.
        'Return Value   :   True or False
        'Function       :   Update Bar_BondedStock while invoice editing,deleting & Locking
        'Comments       :   Date: 13 Sep 2007 ,Issue Id: 21105
        'Revised By     :   Manoj Kr Vaish
        'Revision Date  :   28 Nov 2008 Issue ID : eMpro-20080930-22159
        'History        :   Functionality of Raw Material Invoice through Bar Code
        '----------------------------------------------------------------------------

        On Error GoTo ErrHandler
        Dim rsGetQty As ClsResultSetDB_Invoice
        Dim rsGetBondedQty As ClsResultSetDB_Invoice
        Dim strSql As String
        Dim blnQuantitymatch As Boolean
        rsGetQty = New ClsResultSetDB_Invoice
        rsGetBondedQty = New ClsResultSetDB_Invoice
        BarCodeTracking = False
        mblnQuantityCheck = True
        Dim strItemAlias As String = String.Empty
        Select Case pstrMode
            Case "LOCK"
                Dim isRejectionInvoice As Boolean = Find_Value("Select isnull(is_Rejection_invoice,0) from saleschallan_dtl  WHERE UNIT_CODE='" + gstrUNITID + "' and doc_no='" & Ctlinvoice.Text & "'")
                If isRejectionInvoice = False Then
                    'Changed for Issue ID eMpro-20080930-22159 Starts
                    If UCase(Trim(CmbCategory.Text)) = "RAW MATERIAL" Or UCase(Trim(CmbCategory.Text)) = "INPUTS" Or UCase(Trim(CmbCategory.Text)) = "COMPONENTS" Or UCase(Trim(CmbCategory.Text)) = "SUB ASSEMBLY" Then

                        '**************************Check Barcode Tracking Flag for RM from Item Master********************************
                        strSql = "select A.item_code,B.Itm_itemalias,isnull(Sales_Quantity,0) as Sales_Qty"
                        strSql = strSql & " from sales_dtl A Inner join Item_mst B on A.item_code=B.item_code AND A.UNIT_CODE=B.UNIT_CODE "
                        strSql = strSql & " WHERE A.UNIT_CODE='" + gstrUNITID + "' AND  B.Barcode_tracking=1 and A.doc_no='" & Trim(pstrInvNo) & "'"
                        rsGetQty.GetResult(strSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        If rsGetQty.GetNoRows > 0 Then
                            rsGetQty.MoveFirst()
                            Do While Not rsGetQty.EOFRecord
                                strSql = "select Isnull(Sum(Convert(numeric(16,4),Issue_Qty)),0) as Issue_Qty from Bar_Invoice_Issue "
                                strSql = strSql & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Issue_misno='" & Trim(pstrInvNo) & "' and invoice_status is null and dbo.UFN_BAR_GET_MAPPED_ITEM_ALIAS(unit_code,substring(Issue_partBarCode,1,8))='" & rsGetQty.GetValue("Itm_itemalias") & "'"
                                rsGetBondedQty.GetResult(strSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
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
                            'mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & "Insert into Bar_Issue(Issue_No,Issue_UserId,Issue_Unit,Issue_CurrDtTm,Issue_PTId,Issue_UId,"
                            'mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & "Issue_PartBarcode,Issue_MISNo,Issue_LocNo,Issue_Qty,Issue_Date,Issue_Time,Doc_Type,Location_code,UNIT_CODE)"
                            'mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & "Select Issue_No,Issue_UserId,Issue_Unit,Issue_CurrDtTm,Issue_PTId,Issue_UId,Issue_PartBarcode,"
                            'mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & "Issue_MISNo,Issue_LocNo,Issue_Qty,Issue_Date,Issue_Time,Doc_Type,Location_code,UNIT_CODE "
                            'mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & "from Bar_Invoice_Issue WHERE UNIT_CODE='" + gstrUNITID + "' AND  Issue_MISNo='" & Trim(pstrInvNo) & "' and Invoice_status is null" & vbCrLf

                            'strSql = "select A.CRef_PacketNo,isnull(sum(A.CRef_BalQty),0)as BarQuantity,Isnull(sum(Convert(numeric(16,4),Issue_Qty)),0)as SalesQuantity "
                            'strSql = strSql & "from Bar_CrossReference A,Bar_Invoice_Issue B where  A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE ='" + gstrUNITID + "' AND A.CRef_PacketNo=substring(B.Issue_PartbarCode,9,len(CRef_PacketNo)) and "
                            'strSql = strSql & "A.CRef_PartCode=dbo.UFN_BAR_GET_MAPPED_ITEM_ALIAS(b.unit_code,substring(b.Issue_partBarCode,1,8)) and A.CRef_Stage='B' and B.Issue_MISNO='" & Trim(pstrInvNo) & "' and Invoice_status is null group by A.CRef_PacketNo"
                            'rsGetQty.GetResult(strSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

                            'If rsGetQty.GetNoRows > 0 Then
                            '    rsGetQty.MoveFirst()
                            '    Do While Not rsGetQty.EOFRecord
                            '        If rsGetQty.GetValue("BarQuantity") - rsGetQty.GetValue("SalesQuantity") > 0 Then
                            '            mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & "update A set A.CRef_BalQty=A.CRef_BalQty-" & rsGetQty.GetValue("SalesQuantity") & ""
                            '            mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & " from Bar_CrossReference A,Bar_Invoice_Issue B"
                            '            mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & " Where A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" + gstrUNITID + "' AND  A.CRef_PartCode=dbo.UFN_BAR_GET_MAPPED_ITEM_ALIAS(b.unit_code,substring(b.Issue_partBarCode,1,8)) and A.CRef_Stage='B'"
                            '            mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & " and B.Issue_MISNO='" & Trim(pstrInvNo) & "' and A.CRef_PacketNo='" & rsGetQty.GetValue("CRef_PacketNo") & "'" & vbCrLf
                            '        ElseIf rsGetQty.GetValue("BarQuantity") - rsGetQty.GetValue("SalesQuantity") = 0 Then
                            '            mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & "update A set A.CRef_BalQty=A.CRef_BalQty-" & rsGetQty.GetValue("SalesQuantity") & ",A.CRef_Stage='I'"
                            '            mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & " from Bar_CrossReference A,Bar_Invoice_Issue B"
                            '            mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & " Where A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" + gstrUNITID + "' AND A.CRef_PartCode=dbo.UFN_BAR_GET_MAPPED_ITEM_ALIAS(b.unit_code,substring(b.Issue_partBarCode,1,8)) and A.CRef_Stage='B'"
                            '            mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & " and B.Issue_MISNO='" & Trim(pstrInvNo) & "' and A.CRef_PacketNo='" & rsGetQty.GetValue("CRef_PacketNo") & "'" & vbCrLf
                            '        ElseIf rsGetQty.GetValue("BarQuantity") - rsGetQty.GetValue("SalesQuantity") < 0 Then
                            '            MsgBox("Quantity is not available in this Packet [" & rsGetQty.GetValue("CRef_PacketNo") & "] against issued quantity.", vbInformation, ResolveResString(100))
                            '            mblnQuantityCheck = False
                            '            Exit Function
                            '        End If
                            '        rsGetQty.MoveNext()
                            '    Loop
                            'End If

                            'mstrupdateBarBondedStockFlag = mstrupdateBarBondedStockFlag & "Update Bar_Invoice_Issue Set Issue_Misno='" & mInvNo & "',Invoice_Status=1 WHERE UNIT_CODE='" + gstrUNITID + "' AND  Issue_MisNo='" & Trim(pstrInvNo) & "'" & vbCrLf
                            'mstrupdateBarBondedStockFlag = mstrupdateBarBondedStockFlag & "Update Bar_Issue Set Issue_Misno='" & mInvNo & "' WHERE UNIT_CODE='" + gstrUNITID + "' AND  Issue_MisNo='" & Trim(pstrInvNo) & "'" & vbCrLf
                            If checkBARcrossrefence_Invoicequantity(Ctlinvoice.Text) = False Then
                                mblnQuantityCheck = False
                                Exit Function
                            Else
                                BarCodeTracking = True
                            End If

                            BarCodeTracking = True
                        Else
                            BarCodeTracking = False
                        End If
                        rsGetBondedQty = Nothing
                        rsGetQty = Nothing
                    Else
                        '**************************Check Picked Quantity Against Invocie Quantity********************************
                        strSql = "select A.item_code,B.Itm_itemalias,isnull(Sales_Quantity,0) as Sales_Qty"
                        strSql = strSql & " from sales_dtl A Inner join Item_mst B on A.item_code=B.item_code AND A.UNIT_CODE=B.UNIT_CODE "
                        strSql = strSql & " where A.UNIT_CODE='" + gstrUNITID + "' AND B.Barcode_tracking=1 and A.doc_no='" & Trim(pstrInvNo) & "'"
                        rsGetQty.GetResult(strSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        If rsGetQty.GetNoRows > 0 Then
                            rsGetQty.MoveFirst()
                            Do While Not rsGetQty.EOFRecord
                                strSql = "select isnull(sum(Quantity),0) as BondedStock_Qty from bar_BondedStock_Dtl "
                                'UPGRADE_WARNING: Couldn't resolve default property of object rsGetQty.GetValue(Itm_itemalias). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                strSql = strSql & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  invoice_no='" & Trim(pstrInvNo) & "' and Status_Flag='W' and item_alias='" & rsGetQty.GetValue("Itm_itemalias") & "'"
                                rsGetBondedQty.GetResult(strSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                                'UPGRADE_WARNING: Couldn't resolve default property of object rsGetBondedQty.GetValue(BondedStock_Qty). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                'UPGRADE_WARNING: Couldn't resolve default property of object rsGetQty.GetValue(Sales_Qty). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                If rsGetQty.GetValue("Sales_Qty") = rsGetBondedQty.GetValue("BondedStock_Qty") Then
                                    blnQuantitymatch = True
                                Else
                                    MsgBox("Picked Quantity is less than Invoice Quantity.", MsgBoxStyle.Information, ResolveResString(100))
                                    mblnQuantityCheck = False
                                    Exit Function
                                End If
                                rsGetQty.MoveNext()
                            Loop

                        End If
                        '******************************Update bar Bonded Stock**********************************
                        mstrupdateBarBondedStockQty = ""
                        If blnQuantitymatch = True Then
                            strSql = "select B.Box_label,isnull(sum(A.Quantity),0)as BarQuantity,isnull(sum(B.Quantity),0)as SalesQuantity from Bar_BondedStock A,Bar_BondedStock_Dtl B where "
                            strSql = strSql & " A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" + gstrUNITID + "' AND A.Box_Label=B.Box_label and A.Status='B' and B.Status_Flag='W' and B.Invoice_No='" & Trim(pstrInvNo) & "' Group By B.Box_label"
                            rsGetQty.GetResult(strSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                            If rsGetQty.GetNoRows > 0 Then
                                rsGetQty.MoveFirst()
                                Do While Not rsGetQty.EOFRecord
                                    'UPGRADE_WARNING: Couldn't resolve default property of object rsGetQty.GetValue(SalesQuantity). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                    'UPGRADE_WARNING: Couldn't resolve default property of object rsGetQty.GetValue(BarQuantity). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                    If rsGetQty.GetValue("BarQuantity") - rsGetQty.GetValue("SalesQuantity") > 0 Then
                                        'UPGRADE_WARNING: Couldn't resolve default property of object rsGetQty.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                        mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & "update A set A.Quantity=A.Quantity-" & rsGetQty.GetValue("SalesQuantity") & ""
                                        mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & " from bar_BondedStock A,bar_BondedStock_Dtl B"
                                        mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & " Where A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" + gstrUNITID + "' AND A.Box_label = B.Box_label and A.Status='B' and B.Status_Flag='W'"
                                        'UPGRADE_WARNING: Couldn't resolve default property of object rsGetQty.GetValue(Box_label). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                        mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & " and B.Invoice_no='" & Trim(pstrInvNo) & "' and A.Box_label='" & rsGetQty.GetValue("Box_label") & "'" & vbCrLf
                                        'UPGRADE_WARNING: Couldn't resolve default property of object rsGetQty.GetValue(SalesQuantity). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                        'UPGRADE_WARNING: Couldn't resolve default property of object rsGetQty.GetValue(BarQuantity). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                    ElseIf rsGetQty.GetValue("BarQuantity") - rsGetQty.GetValue("SalesQuantity") = 0 Then
                                        'UPGRADE_WARNING: Couldn't resolve default property of object rsGetQty.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                        mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & "update A set A.Quantity=A.Quantity-" & rsGetQty.GetValue("SalesQuantity") & ",A.Status='I'"
                                        mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & " from bar_BondedStock A,bar_BondedStock_Dtl B"
                                        mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & " Where A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" + gstrUNITID + "' AND A.Box_label = B.Box_label and A.Status='B' and B.Status_Flag='W'"
                                        'UPGRADE_WARNING: Couldn't resolve default property of object rsGetQty.GetValue(Box_label). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                        mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & " and B.Invoice_no='" & Trim(pstrInvNo) & "' and A.Box_label='" & rsGetQty.GetValue("Box_label") & "'" & vbCrLf
                                        'UPGRADE_WARNING: Couldn't resolve default property of object rsGetQty.GetValue(SalesQuantity). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                        'UPGRADE_WARNING: Couldn't resolve default property of object rsGetQty.GetValue(BarQuantity). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                    ElseIf rsGetQty.GetValue("BarQuantity") - rsGetQty.GetValue("SalesQuantity") < 0 Then
                                        'UPGRADE_WARNING: Couldn't resolve default property of object rsGetQty.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                        MsgBox("Quantity is not available in this Box [" & rsGetQty.GetValue("Box_label") & "] against picked quantity.", MsgBoxStyle.Information, ResolveResString(100))
                                        mblnQuantityCheck = False
                                        Exit Function
                                    End If
                                    rsGetQty.MoveNext()
                                Loop
                                mstrupdateBarBondedStockFlag = "Update bar_BondedStock_Dtl set Status_Flag='L',Invoice_no='" & Trim(CStr(mInvNo)) & "' WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_no='" & Trim(pstrInvNo) & "' and Status_Flag='W'"
                                BarCodeTracking = True
                            End If
                        Else
                            BarCodeTracking = False
                        End If
                        'UPGRADE_NOTE: Object rsGetBondedQty may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                        rsGetBondedQty = Nothing
                        'UPGRADE_NOTE: Object rsGetQty may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                        rsGetQty = Nothing
                    End If
                    'Changed for Issue ID eMpro-20080930-22159 Ends

                Else
                    BarCodeTracking = False
                    mblnQuantityCheck = True
                End If
        End Select
        Exit Function
ErrHandler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, err.Source, Err.Description, mp_connection)
    End Function
    Private Function checkBARcrossrefence_Invoicequantity(ByVal pstrInvNo As String) As Boolean

        Dim oCmd As ADODB.Command
        Dim strMsg As String = String.Empty
        Try
            oCmd = New ADODB.Command
            With oCmd
                .let_ActiveConnection(mP_Connection)
                .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                .CommandText = "USP_CHECK_BARCROSSREFERNCE_SALESQUANTITY"
                .Parameters.Append(.CreateParameter("@Unit_Code", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, gstrUNITID))
                .Parameters.Append(.CreateParameter("@TEMP_INV_NO", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, , pstrInvNo))
                .Parameters.Append(.CreateParameter("@Msg", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamOutput, 8000))
                .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End With

            If oCmd.Parameters(oCmd.Parameters.Count - 1).Value <> "" Then
                checkBARcrossrefence_Invoicequantity = False
                MsgBox(oCmd.Parameters(oCmd.Parameters.Count - 1).Value, MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100))
                oCmd = Nothing
                Exit Function
            End If
            oCmd = Nothing

            checkBARcrossrefence_Invoicequantity = True

        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Private Function updateBARcrossrefence_Invoicequantity(ByVal pstrInvNo As String) As Boolean

        Dim oCmd As ADODB.Command
        Dim strMsg As String = String.Empty
        Try
            oCmd = New ADODB.Command
            With oCmd
                .let_ActiveConnection(mP_Connection)
                .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                .CommandText = "USP_UPDATE_BARCROSSREFERNCE_SALESQUANTITY"
                .Parameters.Append(.CreateParameter("@Unit_Code", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, gstrUNITID))
                .Parameters.Append(.CreateParameter("@TEMP_INV_NO", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, , pstrInvNo))
                .Parameters.Append(.CreateParameter("@INV_NO", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, , mInvNo))
                .Parameters.Append(.CreateParameter("@Msg", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamOutput, 1000))
                .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End With

            If oCmd.Parameters(oCmd.Parameters.Count - 1).Value <> "" Then
                updateBARcrossrefence_Invoicequantity = False
                MsgBox(oCmd.Parameters(oCmd.Parameters.Count - 1).Value, MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100))
                oCmd = Nothing
                Exit Function
            End If
            oCmd = Nothing

            updateBARcrossrefence_Invoicequantity = True

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function CheckKnocedOffQuantity(ByVal pstrAccountCode As String, ByVal pdttransdate As Date, ByVal pstrItemCode As String, ByVal pstrCustDrgNo As String) As Boolean
        '-------------------------------------------------------------------------------------------
        ' Author        : Manoj Kr. Vaish
        ' Arguments     : AccountCode,ItemCode,CustDrgNo,TransDate
        ' Return Value  : TRUE  - If Successfull
        '                 FALSE - If error Occured during processing
        ' Function      : To Check the knocked off schedule with dailymktscheudle
        ' Datetime      : 03 JAN 2008
        '--------------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim Com As ADODB.Command

        CheckKnocedOffQuantity = True

        Com = New ADODB.Command
        Com.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Com.CommandText = "MKT_CHECKKNOCKEDOFFQUANTITY"
        Com.Parameters.Append(Com.CreateParameter("@ACCOUNT_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 15, Trim(pstrAccountCode)))
        Com.Parameters.Append(Com.CreateParameter("@ITEM_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 30, Trim(pstrItemCode)))
        Com.Parameters.Append(Com.CreateParameter("@CUSTITEM_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 30, Trim(pstrCustDrgNo)))
        Com.Parameters.Append(Com.CreateParameter("@TODATE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 11, Trim(CStr(pdttransdate))))
        Com.Parameters.Append(Com.CreateParameter("@ERR", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamOutput, 100))
        Com.let_ActiveConnection(mp_connection)
        Com.Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

        If Len(Com.Parameters(4).Value) > 0 Then
            MsgBox(Com.Parameters(4).Value, MsgBoxStyle.Information + MsgBoxStyle.OKOnly, ResolveResString(100))
            CheckKnocedOffQuantity = False
            'UPGRADE_NOTE: Object Com may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            Com = Nothing
            Exit Function
        End If

        'UPGRADE_NOTE: Object Com may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        Com = Nothing

        Exit Function
ErrHandler:
        'UPGRADE_NOTE: Object Com may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        Com = Nothing
        CheckKnocedOffQuantity = False
        Call gobjError.RAISEERROR_INVOICE(Err.Number, err.Source, Err.Description, mp_connection)
    End Function

    Public Function SendMailDSOverKnockoff(ByVal pstrDocNo As String) As Boolean
        '-------------------------------------------------------------------------------------------
        ' Author        : Manoj Kr. Vaish
        ' Arguments     : Doc_No
        ' Return Value  : TRUE  - If Successfull
        '                 FALSE - If error Occured during processing
        ' Function      : To send the mail alert for over knocked off schedule with dailymktscheudle
        ' Datetime      : 07 JAN 2008
        '--------------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim Com As ADODB.Command

        SendMailDSOverKnockoff = True

        Com = New ADODB.Command
        Com.CommandTimeout = 0
        Com.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        Com.CommandText = "MKT_SENDAUTOMAIL"
        Com.Parameters.Append(Com.CreateParameter("@UNITCODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
        Com.Parameters.Append(Com.CreateParameter("@Doc_No", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, Trim(pstrDocNo)))
        Com.Parameters.Append(Com.CreateParameter("@MSG", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamOutput, 500))
        Com.let_ActiveConnection(mp_connection)
        Com.Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

        If Len(Com.Parameters(1).Value) > 0 Then
            mstrError = Com.Parameters(1).Value
            SendMailDSOverKnockoff = False
            'UPGRADE_NOTE: Object Com may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            Com = Nothing
            Exit Function
        End If

        'UPGRADE_NOTE: Object Com may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        Com = Nothing

        Exit Function
ErrHandler:
        'UPGRADE_NOTE: Object Com may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        Com = Nothing
        SendMailDSOverKnockoff = False
        Call gobjError.RAISEERROR_INVOICE(Err.Number, err.Source, Err.Description, mp_connection)
    End Function

    Private Sub Cmdinvoice_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.UCfraRepCmd.ButtonClickEventArgs) Handles Cmdinvoice.ButtonClick

        '=======================================================================================
        'Revision  By       : Ashutosh Verma,issue id:16685
        'Revision On        : 26-12-2005
        'History            : Mark Unlocked as TEMPORARY INVOICE (for SUNVAC).
        '=======================================================================================


        Dim rsSalesConf As ClsResultSetDB_Invoice
        Dim rssaledtl As ClsResultSetDB_Invoice
        Dim rsItembal As ClsResultSetDB_Invoice
        '''Dim rsCompany As ClsResultSetDB_Invoice
        Dim rsBatch As ClsResultSetDB_Invoice
        '''Dim rsSalesChallan As ClsResultSetDB_Invoice
        Dim rsSalesParameter As ClsResultSetDB_Invoice
        Dim rsBatchMst As ClsResultSetDB_Invoice
        Dim rsbom As ClsResultSetDB_Invoice
        Dim rsItemBalance As ClsResultSetDB_Invoice
        '''Changes done By Ashutosh on 13 jun 2007, Issue Id:19934
        Dim rsInvDataInFin As ClsResultSetDB_Invoice
        Dim rsInvoiceDataInFin As ClsResultSetDB_Invoice
        Dim strAsnInvoice As String

        'Added for Issue ID 19992 Starts
        Dim RdAddSold As ReportDocument
        Dim RepPath As String
        Dim Frm As Object = Nothing

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
        Dim objDrCr As prj_DrCrNote.cls_DrCrNote
        'Dim objDrCr As New cls_DrCrNote
        Dim strInvoiceDate As String
        Dim strBarcodeMsg_paratemeter As String = String.Empty
        Dim irowcount As Short
        Dim intRwCount1 As Short
        Dim varItemQty1 As Double
        Dim strToolCode As String
        Dim blnBatchTrack As Boolean
        Dim strBatchQuery As String
        Dim blnFlagTrans As Boolean
        Dim cmdObject As New ADODB.Command
        Dim strVar As String
        '''***** Added by ashutosh.
        Dim varTmp As Object
        Dim varTmp1 As Object
        Dim i As Short
        Dim dblTmpItembal As Double
        Dim dblFinalItembal As Double

        Dim strBalanceCheck As String
        'Added for Issue ID 19992 Ends
        'Added for Issue ID 21840 Starts
        Dim blnInvoiceAgainstMultipleSO As Boolean
        'Added for Issue ID 21840 Ends
        Dim blnIsReportDisplayed As Boolean = False
        Dim intmaxsubreportloop As Integer
        Dim intsubreportloopcounter As Integer
        Dim blnDSTracking As Boolean
        Dim COPYNAME As String
        Dim oCmd As ADODB.Command
        Dim dblewaymaxvalue As Double
        Dim strQry As String
        Dim intMaxLoop As Short
        Dim intNoCopies As Short
        Dim intmainloop As Short
        Dim mintnocopies As Short
        Dim strinvfilename As String = String.Empty
        Dim strpath As String = String.Empty
        Dim strfullpath As String = String.Empty
        Dim rssalesdtl As ClsResultSetDB_Invoice
        Dim blnRoundoff As Boolean
        Dim intLoopCounters As Short
        Dim intnoofRow As Short
        Dim intBasicRoundOffDecimal As Short
        Dim blninvoicelockYES As Boolean = False

        'On Error GoTo Err_Handler
        Try

            If Ctlinvoice.Text <> "" Then
                If optInvYes(0).Checked = True Then
                    dblewaymaxvalue = Find_Value("select total_amount from saleschallan_Dtl where unit_code='" + gstrUNITID + "' and doc_no =" & Ctlinvoice.Text)
                Else
                    dblewaymaxvalue = Find_Value("select total_amount from saleschallan_Dtl where unit_code='" + gstrUNITID + "' and doc_no =" & Ctlinvoice.Text)
                End If
            End If
            If Ctlinvoice.Text <> "" Then
                If optInvYes(0).Checked = True Then
                    blnRoundoff = CBool(Find_Value("select Basic_Roundoff from sales_parameter WHERE UNIT_CODE='" + gstrUNITID + "'"))
                    intBasicRoundOffDecimal = Val(Find_Value("select basic_roundoff_decimal from sales_parameter WHERE UNIT_CODE='" + gstrUNITID + "'"))

                    SALEDTL = "Select * from sales_Dtl where UNIT_CODE= '" & gstrUNITID & "' and  Doc_No = " & Me.Ctlinvoice.Text & " and Location_Code='" & Trim(txtUnitCode.Text) & "'"
                    rssalesdtl = New ClsResultSetDB_Invoice
                    rssalesdtl.GetResult(SALEDTL, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                    intnoofRow = rssalesdtl.GetNoRows
                    rssalesdtl.MoveFirst()
                    For intLoopCounters = 1 To intnoofRow
                        If blnRoundoff = True Then
                            If System.Math.Round((Val(rssalesdtl.GetValue("sales_quantity")) * Val(rssalesdtl.GetValue("rate")))) <> (Val(rssalesdtl.GetValue("basic_amount"))) Then
                                MsgBox("Wrong calculation, please EDIT/Update the invoice", MsgBoxStyle.Information, "eMPro")
                                Exit Sub
                            End If
                        Else
                            If System.Math.Round((Val(rssalesdtl.GetValue("sales_quantity")) * Val(rssalesdtl.GetValue("rate"))), intBasicRoundOffDecimal) <> (Val(rssalesdtl.GetValue("basic_amount"))) Then
                                MsgBox("Wrong Calculation, please EDIT/UPDATE the invoice", MsgBoxStyle.Information, "eMPro")
                                Exit Sub
                            End If
                        End If
                        rssalesdtl.MoveNext()
                        'UPGRADE_WARNING: Couldn't resolve default property of object rssaledtl.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Next

                End If
            End If
            '''rssaledtl = New ClsResultSetDB_Invoice
            Frm = New eMProCrystalReportViewer

            If UCase(Trim(GetPlantName)) = "HILEX" Then
                If mblnEwaybill_Print = False Then
                    Frm.glblnInvoiceform = True
                Else
                    If dblewaymaxvalue <= mdblewaymaximumvalue Then
                        Frm.glblnInvoiceform = True
                    ElseIf chkprintreprint.Checked = True Then
                        'strQry = "set dateformat 'dmy' SELECT TOP 1 1 FROM  Saleschallan_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND invoice_type='" & Trim(Me.lbldescription.Text) & "' and sub_category='" & Trim(Me.lblcategory.Text) & "' and Doc_No < 99000000 and bill_flag = '1' and cancel_flag = '0' and Location_Code='" & Trim(txtUnitCode.Text) & "' and doc_no=" & Ctlinvoice.Text.Trim
                        'strQry += " AND (( INVOICE_DATE < '" & mblnEWAY_BILL_STARTDATE & " '))"
                        'strQry += " UNION  SELECT TOP 1 1  FROM  SALESCHALLAN_DTL WHERE UNIT_CODE='" + gstrUNITID + "' AND INVOICE_TYPE='" & Trim(Me.lbldescription.Text) & "' AND SUB_CATEGORY='" & Trim(Me.lblcategory.Text) & "' AND DOC_NO < 99000000 AND BILL_FLAG = '1' AND CANCEL_FLAG = '0' AND LOCATION_CODE='" & Trim(txtUnitCode.Text) & "'"
                        'strQry += " AND TOTAL_AMOUNT <= " & mdblewaymaximumvalue & " and doc_no=" & Ctlinvoice.Text.Trim & ""
                        'strQry += " UNION  SELECT TOP 1 1 FROM  SALESCHALLAN_DTL WHERE UNIT_CODE='" + gstrUNITID + "' AND INVOICE_TYPE='" & Trim(Me.lbldescription.Text) & "' AND SUB_CATEGORY='" & Trim(Me.lblcategory.Text) & "' AND DOC_NO < 99000000 AND BILL_FLAG = '1' AND CANCEL_FLAG = '0' AND LOCATION_CODE='" & Trim(txtUnitCode.Text) & "'"
                        'strQry += " AND TOTAL_AMOUNT > " & mdblewaymaximumvalue & " AND  INVOICE_DATE >= '" & mblnEWAY_BILL_STARTDATE & " ' "
                        'strQry += " AND ISNULL(EWAY_BILL_NO,'')<>'' AND  INVOICE_DATE >= '" & mblnEWAY_BILL_STARTDATE & "'and doc_no=" & Ctlinvoice.Text.Trim

                        strQry = "set dateformat 'dmy' SELECT TOP 1 1 FROM  Saleschallan_dtl WHERE UNIT_CODE='" & gstrUNITID & "' AND invoice_type='" & Trim(Me.lbldescription.Text) & "' and sub_category='" & Trim(Me.lblcategory.Text) & "' and Doc_No < 99000000 and bill_flag = '1' and cancel_flag = '0' and Location_Code='" & Trim(txtUnitCode.Text) & "' and doc_no=" & Ctlinvoice.Text.Trim
                        strQry += " AND EWAY_IRN_REQUIRED='N' "
                        strQry += " UNION Select TOP 1 1 from SalesChallan_Dtl S LEFT JOIN SALESCHALLAN_DTL_IRN I ON I.UNIT_CODE=S.UNIT_CODE AND I.DOC_NO=S.DOC_NO where  S.UNIT_CODE = '" & gstrUNITID & "' and S.Invoice_Type <> 'EXP' "
                        strQry += " AND ((S.EWAY_IRN_REQUIRED='E' AND ISNULL(S.EWAY_BILL_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='I' AND ISNULL(I.IRN_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='B' AND ISNULL(S.EWAY_BILL_NO,'')<>'' AND ISNULL(I.IRN_NO,'')<>'')) "
                        strQry += " AND S.Location_Code='" & Trim(txtUnitCode.Text) & "'" & " and S.bill_flag =1 and S.CANCEL_FLAG = 0 and S.invoice_type='" & Me.lbldescription.Text & "' and S.sub_category='" & Me.lblcategory.Text & "' and S.Doc_No < 99000000 "
                        strQry += " AND S.Doc_No =" & Ctlinvoice.Text.Trim & " "
                    ElseIf chkprintreprint.Checked = False Then
                        'strQry = "set dateformat 'dmy' SELECT TOP 1 1 FROM  Saleschallan_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND invoice_type='" & Trim(Me.lbldescription.Text) & "' and sub_category='" & Trim(Me.lblcategory.Text) & "' and Doc_No < 99000000 and bill_flag = '1' and cancel_flag = '0' and Location_Code='" & Trim(txtUnitCode.Text) & "' and doc_no=" & Ctlinvoice.Text.Trim
                        'strQry += " AND (( INVOICE_DATE < '" & mblnEWAY_BILL_STARTDATE & " '))"
                        'strQry += " UNION  SELECT TOP 1 1  FROM  SALESCHALLAN_DTL WHERE UNIT_CODE='" + gstrUNITID + "' AND INVOICE_TYPE='" & Trim(Me.lbldescription.Text) & "' AND SUB_CATEGORY='" & Trim(Me.lblcategory.Text) & "' AND DOC_NO < 99000000 AND BILL_FLAG = '1' AND CANCEL_FLAG = '0' AND LOCATION_CODE='" & Trim(txtUnitCode.Text) & "'"
                        'strQry += " AND TOTAL_AMOUNT <= " & mdblewaymaximumvalue & " and doc_no=" & Ctlinvoice.Text.Trim & ""
                        'strQry += " UNION  SELECT TOP 1 1 FROM  SALESCHALLAN_DTL WHERE UNIT_CODE='" + gstrUNITID + "' AND INVOICE_TYPE='" & Trim(Me.lbldescription.Text) & "' AND SUB_CATEGORY='" & Trim(Me.lblcategory.Text) & "' AND DOC_NO < 99000000 AND BILL_FLAG = '1' AND CANCEL_FLAG = '0' AND LOCATION_CODE='" & Trim(txtUnitCode.Text) & "'"
                        'strQry += " AND TOTAL_AMOUNT > " & mdblewaymaximumvalue & " AND  INVOICE_DATE >= '" & mblnEWAY_BILL_STARTDATE & " ' "
                        'strQry += " AND ISNULL(EWAY_BILL_NO,'')<>'' AND  INVOICE_DATE >= '" & mblnEWAY_BILL_STARTDATE & "'and doc_no=" & Ctlinvoice.Text.Trim

                        strQry = "set dateformat 'dmy' Select S.Doc_No,S.Invoice_type from SalesChallan_Dtl S LEFT JOIN SALESCHALLAN_DTL_IRN I ON I.UNIT_CODE=S.UNIT_CODE AND I.DOC_NO=S.DOC_NO where  S.UNIT_CODE = '" & gstrUNITID & "' and S.Invoice_Type <> 'EXP' "
                        strQry += " AND ((S.EWAY_IRN_REQUIRED='E' AND ISNULL(S.EWAY_BILL_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='I' AND ISNULL(I.IRN_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='B' AND ISNULL(S.EWAY_BILL_NO,'')<>'' AND ISNULL(I.IRN_NO,'')<>'')) "
                        strQry += " AND S.Location_Code='" & Trim(txtUnitCode.Text) & "'" & " and S.bill_flag =1 and S.CANCEL_FLAG = 0 and S.invoice_type='" & Me.lbldescription.Text & "' and S.sub_category='" & Me.lblcategory.Text & "' and S.Doc_No < 99000000 "
                        strQry += " AND S.Doc_No = " & Ctlinvoice.Text.Trim & " "

                        If DataExist(strQry) = True Then
                            Frm.glblnInvoiceform = True
                        Else
                            Frm.glblnInvoiceform = False
                        End If
                    Else

                    End If
                End If

            Else
                Frm.glblnInvoiceform = False
            End If


            If e.Button = UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE Then
                Me.Dispose()
                Exit Sub
            Else
                If ValidSelection() = False Then Exit Sub
                '101188073 Start
                If gblnGSTUnit Then
                    If Len(GSTUnitPrefixCode) = 0 Or Val(GSTUnitPrefixCode) = 0 Then
                        'MsgBox("Please first define GST Unit Prefix in Gen_UnitMaster.", MsgBoxStyle.Information, "eMPro")
                        'Exit Sub
                    End If
                End If
                '101188073 End
                'Added for Issue ID 19992 Starts
                Call CheckMultipleSOAllowed(Trim(cmbInvType.Text), Trim(CmbCategory.Text))
                'Added for Issue ID 19992 Ends

            End If

            Dim rsgrin As New ADODB.Recordset
            Dim strSql As String
            Dim dblGRINQty As Double
            Dim dblAvailableQty As Double


            'If UCase(cmbInvType.Text) <> "SERVICE INVOICE"  and mblnServiceInvoiceWithoutSO Then
            If Not (UCase(cmbInvType.Text) = "SERVICE INVOICE" And mblnServiceInvoiceWithoutSO) Then
                'CODE ADDED BY NISHA ON 21/03/2003 FOR FINANCIAL ROLLOVER
                SALEDTL = "select * from Saleschallan_Dtl where UNIT_CODE= '" & gstrUNITID & "' and  Doc_No =" & Me.Ctlinvoice.Text & "  and Location_Code='" & Trim(txtUnitCode.Text) & "'"

                '10804443 - MULTI LOCATION IN BARCODE - HILEX 
                If optInvYes(0).Checked = True Then
                    If ValidateInvoiceStockLocation() = False Then
                        Exit Sub
                    End If
                End If
                ''
                rssaledtl = New ClsResultSetDB_Invoice
                rssaledtl.GetResult(SALEDTL, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                strAccountCode = rssaledtl.GetValue("Account_code")
                strCustRef = rssaledtl.GetValue("Cust_ref")
                StrAmendmentNo = rssaledtl.GetValue("Amendment_No")
                strInvoiceDate = VB6.Format(rssaledtl.GetValue("Invoice_Date"), "dd/mm/yyyy")
                mblncustomerlevel_A4report_functionlity = False
                mblncustomerspecificreport = False
                mblnftsitem = rssaledtl.GetValue("FTS_item")
                mblnftsbarcodeitem = rssaledtl.GetValue("FTS_BARCODE")

                If AllowA4Reports(strAccountCode) = True Then
                    mblncustomerlevel_A4report_functionlity = True
                End If

                'Added for Issue ID 21840 Starts
                blnInvoiceAgainstMultipleSO = rssaledtl.GetValue("InvoiceAgainstMultipleSO")
                rssaledtl.ResultSetClose()
                'Added for Issue ID 21840 Ends
                If UCase(cmbInvType.Text) = "NORMAL INVOICE" And UCase(CmbCategory.Text) = "FINISHED GOODS" Then
                    If AllowCustomerspecificreport(strAccountCode) = True Then
                        mblncustomerspecificreport = True
                    End If
                End If

                'Added for Issue ID eMpro-20080930-22159 Starts (new Column BatchTrackingAllowed)
                strSalesconf = "Select FTS_ENABLED,FTS_STOCK_LOCATION ,UpdatePO_Flag,UpdateStock_Flag,Stock_Location,OpenningBal,Preprinted_Flag,NoCopies,isnull(BatchTrackingAllowed,0) as BatchTrackingAllowed, AllowA4Reports , Noofcopies_A4report   from saleconf where UNIT_CODE= '" & gstrUNITID & "' and "
                strSalesconf = strSalesconf & "Invoice_type = '" & Me.lbldescription.Text & "' and sub_type = '"
                strSalesconf = strSalesconf & Me.lblcategory.Text & "' and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
                'Added for Issue ID eMpro-20080930-22159 Ends (new Column BatchTrackingAllowed)

                'CHANGES ENDS HERE 21/02/2003
                rsSalesConf = New ClsResultSetDB_Invoice
                rsSalesConf.GetResult(strSalesconf, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                updatePOflag = rsSalesConf.GetValue("UpdatePO_Flag")
                updatestockflag = rsSalesConf.GetValue("UpdateStock_Flag")
                strStockLocation = rsSalesConf.GetValue("Stock_Location")
                mOpeeningBalance = Val(rsSalesConf.GetValue("OpenningBal"))
                intNoCopies = rsSalesConf.GetValue("NoCopies")
                mblnA4reports_invoicewise = rsSalesConf.GetValue("AllowA4Reports")
                'Added for Issue ID eMpro-20080930-22159 Starts
                blnBatchTrack = rsSalesConf.GetValue("BatchTrackingAllowed")
                mblnftsenabled = rsSalesConf.GetValue("FTS_ENABLED")
                strFTSstocklocation = rsSalesConf.GetValue("FTS_Stock_Location")

                'Added for Issue ID eMpro-20080930-22159 Ends

                rsSalesConf.ResultSetClose()

                If Len(Trim(strStockLocation)) = 0 Then
                    MsgBox("Please Define Stock Location in Sales Configuration. ")
                    Exit Sub
                End If
                strsaleconfLocation = strStockLocation
                If mblnftsenabled = True And Len(Trim(strFTSstocklocation)) = 0 Then
                    MsgBox("Please Define FTS Stock Location in Sales Configuration. ")
                    Exit Sub
                End If

                If mblnftsenabled = True Then
                    'If mblnftsitem = True And mblnftsbarcodeitem = True Then
                    'strStockLocation = "01P3"
                    'ElseIf mblnftsitem = True And mblnftsbarcodeitem = False Then
                    '   strStockLocation = strFTSstocklocation
                    'End If
                    'Else
                    strStockLocation = Find_Value("Select fts_location from saleschallan_dtl  WHERE UNIT_CODE='" + gstrUNITID + "' and doc_no='" & Ctlinvoice.Text & "'")
                End If

                'priti madam changes revert as on 12 aug 2021
                Dim isRejectionInvoice As Boolean = Find_Value("Select isnull(is_Rejection_invoice,0) from saleschallan_dtl  WHERE UNIT_CODE='" + gstrUNITID + "' and doc_no='" & Ctlinvoice.Text & "'")
                If isRejectionInvoice Then
                    strStockLocation = Find_Value("Select Transfer_RejectionLoc from sales_parameter WHERE UNIT_CODE='" + gstrUNITID + "'")
                    strsaleconfLocation = strStockLocation
                End If
                'priti madam changes ended revert as on 12 aug 2021

                If mblnftsenabled = True Then
                    If mblnftsitem = True And mblnftsbarcodeitem = True Then
                        'FTS RELATED CHANGES
                        If DataExist("Select top 1 1  from saleschallan_dtl WHERE UNIT_CODE='" + gstrUNITID + "' and doc_no='" & Ctlinvoice.Text & "' and fts_item=1 and fts_barcode=1") Then
                            oCmd = New ADODB.Command
                            With oCmd
                                .let_ActiveConnection(mP_Connection)
                                .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                                .CommandText = "USP_FTS_SCANINVOICE"
                                .Parameters.Append(.CreateParameter("@Unit_Code", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                                .Parameters.Append(.CreateParameter("@Doc_No", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, , Trim(Ctlinvoice.Text)))
                                .Parameters.Append(.CreateParameter("@IPADDRESS", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, gstrIpaddressWinSck))
                                .Parameters.Append(.CreateParameter("@ERRMSG", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamOutput, 5000))

                                .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            End With
                            If oCmd.Parameters(oCmd.Parameters.Count - 1).Value <> "" Then

                                MsgBox(oCmd.Parameters(oCmd.Parameters.Count - 1).Value.ToString(), MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100))
                                oCmd = Nothing
                                Exit Sub
                            End If
                            oCmd = Nothing

                        End If
                        'FTS RELATED CHANGES


                    End If

                End If
                'Changed for Issue ID 19992 Starts
                If mblnMultipleSOAllowed = False Then
                    '***********To check if Tool Cost Deduction will be done or Not on 16/02/2004
                    rsSalesParameter = New ClsResultSetDB_Invoice
                    rsSalesParameter.GetResult("Select Batch_Tracking = Isnull(Batch_Tracking,0),CheckToolAmortisation from Sales_Parameter where UNIT_CODE= '" & gstrUNITID & "'")
                    If rsSalesParameter.GetNoRows > 0 Then
                        rsSalesParameter.MoveFirst()
                        'UPGRADE_WARNING: Couldn't resolve default property of object rsSalesParameter.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'

                        'Commented for Issue ID eMpro-20080930-22159 Starts
                        'blnBatchTrack = rsSalesParameter.GetValue("Batch_Tracking")
                        'Commented for Issue ID eMpro-20080930-22159 Ends

                        'UPGRADE_WARNING: Couldn't resolve default property of object rsSalesParameter.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If Len(Trim(rsSalesParameter.GetValue("CheckToolAmortisation"))) = 0 Then
                            MsgBox("First define Check Tool Amortisation in Sales Parameter", MsgBoxStyle.Information, "eMPro")
                            Exit Sub
                        End If
                        'UPGRADE_WARNING: Couldn't resolve default property of object rsSalesParameter.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        blnCheckToolCost = rsSalesParameter.GetValue("CheckToolAmortisation")
                    Else
                        MsgBox("No Data Defined in Sales Parameter", MsgBoxStyle.Information, "eMPro")
                        rsSalesParameter.ResultSetClose()
                        Exit Sub
                    End If
                    rsSalesParameter.ResultSetClose()
                    '*************
                    SALEDTL = "Select Sales_Quantity,Item_code,Cust_Item_Code,toolcost_amount from sales_Dtl where UNIT_CODE= '" & gstrUNITID & "' and  Doc_No = " & Me.Ctlinvoice.Text & " and Location_Code='" & Trim(txtUnitCode.Text) & "'"
                    rssaledtl = New ClsResultSetDB_Invoice
                    rssaledtl.GetResult(SALEDTL, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                    intRow = rssaledtl.GetNoRows
                    rssaledtl.MoveFirst()

                    If optInvYes(0).Checked = True Then
                        '******Check for balance & despatch in Cust_ord_dtl
                        For intLoopCount = 1 To intRow
                            'UPGRADE_WARNING: Couldn't resolve default property of object rssaledtl.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            ItemCode = rssaledtl.GetValue("Item_code")
                            'UPGRADE_WARNING: Couldn't resolve default property of object rssaledtl.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            salesQuantity = Val(rssaledtl.GetValue("Sales_quantity"))
                            'UPGRADE_WARNING: Couldn't resolve default property of object rssaledtl.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            strDrgNo = rssaledtl.GetValue("Cust_Item_code")
                            'UPGRADE_WARNING: Couldn't resolve default property of object rssaledtl.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            'Changed for Issue ID eMpro-20090415-30143 Starts
                            ' dblToolCost = Val(rssaledtl.GetValue("ToolCost_amount"))
                            dblToolCost = IIf(IsDBNull(rssaledtl.GetValue("ToolCost_amount")), 0, Val(rssaledtl.GetValue("ToolCost_amount")))
                            rsItembal = New ClsResultSetDB_Invoice
                            rsItembal.GetResult("Select Cur_bal from Itembal_Mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  Item_code = '" & ItemCode & "'and Location_code ='" & strStockLocation & "'", ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                            If rsItembal.GetNoRows > 0 Then
                                'UPGRADE_WARNING: Couldn't resolve default property of object rsItembal.GetValue(Cur_Bal). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                If salesQuantity > rsItembal.GetValue("Cur_Bal") Then
                                    MsgBox("Balance for item " & ItemCode & " at Location " & strStockLocation & " not available. ", MsgBoxStyle.Information, "eMPro")
                                    rsItembal.ResultSetClose()
                                    Exit Sub
                                End If
                            Else
                                MsgBox("No Item in ItemMaster for Location " & strStockLocation & ".", MsgBoxStyle.OkOnly, "eMPro")
                                '''@@@rsSalesConf.ResultSetClose()
                                'UPGRADE_NOTE: Object rsSalesConf may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                                '''@@@rsSalesConf = Nothing
                                rsItembal.ResultSetClose()
                                Exit Sub
                            End If
                            'Code add by Sourabh Khatri For Batch Tracking
                            '---------------------------------------------
                            If blnBatchTrack = True And updatestockflag = True And UCase(Trim(cmbInvType.Text)) <> "REJECTION" And UCase(Trim(cmbInvType.Text)) <> "JOBWORK INVOICE" Then
                                rsBatch = New ClsResultSetDB_Invoice
                                Call rsBatch.GetResult("Select Batch_No,Batch_Qty from ItemBatch_Dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  Doc_Type = 9999 and Doc_no = '" & Trim(Me.Ctlinvoice.Text) & "' and Item_Code = '" & ItemCode & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                                If rsBatch.RowCount <= 0 Then
                                    MsgBox(" Batch Details is Not Available ", MsgBoxStyle.Information, "eMPro")
                                    Exit Sub
                                End If
                                rsBatch.MoveFirst()
                                While Not rsBatch.EOFRecord
                                    'UPGRADE_WARNING: Couldn't resolve default property of object rsBatch.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                    rsBatchMst = New ClsResultSetDB_Invoice
                                    Call rsBatchMst.GetResult("Select Current_batch_Qty = Isnull(Current_batch_Qty,0) From ItemBatch_Mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  Batch_No = '" & rsBatch.GetValue("Batch_No") & "' and Location_Code = '" & strStockLocation & "' and Item_Code = '" & Trim(ItemCode) & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                                    If rsBatchMst.RowCount > 0 Then
                                        'UPGRADE_WARNING: Couldn't resolve default property of object rsBatch.GetValue(Batch_Qty). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                        'UPGRADE_WARNING: Couldn't resolve default property of object rsBatchMst.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                        If Val(rsBatchMst.GetValue("Current_Batch_Qty")) < rsBatch.GetValue("Batch_Qty") Then
                                            'UPGRADE_WARNING: Couldn't resolve default property of object rsBatchMst.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                            MsgBox("Balance for item " & ItemCode & " at Location " & strStockLocation & " is " & Val(rsBatchMst.GetValue("Current_Batch_Qty")) & " at Batch Master")
                                            rsBatch.ResultSetClose()
                                            rsBatchMst.ResultSetClose()
                                            Exit Sub
                                        Else
                                            'UPGRADE_WARNING: Couldn't resolve default property of object rsBatch.GetValue(Batch_No). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                            'UPGRADE_WARNING: Couldn't resolve default property of object rsBatch.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                            strBatchQuery = strBatchQuery & "  Update ItemBatch_Mst Set Current_batch_Qty = Current_batch_Qty - " & Val(rsBatch.GetValue("Batch_Qty")) & ",Upd_Userid = '" & mP_User & "' ,Upd_Dt = getdate()  WHERE UNIT_CODE='" + gstrUNITID + "' AND  Batch_No = '" & rsBatch.GetValue("Batch_No") & "' and Location_Code = '" & strStockLocation & "' and Item_Code = '" & Trim(ItemCode) & "'"
                                        End If
                                        rsBatchMst.ResultSetClose()
                                    Else
                                        MsgBox("Balance for item " & ItemCode & " at Location " & strStockLocation & " not available in Batch Master. ")
                                        rsBatch.ResultSetClose()
                                        rsBatchMst.ResultSetClose()
                                        Exit Sub
                                    End If
                                    rsBatch.MoveNext()
                                End While
                                rsBatch.ResultSetClose()
                            End If
                            '---------------------------------------------


                            If Len(Trim(strCustRef)) > 0 Then
                                If UCase(cmbInvType.Text) <> "REJECTION" Then
                                    rsItembal = New ClsResultSetDB_Invoice
                                    rsItembal.GetResult("Select balanceQty = order_qty - despatch_Qty,OpenSO from Cust_ord_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  account_code ='" & strAccountCode & "' and Cust_ref ='" & strCustRef & "' and Amendment_No = '" & StrAmendmentNo & "' and Item_code ='" & ItemCode & "' and Cust_drgNo ='" & strDrgNo & "' and Active_flag ='A'", ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                                    If rsItembal.GetNoRows > 0 Then
                                        'Changed by nisha on 15/09/2002 for Open So Check
                                        'UPGRADE_WARNING: Couldn't resolve default property of object rsItembal.GetValue(OpenSO). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                        If rsItembal.GetValue("OpenSO") = False Then
                                            'UPGRADE_WARNING: Couldn't resolve default property of object rsItembal.GetValue(BalanceQty). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                            If salesQuantity > rsItembal.GetValue("BalanceQty") Then
                                                'UPGRADE_WARNING: Couldn't resolve default property of object rsItembal.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                                MsgBox("Balance Quantity in SO for item " & ItemCode & " is " & rsItembal.GetValue("BalanceQty") & ".Check Quantity of Item in Challan.", MsgBoxStyle.Information, "eMPro")
                                                rsItembal.ResultSetClose()
                                                Exit Sub
                                            End If
                                        End If
                                        rsItembal.ResultSetClose()
                                    Else
                                        MsgBox("No Item (" & StrItemCode & ") exist in SO - " & strCustRef & ".", MsgBoxStyle.Information, "eMPro")
                                        rsItembal.ResultSetClose()
                                        Exit Sub
                                    End If
                                End If
                            End If
                            '************To Check for Tool Cost
                            If blnCheckToolCost = True Then
                                strItembal = "select BalanceQty = isnull(a.proj_qty,0) - isnull(a.ClosingValueSMIEL,0),a.Tool_C from Amor_dtl a,Tool_Mst b"
                                strItembal = strItembal & " WHERE a.unit_code = b.Unit_code and  a.UNIT_CODE='" + gstrUNITID + "' AND  account_code = '" & strAccountCode & "'"
                                strItembal = strItembal & " and Item_code = '" & ItemCode & "' and a.Tool_c = b.tool_c and a.Item_code = b.Product_No order by a.tool_c"
                                rsItembal = New ClsResultSetDB_Invoice
                                rsItembal.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                                If rsItembal.GetNoRows > 0 Then
                                    rsItembal.MoveFirst()
                                    'UPGRADE_WARNING: Couldn't resolve default property of object rsItembal.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                    strtoolQuantity = CStr(Val(rsItembal.GetValue("BalanceQty")))
                                    'code added by nisha on 22 Nov for tool code check
                                    'UPGRADE_WARNING: Couldn't resolve default property of object rsItembal.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                    strToolCode = rsItembal.GetValue("Tool_c")
                                    rsItembal.ResultSetClose()

                                    strItembal = "select BalanceQty = sum(isnull(UsedProjQty,0)) from Amor_dtl "
                                    strItembal = strItembal & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  "
                                    strItembal = strItembal & " Item_code = '" & ItemCode & "' and Tool_c = '" & strToolCode & "'"

                                    rsItembal = New ClsResultSetDB_Invoice
                                    rsItembal.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                                    rsItembal.MoveFirst()
                                    'UPGRADE_WARNING: Couldn't resolve default property of object rsItembal.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                    strtoolQuantity = CStr(Val(strtoolQuantity) - Val(rsItembal.GetValue("BalanceQty")))
                                    rsItembal.ResultSetClose()
                                    'Chenges ends here by nisha on 22 Nov 2004
                                    If Val(CStr(salesQuantity)) > Val(strtoolQuantity) Then
                                        If Val(strtoolQuantity) = 0 Then
                                            MsgBox("No Balance Available for Item (" & ItemCode & ") and customer Part Code (" & strDrgNo & ") For Amortisation Calculations. ", MsgBoxStyle.OkOnly, "eMPro")
                                        Else
                                            MsgBox("Quantity should not be Greater then available Balance Quantity for Amortisarion " & strtoolQuantity, MsgBoxStyle.OkOnly, "eMPro")
                                        End If
                                        Exit Sub
                                    End If
                                Else
                                    rsItembal.ResultSetClose()
                                    'code added by nisha on 22 Nov for tool code check
                                    strItembal = "select BalanceQty = isnull(proj_qty,0) - isnull(ClosingValueSMIEL,0) from Amor_dtl"
                                    strItembal = strItembal & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  account_code = '" & strAccountCode & "'"
                                    strItembal = strItembal & " and Item_code = '" & ItemCode & "'"

                                    rsItembal = New ClsResultSetDB_Invoice
                                    rsItembal.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                                    If rsItembal.GetNoRows > 0 Then
                                        rsItembal.MoveFirst()
                                        'UPGRADE_WARNING: Couldn't resolve default property of object rsItembal.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                        strtoolQuantity = CStr(Val(rsItembal.GetValue("BalanceQty")))
                                        rsItembal.ResultSetClose()

                                        strItembal = "select BalanceQty = sum(isnull(UsedProjQty,0)) from Amor_dtl"
                                        strItembal = strItembal & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Item_code = '" & ItemCode & "'"
                                        rsItembal = New ClsResultSetDB_Invoice
                                        rsItembal.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                                        'UPGRADE_WARNING: Couldn't resolve default property of object rsItembal.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                        strtoolQuantity = CStr(CDbl(Val(strtoolQuantity)) - Val(rsItembal.GetValue("BalanceQty")))
                                        rsItembal.ResultSetClose()
                                        If Val(CStr(salesQuantity)) > Val(strtoolQuantity) Then
                                            If Val(strtoolQuantity) = 0 Then
                                                MsgBox("No Balance Available for Item (" & ItemCode & ") and customer Part Code (" & strDrgNo & ") For Amortisation Calculations. ", MsgBoxStyle.OkOnly, "eMPro")
                                            Else
                                                MsgBox("Quantity should not be Greater then available Balance Quantity for Amortisarion " & strtoolQuantity, MsgBoxStyle.OkOnly, "eMPro")
                                            End If
                                            Exit Sub
                                        End If
                                    End If
                                    'Chenges ends here by nisha on 22 Nov 2004
                                End If
                                '************Add Rajani Kant 19/08/2004
                                With mP_Connection
                                    .Execute("DELETE FROM TMPBOM WHERE UNIT_CODE='" + gstrUNITID + "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                    'Added By ekta uniyal
                                    '.Execute("BOMExplosion '" & Trim(ItemCode) & "','" & Trim(ItemCode) & "',1,0", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                    .Execute("BOMExplosion_Hilex'" & Trim(ItemCode) & "','" & Trim(ItemCode) & "',1,0,0,1,'" & gstrIpaddressWinSck & "','" + gstrUNITID + "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                    'End Here
                                End With

                                rsbom = New ClsResultSetDB_Invoice

                                rsbom.GetResult("select * from tmpBOM  WHERE UNIT_CODE='" + gstrUNITID + "' ", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                                If rsbom.GetNoRows > 0 Then
                                    irowcount = rsbom.GetNoRows
                                    rsbom.MoveFirst()
                                    For intRwCount1 = 1 To irowcount
                                        strItembal = "select BalanceQty = isnull(a.proj_qty,0) - isnull(a.ClosingValueSMIEL,0),a.tool_C from Amor_dtl a, tool_mst b "
                                        strItembal = strItembal & " WHERE a.Unit_code = b.Unit_code AND  a.UNIT_CODE='" + gstrUNITID + "' AND  account_code = '" & Trim(strAccountCode) & "'"
                                        'UPGRADE_WARNING: Couldn't resolve default property of object rsbom.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                        strItembal = strItembal & " and Item_code = '" & rsbom.GetValue("item_code") & "' and a.Tool_c = b.Tool_c and a.ITem_code = b.Product_no order by a.tool_c"
                                        rsItembal = New ClsResultSetDB_Invoice
                                        rsItembal.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                                        If rsItembal.GetNoRows > 0 Then
                                            rsItembal.MoveFirst()
                                            'UPGRADE_WARNING: Couldn't resolve default property of object rsItembal.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                            strtoolQuantity = CStr(Val(rsItembal.GetValue("BalanceQty")))
                                            'UPGRADE_WARNING: Couldn't resolve default property of object rsItembal.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                            strToolCode = rsItembal.GetValue("Tool_c")
                                            rsItembal.ResultSetClose()

                                            'code added by nisha on 22 Nov for tool code check
                                            strItembal = "select BalanceQty = sum(isnull(UsedProjQty,0)) from Amor_dtl a"
                                            strItembal = strItembal & " WHERE a.UNIT_CODE='" + gstrUNITID + "' AND  account_code = '" & Trim(strAccountCode) & "'"
                                            'UPGRADE_WARNING: Couldn't resolve default property of object rsbom.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                            strItembal = strItembal & " and Item_code = '" & rsbom.GetValue("item_code") & "' and a.Tool_c = '" & strToolCode & "'"

                                            rsItembal = New ClsResultSetDB_Invoice
                                            rsItembal.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                                            'UPGRADE_WARNING: Couldn't resolve default property of object rsbom.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                            varItemQty1 = (salesQuantity * Val(rsbom.GetValue("grossweight")))
                                            'UPGRADE_WARNING: Couldn't resolve default property of object rsItembal.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                            strtoolQuantity = CStr(Val(strtoolQuantity) - Val(rsItembal.GetValue("BalanceQty")))
                                            rsItembal.ResultSetClose()

                                            'changes end here by nisha on 22 Nov 2004
                                            If Val(CStr(varItemQty1)) > Val(strtoolQuantity) Then
                                                If Val(strtoolQuantity) = 0 Then
                                                    'UPGRADE_WARNING: Couldn't resolve default property of object rsbom.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                                    MsgBox("No Balance Available for Item (" & rsbom.GetValue("item_code") & ") and customer Part Code (" & strDrgNo & ") For Amortisation Calculations. ", MsgBoxStyle.OkOnly, "eMPro")
                                                Else
                                                    'UPGRADE_WARNING: Couldn't resolve default property of object rsbom.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                                    MsgBox("Quantity should not be Greater then available Balance Quantity for Amortisarion of this Item (" & rsbom.GetValue("item_code") & ")" & strtoolQuantity, MsgBoxStyle.OkOnly, "eMPro")
                                                End If
                                                Exit Sub
                                            End If
                                        End If
                                        rsbom.MoveNext()
                                    Next
                                End If
                                rsbom.ResultSetClose()
                            End If
                            rssaledtl.MoveNext()
                        Next

                        '****
                        '****To Check in Rejection Invoice if Grin No Exist
                        If UCase(cmbInvType.Text) = "REJECTION" Then
                            If Len(Trim(strCustRef)) > 0 Then
                                If CheckDataFromGrin(Val(Trim(strCustRef)), strAccountCode) = False Then
                                    Exit Sub
                                End If
                            End If
                        End If
                        '****
                    End If
                    rssaledtl.ResultSetClose()
                    'UPGRADE_NOTE: Object rssaledtl may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                    '''rssaledtl = Nothing


                ElseIf mblnMultipleSOAllowed = True And Val(Ctlinvoice.Text) > 99000000 Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object CheckBalanceForPrinting(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    strBalanceCheck = CheckBalanceForPrinting()
                    If Len(Trim(strBalanceCheck)) > 0 Then
                        If strBalanceCheck = "Error" Then
                            Exit Sub

                        Else
                            MsgBox(Space(20) & "Balance Status:" & vbCrLf & strBalanceCheck, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))

                        End If
                        Exit Sub
                    End If
                    SALEDTL = "Select COUNT(*) as introw from sales_Dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  Doc_No = " & Me.Ctlinvoice.Text & " and Location_Code='" & Trim(txtUnitCode.Text) & "'"
                    rssaledtl = New ClsResultSetDB_Invoice
                    rssaledtl.GetResult(SALEDTL, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                    If rssaledtl.GetNoRows > 0 Then
                        'UPGRADE_WARNING: Couldn't resolve default property of object rssaledtl.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        intRow = rssaledtl.GetValue("introw")
                    End If
                    rssaledtl.ResultSetClose()
                End If
                'Changed for Issue ID 19992 Ends

            Else
                If optInvYes(0).Checked Then
                    'Validation Added By Arshad in Service Invoice

                    strSql = "SELECT NRGP_NO, GRIN_NO, ITEM_CODE, ITEM_QTY FROM NRGP_GRIN_Dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  "
                    strSql = strSql & " NRGP_NO IN (SELECT NRGPNoInCaseOfServiceInvoice FROM SALESCHALLAN_DTL WHERE UNIT_CODE='" + gstrUNITID + "' AND  DOC_NO=" & Trim(Ctlinvoice.Text) & " and Location_code='" & txtUnitCode.Text & "')"
                    If rsgrin.State = ADODB.ObjectStateEnum.adStateOpen Then
                        rsgrin.Close()
                        'UPGRADE_NOTE: Object rsgrin may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                        rsgrin = Nothing
                    End If
                    rsgrin.Open(strSql, mP_Connection)
                    While Not rsgrin.EOF
                        dblGRINQty = rsgrin.Fields("Item_Qty").Value

                        strSql = "SELECT D.ACCEPTED_QUANTITY-D.DESPATCH_QUANTITY AS QUANTITY "
                        strSql = strSql & " FROM GRN_HDR H INNER JOIN GRN_DTL D"
                        strSql = strSql & " ON H.DOC_TYPE = D.DOC_TYPE AND H.DOC_NO = D.DOC_NO AND H.FROM_LOCATION = D.FROM_LOCATION AND H.UNIT_CODE=D.UNIT_CODE  "
                        strSql = strSql & " WHERE H.UNIT_CODE='" + gstrUNITID + " AND H.DOC_CATEGORY='Z' AND H.QA_AUTHORIZED_CODE IS NOT NULL"
                        strSql = strSql & " AND H.DOC_NO='" & rsgrin.Fields("Grin_No").Value & "'"
                        strSql = strSql & " AND D.ITEM_CODE='" & rsgrin.Fields("ITEM_CODE").Value & "'"

                        dblAvailableQty = Val(Find_Value(strSql))

                        If dblGRINQty > dblAvailableQty Then
                            MsgBox("GRIN availbale quantity is less than NRGP quantity.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "eMPro")
                            Exit Sub
                        End If
                        rsgrin.MoveNext()
                    End While
                    If rsgrin.State = ADODB.ObjectStateEnum.adStateOpen Then
                        rsgrin.Close()
                        'UPGRADE_NOTE: Object rsgrin may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                        rsgrin = Nothing
                    End If
                End If
                'Ends here
            End If

            Dim intLoopCounter As Short
            Dim blnPrintFlag As Boolean
            Dim blnPrintActualInvFlag As Boolean
            Dim intNoCopies_A4reports_orignial As Short
            Dim intNoCopies_A4reports_REPRINT As Short
            Dim rsGENERATEBARCODE As ClsResultSetDB_Invoice
            Dim strPrintMethod As String = ""
            Dim ObjBarcodeHMI As New Prj_BCHMI.cls_BCHMI(gstrUNITID)
            Dim intTotalNoofitemsinInvoices As Integer

            'Dim ObjBarcodeHMI As New cls_BCHMI(gstrUNITID)
            Dim strBarcodeMsg As String
            strBarcodeMsg = ""

            Select Case e.Button
                ' crRpt.Destination = crptToWindow
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                    Me.Dispose()
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT_TO_WINDOW
                    '                Write_In_Log_File(GetServerDateTime() & "  : Process Started - Window")
                    '102027599
                    If optInvYes(0).Checked = False Then
                        If mblnEwaybill_Print Then
                            Call IRN_QRBarcode()
                        End If
                    End If
                    If InvoiceGeneration(RdAddSold, RepPath, Frm) = True Then

                        'Changed and Part Added By Arshad on 23/04/2004 for Dos Based Printing
                        If CBool(Find_Value("select TextPrinting from sales_parameter WHERE UNIT_CODE='" + gstrUNITID + "'")) Then
                            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.WaitCursor)
                            objInvoicePrint = New prj_InvoicePrinting.clsInvoicePrinting(gstrDateFormat)
                            'Added for Issue ID eMpro-20080805-20745 Starts
                            objInvoicePrint.mstrDSNforInoivcePrint = gstrDSNName
                            'Added for Issue ID eMpro-20080805-20745 Ends
                            objInvoicePrint.ConnectionString = gstrCONNECTIONSTRING  'mP_Connection.ConnectionString
                            objInvoicePrint.Connection()
                            'objInvoicePrint.FileName = App.Path & "\Reports\InvoicePrint.txt"
                            'shalini
                            'objInvoicePrint.FileName = "C:\InvoicePrint.txt"
                            objInvoicePrint.FileName = strCitrix_Inv_Pronting_Loc & "InvoicePrint.txt"
                            'objInvoicePrint.BCFileName = App.Path & "\BarCode.txt"
                            'shalini
                            'objInvoicePrint.BCFileName = "C:\BarCode.txt"
                            objInvoicePrint.BCFileName = strCitrix_Inv_Pronting_Loc & "BarCode.txt"
                            objInvoicePrint.CompanyName = gstrCOMPANY
                            objInvoicePrint.Address1 = gstr_RGN_ADDRESS1
                            objInvoicePrint.Address2 = gstr_RGN_ADDRESS2
                            objInvoicePrint.Print_Invoice(gstrUNITID, True, (txtUnitCode.Text), (Ctlinvoice.Text), dtpRemoval.Text & " " & dtpRemovalTime.Value.Hour & ":" & dtpRemovalTime.Value.Minute)
                            '                        Write_In_Log_File(GetServerDateTime() & " : Text File Generated Successfully")
                            '''rtbInvoicePreview.LoadFile(objInvoicePrint.FileName)
                            rtbInvoicePreview.LoadFile(objInvoicePrint.FileName, RichTextBoxStreamType.PlainText)
                            rtbInvoicePreview.BackColor = System.Drawing.Color.White
                            cmdPrint.Image = My.Resources.ico231.ToBitmap
                            cmdPrint.ImageAlign = ContentAlignment.BottomCenter
                            cmdPrint.TextAlign = ContentAlignment.TopCenter

                            cmdClose.Image = My.Resources.ico217.ToBitmap
                            cmdClose.ImageAlign = ContentAlignment.TopCenter
                            cmdClose.TextAlign = ContentAlignment.TopCenter

                            FraInvoicePreview.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Height) - 1300)
                            FraInvoicePreview.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Width) - 400)
                            FraInvoicePreview.Left = VB6.TwipsToPixelsX(100)
                            FraInvoicePreview.Top = ctlFormHeader1.Height

                            rtbInvoicePreview.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(FraInvoicePreview.Height) - 1000)
                            rtbInvoicePreview.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(FraInvoicePreview.Width) - 200)
                            rtbInvoicePreview.Left = VB6.TwipsToPixelsX(100)
                            rtbInvoicePreview.Top = VB6.TwipsToPixelsY(900)
                            rtbInvoicePreview.RightMargin = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(rtbInvoicePreview.Width) + 5000)

                            shpInvoice.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(FraInvoicePreview.Width) - VB6.PixelsToTwipsX(shpInvoice.Width)) / 2)
                            cmdPrint.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(shpInvoice.Left) + 100)
                            cmdClose.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(cmdPrint.Left) + VB6.PixelsToTwipsX(cmdPrint.Width) + 100)

                            If DataExist("SELECT TOP 1 1 FROM SALES_PARAMETER WHERE PRINT_WITHOUT_LOCK_REQ = 1 and UNIT_CODE='" + gstrUNITID + "'") Then
                                cmdPrint.Enabled = True
                            Else
                                cmdPrint.Enabled = False
                            End If

                            cmdClose.Enabled = True
                            FraInvoicePreview.Enabled = True : rtbInvoicePreview.Enabled = True : rtbInvoicePreview.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                            ReplaceJunkCharacters()
                            FraInvoicePreview.Visible = True
                            FraInvoicePreview.Enabled = True
                            FraInvoicePreview.BringToFront()
                            rtbInvoicePreview.Focus()
                            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
                        Else
                            'Start 102027599
                            If AllowBarCodePrinting(strAccountCode) = True Then
                                If optInvYes(0).Checked = False And chkprintreprint.Checked = False And mblnEwaybill_Print = True Then
                                    '------------------------------------------------------------------------------------
                                    rsGENERATEBARCODE = New ClsResultSetDB_Invoice
                                    rsGENERATEBARCODE.GetResult("SELECT PRINT_METHOD FROM CUSTOMER_MST C WHERE C.UNIT_CODE='" & gstrUNITID & "' AND C.CUSTOMER_CODE='" & strAccountCode & "'")
                                    strPrintMethod = UCase(rsGENERATEBARCODE.GetValue("PRINT_METHOD").ToString)
                                    rsGENERATEBARCODE.ResultSetClose()
                                    rsGENERATEBARCODE = Nothing

                                    If optInvYes(0).Checked = False And chkprintreprint.Checked = False And mblnEwaybill_Print = True Then
                                        If strPrintMethod = "NORMAL" Then
                                            If blnlinelevelcustomer = True Then
                                                strBarcodeMsg = ObjBarcodeHMI.GenerateBarCode_LINELEVEL_SALESORDER_2dbarcode_Normal_Hilex(gstrUserMyDocPath, Trim(Ctlinvoice.Text), "NORMAL", "", "", True, Trim(Ctlinvoice.Text), gstrCONNECTIONSTRING)

                                            Else
                                                strBarcodeMsg = ObjBarcodeHMI.GenerateBarCode_LINELEVEL_SALESORDER_2dbarcode_hilex(gstrUserMyDocPath, Trim(Ctlinvoice.Text), "NORMAL", "", "", True, Trim(Ctlinvoice.Text), gstrCONNECTIONSTRING)
                                            End If

                                            If VB.Left(strBarcodeMsg, 1) <> "Y" Then
                                                MsgBox("Problem While Generating Barcode Image.", vbInformation, ResolveResString(100))
                                                Exit Sub
                                            Else
                                                If SaveBarCodeImage_singlelevelso_2DBARCODE(Ctlinvoice.Text, gstrUserMyDocPath, Mid(strBarcodeMsg, 3)) = False Then
                                                    MsgBox("Problem While saving Barcode Image.", vbInformation, ResolveResString(100))
                                                    Exit Sub
                                                Else
                                                    mP_Connection.Execute(" UPDATE T SET T.BARCODEIMAGE =SC.BARCODEIMAGE FROM SALESCHALLAN_DTL SC,TMP_INVOICEPRINT  T " &
                                                                           " WHERE SC.UNIT_CODE = T.UNIT_CODE AND SC.DOC_NO =T.DOC_NO AND SC.UNIT_CODE='" & gstrUNITID & "' AND " &
                                                                           " SC.DOC_NO='" & Ctlinvoice.Text.Trim & "' AND T.IP_ADDRESS='" & gstrIpaddressWinSck & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                                End If
                                            End If
                                        End If

                                    End If
                                End If
                            End If

                            'End 102027599
                            Call ReprintQRbarcode()
                            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                            ''  rptinvoice.Destination = Crystal.k.crptToWindow

                            'Added for Issue ID eMpro-20090415-30143 Starts
                            CheckASNExist(Me.Ctlinvoice.Text)
                            'HILEX A4 BARCODE'

                            If AllowBarCodePrinting(strAccountCode) = True Then
                                If optInvYes(0).Checked = True Then
                                    '------------------------------------------------------------------------------------
                                    rsGENERATEBARCODE = New ClsResultSetDB_Invoice
                                    rsGENERATEBARCODE.GetResult("SELECT PRINT_METHOD FROM CUSTOMER_MST C WHERE C.UNIT_CODE='" & gstrUNITID & "' AND C.CUSTOMER_CODE='" & strAccountCode & "'")
                                    strPrintMethod = UCase(rsGENERATEBARCODE.GetValue("PRINT_METHOD").ToString)
                                    rsGENERATEBARCODE.ResultSetClose()
                                    rsGENERATEBARCODE = Nothing

                                    If optInvYes(0).Checked = True Then
                                        If strPrintMethod = "NORMAL" Then
                                            If blnlinelevelcustomer = True Then
                                                strBarcodeMsg = ObjBarcodeHMI.GenerateBarCode_LINELEVEL_SALESORDER_2dbarcode_Normal_Hilex(gstrUserMyDocPath, mInvNo, "NORMAL", "", "", True, Trim(Ctlinvoice.Text), gstrCONNECTIONSTRING)
                                            Else
                                                strBarcodeMsg = ObjBarcodeHMI.GenerateBarCode_LINELEVEL_SALESORDER_2dbarcode_hilex(gstrUserMyDocPath, mInvNo, "NORMAL", "", "", True, Trim(Ctlinvoice.Text), gstrCONNECTIONSTRING)
                                            End If
                                            Dim strQuery As String

                                            If VB.Left(strBarcodeMsg, 1) <> "Y" Then
                                                MsgBox("Problem While Generating Barcode Image.", vbInformation, ResolveResString(100))
                                                Exit Sub
                                            Else
                                                If SaveBarCodeImage_singlelevelso_2DBARCODE(Ctlinvoice.Text, gstrUserMyDocPath, Mid(strBarcodeMsg, 3)) = False Then
                                                    MsgBox("Problem While saving Barcode Image.", vbInformation, ResolveResString(100))
                                                    Exit Sub
                                                Else
                                                    mP_Connection.Execute(" UPDATE T SET T.BARCODEIMAGE =SC.BARCODEIMAGE FROM SALESCHALLAN_DTL SC,TMP_INVOICEPRINT  T " &
                                                                           " WHERE SC.UNIT_CODE = T.UNIT_CODE AND SC.DOC_NO =T.DOC_NO AND SC.UNIT_CODE='" & gstrUNITID & "' AND " &
                                                                           " SC.DOC_NO='" & Ctlinvoice.Text.Trim & "' AND T.IP_ADDRESS='" & gstrIpaddressWinSck & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                                End If
                                            End If
                                            '25 OCT 2017
                                        ElseIf strPrintMethod = "TOYOTA_NEW" Then
                                            strBarcodeMsg = ObjBarcodeHMI.GenerateBarCode_LINELEVEL_SALESORDER_2dbarcode_hilex(gstrUserMyDocPath, mInvNo, "TOYOTA_NEW", "", "", True, Trim(Ctlinvoice.Text), gstrCONNECTIONSTRING)
                                            Dim strQuery As String

                                            If VB.Left(strBarcodeMsg, 1) <> "Y" Then
                                                MsgBox("Problem While Generating Barcode Image.", vbInformation, ResolveResString(100))
                                                Exit Sub
                                            Else
                                                If SaveBarCodeImage_singlelevelso_2DBARCODE(Ctlinvoice.Text, gstrUserMyDocPath, Mid(strBarcodeMsg, 3)) = False Then
                                                    MsgBox("Problem While saving Barcode Image.", vbInformation, ResolveResString(100))
                                                    Exit Sub
                                                Else
                                                    mP_Connection.Execute(" UPDATE T SET T.BARCODEIMAGE =SC.BARCODEIMAGE FROM SALESCHALLAN_DTL SC,TMP_INVOICEPRINT  T " &
                                                                                   " WHERE SC.UNIT_CODE = T.UNIT_CODE AND SC.DOC_NO =T.DOC_NO AND SC.UNIT_CODE='" & gstrUNITID & "' AND " &
                                                                                   " SC.DOC_NO='" & Ctlinvoice.Text.Trim & "' AND T.IP_ADDRESS='" & gstrIpaddressWinSck & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                                End If
                                            End If
                                            '25 OCT 2017


                                            'Added by prit for TATA QR BARCODE on 15 Feb 2021
                                            'If AllowBarCodePrinting(strAccountCode) Then 
                                        ElseIf GetPrintMethod(strAccountCode).ToUpper() = "TATA" Then  '' For Temporary Number
                                            Dim StrTATAsuffix As String
                                            StrTATAsuffix = gstrUNITID & Ctlinvoice.Text.ToString.Trim & DateTime.Now.ToString("ddMMyyHHmmssfff")
                                            strBarcodeMsg = ObjBarcodeHMI.GenerateQRBarCodeForTATAMotors(gstrUserMyDocPath, True, Trim(Ctlinvoice.Text), 0, gstrCONNECTIONSTRING, StrTATAsuffix)
                                            If VB.Left(strBarcodeMsg, 1) <> "Y" Then
                                                CustomRollbackTrans()
                                                MsgBox("Problem While Generating Barcode Image.", vbInformation, ResolveResString(100))
                                                Exit Sub
                                            Else
                                                strBarcodeMsg_paratemeter = Mid(strBarcodeMsg, 3)
                                                If Not SaveQRBarCodeImageTATA(Trim(Ctlinvoice.Text), 0, strBarcodeMsg_paratemeter, StrTATAsuffix) Then
                                                    CustomRollbackTrans()
                                                    MsgBox("Problem While Saving Barcode Image.", vbInformation, ResolveResString(100))
                                                    Exit Sub
                                                End If
                                            End If


                                        Else
                                            rsGENERATEBARCODE = New ClsResultSetDB_Invoice
                                            rsGENERATEBARCODE.GetResult("SELECT PRINT_METHOD,ITEM_CODE,CUST_ITEM_CODE FROM SALES_DTL SD ,SALESCHALLAN_DTL SC ,CUSTOMER_MST C WHERE C.UNIT_CODE=SC.UNIT_CODE AND C.CUSTOMER_CODE=SC.ACCOUNT_CODE AND SC.UNIT_CODE=SD.UNIT_CODE AND " &
                                                                                " SC.DOC_NO=SD.DOC_NO  AND SC.UNIT_CODE='" & gstrUNITID & "' AND  " &
                                                                                " SC.DOC_NO= " & Ctlinvoice.Text & " ORDER BY CUST_ITEM_CODE ")

                                            intTotalNoofitemsinInvoices = rsGENERATEBARCODE.GetNoRows
                                            rsGENERATEBARCODE.MoveFirst()
                                            For intRow = 1 To intTotalNoofitemsinInvoices
                                                strBarcodeMsg = ObjBarcodeHMI.GenerateBarCode_LINELEVEL_SALESORDER(gstrUserMyDocPath, mInvNo, rsGENERATEBARCODE.GetValue("PRINT_METHOD").ToString, rsGENERATEBARCODE.GetValue("Cust_Item_Code").ToString, rsGENERATEBARCODE.GetValue("Item_Code").ToString, True, Trim(Ctlinvoice.Text), gstrCONNECTIONSTRING)
                                                If VB.Left(strBarcodeMsg, 1) <> "Y" Then
                                                    MsgBox("Problem While Generating Barcode Image.", vbInformation, ResolveResString(100))
                                                    Exit Sub
                                                End If
                                                If SaveBarCodeImage_singlelevelso(Ctlinvoice.Text, rsGENERATEBARCODE.GetValue("Cust_Item_Code").ToString, rsGENERATEBARCODE.GetValue("Item_Code").ToString, gstrUserMyDocPath, intRow) = False Then
                                                    MsgBox("Problem While saving Barcode Image.", vbInformation, ResolveResString(100))
                                                    Exit Sub
                                                End If
                                                rsGENERATEBARCODE.MoveNext()
                                            Next

                                            rsGENERATEBARCODE.ResultSetClose()
                                            rsGENERATEBARCODE = Nothing
                                        End If

                                    End If
                                Else
                                    'Added by prit for TATA QR BARCODE on 15 Feb 2021
                                    'If AllowBarCodePrinting(strAccountCode) Then
                                    If GetPrintMethod(strAccountCode).ToUpper() = "TATA" Then '' Reprint  + preview option
                                        Dim StrTATAsuffix As String
                                        StrTATAsuffix = gstrUNITID & Ctlinvoice.Text.ToString.Trim & DateTime.Now.ToString("ddMMyyHHmmssfff")
                                        strBarcodeMsg = ObjBarcodeHMI.GenerateQRBarCodeForTATAMotors(gstrUserMyDocPath, True, Trim(Ctlinvoice.Text), 0, gstrCONNECTIONSTRING, StrTATAsuffix)
                                        If VB.Left(strBarcodeMsg, 1) <> "Y" Then
                                            CustomRollbackTrans()
                                            MsgBox("Problem While Generating Barcode Image.", vbInformation, ResolveResString(100))
                                            Exit Sub
                                        Else
                                            strBarcodeMsg_paratemeter = Mid(strBarcodeMsg, 3)
                                            If Not SaveQRBarCodeImageTATA(Trim(Ctlinvoice.Text), mInvNo, strBarcodeMsg_paratemeter, StrTATAsuffix) Then
                                                CustomRollbackTrans()
                                                MsgBox("Problem While Saving Barcode Image.", vbInformation, ResolveResString(100))
                                                Exit Sub
                                            End If
                                        End If
                                    End If
                                End If

                            End If

                            'Changes for Hero invoice Start  -------------------PRAVEEN KUMAR

                            If (Find_Value("SELECT top 1 1 FROM LISTS WHERE UNIT_CODE='" & gstrUNITID & "' AND KEY1='HEROBAR' and Key2='" & strAccountCode & "' ")) = "1" Then
                                Dim strBarcodestring1 As String = ""
                                Dim strBarcodestring2 As String = ""
                                Dim strBarcodeString3 As String = ""
                                Dim totalAccesibleAmt As Double = 0
                                Dim totalCgst As Double = 0
                                Dim totalIgst As Double = 0
                                Dim totalSgst As Double = 0
                                Dim totaltcs As Double = 0

                                rsGENERATEBARCODE = New ClsResultSetDB_Invoice

                                ' rsGENERATEBARCODE.GetResult("SELECT  C.CUST_VENDOR_CODE,A.CUST_REF,A.DOC_NO,Convert(varchar,A.INVOICE_DATE,104) INVOICE_DATE,GSTIN_ID, Convert(numeric(8,2),ACCESSIBLE_AMOUNT) ACCESSIBLE_AMOUNT, Convert(numeric(8,2),C.TOTAL_AMOUNT) TOTAL_AMOUNT,A.VEHICLE_NO,Convert(numeric(8,2),isnull(SGSTTXRT_AMOUNT,0)) SGSTTXRT_AMOUNT, " & _
                                '" Convert(numeric(8,2),isnull(IGSTTXRT_AMOUNT,0)) IGSTTXRT_AMOUNT, Convert(numeric(8,2),isnull(CGSTTXRT_AMOUNT,0)) CGSTTXRT_AMOUNT,CUST_item_CODE,HSNSACCODE,Convert(numeric(8,2),SALES_QUANTITY) SALES_QUANTITY,Convert(numeric(8,2),Rate) Rate " & _
                                '" FROM SALESCHALLAN_DTL A INNER JOIN GEN_UNITMASTER B ON A.UNIT_CODE=B.UNT_CODEID " & _
                                '" INNER JOIN TMP_INVOICEPRINT C ON B.UNT_CODEID=C.UNIT_CODE AND A.DOC_NO=C.DOC_NO " & _
                                '" WHERE A.DOC_NO='" & Ctlinvoice.Text.Trim & "' AND A.UNIT_CODE='" & gstrUNITID & "' ")

                                rsGENERATEBARCODE.GetResult("SELECT  C.CUST_VENDOR_CODE,A.CUST_REF,A.DOC_NO,Convert(varchar,A.INVOICE_DATE,104) INVOICE_DATE,GSTIN_ID, Convert(numeric(19,2),ACCESSIBLE_AMOUNT) ACCESSIBLE_AMOUNT, Convert(numeric(19,2),C.TOTAL_AMOUNT) TOTAL_AMOUNT,A.VEHICLE_NO,Convert(numeric(19,2),isnull(SGSTTXRT_AMOUNT,0)) SGSTTXRT_AMOUNT, " &
                                " Convert(numeric(19,2),isnull(IGSTTXRT_AMOUNT,0)) IGSTTXRT_AMOUNT, Convert(numeric(19,2),isnull(CGSTTXRT_AMOUNT,0)) CGSTTXRT_AMOUNT,CUST_item_CODE,HSNSACCODE,Convert(numeric(12,2),SALES_QUANTITY) SALES_QUANTITY,Convert(numeric(19,2),Rate) Rate ,Convert(numeric(19,2),isnull(TCSAMOUNT,0)) TCSAMOUNT " &
                                " FROM SALESCHALLAN_DTL A INNER JOIN GEN_UNITMASTER B ON A.UNIT_CODE=B.UNT_CODEID " &
                                " INNER JOIN TMP_INVOICEPRINT C ON B.UNT_CODEID=C.UNIT_CODE AND A.DOC_NO=C.DOC_NO " &
                                " WHERE A.DOC_NO='" & Ctlinvoice.Text.Trim & "' AND A.UNIT_CODE='" & gstrUNITID & "' AND C.IP_ADDRESS= '" & gstrIpaddressWinSck & "'")


                                While Not rsGENERATEBARCODE.EOFRecord
                                    totalAccesibleAmt = totalAccesibleAmt + Convert.ToDouble(rsGENERATEBARCODE.GetValue("ACCESSIBLE_AMOUNT").ToString)
                                    If optInvYes(0).Checked = True Then
                                        strBarcodestring1 = rsGENERATEBARCODE.GetValue("CUST_VENDOR_CODE").ToString & vbTab & rsGENERATEBARCODE.GetValue("CUST_REF").ToString & vbTab & mInvNo.ToString & vbTab & rsGENERATEBARCODE.GetValue("INVOICE_DATE").ToString & vbTab & rsGENERATEBARCODE.GetValue("GSTIN_ID").ToString & vbTab & rsGENERATEBARCODE.GetValue("TOTAL_AMOUNT").ToString
                                    Else
                                        strBarcodestring1 = rsGENERATEBARCODE.GetValue("CUST_VENDOR_CODE").ToString & vbTab & rsGENERATEBARCODE.GetValue("CUST_REF").ToString & vbTab & Ctlinvoice.Text.ToString & vbTab & rsGENERATEBARCODE.GetValue("INVOICE_DATE").ToString & vbTab & rsGENERATEBARCODE.GetValue("GSTIN_ID").ToString & vbTab & rsGENERATEBARCODE.GetValue("TOTAL_AMOUNT").ToString
                                    End If

                                    strBarcodeString3 = rsGENERATEBARCODE.GetValue("VEHICLE_NO").ToString
                                    totalSgst = totalSgst + Convert.ToDouble(rsGENERATEBARCODE.GetValue("SGSTTXRT_AMOUNT").ToString)
                                    totalIgst = totalIgst + Convert.ToDouble(rsGENERATEBARCODE.GetValue("IGSTTXRT_AMOUNT").ToString)
                                    totalCgst = totalCgst + Convert.ToDouble(rsGENERATEBARCODE.GetValue("CGSTTXRT_AMOUNT").ToString)
                                    totaltcs = Convert.ToDouble(rsGENERATEBARCODE.GetValue("TCSAMOUNT").ToString)
                                    strBarcodestring2 = strBarcodestring2 & vbTab & rsGENERATEBARCODE.GetValue("CUST_item_CODE").ToString & vbTab & rsGENERATEBARCODE.GetValue("HSNSACCODE").ToString & vbTab & rsGENERATEBARCODE.GetValue("SALES_QUANTITY").ToString & vbTab & rsGENERATEBARCODE.GetValue("Rate").ToString
                                    rsGENERATEBARCODE.MoveNext()
                                End While
                                rsGENERATEBARCODE.ResultSetClose()


                                Dim PDF417barcode As BarcodeLib.Barcode.PDF417.PDF417 = New BarcodeLib.Barcode.PDF417.PDF417()
                                PDF417barcode.UOM = BarcodeLib.Barcode.Linear.UnitOfMeasure.Pixel
                                PDF417barcode.LeftMargin = 0
                                PDF417barcode.RightMargin = 0
                                PDF417barcode.TopMargin = 0
                                PDF417barcode.BottomMargin = 0
                                PDF417barcode.ImageFormat = System.Drawing.Imaging.ImageFormat.Png

                                PDF417barcode.Data = (strBarcodestring1 & vbTab & totalAccesibleAmt.ToString & vbTab & strBarcodeString3 & vbTab & totalSgst.ToString & vbTab & totalIgst.ToString & vbTab & totalCgst.ToString & vbTab & totaltcs.ToString + strBarcodestring2).ToString().Trim
                                Dim imageData() As Byte = PDF417barcode.drawBarcodeAsBytes()

                                Dim cmd As SqlCommand = Nothing
                                cmd = New System.Data.SqlClient.SqlCommand()
                                With cmd
                                    .CommandType = CommandType.Text
                                    .CommandText = "UPDATE TMP_INVOICEPRINT SET barcodeimage=@QRIMAGE where DOC_NO='" & Ctlinvoice.Text.Trim & "' AND UNIT_CODE='" & gstrUNITID & "' "
                                    .Parameters.Add("@QRIMAGE", SqlDbType.Image).Value = imageData
                                    SqlConnectionclass.ExecuteNonQuery(cmd)

                                End With
                            End If



                            'Changes for Hero invoice End-------------------

                            'HILEX A4 BARCODE
                            If optInvYes(0).Checked = True And AllowASNPrinting(strAccountCode) = True Then
                                If mblnASNExist = True Then
                                    mP_Connection.Execute("Update CreatedASN Set ASN_NO='" & Trim$(txtASNNumber.Text) & "',Updatedon=getdate() WHERE UNIT_CODE='" + gstrUNITID + "' AND  doc_no='" & Trim$(Me.Ctlinvoice.Text) & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                Else
                                    If txtASNNumber.Text.Trim.Length > 0 Then
                                        mP_Connection.Execute("Insert into CreatedASN values('" & Trim$(Me.Ctlinvoice.Text) & "','" & Trim$(txtASNNumber.Text) & "',getdate(),'" & mP_User & "',getdate(),'" + gstrUNITID + "','" & dtpASNDatetime.Value & "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                    End If
                                End If
                            End If
                            'Added for Issue ID eMpro-20090415-30143 Ends

                            If DataExist("SELECT TOP 1 1 FROM SALES_PARAMETER WHERE PRINT_WITHOUT_LOCK_REQ = 1 and  UNIT_CODE='" + gstrUNITID + "'") Then
                                blnIsReportDisplayed = True
                                Frm.Show()

                                ''  intmaxsubreportloop = Me.rptinvoice.GetNSubreports
                                ''  For intsubreportloopcounter = 0 To intmaxsubreportloop - 1
                                ''      Me.rptinvoice.SubreportToChange = Me.rptinvoice.GetNthSubreportName(intsubreportloopcounter)
                                ''     Me.rptinvoice.Connect = gstrREPORTCONNECT


                                ''   Next
                                ''   Me.rptinvoice.SubreportToChange = ""
                                ''   rptinvoice.WindowAllowDrillDown = True
                                ''  rptinvoice.Connect = gstrREPORTCONNECT
                                ''   rptinvoice.Action = 1
                            End If

                            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                        End If
                    Else
                        Exit Sub
                    End If

                    'SATISH KESHARWANI CHANGE
                    If chktkmlbarcode.Checked = True And optInvYes(1).Checked = True Then
                        Print_barcodelabel(Me.Ctlinvoice.Text)
                    End If

                    'SATISH KESHARWANI CHANGE
                    If optInvYes(0).Checked = True And AllowASNPrinting(strAccountCode) = True Then
                        If txtASNNumber.Text.Trim.Length > 0 Then
                            'If DUPLICATEASN(Me.txtASNNumber.Text.Trim) = True Then
                            strAsnInvoice = DUPLICATEASN(Me.txtASNNumber.Text.Trim)
                            If mblnDuplicateASNExist = True Then
                                MsgBox("Already Used ASN Number in Different Invoice. Invoice No:" & strAsnInvoice & " can't save", MsgBoxStyle.Information, ResolveResString(100))
                                mP_Connection.Execute(" Update CreatedASN Set ASN_NO='' where doc_no='" & Trim$(Me.Ctlinvoice.Text) & "' and unit_code='" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                txtASNNumber.Text = ""
                                Exit Sub
                            End If
                        End If
                    End If

                    'Changes By Arshad Ends Here
                    '************Added By Tapan On 20-Aug-2k2*************
                    If DataExist("SELECT TOP 1 1 FROM SALES_DTL WHERE DOC_NO='" & Me.Ctlinvoice.Text.Trim & "' And Rate = 0 ") Then
                        MsgBox("Invoice Rate Cannot be zero , Kindly EDIT/UPDATE ", MsgBoxStyle.Information, ResolveResString(100))
                        Exit Sub
                    End If

                    'PRASHANT RAJPAL CHANGED DATED 13 MAR 2014'
                    If optInvYes(0).Checked = True And InvAgstBarCode() = True And (mstrFGDomestic = "1" Or mstrFGDomestic = "2") Then
                        If BarCodeTracking(Trim(Ctlinvoice.Text), "LOCK") = False Then
                            If mblnQuantityCheck = False Then
                                Exit Sub
                            End If
                        End If
                    End If
                    ''PRASHANT RAJPAL CHANGED ENDED 13 MAR 2014'



                    If chkLockPrintingFlag.CheckState = 1 And optInvYes(0).Checked = True Then
                        Sleep((3000))
                        If ConfirmWindow(10344, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                            Dim strtime As String = GetServerDateTime()
                            blninvoicelockYES = True
                            If CBool(Find_Value("select TextPrinting from sales_parameter WHERE UNIT_CODE='" + gstrUNITID + "'")) Then
                                cmdPrint.Enabled = False      'ENABLE PRINT BUTTON
                            Else
                                If Not blnIsReportDisplayed Then
                                    If UCase(Trim(GetPlantName)) = "SMIEL" Or (UCase(Trim(GetPlantName)) = "SUMIT") Then

                                        Frm.ShowExportButton = False
                                        '' rptinvoice.WindowShowPrintSetupBtn = False
                                        Frm.ShowPrintButton = False
                                    End If


                                    If UCase(Trim(GetPlantName)) <> "SMIEL" And UCase(Trim(GetPlantName)) <> "SUMIT" Then
                                        '10109115 
                                        '' intmaxsubreportloop = Me.rptinvoice.GetNSubreports
                                        ''  For intsubreportloopcounter = 0 To intmaxsubreportloop - 1
                                        ''    Me.rptinvoice.SubreportToChange = Me.rptinvoice.GetNthSubreportName(intsubreportloopcounter)
                                        ''    Me.rptinvoice.Connect = gstrREPORTCONNECT


                                        ''   Next
                                        ''   Me.rptinvoice.SubreportToChange = ""
                                        ''   rptinvoice.WindowAllowDrillDown = True
                                        ''   rptinvoice.Connect = gstrREPORTCONNECT
                                        '10109115  end 
                                        If UCase(Trim(GetPlantName)) = "VF1" Or UCase(Trim(GetPlantName)) = "RSA" Then
                                            If mstrReportFilename = "rptinvoicemate_VF1" Or mstrReportFilename = "rptinvoiceMATE_RSA" Then
                                                RdAddSold.DataDefinition.FormulaFields("Deliverynoteno").Text = "'" + CStr(mInvNo) + "'"
                                            End If
                                        End If
                                        '' rptinvoice.Action = 1
                                        'Frm.Show()
                                    End If

                                    'rptinvoice.Action = 1       'PRINT INVOICE REPORT
                                End If
                            End If

                            'Added for Issue ID 22207 Starts
                            If Val(Trim(CStr(mInvNo))) = 0 Then
                                MsgBox("Invoice Can't save with number 0. Please Check Current number in Sales Configuration master. ", MsgBoxStyle.Information, ResolveResString(100))
                                Exit Sub
                            End If
                            'Added for Issue ID 22207 Ends
                            'prashant changed on 19 apr 2011 
                            If DataExist("select total_amount from SalesChallan_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  location_code='" & Trim(txtUnitCode.Text) & "' and doc_no='" & mInvNo & "' And total_amount = 0 ") Then
                                MsgBox("Kindly EDIT/UPDATE the Invoice Again", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "eMPro")
                                Exit Sub
                            End If
                            'prashant changed on 19 apr 2011
                            'prashant changed on 06 May 2011 for issue id : 1093232
                            If DataExist("select doc_no from Supplementaryinv_hdr WHERE UNIT_CODE='" + gstrUNITID + "' AND  location_code='" & Trim(txtUnitCode.Text) & "' and doc_no='" & mInvNo & "'") Then
                                MsgBox("Already Exist with the same Number , Please Try Again  ", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "eMPro")
                                Exit Sub
                            End If
                            'prashant changed ended on 06 May 2011 for issue id : 1093232
                            'Code Added By Arshad on 28/04/2004
                            If Len(Find_Value("select doc_no from SalesChallan_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  location_code='" & Trim(txtUnitCode.Text) & "' and doc_no='" & mInvNo & "'")) > 0 Then
                                MsgBox("Next Invoice number already generated." & vbCrLf & "Please skip current no either backward or forward" & vbCrLf & "in Sales Configuration Master Form.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "eMPro")
                                Exit Sub
                            End If
                            'changes By Arshad end here

                            ' Add by Sandeep
                            mstrInvRejSQL = ""
                            If CBool(Find_Value("Select REJINV_Tracking from Sales_Parameter WHERE UNIT_CODE='" + gstrUNITID + "'")) = True Then
                                mstrInvRejSQL = "Update MKT_INVREJ_DTL Set Invoice_No='" & mInvNo & "' WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_No='" & Trim(Ctlinvoice.Text) & "'"
                            End If
                            ' 'End Here

                            'Added for Issue ID eMpro-20090226-27911 Starts
                            Call CheckInvoiceExistInFinance(mInvNo)
                            'mstrInvRejSQL = ""
                            'If CBool(Find_Value("Select REJINV_Tracking from Sales_Parameter WHERE UNIT_CODE='" + gstrUNITID + "'")) = True Then
                            '    mstrInvRejSQL = "Update MKT_INVREJ_DTL Set Invoice_No='" & mInvNo & "' WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_No='" & Trim(Ctlinvoice.Text) & "'"
                            'End If
                            'Added for Issue ID eMpro-20090226-27911 Ends
                            ResetDatabaseConnection()
                            mP_Connection.BeginTrans()
                            mP_Connection.Execute("set Dateformat 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

                            'Added by prit for TATA QR BARCODE on 15 Feb 2021
                            If AllowBarCodePrinting(strAccountCode) Then
                                If GetPrintMethod(strAccountCode).ToUpper() = "TATA" Then  ' For Locking at Preview button.
                                    Dim StrTATAsuffix As String
                                    StrTATAsuffix = gstrUNITID & mInvNo.ToString.Trim & DateTime.Now.ToString("ddMMyyHHmmssfff")

                                    strBarcodeMsg = ObjBarcodeHMI.GenerateQRBarCodeForTATAMotors(gstrUserMyDocPath, True, Trim(Ctlinvoice.Text), mInvNo, gstrCONNECTIONSTRING, StrTATAsuffix)
                                    If VB.Left(strBarcodeMsg, 1) <> "Y" Then
                                        CustomRollbackTrans()
                                        MsgBox("Problem While Generating Barcode Image.", vbInformation, ResolveResString(100))
                                        Exit Sub
                                    Else
                                        strBarcodeMsg_paratemeter = Mid(strBarcodeMsg, 3)
                                        If Not SaveQRBarCodeImageTATA(Trim(Ctlinvoice.Text), mInvNo, strBarcodeMsg_paratemeter, StrTATAsuffix) Then
                                            CustomRollbackTrans()
                                            MsgBox("Problem While Saving Barcode Image.", vbInformation, ResolveResString(100))
                                            Exit Sub
                                        End If
                                    End If
                                End If
                            End If

                            TempInvNo = Ctlinvoice.Text
                            'Added for Issue ID 21105 Starts
                            'issue id 10192547 
                            'If InvAgstBarCode() = True And mstrFGDomestic = "1" Then
                            'FTS RELATED CHANGES
                            If DataExist("Select top 1 1  from saleschallan_dtl WHERE UNIT_CODE='" + gstrUNITID + "' and doc_no='" & Ctlinvoice.Text & "' and fts_item=1 and isnull(is_Rejection_invoice,0)=0 ") Then
                                oCmd = New ADODB.Command
                                With oCmd
                                    .let_ActiveConnection(mP_Connection)
                                    .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                                    .CommandText = "USP_FTS_PRODUCTIONSLIP"
                                    .Parameters.Append(.CreateParameter("@Unit_Code", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                                    .Parameters.Append(.CreateParameter("@Doc_No", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, , Trim(Ctlinvoice.Text)))
                                    .Parameters.Append(.CreateParameter("@ERRMSG", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamOutput, 1000))
                                    .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                End With
                                If oCmd.Parameters(oCmd.Parameters.Count - 1).Value <> "" Then
                                    CustomRollbackTrans()
                                    MsgBox(oCmd.Parameters(oCmd.Parameters.Count - 1).Value, MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100))
                                    oCmd = Nothing
                                    Exit Sub
                                End If
                                oCmd = Nothing
                                mP_Connection.Execute("update Saleschallan_dtl set from_location=  '" & strsaleconfLocation & "' WHERE UNIT_CODE='" + gstrUNITID + "' AND  Doc_no = " & Ctlinvoice.Text, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                mP_Connection.Execute("update FTS_LABEL_ISSUE set AUTHORISED= 1 ,DOC_NO='" & mInvNo & "' WHERE UNIT_CODE='" + gstrUNITID + "' AND DOC_TYPE=9999  AND  Doc_no = '" & Ctlinvoice.Text & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                mP_Connection.Execute("update FTS_LABEL_RELATIONSHIP set Ref_DOC_NO='" & mInvNo & "' WHERE UNIT_CODE='" + gstrUNITID + "' AND Ref_Doc_no = '" & Ctlinvoice.Text & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                mP_Connection.Execute("update FTS_FG_PICKLIST_INV_DTL set INVOICENO='" & mInvNo & "' WHERE UNIT_CODE='" + gstrUNITID + "' AND TMP_INVOICENO = '" & Ctlinvoice.Text & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                mP_Connection.Execute("update PRODRECEIPT_DTL set INVOICE_NO='" & mInvNo & "' WHERE UNIT_CODE='" + gstrUNITID + "' AND INVOICE_NO = '" & Ctlinvoice.Text & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

                            Else
                                mP_Connection.Execute("update Saleschallan_dtl set from_location=  '" & strsaleconfLocation & "' WHERE UNIT_CODE='" + gstrUNITID + "' AND  Doc_no = " & Ctlinvoice.Text, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            End If
                            'FTS RELATED CHANGES

                            'If InvAgstBarCode() = True And (mstrFGDomestic = "1" Or mstrFGDomestic = "2") Then
                            '    'issue id 10192547 end 
                            '    If BarCodeTracking(Trim(Ctlinvoice.Text), "LOCK") = True Then
                            '        mP_Connection.Execute(mstrupdateBarBondedStockQty, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            '        mP_Connection.Execute(mstrupdateBarBondedStockFlag, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            '    ElseIf mblnQuantityCheck = False Then
                            '        CustomRollbackTrans
                            '        Exit Sub
                            '    End If
                            'End If
                            'Added for Issue ID 21105 Starts
                            If InvAgstBarCode() = True And (mstrFGDomestic = "1" Or mstrFGDomestic = "2") Then
                                'issue id 10192547 end 
                                If BarCodeTracking(Trim(Ctlinvoice.Text), "LOCK") = True Then
                                    If UCase(Trim(CmbCategory.Text)) = "RAW MATERIAL" Or UCase(Trim(CmbCategory.Text)) = "INPUTS" Or UCase(Trim(CmbCategory.Text)) = "COMPONENTS" Or UCase(Trim(CmbCategory.Text)) = "SUB ASSEMBLY" Then
                                        Call updateBARcrossrefence_Invoicequantity(Ctlinvoice.Text)
                                    Else
                                        If mstrupdateBarBondedStockQty <> "" Then
                                            mP_Connection.Execute(mstrupdateBarBondedStockQty, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                        End If
                                        If mstrupdateBarBondedStockFlag <> "" Then
                                            mP_Connection.Execute(mstrupdateBarBondedStockFlag, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                        End If
                                    End If

                                    'mP_Connection.Execute(mstrupdateBarBondedStockQty, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                    'mP_Connection.Execute(mstrupdateBarBondedStockFlag, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                ElseIf mblnQuantityCheck = False Then
                                    CustomRollbackTrans()
                                    Exit Sub
                                End If

                            End If

                            ' Code add by Sourabh
                            If Trim(UCase(Me.txtUnitCode.Text)) = "SML" Or Trim(UCase(Me.txtUnitCode.Text)) = "SMT" Then
                                If CBool(Find_Value("Select isnull(MultiUnitInvoice,0) from sales_parameter WHERE UNIT_CODE='" + gstrUNITID + "'")) Then
                                    mP_Connection.Execute("Set XACT_ABORT  on", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                    blnFlagTrans = True
                                    With cmdObject
                                        .let_ActiveConnection(mP_Connection)
                                        .CommandTimeout = 0
                                        .CommandType = ADODB.CommandTypeEnum.adCmdText
                                        If Me.txtUnitCode.Text = "SML" Then
                                            strVar = Replace(salesconf, "saleconf", "SumitLive2008.DBO.Saleconf")
                                            strVar = Replace(strVar, "Location_Code='SML'", "Location_Code='SMT'")
                                        Else
                                            strVar = Replace(salesconf, "saleconf", "SMIEL_FIN2008.DBO.Saleconf")
                                            strVar = Replace(strVar, "Location_Code='SMT'", "Location_Code='SML'")
                                        End If
                                        .CommandText = strVar
                                        .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                    End With
                                    'UPGRADE_NOTE: Object cmdObject may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                                    cmdObject = Nothing
                                    mP_Connection.Execute("Set XACT_ABORT  Off", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                    blnFlagTrans = False
                                End If
                            End If
                            'Code end here

                            mP_Connection.Execute(salesconf, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

                            ' Add by Sandeep To Update Form_Details.
                            mP_Connection.Execute("UpDate Forms_Dtl Set PO_No='" & mInvNo & "' where  UNIT_CODE='" + gstrUNITID + "' AND  PO_No='" & Trim(Ctlinvoice.Text) & "' and Doc_Type='9999'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            ' End
                            If Len(Trim(mstrExcisePriorityUpdationString)) > 0 Then
                                mP_Connection.Execute("update Saleschallan_dtl set Excise_type = '" & mstrExcisePriorityUpdationString & "' WHERE UNIT_CODE='" + gstrUNITID + "' AND  Doc_no = " & Ctlinvoice.Text, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            End If
                            mP_Connection.Execute("INSERT INTO INV_ERROR_DTL(QUERY,UNIT_CODE,INVOICENO) VALUES('" & Replace(saleschallan, "'", "") & "','" & gstrUNITID & "','" & mInvNo & "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            mP_Connection.Execute(saleschallan, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            'hilex
                            Call UPDATETRANFERINVOICE_HILEX(Ctlinvoice.Text, mInvNo, "L")
                            'hilex


                            ' Add by Sandeep Chadha
                            If Len(Trim(mstrInvRejSQL)) <> 0 Then
                                mP_Connection.Execute(mstrInvRejSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            End If
                            'End

                            If updatePOflag = True And Len(strupdatecustodtdtl) > 0 Then
                                mP_Connection.Execute(strupdatecustodtdtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            End If

                            '''***** changes done By ashutosh on 13-06-2006, issue Id:18099
                            If updatestockflag = True Then
                                'UPGRADE_WARNING: Couldn't resolve default property of object varTmp1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                varTmp1 = Split(strSelectItmbalmst, "»")

                                varTmp = Split(strupdateitbalmst, "»")

                                For i = 0 To (intRow - 1)
                                    'UPGRADE_WARNING: Couldn't resolve default property of object varTmp1(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                    rsItemBalance = New ClsResultSetDB_Invoice
                                    rsItemBalance.GetResult(varTmp1(i))
                                    If rsItemBalance.GetNoRows > 0 Then
                                        'UPGRADE_WARNING: Couldn't resolve default property of object rsItemBalance.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                        dblTmpItembal = Val(rsItemBalance.GetValue("Cur_Bal"))
                                    End If
                                    rsItemBalance.ResultSetClose()
                                    If varTmp(i) <> "" Then
                                        mP_Connection.Execute(varTmp(i), , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                    End If

                                    'UPGRADE_WARNING: Couldn't resolve default property of object varTmp1(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                    rsItemBalance = New ClsResultSetDB_Invoice
                                    rsItemBalance.GetResult(varTmp1(i))
                                    If rsItemBalance.GetNoRows > 0 Then
                                        'UPGRADE_WARNING: Couldn't resolve default property of object rsItemBalance.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                        dblFinalItembal = Val(rsItemBalance.GetValue("Cur_Bal"))
                                    End If
                                    rsItemBalance.ResultSetClose()

                                    ''''                    If dblFinalItembal = dblTmpItembal Then
                                    ''''                        mP_Connection.Execute varTmp(i)
                                    ''''                    End If
                                Next i
                                '''mP_Connection.Execute strupdateitbalmst
                            End If
                            '''rsItemBalance.ResultSetClose()
                            'UPGRADE_NOTE: Object rsItemBalance may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                            '''rsItemBalance = Nothing
                            '''***** Cahnges for Issue Id:18099 end here.

                            'Code Added By Arshad Ali
                            If UCase(cmbInvType.Text) = "SERVICE INVOICE" And mblnServiceInvoiceWithoutSO Then
                                If Len(mstrGrinQtyUpdate) > 0 Then
                                    mP_Connection.Execute(mstrGrinQtyUpdate, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                End If
                            End If
                            'Ends here
                            'Code add By Sourabh Khatri If Batch Tracking is On Then Query Will Update ItemBatch_mst
                            If Len(Trim(strBatchQuery)) > 0 Then
                                mP_Connection.Execute(strBatchQuery, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                mP_Connection.Execute("Update ItemBatch_dtl Set Doc_no = '" & Trim(CStr(mInvNo)) & "' WHERE UNIT_CODE='" + gstrUNITID + "' AND  Doc_no = '" & Trim(Me.Ctlinvoice.Text) & "' and Doc_Type = 9999 ", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            End If
                            ' Code add by sourabh on 22 nov 2005 against issue id 2004-04-003-16261
                            If mblnInvoiceAgainstBarCode = True Then
                                mP_Connection.Execute("Update tbl_BarCode_MaterialOut Set Invoice_No = '" & Trim(CStr(mInvNo)) & "' WHERE UNIT_CODE='" + gstrUNITID + "' AND  invoice_no = '" & Trim(Me.Ctlinvoice.Text) & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            End If
                            '***********To check if Tool Cost Deduction will be done or Not on 16/02/2004
                            If blnCheckToolCost = True Then
                                If Len(Trim(strUpdateAmorDtl)) > 0 Then
                                    mP_Connection.Execute(strUpdateAmorDtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                    '****************Add Rajani Kant On 19/08/2004
                                    If Len(Trim(strupdateamordtlbom)) > 0 Then
                                        mP_Connection.Execute(strupdateamordtlbom, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                    End If
                                    '*************************************
                                End If
                            End If

                            '*********************
                            'changes done by nisha on 13/05/2003
                            'Change by Sandeep
                            If UCase(cmbInvType.Text) = "JOBWORK INVOICE" Then

                                If GetBOMCheckFlagValue("BomCheck_Flag") Then
                                    'changes ends here 13/05/2003
                                    mP_Connection.Execute("SET DATEFORMAT 'DMY'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                    mP_Connection.Execute(mstrAnnex, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                ElseIf CBool(mblnJobWkFormulation) = True Then
                                    If CheckforCustSupplyMaterial(True) = False Then
                                        CustomRollbackTrans()
                                        MsgBox("Customer Supplied Material is not reconcile with Customer Formulation. Transaction cannot Saved.", MsgBoxStyle.Critical, "eMPro")
                                        Exit Sub
                                    End If
                                End If

                            End If

                            If UCase(Me.lbldescription.Text) = "REJ" Then
                                If Len(Trim(mCust_Ref)) > 0 Then
                                    mP_Connection.Execute(strupdateGrinhdr, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                End If
                            End If

                            '--------------updation of Schedules-------------------------------------------------------
                            'changes done by nisha on 06/10/2004 for schedule check on condition
                            'If ((UCase(Trim(cmbInvType.Text)) = "NORMAL INVOICE") And (UCase(CStr((Trim(CmbCategory.Text)) = "FINISHED GOODS")) Or (UCase(Trim(CmbCategory.Text)) = "TRADING GOODS"))) Or (UCase(Trim(cmbInvType.Text)) = "JOBWORK INVOICE") Or (UCase(Trim(cmbInvType.Text)) = "EXPORT INVOICE") Or (UCase(Trim(cmbInvType.Text)) = "SERVICE INVOICE") Then
                            If ((UCase(Trim(cmbInvType.Text)) = "TRANSFER INVOICE") And (UCase(CStr((Trim(CmbCategory.Text)) = "FINISHED GOODS")))) Then
                                '-------JS on 05/09/2004-------------------------------------------------------------------
                                'Condition added by Arshad for Service Invoice
                                If UCase(Trim(cmbInvType.Text)) <> "SERVICE INVOICE" Then
                                    'Added for Issue Id 21840 Starts
                                    If blnInvoiceAgainstMultipleSO = False Then
                                        blnDSTracking = CBool(Find_Value("SELECT isnull(DSWiseTracking,0)as DSWiseTracking FROM sales_parameter Where unit_code='" & gstrUNITID & "'"))
                                        If blnDSTracking = True And AllowTextFileGeneration_SMIIEL(strAccountCode) = True And CBool(Find_Value("SELECT SMIEL_FTP_INVOICE FROM SALES_PARAMETER Where unit_code='" & gstrUNITID & "'")) = True Then
                                            If Not UpdateMktSchedule() Then
                                                CustomRollbackTrans()
                                                Exit Sub
                                            End If
                                        End If
                                    End If
                                    'Added for Issue Id 21840 Starts
                                ElseIf UCase(Trim(cmbInvType.Text)) = "SERVICE INVOICE" Then
                                    If Not mblnServiceInvoiceWithoutSO Then
                                        If Not UpdateMktSchedule() Then
                                            'Commented for Issue ID eMpro -20080516 - 18915 Starts
                                            'MsgBox "Error in Schedule updations", vbInformation, "eMPro"
                                            'Commented for Issue ID eMpro -20080516 - 18915 Ends
                                            CustomRollbackTrans()
                                            Exit Sub
                                        End If
                                    End If
                                End If
                            End If
                            '--------------updation of Schedules-------------------------------------------------------
                            '10237233
                            mstrupdateASNdtl = ""
                            mstrupdateASNCumFig = ""
                            If AllowASNTextFileGeneration(Trim(strAccountCode)) = True Then
                                mP_Connection.Execute("UPDATE MKT_ASN_INVDTL SET DOC_NO=" & Trim(mInvNo) & " where dOC_NO=" & Trim(Ctlinvoice.Text) & " and unit_code='" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                If FordASNFileGeneration(mInvNo) = False Then
                                    CustomRollbackTrans()
                                    Exit Sub
                                Else
                                    If Len(mstrupdateASNdtl) > 0 Then
                                        mP_Connection.Execute(mstrupdateASNdtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                        mP_Connection.Execute(mstrupdateASNCumFig, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                    End If
                                End If
                            End If
                            '10237233
                            If optInvYes(0).Checked = True And AllowASNPrinting(strAccountCode) = True Then
                                If txtASNNumber.Text.Trim.Length > 0 Then
                                    mP_Connection.Execute("Update CreatedASN Set DOC_NO='" & Trim$(mInvNo) & "',Updatedon=getdate() where doc_no='" & Trim$(Me.Ctlinvoice.Text) & "' and UNIT_CODE='" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                End If
                            End If
                            If optInvYes(0).Checked = True Then
                                If DataExist("SELECT TOP 1 1 FROM SALES_PARAMETER WHERE INVOICE_LOCKING_ENTRY_SAMEDATE=1  and UNIT_CODE = '" & gstrUNITID & "'") Then
                                    mP_Connection.Execute("update Saleschallan_dtl set invoice_date= Convert(varchar(12), getdate(), 106) WHERE UNIT_CODE='" + gstrUNITID + "' AND  Doc_no = " & mInvNo, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                End If
                            End If
                            'Accounts Posting is done here
                            If mblnpostinfin = True Then
                                objDrCr = New prj_DrCrNote.cls_DrCrNote(GetServerDate)
                                If UCase(Trim(cmbInvType.Text)) <> "REJECTION" Then
                                    strRetVal = objDrCr.SetARInvoiceDocument(gstrUNITID, mstrMasterString, mstrDetailString, prj_GLTransactions.cls_GLTransactions.udtOperationType.optInsert, gstrCONNECTIONSTRING, "", mP_Connection)
                                Else
                                    '''Changes done By ashutosh on 12 jun 2007, Issue Id:19934.
                                    If RejInvOptionalPostingFlag() = True Then
                                        'changes done by nisha for Rejection Accounts Posting
                                        If (chkAcceffects.CheckState = System.Windows.Forms.CheckState.Checked) And (chkAcceffects.Enabled = True) Then
                                            strRetVal = "Y"
                                        Else
                                            prj_DocGenerator.cls_DocumentGenerator.gbln_AR_AP_Dr_Cr_Doc_Sub_Category = "DR"
                                            strRetVal = objDrCr.SetAPDocument(gstrUNITID, mstrMasterString, mstrDetailString, prj_GLTransactions.cls_GLTransactions.udtOperationType.optInsert, gstrCONNECTIONSTRING, "", mP_Connection)
                                            prj_DocGenerator.cls_DocumentGenerator.gbln_AR_AP_Dr_Cr_Doc_Sub_Category = ""
                                        End If
                                    Else
                                        prj_DocGenerator.cls_DocumentGenerator.gbln_AR_AP_Dr_Cr_Doc_Sub_Category = "DR"
                                        strRetVal = objDrCr.SetAPDocument(gstrUNITID, mstrMasterString, mstrDetailString, prj_GLTransactions.cls_GLTransactions.udtOperationType.optInsert, gstrCONNECTIONSTRING, "", mP_Connection)
                                        prj_DocGenerator.cls_DocumentGenerator.gbln_AR_AP_Dr_Cr_Doc_Sub_Category = ""
                                    End If
                                    '''Changes for Issue Id:19934 end here.
                                End If
                                strRetVal = CheckString(strRetVal)
                                objDrCr = Nothing
                            Else
                                strRetVal = "Y"
                            End If

                            If Not strRetVal = "Y" Then
                                MsgBox(strRetVal, MsgBoxStyle.Information, "eMPro")
                                CustomRollbackTrans()
                                Exit Sub
                            Else
                                Frm.ShowPrintButton = False


                                If mblncustomerspecificreport = False Then
                                    Frm.Show()
                                End If
                                blnDSTracking = CBool(Find_Value("SELECT isnull(DSWiseTracking,0)as DSWiseTracking FROM sales_parameter Where Unit_code='" & gstrUNITID & "' "))
                                If ((UCase(Trim(cmbInvType.Text)) = "TRANSFER INVOICE") And (UCase(CStr((Trim(CmbCategory.Text)) = "FINISHED GOODS")))) Then
                                    If blnDSTracking = True And AllowTextFileGeneration_SMIIEL(strAccountCode) = True And CBool(Find_Value("SELECT SMIEL_FTP_INVOICE FROM SALES_PARAMETER Where Unit_code='" & gstrUNITID & "' ")) = True Then
                                        If CheckInvoices(mInvNo, strAccountCode) = False Then
                                            CustomRollbackTrans()
                                            Exit Sub
                                        End If
                                    End If
                                End If
                                '''Changes done by Ashutosh on 12 Jun 2007, Issue Id:19934
                                If VerifyInvPostingFlag() = True Then
                                    If InvoicePostingFlag() = True Then
                                        If UCase(Trim(cmbInvType.Text)) <> "REJECTION" Then
                                            rsInvDataInFin = New ClsResultSetDB_Invoice
                                            rsInvoiceDataInFin = New ClsResultSetDB_Invoice
                                            rsInvDataInFin.GetResult("Select top 1 * from ar_docMaster where  docM_unit ='" + gstrUNITID + "' AND docm_vono='" & Trim(CStr(mInvNo)) & "'")
                                            If rsInvDataInFin.GetNoRows > 0 Then
                                                rsInvoiceDataInFin.GetResult("Select top 1 * from saleschallan_dtl where  unit_code='" + gstrUNITID + "' AND doc_no='" & Trim(CStr(mInvNo)) & "'")
                                                If rsInvoiceDataInFin.GetNoRows > 0 Then
                                                    mP_Connection.CommitTrans()

                                                    If mblncustomerspecificreport = True And GetPrintMethod(strAccountCode).ToUpper() <> "TATA" Then
                                                        strSql = "{SalesChallan_Dtl.Location_Code}='" & Trim(txtUnitCode.Text) & "'  and {SalesChallan_Dtl.Unit_Code}='" & gstrUNITID & "'  and {SalesChallan_Dtl.Doc_No} =" & mInvNo & " and {SalesChallan_Dtl.Invoice_Type}"
                                                        strSql = strSql & " = '" & Trim(Me.lbldescription.Text) & "'  and {SalesChallan_Dtl.Sub_Category} = '" & Trim(Me.lblcategory.Text) & "'"
                                                        RdAddSold.DataDefinition.RecordSelectionFormula = strSql
                                                        Frm.Show()
                                                    Else
                                                        If mblncustomerspecificreport = True And GetPrintMethod(strAccountCode).ToUpper() = "TATA" Then
                                                            Frm.Show()
                                                        End If
                                                    End If
                                                Else
                                                    MsgBox("Invoice not posted to Invoice Table. Try Again!!! ", MsgBoxStyle.Information, ResolveResString(100))
                                                    CustomRollbackTrans()
                                                    mP_Connection.Execute("delete from ar_docmaster WHERE docM_unit='" + gstrUNITID + "' AND  docm_vono='" & mInvNo & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                                    mP_Connection.Execute("delete from ar_docdtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  docd_vono='" & mInvNo & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                                    mP_Connection.Execute("delete from fin_gltrans WHERE glt_UntCodeID='" + gstrUNITID + "' AND  glt_srcdocno='" & mInvNo & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                                    rsInvoiceDataInFin.ResultSetClose()
                                                    rsInvDataInFin.ResultSetClose()
                                                    Exit Sub
                                                End If
                                            Else
                                                MsgBox("Invoice not posted to accounts. Try Again!!! ", MsgBoxStyle.Information, ResolveResString(100))
                                                CustomRollbackTrans()
                                                mP_Connection.Execute("delete from ar_docmaster WHERE docM_unit='" + gstrUNITID + "' AND  docm_vono='" & mInvNo & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                                mP_Connection.Execute("delete from ar_docdtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  docd_vono='" & mInvNo & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                                mP_Connection.Execute("delete from fin_gltrans WHERE glt_UntCodeID='" + gstrUNITID + "' AND  glt_srcdocno='" & mInvNo & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

                                                rsInvDataInFin.ResultSetClose()
                                                Exit Sub
                                            End If
                                            rsInvDataInFin.ResultSetClose()
                                        Else
                                            If (chkAcceffects.CheckState <> System.Windows.Forms.CheckState.Checked) Then
                                                rsInvDataInFin = New ClsResultSetDB_Invoice
                                                rsInvDataInFin.GetResult("Select top 1 apdocm_venprdocNo from ap_docMaster where APdocM_unit ='" + gstrUNITID + "' AND apdocm_venprdocNo='" & Trim(CStr(mInvNo)) & "'")
                                                If rsInvDataInFin.GetNoRows > 0 Then
                                                    mP_Connection.CommitTrans()
                                                Else
                                                    MsgBox("Invoice not posted to accounts. Try Again!!! ", MsgBoxStyle.Information, ResolveResString(100))
                                                    CustomRollbackTrans()
                                                    rsInvDataInFin.ResultSetClose()
                                                    Exit Sub
                                                End If
                                                rsInvDataInFin.ResultSetClose()
                                            Else
                                                mP_Connection.CommitTrans()
                                            End If
                                        End If
                                    Else
                                        mP_Connection.CommitTrans()
                                    End If
                                Else


                                    mP_Connection.CommitTrans()
                                End If
                                '''Changes for Issue Id:19934 end here.
                                If gBlnWIPFGProcess AndAlso WIP_FG_Customer(strAccountCode) Then
                                    If mblnskipdacinvoicebincheck = False Then
                                        Dim svcWIPFGInv As New eMPRO_WIP_FG_INVOICE_COMPLETION
                                        Dim strResult As String = String.Empty
                                        If DataExist("SELECT TOP 1 1 FROM WIP_FG_PICKLIST_INV_DTL  WHERE  UNIT_CODE='" & gstrUNITID & "' AND TMP_INVOICENO =" & Trim(Ctlinvoice.Text)) Then
                                            mP_Connection.Execute("UPDATE WIP_FG_PICKLIST_INV_DTL  SET INVOICENO = '" & mInvNo & "' WHERE UNIT_CODE='" + gstrUNITID + "' AND   TMP_INVOICENO = " & Trim(Ctlinvoice.Text), , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                            strResult = svcWIPFGInv.ProcessInvoice(gstrUNITID, "R", Trim(Ctlinvoice.Text))
                                            If strResult.ToUpper() <> "SUCCESS|" Then
                                                WIP_FG_SAVE_EXCEPTION_LOG(Convert.ToInt32(mInvNo), "svcWIPFGInv.ProcessInvoice", strResult, gstrUNITID, "R", Trim(Ctlinvoice.Text), "")
                                            End If
                                        End If
                                        svcWIPFGInv.Dispose()
                                    Else
                                        '10623079
                                        mP_Connection.Execute("UPDATE WIP_FG_PICKLIST_INV_DTL  SET INVOICENO = '" & mInvNo & "' WHERE UNIT_CODE='" + gstrUNITID + "' AND   TMP_INVOICENO = " & Trim(Ctlinvoice.Text), , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                        '10623079
                                    End If

                                End If

                                'Changed by nisha on 05/09/2003
                                If UCase(Trim(cmbInvType.Text)) = "REJECTION" Then
                                    If UCase(Trim(CmbCategory.Text)) = "REJECTION" Then
                                        '''Changes for Issue Id:19934, on 05 jun 2007
                                        If RejInvOptionalPostingFlag() = True And chkAcceffects.CheckState = System.Windows.Forms.CheckState.Checked Then
                                            mP_Connection.Execute("update salesChallan_Dtl set RejectionPosting = 1 WHERE UNIT_CODE='" + gstrUNITID + "' AND  doc_no = " & mInvNo, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                        Else
                                            mP_Connection.Execute("update salesChallan_Dtl set RejectionPosting = 0 WHERE UNIT_CODE='" + gstrUNITID + "' AND  doc_no = " & mInvNo, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                        End If
                                    End If
                                End If
                                'Changed by nisha on 05/09/2003 ends here
                                Call Logging_Starting_End_Time("Invoice Locking ", strtime, "Saved", mInvNo)
                                MsgBox("Invoice has been locked successfully with number " & mInvNo, MsgBoxStyle.Information, "eMPro")
                                '                            Write_In_Log_File(GetServerDateTime() & " : Invoice has been locked successfully")
                                'Added by Prashant dhingra for Auto Invoicing 10140220 
                                'SATISH KESHAERWANI CHANGE
                                If chktkmlbarcode.Checked = True Then
                                    Print_barcodelabel(mInvNo)
                                End If
                                'SATISH KESHAERWANI CHANGE

                                txtASNNumber.Text = String.Empty
                                mP_Connection.Execute("Exec USP_AUTOINVOICELOCKING '" & gstrUNITID & "', '" & TempInvNo & "','" & mInvNo & "', '" & mP_User & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                'end
                                If CBool(Find_Value("select TextPrinting from sales_parameter WHERE UNIT_CODE='" + gstrUNITID + "'")) Then
                                    cmdPrint.Enabled = True 'ENABLE PRINT BUTTON
                                End If



                                '****CODE ADDED BY SIDDHARTH RANJAN***********
                                If UCase(Trim(GetPlantName)) = "SMIEL" Or UCase(Trim(GetPlantName)) = "SUMIT" Then
                                    Send_Report_Printer_Smiel(RdAddSold, RepPath, Frm)
                                    Frm.ShowExportButton = False
                                    Frm.ShowPrintButton = False
                                    Frm.Show()
                                    '' rptinvoice.WindowShowExportBtn = False
                                    '' rptinvoice.WindowShowPrintSetupBtn = False
                                    ''  rptinvoice.WindowShowPrintBtn = False
                                    ''  rptinvoice.Destination = Crystal.DestinationConstants.crptToWindow
                                    ''  rptinvoice.Action = 1
                                End If

                                '****END OF CODE ADDED BY SIDDHARTH RANJAN***********

                                Ctlinvoice.Text = ""
                                'Added for Issue ID 22035 Starts
                                If GetPlantName() = "SMIEL" Then
                                    If CBool(Find_Value("SELECT ISnull(DSWiseTracking,0) FROM sales_parameter WHERE UNIT_CODE='" + gstrUNITID + "'")) = True Then
                                        If SendMailDSOverKnockoff(CStr(mInvNo)) = False And Len(mstrError) > 0 Then
                                            MsgBox(mstrError, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                                        End If
                                    End If
                                End If
                                'Added for Issue ID 22035 Ends
                            End If

                            '''***** Changes done by Ashutosh on 24-12-2005,Issue Id:16685
                            If UCase(Trim(txtUnitCode.Text)) = "SUN" Then
                                'On Error Resume Next
                                'shalini
                                'Kill("C:\InvoicePrint.txt")
                                Kill(strCitrix_Inv_Pronting_Loc & "InvoicePrint.txt")
                                'On Error GoTo Err_Handler
                                objInvoicePrint = New prj_InvoicePrinting.clsInvoicePrinting(gstrDateFormat)
                                objInvoicePrint.Print_Invoice(gstrUNITID, True, (txtUnitCode.Text), CStr(mInvNo), dtpRemoval.Text & " " & dtpRemovalTime.Value.Hour & ":" & dtpRemovalTime.Value.Minute)
                                rtbInvoicePreview.LoadFile(objInvoicePrint.FileName, RichTextBoxStreamType.PlainText)
                                rtbInvoicePreview.BackColor = System.Drawing.Color.White
                                cmdPrint.Image = My.Resources.ico231.ToBitmap
                                cmdClose.Image = My.Resources.ico217.ToBitmap

                                FraInvoicePreview.Height = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Me.Height) - 1300)
                                FraInvoicePreview.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Width) - 400)
                                FraInvoicePreview.Left = VB6.TwipsToPixelsX(100)
                                FraInvoicePreview.Top = ctlFormHeader1.Height

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
                                FraInvoicePreview.Visible = True
                                FraInvoicePreview.Enabled = True
                                FraInvoicePreview.BringToFront()
                                rtbInvoicePreview.Focus()
                                Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
                            End If
                            '''***** Changes on 24-12-2005 end here.
                        Else
                            blninvoicelockYES = False
                            If (Not (UCase(Trim(GetPlantName)) = "SMIEL" Or (UCase(Trim(GetPlantName)) = "SUMIT"))) And DataExist("SELECT TOP 1 1 FROM SALES_PARAMETER WHERE UNIT_CODE='" + gstrUNITID + "' AND  PRINT_WITHOUT_LOCK_REQ = 0") Then
                                ''  rptinvoice.WindowShowExportBtn = False
                                ''   rptinvoice.WindowShowPrintSetupBtn = False
                                ''  rptinvoice.WindowShowPrintBtn = False
                                ''   intmaxsubreportloop = Me.rptinvoice.GetNSubreports
                                ''    For intsubreportloopcounter = 0 To intmaxsubreportloop - 1
                                ''      Me.rptinvoice.SubreportToChange = Me.rptinvoice.GetNthSubreportName(intsubreportloopcounter)
                                ''     Me.rptinvoice.Connect = gstrREPORTCONNECT
                                Frm.ShowExportButton = False
                                Frm.ShowPrintButton = False
                                Frm.glbblnHidePrintButton = True
                                Frm.Show()


                                ''    Next
                                ''    Me.rptinvoice.SubreportToChange = ""
                                ''    rptinvoice.WindowAllowDrillDown = True
                                ''    rptinvoice.Connect = gstrREPORTCONNECT
                                '10109115  end 
                                If UCase(Trim(GetPlantName)) = "VF1" Or UCase(Trim(GetPlantName)) = "RSA" Then
                                    If mstrReportFilename = "rptinvoicemate_VF1" Or mstrReportFilename = "rptinvoiceMATE_RSA" Then
                                        '' rptinvoice.set_Formulas(30, "Deliverynoteno='" & mInvNo & "'")
                                        RdAddSold.DataDefinition.FormulaFields("Deliverynoteno").Text = "'" + CStr(mInvNo) + "'"
                                    End If
                                    Frm.EnableDrillDown = True
                                    'Frm.Show()
                                End If

                                'rptinvoice.Action = 1
                            ElseIf (UCase(Trim(GetPlantName)) = "SMIEL" Or (UCase(Trim(GetPlantName)) = "SUMIT")) Then
                                '' rptinvoice.WindowShowExportBtn = False
                                ''  rptinvoice.WindowShowPrintSetupBtn = False
                                ''  rptinvoice.WindowShowPrintBtn = False
                                ''   rptinvoice.Action = 1
                                Frm.ShowExportButton = False
                                Frm.ShowPrintButton = False
                                Frm.Show()
                            End If



                        End If
                    Else
                        If CBool(Find_Value("select TextPrinting from sales_parameter where UNIT_CODE= '" & gstrUNITID & "'")) Then
                            cmdPrint.Enabled = True     'ENABLE PRINT BUTTON
                        Else

                            If Not blnIsReportDisplayed Then
                                'issue id 10109115 
                                '' intmaxsubreportloop = Me.rptinvoice.GetNSubreports
                                '' For intsubreportloopcounter = 0 To intmaxsubreportloop - 1
                                ''     Me.rptinvoice.SubreportToChange = Me.rptinvoice.GetNthSubreportName(intsubreportloopcounter)
                                '                            Me.rptinvoice.Connect = gstrREPORTCONNECT
                                ''  Next
                                ''   Me.rptinvoice.SubreportToChange = ""
                                ''   rptinvoice.WindowAllowDrillDown = True
                                ''   rptinvoice.Connect = gstrREPORTCONNECT
                                If UCase(Trim(GetPlantName)) = "VF1" Or UCase(Trim(GetPlantName)) = "RSA" Then
                                    If mstrReportFilename = "rptinvoicemate_VF1" Or mstrReportFilename = "rptinvoiceMATE_RSA" Then
                                        ''rptinvoice.set_Formulas(30, "Deliverynoteno='" & Ctlinvoice.Text & "'")
                                        RdAddSold.DataDefinition.FormulaFields("Deliverynoteno").Text = "'" + Ctlinvoice.Text + "'"
                                    End If
                                    Frm.EnableDrillDown = True
                                    Frm.Show()
                                Else
                                    'RdAddSold.PrintOptions.PaperSize = PaperSize.PaperLetter
                                    'RdAddSold.PrintToPrinter(1, False, 1, 1)
                                    Frm.ShowPrintButton = False
                                    Frm.Show()
                                End If

                                'issue id 10109115 
                                ''rptinvoice.Action = 1       'PRINT INVOICE REPORT
                            End If
                        End If
                        'If Not blnIsReportDisplayed Then
                        '    rptinvoice.Action = 1
                        'End If
                    End If
                    '*********************Addition Ends*********************************************

                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT_TO_PRINTER
                    If CBool(Find_Value("select PRINT_DISABLED from sales_parameter WHERE UNIT_CODE='" + gstrUNITID + "'")) = True Then
                        MsgBox("Kindly Click on PREVIEW", MsgBoxStyle.Information, ResolveResString(100))
                        Exit Sub
                    End If
                    ''HILEX A4 BARCODE'
                    'If AllowBarCodePrinting(strAccountCode) = True Then
                    '    If optInvYes(0).Checked = True Then
                    '        '------------------------------------------------------------------------------------
                    '        rsGENERATEBARCODE = New ClsResultSetDB_Invoice
                    '        rsGENERATEBARCODE.GetResult("SELECT PRINT_METHOD FROM CUSTOMER_MST C WHERE C.UNIT_CODE='" & gstrUNITID & "' AND C.CUSTOMER_CODE='" & strAccountCode & "'")
                    '        strPrintMethod = UCase(rsGENERATEBARCODE.GetValue("PRINT_METHOD").ToString)
                    '        rsGENERATEBARCODE.ResultSetClose()
                    '        rsGENERATEBARCODE = Nothing

                    '        If optInvYes(0).Checked = True Then
                    '            If strPrintMethod = "NORMAL" Then
                    '                strBarcodeMsg = ObjBarcodeHMI.GenerateBarCode_LINELEVEL_SALESORDER_2dbarcode_hilex(gstrUserMyDocPath, mInvNo, "NORMAL", "", "", True, Trim(Ctlinvoice.Text), gstrCONNECTIONSTRING)
                    '                Dim strQuery As String

                    '                If VB.Left(strBarcodeMsg, 1) <> "Y" Then
                    '                    MsgBox("Problem While Generating Barcode Image.", vbInformation, ResolveResString(100))
                    '                    Exit Sub
                    '                Else
                    '                    If SaveBarCodeImage_singlelevelso_2DBARCODE(Ctlinvoice.Text, gstrUserMyDocPath, Mid(strBarcodeMsg, 3)) = False Then
                    '                        MsgBox("Problem While saving Barcode Image.", vbInformation, ResolveResString(100))
                    '                        Exit Sub
                    '                    Else
                    '                        mP_Connection.Execute(" UPDATE T SET T.BARCODEIMAGE =SC.BARCODEIMAGE FROM SALESCHALLAN_DTL SC,TMP_INVOICEPRINT  T " & _
                    '                                                       " WHERE SC.UNIT_CODE = T.UNIT_CODE AND SC.DOC_NO =T.DOC_NO AND SC.UNIT_CODE='" & gstrUNITID & "' AND " & _
                    '                                                       " SC.DOC_NO='" & Ctlinvoice.Text.Trim & "' AND T.IP_ADDRESS='" & gstrIpaddressWinSck & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    '                    End If
                    '                End If
                    '            Else
                    '                rsGENERATEBARCODE = New ClsResultSetDB_Invoice
                    '                rsGENERATEBARCODE.GetResult("SELECT PRINT_METHOD,ITEM_CODE,CUST_ITEM_CODE FROM SALES_DTL SD ,SALESCHALLAN_DTL SC ,CUSTOMER_MST C WHERE C.UNIT_CODE=SC.UNIT_CODE AND C.CUSTOMER_CODE=SC.ACCOUNT_CODE AND SC.UNIT_CODE=SD.UNIT_CODE AND " & _
                    '                                                    " SC.DOC_NO=SD.DOC_NO  AND SC.UNIT_CODE='" & gstrUNITID & "' AND  " & _
                    '                                                    " SC.DOC_NO= " & Ctlinvoice.Text & " ORDER BY CUST_ITEM_CODE ")

                    '                intTotalNoofitemsinInvoices = rsGENERATEBARCODE.GetNoRows
                    '                rsGENERATEBARCODE.MoveFirst()
                    '                For intRow = 1 To intTotalNoofitemsinInvoices
                    '                    strBarcodeMsg = ObjBarcodeHMI.GenerateBarCode_LINELEVEL_SALESORDER(gstrUserMyDocPath, mInvNo, rsGENERATEBARCODE.GetValue("PRINT_METHOD").ToString, rsGENERATEBARCODE.GetValue("Cust_Item_Code").ToString, rsGENERATEBARCODE.GetValue("Item_Code").ToString, True, Trim(Ctlinvoice.Text), gstrCONNECTIONSTRING)
                    '                    If VB.Left(strBarcodeMsg, 1) <> "Y" Then
                    '                        MsgBox("Problem While Generating Barcode Image.", vbInformation, ResolveResString(100))
                    '                        Exit Sub
                    '                    End If
                    '                    If SaveBarCodeImage_singlelevelso(Ctlinvoice.Text, rsGENERATEBARCODE.GetValue("Cust_Item_Code").ToString, rsGENERATEBARCODE.GetValue("Item_Code").ToString, gstrUserMyDocPath, intRow) = False Then
                    '                        MsgBox("Problem While saving Barcode Image.", vbInformation, ResolveResString(100))
                    '                        Exit Sub
                    '                    End If
                    '                    rsGENERATEBARCODE.MoveNext()
                    '                Next

                    '                rsGENERATEBARCODE.ResultSetClose()
                    '                rsGENERATEBARCODE = Nothing
                    '            End If

                    '        End If
                    '    End If
                    'End If

                    ''HILEX A4 BARCODE
                    If optInvYes(0).Checked = True And AllowASNPrinting(strAccountCode) = True Then
                        If txtASNNumber.Text.Trim.Length > 0 Then
                            'If DUPLICATEASN(Me.txtASNNumber.Text.Trim) = True Then
                            strAsnInvoice = DUPLICATEASN(Me.txtASNNumber.Text.Trim)
                            If mblnDuplicateASNExist = True Then
                                MsgBox("Already Used ASN Number in Different Invoice. Invoice No:" & strAsnInvoice & " can't save", MsgBoxStyle.Information, ResolveResString(100))
                                mP_Connection.Execute("Update CreatedASN Set ASN_NO='' where doc_no='" & Trim$(Me.Ctlinvoice.Text) & "'  and unit_code='" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                txtASNNumber.Text = ""
                                Exit Sub
                            End If
                        End If
                    End If

                    If DataExist("SELECT TOP 1 1 FROM SALES_DTL WHERE DOC_NO='" & Me.Ctlinvoice.Text.Trim & "' And Rate = 0  and unit_code='" & gstrUNITID & "'") Then
                        MsgBox("Invoice Rate Cannot be zero , Kindly EDIT/UPDATE ", MsgBoxStyle.Information, ResolveResString(100))
                        Exit Sub
                    End If

                    'PRASHANT RAJPAL CHANGED DATED 13 MAR 2014'
                    If optInvYes(0).Checked = True And InvAgstBarCode() = True And (mstrFGDomestic = "1" Or mstrFGDomestic = "2") Then
                        If BarCodeTracking(Trim(Ctlinvoice.Text), "LOCK") = False Then
                            If mblnQuantityCheck = False Then
                                Exit Sub
                            End If
                        End If
                    End If
                    'PRASHANT RAJPAL CHANGED ENDED 13 MAR 2014'

                    '                Write_In_Log_File(GetServerDateTime() & "  : Process Started - Printer")
                    '102027599
                    If optInvYes(0).Checked = False Then
                        If mblnEwaybill_Print Then
                            Call IRN_QRBarcode()
                        End If
                    End If
                    If InvoiceGeneration(RdAddSold, RepPath, Frm) = True Then


                        'Changes for Hero invoice Start-------------------
                        If (Find_Value("SELECT top 1 1 FROM LISTS WHERE UNIT_CODE='" & gstrUNITID & "' AND KEY1='HEROBAR' and Key2='" & strAccountCode & "' ")) = "1" Then
                            Dim strBarcodestring1 As String = ""
                            Dim strBarcodestring2 As String = ""
                            Dim strBarcodeString3 As String = ""
                            Dim totalAccesibleAmt As Double = 0
                            Dim totalCgst As Double = 0
                            Dim totalIgst As Double = 0
                            Dim totalSgst As Double = 0
                            Dim totaltcs As Double = 0
                            rsGENERATEBARCODE = New ClsResultSetDB_Invoice

                            ' rsGENERATEBARCODE.GetResult("SELECT  C.CUST_VENDOR_CODE,A.CUST_REF,A.DOC_NO,Convert(varchar,A.INVOICE_DATE,104) INVOICE_DATE,GSTIN_ID, Convert(numeric(8,2),ACCESSIBLE_AMOUNT) ACCESSIBLE_AMOUNT, Convert(numeric(8,2),C.TOTAL_AMOUNT) TOTAL_AMOUNT,A.VEHICLE_NO,Convert(numeric(8,2),isnull(SGSTTXRT_AMOUNT,0)) SGSTTXRT_AMOUNT, " & _
                            '" Convert(numeric(8,2),isnull(IGSTTXRT_AMOUNT,0)) IGSTTXRT_AMOUNT, Convert(numeric(8,2),isnull(CGSTTXRT_AMOUNT,0)) CGSTTXRT_AMOUNT,CUST_item_CODE,HSNSACCODE,Convert(numeric(8,2),SALES_QUANTITY) SALES_QUANTITY,Convert(numeric(8,2),Rate) Rate " & _
                            '" FROM SALESCHALLAN_DTL A INNER JOIN GEN_UNITMASTER B ON A.UNIT_CODE=B.UNT_CODEID " & _
                            '" INNER JOIN TMP_INVOICEPRINT C ON B.UNT_CODEID=C.UNIT_CODE AND A.DOC_NO=C.DOC_NO " & _
                            '" WHERE A.DOC_NO='" & Ctlinvoice.Text.Trim & "' AND A.UNIT_CODE='" & gstrUNITID & "' ")

                            rsGENERATEBARCODE.GetResult("SELECT  C.CUST_VENDOR_CODE,A.CUST_REF,A.DOC_NO,Convert(varchar,A.INVOICE_DATE,104) INVOICE_DATE,GSTIN_ID, Convert(numeric(19,2),ACCESSIBLE_AMOUNT) ACCESSIBLE_AMOUNT, Convert(numeric(19,2),C.TOTAL_AMOUNT) TOTAL_AMOUNT,A.VEHICLE_NO,Convert(numeric(19,2),isnull(SGSTTXRT_AMOUNT,0)) SGSTTXRT_AMOUNT, " &
                           " Convert(numeric(19,2),isnull(IGSTTXRT_AMOUNT,0)) IGSTTXRT_AMOUNT, Convert(numeric(19,2),isnull(CGSTTXRT_AMOUNT,0)) CGSTTXRT_AMOUNT,CUST_item_CODE,HSNSACCODE,Convert(numeric(12,2),SALES_QUANTITY) SALES_QUANTITY,Convert(numeric(19,2),Rate) Rate ,Convert(numeric(19,2),isnull(TCSAMOUNT,0)) TCSAMOUNT " &
                           " FROM SALESCHALLAN_DTL A INNER JOIN GEN_UNITMASTER B ON A.UNIT_CODE=B.UNT_CODEID " &
                           " INNER JOIN TMP_INVOICEPRINT C ON B.UNT_CODEID=C.UNIT_CODE AND A.DOC_NO=C.DOC_NO " &
                           " WHERE A.DOC_NO='" & Ctlinvoice.Text.Trim & "' AND A.UNIT_CODE='" & gstrUNITID & "' AND C.IP_ADDRESS='" & gstrIpaddressWinSck & "'")


                            While Not rsGENERATEBARCODE.EOFRecord
                                totalAccesibleAmt = totalAccesibleAmt + Convert.ToDouble(rsGENERATEBARCODE.GetValue("ACCESSIBLE_AMOUNT").ToString)
                                If optInvYes(0).Checked = True Then
                                    strBarcodestring1 = rsGENERATEBARCODE.GetValue("CUST_VENDOR_CODE").ToString & vbTab & rsGENERATEBARCODE.GetValue("CUST_REF").ToString & vbTab & mInvNo.ToString & vbTab & rsGENERATEBARCODE.GetValue("INVOICE_DATE").ToString & vbTab & rsGENERATEBARCODE.GetValue("GSTIN_ID").ToString & vbTab & rsGENERATEBARCODE.GetValue("TOTAL_AMOUNT").ToString
                                Else
                                    strBarcodestring1 = rsGENERATEBARCODE.GetValue("CUST_VENDOR_CODE").ToString & vbTab & rsGENERATEBARCODE.GetValue("CUST_REF").ToString & vbTab & Ctlinvoice.Text.ToString & vbTab & rsGENERATEBARCODE.GetValue("INVOICE_DATE").ToString & vbTab & rsGENERATEBARCODE.GetValue("GSTIN_ID").ToString & vbTab & rsGENERATEBARCODE.GetValue("TOTAL_AMOUNT").ToString
                                End If

                                strBarcodeString3 = rsGENERATEBARCODE.GetValue("VEHICLE_NO").ToString
                                totalSgst = totalSgst + Convert.ToDouble(rsGENERATEBARCODE.GetValue("SGSTTXRT_AMOUNT").ToString)
                                totalIgst = totalIgst + Convert.ToDouble(rsGENERATEBARCODE.GetValue("IGSTTXRT_AMOUNT").ToString)
                                totalCgst = totalCgst + Convert.ToDouble(rsGENERATEBARCODE.GetValue("CGSTTXRT_AMOUNT").ToString)
                                totaltcs = Convert.ToDouble(rsGENERATEBARCODE.GetValue("TCSAMOUNT").ToString)
                                strBarcodestring2 = strBarcodestring2 & vbTab & rsGENERATEBARCODE.GetValue("CUST_item_CODE").ToString & vbTab & rsGENERATEBARCODE.GetValue("HSNSACCODE").ToString & vbTab & rsGENERATEBARCODE.GetValue("SALES_QUANTITY").ToString & vbTab & rsGENERATEBARCODE.GetValue("Rate").ToString
                                rsGENERATEBARCODE.MoveNext()
                            End While
                            'While Not rsGENERATEBARCODE.EOFRecord
                            '    totalAccesibleAmt = totalAccesibleAmt + Convert.ToDouble(rsGENERATEBARCODE.GetValue("ACCESSIBLE_AMOUNT").ToString)
                            '    strBarcodestring1 = rsGENERATEBARCODE.GetValue("CUST_VENDOR_CODE").ToString + " " + rsGENERATEBARCODE.GetValue("CUST_REF").ToString + " " + rsGENERATEBARCODE.GetValue("DOC_NO").ToString + " " + rsGENERATEBARCODE.GetValue("INVOICE_DATE").ToString + " " + rsGENERATEBARCODE.GetValue("GSTIN_ID").ToString + " " + rsGENERATEBARCODE.GetValue("TOTAL_AMOUNT").ToString
                            '    strBarcodeString3 = rsGENERATEBARCODE.GetValue("VEHICLE_NO").ToString
                            '    totalSgst = totalSgst + Convert.ToDouble(rsGENERATEBARCODE.GetValue("SGSTTXRT_AMOUNT").ToString)
                            '    totalIgst = totalIgst + Convert.ToDouble(rsGENERATEBARCODE.GetValue("IGSTTXRT_AMOUNT").ToString)
                            '    totalCgst = totalCgst + Convert.ToDouble(rsGENERATEBARCODE.GetValue("CGSTTXRT_AMOUNT").ToString)
                            '    strBarcodestring2 = strBarcodestring2 + " " + rsGENERATEBARCODE.GetValue("CUST_item_CODE").ToString + " " + rsGENERATEBARCODE.GetValue("HSNSACCODE").ToString + " " + rsGENERATEBARCODE.GetValue("SALES_QUANTITY").ToString + " " + rsGENERATEBARCODE.GetValue("Rate").ToString
                            '    rsGENERATEBARCODE.MoveNext()
                            'End While
                            rsGENERATEBARCODE.ResultSetClose()


                            Dim PDF417barcode As BarcodeLib.Barcode.PDF417.PDF417 = New BarcodeLib.Barcode.PDF417.PDF417()
                            PDF417barcode.UOM = BarcodeLib.Barcode.Linear.UnitOfMeasure.Pixel
                            PDF417barcode.LeftMargin = 0
                            PDF417barcode.RightMargin = 0
                            PDF417barcode.TopMargin = 0
                            PDF417barcode.BottomMargin = 0
                            PDF417barcode.ImageFormat = System.Drawing.Imaging.ImageFormat.Png

                            'PDF417barcode.Data = (strBarcodestring1 & vbTab & totalAccesibleAmt.ToString & vbTab & strBarcodeString3 & vbTab & totalSgst.ToString & vbTab & totalIgst.ToString & vbTab & totalCgst.ToString + strBarcodestring2).ToString().Trim
                            PDF417barcode.Data = (strBarcodestring1 & vbTab & totalAccesibleAmt.ToString & vbTab & strBarcodeString3 & vbTab & totalSgst.ToString & vbTab & totalIgst.ToString & vbTab & totalCgst.ToString & vbTab & totaltcs.ToString & strBarcodestring2).ToString().Trim
                            'PDF417barcode.Data = (strBarcodestring1 + " " + totalAccesibleAmt.ToString + " " + strBarcodeString3 + " " + totalSgst.ToString + " " + totalIgst.ToString + " " + totalCgst.ToString + strBarcodestring2).ToString().Trim
                            Dim imageData() As Byte = PDF417barcode.drawBarcodeAsBytes()

                            Dim cmd As SqlCommand = Nothing
                            cmd = New System.Data.SqlClient.SqlCommand()
                            With cmd
                                .CommandType = CommandType.Text
                                .CommandText = "UPDATE TMP_INVOICEPRINT SET barcodeimage=@QRIMAGE where DOC_NO='" & Ctlinvoice.Text.Trim & "' AND UNIT_CODE='" & gstrUNITID & "' "
                                .Parameters.Add("@QRIMAGE", SqlDbType.Image).Value = imageData
                                SqlConnectionclass.ExecuteNonQuery(cmd)

                            End With
                            'Dim cmd As SqlCommand = Nothing
                            'cmd = New System.Data.SqlClient.SqlCommand()
                            'cmd.Connection = SqlConnectionclass.GetConnection
                            'cmd.Transaction = cmd.Connection.BeginTransaction
                            'With cmd
                            '    .CommandText = String.Empty
                            '    .Parameters.Clear()
                            '    .CommandType = CommandType.Text
                            '    .CommandText = "UPDATE TMP_INVOICEPRINT SET barcodeimage=@QRIMAGE where DOC_NO='" & Ctlinvoice.Text.Trim & "' AND UNIT_CODE='" & gstrUNITID & "' "
                            '    .Parameters.Add("@QRIMAGE", SqlDbType.Image).Value = imageData

                            '    If cmd.ExecuteNonQuery() Then
                            '        cmd.Transaction.Commit()
                            '    Else
                            '        cmd.Transaction.Rollback()
                            '        MsgBox("Problem While saving Barcode Image.", vbInformation, ResolveResString(100))
                            '        Exit Sub
                            '    End If
                            'End With
                        End If

                        'Changes for Hero invoice End-------------------
                        'Changed for Issue ID eMpro-20090216-27468 Starts
                        'If chkLockPrintingFlag.CheckState = 0 Then
                        'Start 102027599
                        If AllowBarCodePrinting(strAccountCode) = True Then
                            If optInvYes(0).Checked = False And chkprintreprint.Checked = False And mblnEwaybill_Print = True Then
                                '------------------------------------------------------------------------------------
                                rsGENERATEBARCODE = New ClsResultSetDB_Invoice
                                rsGENERATEBARCODE.GetResult("SELECT PRINT_METHOD FROM CUSTOMER_MST C WHERE C.UNIT_CODE='" & gstrUNITID & "' AND C.CUSTOMER_CODE='" & strAccountCode & "'")
                                strPrintMethod = UCase(rsGENERATEBARCODE.GetValue("PRINT_METHOD").ToString)
                                rsGENERATEBARCODE.ResultSetClose()
                                rsGENERATEBARCODE = Nothing

                                If optInvYes(0).Checked = False And chkprintreprint.Checked = False And mblnEwaybill_Print = True Then
                                    If strPrintMethod = "NORMAL" Then
                                        If blnlinelevelcustomer = True Then
                                            strBarcodeMsg = ObjBarcodeHMI.GenerateBarCode_LINELEVEL_SALESORDER_2dbarcode_Normal_Hilex(gstrUserMyDocPath, Trim(Ctlinvoice.Text), "NORMAL", "", "", True, Trim(Ctlinvoice.Text), gstrCONNECTIONSTRING)

                                        Else
                                            strBarcodeMsg = ObjBarcodeHMI.GenerateBarCode_LINELEVEL_SALESORDER_2dbarcode_hilex(gstrUserMyDocPath, Trim(Ctlinvoice.Text), "NORMAL", "", "", True, Trim(Ctlinvoice.Text), gstrCONNECTIONSTRING)
                                        End If

                                        If VB.Left(strBarcodeMsg, 1) <> "Y" Then
                                            MsgBox("Problem While Generating Barcode Image.", vbInformation, ResolveResString(100))
                                            Exit Sub
                                        Else
                                            If SaveBarCodeImage_singlelevelso_2DBARCODE(Ctlinvoice.Text, gstrUserMyDocPath, Mid(strBarcodeMsg, 3)) = False Then
                                                MsgBox("Problem While saving Barcode Image.", vbInformation, ResolveResString(100))
                                                Exit Sub
                                            Else
                                                mP_Connection.Execute(" UPDATE T SET T.BARCODEIMAGE =SC.BARCODEIMAGE FROM SALESCHALLAN_DTL SC,TMP_INVOICEPRINT  T " &
                                                                       " WHERE SC.UNIT_CODE = T.UNIT_CODE AND SC.DOC_NO =T.DOC_NO AND SC.UNIT_CODE='" & gstrUNITID & "' AND " &
                                                                       " SC.DOC_NO='" & Ctlinvoice.Text.Trim & "' AND T.IP_ADDRESS='" & gstrIpaddressWinSck & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If

                        'End 102027599

                        Call ReprintQRbarcode()
                        If chkLockPrintingFlag.CheckState = 1 Then
                            'Changed for Issue ID eMpro-20090216-27468 Ends
                            If optInvYes(0).Checked = True Then
                                'HILEX A4 BARCODE'
                                If AllowBarCodePrinting(strAccountCode) = True Then
                                    If optInvYes(0).Checked = True Then
                                        '------------------------------------------------------------------------------------
                                        rsGENERATEBARCODE = New ClsResultSetDB_Invoice
                                        rsGENERATEBARCODE.GetResult("SELECT PRINT_METHOD FROM CUSTOMER_MST C WHERE C.UNIT_CODE='" & gstrUNITID & "' AND C.CUSTOMER_CODE='" & strAccountCode & "'")
                                        strPrintMethod = UCase(rsGENERATEBARCODE.GetValue("PRINT_METHOD").ToString)
                                        rsGENERATEBARCODE.ResultSetClose()
                                        rsGENERATEBARCODE = Nothing

                                        If optInvYes(0).Checked = True Then
                                            If strPrintMethod = "NORMAL" Then
                                                If blnlinelevelcustomer = True Then
                                                    strBarcodeMsg = ObjBarcodeHMI.GenerateBarCode_LINELEVEL_SALESORDER_2dbarcode_Normal_Hilex(gstrUserMyDocPath, mInvNo, "NORMAL", "", "", True, Trim(Ctlinvoice.Text), gstrCONNECTIONSTRING)
                                                Else
                                                    strBarcodeMsg = ObjBarcodeHMI.GenerateBarCode_LINELEVEL_SALESORDER_2dbarcode_hilex(gstrUserMyDocPath, mInvNo, "NORMAL", "", "", True, Trim(Ctlinvoice.Text), gstrCONNECTIONSTRING)
                                                End If

                                                Dim strQuery As String

                                                If VB.Left(strBarcodeMsg, 1) <> "Y" Then
                                                    MsgBox("Problem While Generating Barcode Image.", vbInformation, ResolveResString(100))
                                                    Exit Sub
                                                Else
                                                    If SaveBarCodeImage_singlelevelso_2DBARCODE(Ctlinvoice.Text, gstrUserMyDocPath, Mid(strBarcodeMsg, 3)) = False Then
                                                        MsgBox("Problem While saving Barcode Image.", vbInformation, ResolveResString(100))
                                                        Exit Sub
                                                    Else
                                                        mP_Connection.Execute(" UPDATE T SET T.BARCODEIMAGE =SC.BARCODEIMAGE FROM SALESCHALLAN_DTL SC,TMP_INVOICEPRINT  T " &
                                                                                       " WHERE SC.UNIT_CODE = T.UNIT_CODE AND SC.DOC_NO =T.DOC_NO AND SC.UNIT_CODE='" & gstrUNITID & "' AND " &
                                                                                       " SC.DOC_NO='" & Ctlinvoice.Text.Trim & "' AND T.IP_ADDRESS='" & gstrIpaddressWinSck & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                                    End If
                                                End If
                                                '25 oct 2017
                                            ElseIf strPrintMethod = "TOYOTA_NEW" Then
                                                strBarcodeMsg = ObjBarcodeHMI.GenerateBarCode_LINELEVEL_SALESORDER_2dbarcode_hilex(gstrUserMyDocPath, mInvNo, "TOYOTA_NEW", "", "", True, Trim(Ctlinvoice.Text), gstrCONNECTIONSTRING)
                                                Dim strQuery As String

                                                If VB.Left(strBarcodeMsg, 1) <> "Y" Then
                                                    MsgBox("Problem While Generating Barcode Image.", vbInformation, ResolveResString(100))
                                                    Exit Sub
                                                Else
                                                    If SaveBarCodeImage_singlelevelso_2DBARCODE(Ctlinvoice.Text, gstrUserMyDocPath, Mid(strBarcodeMsg, 3)) = False Then
                                                        MsgBox("Problem While saving Barcode Image.", vbInformation, ResolveResString(100))
                                                        Exit Sub
                                                    Else
                                                        mP_Connection.Execute(" UPDATE T SET T.BARCODEIMAGE =SC.BARCODEIMAGE FROM SALESCHALLAN_DTL SC,TMP_INVOICEPRINT  T " &
                                                                                       " WHERE SC.UNIT_CODE = T.UNIT_CODE AND SC.DOC_NO =T.DOC_NO AND SC.UNIT_CODE='" & gstrUNITID & "' AND " &
                                                                                       " SC.DOC_NO='" & Ctlinvoice.Text.Trim & "' AND T.IP_ADDRESS='" & gstrIpaddressWinSck & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                                    End If
                                                End If
                                                'Added by prit for TATA QR BARCODE on 15 Feb 2021
                                                'If AllowBarCodePrinting(strAccountCode) Then
                                            ElseIf GetPrintMethod(strAccountCode).ToUpper() = "TATA" Then  '' On Print button for temporary No
                                                Dim StrTATAsuffix As String
                                                StrTATAsuffix = gstrUNITID & Ctlinvoice.Text.ToString.Trim & DateTime.Now.ToString("ddMMyyHHmmssfff")

                                                strBarcodeMsg = ObjBarcodeHMI.GenerateQRBarCodeForTATAMotors(gstrUserMyDocPath, True, Trim(Ctlinvoice.Text), 0, gstrCONNECTIONSTRING, StrTATAsuffix)
                                                If VB.Left(strBarcodeMsg, 1) <> "Y" Then
                                                    CustomRollbackTrans()
                                                    MsgBox("Problem While Generating Barcode Image.", vbInformation, ResolveResString(100))
                                                    Exit Sub
                                                Else
                                                    strBarcodeMsg_paratemeter = Mid(strBarcodeMsg, 3)
                                                    If Not SaveQRBarCodeImageTATA(Trim(Ctlinvoice.Text), 0, strBarcodeMsg_paratemeter, StrTATAsuffix) Then
                                                        CustomRollbackTrans()
                                                        MsgBox("Problem While Saving Barcode Image.", vbInformation, ResolveResString(100))
                                                        Exit Sub
                                                    End If
                                                End If

                                                '25 oct 2017
                                            Else
                                                rsGENERATEBARCODE = New ClsResultSetDB_Invoice
                                                rsGENERATEBARCODE.GetResult("SELECT PRINT_METHOD,ITEM_CODE,CUST_ITEM_CODE FROM SALES_DTL SD ,SALESCHALLAN_DTL SC ,CUSTOMER_MST C WHERE C.UNIT_CODE=SC.UNIT_CODE AND C.CUSTOMER_CODE=SC.ACCOUNT_CODE AND SC.UNIT_CODE=SD.UNIT_CODE AND " &
                                                                                    " SC.DOC_NO=SD.DOC_NO  AND SC.UNIT_CODE='" & gstrUNITID & "' AND  " &
                                                                                    " SC.DOC_NO= " & Ctlinvoice.Text & " ORDER BY CUST_ITEM_CODE ")

                                                intTotalNoofitemsinInvoices = rsGENERATEBARCODE.GetNoRows
                                                rsGENERATEBARCODE.MoveFirst()
                                                For intRow = 1 To intTotalNoofitemsinInvoices
                                                    strBarcodeMsg = ObjBarcodeHMI.GenerateBarCode_LINELEVEL_SALESORDER(gstrUserMyDocPath, mInvNo, rsGENERATEBARCODE.GetValue("PRINT_METHOD").ToString, rsGENERATEBARCODE.GetValue("Cust_Item_Code").ToString, rsGENERATEBARCODE.GetValue("Item_Code").ToString, True, Trim(Ctlinvoice.Text), gstrCONNECTIONSTRING)
                                                    If VB.Left(strBarcodeMsg, 1) <> "Y" Then
                                                        MsgBox("Problem While Generating Barcode Image.", vbInformation, ResolveResString(100))
                                                        Exit Sub
                                                    End If
                                                    If SaveBarCodeImage_singlelevelso(Ctlinvoice.Text, rsGENERATEBARCODE.GetValue("Cust_Item_Code").ToString, rsGENERATEBARCODE.GetValue("Item_Code").ToString, gstrUserMyDocPath, intRow) = False Then
                                                        MsgBox("Problem While saving Barcode Image.", vbInformation, ResolveResString(100))
                                                        Exit Sub
                                                    End If
                                                    rsGENERATEBARCODE.MoveNext()
                                                Next

                                                rsGENERATEBARCODE.ResultSetClose()
                                                rsGENERATEBARCODE = Nothing
                                            End If

                                        End If
                                    Else
                                        'Added by prit for TATA QR BARCODE on 15 Feb 2021
                                        'If AllowBarCodePrinting(strAccountCode) Then
                                        If GetPrintMethod(strAccountCode).ToUpper() = "TATA" Then '' On print button
                                            Dim StrTATAsuffix As String
                                            StrTATAsuffix = gstrUNITID & Ctlinvoice.Text.ToString.Trim & DateTime.Now.ToString("ddMMyyHHmmssfff")
                                            strBarcodeMsg = ObjBarcodeHMI.GenerateQRBarCodeForTATAMotors(gstrUserMyDocPath, True, Trim(Ctlinvoice.Text), 0, gstrCONNECTIONSTRING, StrTATAsuffix)
                                            If VB.Left(strBarcodeMsg, 1) <> "Y" Then
                                                CustomRollbackTrans()
                                                MsgBox("Problem While Generating Barcode Image.", vbInformation, ResolveResString(100))
                                                Exit Sub
                                            Else
                                                strBarcodeMsg_paratemeter = Mid(strBarcodeMsg, 3)
                                                If Not SaveQRBarCodeImageTATA(Trim(Ctlinvoice.Text), 0, strBarcodeMsg_paratemeter, StrTATAsuffix) Then
                                                    CustomRollbackTrans()
                                                    MsgBox("Problem While Saving Barcode Image.", vbInformation, ResolveResString(100))
                                                    Exit Sub
                                                End If
                                            End If
                                        End If
                                    End If
                                End If

                                'HILEX A4 BARCODE'
                                If ConfirmWindow(10344, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                                    Dim strtime As String = GetServerDateTime()
                                    blninvoicelockYES = True
                                    'Added for Issue ID eMpro-20090226-27911 Starts
                                    Call CheckInvoiceExistInFinance(mInvNo)
                                    'Added for Issue ID eMpro-20090226-27911 Ends

                                    mP_Connection.BeginTrans()
                                    mP_Connection.Execute("set Dateformat 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

                                    If AllowBarCodePrinting(strAccountCode) Then
                                        If GetPrintMethod(strAccountCode).ToUpper() = "TATA" Then '' On print button for locking
                                            Dim StrTATAsuffix As String
                                            StrTATAsuffix = gstrUNITID & mInvNo.ToString.Trim & DateTime.Now.ToString("ddMMyyHHmmssfff")
                                            strBarcodeMsg = ObjBarcodeHMI.GenerateQRBarCodeForTATAMotors(gstrUserMyDocPath, True, Trim(Ctlinvoice.Text), mInvNo, gstrCONNECTIONSTRING, StrTATAsuffix)
                                            If VB.Left(strBarcodeMsg, 1) <> "Y" Then
                                                CustomRollbackTrans()
                                                MsgBox("Problem While Generating Barcode Image.", vbInformation, ResolveResString(100))
                                                Exit Sub
                                            Else
                                                strBarcodeMsg_paratemeter = Mid(strBarcodeMsg, 3)
                                                If Not SaveQRBarCodeImageTATA(Trim(Ctlinvoice.Text), mInvNo, strBarcodeMsg_paratemeter, StrTATAsuffix) Then
                                                    CustomRollbackTrans()
                                                    MsgBox("Problem While Saving Barcode Image.", vbInformation, ResolveResString(100))
                                                    Exit Sub
                                                End If
                                            End If
                                        End If
                                    End If

                                    'mP_Connection.Execute strsaledetails

                                    'Added for Issue ID 21105 Starts
                                    'If InvAgstBarCode() = True And mstrFGDomestic = "1" Then
                                    'If InvAgstBarCode() = True And (mstrFGDomestic = "1" Or mstrFGDomestic = "2") Then
                                    '    'issue id 10192547
                                    '    If BarCodeTracking(Trim(Ctlinvoice.Text), "LOCK") = True Then
                                    '        mP_Connection.Execute(mstrupdateBarBondedStockQty, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                    '        mP_Connection.Execute(mstrupdateBarBondedStockFlag, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                    '    ElseIf mblnQuantityCheck = False Then
                                    '        CustomRollbackTrans
                                    '        Exit Sub
                                    '    End If
                                    'End If
                                    If InvAgstBarCode() = True And (mstrFGDomestic = "1" Or mstrFGDomestic = "2") Then
                                        'issue id 10192547
                                        If BarCodeTracking(Trim(Ctlinvoice.Text), "LOCK") = True Then
                                            If UCase(Trim(CmbCategory.Text)) = "RAW MATERIAL" Or UCase(Trim(CmbCategory.Text)) = "INPUTS" Or UCase(Trim(CmbCategory.Text)) = "COMPONENTS" Or UCase(Trim(CmbCategory.Text)) = "SUB ASSEMBLY" Then
                                                Call updateBARcrossrefence_Invoicequantity(Ctlinvoice.Text)
                                            Else
                                                If mstrupdateBarBondedStockQty <> "" Then
                                                    mP_Connection.Execute(mstrupdateBarBondedStockQty, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                                End If
                                                If mstrupdateBarBondedStockFlag <> "" Then
                                                    mP_Connection.Execute(mstrupdateBarBondedStockFlag, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                                End If
                                            End If

                                            'mP_Connection.Execute(mstrupdateBarBondedStockQty, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                            'mP_Connection.Execute(mstrupdateBarBondedStockFlag, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                        ElseIf mblnQuantityCheck = False Then
                                            CustomRollbackTrans()
                                            Exit Sub
                                        End If
                                    End If

                                    ''Added for Issue ID 21105 Starts
                                    'FTS RELATED CHANGES
                                    If DataExist("Select top 1 1  from saleschallan_dtl WHERE UNIT_CODE='" + gstrUNITID + "' and doc_no='" & Ctlinvoice.Text & "' and fts_item=1") Then
                                        oCmd = New ADODB.Command
                                        With oCmd
                                            .let_ActiveConnection(mP_Connection)
                                            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                                            .CommandText = "USP_FTS_PRODUCTIONSLIP"
                                            .Parameters.Append(.CreateParameter("@Unit_Code", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                                            .Parameters.Append(.CreateParameter("@Doc_No", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, , Trim(Ctlinvoice.Text)))
                                            .Parameters.Append(.CreateParameter("@ErrMsg", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamOutput, 1000))
                                            .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                        End With
                                        If oCmd.Parameters(oCmd.Parameters.Count - 1).Value <> "" Then
                                            CustomRollbackTrans()
                                            MsgBox(oCmd.Parameters(oCmd.Parameters.Count - 1).Value, MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100))
                                            oCmd = Nothing
                                            Exit Sub
                                        End If
                                        oCmd = Nothing
                                        'mP_Connection.Execute("update Saleschallan_dtl set Location_code= '" & strStockLocation & "' WHERE UNIT_CODE='" + gstrUNITID + "' AND  Doc_no  " & Ctlinvoice.Text, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                        mP_Connection.Execute("update Saleschallan_dtl set from_location=  '" & strsaleconfLocation & "' WHERE UNIT_CODE='" + gstrUNITID + "' AND  Doc_no = " & Ctlinvoice.Text, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                        mP_Connection.Execute("update FTS_LABEL_ISSUE set AUTHORISED= 1 ,DOC_NO='" & mInvNo & "' WHERE UNIT_CODE='" + gstrUNITID + "' AND DOC_TYPE=9999  AND  Doc_no = '" & Ctlinvoice.Text & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                        mP_Connection.Execute("update FTS_LABEL_RELATIONSHIP set Ref_DOC_NO='" & mInvNo & "' WHERE UNIT_CODE='" + gstrUNITID + "' AND Ref_Doc_no = '" & Ctlinvoice.Text & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                        mP_Connection.Execute("update FTS_FG_PICKLIST_INV_DTL set INVOICENO='" & mInvNo & "' WHERE UNIT_CODE='" + gstrUNITID + "' AND TMP_INVOICENO = '" & Ctlinvoice.Text & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

                                    End If
                                    'FTS RELATED CHANGES

                                    mP_Connection.Execute("INSERT INTO INV_ERROR_DTL(QUERY,UNIT_CODE,INVOICENO) VALUES('" & Replace(saleschallan, "'", "") & "','" & gstrUNITID & "','" & mInvNo & "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                    mP_Connection.Execute(saleschallan, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                    'hilex
                                    Call UPDATETRANFERINVOICE_HILEX(Ctlinvoice.Text, mInvNo, "L")
                                    'hilex
                                    If Len(Trim(mstrExcisePriorityUpdationString)) > 0 Then
                                        mP_Connection.Execute("update Saleschallan_dtl set Excise_type = '" & mstrExcisePriorityUpdationString & "' WHERE UNIT_CODE='" + gstrUNITID + "' AND  Doc_no = " & Ctlinvoice.Text, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                    End If
                                    mstrInvRejSQL = ""
                                    If CBool(Find_Value("Select REJINV_Tracking from Sales_Parameter WHERE UNIT_CODE='" + gstrUNITID + "'")) = True Then
                                        mstrInvRejSQL = "Update MKT_INVREJ_DTL Set Invoice_No='" & mInvNo & "' WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_No='" & Trim(Ctlinvoice.Text) & "'"
                                    End If

                                    If Len(Trim(mstrInvRejSQL)) <> 0 Then
                                        mP_Connection.Execute(mstrInvRejSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                    End If
                                    If optInvYes(0).Checked = True Then
                                        If DataExist("SELECT TOP 1 1 FROM SALES_PARAMETER WHERE INVOICE_LOCKING_ENTRY_SAMEDATE=1  and UNIT_CODE = '" & gstrUNITID & "'") Then
                                            mP_Connection.Execute("update Saleschallan_dtl set invoice_date= Convert(varchar(12), getdate(), 106) WHERE UNIT_CODE='" + gstrUNITID + "' AND  Doc_no = " & mInvNo, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                        End If
                                    End If

                                    ' Code add by Sourabh
                                    If Trim(UCase(Me.txtUnitCode.Text)) = "SML" Or Trim(UCase(Me.txtUnitCode.Text)) = "SMT" Then
                                        If CBool(Find_Value("Select isnull(MultiUnitInvoice,0) from sales_parameter WHERE UNIT_CODE='" + gstrUNITID + "'")) Then
                                            mP_Connection.Execute("Set XACT_ABORT  on", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                            blnFlagTrans = True
                                            With cmdObject
                                                .let_ActiveConnection(mP_Connection)
                                                .CommandTimeout = 0
                                                .CommandType = ADODB.CommandTypeEnum.adCmdText

                                                If Me.txtUnitCode.Text = "SML" Then
                                                    strVar = Replace(salesconf, "saleconf", "SUMITLIVE2008.DBO.Saleconf")
                                                    strVar = Replace(strVar, "Location_Code='SML'", "Location_Code='SMT'")
                                                Else
                                                    strVar = Replace(salesconf, "saleconf", "SMIEL_FIN2008.DBO.Saleconf")
                                                    strVar = Replace(strVar, "Location_Code='SMT'", "Location_Code='SML'")
                                                End If

                                                .CommandText = strVar
                                                .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                            End With
                                            'UPGRADE_NOTE: Object cmdObject may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
                                            cmdObject = Nothing
                                            mP_Connection.Execute("Set XACT_ABORT  Off", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                            blnFlagTrans = False
                                        End If
                                    End If
                                    'Code end here


                                    mP_Connection.Execute(salesconf, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                    If updatePOflag = True Then
                                        mP_Connection.Execute(strupdatecustodtdtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                    End If

                                    '''***** changes done By ashutosh on 13-06-2006, issue Id:18099

                                    If updatestockflag = True Then
                                        'UPGRADE_WARNING: Couldn't resolve default property of object varTmp1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                        varTmp1 = Split(strSelectItmbalmst, "»")

                                        varTmp = Split(strupdateitbalmst, "»")

                                        For i = 0 To (intRow - 1)
                                            'UPGRADE_WARNING: Couldn't resolve default property of object varTmp1(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                            rsItemBalance = New ClsResultSetDB_Invoice
                                            rsItemBalance.GetResult(varTmp1(i))
                                            If rsItemBalance.GetNoRows > 0 Then
                                                'UPGRADE_WARNING: Couldn't resolve default property of object rsItemBalance.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                                dblTmpItembal = Val(rsItemBalance.GetValue("Cur_Bal"))
                                            End If
                                            rsItemBalance.ResultSetClose()
                                            rsItemBalance = Nothing


                                            mP_Connection.Execute(varTmp(i), , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

                                            'UPGRADE_WARNING: Couldn't resolve default property of object varTmp1(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                            rsItemBalance = New ClsResultSetDB_Invoice
                                            rsItemBalance.GetResult(varTmp1(i))
                                            If rsItemBalance.GetNoRows > 0 Then
                                                'UPGRADE_WARNING: Couldn't resolve default property of object rsItemBalance.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                                                dblFinalItembal = Val(rsItemBalance.GetValue("Cur_Bal"))
                                            End If
                                            rsItemBalance = New ClsResultSetDB_Invoice

                                            ''''                           If dblFinalItembal = dblTmpItembal Then
                                            ''''                               mP_Connection.Execute varTmp(i)
                                            ''''                           End If
                                        Next i
                                        '''mP_Connection.Execute strupdateitbalmst
                                    End If

                                    If UCase(cmbInvType.Text) = "JOBWORK INVOICE" Then
                                        mP_Connection.Execute("SET DATEFORMAT 'DMY'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                        mP_Connection.Execute(mstrAnnex, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                    End If
                                    If UCase(Me.lbldescription.Text) = "REJ" Then
                                        If Len(Trim(mCust_Ref)) > 0 Then
                                            mP_Connection.Execute(strupdateGrinhdr, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                        End If
                                    End If
                                    If Len(Trim(strBatchQuery)) > 0 Then
                                        mP_Connection.Execute(strBatchQuery, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                        mP_Connection.Execute("Update ItemBatch_dtl Set Doc_no = '" & Trim(CStr(mInvNo)) & "' Where Doc_no = '" & Trim(Me.Ctlinvoice.Text) & "' and Doc_Type = 9999 ", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                    End If

                                    ' Code add by sourabh on 22 nov 2005 against issue id 2004-04-003-16261
                                    If mblnInvoiceAgainstBarCode = True Then
                                        mP_Connection.Execute("Update tbl_BarCode_MaterialOut Set Invoice_No = '" & Trim(CStr(mInvNo)) & "' WHERE UNIT_CODE='" + gstrUNITID + "' AND  invoice_no = '" & Trim(Me.Ctlinvoice.Text) & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                    End If

                                    '--------------updation of Schedules-------------------------------------------------------
                                    '-------JS on 05/09/2004-------------------------------------------------------------------
                                    'Added for Issue Id 21840 Starts
                                    If blnInvoiceAgainstMultipleSO = False Then
                                        If ((UCase(Trim(cmbInvType.Text)) = "TRANSFER INVOICE") And (UCase(CStr((Trim(CmbCategory.Text)) = "FINISHED GOODS")))) Then
                                            blnDSTracking = CBool(Find_Value("SELECT isnull(DSWiseTracking,0)as DSWiseTracking FROM sales_parameter Where Unit_code='" & gstrUNITID & "'  "))
                                            If blnDSTracking = True And AllowTextFileGeneration_SMIIEL(strAccountCode) = True And CBool(Find_Value("SELECT SMIEL_FTP_INVOICE FROM SALES_PARAMETER  where   unit_code='" & gstrUNITID & "'")) = True Then
                                                If Not UpdateMktSchedule() Then
                                                    CustomRollbackTrans()
                                                    Exit Sub
                                                End If
                                            End If
                                        End If
                                    End If
                                    'Added for Issue Id 21840 Ends
                                    '--------------updation of Schedules-------------------------------------------------------
                                    '10237233
                                    mstrupdateASNdtl = ""
                                    mstrupdateASNCumFig = ""
                                    If AllowASNTextFileGeneration(Trim(strAccountCode)) = True Then
                                        mP_Connection.Execute("UPDATE MKT_ASN_INVDTL SET DOC_NO=" & Trim(mInvNo) & " where dOC_NO=" & Trim(Ctlinvoice.Text) & " and unit_code='" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                        If FordASNFileGeneration(mInvNo) = False Then
                                            CustomRollbackTrans()
                                            Exit Sub
                                        Else
                                            If Len(mstrupdateASNdtl) > 0 Then
                                                mP_Connection.Execute(mstrupdateASNdtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                                mP_Connection.Execute(mstrupdateASNCumFig, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                            End If
                                        End If
                                    End If
                                    '10237233

                                    'Accounts Posting is done here
                                    If mblnpostinfin = True Then
                                        objDrCr = New prj_DrCrNote.cls_DrCrNote(GetServerDate)
                                        If UCase(Trim(cmbInvType.Text)) <> "REJECTION" Then
                                            strRetVal = objDrCr.SetARInvoiceDocument(gstrUNITID, mstrMasterString, mstrDetailString, prj_GLTransactions.cls_GLTransactions.udtOperationType.optInsert, gstrCONNECTIONSTRING, "", mP_Connection)
                                        Else
                                            '''Changes done by ashutosh on 12 jun 2007 , Issue Id:19934
                                            If RejInvOptionalPostingFlag() = True Then
                                                If (chkAcceffects.CheckState = System.Windows.Forms.CheckState.Checked) And (chkAcceffects.Enabled = True) Then
                                                    strRetVal = "Y"
                                                Else
                                                    prj_DocGenerator.cls_DocumentGenerator.gbln_AR_AP_Dr_Cr_Doc_Sub_Category = "DR"
                                                    strRetVal = objDrCr.SetAPDocument(gstrUNITID, mstrMasterString, mstrDetailString, prj_GLTransactions.cls_GLTransactions.udtOperationType.optInsert, gstrCONNECTIONSTRING, "", mP_Connection)
                                                    prj_DocGenerator.cls_DocumentGenerator.gbln_AR_AP_Dr_Cr_Doc_Sub_Category = ""
                                                End If
                                            Else
                                                prj_DocGenerator.cls_DocumentGenerator.gbln_AR_AP_Dr_Cr_Doc_Sub_Category = "DR"
                                                strRetVal = objDrCr.SetAPDocument(gstrUNITID, mstrMasterString, mstrDetailString, prj_GLTransactions.cls_GLTransactions.udtOperationType.optInsert, gstrCONNECTIONSTRING, "", mP_Connection)
                                                prj_DocGenerator.cls_DocumentGenerator.gbln_AR_AP_Dr_Cr_Doc_Sub_Category = ""
                                            End If
                                            '''Changes for Issue Id:19934 end here.
                                        End If
                                    Else
                                        strRetVal = "Y"
                                    End If

                                    strRetVal = CheckString(strRetVal)
                                    If Not strRetVal = "Y" Then
                                        MsgBox(strRetVal, MsgBoxStyle.Information, "eMPro")
                                        CustomRollbackTrans()
                                        Exit Sub
                                    Else
                                        Frm.ShowPrintButton = False
                                        If mblnCSMspecificreport = False Then
                                            Frm.show()
                                        End If

                                        CheckASNExist(Me.Ctlinvoice.Text)
                                        If AllowASNPrinting(strAccountCode) = True Then
                                            If mblnASNExist = True Then
                                                mP_Connection.Execute("Update CreatedASN Set ASN_NO='" & Trim$(txtASNNumber.Text) & "' where doc_no='" & Trim$(Me.Ctlinvoice.Text) & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                            Else
                                                mP_Connection.Execute("Insert into CreatedASN values('" & Trim$(Me.Ctlinvoice.Text) & "','" & Trim$(txtASNNumber.Text) & "',getdate(),'" & mP_User & "',getdate(),'" & dtpASNDatetime.Value & "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                            End If
                                        End If

                                        blnDSTracking = CBool(Find_Value("SELECT isnull(DSWiseTracking,0)as DSWiseTracking FROM sales_parameter where unit_code='" & gstrUNITID & "'"))
                                        If ((UCase(Trim(cmbInvType.Text)) = "TRANSFER INVOICE") And (UCase(CStr((Trim(CmbCategory.Text)) = "FINISHED GOODS")))) Then
                                            If blnDSTracking = True And AllowTextFileGeneration_SMIIEL(strAccountCode) = True And CBool(Find_Value("SELECT SMIEL_FTP_INVOICE FROM SALES_PARAMETER WHERE UNIT_CODE='" & gstrUNITID & "'")) = True Then
                                                If CheckInvoices(mInvNo, strAccountCode) = False Then
                                                    CustomRollbackTrans()
                                                    Exit Sub
                                                End If
                                            End If
                                        End If

                                        '''Changes done by Ashutosh on 11 May 2007, Issue Id:19934
                                        If VerifyInvPostingFlag() = True Then
                                            If InvoicePostingFlag() = True Then
                                                If UCase(Trim(cmbInvType.Text)) <> "REJECTION" Then

                                                    rsInvDataInFin = New ClsResultSetDB_Invoice
                                                    rsInvoiceDataInFin = New ClsResultSetDB_Invoice
                                                    rsInvDataInFin.GetResult("Select top 1 * from ar_docMaster WHERE DocM_unit='" + gstrUNITID + "' AND  docm_vono='" & Trim(CStr(mInvNo)) & "'")
                                                    If rsInvDataInFin.GetNoRows > 0 Then
                                                        rsInvoiceDataInFin.GetResult("Select top 1 * from saleschallan_dtl where  unit_code='" + gstrUNITID + "' AND doc_no='" & Trim(CStr(mInvNo)) & "'")
                                                        If rsInvoiceDataInFin.GetNoRows > 0 Then
                                                            mP_Connection.CommitTrans()
                                                            If mblncustomerspecificreport = True And GetPrintMethod(strAccountCode).ToUpper() <> "TATA" Then
                                                                strSql = "{SalesChallan_Dtl.Location_Code}='" & Trim(txtUnitCode.Text) & "'  and {SalesChallan_Dtl.Unit_Code}='" & gstrUNITID & "'  and {SalesChallan_Dtl.Doc_No} =" & mInvNo & " and {SalesChallan_Dtl.Invoice_Type}"
                                                                strSql = strSql & " = '" & Trim(Me.lbldescription.Text) & "'  and {SalesChallan_Dtl.Sub_Category} = '" & Trim(Me.lblcategory.Text) & "'"
                                                                RdAddSold.DataDefinition.RecordSelectionFormula = strSql
                                                                Frm.Show()
                                                            Else
                                                                If mblncustomerspecificreport = True And GetPrintMethod(strAccountCode).ToUpper() = "TATA" Then
                                                                    Frm.Show()
                                                                End If
                                                            End If
                                                        Else
                                                            MsgBox("Invoice not posted to Invoice Table. Try Again!!! ", MsgBoxStyle.Information, ResolveResString(100))
                                                            CustomRollbackTrans()
                                                            rsInvDataInFin.ResultSetClose()
                                                            rsInvoiceDataInFin.ResultSetClose()
                                                            Exit Sub
                                                        End If
                                                    Else
                                                        MsgBox("Invoice not posted to accounts. Try Again!!! ", MsgBoxStyle.Information, ResolveResString(100))
                                                        CustomRollbackTrans()
                                                        rsInvDataInFin.ResultSetClose()
                                                        Exit Sub
                                                    End If
                                                    rsInvDataInFin.ResultSetClose()
                                                Else
                                                    If (chkAcceffects.CheckState <> System.Windows.Forms.CheckState.Checked) Then
                                                        rsInvDataInFin = New ClsResultSetDB_Invoice
                                                        rsInvDataInFin.GetResult("Select top 1 apdocm_venprdocNo from ap_docMaster WHERE apDocM_unit='" + gstrUNITID + "' AND apdocm_venprdocNo='" & Trim(CStr(mInvNo)) & "'")
                                                        If rsInvDataInFin.GetNoRows > 0 Then
                                                            mP_Connection.CommitTrans()
                                                        Else
                                                            MsgBox("Invoice not posted to accounts. Try Again!!! ", MsgBoxStyle.Information, ResolveResString(100))
                                                            CustomRollbackTrans()
                                                            rsInvDataInFin.ResultSetClose()
                                                            Exit Sub
                                                        End If
                                                        rsInvDataInFin.ResultSetClose()
                                                    Else
                                                        mP_Connection.CommitTrans()
                                                    End If
                                                End If
                                            Else
                                                mP_Connection.CommitTrans()
                                            End If
                                        Else
                                            mP_Connection.CommitTrans()
                                        End If

                                        If gBlnWIPFGProcess AndAlso WIP_FG_Customer(strAccountCode) Then
                                            If mblnskipdacinvoicebincheck = False Then
                                                Dim svcWIPFGInv As New eMPRO_WIP_FG_INVOICE_COMPLETION
                                                Dim strResult As String = String.Empty
                                                If DataExist("SELECT TOP 1 1 FROM WIP_FG_PICKLIST_INV_DTL  WHERE  UNIT_CODE='" & gstrUNITID & "' AND TMP_INVOICENO =" & Trim(Ctlinvoice.Text)) Then
                                                    mP_Connection.Execute("UPDATE WIP_FG_PICKLIST_INV_DTL  SET INVOICENO = '" & mInvNo & "' WHERE UNIT_CODE='" + gstrUNITID + "' AND   TMP_INVOICENO = " & Trim(Ctlinvoice.Text), , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                                    strResult = svcWIPFGInv.ProcessInvoice(gstrUNITID, "R", Trim(Ctlinvoice.Text))
                                                    If strResult.ToUpper() <> "SUCCESS|" Then
                                                        WIP_FG_SAVE_EXCEPTION_LOG(Convert.ToInt32(mInvNo), "svcWIPFGInv.ProcessInvoice", strResult, gstrUNITID, "R", Trim(Ctlinvoice.Text), "")
                                                    End If
                                                End If
                                                svcWIPFGInv.Dispose()
                                            Else
                                                '10623079
                                                mP_Connection.Execute("UPDATE WIP_FG_PICKLIST_INV_DTL  SET INVOICENO = '" & mInvNo & "' WHERE UNIT_CODE='" + gstrUNITID + "' AND   TMP_INVOICENO = " & Trim(Ctlinvoice.Text), , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                                '10623079
                                            End If

                                        End If

                                        '''Changes for Issue Id:19934 end here.
                                        If optInvYes(0).Checked = True And AllowASNPrinting(strAccountCode) = True Then
                                            If txtASNNumber.Text.Trim.Length > 0 Then
                                                mP_Connection.Execute("Update CreatedASN Set DOC_NO='" & Trim$(mInvNo) & "',Updatedon=getdate() where doc_no='" & Trim$(Me.Ctlinvoice.Text) & "'  and unit_code='" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                            End If
                                        End If

                                        'Changed by nisha on 05/09/2003
                                        If UCase(Trim(cmbInvType.Text)) = "REJECTION" Then
                                            If UCase(Trim(CmbCategory.Text)) = "REJECTION" Then
                                                '''Changes for Issue Id:19934 , on 12 Jun 2007.
                                                If RejInvOptionalPostingFlag() = True And chkAcceffects.CheckState = System.Windows.Forms.CheckState.Checked Then
                                                    '''Changes for Issue Id:19934 end here.
                                                    mP_Connection.Execute("update salesChallan_Dtl set RejectionPosting = 1 WHERE UNIT_CODE='" + gstrUNITID + "' AND  doc_no = " & mInvNo, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                                Else
                                                    mP_Connection.Execute("update salesChallan_Dtl set RejectionPosting = 0 WHERE UNIT_CODE='" + gstrUNITID + "' AND  doc_no = " & mInvNo, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                                End If
                                            End If
                                        End If
                                        'Changed by nisha on 05/09/2003 ends here
                                        Call Logging_Starting_End_Time("Invoice Locking ", strtime, "Saved", mInvNo)
                                        MsgBox("Invoice has been locked successfully with number " & mInvNo, MsgBoxStyle.Information, "eMPro")
                                        '                                    Write_In_Log_File(GetServerDateTime() & "  : Invoice Locked Successfully - Printer")
                                        'SATISH KESHAERWANI CHANGE
                                        If chktkmlbarcode.Checked = True Then
                                            Print_barcodelabel(mInvNo)
                                        End If
                                        'SATISH KESHAERWANI CHANGE

                                        '''***** Code added by Ashutosh on 24-12-2005,Issue Id:16685
                                        blnPrintActualInvFlag = True
                                        '''***** Changes on 24-12-2005 end here.

                                    End If
                                Else
                                    'Changed for Issue ID eMpro-20090216-27468 Starts(Unlocked invoice cannot be send to printer)
                                    MessageBox.Show("Please lock the invoice before printing", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                                    Ctlinvoice.Text = ""
                                    Exit Sub
                                    'Changed for Issue ID eMpro-20090216-27468 Ends(Unlocked invoice cannot be send to printer)
                                End If

                                If Trim(UCase(Me.txtUnitCode.Text)) = "SUN" Then
                                    If blnPrintActualInvFlag = False Then
                                        objInvoicePrint = New prj_InvoicePrinting.clsInvoicePrinting(gstrDateFormat)
                                        objInvoicePrint.Print_Invoice(gstrUNITID, True, (txtUnitCode.Text), (Ctlinvoice.Text), dtpRemoval.Text & " " & dtpRemovalTime.Value.Hour & ":" & dtpRemovalTime.Value.Minute)
                                        cmdPrint_Click(cmdPrint, New System.EventArgs())
                                    Else
                                        objInvoicePrint = New prj_InvoicePrinting.clsInvoicePrinting(gstrDateFormat)
                                        'Added for Issue ID eMpro-20080805-20745 Starts
                                        objInvoicePrint.mstrDSNforInoivcePrint = gstrDSNName
                                        'Added for Issue ID eMpro-20080805-20745 Ends
                                        objInvoicePrint.Print_Invoice(gstrUNITID, True, (txtUnitCode.Text), CStr(mInvNo), dtpRemoval.Value & " " & dtpRemovalTime.Hour & ":" & dtpRemovalTime.Minute)
                                        cmdPrint_Click(cmdPrint, New System.EventArgs())
                                    End If
                                End If

                                'Changed for Issue ID eMpro-20090216-27468 Starts(Only Locked invoice can send to printer)
                                If CBool(Find_Value("select TextPrinting from sales_parameter WHERE UNIT_CODE='" + gstrUNITID + "'")) Then
                                    objInvoicePrint = New prj_InvoicePrinting.clsInvoicePrinting(gstrDateFormat)
                                    objInvoicePrint.ConnectionString = gstrCONNECTIONSTRING 'mP_Connection.ConnectionString
                                    'Added for Issue ID eMpro-20080805-20745 Starts
                                    objInvoicePrint.mstrDSNforInoivcePrint = gstrDSNName
                                    'Added for Issue ID eMpro-20080805-20745 Ends
                                    objInvoicePrint.Connection()
                                    'shalini
                                    'objInvoicePrint.FileName = "c:\InvoicePrint.txt"
                                    'objInvoicePrint.BCFileName = "c:\BarCode.txt"
                                    objInvoicePrint.FileName = strCitrix_Inv_Pronting_Loc & "InvoicePrint.txt"
                                    objInvoicePrint.BCFileName = strCitrix_Inv_Pronting_Loc & "BarCode.txt"
                                    objInvoicePrint.CompanyName = gstrCOMPANY
                                    objInvoicePrint.Address1 = gstr_RGN_ADDRESS1
                                    objInvoicePrint.Address2 = gstr_RGN_ADDRESS2
                                    objInvoicePrint.Print_Invoice(gstrUNITID, True, (txtUnitCode.Text), CStr(mInvNo), dtpRemoval.Text & " " & dtpRemovalTime.Value.Hour & ":" & dtpRemovalTime.Value.Minute)
                                    '                                Write_In_Log_File(GetServerDateTime() & " : Text File Generated Successfully - Printer")
                                    cmdPrint_Click(cmdPrint, New System.EventArgs())

                                    Ctlinvoice.Text = ""
                                    '''***** Changes on 24-12-2005 end here.
                                Else
                                    If mblncustomerlevel_A4report_functionlity = True And mblnA4reports_invoicewise = True Then 'FOR A4 CUSTOMERS


                                        intNoCopies_A4reports_orignial = CInt(Find_Value("select isnull(MAX(SERIALNO),0) SERIALNO from A4CUSTOMER_INVOICEPRINTINGTAG  WHERE UNIT_CODE='" + gstrUNITID + "'AND CUSTOMER_CODE='" & strAccountCode & "' AND ORIGINAL_REPRINT='O'"))
                                        intNoCopies_A4reports_REPRINT = CInt(Find_Value("select isnull(MAX(SERIALNO),0) SERIALNO from A4CUSTOMER_INVOICEPRINTINGTAG  WHERE UNIT_CODE='" + gstrUNITID + "'AND CUSTOMER_CODE='" & strAccountCode & "' AND ORIGINAL_REPRINT='R'"))
                                        If optInvYes(0).Checked = True Then
                                            intMaxLoop = intNoCopies_A4reports_orignial
                                        Else
                                            If intNoCopies_A4reports > 1 Then
                                                intMaxLoop = intNoCopies_A4reports_REPRINT
                                            End If
                                        End If

                                        For intLoopCounter = 1 To intMaxLoop
                                            If optInvYes(0).Checked = True Then
                                                COPYNAME = Find_Value("Select TEXTHEADING FROM  A4CUSTOMER_INVOICEPRINTINGTAG WHERE UNIT_CODE='" + gstrUNITID + "' AND CUSTOMER_CODE='" & strAccountCode & "' AND ORIGINAL_REPRINT='O' AND SERIALNO=" & intLoopCounter)
                                                RdAddSold.DataDefinition.FormulaFields("CopyName").Text = "'" & COPYNAME & "'"
                                            Else
                                                COPYNAME = Find_Value("Select TEXTHEADING FROM  A4CUSTOMER_INVOICEPRINTINGTAG WHERE UNIT_CODE='" + gstrUNITID + "' AND CUSTOMER_CODE='" & strAccountCode & "' AND ORIGINAL_REPRINT='R' AND SERIALNO=" & intLoopCounter)
                                                RdAddSold.DataDefinition.FormulaFields("CopyName").Text = "'" & COPYNAME & "'"
                                            End If
                                            RdAddSold.DataDefinition.FormulaFields("InsuranceFlag").Text = "'" + CStr(mblnInsuranceFlag) + "'"
                                            Frm.SetReportDocument()
                                            'RdAddSold.PrintToPrinter(1, False, 0, 0)

                                            'If optInvYes(0).Checked = True Then
                                            '    dblewaymaxvalue = Find_Value("select total_amount from saleschallan_Dtl where unit_code='" + gstrUNITID + "' and doc_no =" & mInvNo)
                                            'Else
                                            '    dblewaymaxvalue = Find_Value("select total_amount from saleschallan_Dtl where unit_code='" + gstrUNITID + "' and doc_no =" & Ctlinvoice.Text)
                                            'End If

                                            If mblnEwaybill_Print = False Then
                                                RdAddSold.PrintToPrinter(1, False, 0, 0)
                                            Else
                                                'If dblewaymaxvalue <= mdblewaymaximumvalue Then
                                                '    RdAddSold.PrintToPrinter(1, False, 0, 0)
                                                If chkprintreprint.Checked = True And optInvYes(1).Checked = True Then
                                                    RdAddSold.PrintToPrinter(1, False, 0, 0)
                                                Else
                                                    If optInvYes(0).Checked = True Then
                                                        If Not DataExist("SELECT TOP 1 1 FROM FIRSTTIME_INVOICEPRINTING  WHERE UNIT_CODE = '" & gstrUNITID & "' and doc_no= " & Trim(Ctlinvoice.Text) & "") = True Then
                                                            mP_Connection.Execute("Insert into FIRSTTIME_INVOICEPRINTING(unit_code,doc_no,ent_dt,ent_userid) values('" & gstrUNITID & "','" & Trim$(Ctlinvoice.Text) & "',getdate(),'" & mP_User & "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                                        End If
                                                    End If
                                                End If
                                            End If

                                        Next
                                    Else ' FOR NON A4 CUSTROMERS

                                        If optInvYes(0).Checked = True Then
                                            intMaxLoop = intNoCopies
                                        Else
                                            If intNoCopies > 1 Then
                                                intMaxLoop = intNoCopies - 1
                                            Else
                                                intMaxLoop = intNoCopies
                                            End If
                                        End If

                                        For intLoopCounter = 1 To intMaxLoop
                                            Select Case intLoopCounter
                                                Case 1
                                                    If optInvYes(0).Checked = True Then
                                                        RdAddSold.DataDefinition.FormulaFields("CopyName").Text = "'ORIGINAL FOR BUYER'"
                                                    Else
                                                        RdAddSold.DataDefinition.FormulaFields("CopyName").Text = "'DUPLICATE FOR TRANSPORTER'"
                                                    End If
                                                Case 2
                                                    If optInvYes(0).Checked = True Then
                                                        RdAddSold.DataDefinition.FormulaFields("CopyName").Text = "'DUPLICATE FOR TRANSPORTER'"
                                                    Else
                                                        RdAddSold.DataDefinition.FormulaFields("CopyName").Text = "'TRIPLICATE FOR ASSESSEE'"
                                                    End If
                                                Case 3
                                                    If optInvYes(0).Checked = True Then
                                                        RdAddSold.DataDefinition.FormulaFields("CopyName").Text = "'TRIPLICATE FOR ASSESSEE'"
                                                    Else
                                                        RdAddSold.DataDefinition.FormulaFields("CopyName").Text = "'EXTRA COPY'"
                                                    End If
                                                Case Is >= 4
                                                    RdAddSold.DataDefinition.FormulaFields("CopyName").Text = "'EXTRA COPY'"
                                            End Select
                                            RdAddSold.DataDefinition.FormulaFields("InsuranceFlag").Text = "'" + CStr(mblnInsuranceFlag) + "'"
                                            Frm.SetReportDocument()
                                            'RdAddSold.PrintToPrinter(1, False, 0, 0)
                                            If mblnEwaybill_Print = False Then
                                                RdAddSold.PrintToPrinter(1, False, 0, 0)
                                            Else
                                                'If dblewaymaxvalue <= mdblewaymaximumvalue Then
                                                '    RdAddSold.PrintToPrinter(1, False, 0, 0)
                                                If chkprintreprint.Checked = True And optInvYes(1).Checked = True Then
                                                    RdAddSold.PrintToPrinter(1, False, 0, 0)
                                                Else
                                                    If optInvYes(0).Checked = True Then
                                                        If Not DataExist("SELECT TOP 1 1 FROM FIRSTTIME_INVOICEPRINTING  WHERE UNIT_CODE = '" & gstrUNITID & "' and doc_no= " & Trim(Ctlinvoice.Text) & "") = True Then
                                                            mP_Connection.Execute("Insert into FIRSTTIME_INVOICEPRINTING(unit_code,doc_no,ent_dt,ent_userid) values('" & gstrUNITID & "','" & Trim$(Ctlinvoice.Text) & "',getdate(),'" & mP_User & "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                                        End If
                                                    End If
                                                End If
                                            End If

                                        Next
                                    End If
                                End If
                                'Changed for Issue ID eMpro-20090216-27468 Ends(First Lock the invoice then send to printer)

                            Else
                                'Changed for Issue ID eMpro-20090216-27468 Starts(If Invoice is already then send it to printer)
                                If CBool(Find_Value("select TextPrinting from sales_parameter WHERE UNIT_CODE='" + gstrUNITID + "'")) Then
                                    objInvoicePrint = New prj_InvoicePrinting.clsInvoicePrinting(gstrDateFormat)
                                    objInvoicePrint.ConnectionString = gstrCONNECTIONSTRING 'mP_Connection.ConnectionString
                                    'Added for Issue ID eMpro-20080805-20745 Starts
                                    objInvoicePrint.mstrDSNforInoivcePrint = gstrDSNName
                                    'Added for Issue ID eMpro-20080805-20745 Ends
                                    objInvoicePrint.Connection()
                                    'shalini
                                    'objInvoicePrint.FileName = "c:\InvoicePrint.txt"
                                    'objInvoicePrint.BCFileName = "c:\BarCode.txt"
                                    objInvoicePrint.FileName = strCitrix_Inv_Pronting_Loc & "InvoicePrint.txt"
                                    objInvoicePrint.BCFileName = strCitrix_Inv_Pronting_Loc & "BarCode.txt"
                                    objInvoicePrint.CompanyName = gstrCOMPANY
                                    objInvoicePrint.Address1 = gstr_RGN_ADDRESS1
                                    objInvoicePrint.Address2 = gstr_RGN_ADDRESS2
                                    objInvoicePrint.Print_Invoice(gstrUNITID, True, (txtUnitCode.Text), (Ctlinvoice.Text), dtpRemoval.Text & " " & dtpRemovalTime.Value.Hour & ":" & dtpRemovalTime.Value.Minute)
                                    '                                Write_In_Log_File(GetServerDateTime() & " : Text File Generated Successfully - Printer 2")
                                    If Not cmdPrint.Enabled Then
                                        cmdPrint.Enabled = True
                                    End If
                                    cmdPrint_Click(cmdPrint, New System.EventArgs())
                                Else
                                    'hilex change

                                    'Added by prit for TATA QR BARCODE on 15 Feb 2021  
                                    If AllowBarCodePrinting(strAccountCode) Then  '' Reprint + Print option
                                        If GetPrintMethod(strAccountCode).ToUpper() = "TATA" Then
                                            Dim StrTATAsuffix As String
                                            StrTATAsuffix = gstrUNITID & Ctlinvoice.Text.ToString.Trim & DateTime.Now.ToString("ddMMyyHHmmssfff")
                                            strBarcodeMsg = ObjBarcodeHMI.GenerateQRBarCodeForTATAMotors(gstrUserMyDocPath, True, Trim(Ctlinvoice.Text), 0, gstrCONNECTIONSTRING, StrTATAsuffix)
                                            If VB.Left(strBarcodeMsg, 1) <> "Y" Then
                                                CustomRollbackTrans()
                                                MsgBox("Problem While Generating Barcode Image.", vbInformation, ResolveResString(100))
                                                Exit Sub
                                            Else
                                                strBarcodeMsg_paratemeter = Mid(strBarcodeMsg, 3)
                                                If Not SaveQRBarCodeImageTATA(Trim(Ctlinvoice.Text), 0, strBarcodeMsg_paratemeter, StrTATAsuffix) Then
                                                    CustomRollbackTrans()
                                                    MsgBox("Problem While Saving Barcode Image.", vbInformation, ResolveResString(100))
                                                    Exit Sub
                                                End If
                                            End If
                                        End If
                                    End If

                                    If mblncustomerlevel_A4report_functionlity = True And mblnA4reports_invoicewise = True Then 'FOR A4 CUSTOMERS 
                                        strAccountCode = Find_Value("select account_code from saleschallan_dtl where UNIT_CODE = '" & gstrUNITID & "' and doc_no='" & Trim(Me.Ctlinvoice.Text) & "'")
                                        intNoCopies_A4reports_orignial = CInt(Find_Value("select isnull(MAX(SERIALNO),0) SERIALNO from A4CUSTOMER_INVOICEPRINTINGTAG  WHERE UNIT_CODE='" + gstrUNITID + "'AND CUSTOMER_CODE='" & strAccountCode & "' AND ORIGINAL_REPRINT='O'"))
                                        intNoCopies_A4reports_REPRINT = CInt(Find_Value("select isnull(MAX(SERIALNO),0) SERIALNO from A4CUSTOMER_INVOICEPRINTINGTAG  WHERE UNIT_CODE='" + gstrUNITID + "'AND CUSTOMER_CODE='" & strAccountCode & "' AND ORIGINAL_REPRINT='R'"))
                                        If optInvYes(0).Checked = True Then
                                            intMaxLoop = intNoCopies_A4reports_orignial
                                        Else
                                            If intNoCopies_A4reports > 1 Then
                                                intMaxLoop = intNoCopies_A4reports_REPRINT
                                            End If
                                        End If

                                        For intLoopCounter = 1 To intMaxLoop
                                            If mblnEwaybill_Print = False Then
                                                If optInvYes(0).Checked = True Then
                                                    COPYNAME = Find_Value("Select TEXTHEADING FROM  A4CUSTOMER_INVOICEPRINTINGTAG WHERE UNIT_CODE='" + gstrUNITID + "' AND CUSTOMER_CODE='" & strAccountCode & "' AND ORIGINAL_REPRINT='O' AND SERIALNO=" & intLoopCounter)
                                                Else
                                                    COPYNAME = Find_Value("Select TEXTHEADING FROM  A4CUSTOMER_INVOICEPRINTINGTAG WHERE UNIT_CODE='" + gstrUNITID + "' AND CUSTOMER_CODE='" & strAccountCode & "' AND ORIGINAL_REPRINT='R' AND SERIALNO=" & intLoopCounter)
                                                End If
                                            Else
                                                If chkprintreprint.Checked = True And optInvYes(1).Checked = True Then
                                                    COPYNAME = Find_Value("Select TEXTHEADING FROM  A4CUSTOMER_INVOICEPRINTINGTAG WHERE UNIT_CODE='" + gstrUNITID + "' AND CUSTOMER_CODE='" & strAccountCode & "' AND ORIGINAL_REPRINT='R' AND SERIALNO=" & intLoopCounter)
                                                ElseIf chkprintreprint.Checked = False And optInvYes(1).Checked = True Then
                                                    COPYNAME = Find_Value("Select TEXTHEADING FROM  A4CUSTOMER_INVOICEPRINTINGTAG WHERE UNIT_CODE='" + gstrUNITID + "' AND CUSTOMER_CODE='" & strAccountCode & "' AND ORIGINAL_REPRINT='O' AND SERIALNO=" & intLoopCounter)
                                                Else
                                                    COPYNAME = Find_Value("Select TEXTHEADING FROM  A4CUSTOMER_INVOICEPRINTINGTAG WHERE UNIT_CODE='" + gstrUNITID + "' AND CUSTOMER_CODE='" & strAccountCode & "' AND ORIGINAL_REPRINT='R' AND SERIALNO=" & intLoopCounter)
                                                End If
                                            End If
                                            RdAddSold.DataDefinition.FormulaFields("InsuranceFlag").Text = "'" + CStr(mblnInsuranceFlag) + "'"
                                            Frm.SetReportDocument()

                                            'If optInvYes(1).Checked = True Then
                                            '    dblewaymaxvalue = Find_Value("select total_amount from saleschallan_Dtl where unit_code='" + gstrUNITID + "' and doc_no =" & mInvNo)
                                            'Else
                                            '    dblewaymaxvalue = Find_Value("select total_amount from saleschallan_Dtl where unit_code='" + gstrUNITID + "' and doc_no =" & Ctlinvoice.Text)
                                            'End If

                                            'RdAddSold.PrintToPrinter(1, False, 0, 0)
                                            If mblnEwaybill_Print = False Then
                                                RdAddSold.PrintToPrinter(1, False, 0, 0)
                                            Else
                                                If chkprintreprint.Checked = True And optInvYes(1).Checked = True Then
                                                    RdAddSold.PrintToPrinter(1, False, 0, 0)
                                                Else
                                                    If optInvYes(1).Checked = True Then
                                                        If Not DataExist("SELECT TOP 1 1 FROM FIRSTTIME_INVOICEPRINTING  WHERE UNIT_CODE = '" & gstrUNITID & "' and doc_no= " & Trim(mInvNo) & "") = True Then
                                                            mP_Connection.Execute("Insert into FIRSTTIME_INVOICEPRINTING(unit_code,doc_no,ent_dt,ent_userid) values('" & gstrUNITID & "','" & Trim$(mInvNo) & "',',getdate(),'" & mP_User & "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                                            RdAddSold.PrintToPrinter(1, False, 0, 0)
                                                        Else
                                                            RdAddSold.PrintToPrinter(1, False, 0, 0)
                                                        End If
                                                    End If

                                                End If
                                            End If
                                        Next
                                    Else
                                        'hilex change 
                                        If optInvYes(0).Checked = True Then
                                            intMaxLoop = intNoCopies
                                        Else
                                            If intNoCopies > 1 Then
                                                intMaxLoop = intNoCopies - 1
                                            Else
                                                intMaxLoop = intNoCopies
                                            End If
                                        End If
                                    End If

                                    CheckASNExist(Me.Ctlinvoice.Text)
                                    'Added for Issue ID eMpro-20090415-30143 Starts
                                    If AllowASNPrinting(strAccountCode) = True Then
                                        If mblnASNExist = True Then
                                            mP_Connection.Execute("Update CreatedASN Set ASN_NO='" & Trim$(txtASNNumber.Text) & "' WHERE UNIT_CODE='" + gstrUNITID + "' AND  doc_no='" & Trim$(Me.Ctlinvoice.Text) & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                        Else
                                            mP_Connection.Execute("Insert into CreatedASN values('" & Trim$(Me.Ctlinvoice.Text) & "','" & Trim$(txtASNNumber.Text) & "',getdate(),'" & mP_User & "',getdate(),'" + gstrUNITID + "','" & dtpASNDatetime.Value & "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                        End If
                                    End If
                                    'Added for Issue ID eMpro-20090415-30143 Ends

                                    If mblncustomerlevel_A4report_functionlity = True And mblnA4reports_invoicewise = True Then 'FOR A4 CUSTOMERS
                                        If optInvYes(0).Checked = True Then
                                            intNoCopies_A4reports_orignial = CInt(Find_Value("select isnull(MAX(SERIALNO),0) SERIALNO from A4CUSTOMER_INVOICEPRINTINGTAG  WHERE UNIT_CODE='" + gstrUNITID + "'AND CUSTOMER_CODE='" & strAccountCode & "' AND ORIGINAL_REPRINT='O'"))
                                            intMaxLoop = intNoCopies_A4reports_orignial
                                        Else
                                            intNoCopies_A4reports_REPRINT = CInt(Find_Value("select isnull(MAX(SERIALNO),0) SERIALNO  from A4CUSTOMER_INVOICEPRINTINGTAG  WHERE UNIT_CODE='" + gstrUNITID + "'AND CUSTOMER_CODE='" & strAccountCode & "' AND ORIGINAL_REPRINT='R'"))
                                            intMaxLoop = intNoCopies_A4reports_REPRINT
                                        End If
                                        For intLoopCounter = 1 To intMaxLoop
                                            If optInvYes(0).Checked = True Then
                                                COPYNAME = Find_Value("Select TEXTHEADING FROM  A4CUSTOMER_INVOICEPRINTINGTAG WHERE UNIT_CODE='" + gstrUNITID + "' AND CUSTOMER_CODE='" & strAccountCode & "' AND ORIGINAL_REPRINT='O' AND SERIALNO=" & intLoopCounter)
                                            Else
                                                COPYNAME = Find_Value("Select TEXTHEADING FROM  A4CUSTOMER_INVOICEPRINTINGTAG WHERE UNIT_CODE='" + gstrUNITID + "' AND CUSTOMER_CODE='" & strAccountCode & "' AND ORIGINAL_REPRINT='R' AND SERIALNO=" & intLoopCounter)
                                            End If
                                            RdAddSold.DataDefinition.FormulaFields("CopyName").Text = "'" & COPYNAME & "'"
                                            RdAddSold.DataDefinition.FormulaFields("InsuranceFlag").Text = "'" + CStr(mblnInsuranceFlag) + "'"
                                            Frm.SetReportDocument()
                                            RdAddSold.PrintToPrinter(1, False, 0, 0)
                                            If optInvYes(0).Checked = False Then
                                                If Not DataExist("SELECT TOP 1 1 FROM FIRSTTIME_INVOICEPRINTING  WHERE UNIT_CODE = '" & gstrUNITID & "' and doc_no= " & Trim(Ctlinvoice.Text) & "") = True Then
                                                    mP_Connection.Execute("Insert into FIRSTTIME_INVOICEPRINTING(unit_code,doc_no,ent_dt,ent_userid) values('" & gstrUNITID & "','" & Trim$(Ctlinvoice.Text) & "',getdate(),'" & mP_User & "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                                End If
                                            End If
                                        Next
                                    Else 'FOR NON A4 CUSTOMERS 

                                        For intLoopCounter = 1 To intMaxLoop
                                            Select Case intLoopCounter
                                                Case 1
                                                    If optInvYes(0).Checked = True Then
                                                        RdAddSold.DataDefinition.FormulaFields("CopyName").Text = "'ORIGINAL FOR BUYER'"
                                                    Else
                                                        RdAddSold.DataDefinition.FormulaFields("CopyName").Text = "'DUPLICATE FOR TRANSPORTER'"
                                                    End If
                                                Case 2
                                                    If optInvYes(0).Checked = True Then
                                                        RdAddSold.DataDefinition.FormulaFields("CopyName").Text = "'DUPLICATE FOR TRANSPORTER'"
                                                    Else
                                                        RdAddSold.DataDefinition.FormulaFields("CopyName").Text = "'TRIPLICATE FOR ASSESSEE'"
                                                    End If
                                                Case 3
                                                    If optInvYes(0).Checked = True Then
                                                        RdAddSold.DataDefinition.FormulaFields("CopyName").Text = "'TRIPLICATE FOR ASSESSEE'"
                                                    Else
                                                        RdAddSold.DataDefinition.FormulaFields("CopyName").Text = "'EXTRA COPY'"
                                                    End If
                                                Case Is >= 4
                                                    RdAddSold.DataDefinition.FormulaFields("CopyName").Text = "'EXTRA COPY'"
                                            End Select
                                            RdAddSold.DataDefinition.FormulaFields("InsuranceFlag").Text = "'" + CStr(mblnInsuranceFlag) + "'"
                                            If UCase(Trim(GetPlantName)) = "VF1" Or UCase(Trim(GetPlantName)) = "RSA" Then
                                                If mstrReportFilename = "rptinvoicemate_VF1" Or mstrReportFilename = "rptinvoiceMATE_RSA" Then
                                                    'rptinvoice.set_Formulas(30, "Deliverynoteno='" & mInvNo & "'")
                                                    RdAddSold.DataDefinition.FormulaFields("Deliverynoteno").Text = "'" + CStr(mInvNo) + "'"
                                                End If
                                            End If
                                        Next
                                    End If
                                End If
                                'Changed for Issue ID eMpro-20090216-27468 Ends(First Lock the invoice then send to printer)

                            End If
                        End If

                    End If

                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT_TO_FILE
                    'Reset the mouse pointer
                    If chkPDFExport.Checked = True Then
                        If txtPDFpath.Text = "" Then
                            MsgBox("Please define Folder Path first.", MsgBoxStyle.Information, ResolveResString(100))
                            Exit Sub
                        End If
                        If optInvYes(0).Checked = True Then
                            MsgBox("Invoice Text File cannot be exported For Temporary invoice's .", MsgBoxStyle.Information, ResolveResString(100))
                            txtPDFpath.Text = String.Empty
                            Exit Sub
                        Else
                            'added by priti IRN barcode in export option start here on 19 Feb 2025
                            If mblnEwaybill_Print Then
                                Call IRN_QRBarcode()
                            End If
                            'end by priti IRN barcode in export option start here on 19 Feb 2025


                            If InvoiceGeneration(RdAddSold, RepPath, Frm) = True Then
                                strAccountCode = Find_Value("select account_code from saleschallan_dtl where UNIT_CODE = '" & gstrUNITID & "' and doc_no='" & Trim(Me.Ctlinvoice.Text) & "'")
                                If optInvYes(0).Checked = True Then
                                    mintnocopies = CInt(Find_Value("select isnull(MAX(SERIALNO),0) SERIALNO from A4CUSTOMER_INVOICEPRINTINGTAG  WHERE UNIT_CODE='" + gstrUNITID + "'AND CUSTOMER_CODE='" & strAccountCode & "' AND ORIGINAL_REPRINT='O'"))
                                Else
                                    mintnocopies = CInt(Find_Value("select isnull(MAX(SERIALNO),0) SERIALNO from A4CUSTOMER_INVOICEPRINTINGTAG  WHERE UNIT_CODE='" + gstrUNITID + "'AND CUSTOMER_CODE='" & strAccountCode & "' AND ORIGINAL_REPRINT='R'"))
                                End If

                                ''added by priti Hero QR Code start here for export option on 23 Jan 2025 
                                If (Find_Value("SELECT top 1 1 FROM LISTS WHERE UNIT_CODE='" & gstrUNITID & "' AND KEY1='HEROBAR' and Key2='" & strAccountCode & "' ")) = "1" Then
                                    Dim strBarcodestring1 As String = ""
                                    Dim strBarcodestring2 As String = ""
                                    Dim strBarcodeString3 As String = ""
                                    Dim totalAccesibleAmt As Double = 0
                                    Dim totalCgst As Double = 0
                                    Dim totalIgst As Double = 0
                                    Dim totalSgst As Double = 0
                                    Dim totaltcs As Double = 0

                                    rsGENERATEBARCODE = New ClsResultSetDB_Invoice

                                    ' rsGENERATEBARCODE.GetResult("SELECT  C.CUST_VENDOR_CODE,A.CUST_REF,A.DOC_NO,Convert(varchar,A.INVOICE_DATE,104) INVOICE_DATE,GSTIN_ID, Convert(numeric(8,2),ACCESSIBLE_AMOUNT) ACCESSIBLE_AMOUNT, Convert(numeric(8,2),C.TOTAL_AMOUNT) TOTAL_AMOUNT,A.VEHICLE_NO,Convert(numeric(8,2),isnull(SGSTTXRT_AMOUNT,0)) SGSTTXRT_AMOUNT, " & _
                                    '" Convert(numeric(8,2),isnull(IGSTTXRT_AMOUNT,0)) IGSTTXRT_AMOUNT, Convert(numeric(8,2),isnull(CGSTTXRT_AMOUNT,0)) CGSTTXRT_AMOUNT,CUST_item_CODE,HSNSACCODE,Convert(numeric(8,2),SALES_QUANTITY) SALES_QUANTITY,Convert(numeric(8,2),Rate) Rate " & _
                                    '" FROM SALESCHALLAN_DTL A INNER JOIN GEN_UNITMASTER B ON A.UNIT_CODE=B.UNT_CODEID " & _
                                    '" INNER JOIN TMP_INVOICEPRINT C ON B.UNT_CODEID=C.UNIT_CODE AND A.DOC_NO=C.DOC_NO " & _
                                    '" WHERE A.DOC_NO='" & Ctlinvoice.Text.Trim & "' AND A.UNIT_CODE='" & gstrUNITID & "' ")

                                    rsGENERATEBARCODE.GetResult("SELECT  C.CUST_VENDOR_CODE,A.CUST_REF,A.DOC_NO,Convert(varchar,A.INVOICE_DATE,104) INVOICE_DATE,GSTIN_ID, Convert(numeric(19,2),ACCESSIBLE_AMOUNT) ACCESSIBLE_AMOUNT, Convert(numeric(19,2),C.TOTAL_AMOUNT) TOTAL_AMOUNT,A.VEHICLE_NO,Convert(numeric(19,2),isnull(SGSTTXRT_AMOUNT,0)) SGSTTXRT_AMOUNT, " &
                                    " Convert(numeric(19,2),isnull(IGSTTXRT_AMOUNT,0)) IGSTTXRT_AMOUNT, Convert(numeric(19,2),isnull(CGSTTXRT_AMOUNT,0)) CGSTTXRT_AMOUNT,CUST_item_CODE,HSNSACCODE,Convert(numeric(12,2),SALES_QUANTITY) SALES_QUANTITY,Convert(numeric(19,2),Rate) Rate ,Convert(numeric(19,2),isnull(TCSAMOUNT,0)) TCSAMOUNT " &
                                    " FROM SALESCHALLAN_DTL A INNER JOIN GEN_UNITMASTER B ON A.UNIT_CODE=B.UNT_CODEID " &
                                    " INNER JOIN TMP_INVOICEPRINT C ON B.UNT_CODEID=C.UNIT_CODE AND A.DOC_NO=C.DOC_NO " &
                                    " WHERE A.DOC_NO='" & Ctlinvoice.Text.Trim & "' AND A.UNIT_CODE='" & gstrUNITID & "' AND C.IP_ADDRESS= '" & gstrIpaddressWinSck & "'")


                                    While Not rsGENERATEBARCODE.EOFRecord
                                        totalAccesibleAmt = totalAccesibleAmt + Convert.ToDouble(rsGENERATEBARCODE.GetValue("ACCESSIBLE_AMOUNT").ToString)
                                        If optInvYes(0).Checked = True Then
                                            strBarcodestring1 = rsGENERATEBARCODE.GetValue("CUST_VENDOR_CODE").ToString & vbTab & rsGENERATEBARCODE.GetValue("CUST_REF").ToString & vbTab & mInvNo.ToString & vbTab & rsGENERATEBARCODE.GetValue("INVOICE_DATE").ToString & vbTab & rsGENERATEBARCODE.GetValue("GSTIN_ID").ToString & vbTab & rsGENERATEBARCODE.GetValue("TOTAL_AMOUNT").ToString
                                        Else
                                            strBarcodestring1 = rsGENERATEBARCODE.GetValue("CUST_VENDOR_CODE").ToString & vbTab & rsGENERATEBARCODE.GetValue("CUST_REF").ToString & vbTab & Ctlinvoice.Text.ToString & vbTab & rsGENERATEBARCODE.GetValue("INVOICE_DATE").ToString & vbTab & rsGENERATEBARCODE.GetValue("GSTIN_ID").ToString & vbTab & rsGENERATEBARCODE.GetValue("TOTAL_AMOUNT").ToString
                                        End If

                                        strBarcodeString3 = rsGENERATEBARCODE.GetValue("VEHICLE_NO").ToString
                                        totalSgst = totalSgst + Convert.ToDouble(rsGENERATEBARCODE.GetValue("SGSTTXRT_AMOUNT").ToString)
                                        totalIgst = totalIgst + Convert.ToDouble(rsGENERATEBARCODE.GetValue("IGSTTXRT_AMOUNT").ToString)
                                        totalCgst = totalCgst + Convert.ToDouble(rsGENERATEBARCODE.GetValue("CGSTTXRT_AMOUNT").ToString)
                                        totaltcs = Convert.ToDouble(rsGENERATEBARCODE.GetValue("TCSAMOUNT").ToString)
                                        strBarcodestring2 = strBarcodestring2 & vbTab & rsGENERATEBARCODE.GetValue("CUST_item_CODE").ToString & vbTab & rsGENERATEBARCODE.GetValue("HSNSACCODE").ToString & vbTab & rsGENERATEBARCODE.GetValue("SALES_QUANTITY").ToString & vbTab & rsGENERATEBARCODE.GetValue("Rate").ToString
                                        rsGENERATEBARCODE.MoveNext()
                                    End While
                                    rsGENERATEBARCODE.ResultSetClose()


                                    Dim PDF417barcode As BarcodeLib.Barcode.PDF417.PDF417 = New BarcodeLib.Barcode.PDF417.PDF417()
                                    PDF417barcode.UOM = BarcodeLib.Barcode.Linear.UnitOfMeasure.Pixel
                                    PDF417barcode.LeftMargin = 0
                                    PDF417barcode.RightMargin = 0
                                    PDF417barcode.TopMargin = 0
                                    PDF417barcode.BottomMargin = 0
                                    PDF417barcode.ImageFormat = System.Drawing.Imaging.ImageFormat.Png

                                    PDF417barcode.Data = (strBarcodestring1 & vbTab & totalAccesibleAmt.ToString & vbTab & strBarcodeString3 & vbTab & totalSgst.ToString & vbTab & totalIgst.ToString & vbTab & totalCgst.ToString & vbTab & totaltcs.ToString + strBarcodestring2).ToString().Trim
                                    Dim imageData() As Byte = PDF417barcode.drawBarcodeAsBytes()

                                    Dim cmd As SqlCommand = Nothing
                                    cmd = New System.Data.SqlClient.SqlCommand()
                                    With cmd
                                        .CommandType = CommandType.Text
                                        .CommandText = "UPDATE TMP_INVOICEPRINT SET barcodeimage=@QRIMAGE where DOC_NO='" & Ctlinvoice.Text.Trim & "' AND UNIT_CODE='" & gstrUNITID & "' "
                                        .Parameters.Add("@QRIMAGE", SqlDbType.Image).Value = imageData
                                        SqlConnectionclass.ExecuteNonQuery(cmd)

                                    End With
                                End If

                                ''Hero QR Code ends here for export option on 23 Jan 2025 


                                ''added by priti TATA QR Code start here for export option on 05 May 2025 

                                If GetPrintMethod(strAccountCode).ToUpper() = "TATA" Then  '' For Export Invoice
                                    Dim StrTATAsuffix As String
                                    StrTATAsuffix = gstrUNITID & Ctlinvoice.Text.ToString.Trim & DateTime.Now.ToString("ddMMyyHHmmssfff")
                                    strBarcodeMsg = ObjBarcodeHMI.GenerateQRBarCodeForTATAMotors(gstrUserMyDocPath, True, Trim(Ctlinvoice.Text), Trim(Ctlinvoice.Text), gstrCONNECTIONSTRING, StrTATAsuffix)
                                    If VB.Left(strBarcodeMsg, 1) <> "Y" Then
                                        CustomRollbackTrans()
                                        MsgBox("Problem While Generating Barcode Image.", vbInformation, ResolveResString(100))
                                        Exit Sub
                                    Else
                                        strBarcodeMsg_paratemeter = Mid(strBarcodeMsg, 3)
                                        If Not SaveQRBarCodeImageTATA(Trim(Ctlinvoice.Text), 0, strBarcodeMsg_paratemeter, StrTATAsuffix) Then
                                            CustomRollbackTrans()
                                            MsgBox("Problem While Saving Barcode Image.", vbInformation, ResolveResString(100))
                                            Exit Sub
                                        End If
                                    End If
                                End If


                                intMaxLoop = mintnocopies
                                For intLoopCounter = 1 To intMaxLoop
                                    Dim HeadName As String = ""
                                    Dim PDFNAME_SERIALNO As Boolean = SqlConnectionclass.ExecuteScalar("Select isnull(PDFNAME_SERIALNO,0) from customer_mst where unit_code='" + gstrUNITID + "' and customer_code='" + strAccountCode + "' ")
                                    If PDFNAME_SERIALNO = True Then
                                        If optInvYes(0).Checked = True Then
                                            HeadName = Find_Value("Select serialno FROM  A4CUSTOMER_INVOICEPRINTINGTAG WHERE UNIT_CODE='" + gstrUNITID + "' AND CUSTOMER_CODE='" & strAccountCode & "' AND ORIGINAL_REPRINT='O' AND SERIALNO=" & intLoopCounter)
                                            COPYNAME = Find_Value("Select TEXTHEADING FROM  A4CUSTOMER_INVOICEPRINTINGTAG WHERE UNIT_CODE='" + gstrUNITID + "' AND CUSTOMER_CODE='" & strAccountCode & "' AND ORIGINAL_REPRINT='O' AND SERIALNO=" & intLoopCounter)
                                        Else
                                            HeadName = Find_Value("Select serialno FROM  A4CUSTOMER_INVOICEPRINTINGTAG WHERE UNIT_CODE='" + gstrUNITID + "' AND CUSTOMER_CODE='" & strAccountCode & "' AND ORIGINAL_REPRINT='R' AND SERIALNO=" & intLoopCounter)
                                            COPYNAME = Find_Value("Select TEXTHEADING FROM  A4CUSTOMER_INVOICEPRINTINGTAG WHERE UNIT_CODE='" + gstrUNITID + "' AND CUSTOMER_CODE='" & strAccountCode & "' AND ORIGINAL_REPRINT='R' AND SERIALNO=" & intLoopCounter)

                                        End If
                                    Else

                                        If optInvYes(0).Checked = True Then
                                            COPYNAME = Find_Value("Select TEXTHEADING FROM  A4CUSTOMER_INVOICEPRINTINGTAG WHERE UNIT_CODE='" + gstrUNITID + "' AND CUSTOMER_CODE='" & strAccountCode & "' AND ORIGINAL_REPRINT='O' AND SERIALNO=" & intLoopCounter)
                                        Else
                                            COPYNAME = Find_Value("Select TEXTHEADING FROM  A4CUSTOMER_INVOICEPRINTINGTAG WHERE UNIT_CODE='" + gstrUNITID + "' AND CUSTOMER_CODE='" & strAccountCode & "' AND ORIGINAL_REPRINT='R' AND SERIALNO=" & intLoopCounter)
                                        End If
                                        HeadName = COPYNAME
                                    End If

                                    RdAddSold.DataDefinition.FormulaFields("CopyName").Text = "'" & COPYNAME & "'"
                                    Frm.SetReportDocument()
                                    strinvfilename = Ctlinvoice.Text.ToString + "_" + strAccountCode + "_" + HeadName + ".PDF"
                                    strpath = txtPDFpath.Text
                                    strfullpath = strpath + "\" + strinvfilename
                                    If File.Exists(strfullpath) = True Then
                                        Kill(strfullpath)
                                    End If
                                    RdAddSold.ExportToDisk(ExportFormatType.PortableDocFormat, strfullpath)

                                Next

                                'Added by priti on 16 Jan 2025 to skip first time invoice printing in case of PDF generation
                                If UCase(cmbInvType.Text) <> "REJECTION" Then
                                    Dim strPDFPath = SqlConnectionclass.ExecuteScalar("Select isnull(Invoice_PDFCOPYPATH_PRINT,'') from Customer_mst where unit_code='" & gstrUNITID & "'  and customer_code='" & strAccountCode & "'")
                                    If Len(strPDFPath.ToString) > 0 Then
                                        If Not DataExist("SELECT TOP 1 1 FROM FIRSTTIME_INVOICEPRINTING  WHERE UNIT_CODE = '" & gstrUNITID & "' and doc_no= " & Trim(Ctlinvoice.Text) & "") = True Then
                                            SqlConnectionclass.ExecuteNonQuery("Insert into FIRSTTIME_INVOICEPRINTING(unit_code,doc_no,ent_dt,ent_userid) values('" & gstrUNITID & "','" & Trim$(Ctlinvoice.Text) & "',getdate(),'" & mP_User & "')")
                                        End If
                                    End If
                                End If
                                'End by priti on 16 Jan 2025 to skip first time invoice printing in case of PDF generation

                                MsgBox("Invoice PDF Files are exported at location." + strpath, MsgBoxStyle.Information, ResolveResString(100))


                            End If

                        End If
                        If CBool(Find_Value("select TextPrinting from sales_parameter WHERE UNIT_CODE='" + gstrUNITID + "'")) Then
                            MsgBox("Invoice Text File report cannot be exported.", MsgBoxStyle.Information, ResolveResString(100))
                            Exit Sub
                        End If

                    Else

                        If InvoiceGeneration(RdAddSold, RepPath, Frm) = True Then
                            'Added for Issue ID eMpro-20080805-20745 Starts
                            If CBool(Find_Value("select TextPrinting from sales_parameter WHERE UNIT_CODE='" + gstrUNITID + "'")) Then
                                MsgBox("Invoice Text File report cannot be exported.", MsgBoxStyle.Information, ResolveResString(100))
                                Exit Sub
                            End If

                            Frm.ExportToFile()
                            'Added for Issue ID eMpro-20080805-20745 Ends
                            '' frmExport.ShowDialog()
                            '' Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
                            '' If gblnCancelExport Then Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default) : Exit Sub
                            ''  Me.rptinvoice.PrintFileType = ModDeclares.genmFileType
                            ''  Me.rptinvoice.Destination = Crystal.DestinationConstants.crptToFile
                            ''  Me.rptinvoice.Action = 1
                            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                        End If

                    End If
            End Select
            '''objInvoicePrint = Nothing
            '''rsItembal.ResultSetClose()
            'UPGRADE_NOTE: Object rsItembal may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            '''rsItembal = Nothing

            '''Added By Ashutosh on 20 Aug 2007 ,Issue Id: 20876
            Dim strMessage As String
            If Me.optInvYes(0).Checked = True Then
                If Find_Value("Select Cust_name from saleschallan_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  doc_no ='" & Trim(CStr(mInvNo)) & "' and cust_name like '%toyota%' ") <> "" Then
                    strMessage = ToyotaTextFile(Trim(CStr(mInvNo)))
                End If
            Else
                If Find_Value("Select Cust_name from saleschallan_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  doc_no ='" & Trim(Ctlinvoice.Text) & "' and cust_name like '%toyota%' ") <> "" Then
                    strMessage = ToyotaTextFile(Trim(Ctlinvoice.Text))
                End If
            End If
            '''Changes for ISsue ID:20876 Ends here.


            Exit Sub
            'Err_Handler:
            '            If Err.Number = 20545 Then
            '                Resume Next
            '            Else
            '                objInvoicePrint = Nothing
            '                Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
            '            End If
        Catch Ex As Exception
            Dim IsPrimaryKeyOccured As Boolean = False
            If Ex.Message.ToUpper.ToString.Contains("VIOLATION OF PRIMARY KEY CONSTRAINT") And Ex.Message.ToUpper.ToString.Contains("SALESCHALLAN_DTL") Then
                IsPrimaryKeyOccured = True
                Try
                    CustomRollbackTrans()
                Catch
                End Try
                updatesalesconfandsaleschallan_Contingency(mInvNo.ToString, mAssessableValue.ToString())
                SqlConnectionclass.ExecuteNonQuery("INSERT INTO INV_ERR_PRIMARYKEY(ERR_NUMBER,ERR_DESCRIPTION,TMPINVOICENO,INVOICE_NO,UNIT_CODE,FUNCTIONNAME ) VALUES('" & Err.Number & "','" & Replace(Ex.Message, "'", "") & "','" & Ctlinvoice.Text & "','" & mInvNo & "','" & gstrUNITID & "','cmdinvoice: PK Correction Attemped')")
                MsgBox(Ex.Message + " :Attemped to correct Internal PK Issue. Please Try Again!!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, My.Resources.resEmpower.STR100)
            End If

            If Err.Number = 20545 Then
                'Resume Next
            Else
                SqlConnectionclass.ExecuteNonQuery("INSERT INTO INV_ERR_PRIMARYKEY(ERR_NUMBER,ERR_DESCRIPTION,TMPINVOICENO,INVOICE_NO,UNIT_CODE,FUNCTIONNAME ) VALUES('" & Err.Number & "','" & Replace(Ex.Message, "'", "") & "','" & Ctlinvoice.Text & "','" & mInvNo & "','" & gstrUNITID & "','cmdinvoice')")
                Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Ex.Message, mP_Connection, "CMDINVOICE")
            End If
        Finally
            '-- changes started by prashant rajpal on 09th Mar 2023
            If optInvYes(0).Checked = True And Val(mInvNo) <> 0 And blninvoicelockYES = True Then
                If Not strRetVal = "Y" Then
                    If Not DataExist("select doc_no from Saleschallan_Dtl where doc_no='" & mInvNo & "' and UNIT_CODE = '" & gstrUNITID & "'") Then
                        CustomRollbackTrans()
                        mP_Connection.BeginTrans()
                        mP_Connection.Execute("delete from ar_docmaster WHERE docM_unit='" + gstrUNITID + "' AND  docm_vono='" & mInvNo & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        mP_Connection.Execute("delete from ar_docdtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  docd_vono='" & mInvNo & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        mP_Connection.Execute("delete from fin_gltrans WHERE glt_UntCodeID='" + gstrUNITID + "' AND  glt_srcdocno='" & mInvNo & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        mP_Connection.CommitTrans()
                        Call Logging_Starting_End_Time("Deletion of Finance Existance Tables : " + gstrUNITID + ":" + mInvNo.ToString(), DateTime.Now.ToString(), "Saved", mInvNo.ToString())
                        MsgBox("Internal Issue Occured !! ,  Please try Again.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, My.Resources.resEmpower.STR100)
                    End If
                End If
            End If
            '-- changes done by prashant rajpal on 09th Mar 2023
        End Try

    End Sub

    Private Function GetPrintMethod(ByVal pstraccoutncode As String) As String
        On Error GoTo ErrHandler
        Dim strQry As String
        Dim Rs As ClsResultSetDB_Invoice
        GetPrintMethod = String.Empty
        strQry = "Select isnull(PRINT_METHOD,'') as PRINT_METHOD from customer_mst (nolock) where Customer_Code='" & Trim(pstraccoutncode) & "' and UNIT_CODE = '" & gstrUNITID & "'"
        Rs = New ClsResultSetDB_Invoice
        If Rs.GetResult(strQry) = False Then GoTo ErrHandler
        If Convert.ToString(Rs.GetValue("PRINT_METHOD")).Length > 0 Then
            GetPrintMethod = Convert.ToString(Rs.GetValue("PRINT_METHOD"))
        Else
            GetPrintMethod = String.Empty
        End If
        Rs.ResultSetClose()
        Rs = Nothing
        Exit Function
ErrHandler:
        Rs = Nothing
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

    Public Function SaveQRBarCodeImageTATA(ByVal tempInvoiceNo As Long, ByVal invoiceNo As Long, ByVal barcodeString As String, ByVal strsuffix As String) As Boolean

        On Error GoTo ErrHandler

        Dim stimage As ADODB.Stream
        Dim strQuery As String
        Dim pstrPath As String = ""
        Dim blnCROP_QRIMAGE As Boolean = False
        pstrPath = gstrUserMyDocPath
        SaveQRBarCodeImageTATA = True
        Dim size As Integer = 0

        'pstrPath = pstrPath & "QRBarcodeImgTataMotors.wmf"
        pstrPath = pstrPath & "QRBarcodeImgTataMotors" & strsuffix & ".wmf"
        blnCROP_QRIMAGE = CBool(Find_Value("SELECT CROP_QRBARCODE FROM SALES_PARAMETER (NOLOCK) WHERE UNIT_CODE='" + gstrUNITID + "'"))
        If blnCROP_QRIMAGE = True Then
            Dim bmp As New Bitmap(pstrPath)
            Dim picturebox1 As New PictureBox
            picturebox1.Image = ImageTrim(bmp)
            picturebox1.Image.Save(pstrPath)
            picturebox1 = Nothing
        End If
        Dim fs As FileStream = New FileStream(pstrPath, FileMode.Open, FileAccess.Read)
        Dim br As BinaryReader = New BinaryReader(fs)
        Dim bytes As Byte() = br.ReadBytes(Convert.ToInt32(fs.Length))
        br.Close()
        fs.Close()

        If bytes Is Nothing OrElse bytes.Length = 0 Then
            size = 8000
        Else
            size = bytes.Length
        End If
        Dim objComm As New ADODB.Command
        With objComm
            .ActiveConnection = mP_Connection
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandText = "USP_QR_CODE_TATA_MOTORS"
            .CommandTimeout = 0
            .Parameters.Append(.CreateParameter("@TEMP_INVOICE_NO", ADODB.DataTypeEnum.adNumeric, ADODB.ParameterDirectionEnum.adParamInput, , tempInvoiceNo))
            .Parameters("@TEMP_INVOICE_NO").Precision = 18 : .Parameters("@TEMP_INVOICE_NO").NumericScale = 0
            .Parameters.Append(.CreateParameter("@INVOICE_NO", ADODB.DataTypeEnum.adNumeric, ADODB.ParameterDirectionEnum.adParamInput, , invoiceNo))
            .Parameters("@INVOICE_NO").Precision = 18 : .Parameters("@INVOICE_NO").NumericScale = 0
            .Parameters.Append(.CreateParameter("@IS_LOCK", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, , 0))
            .Parameters.Append(.CreateParameter("@IS_REPRINT", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, , IIf(optInvYes(0).Checked, 0, 1)))
            .Parameters.Append(.CreateParameter("@UNIT_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
            .Parameters.Append(.CreateParameter("@BARCODE_IMAGE", ADODB.DataTypeEnum.adVarBinary, ADODB.ParameterDirectionEnum.adParamInput, size, bytes))
            .Parameters.Append(.CreateParameter("@BARCODE_STRING", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 8000, barcodeString))
            .Parameters.Append(.CreateParameter("@USER_ID", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 50, mP_User))
            .Parameters.Append(.CreateParameter("@IP_ADDRESS", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 50, gstrIpaddressWinSck))
            .Parameters.Append(.CreateParameter("@OPERATION_TYPE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 50, "SAVE_QR_CODE_STRING"))
            .Parameters.Append(.CreateParameter("@MESSAGE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamOutput, 8000, ""))
            .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

            If Convert.ToString(.Parameters(.Parameters.Count - 1).Value) <> String.Empty Then
                SaveQRBarCodeImageTATA = False
                MessageBox.Show(.Parameters(.Parameters.Count - 1).Value.ToString(), ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error)
                objComm = Nothing
                Exit Function
            End If
        End With
        objComm = Nothing

        Exit Function
ErrHandler:
        SaveQRBarCodeImageTATA = False
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

    Private Sub ctlFormHeader1_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlFormHeader1.Click
        Call ShowHelp("HLPMKTTRN0008.htm")
    End Sub
    Private Function CheckInvoiceExistInFinance(ByVal pstrdoc_no As String)
        '*******************************************************************
        'Revised By     : Manoj Vaish
        'Return Value   : NIL
        'Revised Date   : 26 Feb 2009
        'Issue ID       : eMpro-20090226-27911
        'Revised History: If the Data is already present table in Finance Table
        '                 and invoice no. is not updated in saleconf
        '*******************************************************************
        On Error GoTo Err_Handler

        Dim rsCheckFindata As ClsResultSetDB_Invoice
        Dim strsql As String
        rsCheckFindata = New ClsResultSetDB_Invoice

        'Changed for Issue ID eMpro-20090713-33572 Starts
        If Len(Find_Value("select doc_no from SalesChallan_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  location_code='" & Trim(txtUnitCode.Text) & "' and doc_no='" & pstrdoc_no & "'")) > 0 Then
            MsgBox("Next Invoice number already generated." & vbCrLf & "Please skip current no either backward or forward" & vbCrLf & "in Sales Configuration Master Form.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "eMPro")
            Exit Function
        Else
            strsql = "select docm_vono from ar_docmaster nolock WHERE docM_unit='" + gstrUNITID + "' AND  docm_vono='" & pstrdoc_no & "'"

            rsCheckFindata.GetResult(strsql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If rsCheckFindata.GetNoRows > 0 Then
                mP_Connection.BeginTrans()
                mP_Connection.Execute("delete from ar_docmaster WHERE docM_unit='" + gstrUNITID + "' AND  docm_vono='" & pstrdoc_no & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                mP_Connection.Execute("delete from ar_docdtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  docd_vono='" & pstrdoc_no & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                mP_Connection.Execute("delete from fin_gltrans WHERE glt_UntCodeID='" + gstrUNITID + "' AND  glt_srcdocno='" & pstrdoc_no & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                mP_Connection.Execute("execute updatefinbalances '" + gstrUNITID + "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                mP_Connection.CommitTrans()
            End If

            rsCheckFindata.ResultSetClose()
            rsCheckFindata = Nothing
        End If
        'Changed for Issue ID eMpro-20090713-33572 Ends

        Exit Function
Err_Handler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)

    End Function

    Private Sub txtASNNumber_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtASNNumber.KeyPress
        'Created  By     : Manoj Kr Vaish
        'Creation Date   : 09 Mar 2009
        'Issue ID        : eMpro-20090415-30143
        Dim KeyAscii As Short = Asc(e.KeyChar)
        Select Case KeyAscii
            Case 39, 34, 96
                KeyAscii = 0
            Case 13
                Cmdinvoice.Focus()
        End Select
        e.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            e.Handled = True
        End If
        AllowNumericValueInTextBox(txtASNNumber, e)
    End Sub
    Private Function AllowASNPrinting(ByVal pstraccountcode As String) As Boolean
        'Revised By     : Manoj Kr. Vaish
        'Revised On     : 24 Mar  2009
        'Arguments      : Account Code
        'Return Value   : True/False
        'Issue ID       : eMpro-20090415-30143
        'Reason         : Check ASNPrinting from Customer Master
        '--------------------------------------------------------------------------------------
        On Error GoTo ErrHandler

        Dim strQry As String
        Dim Rs As ClsResultSetDB_Invoice
        AllowASNPrinting = False
        strQry = "Select isnull(AllowASNPrinting,0) as AllowASNPrinting from customer_mst WHERE UNIT_CODE='" + gstrUNITID + "' AND  Customer_Code='" & Trim(pstraccountcode) & "'"
        Rs = New ClsResultSetDB_Invoice
        If Rs.GetResult(strQry) = False Then GoTo ErrHandler

        If Rs.GetValue("AllowASNPrinting") = "True" Then
            AllowASNPrinting = True
        Else
            AllowASNPrinting = False
        End If

        Rs.ResultSetClose()
        Rs = Nothing
        Exit Function
ErrHandler:
        Rs = Nothing
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function CheckASNExist(ByVal pstrInvoiceNo As String) As String
        'Revised By     : Manoj Kr. Vaish
        'Revised On     : 24 Mar 2009
        'Arguments      : Invoice Number
        'Return Value   : ASN Number
        'Issue ID       : eMpro-20090415-30143
        'Reason         : Check ASN already exist for invoice
        '--------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim rsgetASNNumber As ClsResultSetDB_Invoice
        Dim strsql As String

        rsgetASNNumber = New ClsResultSetDB_Invoice
        strsql = "select ASN_NO from CreatedASN WHERE UNIT_CODE='" + gstrUNITID + "' AND  doc_no='" & Trim(pstrInvoiceNo) & "'"

        rsgetASNNumber.GetResult(strsql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
        If rsgetASNNumber.GetNoRows > 0 Then
            CheckASNExist = IIf(IsDBNull(rsgetASNNumber.GetValue("ASN_NO")), "", rsgetASNNumber.GetValue("ASN_NO"))
            mblnASNExist = True
        Else
            mblnASNExist = False
        End If

        rsgetASNNumber.ResultSetClose()
        rsgetASNNumber = Nothing

        Exit Function
ErrHandler:
        rsgetASNNumber = Nothing
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)

    End Function
    Private Sub _optInvYes_0_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _optInvYes_0.CheckedChanged
        On Error GoTo ErrHandler
        If optInvYes(0).Checked = True Then
            Me.txtASNNumber.Visible = False
            Me.txtASNNumber.Enabled = False
            Me.txtASNNumber.Text = ""
            Me.txtASNNumber.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            Me.lblASNNumber.Visible = False
            Me.dtpASNDatetime.Visible = False
            Me.lblASNDateTime.Visible = False
            txtPDFpath.Text = ""
        End If
ErrHandler:

        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)

    End Sub
    Public Sub Send_Report_Printer_Smiel(ByRef RdAddSold As ReportDocument, ByRef RepPath As String, ByRef Frm As Object)
        Dim Phone, Range, RegNo, EccNo, Address, Invoice_Rule, TinNo, DeliveredAdd As String
        Dim CST, PLA, Fax, EMail, UPST, Division, Commissionerate, ExpCode, strQry As String
        Dim rsCompMst As ClsResultSetDB_Invoice
        Dim blnPrintTinNo As Boolean
        Dim rsGrnHdr As ClsResultSetDB_Invoice
        Dim strGRNDate As String = String.Empty
        Dim strVendorInvDate As String = String.Empty
        Dim strVendorInvNo As String = String.Empty
        Dim strCustRefForGrn As String = String.Empty
        Dim oCmd As ADODB.Command
        Dim STRCUSTOMERCODE As String
        Dim strIPAddress As String
        On Error GoTo ErrHandler

        rsCompMst = New ClsResultSetDB_Invoice
        rsCompMst.GetResult("Select Reg_NO,Ecc_No,Range_1,Phone,Fax,Email,PLA_No,LST_No,CST_No,Division,Commissionerate,Invoice_Rule,Exporter_Code,Tin_no from Company_Mst WHERE UNIT_CODE='" + gstrUNITID + "'")
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
            ExpCode = rsCompMst.GetValue("Exporter_Code")
            TinNo = rsCompMst.GetValue("Tin_no")
        End If
        rsCompMst.ResultSetClose()

        RdAddSold = Frm.GetReportDocument()
        STRCUSTOMERCODE = Trim(Find_Value("SELECT ACCOUNT_CODE FROM SALESCHALLAN_DTL WHERE UNIT_CODE='" + gstrUNITID + "'AND DOC_NO='" & Trim(Ctlinvoice.Text) & "'"))
        If IsGSTINSAME(STRCUSTOMERCODE) = True And UCase(Trim(cmbInvType.Text)) = "TRANSFER INVOICE" And System.IO.File.Exists(My.Application.Info.DirectoryPath & "\Reports\Delivery_Challan_Hilex_GST_A4REPORTS.rpt") Then
            RepPath = My.Application.Info.DirectoryPath & "\Reports\Delivery_Challan_Hilex_GST_A4REPORTS.rpt"
        Else
            RepPath = My.Application.Info.DirectoryPath & "\Reports\" & mstrReportFilename & ".rpt"
        End If
        RdAddSold.Load(RepPath)
        '' With rptinvoice
        ''.Reset()
        '' .DiscardSavedData = True
        '' .Connect = gstrREPORTCONNECT
        '' .WindowShowPrintSetupBtn = True
        '' .WindowShowCloseBtn = True
        '' .WindowShowCancelBtn = True
        '' .WindowShowPrintBtn = True
        ''  .WindowShowExportBtn = True
        '' .WindowShowSearchBtn = True
        '' .WindowState = Crystal.WindowStateConstants.crptMaximized
        If UCase(cmbInvType.Text) <> "JOBWORK INVOICE" Then
            RdAddSold.DataDefinition.FormulaFields("Category").Text = "'" + Me.lblcategory.Text + "'"
        End If

        RdAddSold.DataDefinition.FormulaFields("Registration").Text = "'" + RegNo + "'"
        RdAddSold.DataDefinition.FormulaFields("ECC").Text = "'" + EccNo + "'"
        RdAddSold.DataDefinition.FormulaFields("Range").Text = "'" + Range + "'"
        RdAddSold.DataDefinition.FormulaFields("CompanyName").Text = "'" + gstrCOMPANY + "'"
        RdAddSold.DataDefinition.FormulaFields("CompanyAddress").Text = "'" + Address + "'"
        RdAddSold.DataDefinition.FormulaFields("Phone").Text = "'" + Phone + "'"
        RdAddSold.DataDefinition.FormulaFields("Fax").Text = "'" + Fax + "'"

        If UCase(cmbInvType.Text) <> "JOBWORK INVOICE" Then
            RdAddSold.DataDefinition.FormulaFields("EMail").Text = "'" + EMail + "'"
        End If
        RdAddSold.DataDefinition.FormulaFields("PLA").Text = "'" + PLA + "'"
        RdAddSold.DataDefinition.FormulaFields("UPST").Text = "'" + UPST + "'"
        RdAddSold.DataDefinition.FormulaFields("CST").Text = "'" + CST + "'"
        RdAddSold.DataDefinition.FormulaFields("Division").Text = "'" + Division + "'"
        RdAddSold.DataDefinition.FormulaFields("commissionerate").Text = "'" + Commissionerate + "'"
        RdAddSold.DataDefinition.FormulaFields("InvoiceRule").Text = "'" + Invoice_Rule + "'"
        RdAddSold.DataDefinition.FormulaFields("EOUFlag").Text = "'" + CStr(mblnEOUUnit) + "'"

        If optYes(0).Checked = True Then
            RdAddSold.DataDefinition.FormulaFields("DeliveredAt").Text = "' Delivered At '"
            RdAddSold.DataDefinition.FormulaFields("Address2").Text = "'" + DeliveredAdd + "'"
        Else
            RdAddSold.DataDefinition.FormulaFields("DeliveredAt").Text = "' Delivered At '"
            RdAddSold.DataDefinition.FormulaFields("Address2").Text = "''"
            '' .set_Formulas(18, "Address2=''") 'to pass blanck Address in this case will overwrite this Formula written in Crystal Report for else case
        End If

        'rptinvoice.Formulas(16) = "EOUFlag='" & blnEOUFlag & "'"
        RdAddSold.DataDefinition.FormulaFields("PLADuty").Text = "'" + Trim(txtPLA.Text) + "'"
        RdAddSold.DataDefinition.FormulaFields("InsuranceFlag").Text = "'" + CStr(mblnInsuranceFlag) + "'"
        RdAddSold.DataDefinition.FormulaFields("StringYear").Text = "'" + CStr(Year(GetServerDate)) + "'"
        RdAddSold.DataDefinition.FormulaFields("DateOfRemoval").Text = "'" + dtpRemoval.Text + "'"

        If optInvYes(0).Checked = True Then
            RdAddSold.DataDefinition.FormulaFields("InvoiceNo").Text = "'" + CStr(mInvNo) + "'"
        Else
            RdAddSold.DataDefinition.FormulaFields("InvoiceNo").Text = "'" + CStr(mInvNo) + "'"
        End If

        strQry = "Select isnull(SC.RecordsPerPage,7) as RecordsPerPage,isnull(SP.MoreThan7ItemInInvoice,0) as MoreThan7ItemInInvoice"
        strQry = strQry & " From saleschallan_dtl SCD inner join SaleConf SC on SCD.Invoice_Type = SC.Invoice_Type And SCD.Sub_Category = SC.Sub_Type And SCD.UNIT_CODE = SC.UNIT_CODE"
        strQry = strQry & " inner join Sales_parameter SP on Not (isnull(SP.maruti_ac,'') = SCD.Account_code or isnull(SP.maruti_ac1,'') = SCD.Account_code) AND SP.UNIT_CODE = SCD.UNIT_CODE"
        strQry = strQry & " where SCD.UNIT_CODE='" + gstrUNITID + "' AND datediff(dd,SCD.Invoice_Date,SC.fin_start_date)<=0 And datediff(dd,SC.fin_end_date,SCD.Invoice_Date)<=0 And SCD.doc_no=" & Ctlinvoice.Text

        rsCompMst = New ClsResultSetDB_Invoice
        Call rsCompMst.GetResult(strQry, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsCompMst.RowCount > 0 Then
            If rsCompMst.GetValue("MoreThan7ItemInInvoice") = "True" Then
                RdAddSold.DataDefinition.FormulaFields("RecPerPage").Text = "'" + rsCompMst.GetValue("RecordsPerPage") + "'"
            End If
        End If
        rsCompMst.ResultSetClose()

        blnPrintTinNo = CBool(Find_Value("Select isnull(PrintTinNO,0) as PrintTinNO from sales_parameter WHERE UNIT_CODE='" + gstrUNITID + "'"))
        If blnPrintTinNo = True Then
            RdAddSold.DataDefinition.FormulaFields("TinNo").Text = "'" + TinNo + "'"
        End If

        If UCase(cmbInvType.Text) = "REJECTION" Then
            rsGrnHdr = New ClsResultSetDB_Invoice
            strGRNDate = "" : strVendorInvDate = "" : strVendorInvNo = "" : strCustRefForGrn = ""
            rsGrnHdr.GetResult("Select Cust_ref from salesChallan_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  doc_No = " & Ctlinvoice.Text)
            If rsGrnHdr.GetNoRows > 0 Then
                rsGrnHdr.MoveFirst()
                strCustRefForGrn = rsGrnHdr.GetValue("Cust_ref")
            End If
            rsGrnHdr.ResultSetClose()
            If Len(Trim(strCustRefForGrn)) > 0 Then
                rsGrnHdr = New ClsResultSetDB_Invoice
                rsGrnHdr.GetResult("select grn_date,Invoice_no,Invoice_date from grn_hdr WHERE UNIT_CODE='" + gstrUNITID + "' AND  From_Location ='01R1' and doc_No = " & strCustRefForGrn)
                If rsGrnHdr.GetNoRows > 0 Then
                    rsGrnHdr.MoveFirst()
                    strGRNDate = IIf(IsDBNull(rsGrnHdr.GetValue("grn_date")), "", VB6.Format(rsGrnHdr.GetValue("grn_date"), gstrDateFormat))
                    strVendorInvDate = IIf(IsDBNull(rsGrnHdr.GetValue("invoice_date")), "", VB6.Format(rsGrnHdr.GetValue("invoice_date"), gstrDateFormat))
                    strVendorInvNo = rsGrnHdr.GetValue("Invoice_No")
                End If
                rsGrnHdr.ResultSetClose()
            End If
            RdAddSold.DataDefinition.FormulaFields("GrinDate").Text = "'" + strGRNDate + "'"
            RdAddSold.DataDefinition.FormulaFields("GrinInvoiceNo").Text = "'" + strVendorInvNo + "'"
            RdAddSold.DataDefinition.FormulaFields("GrinInvoiceDate").Text = "'" + strVendorInvDate + "'"
        End If
        oCmd = New ADODB.Command
        With oCmd
            .ActiveConnection = mP_Connection
            .CommandTimeout = 0
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            'added by ekta uniyal
            '.CommandText = "PRC_INVOICEPRINTING"
            .CommandText = "PRC_INVOICEPRINTING_HILEX"
            'End Here
            .Parameters.Append(.CreateParameter("@UNITCODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
            .Parameters.Append(.CreateParameter("@LOC_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 8, Trim(txtUnitCode.Text)))
            .Parameters.Append(.CreateParameter("@DOC_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 8, Trim(mInvNo)))
            .Parameters.Append(.CreateParameter("@INV_TYPE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 3, Trim(Me.lbldescription.Text)))
            .Parameters.Append(.CreateParameter("@INV_SUBTYPE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Me.lblcategory.Text)))
            .Parameters.Append(.CreateParameter("@IP_ADDRESS", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 16, gstrIpaddressWinSck))
            .Parameters.Append(.CreateParameter("@ERRCODE", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInputOutput))
            .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        End With

        If oCmd.Parameters("@ERRCODE").Value <> 0 Then
            MsgBox("Error encountered while generating data for report.Please try Again.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, My.Resources.resEmpower.STR100)
            oCmd = Nothing
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
            Exit Sub
        End If
        oCmd = Nothing

        '' .ReportFileName = My.Application.Info.DirectoryPath & "\Reports\" & mstrReportFilename & ".rpt"
        '' .Destination = Crystal.DestinationConstants.crptToPrinter
        '' .SelectionFormula = "{TMP_INVOICEPRINT_SMIEL.IP_ADDRESS}='" & gstrIpaddressWinSck.Trim() & "'"
        ''  .Action = 1
        ''  End With
        RdAddSold.DataDefinition.RecordSelectionFormula = "{TMP_INVOICEPRINT_SMIEL.IP_ADDRESS}='" & gstrIpaddressWinSck.Trim() & "' and {TMP_INVOICEPRINT_SMIEL.unit_code}='" & gstrUNITID & "'"
        Frm.SetReportDocument()
        RdAddSold.PrintToPrinter(1, False, 0, 0)

        Exit Sub
ErrHandler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Function VerifyGLCCMappingFlag(ByVal Glcode As String) As Boolean

        On Error GoTo ErrHandler
        Dim rsVerifyGLCC_CODEFlag As ClsResultSetDB_Invoice
        rsVerifyGLCC_CODEFlag = New ClsResultSetDB_Invoice

        rsVerifyGLCC_CODEFlag.GetResult("select isnull(GLM_GLCODE,0) as GLM_GLCODE from FIN_GLMASTER WHERE UNIT_CODE='" + gstrUNITID + "' AND  GLM_GLCODE='" & Glcode & "'")
        If rsVerifyGLCC_CODEFlag.GetNoRows > 0 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object rsVerifyGLCC_CODEFlag.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            VerifyGLCCMappingFlag = True
        Else
            VerifyGLCCMappingFlag = False
        End If
        rsVerifyGLCC_CODEFlag.ResultSetClose()

        Exit Function
ErrHandler:
        VerifyGLCCMappingFlag = False
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function GetCommonCCCode(ByVal PurposeCode As String) As String

        Dim objRecordSet As New ADODB.Recordset
        Dim strCCCode As String

        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close()
        objRecordSet.Open("SELECT gbl_ccCode FROM fin_globalgl WHERE UNIT_CODE='" + gstrUNITID + "' AND  gbl_prpsCode = '" & PurposeCode & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
        If objRecordSet.EOF Then
            GetCommonCCCode = "N"
            MsgBox("CC Code not defined for  : " & PurposeCode, MsgBoxStyle.Information, "eMPro")
            If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                objRecordSet.Close()

                objRecordSet = Nothing
            End If
            Exit Function
        Else

            strCCCode = Trim(IIf(IsDBNull(objRecordSet.Fields("gbl_ccCode").Value), "", objRecordSet.Fields("gbl_ccCode").Value))

        End If
        If strCCCode = "" Then
            GetCommonCCCode = "N"
            MsgBox("CC Code not defined  for Purpose Code:" & PurposeCode, MsgBoxStyle.Information, "eMPro")
            If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                objRecordSet.Close()

                objRecordSet = Nothing
            End If
            Exit Function
        End If
        GetCommonCCCode = strCCCode
        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
            objRecordSet.Close()

            objRecordSet = Nothing
        End If
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        GetCommonCCCode = "N"
        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
            objRecordSet.Close()

            objRecordSet = Nothing
        End If
    End Function
    Private Sub PrintViaBatchFile(ByVal strFilename As String, ByVal strparameter As String)
        'issue id : 10158952
        Dim SProcess As New System.Diagnostics.Process
        Try
            'SProcess.StartInfo.FileName = gStrDriveLocation & "TypeToPrn.bat"
            'SProcess.StartInfo.Arguments = "C:\InvoicePrint.txt"
            SProcess.StartInfo.FileName = strFilename
            SProcess.StartInfo.Arguments = strparameter
            'SProcess.StartInfo.WindowStyle = ProcessWindowStyle.Hidden
            SProcess.Start()
        Catch Ex As Exception
            MsgBox(Ex.Message, MsgBoxStyle.Information, ResolveResString(100))
        Finally
            If SProcess.HasExited = False Then
                SProcess.WaitForExit(8000)
                Try
                    SProcess.Kill()
                Catch Ex As Exception
                End Try
            End If
        End Try
    End Sub
    '    Private Function FordASNFileGeneration(ByVal pintdocno As Integer, ByVal pstraccountcode As String) As Boolean
    '        'Revised By     : Manoj Kr. Vaish
    '        'Revised On     : 14 May 2009
    '        'Arguments      : INvoice No
    '        'Issue ID       : eMpro-20090513-31282
    '        'Reason         : Generate ASN File for FORD
    '        '--------------------------------------------------------------------------------------
    '        On Error GoTo ErrHandler

    '        Dim rsgetData As New ClsResultSetDB_Invoice
    '        Dim strquery As String
    '        Dim strASNdata As String
    '        Dim Dcount As Integer
    '        Dim TotalQty As Double
    '        Dim dblSalesQty As Double
    '        Dim strcontainerdespQty As String
    '        Dim dblcummulativeQty As Double
    '        Dim dblContainerQty As Double
    '        Dim strTotalQty As String
    '        Dim strASNFilepath As String
    '        Dim strASNFilepathforEDI As String

    '        strASNdata = ""
    '        strquery = "select * from dbo.FN_GETASNDETAIL(" & pintdocno & ",'" & gstrUNITID & "')"
    '        rsgetData.GetResult(strquery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
    '        If rsgetData.GetNoRows > 0 Then
    '            rsgetData.MoveFirst()
    '            Do While Not rsgetData.EOFRecord

    '                dblcummulativeQty = 0
    '                dblSalesQty = 0
    '                dblContainerQty = 0
    '                dblcummulativeQty = Find_Value("SELECT DBO.UDF_GET_CUMMULATIVEQTY_Hilex('" & gstrUNITID & "','" & rsgetData.GetValue("CUST_PLANTCODE").ToString() & "','" & rsgetData.GetValue("CUST_PART_CODE").ToString() & "'," & pintdocno & ")")

    '                dblSalesQty = rsgetData.GetValue("SALES_QUANTITY")
    '                dblcummulativeQty = dblcummulativeQty + dblSalesQty
    '                dblContainerQty = rsgetData.GetValue("CONTAINER_QTY")

    '                dblSalesQty = rsgetData.GetValue("Sales_Quantity")

    '                strTotalQty = Val(strTotalQty) + Val(rsgetData.GetValue("SALES_QUANTITY"))

    '                mstrupdateASNdtl = Trim(mstrupdateASNdtl) & "UPDATE MKT_ASN_INVDTL SET ASN_STATUS=1,CUMMULATIVE_QTY=" & dblcummulativeQty & " WHERE DOC_NO=" & pintdocno & " AND CUST_PART_CODE='" & rsgetData.GetValue("CUST_PART_CODE").ToString().Trim() & "' AND CUST_PLANTCODE='" & rsgetData.GetValue("CUST_PLANTCODE").ToString().Trim & "' AND UNIT_CODE='" & gstrUNITID & "'" & vbCrLf
    '                mstrupdateASNCumFig = Trim(mstrupdateASNCumFig) & "UPDATE MKT_ASN_CUMFIG SET CUMMULATIVE_QTY=" & dblcummulativeQty & " WHERE CUST_PART_CODE='" & rsgetData.GetValue("CUST_PART_CODE").ToString().Trim() & "' AND CUST_PLANTCODE='" & rsgetData.GetValue("CUST_PLANTCODE").ToString().Trim & "' AND UNIT_CODE='" & gstrUNITID & "'" & vbCrLf
    '                'Changed for Issue Id eMpro-20090709-33409 Ends
    '                rsgetData.MoveNext()
    '            Loop

    '            Dcount = Dcount + 1
    '            rsgetData.ResultSetClose()
    '            rsgetData = Nothing

    '            FordASNFileGeneration = True
    '        Else
    '            MessageBox.Show("Unable To Generate ASN File. Invoice Can't Be Locked", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
    '            FordASNFileGeneration = False
    '        End If

    '        Exit Function
    'ErrHandler:
    '        FordASNFileGeneration = False
    '        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    '    End Function

    Private Function FordASNFileGeneration(ByVal pintdocno As Integer) As Boolean
        '--------------------------------------------------------------------------------------
        'Modified By    : Shubhra Verma 
        'Modified On    : 06 Jun 2016
        '--------------------------------------------------------------------------------------

        Dim rsgetData As New ClsResultSetDB_Invoice
        Dim strASNdata As String = String.Empty
        Dim strupdateASNdtl As String = String.Empty
        Dim strupdateASNCumFig As String = String.Empty
        Dim strASNFilepath As String = String.Empty
        Dim strASNFilepathforEDI As String = String.Empty
        Dim strSQL As String
        Dim dblcummulativeQty As Double = 0
        Dim intCtr As Integer = 0
        Dim fs As FileStream
        Dim sw As StreamWriter

        Try

            strSQL = "select * from DBO.UFN_FORDASN('" & gstrUNITID & "'," & pintdocno & ")"
            rsgetData.GetResult(strSQL)

            If rsgetData.GetNoRows > 0 Then
                rsgetData.MoveFirst()

                Do While Not rsgetData.EOFRecord
                    strASNdata = strASNdata & rsgetData.GetValue("DataString").ToString + vbCrLf

                    strSQL = "SELECT DBO.UDF_GET_CUMMULATIVEQTY('" & gstrUNITID & "'," &
                        " '" & rsgetData.GetValue("CUST_PLANTCODE").ToString() & "'," &
                        " '" & rsgetData.GetValue("Cust_Item_Code").ToString() & "'," & pintdocno & ")"

                    dblcummulativeQty = Find_Value(strSQL)

                    strupdateASNdtl = "UPDATE MKT_ASN_INVDTL SET ASN_STATUS=1,CUMMULATIVE_QTY=" & dblcummulativeQty & "" &
                        " WHERE UNIT_CODE = '" & gstrUNITID & "' AND DOC_NO=" & pintdocno & "" &
                        " AND CUST_PART_CODE='" & rsgetData.GetValue("Cust_Item_Code").ToString() & "'" &
                        " AND CUST_PLANTCODE='" & rsgetData.GetValue("CUST_PLANTCODE").ToString().Trim & "'"

                    mP_Connection.Execute(strupdateASNdtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

                    strupdateASNCumFig = strupdateASNCumFig & "UPDATE MKT_ASN_CUMFIG SET CUMMULATIVE_QTY=" & dblcummulativeQty & "" &
                        " WHERE UNIT_CODE = '" & gstrUNITID & "'" &
                        " AND CUST_PART_CODE='" & rsgetData.GetValue("CUST_ITEM_CODE").ToString().Trim() & "'" &
                        " AND CUST_PLANTCODE='" & rsgetData.GetValue("CUST_PLANTCODE").ToString().Trim & "'"

                    mP_Connection.Execute(strupdateASNCumFig, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    rsgetData.MoveNext()
                Loop

                rsgetData.ResultSetClose()
                rsgetData = Nothing

                gstrASNPath = gstrUserMyDocPath
                gstrASNPathForEDI = ReadValueFromINI(Application.StartupPath & "\mind.cfg", "ASNPATH-" & gstrUNITID, "FilepathforEDI")

                If Directory.Exists(gstrASNPath) = False Then
                    Directory.CreateDirectory(gstrASNPath)
                End If

                If Directory.Exists(gstrASNPathForEDI) = False Then
                    Directory.CreateDirectory(gstrASNPathForEDI)
                End If

                strASNFilepath = gstrASNPath & "\" & mInvNo.ToString() & ".dat"
                strASNFilepathforEDI = gstrASNPathForEDI & "\" & mInvNo.ToString() & ".dat"
                fs = File.Create(strASNFilepath)
                sw = New StreamWriter(fs)
                sw.WriteLine(strASNdata)
                sw.Close()
                fs.Close()

                If File.Exists(strASNFilepathforEDI) = False Then
                    File.Copy(strASNFilepath, strASNFilepathforEDI)
                End If

                FordASNFileGeneration = True
            Else
                MessageBox.Show("Unable To Generate ASN File. Invoice Can't Be Locked", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                FordASNFileGeneration = False
            End If

        Catch EX As Exception
            FordASNFileGeneration = False
            RaiseException(EX)
        Finally

        End Try

    End Function

    Private Function AllowASNTextFileGeneration(ByVal pstraccountcode As String) As Boolean
        On Error GoTo ErrHandler

        Dim strQry As String
        Dim Rs As ClsResultSetDB_Invoice
        AllowASNTextFileGeneration = False

        'Change for Issue ID eMpro-20090624-32847 Starts
        If (UCase(Trim(cmbInvType.Text)) = "NORMAL INVOICE" And UCase(Trim(CmbCategory.Text)) = "FINISHED GOODS") Then
            strQry = "Select isnull(AllowASNTextGeneration,0) as AllowASNTextGeneration from customer_mst where Customer_Code='" & Trim(pstraccountcode) & "' And Unit_Code='" & gstrUNITID & "'"
            Rs = New ClsResultSetDB_Invoice
            If Rs.GetResult(strQry) = False Then GoTo ErrHandler

            If Rs.GetValue("AllowASNTextGeneration") = "True" Then
                AllowASNTextFileGeneration = True
            Else
                AllowASNTextFileGeneration = False
            End If

            Rs.ResultSetClose()
            Rs = Nothing
        End If

        Exit Function
ErrHandler:
        Rs = Nothing
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

    Private Sub txtASNNumber_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtASNNumber.TextChanged
        Try
            If IsNumeric(Me.txtASNNumber.Text.Trim) = False Then
                Me.txtASNNumber.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Function DUPLICATEASN(ByVal pstrASNNO As String) As String
        On Error GoTo ErrHandler

        Dim rsgetASNNumber As ClsResultSetDB_Invoice
        Dim strsql As String
        mblnDuplicateASNExist = False
        rsgetASNNumber = New ClsResultSetDB_Invoice
        strsql = "select doc_no from CreatedASN where ASN_NO='" & Trim(pstrASNNO) & "' and doc_no <> '" & Me.Ctlinvoice.Text & "' And Unit_Code='" & gstrUNITID & "'"

        rsgetASNNumber.GetResult(strsql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
        If rsgetASNNumber.GetNoRows > 0 Then
            DUPLICATEASN = IIf(IsDBNull(rsgetASNNumber.GetValue("doc_no")), "", rsgetASNNumber.GetValue("doc_no"))
            mblnDuplicateASNExist = True
        Else
            mblnDuplicateASNExist = False
        End If

        rsgetASNNumber.ResultSetClose()
        rsgetASNNumber = Nothing

        Exit Function
ErrHandler:
        rsgetASNNumber = Nothing
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)

    End Function
    Public Function AllowNumericValueInTextBox(ByRef TxtBox As TextBox, ByRef e As System.Windows.Forms.KeyPressEventArgs) As Int32
        Dim strNumbers As String = "0123456789." + vbBack
        Dim KeyAscii As Short = Asc(e.KeyChar)
        Try
            If InStr(TxtBox.Text, ".") = 0 Then
                If InStr(strNumbers & ".", Chr(KeyAscii)) = 0 Then KeyAscii = 0
            ElseIf InStr(strNumbers, Chr(KeyAscii)) <> 0 And KeyAscii = 46 Then
                KeyAscii = 0
            Else
                If InStr(strNumbers, Chr(KeyAscii)) = 0 Then KeyAscii = 0
            End If

            If Len(TxtBox.Text) = 0 And Chr(KeyAscii) = "." Then KeyAscii = 0
            e.KeyChar = Chr(KeyAscii)
            If KeyAscii = 0 Then
                e.Handled = True
            End If
        Catch ex As Exception
            Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Function
    Private Sub UPDATETRANFERINVOICE_HILEX(ByVal pstrdoc_no As Integer, ByVal pstrINV_no As Integer, ByVal pstrdoCTYPE As String)

        On Error GoTo ErrHandler

        Dim com As ADODB.Command
        com = New ADODB.Command

        com.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
        com.ActiveConnection = mP_Connection
        com.CommandText = "USP_BAR_UPDATE_TRF_INVOICE"
        com.Parameters.Append(com.CreateParameter("@UNIT_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
        com.Parameters.Append(com.CreateParameter("@TEMPINVNO", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, , pstrdoc_no))
        com.Parameters.Append(com.CreateParameter("@NEWINVNO", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, , pstrINV_no))
        com.Parameters.Append(com.CreateParameter("@INV_MODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 1, pstrdoCTYPE))

        com.Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

        com = Nothing
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Public Function CheckInvoices(ByVal strInvoiceno As String, ByVal straccountcode As String) As Boolean
        'ISSUE ID : 10341052
        Dim minv As Object

        On Error GoTo errhand
        Dim strSBUFolder As String
        Dim strHdr, strSql, mSuff As Object
        Dim strRec, strDtl, mfile As Object
        Dim sFolderName, VCode, strUserMsg As String
        Dim rs_hdr As ClsResultSetDB_Invoice
        Dim rs_dtl As ClsResultSetDB_Invoice
        Dim nInvWrite, intInvoice, intDSWrite As Short
        Dim mValue As Double
        Dim invnotopost As String
        Dim objGetDSData As ClsResultSetDB_Invoice
        Dim varDSFile As Object
        Dim strDSFileData As String
        Dim objFSO As Scripting.FileSystemObject
        Dim strInvList As String

        Dim strEDIFolder As String
        Dim minvEDIFile As String
        Dim mdsEDIFile As String
        Dim intdsEDIFile As Short
        Dim intinvEDIFile As Short
        Dim strLogMsg As String
        Dim pstrTempPath As String
        Dim pstrLocalDSPath As String
        Dim pstrTempPathForEDI As String
        Dim blnFTPwithEDI As Boolean
        Dim blnNewData As Boolean
        Dim strBuffer(14) As String
        Dim nInFile As Short 'File Handle of the arguments file
        '''strLogMsg = "Please Wait, Checking For New Invoices....."
        '''frmmsg.txtmsg.Text = frmmsg.txtmsg.Text & vbCrLf & strLogMsg
        '''PrintLine(nOutFile, strLogMsg)

        VCode = "M554"

        nInFile = FreeFile()
        'mstrins = ""

        FileOpen(nInFile, My.Application.Info.DirectoryPath & "\" & "FTParguments.cfg", OpenMode.Input)
        Dim counter As Short
        counter = 1
        'Read the arguments file and store all the values in strBuffer
        While Not EOF(nInFile)
            strBuffer(counter) = LineInput(nInFile)
            counter = counter + 1
        End While
        blnFTPwithEDI = GetArgumentValue(strBuffer(13))
        pstrTempPathForEDI = GetArgumentValue(strBuffer(14)) ' Path of temp. files for EDI
        pstrTempPath = GetArgumentValue(strBuffer(3)) ' Path of temp. files for FTP

        'FileClose(nInFile)
        'If Dir(pstrTempPath & "\Invoices", FileAttribute.Directory) = "" Then MkDir(pstrTempPath & "\Invoices\")

        'Check whether folder c:\temp exist or not for EDI
        If Not Directory.Exists(pstrTempPath & "\Invoices") Then
            Directory.CreateDirectory(pstrTempPath & "\Invoices")
        End If
        'If Dir(pstrTempPathForEDI & "\Invoices\", FileAttribute.Directory) = "" Then MkDir(pstrTempPathForEDI & "\Invoices\")
        If Not Directory.Exists(pstrTempPathForEDI & "\Invoices") Then
            Directory.CreateDirectory(pstrTempPathForEDI & "\Invoices")
        End If

        'strSql = "Select h.account_code,h.doc_no,h.suffix,mdt=Convert(char(10),h.invoice_date,103),h.cust_ref,h.sales_tax_amount"
        'strSql = strSql & ",b.schedulecode From saleschallan_dtl h  , customer_mst b Where h.unit_code=b.unit_code and h.ftp=0"
        'strSql = strSql & " And h.location_code='SML' and h.doc_no=" & strInvoiceno
        strSql = "select * from dbo.FN_GET_FTP_FILEDATA(" & strInvoiceno & ",'" & gstrUNITID & "')"
        rs_hdr = New ClsResultSetDB_Invoice
        rs_hdr.GetResult(strSql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)

        Dim rsGetAmount As ClsResultSetDB_Invoice
        Dim stramt As String
        Do While Not rs_hdr.EOFRecord
            strSBUFolder = pstrTempPath & "\Invoices\"
            strEDIFolder = pstrTempPathForEDI & "\Invoices\"

            If (rs_hdr.GetValue("schedulecode").ToString).Trim.Length = 0 Then
                strSBUFolder = strSBUFolder & "XXX\"
                strEDIFolder = strEDIFolder & "XXX\"
            Else
                If UCase(Trim(rs_hdr.GetValue("schedulecode").ToString)) = "U5J" Or UCase(Trim(rs_hdr.GetValue("schedulecode").ToString)) = "UJW" Or UCase(Trim(rs_hdr.GetValue("schedulecode").ToString)) = "19J" Then
                    strSBUFolder = strSBUFolder & "U09\"
                    strEDIFolder = strEDIFolder & "U09\"
                Else
                    strSBUFolder = strSBUFolder & Trim(rs_hdr.GetValue("schedulecode").ToString) & "\"
                    strEDIFolder = strEDIFolder & Trim(rs_hdr.GetValue("schedulecode").ToString) & "\"
                End If
            End If

            'If Dir(strSBUFolder, FileAttribute.Directory) = "" Then
            If Not Directory.Exists(strSBUFolder) Then
                Directory.CreateDirectory(strSBUFolder)
            End If
            If Not Directory.Exists(strEDIFolder) Then
                Directory.CreateDirectory(strEDIFolder)
            End If

            minv = strInvoiceno ' rs_hdr.GetValue("doc_no").ToString
            strInvList = strInvList & "'" & minv & "',"
            invnotopost = minv
            mSuff = ""
            invnotopost = CStr(Val(invnotopost))
            '**************.inv file*********************
            If UCase(Trim(rs_hdr.GetValue("schedulecode").ToString)) = "U5J" Or UCase(Trim(rs_hdr.GetValue("schedulecode").ToString)) = "UJW" Or UCase(Trim(rs_hdr.GetValue("schedulecode").ToString)) = "19J" Then
                mfile = Trim(strSBUFolder) & sFolderName & "U09INV" & Trim(VCode) & Trim(Str(CDbl(invnotopost))) & ".inv"
                minvEDIFile = Trim(strEDIFolder) & "U09INV" & Trim(VCode) & Trim(Str(CDbl(invnotopost))) & ".inv"
            Else
                mfile = Trim(strSBUFolder) & sFolderName & "" & Trim(rs_hdr.GetValue("schedulecode").ToString) & "INV" & Trim(VCode) & Trim(Str(CDbl(invnotopost))) & ".inv"
                minvEDIFile = Trim(strEDIFolder) & Trim(rs_hdr.GetValue("schedulecode").ToString) & "INV" & Trim(VCode) & Trim(Str(CDbl(invnotopost))) & ".inv"
            End If

            '**************.ds file*********************
            If UCase(Trim(rs_hdr.GetValue("schedulecode").ToString)) = "U5J" Or UCase(Trim(rs_hdr.GetValue("schedulecode").ToString)) = "UJW" Or UCase(Trim(rs_hdr.GetValue("schedulecode").ToString)) = "19J" Then
                varDSFile = Trim(strSBUFolder) & sFolderName & "U09INV" & Trim(VCode) & Trim(Str(CDbl(invnotopost))) & ".ds"
                mdsEDIFile = Trim(strEDIFolder) & "U09INV" & Trim(VCode) & Trim(Str(CDbl(invnotopost))) & ".ds"
            Else
                varDSFile = Trim(strSBUFolder) & sFolderName & "" & Trim(rs_hdr.GetValue("schedulecode").ToString) & "INV" & Trim(VCode) & Trim(Str(CDbl(invnotopost))) & ".ds"
                mdsEDIFile = Trim(strEDIFolder) & Trim(rs_hdr.GetValue("schedulecode").ToString) & "INV" & Trim(VCode) & Trim(Str(CDbl(invnotopost))) & ".ds"
            End If
            '************************************************************

            strHdr = Trim(rs_hdr.GetValue("account_code").ToString) & "|" & Trim(invnotopost) & "|" & rs_hdr.GetValue("INVOICE_DATE").ToString & "|" & Trim(invnotopost) & "|" & rs_hdr.GetValue("INVOICE_DATE").ToString & "|"

            strSql = "select b.cust_item_code,b.Rate, b.sales_quantity,b.to_box From sales_dtl b "
            strSql = strSql & " Where b.unit_code='" & gstrUNITID & "' and  b.doc_no=" & minv & " And b.suffix = '" & mSuff & "' "
            strSql = strSql & " Order by b.cust_item_code"

            rs_dtl = New ClsResultSetDB_Invoice
            rs_dtl.GetResult(strSql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)

            strDtl = ""
            mValue = 0
            Do While Not rs_dtl.EOFRecord
                strDtl = strDtl & Trim(rs_dtl.GetValue("cust_item_code").ToString) & "|" & Trim(Str(rs_dtl.GetValue("sales_quantity").ToString)) & "|" & Trim(Str(rs_dtl.GetValue("to_box").ToString)) & "^"
                mValue = mValue + (rs_dtl.GetValue("sales_quantity") * rs_dtl.GetValue("Rate"))
                'mstrins += "insert into invoiceupload_dtl values (" & minv & " , '" & rs_dtl.GetValue("cust_item_code").ToString & "' , " & rs_dtl.GetValue("sales_quantity").ToString & "," & rs_dtl.GetValue("rate").ToString & ",0,'" & gstrUNITID & "')"
                'mP_Connection.Execute(strins)
                rs_dtl.MoveNext()
            Loop

            '***Query to Fetch Data DS File**********
            strSql = "SELECT CUST_PART_CODE, DSNO,QUANTITYKNOCKEDOFF " & " FROM MKT_INVDSHISTORY H " & " INNER JOIN SALESCHALLAN_DTL SC ON SC.UNIT_CODE=H.UNIT_CODE AND SC.DOC_NO = H.DOC_NO " & " AND SC.LOCATION_CODE = H.LOCATION_CODE " & " AND SC.ACCOUNT_CODE = H.CUSTOMER_CODE" & " INNER JOIN SALES_DTL SD ON SC.UNIT_CODE=SD.UNIT_CODE AND SC.DOC_NO = SD.DOC_NO" & " AND SC.LOCATION_CODE = SD.LOCATION_CODE" & " AND SD.ITEM_CODE = H.ITEM_CODE" & " AND SC.UNIT_CODE='" & gstrUNITID & "' AND SC.DOC_NO = " & minv
            objGetDSData = New ClsResultSetDB_Invoice
            objGetDSData.GetResult(strSql)
            strDSFileData = ""
            Do While Not objGetDSData.EOFRecord
                strDSFileData = strDSFileData & objGetDSData.GetValue("CUST_PART_CODE").ToString & "|" & objGetDSData.GetValue("DSNO").ToString & "|" & objGetDSData.GetValue("QUANTITYKNOCKEDOFF").ToString & "^"
                objGetDSData.MoveNext()
            Loop
            objGetDSData.ResultSetClose()

            rsGetAmount = New ClsResultSetDB_Invoice
            stramt = " SELECT TOTAL_AMOUNT FROM SALESCHALLAN_DTL WHERE UNIT_CODE='" & gstrUNITID & "' AND DOC_NO=" & minv & ""
            rsGetAmount.GetResult(stramt, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
            If rsGetAmount.GetNoRows > 0 Then
                stramt = rsGetAmount.GetValue("total_amount").ToString
            End If

            'strHdr = strHdr & (stramt) & "|" & rs_hdr.GetValue("cust_ref")& "##"
            strHdr = strHdr & (stramt) & "|" & rs_hdr.GetValue("cust_ref").ToString & "##"
            rsGetAmount.ResultSetClose()
            rsGetAmount = Nothing


            'Combining Header & Detail information
            strRec = strHdr & strDtl

            ' ***----  working with text file
            nInvWrite = FreeFile()
            FileOpen(nInvWrite, mfile, OpenMode.Output)
            PrintLine(nInvWrite, strRec)
            FileClose(nInvWrite)

            '*********** Working in .DS file*************
            intDSWrite = FreeFile()
            FileOpen(intDSWrite, varDSFile, OpenMode.Output)
            PrintLine(intDSWrite, strDSFileData)
            FileClose(intDSWrite)

            objFSO = New Scripting.FileSystemObject
            ' ***----  working with text file for EDI
            objFSO.CopyFile(mfile, minvEDIFile)
            'intinvEDIFile = FreeFile()
            'FileOpen(intinvEDIFile, minvEDIFile, OpenMode.Output)
            'PrintLine(intinvEDIFile, strRec)
            'FileClose(intinvEDIFile)

            '*********** Working in .DS file*********for EDI
            objFSO.CopyFile(varDSFile, mdsEDIFile)
            'intdsEDIFile = FreeFile()
            'FileOpen(intdsEDIFile, mdsEDIFile, OpenMode.Output)
            'PrintLine(intdsEDIFile, strDSFileData)
            'FileClose(intdsEDIFile)
            objFSO = Nothing
            rs_dtl.ResultSetClose()


            'intInvoice = intInvoice + 1
            rs_hdr.MoveNext()
        Loop
        'strInvList = Left(strInvList, Len(strInvList) - 1)
        strInvList = Mid(strInvList, Len(Trim(strInvList)) + 1)
        If blnFTPwithEDI = False Then
            '    cn.BeginTrans()
            'cn.CommandTimeout = 0
            'cn.Execute("Update SalesChallan_Dtl set ftp = '1' Where doc_no in (" & strInvList & ")")
            'cn.Execute("Update invoiceupload_dtl set Updated=1 where doc_no in (" & strInvList & ")")
            'cn.CommitTrans()
        End If
        '**********Copy Folder In TempEDI folder***************************

        If rs_hdr.GetNoRows > 0 Then
            strLogMsg = Trim(Str(intInvoice)) & " Invoice(s) Found."
            blnNewData = True
        Else
            strLogMsg = "No Invoices Found For MSSL."
            blnNewData = False
        End If
        rs_hdr.ResultSetClose()

        FileClose(nInFile)
        CheckInvoices = True
        Exit Function
errhand:
        MsgBox("Error " & Err.Number & " ::" & Err.Description)
        strLogMsg = "Error " & Err.Description & " " & CStr(Now)
        FileClose(nInFile)
        CheckInvoices = False
    End Function
    Private Function GetArgumentValue(ByRef strArgument As String) As Object
        If InStr(strArgument, "~") > 0 Then
            GetArgumentValue = Trim(Mid(strArgument, 1, InStr(strArgument, "~") - 1))
        Else
            GetArgumentValue = Trim(strArgument)
        End If
    End Function
    Private Function AllowTextFileGeneration_SMIIEL(ByVal pstraccountcode As String) As Boolean
        On Error GoTo ErrHandler
        Dim strQry As String
        AllowTextFileGeneration_SMIIEL = False

        If DataExist("SELECT TOP 1 1 FROM CUSTOMER_MST where Customer_Code='" & Trim(pstraccountcode) & "' and UNIT_CODE = '" & gstrUNITID & "' and schedulecode is not null") = True Then
            AllowTextFileGeneration_SMIIEL = True
        Else
            AllowTextFileGeneration_SMIIEL = False
        End If

        Exit Function
ErrHandler:

        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

    Private Function Print_barcodelabel(ByVal pstrInvNo As String) As Boolean
        On Error GoTo ErrHandler

        Dim rsTKMLBARCODE As ClsResultSetDB_Invoice
        Dim rsSALESDTL As ClsResultSetDB_Invoice
        Dim rsSALESCHALLANDTL As ClsResultSetDB_Invoice
        Dim StrBarcodelabelFormat As String = String.Empty
        Dim StrNewlabel As String = String.Empty
        Dim SW As StreamWriter
        Dim intmaxRows As Integer
        Dim strBarcode As String = String.Empty
        Dim intRow As Integer
        Dim strtotalBasicAmount As String
        Dim strtotalExciseAmount As String
        Dim strtotalCVDAmount As String
        rsTKMLBARCODE = New ClsResultSetDB_Invoice
        rsSALESCHALLANDTL = New ClsResultSetDB_Invoice
        rsSALESDTL = New ClsResultSetDB_Invoice
        rsTKMLBARCODE.GetResult("SELECT ISNULL(TKML_BARCODELABEL_FORMAT,'') TKML_BARCODELABEL_FORMAT FROM SALES_PARAMETER where  UNIT_CODE= '" & gstrUNITID & "'")
        If rsTKMLBARCODE.GetNoRows > 0 Then
            StrBarcodelabelFormat = rsTKMLBARCODE.GetValue("TKML_BARCODELABEL_FORMAT").ToString
            SW = File.CreateText(gstrUserMyDocPath + "TKML_BARCODELABEL.TXT")
            StrNewlabel = StrBarcodelabelFormat
            rsSALESCHALLANDTL.GetResult("SELECT Distinct sd.srvdino,checksheetno,SD.doc_no,invoice_date,total_amount,Ecess_amount,Secess_amount, Sales_tax_amount FROM SALES_DTL SD ,SALESCHALLAN_DTL SC WHERE " &
                                    " SC.DOC_NO=SD.DOC_NO  AND SC.UNIT_CODE=SD.UNIT_CODE  AND  " &
                                    " SC.DOC_NO= " & pstrInvNo & " and SC.UNIT_CODE= '" & gstrUNITID & "'")
            intmaxRows = rsSALESCHALLANDTL.GetNoRows
            If intmaxRows > 0 Then
                strBarcode = rsSALESCHALLANDTL.GetValue("checksheetno").ToString & ","
                strBarcode = strBarcode + rsSALESCHALLANDTL.GetValue("doc_no").ToString & ","
                strBarcode = strBarcode + VB6.Format(rsSALESCHALLANDTL.GetValue("invoice_date"), "ddmmyy") & ","
                strtotalBasicAmount = Find_Value("SELECT SUM(ISNULL(BASIC_AMOUNT,0)) AS TOTALBASIC_AMOUNT FROM SALES_DTL WHERE UNIT_CODE= '" & gstrUNITID & "' and DOC_NO=" & pstrInvNo)
                strBarcode = strBarcode + VB6.Format(strtotalBasicAmount, "###.00").ToString & ","
                strtotalExciseAmount = Find_Value("SELECT SUM(ISNULL(EXCISE_TAX,0)) AS TOTALEXCISE_AMOUNT FROM SALES_DTL WHERE  UNIT_CODE= '" & gstrUNITID & "' and DOC_NO=" & pstrInvNo)

                If strtotalExciseAmount > 0 Then
                    strBarcode = strBarcode + VB6.Format(strtotalExciseAmount, "###").ToString & ","
                Else
                    strBarcode = strBarcode + "0" & ","
                End If


                strtotalCVDAmount = Find_Value("SELECT SUM(ISNULL(CVD_AMOUNT ,0)) AS CVD_AMOUNT FROM SALES_DTL WHERE UNIT_CODE= '" & gstrUNITID & "' and  DOC_NO=" & pstrInvNo)
                If strtotalCVDAmount > 0 Then
                    strBarcode = strBarcode + VB6.Format(strtotalCVDAmount, "###.0000").ToString & ","
                Else
                    strBarcode = strBarcode + "0.0000" & ","
                End If

                strBarcode = strBarcode + rsSALESCHALLANDTL.GetValue("ecess_amount").ToString & ","
                strBarcode = strBarcode + rsSALESCHALLANDTL.GetValue("secess_amount").ToString & ","
                If rsSALESCHALLANDTL.GetValue("Sales_tax_amount") > 0 Then
                    strBarcode = strBarcode + VB6.Format(rsSALESCHALLANDTL.GetValue("Sales_tax_amount").ToString, "###") & ","
                Else
                    strBarcode = strBarcode + "0.0000" & ","
                End If

                'strBarcode = strBarcode + VB6.Format(rsSALESCHALLANDTL.GetValue("total_amount").ToString, "###.00") & ","
                strBarcode = strBarcode + "1/1~"

                rsSALESDTL.GetResult("SELECT CUST_ITEM_CODE,SALES_QUANTITY  FROM SALES_DTL SD ,SALESCHALLAN_DTL SC WHERE " &
                                    " SC.DOC_NO=SD.DOC_NO  AND SC.UNIT_CODE=SD.UNIT_CODE  AND" &
                                    " SC.DOC_NO= " & pstrInvNo & " and SC.UNIT_CODE= '" & gstrUNITID & "' ")

                intmaxRows = rsSALESDTL.GetNoRows

                rsSALESDTL.MoveFirst()
                For intRow = 1 To intmaxRows
                    'strBarcode = strBarcode + rsSALESDTL.GetValue("Cust_Item_Code").ToString & "," + Val(rsSALESDTL.GetValue("sales_quantity")).ToString & "~"
                    strBarcode = strBarcode + rsSALESDTL.GetValue("Cust_Item_Code").ToString & "," + rsSALESDTL.GetValue("sales_quantity").ToString & "~"
                    rsSALESDTL.MoveNext()
                Next

            End If

            StrNewlabel = StrNewlabel.Replace("V_SRVDINO", rsSALESCHALLANDTL.GetValue("checksheetno").ToString)
            StrNewlabel = StrNewlabel.Replace("V_BARCODE", strBarcode)
            StrNewlabel = StrNewlabel.Replace("V_INVNO", rsSALESCHALLANDTL.GetValue("doc_no").ToString)
            StrNewlabel = StrNewlabel.Replace("V_Invdt", VB6.Format(rsSALESCHALLANDTL.GetValue("invoice_date"), "dd/mm/yyyy"))
            SW.WriteLine(StrNewlabel)
            '            print_barcodelabel = True
        Else
            Print_barcodelabel = False
        End If
        rsTKMLBARCODE.ResultSetClose()
        rsSALESDTL.ResultSetClose()
        rsSALESCHALLANDTL.ResultSetClose()
        SW.Close()
        If File.Exists(gstrUserMyDocPath + "TKML_BARCODELABEL.BAT") = False Then
            SW = File.CreateText(gstrUserMyDocPath + "TKML_BARCODELABEL.BAT")
            SW.WriteLine("CD\")
            SW.WriteLine("C:")
            SW.WriteLine("MODE:LPT1")
            SW.WriteLine("COPY """ & gstrUserMyDocPath & "TKML_BARCODELABEL.TXT"" LPT1")
            SW.Close()
        End If

        Shell("cmd.exe /c """ & gstrUserMyDocPath & "TKML_BARCODELABEL.BAT""", AppWinStyle.MinimizedNoFocus)
        MsgBox("Labels printed successfully.", MsgBoxStyle.Information, ResolveResString(100))
        Exit Function
ErrHandler:
        Print_barcodelabel = False
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

    Private Function ValidateInvoiceStockLocation() As Boolean
        '10804443 - MULTI LOCATION IN BARCODE - HILEX 
        Dim strMsg As String = String.Empty
        Try
            Using sqlcmd As SqlCommand = New SqlCommand
                With sqlcmd
                    .CommandText = "USP_VALIDATE_INVOICE_LOCATION"
                    .CommandType = CommandType.StoredProcedure
                    .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                    .Parameters.Add("@INVOICE_NO", SqlDbType.Int).Value = Val(Me.Ctlinvoice.Text)
                    .Parameters.Add("@MSG", SqlDbType.VarChar, 1000).Direction = ParameterDirection.Output
                    SqlConnectionclass.ExecuteNonQuery(sqlcmd)
                    strMsg = Convert.ToString(.Parameters("@MSG").Value)
                    If Len(strMsg) > 0 Then
                        MsgBox(strMsg, MsgBoxStyle.Exclamation, ResolveResString(100))
                        Return False
                    End If
                End With
            End Using
            Return True
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Private Function AllowA4Reports(ByVal pstraccountcode As String) As Boolean
        On Error GoTo ErrHandler

        Dim strQry As String
        Dim Rs As ClsResultSetDB_Invoice
        AllowA4Reports = False

        strQry = "Select isnull(AllowA4Reports,0) as AllowA4Reports from CUSTOMER_VENDOR_VW where UNIT_CODE='" + gstrUNITID + "' AND  Customer_Code='" & Trim(pstraccountcode) & "'"
        Rs = New ClsResultSetDB_Invoice
        If Rs.GetResult(strQry) = False Then GoTo ErrHandler

        If Rs.GetValue("AllowA4Reports") = "True" Then
            AllowA4Reports = True
        Else
            AllowA4Reports = False
        End If

        Rs.ResultSetClose()
        Rs = Nothing

        Exit Function
ErrHandler:
        Rs = Nothing
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function AllowCustomerspecificreport(ByVal pstraccountcode As String) As Boolean
        On Error GoTo ErrHandler

        Dim strQry As String
        Dim Rs As ClsResultSetDB_Invoice
        AllowCustomerspecificreport = False

        strQry = "Select isnull(AllowCustomerSpecificReport,0) as AllowCustomerSpecificReport , CustomerSpecificReportname from customer_mst where UNIT_CODE='" + gstrUNITID + "' AND  Customer_Code='" & Trim(pstraccountcode) & "'"
        Rs = New ClsResultSetDB_Invoice
        If Rs.GetResult(strQry) = False Then GoTo ErrHandler

        If Rs.GetValue("AllowCustomerSpecificReport") = "True" Then
            AllowCustomerspecificreport = True
            mstrReportFilename = Rs.GetValue("CustomerSpecificReportname").ToString
        Else
            AllowCustomerspecificreport = False
        End If

        Rs.ResultSetClose()
        Rs = Nothing

        Exit Function
ErrHandler:
        Rs = Nothing
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

    Private Function AllowBarCodePrinting(ByVal pstraccoutncode As String) As Boolean
        On Error GoTo ErrHandler
        Dim strQry As String
        Dim Rs As ClsResultSetDB_Invoice
        AllowBarCodePrinting = False
        strQry = "Select isnull(AllowBarCodePrinting,0) as AllowBarCodePrinting from customer_mst where Customer_Code='" & Trim(pstraccoutncode) & "' and UNIT_CODE = '" & gstrUNITID & "'"
        Rs = New ClsResultSetDB_Invoice
        If Rs.GetResult(strQry) = False Then GoTo ErrHandler
        If Rs.GetValue("AllowBarCodePrinting") = "True" Then
            AllowBarCodePrinting = True
        Else
            AllowBarCodePrinting = False
        End If
        Rs.ResultSetClose()
        Rs = Nothing
        Exit Function
ErrHandler:
        Rs = Nothing
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Public Function SaveBarCodeImage_singlelevelso(ByVal pstrInvNo As String, ByVal pstrCUSTITEMCODE As String, ByVal pstrItemcode As String, ByVal pstrPath As String, ByVal intcounter As Integer) As Boolean
        On Error GoTo ErrHandler
        Dim stimage As ADODB.Stream
        Dim strQuery As String
        Dim Rs As ADODB.Recordset
        SaveBarCodeImage_singlelevelso = True
        stimage = New ADODB.Stream
        stimage.Type = ADODB.StreamTypeEnum.adTypeBinary
        stimage.Open()
        pstrPath = pstrPath & "\BarcodeImg.bmp"
        stimage.LoadFromFile(pstrPath)
        If intcounter = 1 Then
            strQuery = "select  barCodeImage from saleschallan_dtl where doc_no='" & Trim(pstrInvNo) & "' and UNIT_CODE = '" & gstrUNITID & "'"
            Rs = New ADODB.Recordset
            Rs.Open(strQuery, mP_Connection, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
            Rs.Fields("barCodeImage").Value = stimage.Read
        End If
        If intcounter = 2 Then
            strQuery = "select  barCodeImage1 from saleschallan_dtl where doc_no='" & Trim(pstrInvNo) & "' and UNIT_CODE = '" & gstrUNITID & "'"
            Rs = New ADODB.Recordset
            Rs.Open(strQuery, mP_Connection, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
            Rs.Fields("barCodeImage1").Value = stimage.Read
        End If

        If intcounter = 3 Then
            strQuery = "select  barCodeImage2 from saleschallan_dtl where doc_no='" & Trim(pstrInvNo) & "' and UNIT_CODE = '" & gstrUNITID & "'"
            Rs = New ADODB.Recordset
            Rs.Open(strQuery, mP_Connection, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
            Rs.Fields("barCodeImage2").Value = stimage.Read
        End If

        If intcounter = 4 Then
            strQuery = "select  barCodeImage3 from saleschallan_dtl where doc_no='" & Trim(pstrInvNo) & "' and UNIT_CODE = '" & gstrUNITID & "'"
            Rs = New ADODB.Recordset
            Rs.Open(strQuery, mP_Connection, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
            Rs.Fields("barCodeImage3").Value = stimage.Read
        End If
        Rs.Update()
        Rs.Close()
        Rs = Nothing
        Exit Function

ErrHandler:
        SaveBarCodeImage_singlelevelso = False
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Public Function SaveBarCodeImage_singlelevelso_2DBARCODE(ByVal pstrInvNo As String, ByVal pstrPath As String, ByVal strbarcodestring As String) As Boolean
        On Error GoTo ErrHandler
        Dim stimage As ADODB.Stream
        Dim strQuery As String
        Dim Rs As ADODB.Recordset
        SaveBarCodeImage_singlelevelso_2DBARCODE = True
        stimage = New ADODB.Stream
        stimage.Type = ADODB.StreamTypeEnum.adTypeBinary
        stimage.Open()
        pstrPath = pstrPath & "BarcodeImg.JPEG"
        stimage.LoadFromFile(pstrPath)
        strQuery = "select  barcodeimage  from saleschallan_dtl where doc_no='" & Trim(pstrInvNo) & "' and UNIT_CODE = '" & gstrUNITID & "'"
        Rs = New ADODB.Recordset
        Rs.Open(strQuery, mP_Connection, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
        Rs.Fields("barcodeimage").Value = stimage.Read
        Rs.Update()
        Rs.Close()
        Rs = Nothing
        Exit Function

ErrHandler:
        SaveBarCodeImage_singlelevelso_2DBARCODE = False
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function IsGSTINSAME(ByVal strCustomerCode As String) As Boolean
        If SqlConnectionclass.ExecuteScalar("Select ISNULL(GSTIN_Id,'') GSTIN_Id From Customer_Mst Where UNIT_CODE='" & gstrUNITID & "' And Customer_Code='" & strCustomerCode & "'") = SqlConnectionclass.ExecuteScalar("Select ISNULL(GSTIN_ID,'') GSTIN_ID From Gen_UnitMaster Where Unt_CodeId='" & gstrUNITID & "'") Then
            Return True
        Else
            Return False
        End If
    End Function
    Private Sub chkPrintReprint_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkprintreprint.CheckedChanged
        If chkprintreprint.Checked = True Then
            Ctlinvoice.Text = ""
            _optInvYes_1.Text = "Reprint"
        Else
            Ctlinvoice.Text = ""
            _optInvYes_1.Text = "Print"
        End If
    End Sub
    Private Sub ReprintQRbarcode()

        On Error GoTo ErrHandler

        Dim rsGENERATEBARCODE As ClsResultSetDB_Invoice
        Dim straccountcode As String
        Dim strPrintMethod As String = ""
        Dim strSQL As String = ""
        Dim intTotalNoofSlabs As Integer = 0
        Dim intRow As Short
        Dim strBarcodeMsg As String
        Dim strBarcodeMsg_paratemeter As String
        Dim ObjBarcodeHMI As New Prj_BCHMI.cls_BCHMI(gstrUNITID)

        straccountcode = Find_Value("select account_code from saleschallan_dtl where UNIT_CODE = '" & gstrUNITID & "' and doc_no='" & Trim(Me.Ctlinvoice.Text) & "'")
        If AllowBarCodePrinting(straccountcode) = True And ChkQrbarcodereprint.Checked = True Then
            rsGENERATEBARCODE = New ClsResultSetDB_Invoice
            rsGENERATEBARCODE.GetResult("SELECT PRINT_METHOD FROM CUSTOMER_MST C WHERE C.UNIT_CODE='" & gstrUNITID & "' AND C.CUSTOMER_CODE='" & straccountcode & "'")
            strPrintMethod = UCase(rsGENERATEBARCODE.GetValue("PRINT_METHOD").ToString)
            rsGENERATEBARCODE.ResultSetClose()
            rsGENERATEBARCODE = Nothing

            If optInvYes(0).Checked = False Then 'only reprint 
                If strPrintMethod = "TOYOTA" Then
                    strSQL = "select  * from dbo.UFN_INV_BARCODE('" & gstrUNITID & "'," & Ctlinvoice.Text.Trim & ")"
                    rsGENERATEBARCODE = New ClsResultSetDB_Invoice
                    rsGENERATEBARCODE.GetResult(strSQL)
                    intTotalNoofSlabs = rsGENERATEBARCODE.GetNoRows
                    If intTotalNoofSlabs > 0 Then
                        rsGENERATEBARCODE.MoveFirst()
                        For intRow = 1 To intTotalNoofSlabs
                            strBarcodeMsg = ObjBarcodeHMI.GenerateQRBarCode(gstrUserMyDocPath, True, Trim(Ctlinvoice.Text), Trim(Ctlinvoice.Text), Val(rsGENERATEBARCODE.GetValue("FromLineNo").ToString), Val(rsGENERATEBARCODE.GetValue("ToLineNo").ToString), intRow, intTotalNoofSlabs, gstrCONNECTIONSTRING)
                            If VB.Left(strBarcodeMsg, 1) <> "Y" Then
                                MsgBox("Problem While Generating Barcode Image.", vbInformation, ResolveResString(100))
                                rsGENERATEBARCODE.ResultSetClose()
                                rsGENERATEBARCODE = Nothing
                                Exit Sub
                            Else
                                '10812364
                                strBarcodeMsg_paratemeter = Mid(strBarcodeMsg, 3)
                                '10812364
                                If UPDATEQRBarCodeImage(Trim(Ctlinvoice.Text), Val(rsGENERATEBARCODE.GetValue("FromLineNo").ToString), Val(rsGENERATEBARCODE.GetValue("ToLineNo").ToString), strBarcodeMsg_paratemeter, intRow, intTotalNoofSlabs, "") = False Then
                                    MsgBox("Problem While saving Barcode Image.", vbInformation, ResolveResString(100))
                                    rsGENERATEBARCODE.ResultSetClose()
                                    rsGENERATEBARCODE = Nothing
                                    Exit Sub
                                End If
                            End If
                            rsGENERATEBARCODE.MoveNext()
                        Next
                    End If
                    rsGENERATEBARCODE.ResultSetClose()
                    rsGENERATEBARCODE = Nothing
                    'changes 
                ElseIf strPrintMethod = "NORMAL" Then
                    If blnlinelevelcustomer = True Then
                        strBarcodeMsg = ObjBarcodeHMI.GenerateBarCode_LINELEVEL_SALESORDER_2dbarcode_Normal_Hilex(gstrUserMyDocPath, Ctlinvoice.Text, "NORMAL", "", "", True, Trim(Ctlinvoice.Text), gstrCONNECTIONSTRING)
                    Else
                        strBarcodeMsg = ObjBarcodeHMI.GenerateBarCode_LINELEVEL_SALESORDER_2dbarcode_hilex(gstrUserMyDocPath, Ctlinvoice.Text, "NORMAL", "", "", True, Trim(Ctlinvoice.Text), gstrCONNECTIONSTRING)
                    End If
                    Dim strQuery As String

                    If VB.Left(strBarcodeMsg, 1) <> "Y" Then
                        MsgBox("Problem While Generating Barcode Image.", vbInformation, ResolveResString(100))
                        Exit Sub
                    Else
                        If SaveBarCodeImage_singlelevelso_2DBARCODE(Ctlinvoice.Text, gstrUserMyDocPath, Mid(strBarcodeMsg, 3)) = False Then
                            MsgBox("Problem While saving Barcode Image.", vbInformation, ResolveResString(100))
                            Exit Sub
                        Else
                            mP_Connection.Execute(" UPDATE T SET T.BARCODEIMAGE =SC.BARCODEIMAGE FROM SALESCHALLAN_DTL SC,TMP_INVOICEPRINT  T " &
                                                   " WHERE SC.UNIT_CODE = T.UNIT_CODE AND SC.DOC_NO =T.DOC_NO AND SC.UNIT_CODE='" & gstrUNITID & "' AND " &
                                                   " SC.DOC_NO='" & Ctlinvoice.Text.Trim & "' AND T.IP_ADDRESS='" & gstrIpaddressWinSck & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        End If
                    End If
                    'changes
                Else

                    strBarcodeMsg = ObjBarcodeHMI.GenerateBarCode(gstrUserMyDocPath, mInvNo, True, Trim(Ctlinvoice.Text), gstrCONNECTIONSTRING)
                    If VB.Left(strBarcodeMsg, 1) <> "Y" Then
                        MsgBox("Problem While Generating Barcode Image.", vbInformation, ResolveResString(100))
                        Exit Sub
                    End If

                End If
            End If
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Function UPDATEQRBarCodeImage(ByVal pstrInvNo As String, ByVal intFromLineNo As Integer, ByVal intToLineNo As Integer, ByVal strbarcodestring As String, ByVal intRow As Integer, ByVal intTotalNoofSlabs As Integer,
                                        Optional ByVal pstrActualInvNo As String = "") As Boolean

        On Error GoTo ErrHandler

        Dim stimage As ADODB.Stream
        Dim strQuery As String
        Dim Rs As ADODB.Recordset
        Dim pstrPath As String = ""
        Dim strSql As String = ""
        Dim strserial As String = ""
        Dim strTEMPINVOICENO As String = ""
        Dim strdeleteSql As String = ""
        Dim blnCROP_QRIMAGE As Boolean = False
        pstrPath = gstrUserMyDocPath
        UPDATEQRBarCodeImage = True


        'If pstrActualInvNo.Trim.Length > 0 Then
        'strSql = "Select Top 1 1 from INVOICE_QRIMAGE WHERE unit_Code='" & gstrUNITID & "' And INVOICE_NO='" & pstrActualInvNo & "' And FROMLINENo=" & intFromLineNo & " AND TOLINENO=" & intToLineNo & ""
        'Else
        strSql = "Select Top 1 1 from INVOICE_QRIMAGE WHERE unit_Code='" & gstrUNITID & "' And INVOICE_NO='" & pstrInvNo & "'"

        'End If
        strserial = CStr(intRow) + "/" + CStr(intTotalNoofSlabs)

        If DataExist(strSql) = True Then

            strTEMPINVOICENO = Find_Value("select  TOP 1 TMP_INVOICENO from INVOICE_QRIMAGE where UNIT_CODE = '" & gstrUNITID & "' and invoice_no='" & Trim(Me.Ctlinvoice.Text) & "'")
            strdeleteSql = "DELETE  FROM INVOICE_QRIMAGE where UNIT_CODE = '" & gstrUNITID & "' and invoice_no='" & pstrInvNo & "' And FROMLINENo=" & intFromLineNo & " AND TOLINENO=" & intToLineNo & ""
            mP_Connection.Execute(strdeleteSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            strSql = "INSERT INTO INVOICE_QRIMAGE (INVOICE_NO,TMP_INVOICENO,FROMLINENO,TOLINENO,UNIT_CODE ,SERIALHEADING,BARCODESTRING ) VALUES('" & pstrInvNo & "','" & strTEMPINVOICENO & "'," & intFromLineNo & "," & intToLineNo & ",'" & gstrUNITID & "','" & strserial & "','" & strbarcodestring & "')"
            mP_Connection.Execute(strSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        End If

        stimage = New ADODB.Stream
        stimage.Type = ADODB.StreamTypeEnum.adTypeBinary
        stimage.Open()
        pstrPath = pstrPath & "QRBarcodeImg.wmf"
        blnCROP_QRIMAGE = CBool(Find_Value("SELECT CROP_QRBARCODE  FROM SALES_PARAMETER WHERE UNIT_CODE='" + gstrUNITID + "'"))
        If blnCROP_QRIMAGE = True Then
            Dim bmp As New Bitmap(pstrPath)
            Dim picturebox1 As New PictureBox
            picturebox1.Image = ImageTrim(bmp)
            picturebox1.Image.Save(pstrPath)
            picturebox1 = Nothing
        End If

        stimage.LoadFromFile(pstrPath)

        strQuery = "select  BARCODEIMG,INVOICE_NO  from INVOICE_QRIMAGE where UNIT_CODE = '" & gstrUNITID & "' AND INVOICE_NO='" & pstrInvNo & "' And FROMLINENO =" & intFromLineNo & " and TOLINENO =" & intToLineNo & ""
        Rs = New ADODB.Recordset
        Rs.Open(strQuery, mP_Connection, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)

        Rs.Fields("BARCODEIMG").Value = stimage.Read

        Rs.Update()
        Rs.Close()
        Rs = Nothing

        Exit Function
ErrHandler:
        UPDATEQRBarCodeImage = False
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function ImageTrim(ByVal img As Bitmap) As Bitmap
        'get image data
        Dim bd As BitmapData = img.LockBits(New Rectangle(Point.Empty, img.Size), ImageLockMode.[ReadOnly], PixelFormat.Format32bppArgb)
        Dim rgbValues As Integer() = New Integer(img.Height * img.Width - 1) {}
        Marshal.Copy(bd.Scan0, rgbValues, 0, rgbValues.Length)
        img.UnlockBits(bd)


        '#Region "determine bounds"
        Dim left As Integer = bd.Width
        Dim top As Integer = bd.Height
        Dim right As Integer = 0
        Dim bottom As Integer = 0

        'determine top
        For i As Integer = 0 To rgbValues.Length - 1
            Dim color As Integer = rgbValues(i) And &HFFFFFF
            If color <> &HFFFFFF Then
                Dim r As Integer = i / bd.Width
                Dim c As Integer = i Mod bd.Width

                If left > c Then
                    left = c
                End If
                If right < c Then
                    right = c
                End If
                bottom = r
                top = r
                Exit For
            End If
        Next

        'determine bottom
        For i As Integer = rgbValues.Length - 1 To 0 Step -1
            Dim color As Integer = rgbValues(i) And &HFFFFFF
            If color <> &HFFFFFF Then
                Dim r As Integer = i / bd.Width
                Dim c As Integer = i Mod bd.Width

                If left > c Then
                    left = c
                End If
                If right < c Then
                    right = c
                End If
                bottom = r
                Exit For
            End If
        Next

        If bottom > top Then
            For r As Integer = top + 1 To bottom - 1
                'determine left
                For c As Integer = 0 To left - 1
                    Dim color As Integer = rgbValues(r * bd.Width + c) And &HFFFFFF
                    If color <> &HFFFFFF Then
                        If left > c Then
                            left = c
                            Exit For
                        End If
                    End If
                Next

                'determine right
                For c As Integer = bd.Width - 1 To right + 1 Step -1
                    Dim color As Integer = rgbValues(r * bd.Width + c) And &HFFFFFF
                    If color <> &HFFFFFF Then
                        If right < c Then
                            right = c
                            Exit For
                        End If
                    End If
                Next
            Next
        End If

        Dim width As Integer = right - left + 1
        Dim height As Integer = bottom - top + 1
        '#End Region

        'copy image data
        Dim imgData As Integer() = New Integer(width * height - 1) {}
        For r As Integer = top To bottom
            Array.Copy(rgbValues, r * bd.Width + left, imgData, (r - top) * width, width)
        Next

        'create new image
        Dim newImage As New Bitmap(width, height, PixelFormat.Format32bppArgb)
        Dim nbd As BitmapData = newImage.LockBits(New Rectangle(0, 0, width, height), ImageLockMode.[WriteOnly], PixelFormat.Format32bppArgb)
        Marshal.Copy(imgData, 0, nbd.Scan0, imgData.Length)
        newImage.UnlockBits(nbd)

        ImageTrim = newImage
    End Function

    Private Sub Cmdinvoice_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cmdinvoice.Load

    End Sub

    Private Sub rptinvoice_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub IRN_QRBarcode()
        Try
            Dim rsGENERATEBARCODE As ClsResultSetDB_Invoice
            Dim straccountcode As String
            Dim strPrintMethod As String = ""
            Dim strSQL As String = ""
            Dim intTotalNoofSlabs As Integer = 0
            Dim intRow As Short
            Dim strBarcodeMsg As String
            Dim strBarcodeMsg_paratemeter As String
            Dim ObjBarcodeHMI As New Prj_BCHMI.cls_BCHMI(gstrUNITID)
            Dim stimage As ADODB.Stream
            Dim strQuery As String
            Dim Rs As ADODB.Recordset
            Dim pstrPath As String = ""
            Dim blnCROP_QRIMAGE As Boolean = False


            pstrPath = gstrUserMyDocPath
            strSQL = "SELECT TOP 1 1 FROM SALESCHALLAN_DTL_IRN I INNER JOIN SALESCHALLAN_DTL_IRN_BARCODE B ON I.UNIT_CODE=B.UNIT_CODE AND I.DOC_NO=B.DOC_NO WHERE I.UNIT_CODE = '" & gstrUNITID & "' AND I.DOC_NO='" & Trim(Me.Ctlinvoice.Text) & "'" & " AND ISNULL(I.IRN_NO,'')<>'' AND ISNULL(B.BARCODE_DATA,'')<>'' "

            If DataExist(strSQL) = True Then
                strBarcodeMsg = ObjBarcodeHMI.GenerateQRBarCodeForIRN(gstrUserMyDocPath, Trim(Ctlinvoice.Text), gstrCONNECTIONSTRING)

                If VB.Left(strBarcodeMsg, 1) <> "Y" Then
                    MsgBox("Problem While Generating Barcode Image.", vbInformation, ResolveResString(100))
                    Exit Sub
                Else
                    strBarcodeMsg_paratemeter = Mid(strBarcodeMsg, 3)
                    stimage = New ADODB.Stream
                    stimage.Type = ADODB.StreamTypeEnum.adTypeBinary
                    stimage.Open()
                    pstrPath = pstrPath & "QRBarcodeImgIRN.wmf"

                    blnCROP_QRIMAGE = CBool(Find_Value("SELECT CROP_IRN_QRBARCODE  FROM SALES_PARAMETER (NOLOCK) WHERE UNIT_CODE='" + gstrUNITID + "'"))
                    If blnCROP_QRIMAGE = True Then
                        Dim bmp As New Bitmap(pstrPath)
                        Dim picturebox1 As New PictureBox
                        picturebox1.Image = ImageTrim(bmp)
                        picturebox1.Image.Save(pstrPath)
                        picturebox1 = Nothing
                    End If

                    stimage.LoadFromFile(pstrPath)

                    strQuery = "select  BARCODE_DATA,Doc_No ,barcodeimage from SALESCHALLAN_DTL_IRN_BARCODE where UNIT_CODE = '" & gstrUNITID & "' AND Doc_No=" & Trim(Me.Ctlinvoice.Text)

                    Rs = New ADODB.Recordset
                    Rs.Open(strQuery, mP_Connection, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)

                    If Not (Rs.EOF And Rs.BOF) Then
                        Rs.Fields("barcodeimage").Value = stimage.Read
                        Rs.Update()
                    End If

                    Rs.Update()
                    Rs.Close()
                    Rs = Nothing


                End If

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Sub

    Private Sub btnExceptionInvoices_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExceptionInvoices.Click
        Try
            Dim objExceptionInvoices As New frmExceptionInvoices
            objExceptionInvoices.ShowDialog()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub cmdpdfpath_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdpdfpath.Click
        If (FBD.ShowDialog() = DialogResult.OK) Then
            txtPDFpath.Text = FBD.SelectedPath
        End If
    End Sub

    Private Sub _optInvYes_1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _optInvYes_1.CheckedChanged
        txtPDFpath.Text = ""
    End Sub
    Public Sub updatesalesconfandsaleschallan_Contingency(ByVal strPermanentNumber As String, ByVal strAssessableValue As String)
        Dim strSql As String
        Dim rsSalesChallan As ClsResultSetDB_Invoice
        Dim dblInvoiceAmt As Double
        Dim strInvoiceDate As String
        Dim strNewSaldConfNo As String
        On Error GoTo Err_Handler
        Dim objRet As Object = 0
        strSql = "select * from Saleschallan_dtl where Doc_No = " & strPermanentNumber.ToString()
        strSql = strSql & " and Invoice_type = '" & mInvType & "'  and  sub_category =  '" & mSubCat & "' and Location_Code='" & Trim(txtUnitCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "'"
        rsSalesChallan = New ClsResultSetDB_Invoice
        rsSalesChallan.GetResult(strSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
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
            If UCase(lbldescription.Text) <> "EXP" Then
                If Not mblnSameSeries Then
                    If IsGSTINSAME(mAccount_Code) And (lbldescription.Text.ToUpper = "TRF" Or lbldescription.Text.ToUpper = "ITD") Then

                        salesconf = "update saleconf set CURRENT_NO_TRF_SAMEGSTIN = " & mSaleConfNo.ToString & ", OpenningBal = openningBal - " & strAssessableValue & " where  UNIT_CODE = '" & gstrUNITID & "' and Invoice_type <> 'EXP' and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
                    Else
                        salesconf = "update saleconf set current_No = " & mSaleConfNo.ToString & ", OpenningBal = openningBal - " & strAssessableValue & " where  UNIT_CODE = '" & gstrUNITID & "' and Invoice_type <> 'EXP' and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
                    End If

                Else
                    If IsGSTINSAME(mAccount_Code) And (lbldescription.Text.ToUpper = "TRF" Or lbldescription.Text.ToUpper = "ITD") Then
                        salesconf = "update saleconf set CURRENT_NO_TRF_SAMEGSTIN = " & mSaleConfNo.ToString & " where  UNIT_CODE = '" & gstrUNITID & "' and Single_Series = 1 and Invoice_type <> 'EXP' and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0" & vbCrLf
                        salesconf = salesconf & "update saleconf set OpenningBal = openningBal - " & strAssessableValue & " where UNIT_CODE = '" & gstrUNITID & "' and  Invoice_type <> 'EXP' and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
                    Else
                        salesconf = "update saleconf set current_No = " & mSaleConfNo.ToString & " where  UNIT_CODE = '" & gstrUNITID & "' and Single_Series = 1 and Invoice_type <> 'EXP' and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0" & vbCrLf
                        salesconf = salesconf & "update saleconf set OpenningBal = openningBal - " & strAssessableValue & " where UNIT_CODE = '" & gstrUNITID & "' and  Invoice_type <> 'EXP' and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
                    End If

                End If
            Else
                salesconf = "update saleconf set current_No = " & mSaleConfNo.ToString & " where  UNIT_CODE = '" & gstrUNITID & "' and Invoice_type = 'EXP' and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
            End If
        Else

            If DataExist("SELECT TOP 1 1 FROM SALES_PARAMETER  WHERE SINGLE_INVOICE_SERIES= 1 and UNIT_CODE='" + gstrUNITID + "'") Then
                If Not mblnSameSeries Then
                    '10869291
                    If lbldescription.Text.ToUpper = "TRF" Or lbldescription.Text.ToUpper = "ITD" Then
                        If IsGSTINSAME(mAccount_Code) And (lbldescription.Text.ToUpper = "TRF" Or lbldescription.Text.ToUpper = "ITD") Then
                            salesconf = "update saleconf set CURRENT_NO_TRF_SAMEGSTIN = " & mSaleConfNo.ToString & " WHERE UNIT_CODE IN( SELECT UNIT_CODE FROM SALES_PARAMETER WHERE SINGLE_INVOICE_SERIES=1)  AND  Invoice_type IN('TRF','ITD') and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
                        Else
                            salesconf = "update saleconf set current_No = " & mSaleConfNo.ToString & " WHERE UNIT_CODE IN( SELECT UNIT_CODE FROM SALES_PARAMETER WHERE SINGLE_INVOICE_SERIES=1)  AND  Invoice_type IN('TRF','ITD') and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
                        End If
                    Else
                        If IsGSTINSAME(mAccount_Code) And (lbldescription.Text.ToUpper = "TRF" Or lbldescription.Text.ToUpper = "ITD") Then
                            salesconf = "update saleconf set CURRENT_NO_TRF_SAMEGSTIN = " & mSaleConfNo.ToString & " WHERE UNIT_CODE IN( SELECT UNIT_CODE FROM SALES_PARAMETER WHERE SINGLE_INVOICE_SERIES=1)  AND  Invoice_type = '" & Me.lbldescription.Text & "' " & " and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
                        Else
                            salesconf = "update saleconf set current_No = " & mSaleConfNo.ToString & " WHERE UNIT_CODE IN( SELECT UNIT_CODE FROM SALES_PARAMETER WHERE SINGLE_INVOICE_SERIES=1)  AND  Invoice_type = '" & Me.lbldescription.Text & "' " & " and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
                        End If

                    End If
                Else
                    salesconf = "update saleconf set current_No = " & mSaleConfNo.ToString & " WHERE  UNIT_CODE IN( SELECT UNIT_CODE FROM SALES_PARAMETER WHERE SINGLE_INVOICE_SERIES=1)  AND Single_Series = 1 " & " and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
                End If
            Else
                If Not mblnSameSeries Then
                    '10869291
                    If lbldescription.Text.ToUpper = "TRF" Or lbldescription.Text.ToUpper = "ITD" Then
                        If IsGSTINSAME(mAccount_Code) And (lbldescription.Text.ToUpper = "TRF" Or lbldescription.Text.ToUpper = "ITD") Then
                            salesconf = "update saleconf set CURRENT_NO_TRF_SAMEGSTIN = " & mSaleConfNo.ToString & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_type in ('ITD','TRF') and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
                        Else
                            salesconf = "update saleconf set current_No = " & mSaleConfNo.ToString & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_type in ('ITD','TRF') and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
                        End If
                    Else
                        If IsGSTINSAME(mAccount_Code) And (lbldescription.Text.ToUpper = "TRF" Or lbldescription.Text.ToUpper = "ITD") Then
                            salesconf = "update saleconf set CURRENT_NO_TRF_SAMEGSTIN = " & mSaleConfNo.ToString & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_type = '" & Me.lbldescription.Text & "' and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
                        Else
                            salesconf = "update saleconf set current_No = " & mSaleConfNo.ToString & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_type = '" & Me.lbldescription.Text & "' and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
                        End If
                    End If
                Else
                    If IsGSTINSAME(mAccount_Code) And (lbldescription.Text.ToUpper = "TRF" Or lbldescription.Text.ToUpper = "ITD") Then
                        salesconf = "update saleconf set CURRENT_NO_TRF_SAMEGSTIN = " & mSaleConfNo.ToString & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Single_Series = 1 and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
                    Else

                        salesconf = "update saleconf set current_No = " & mSaleConfNo.ToString & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Single_Series = 1 and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
                    End If

                End If
            End If
        End If
        mP_Connection.Execute("INSERT INTO INV_ERROR_DTL(QUERY,UNIT_CODE) VALUES('" & Replace(salesconf, "'", "") & "','" & gstrUNITID & "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        mP_Connection.Execute(salesconf, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        Exit Sub
Err_Handler:
        SqlConnectionclass.ExecuteNonQuery("INSERT INTO INV_ERR_PRIMARYKEY(ERR_NUMBER,ERR_DESCRIPTION,TMPINVOICENO,INVOICE_NO,UNIT_CODE,FUNCTIONNAME ) VALUES('" & Err.Number & "','" & Replace(Err.Description, "'", "") & "','" & Ctlinvoice.Text & "','" & mInvNo & "','" & gstrUNITID & "','updatesalesconfandsaleschallan_Contingency')")
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)

    End Sub
    Private Sub CustomRollbackTrans()
        Try
            mP_Connection.RollbackTrans()
        Catch ex As Exception
        End Try
    End Sub
    Private Function GetTransCountIfAvailable() As Integer
        Dim RTRANSCOUNT As ADODB.Recordset
        RTRANSCOUNT = mP_Connection.Execute("SELECT @@TRANCOUNT")
        GetTransCountIfAvailable = RTRANSCOUNT.Fields(0).Value
        RTRANSCOUNT = Nothing
    End Function
End Class