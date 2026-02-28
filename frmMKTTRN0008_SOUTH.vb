Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports System.IO
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Drawing.Text
Imports System.Data.SqlClient
Imports System.Drawing.Imaging
Imports System.Runtime.InteropServices
Imports System.Drawing.Drawing2D
Imports System.Drawing
Imports System.Collections
'Imports iTextSharp.text
'Imports iTextSharp.text.pdf

Imports iText.Kernel.Pdf
Imports iText.Layout
Imports iText.Layout.Element
Imports iText.IO.Image
Imports iText.Kernel.Geom
Imports iText.Kernel.Utils




Friend Class frmMKTTRN0008_SOUTH
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
    '===================================================================================
    'Changed By     : Arul Mozhi on 06-04-2004
    'Description    : Customer Supplied Material Rejection Invoice Provision Provided
    '===================================================================================
    'Changed By     : Arul Mozhi on 20-04-2004
    'Description    : To Avoide the Customer supplied Material Rejection Invoice Postings
    '===================================================================================
    'Changed By     : Arul Mozhi on 27-07-2004
    'Description    : A) To Check the GL & SL Linking for Customer Supplied Material Reejection Invoices
    '                 B) To Avoid the Basic Amount Posting for Customer Supplied Material Rejection Invoices
    '                 C) "rptInvoicerejcustomerMateCh" Seperate report prepared for Customer Supplied Material Rejection Invoices
    '===================================================================================
    'Changed By     : Arul Mozhi on 15-11-2004
    'Description    : Checking of Tool Amortization Sales posting required or not
    '====================================================================================
    'Changed By     : Arul Mozhi on 03-12-2004
    'Description    : Changes made for Customer Supplied Material Excise Value and Those Cess Value Posting to Seperate Account
    '====================================================================================
    'Code Changed By : Arul on 07-03-2005
    'Reason         : Multiple selection of sales Order in one Invoice Option provided to users
    '                 Emp_InvoiceSOLinkage table used to save the multiple Sales Order for a single invoice
    '====================================================================================
    'Changed By     : Arul Mozhi on 09-04-2005
    'Description    : To show the print setup Button on crystal report
    '====================================================================================
    'Changed By     : Arul Mozhi on 19-04-2005
    'Description    : Changes made for Vendor Rejection invoice posting goes to AP Module
    'Description    : Account Creation String for every items has been Cut & Pasted from form sended by Nisha
    '====================================================================================
    'Code Changed   : Arul Mozhi On 07-05-2005
    'Description    : Export Invoice Basic Currency Amount Posted wrongly in AR_Docmaster
    '                 Due to Rounding off Figures
    '=========================================================================================
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
    'Revised By     : Manoj Kr. Vaish
    'Revised On     : 25 Feb 2008
    'Issue ID       : 22486
    'Reason         : HMI bar Code Changes for Mate unit 3
    '--------------------------------------------------------------------------------------
    'Revised By     : Manoj Kr. Vaish
    'Revised On     : 26 MAY 2008
    'Issue ID       : eMpro-20080526-19267
    'Reason         : Preview button should be use only for Invoice Preview and Invoice shall be printed
    '                 only when the user clicks Print button while locking the invoice.
    '--------------------------------------------------------------------------------------
    'Revised By     : Manoj Kr. Vaish
    'Revised On     : 15 Jul 2008
    'Issue ID       : eMpro-20080715-20329
    'Reason         : Place a sleep command between BarCode Generation and SaveImage on db function
    '--------------------------------------------------------------------------------------
    'Revised By      : Manoj Kr.Vaish
    'Issue ID        : eMpro-20090209-27201
    'Revision Date   : 20 Mar 2009
    'History         : BatchWise Tracking of Invoices Made from 01M1 Location including BarCode Tracking
    '*******************************************************************************************************
    'Revised By      : Manoj Kr. Vaish
    'Revised On      : 24 Mar 2009
    'Issue ID        : eMpro-20090204-27027
    'Reason          : ASN File Printing for Mahindra & Mahindra-Mate 1
    '--------------------------------------------------------------------------------------
    'Revised By      : Manoj Kr.Vaish
    'Issue ID        : eMpro-20090513-31282
    'Revision Date   : 14 May 2009
    'History         : Intergeration of Ford ASN File Generation for Mate South Units
    '*******************************************************************************************************
    'Revised By      : Manoj Kr. Vaish
    'Issue ID        : eMpro-20090624-32847
    'Revision Date   : 24 Jun 2009
    'History         : Check for Ford ASN Generation only for Normal-Finished Good
    '****************************************************************************************
    'Revised By      : Manoj Kr. Vaish
    'Issue ID        : eMpro-20090703-33213
    'Revision Date   : 04 Jul 2009
    'History         : Credit term should be fetched from Saleschallan_dtl while locking the invoice
    '****************************************************************************************
    'Revised By      : Manoj Kr. Vaish
    'Issue ID        : eMpro-20090709-33409
    'Revision Date   : 09 Jul 2009
    'History         : Cummulative Qunatity mismatch problem in FORD ASN Invoice for Mate 1,2 and 4
    '****************************************************************************************
    'Revised By        : Siddharth Ranjan
    'Issue ID          : eMpro-20090709-33428
    'Revision Date     : 15 Jul 2009
    'History           : CSI functionality
    '****************************************************************************************
    'Revised By        : Siddharth Ranjan
    'Issue ID          : eMpro-20091026-37855
    'Revision Date     : 26 Oct 2009
    'History           : Wrong calculation of Cummulative qty in ASN for ford new logic implemented
    '****************************************************************************************
    'Revised By        : Siddharth Ranjan
    'Issue ID          : eMpro-20100108-40881
    'Revision Date     : 09 Dec 2009
    'History           : New CSM FIFO KnockedOff functionality
    '****************************************************************************************
    'Revised By      : PRASHANT RAJPAL
    'Revision On     : 05 APR 2011
    'Issue ID        : 1084018
    'History         : CC CODE CHANGES ADDED IN TRANSFER ,SAMPLE AND REJECTION INVOICE ( CONFIGURABLE FUNCTIONALITY )
    '******************************************************************************************************
    ' Revised By                 -   Roshan Singh
    ' Revision Date              -   06 JUN 2011
    ' Description                -   FOR MULTIUNIT FUNCTIONALITY
    '***********************************************************************************
    '****************************************************************************************
    'Revised By      : PRASHANT RAJPAL
    'Revision On     : 28 AUG 2011
    'Issue ID        : 10129872 
    'History         : ASN  functionality-for CUSTOMER REJECTION
    '----------------------------------------------------------------
    ' Revised By     :   Pankaj Kumar
    ' Revision Date  :   14 Oct 2011
    ' Description    :   Modified for MultiUnit Change Management
    '***********************************************************************************
    'Revised By      : PRASHANT RAJPAL
    'Revision On     : 21 OCT 2011
    'Issue ID        : 10151004  
    'History         : CC code GL Code FOR REJECTION INVOICE ERROR
    '***********************************************************************************
    'Revised By Roshan Singh on 09 Nov 2011 for Multi unit Change management
    '***********************************************************************************
    'Revised By      : PRASHANT RAJPAL
    'Revision On     : 16 Nov 2011
    'Issue ID        : 10160094   
    'History         : Changes for ASN Path for Multi-Unit 
    '***********************************************************************************
    'Modified By Roshan Singh on 19 Dec 2011 for multi unit Change Management.
    '***********************************************************************************
    'Revised By      : PRASHANT RAJPAL
    'Issue ID        : 10192547 
    'Revision Date   : 08 FEB 2012
    'History         : Changes in Invoice Entry FOR barcode process (At Main Store )
    '***********************************************************************************
    '-- Modified by Roshan Singh on 14 FEB 2012 for multiunit change

    '***********************************************************************************
    'Revised By      : PRASHANT RAJPAL
    'Issue ID        : 10245888  
    'Revision Date   : 06 JULY 2012
    'History         : Changes FOR ASN GENERATION : FORD FILE 
    '***********************************************************************************
    'Revised By      : PRASHANT RAJPAL
    'Issue ID        : 10213067  
    'Revision Date   : 15-july-2012 to 24-JULY-2012
    'History         : Changes in Invoice Entry :Temporary Invoice Printed with zero invoice number for Unlocked Invoice 
    '***********************************************************************************
    'Revised By      : PRASHANT RAJPAL
    'Issue ID        : 10178530     
    'Revision Date   : 24 july 2012
    'History         : Changes in Invoice Entry :-LRN concept for REJECTION INVOICE
    '***********************************************************************************
    'Revised By      : PRASHANT RAJPAL
    'Issue ID        : 10245698       
    'Revision Date   : 21 nov 2012
    'History         : Changes in Invoice Entry :-ASN for REJECTION INVOICE
    '***********************************************************************************
    'Revision Date      :    30 jan 2013
    'Revised By         :    Prashant Rajpal
    'Revision for       :    Issue No.10331478
    'purpose            :   EDI Functionality for Mate Unit3  export     
    '*****************************************************************************
    'Revision Date      :    07 Feb  2013
    'Revised By         :    Prashant Rajpal
    'Revision for       :    Issue No.10341569 
    'purpose            :   EDI Functionality for Mate Unit3  export :Transfer invoice
    '*****************************************************************************
    'Revised By      : PRASHANT RAJPAL
    'Issue ID        : 10346772  
    'Revision Date   : 19-Feb-2013-20-feb 2013
    'History         : Changes for SMIEL-FTP PROBLEM 
    '***********************************************************************************
    'Revised By      : PRASHANT RAJPAL
    'Issue ID        : 10365277  
    'Revision Date   : 01 apr 2013
    'History         : Changes for Invoice Printing master Record Insertion Failure 
    '***********************************************************************************
    'Revised By      : PRASHANT RAJPAL
    'Issue ID        : 10404992   
    'Revision Date   : 12-June-2013-13-June-2013
    'History         : Including TCS Value in Normal Invoice with sub type SCRAP
    '***********************************************************************************
    'REVISED BY     :  VINOD SINGH
    'REVISED DATE   :  30 AUG 2013
    'ISSUE ID       :  10378778
    'PURPOSE        :  GLOBAL TOOL CHAGES
    '************************************************************************************
    'Revised By      : PRASHANT RAJPAL
    'Issue ID        : 10455376  
    'Revision Date   : 25-OCT-2013
    'History         : REJECTION INVOICE : ASN CUMMS CAN BE ZERO (CONFIGURABLE : SALES_PARAMETER :ZEROASNCUMMS_REJECTION)
    '***********************************************************************************
    'Revised By      : PRASHANT RAJPAL
    'Issue ID        : 10564189   
    'Revision Date   : 13 May 2014
    'History         : Disable the Auto ASN for Rejection invoice and Normal Invoice (Configurable)
    '***************************************************************************************************
    'Revised By      : PRASHANT RAJPAL
    'Issue ID        : 10597202    
    'Revision Date   : 15 May 2014
    'History         : Single  Invoice series  for unit 3 : telecom , domestic and export 
    '***************************************************************************************************
    'Revised By      : PRASHANT RAJPAL
    'Issue ID        : 10617093 
    'Revision Date   : 20 june 2014
    'History         : A4 REPORT FOR MATE BGLORE WITH ANNEXTURE CONCEPT
    '***************************************************************************************************
    ' REVISION DATE     : 01 SEP2014
    ' REVISED BY        : PRASHANT RAJPAL
    ' ISSUE ID          : 10665764
    ' REVISION HISTORY  : SCHEDULE NOT UPDATED CORRECTLY
    '************************************************************************
    ' REVISION DATE     : 28-OCT 2014-29 OCT 2014
    ' REVISED BY        : PRASHANT RAJPAL
    ' ISSUE ID          : 10688760 
    ' REVISION HISTORY  : SHIPPING DETAILS DISPLAYED IN REPORTS (CONFIGURABLE )
    '************************************************************************
    'REVISED BY     :  PRASHANT RAJPAL
    'REVISED DATE   :  20-NOV-2014 - 21-NOV-2014
    'ISSUE ID       :  10706455  
    'PURPOSE        :  TO ADD ADDITIONAL VAT CALCULATION
    '****************************************************************************************
    ' REVISION DATE     : 15 JAN 2015 
    ' REVISED BY        : Abhinav Kumar 
    ' ISSUE ID          : 10736222  
    ' REVISION HISTORY  : CT2 - ARE3 functionality 
    '*********************************************************************************************************************************
    ' REVISION DATE     : 21 May 2015
    ' REVISED BY        : Prashant Rajpal
    ' ISSUE ID          : 10812364 
    ' REVISION HISTORY  : Qr Barcode missing in some invoices
    '*********************************************************************************************************************************
    'REVISED BY     -  PRASHANT RAJPAL
    'REVISED ON     -  18 JUNE 2015
    'PURPOSE        -  10825102 -A4 INVOICE PRINTING FUNCTIONALITY
    '-------------------------------------------------------------------------------------------------------------------------------------------------------
    'REVISED BY     -  PRASHANT RAJPAL
    'REVISED ON     -  25 AUG 2015
    'PURPOSE        -  10856126 -ASN CHANGES FOR LOGISTICS 
    '-------------------------------------------------------------------------------------------------------------------------------------------------------
    'REVISED BY     -  PRASHANT RAJPAL
    'REVISED ON     -  25 AUG 2015
    'PURPOSE        -  10869290  -SERVICE INVOICE FUNCTIONAITY
    '-------------------------------------------------------------------------------------------------------------------------------------------------------
    'REVISED BY     -  PRASHANT RAJPAL
    'REVISED ON     -  14 Jan 2016
    'PURPOSE        -  10910711  - ASSETS TRANSFER NOTE (ATN) CHANGES IN INVOICE
    '-------------------------------------------------------------------------------------------------------------------------------------------------------
    'REVISED BY     -  PRASHANT RAJPAL
    'REVISED ON     -  30 may 2016
    'PURPOSE        -  101041360  - EDI INVOICE FOR FSP
    '-------------------------------------------------------------------------------------------------------------------------------------------------------
    'REVISED BY     -  PRASHANT RAJPAL
    'REVISED ON     -  07 OCT 2016
    'PURPOSE        -  10869291 — eMPro- Inter-Division Invoice 

    'REVISED BY     -  ASHISH SHARMA    
    'REVISED ON     -  21 JUN 2017
    'PURPOSE        -  101188073 — GST CHANGES
    '-------------------------------------------------------------------------------------------------------------------------------------------------------
    'REVISED BY     -  ASHISH SHARMA
    'REVISED ON     -  10 SEP 2018
    'PURPOSE        -  101375632 - REG Bar code implementation - BM1 UNIT
    '-------------------------------------------------------------------------------------------------------------------------------------------------------
    'REVISED BY     -  ASHISH SHARMA
    'REVISED ON     -  11 AUG 2020
    'PURPOSE        -  102027599 - IRN CHANGES
    '-------------------------------------------------------------------------------------------------------------------------------------------------------
    '-------------------------------------------------------------------------------------------------------------------------------------------------------
    'REVISED BY     -  PRAVEEN KUMAR
    'REVISED ON     -  27 JULY 2023
    'PURPOSE        -  102853899  - DIGITAL SIGN CHANGES
    '-------------------------------------------------------------------------------------------------------------------------------------------------------

    Public gobjDB As New ClsResultSetDB_Invoice
    Dim mStrCustMst As String
    Dim mresult As ClsResultSetDB_Invoice
    Dim mintFormIndex As Short
    Dim salesconf As String
    Dim msubTotal, mInvNo, mExDuty, mBasicAmt, mOtherAmt As Double
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
    Dim strupdatecustodtdtl As String
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
    Dim mblnExciseRoundOFFFlag As Boolean
    Dim mSaleConfNo As Double
    Dim blnCSIEX_Inc As Boolean
    Dim mblnBOMcheck_flag As Boolean
    Dim CUST_REJECTION_FLAG As Boolean 'To store the Customer rejection invioice flag
    Dim strCustRef As String
    Dim strUpdateAmorDtl As String
    Dim strupdateamordtlbom As String
    Dim mstrupdateBarBondedStockFlag As String
    Dim mstrupdateBarBondedStockQty As String
    Dim mblnQuantityCheck As Boolean
    Dim mstrFGDomestic As String
    Dim mblnASNExist As Boolean
    Dim mstrupdateASNdtl As String
    Dim mstrupdateASNCumFig As String
    Dim mblnCSM_Knockingoff_req As Boolean
    Dim mblnLock_Clicked As Boolean
    Dim mblnCCFlag As Boolean
    Dim blnlinelevelcustomer As Boolean = False
    Dim mblnZEROASNCUMMS_REJECTIONINVOICE As Boolean
    '10617093 
    Dim mblncustomerlevel_A4report_functionlity As Boolean
    Dim mblncustomerspecificreport As Boolean
    Dim mblnA4reports_invoicewise As Boolean
    Dim intNoCopies_A4reports As Short
    Dim gstrIntNoCopies As Short
    Dim mblnAllowCustomerSpecificReport_COMP As Boolean
    Dim mblnAllowCustomerSpecificReport_RAW As Boolean
    Dim mblncustomerlevel_Annexture_printing As Boolean
    Dim mintmaxnoofitems_Annexture As Integer
    '10688760 
    Dim mbln_SHIPPING_ADDRESS As Boolean
    Dim mbln_SHIPPING_ADDRESS_INVOICEWISE As Boolean
    Dim mintmaxnoofitems_barcodeToyota As Integer
    '10869290
    Dim mblnServiceInvoicemate As Boolean = False
    Dim mblnTotalInvoiceAmountRoundOff As Boolean
    Dim mintTotalInvoiceAmountRoundOff As Short
    Dim mblnlorryno As Boolean
    Dim MSTRREJECTIONNOTE As String
    Dim mstracountcode As String
    Dim mlbnprintforexcurrency As Boolean = False
    Dim mblnEwaybill_Print As Boolean = False
    Dim mblnEWAY_BILL_STARTDATE As String
    Dim mdblewaymaximumvalue As Double
    Dim mblnTOYOTA_MULTIPLESO_ONEPDS_STDATE As String
    Dim mblnAnnextureExport As Boolean = False
    Dim PCA_SCHEDULEINVOICE As String
    Dim mCheckValARana As String = String.Empty
    Dim mPdfReaderPath As String = String.Empty
    Dim pdfPrintProcID As Integer = 0
    Dim mblnSEZ_Knockingoff_req As Boolean
    Dim blnisCreditLimitMandatory As Boolean = False
    Dim dblCreditLimit As Double = 0
    Dim dblTotalInvoiceAmt As Double = 0
    Dim dblOutstandingLimit As Double = 0


    Private Sub ChkCustDetails_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ChkCustDetails.Enter
        Me.ShpCustDetails.Visible = False
    End Sub

    Private Sub ChkCustDetails_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ChkCustDetails.Leave
        Me.ShpCustDetails.Visible = True
    End Sub

    Private Sub cmbInvType_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbInvType.TextChanged
        On Error GoTo ErrHandler
        Call cmbInvType_Validating(cmbInvType, New System.ComponentModel.CancelEventArgs(False))
        Exit Sub
ErrHandler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdUnitCodeList_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdUnitCodeList.Click
        On Error GoTo ErrHandler
        Call ShowCode_Desc("SELECT Unt_CodeID,unt_unitname FROM Gen_UnitMaster WHERE Unt_Status=1 and Unt_CodeID= '" & gstrUNITID & "'", txtUnitCode)
        Exit Sub
ErrHandler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtUnitCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUnitCode.TextChanged
        On Error GoTo ErrHandler
        If Trim(txtUnitCode.Text) = "" Then
            cmbInvType.Enabled = False
            cmbInvType.SelectedIndex = -1
            CmbCategory.Enabled = False
            CmbCategory.SelectedIndex = -1
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtUnitCode_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUnitCode.Enter
        On Error GoTo ErrHandler
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
        If Shift <> 0 Then Exit Sub
        If KeyCode = System.Windows.Forms.Keys.F1 Then Call cmdUnitCodeList_Click(cmdUnitCodeList, New System.EventArgs())
        Exit Sub 'This is to avoid the execution of the error handler
        If KeyCode = Keys.Enter Then System.Windows.Forms.SendKeys.Send("{TAB}")
ErrHandler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub txtUnitCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtUnitCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then
            System.Windows.Forms.SendKeys.Send("{TAB}")
        ElseIf KeyAscii = 187 Or KeyAscii = 166 Or KeyAscii = 164 Or KeyAscii = 172 Or KeyAscii = 39 Or KeyAscii = 34 Or KeyAscii = 96 Then
            KeyAscii = 0
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtUnitCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtUnitCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        Dim strUnitDesc As String
        Dim mobjGLTrans As prj_GLTransactions.cls_GLTransactions
        If Trim(txtUnitCode.Text) = "" Then GoTo EventExitSub
        mobjGLTrans = New prj_GLTransactions.cls_GLTransactions(gstrUNITID, GetServerDate)
        strUnitDesc = mobjGLTrans.GetUnit(Trim(txtUnitCode.Text), ConnectionString:=gstrCONNECTIONSTRING)
        mobjGLTrans = Nothing
        If CheckString(strUnitDesc) <> "Y" Then
            MsgBox(CheckString(strUnitDesc), MsgBoxStyle.Critical, "empower")
            txtUnitCode.Text = ""
            cmbInvType.Enabled = True
            cmbInvType.SelectedIndex = -1
            CmbCategory.Enabled = True
            CmbCategory.SelectedIndex = -1
            Cancel = True
        Else
            If mblnEOUUnit = True Then
                '10869291
                Call selectDataFromSaleConf(Trim(txtUnitCode.Text), cmbInvType, "Description", "'INV','SMP','TRF','REJ','JOB','SRC','ITD'", "datediff(dd,getdate(),fin_start_date)<=0  and datediff(dd,fin_end_date,getdate())<=0")
            Else
                Call selectDataFromSaleConf(Trim(txtUnitCode.Text), cmbInvType, "Description", "'INV','SMP','TRF','REJ','JOB','EXP','SRC','ITD'", "datediff(dd,getdate(),fin_start_date)<=0  and datediff(dd,fin_end_date,getdate())<=0")
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
    Private Sub CmbCategory_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmbCategory.SelectedIndexChanged
        On Error GoTo Err_Handler
        If Len(Trim(CmbCategory.Text)) = 0 Or Trim(CmbCategory.Text) = "-None-" Or Len(Trim(CmbCategory.Text)) > 0 Then
            lblcategory.Text = ""
            Ctlinvoice.Text = ""
        End If

        If Trim(CmbCategory.Text) = "REJECTION" Then
            If MsgBox("Do you want to make a Rejection Invoice for Customer", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Empower") = MsgBoxResult.Yes Then
                CUST_REJECTION_FLAG = True
            Else
                CUST_REJECTION_FLAG = False
            End If
        Else
            CUST_REJECTION_FLAG = False 'Default it should be False
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
        If Not (Len(CmbCategory.Text) <= 0) Then 'Checking if Item Field is not Blank
            If UCase(lbldescription.Text) = "SMP" And mblnpostinfin = True Then
                If rsSalesConf.State = ADODB.ObjectStateEnum.adStateOpen Then rsSalesConf.Close()
                rsSalesConf.Open("SELECT * FROM fin_GlobalGl WHERE gbl_prpsCode='Sample_Expences' and UNIT_CODE = '" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
                If rsSalesConf.EOF Then
                    MsgBox("Please define Sample Expences Account in Global Gl Definition", MsgBoxStyle.Information, "empower")
                End If
            End If
            If rsSalesConf.State = ADODB.ObjectStateEnum.adStateOpen Then rsSalesConf.Close()
            rsSalesConf.CursorLocation = ADODB.CursorLocationEnum.adUseClient
            rsSalesConf.Open("SELECT * FROM SaleConf (nolock) WHERE  UNIT_CODE = '" & gstrUNITID & "' and Invoice_Type='" & lbldescription.Text & "' AND Sub_Type_description ='" & CmbCategory.Text & "' AND Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,getdate(),fin_start_date)<=0  and datediff(dd,fin_end_date,getdate())<=0 ", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
            If Not rsSalesConf.EOF Then
                '13 dec 2017'
                mlbnprintforexcurrency = rsSalesConf.Fields("PRINTINVOICE_FOREXCURRENCY").Value
                mblnEwaybill_Print = rsSalesConf.Fields("EWAY_BILL_FUNCTIONALITY").Value

                If mblnEwaybill_Print = True Then
                    chkprintreprint.Enabled = True
                    chkprintreprint.Checked = True
                Else
                    chkprintreprint.Enabled = False
                    chkprintreprint.Checked = False
                End If
                If mlbnprintforexcurrency = False Then
                    lblprintinBasecurrency.Visible = False
                    ChkPrintForex.Visible = False
                    ChkPrintForex.Checked = False
                Else
                    lblprintinBasecurrency.Visible = True
                    ChkPrintForex.Visible = True
                    ChkPrintForex.Enabled = True
                End If
                '13 dec 2017'
                mstrPurposeCode = IIf(IsDBNull(rsSalesConf.Fields("inv_GLD_prpsCode").Value), "", Trim(rsSalesConf.Fields("inv_GLD_prpsCode").Value))
                mblnSameSeries = rsSalesConf.Fields("Single_Series").Value
                If ChkCustDetails.CheckState = System.Windows.Forms.CheckState.Checked Then
                    If IsDBNull(rsSalesConf.Fields("Report_filenameII").Value) = True Then
                        mstrReportFilename = ""
                    Else
                        mstrReportFilename = Trim(rsSalesConf.Fields("Report_filenameII").Value)
                    End If
                Else
                    mstrReportFilename = IIf(IsDBNull(rsSalesConf.Fields("Report_filename").Value), "", rsSalesConf.Fields("Report_filename").Value.ToString)
                End If
                If mstrPurposeCode = "" Then
                    MsgBox("Please select a Purpose Code in Sales Configuration", MsgBoxStyle.Information, "empower")
                    Me.CmbCategory.SelectedIndex = 0
                    Me.lblcategory.Text = ""
                    Me.cmbInvType.SelectedIndex = 3
                    Me.lbldescription.Text = ""
                    Me.cmbInvType.Focus()
                    mstrPurposeCode = ""
                    GoTo EventExitSub
                End If
            Else
                MsgBox("No record found in Sales Configuration for the selected Location, Invoice Type and Sub-Category", MsgBoxStyle.Information, "empower")
                Me.CmbCategory.SelectedIndex = 0
                Me.lblcategory.Text = ""
                Me.cmbInvType.SelectedIndex = 3
                Me.lbldescription.Text = ""
                Me.cmbInvType.Focus()
                mstrPurposeCode = ""
                GoTo EventExitSub
            End If
            mresult = New ClsResultSetDB_Invoice
            mresult.GetResult("Select sub_type,Sub_Type_Description,Stock_Location,updateStock_Flag  from SaleConf (nolock) where UNIT_CODE = '" & gstrUNITID & "' AND  Invoice_type = '" & Trim(Me.lbldescription.Text) & "' and sub_Type_Description = '" & Trim(Me.CmbCategory.Text) & "' and Location_Code ='" & Trim(txtUnitCode.Text) & "' and datediff(dd,GETDATE(),fin_start_date)<=0  and datediff(dd,fin_end_date,GETDATE())<=0")
            If (mresult.GetNoRows = 0) Then
                Call ConfirmWindow(10002, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
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
                        Ctlinvoice.Enabled = True
                        Ctlinvoice.BackColor = System.Drawing.Color.White
                        frachkRequired.Enabled = True
                        optYes(0).Enabled = True
                        optYes(1).Enabled = True
                        cmdHelp(2).Enabled = True
                        Me.Ctlinvoice.Focus()
                    Else
                        Call ConfirmWindow(10439, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
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
    Private Sub CmbInvType_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbInvType.SelectedIndexChanged
        On Error GoTo Err_Handler
        If Len(Trim(cmbInvType.Text)) = 0 Then
            lbldescription.Text = ""
        End If
        If Len(cmbInvType.Text) > 0 Or cmbInvType.Text = "-None-" Then 'Checking if Item Field is not Blank
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
            mresult = New ClsResultSetDB_Invoice
            If mblnEOUUnit = True Then
                '10869291
                mresult.GetResult("Select distinct(Invoice_type),Description from SaleConf (nolock) where  UNIT_CODE = '" & gstrUNITID & "' and Invoice_Type in('INV','SMP','REJ','TRF','JOB','SRC','ITD') and Description = '" & cmbInvType.Text & "' and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,GETDATE(),fin_start_date)<=0  and datediff(dd,fin_end_date,GETDATE())<=0")
            Else
                mresult.GetResult("Select distinct(Invoice_type),Description from SaleConf (nolock) where  UNIT_CODE = '" & gstrUNITID & "' and Invoice_Type in('INV','SMP','REJ','TRF','JOB','EXP','SRC','ITD') and Description = '" & cmbInvType.Text & "' and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,GETDATE(),fin_start_date)<=0  and datediff(dd,fin_end_date,GETDATE())<=0")
            End If
            If mresult.GetNoRows = 0 Then
                Call ConfirmWindow(10002, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_EXCLAMATION)
                Me.cmbInvType.SelectedIndex = 0
                Me.lbldescription.Text = ""
                mresult.ResultSetClose()
                Cancel = True
                GoTo EventExitSub
            Else
                lbldescription.Text = mresult.GetValue("Invoice_type")
                mresult.ResultSetClose()
                CmbCategory.Enabled = True
                CmbCategory.BackColor = System.Drawing.Color.White
                Call selectDataFromSaleConf(Trim(txtUnitCode.Text), CmbCategory, "Sub_Type_Description", "'" & Trim(lbldescription.Text) & "'", " datediff(dd,getdate(),fin_start_date)<=0  and datediff(dd,fin_end_date,getdate())<=0")
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
        Dim strHelp As Object
        Dim dblcustmtrl As Object
        Dim strQry As String
        On Error GoTo Err_Handler
        gobjDB = New ClsResultSetDB_Invoice
        Select Case Index
            Case 2
                With Me.Ctlinvoice
                    If optInvYes(0).Checked = True Then
                        'strHelp = ShowList(1, .Maxlength, "", "Doc_No", "Invoice_Type", "SalesChallan_dtl", " and Invoice_Type = '" & Me.lbldescription.Text & "' and Sub_category = '" & Me.lblcategory.Text & "' and Doc_No >99000000 and bill_flag = 0 and Location_Code='" & Trim(txtUnitCode.Text) & "'")
                        strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "SELECT convert(bigint,Doc_no) as doc_no,Invoice_Type,cust_name FROM Saleschallan_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  invoice_type='" & Me.lbldescription.Text & "' and sub_category='" & Me.lblcategory.Text & "' and bill_flag = 0 and Location_Code='" & Trim(txtUnitCode.Text) & "' ")
                    Else
                        'strHelp = ShowList(1, .Maxlength, "", "Doc_No", "Invoice_Type", "SalesChallan_dtl", " and Invoice_Type = '" & Me.lbldescription.Text & "' and Sub_category = '" & Me.lblcategory.Text & "' and Doc_No < 99000000 and bill_flag = 1 and cancel_flag = 0 and Location_Code='" & Trim(txtUnitCode.Text) & "'")
                        If mblnEwaybill_Print = True Then
                            If chkprintreprint.Checked = True Then
                                'strQry = "set dateformat 'dmy' SELECT distinct convert(varchar,Doc_no) as doc_no,Invoice_Type,cust_name  ,'' as EWAY_BILL_NO  FROM  Saleschallan_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND invoice_type='" & Trim(Me.lbldescription.Text) & "' and sub_category='" & Trim(Me.lblcategory.Text) & "' and bill_flag = '1' and cancel_flag = '0' and Location_Code='" & Trim(txtUnitCode.Text) & "'"
                                'strQry += " AND (( INVOICE_DATE < '" & mblnEWAY_BILL_STARTDATE & " '))"
                                'strQry += " UNION  SELECT distinct CONVERT(VARCHAR,DOC_NO) AS DOC_NO,INVOICE_TYPE,CUST_NAME,EWAY_BILL_NO  FROM  SALESCHALLAN_DTL WHERE UNIT_CODE='" + gstrUNITID + "' AND INVOICE_TYPE='" & Trim(Me.lbldescription.Text) & "' AND SUB_CATEGORY='" & Trim(Me.lblcategory.Text) & "' AND BILL_FLAG = '1' AND CANCEL_FLAG = '0' AND LOCATION_CODE='" & Trim(txtUnitCode.Text) & "'"
                                'strQry += " AND TOTAL_AMOUNT <= " & mdblewaymaximumvalue
                                'strQry += " UNION  SELECT distinct CONVERT(VARCHAR,DOC_NO) AS DOC_NO,INVOICE_TYPE,CUST_NAME,EWAY_BILL_NO  FROM  SALESCHALLAN_DTL WHERE UNIT_CODE='" + gstrUNITID + "' AND INVOICE_TYPE='" & Trim(Me.lbldescription.Text) & "' AND SUB_CATEGORY='" & Trim(Me.lblcategory.Text) & "' AND BILL_FLAG = '1' AND CANCEL_FLAG = '0' AND LOCATION_CODE='" & Trim(txtUnitCode.Text) & "'"
                                'strQry += " AND TOTAL_AMOUNT > " & mdblewaymaximumvalue & " AND  INVOICE_DATE >= '" & mblnEWAY_BILL_STARTDATE & " ' "
                                'strQry += " AND ISNULL(EWAY_BILL_NO,'')<>'' AND  INVOICE_DATE >= '" & mblnEWAY_BILL_STARTDATE & " '"
                                'strQry += " AND  EXISTS ( SELECT TOP 1 1 FROM FIRSTTIME_INVOICEPRINTING (NOLOCK) WHERE UNIT_CODE =SALESCHALLAN_DTL.UNIT_CODE AND FIRSTTIME_INVOICEPRINTING.DOC_NO=SALESCHALLAN_DTL.DOC_NO )"

                                '102027599
                                strQry = "SELECT DOC_NO,INVOICE_TYPE,CUST_NAME,EWAY_BILL_NO FROM DBO.UDF_GET_LOCKED_INVOICES_FOR_INVOICE_PRINTING('" & gstrUnitId & "','" & Trim(txtUnitCode.Text) & "',1,'" & Me.lbldescription.Text & "','" & Me.lblcategory.Text & "') ORDER BY DOC_NO "
                            Else
                                'strQry = "set dateformat 'dmy' "
                                'strQry += " SELECT CONVERT(VARCHAR,DOC_NO) AS DOC_NO,INVOICE_TYPE,CUST_NAME ,EWAY_BILL_NO FROM  SALESCHALLAN_DTL WHERE UNIT_CODE='" + gstrUNITID + "' AND INVOICE_TYPE='" & Trim(Me.lbldescription.Text) & "' AND SUB_CATEGORY='" & Trim(Me.lblcategory.Text) & "' AND BILL_FLAG = '1' AND CANCEL_FLAG = '0' AND LOCATION_CODE='" & Trim(txtUnitCode.Text) & "'"
                                'strQry += " AND TOTAL_AMOUNT > " & mdblewaymaximumvalue & " AND  INVOICE_DATE >= '" & mblnEWAY_BILL_STARTDATE & " ' "
                                'strQry += " AND ISNULL(EWAY_BILL_NO,'')<>'' AND  INVOICE_DATE >= '" & mblnEWAY_BILL_STARTDATE & " ' "
                                'strQry += " AND  NOT EXISTS ( SELECT TOP 1 1 FROM FIRSTTIME_INVOICEPRINTING (NOLOCK) WHERE UNIT_CODE =SALESCHALLAN_DTL.UNIT_CODE AND FIRSTTIME_INVOICEPRINTING.DOC_NO=SALESCHALLAN_DTL.DOC_NO )"
                                'strQry += " ORDER BY DOC_NO "

                                '102027599
                                strQry = "SELECT DOC_NO,INVOICE_TYPE,CUST_NAME,EWAY_BILL_NO FROM DBO.UDF_GET_LOCKED_INVOICES_FOR_INVOICE_PRINTING('" & gstrUnitId & "','" & Trim(txtUnitCode.Text) & "',0,'" & Me.lbldescription.Text & "','" & Me.lblcategory.Text & "') ORDER BY DOC_NO "
                            End If
                        Else
                            strQry = "SELECT convert(varchar,Doc_no) as doc_no,Invoice_Type,cust_name  FROM  Saleschallan_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND invoice_type='" & Trim(Me.lbldescription.Text) & "' and sub_category='" & Trim(Me.lblcategory.Text) & "' and bill_flag = '1' and cancel_flag = '0' and Location_Code='" & Trim(txtUnitCode.Text) & "'"
                        End If
                        strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQry)
                    End If

                End With
                If UBound(strHelp) = "-1" Then ' No record
                    Call ConfirmWindow(10512, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION)
                Else
                    'Me.Ctlinvoice.Text = strHelp
                    Ctlinvoice.Text = Trim(strHelp(0))
                    Ctlinvoice.Focus()
                    If optInvYes(0).Checked = True Then
                        gobjDB.GetResult("SELECT Doc_NO,Invoice_Type FROM SalesChallan_Dtl Where  UNIT_CODE = '" & gstrUNITID & "' and Invoice_Type = '" & Me.lbldescription.Text & "' and Sub_category = '" & Me.lblcategory.Text & "' and bill_flag =0 and Doc_No = '" & Ctlinvoice.Text & "' and Location_Code='" & Trim(txtUnitCode.Text) & "'")
                    Else
                        gobjDB.GetResult("SELECT Doc_NO,Invoice_Type FROM SalesChallan_Dtl Where  UNIT_CODE = '" & gstrUNITID & "' and Invoice_Type = '" & Me.lbldescription.Text & "' and Sub_category = '" & Me.lblcategory.Text & "' and bill_flag =1 and Doc_No = '" & Ctlinvoice.Text & "'  and Location_Code='" & Trim(txtUnitCode.Text) & "'")
                    End If
                    If gobjDB.GetNoRows > 0 Then 'RECORD FOUND
                    End If
                    gobjDB.GetResult("SELECT cust_mtrl=sum(cust_mtrl)FROM Sales_Dtl Where Doc_No >99000000 and  Doc_No = '" & Ctlinvoice.Text & "' and Location_Code='" & Trim(txtUnitCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "'")

                    dblcustmtrl = Val(gobjDB.GetValue("cust_mtrl"))
                    If dblcustmtrl > 0 Then
                        ChkCustDetails.CheckState = System.Windows.Forms.CheckState.Checked
                        Call CmbCategory_Validating(CmbCategory, New System.ComponentModel.CancelEventArgs(False))
                    Else
                        ChkCustDetails.CheckState = System.Windows.Forms.CheckState.Unchecked
                    End If
                End If
        End Select
        gobjDB.ResultSetClose()
        gobjDB = Nothing
        'SATISH KESHAERWANI CHANGE
        If Len(Ctlinvoice.Text.Trim) > 0 Then
            If DataExist("Select TOP 1 1 FROM SALESCHALLAN_DTL SC WHERE UNIT_CODE='" & gstrUNITID & "' AND DOC_NO=" & Ctlinvoice.Text.Trim & " AND EXISTS(SELECT  TOP 1 1  from customer_mst C WHERE C.UNIT_CODE=SC.UNIT_CODE AND C.CUSTOMER_CODE=SC.ACCOUNT_CODE AND C.UNIT_CODE='" + gstrUNITID + "' and Allow_BARCODELABELFORMAT =1 )") = True Then
                llbtkmlPrintingFlag.Enabled = True
                chktkmlbarcode.Enabled = True
                chktkmlbarcode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                chktkmlbarcode.Checked = False
                If optInvYes(0).Checked = False Then
                    ChkQrbarcodereprint.Enabled = True
                    ChkQrbarcodereprint.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    ChkQrbarcodereprint.Checked = False
                End If
            ElseIf DataExist("Select TOP 1 1 FROM SALESCHALLAN_DTL SC WHERE UNIT_CODE='" & gstrUNITID & "' AND DOC_NO=" & Ctlinvoice.Text.Trim & " AND EXISTS(SELECT  TOP 1 1  from customer_mst C WHERE C.UNIT_CODE=SC.UNIT_CODE AND C.CUSTOMER_CODE=SC.ACCOUNT_CODE AND C.UNIT_CODE='" + gstrUNITID + "' AND PRINT_METHOD='TATA')") = True Then
                If optInvYes(0).Checked = False Then
                    ChkQrbarcodereprint.Enabled = True
                    ChkQrbarcodereprint.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    ChkQrbarcodereprint.Checked = False
                End If
            Else
                llbtkmlPrintingFlag.Enabled = False
                chktkmlbarcode.Enabled = False
                chktkmlbarcode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                chktkmlbarcode.Checked = False
                'If optInvYes(0).Checked = True Then
                ChkQrbarcodereprint.Enabled = False
                ChkQrbarcodereprint.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                ChkQrbarcodereprint.Checked = False
                'End If
            End If
        End If
        'SATISH KESHAERWANI CHANGE

        Exit Sub
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
                txtPLA.Enabled = True : txtPLA.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                dtpRemoval.Enabled = True
                txtPLA.Focus()
                Exit Sub
            End If
        End If
        'SATISH KESHAERWANI CHANGE
        If DataExist("Select TOP 1 1 FROM SALESCHALLAN_DTL SC WHERE UNIT_CODE='" & gstrUNITID & "' AND DOC_NO=" & mInvNo & " AND EXISTS(SELECT  TOP 1 1  from customer_mst C WHERE C.UNIT_CODE=SC.UNIT_CODE AND C.CUSTOMER_CODE=SC.ACCOUNT_CODE AND C.UNIT_CODE='" + gstrUNITID + "' and Allow_BARCODELABELFORMAT =1 )") = True Then
            llblLockPrintingFlag.Enabled = True
            chktkmlbarcode.Enabled = True
            If optInvYes(0).Checked = False Then
                ChkQrbarcodereprint.Enabled = True
            End If
        Else
            llblLockPrintingFlag.Enabled = False
            chktkmlbarcode.Enabled = False
            If optInvYes(0).Checked = True Then
                ChkQrbarcodereprint.Enabled = False
            End If
        End If
        'SATISH KESHAERWANI CHANGE
        If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
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
        Dim strAccountCode As String
        Dim strInvoiceDate As String

        On Error GoTo Err_Handler
        If Len(Ctlinvoice.Text) = 0 Then GoTo EventExitSub
        If mblnEOUUnit = True Then
            mStrCustMst = "Select Doc_No,Invoice_type from SalesChallan_Dtl where  UNIT_CODE = '" & gstrUNITID & "' and Invoice_Type <> 'EXP' and "
        Else
            mStrCustMst = "Select Doc_No,Invoice_type from SalesChallan_Dtl where UNIT_CODE = '" & gstrUNITID & "' and "
        End If
        'If mblnEwaybill_Print = True Then
        '    If optInvYes(0).Checked = False Then
        '        'mStrCustMst = mStrCustMst & " (( INVOICE_DATE < '" & mblnEWAY_BILL_STARTDATE & " ') or ( ISNULL(EWAY_BILL_NO,'')<>'' AND  INVOICE_DATE >= '" & mblnEWAY_BILL_STARTDATE & "')) and "
        '        mStrCustMst = mStrCustMst & " ( ( INVOICE_DATE < '" & mblnEWAY_BILL_STARTDATE & " ') or (INVOICE_DATE >= '" & mblnEWAY_BILL_STARTDATE & "' and total_amount <= " & mdblewaymaximumvalue & ")  or ( ISNULL(EWAY_BILL_NO,'')<>'' AND  INVOICE_DATE >= '" & mblnEWAY_BILL_STARTDATE & "')) and "
        '        'If chkprintreprint.Checked = True Then
        '        'mStrCustMst = mStrCustMst & " EXISTS ( SELECT TOP 1 1 FROM FIRSTTIME_INVOICEPRINTING (NOLOCK) WHERE UNIT_CODE =SALESCHALLAN_DTL.UNIT_CODE AND FIRSTTIME_INVOICEPRINTING.DOC_NO=SALESCHALLAN_DTL.DOC_NO ) AND  "
        '        'Else
        '        'mStrCustMst = mStrCustMst & "NOT EXISTS ( SELECT TOP 1 1 FROM FIRSTTIME_INVOICEPRINTING (NOLOCK) WHERE UNIT_CODE =SALESCHALLAN_DTL.UNIT_CODE AND FIRSTTIME_INVOICEPRINTING.DOC_NO=SALESCHALLAN_DTL.DOC_NO ) AND "
        '        'End If
        '    End If
        'End If

        If optInvYes(0).Checked = True Then
            mStrCustMst = mStrCustMst & " Location_Code='" & Trim(txtUnitCode.Text) & "'" & " and bill_flag =0 and Doc_No ="
        Else
            mStrCustMst = mStrCustMst & " Location_Code='" & Trim(txtUnitCode.Text) & "'" & " and bill_flag =1 and CANCEL_FLAG = 0 and Doc_No ="
        End If

        '102027599
        If mblnEwaybill_Print = True AndAlso optInvYes(0).Checked = False Then
            If mblnEOUUnit = True Then
                If chkprintreprint.Checked = True Then
                    mStrCustMst = "Select S.Doc_No,S.Invoice_type from SalesChallan_Dtl S where S.UNIT_CODE = '" & gstrUNITID & "' and S.Invoice_Type <> 'EXP' "
                    mStrCustMst += " AND S.EWAY_IRN_REQUIRED='N' "
                    mStrCustMst += " AND S.Location_Code='" & Trim(txtUnitCode.Text) & "'" & " and S.bill_flag =1 and S.CANCEL_FLAG = 0 and S.Doc_No = " & Ctlinvoice.Text & " "
                    mStrCustMst += " UNION Select S.Doc_No,S.Invoice_type from SalesChallan_Dtl S LEFT JOIN SALESCHALLAN_DTL_IRN I ON I.UNIT_CODE=S.UNIT_CODE AND I.DOC_NO=S.DOC_NO where  S.UNIT_CODE = '" & gstrUNITID & "' and S.Invoice_Type <> 'EXP' "
                    mStrCustMst += " AND ((S.EWAY_IRN_REQUIRED='E' AND ISNULL(S.EWAY_BILL_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='I' AND ISNULL(I.IRN_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='B' AND ISNULL(S.EWAY_BILL_NO,'')<>'' AND ISNULL(I.IRN_NO,'')<>'')) "
                    mStrCustMst += " AND S.Location_Code='" & Trim(txtUnitCode.Text) & "'" & " and S.bill_flag =1 and S.CANCEL_FLAG = 0 "
                    mStrCustMst += " AND EXISTS (SELECT TOP 1 1 FROM FIRSTTIME_INVOICEPRINTING F (NOLOCK) WHERE F.UNIT_CODE =S.UNIT_CODE AND F.DOC_NO=S.DOC_NO) and S.Doc_No = "
                Else
                    mStrCustMst = " Select S.Doc_No,S.Invoice_type from SalesChallan_Dtl S LEFT JOIN SALESCHALLAN_DTL_IRN I ON I.UNIT_CODE=S.UNIT_CODE AND I.DOC_NO=S.DOC_NO where  S.UNIT_CODE = '" & gstrUNITID & "' and S.Invoice_Type <> 'EXP' "
                    mStrCustMst += " AND ((S.EWAY_IRN_REQUIRED='E' AND ISNULL(S.EWAY_BILL_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='I' AND ISNULL(I.IRN_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='B' AND ISNULL(S.EWAY_BILL_NO,'')<>'' AND ISNULL(I.IRN_NO,'')<>'')) "
                    mStrCustMst += " AND S.Location_Code='" & Trim(txtUnitCode.Text) & "'" & " and S.bill_flag =1 and S.CANCEL_FLAG = 0 "
                    mStrCustMst += " AND NOT EXISTS (SELECT TOP 1 1 FROM FIRSTTIME_INVOICEPRINTING F (NOLOCK) WHERE F.UNIT_CODE =S.UNIT_CODE AND F.DOC_NO=S.DOC_NO) and S.Doc_No = "
                End If
            Else
                If chkprintreprint.Checked = True Then
                    mStrCustMst = "Select S.Doc_No,S.Invoice_type from SalesChallan_Dtl S where S.UNIT_CODE = '" & gstrUNITID & "' "
                    mStrCustMst += " AND S.EWAY_IRN_REQUIRED='N' "
                    mStrCustMst += " AND S.Location_Code='" & Trim(txtUnitCode.Text) & "'" & " and S.bill_flag =1 and S.CANCEL_FLAG = 0 and S.Doc_No = " & Ctlinvoice.Text & " "
                    mStrCustMst += " UNION Select S.Doc_No,S.Invoice_type from SalesChallan_Dtl S LEFT JOIN SALESCHALLAN_DTL_IRN I ON I.UNIT_CODE=S.UNIT_CODE AND I.DOC_NO=S.DOC_NO where  S.UNIT_CODE = '" & gstrUNITID & "' "
                    mStrCustMst += " AND ((S.EWAY_IRN_REQUIRED='E' AND ISNULL(S.EWAY_BILL_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='I' AND ISNULL(I.IRN_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='B' AND ISNULL(S.EWAY_BILL_NO,'')<>'' AND ISNULL(I.IRN_NO,'')<>'')) "
                    mStrCustMst += " AND S.Location_Code='" & Trim(txtUnitCode.Text) & "'" & " and S.bill_flag =1 and S.CANCEL_FLAG = 0 "
                    mStrCustMst += " AND EXISTS (SELECT TOP 1 1 FROM FIRSTTIME_INVOICEPRINTING F (NOLOCK) WHERE F.UNIT_CODE =S.UNIT_CODE AND F.DOC_NO=S.DOC_NO) and S.Doc_No = "
                Else
                    mStrCustMst = " Select S.Doc_No,S.Invoice_type from SalesChallan_Dtl S LEFT JOIN SALESCHALLAN_DTL_IRN I ON I.UNIT_CODE=S.UNIT_CODE AND I.DOC_NO=S.DOC_NO where  S.UNIT_CODE = '" & gstrUNITID & "' "
                    mStrCustMst += " AND ((S.EWAY_IRN_REQUIRED='E' AND ISNULL(S.EWAY_BILL_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='I' AND ISNULL(I.IRN_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='B' AND ISNULL(S.EWAY_BILL_NO,'')<>'' AND ISNULL(I.IRN_NO,'')<>'')) "
                    mStrCustMst += " AND S.Location_Code='" & Trim(txtUnitCode.Text) & "'" & " and S.bill_flag =1 and S.CANCEL_FLAG = 0 "
                    mStrCustMst += " AND NOT EXISTS (SELECT TOP 1 1 FROM FIRSTTIME_INVOICEPRINTING F (NOLOCK) WHERE F.UNIT_CODE =S.UNIT_CODE AND F.DOC_NO=S.DOC_NO) and S.Doc_No = "
                End If
            End If
        End If

        strSql = mStrCustMst & Ctlinvoice.Text

        Me.Ctlinvoice.ExistRecQry = mStrCustMst
        If Len(Ctlinvoice.Text) > 0 Then 'Checking if Item Field is not Blank
            If Ctlinvoice.ExistsRec = True Then 'Checking if the Record Exists
                Me.Cmdinvoice.Focus()

                'Added by priti to update new digital sign print
                If mblnISTrueSignRequired Then
                    chkNewDigitalSign.Enabled = True
                    chkNewDigitalSign.Checked = False
                    chkNewDigitalSign.Visible = True
                Else
                    chkNewDigitalSign.Checked = False
                    chkNewDigitalSign.Visible = False
                End If
                'end by priti to update new digital sign print

                strAccountCode = Find_Value("select account_code from saleschallan_dtl where UNIT_CODE = '" & gstrUNITID & "' and doc_no='" & Trim(Me.Ctlinvoice.Text) & "'")
                If Me.lbldescription.Text = "REJ" Then
                    blnlinelevelcustomer = False
                Else
                    blnlinelevelcustomer = Find_Value("SELECT SOUPLD_LINE_LEVEL_SALESORDER FROM CUSTOMER_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & strAccountCode & "'")

                End If
                '26 sep 2016
                mblnlorryno = False
                If (UCase(cmbInvType.Text) = "NORMAL INVOICE" And UCase(CmbCategory.Text) = "FINISHED GOODS") And optInvYes(0).Checked = True Then
                    strSql = "select dbo.UDF_ISDOCKCODE_INVOICE( '" & gstrUNITID & "','" & strAccountCode & "','NORMAL INVOICE','FINISHED GOODS' )"
                    If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSql)) = True Then
                        txtlorryno.Visible = True
                        txtlorryno.Enabled = True
                        txtlorryno.Text = Find_Value("select Lorryno_date from saleschallan_dtl where UNIT_CODE = '" & gstrUNITID & "' and doc_no='" & Trim(Me.Ctlinvoice.Text) & "'")
                        txtlorryno.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        lblLorry.Visible = True
                        mblnlorryno = True
                    End If
                Else
                    txtlorryno.Visible = False
                    txtlorryno.Enabled = False
                    txtlorryno.Text = ""
                    txtlorryno.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    lblLorry.Visible = False
                    mblnlorryno = False
                End If


                '26 sep 2016
                If optInvYes(1).Checked = True And AllowASNPrinting(strAccountCode) = True Then
                    Me.txtASNNumber.Visible = True
                    Me.txtASNNumber.Enabled = True
                    Me.txtASNNumber.Text = ""
                    Me.txtASNNumber.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    Me.lblASNNumber.Visible = True
                    Me.txtASNNumber.Text = CheckASNExist(Me.Ctlinvoice.Text)        'Get Saved ASN Number
                    Me.txtASNNumber.Focus()
                Else
                    Me.txtASNNumber.Visible = False
                    Me.txtASNNumber.Enabled = False
                    Me.txtASNNumber.Text = ""
                    Me.txtASNNumber.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    Me.lblASNNumber.Visible = False
                End If
            Else
                Cancel = True
                Ctlinvoice.Text = ""
                If optInvYes(1).Checked = True Then
                    Me.txtASNNumber.Visible = False
                    Me.txtASNNumber.Text = ""
                    Me.txtASNNumber.Enabled = False
                    Me.txtASNNumber.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    Me.lblASNNumber.Visible = False
                    Me.txtlorryno.Visible = False
                    Me.txtlorryno.Text = ""
                    Me.txtlorryno.Enabled = False
                    Me.txtlorryno.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    Me.lblLorry.Visible = False
                End If
                Ctlinvoice.Focus()
                GoTo EventExitSub
            End If
        End If
        'SATISH KESHAERWANI CHANGE
        If UCase(cmbInvType.Text) = "NORMAL INVOICE" And UCase(CmbCategory.Text) = "FINISHED GOODS" Then
            If DataExist("Select TOP 1 1 FROM SALESCHALLAN_DTL SC WHERE UNIT_CODE='" & gstrUNITID & "' AND DOC_NO=" & Ctlinvoice.Text.Trim & " AND EXISTS(SELECT  TOP 1 1  from customer_mst C WHERE C.UNIT_CODE=SC.UNIT_CODE AND C.CUSTOMER_CODE=SC.ACCOUNT_CODE AND C.UNIT_CODE='" + gstrUNITID + "' and Allow_BARCODELABELFORMAT =1 )") = True Then
                llbtkmlPrintingFlag.Enabled = True
                chktkmlbarcode.Enabled = True
                chktkmlbarcode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                chktkmlbarcode.Checked = False
                If optInvYes(0).Checked = False Then
                    ChkQrbarcodereprint.Enabled = True
                    ChkQrbarcodereprint.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    ChkQrbarcodereprint.Checked = False
                End If
            Else
                llbtkmlPrintingFlag.Enabled = False
                chktkmlbarcode.Enabled = False
                chktkmlbarcode.Checked = False
                chktkmlbarcode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                If optInvYes(0).Checked = True Then
                    ChkQrbarcodereprint.Enabled = False
                    ChkQrbarcodereprint.Checked = False
                    ChkQrbarcodereprint.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                End If
            End If
        Else
            llbtkmlPrintingFlag.Enabled = False
            chktkmlbarcode.Enabled = False
            chktkmlbarcode.Checked = False
            chktkmlbarcode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        End If
        'SATISH KESHAERWANI CHANGE
        '06 nov 2017
        If optInvYes(0).Checked = False Then
            strInvoiceDate = Find_Value("select invoice_date from saleschallan_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  doc_no='" & Trim(Me.Ctlinvoice.Text) & "'")
            mstrReportFilename = Find_Value("SELECT Report_filename FROM SaleConf (Nolock) WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_Type='" & lbldescription.Text & "' AND Sub_Type_description ='" & CmbCategory.Text & "' AND Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & strInvoiceDate & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & strInvoiceDate & "')<=0 ")
        End If

        '06 nov 2017

        GoTo EventExitSub
Err_Handler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub frmMKTTRN0008_SOUTH_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo Err_Handler
        mdifrmMain.CheckFormName = mintFormIndex
        frmModules.NodeFontBold(Tag) = True
        Exit Sub
Err_Handler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0008_SOUTH_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        On Error GoTo Err_Handler
        frmModules.NodeFontBold(Tag) = False
        Exit Sub
Err_Handler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub Form_Initialize_Renamed()
        On Error GoTo Err_Handler
        gobjDB = New ClsResultSetDB_Invoice
        gobjDB.GetResult("SELECT isnull(EWAY_INV_MAXRANGE,0)EWAY_INV_MAXRANGE ,EWAY_BILL_STARTDATE, EOU_Flag, CustSupp_Inc,InsExc_Excise,postinfin,Excise_RoundOFF,Bomcheck_flag,TOYOTA_MULTIPLESO_ONEPDS_STDATE FROM sales_parameter (NOLOCK) where UNIT_CODE = '" & gstrUNITID & "' ")
        If gobjDB.GetValue("EOU_Flag") = True Then
            mStrCustMst = "Select Doc_No,Invoice_type from SalesChallan_Dtl where Invoice_Type <> 'EXP' and Location_Code='" & Trim(txtUnitCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "' "
            mblnEOUUnit = True
        Else
            mStrCustMst = "Select Doc_No,Invoice_type from SalesChallan_Dtl where Location_Code='" & Trim(txtUnitCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "'"
            mblnEOUUnit = False
        End If
        mblnAddCustomerMaterial = gobjDB.GetValue("CustSupp_Inc")
        mblnInsuranceFlag = gobjDB.GetValue("InsExc_Excise")
        mblnpostinfin = gobjDB.GetValue("postinfin")
        mblnExciseRoundOFFFlag = gobjDB.GetValue("Excise_RoundOFF")
        mblnBOMcheck_flag = gobjDB.GetValue("BOMcheck_flag")
        mblnEWAY_BILL_STARTDATE = gobjDB.GetValue("EWAY_BILL_STARTDATE")
        mdblewaymaximumvalue = gobjDB.GetValue("EWAY_INV_MAXRANGE")
        mblnTOYOTA_MULTIPLESO_ONEPDS_STDATE = gobjDB.GetValue("TOYOTA_MULTIPLESO_ONEPDS_STDATE")
        gobjDB.ResultSetClose()
        gobjDB = Nothing
        Exit Sub
Err_Handler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0008_SOUTH_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0 Then Call ctlFormHeader1_Click(ctlFormHeader1, New System.EventArgs()) : Exit Sub
    End Sub
    Private Sub frmMKTTRN0008_SOUTH_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo Err_Handler
        Dim objFinconfRecordset As New ADODB.Recordset
        mintFormIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader1.Tag)
        Call FillLabelFromResFile(Me) 'To Fill label description from Resource file
        Call FitToClient(Me, fraInvoice, ctlFormHeader1, Cmdinvoice) 'To fit the form in the MDI
        Call EnableControls(False, Me) 'To Disable controls
        optInvYes(0).Enabled = True : optInvYes(1).Enabled = True
        gblnCancelUnload = False
        txtUnitCode.Enabled = True
        txtUnitCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        cmdUnitCodeList.Enabled = True
        lbldescription.Visible = False
        lblcategory.Visible = False
        cmdHelp(2).Image = My.Resources.ico111.ToBitmap

        'optYes(1).Checked = True 'commented by abhijit
        optYes(0).Checked = True 'added by abhijit

        optInvYes(0).Checked = True

        dtpRemoval.Format = DateTimePickerFormat.Custom
        dtpRemoval.CustomFormat = gstrDateFormat
        dtpRemoval.Value = GetServerDate()

        Me.ChkCustDetails.Enabled = True
        Me.ChkCustDetails.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        Me.ShpCustDetails.Visible = False
        Me.chkprintreprint.Enabled = True
        Me.chkprintreprint.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)

        mPdfReaderPath = (SqlConnectionclass.ExecuteScalar("SELECT PDFREADERPATH FROM GLOBAL_FLAG (NOLOCK)")).ToString()

        mstrFGDomestic = Find_Value("Select FG_DOMESTIC from BarCode_config_mst where UNIT_CODE = '" & gstrUNITID & "'")
        mblnCSM_Knockingoff_req = DataExist("SELECT TOP 1 1 FROM SALES_PARAMETER WHERE CSM_KNOCKINGOFF_REQ = 1 and UNIT_CODE = '" & gstrUNITID & "'")
        Call Form_Initialize_Renamed()
        'chkPrintReprint.Enabled = True
        mblnCCFlag = False
        'ISSUE ID 10151004  
        If objFinconfRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objFinconfRecordset.Close() : objFinconfRecordset = Nothing
        objFinconfRecordset.Open("SELECT * FROM FIN_CONF WHERE Unit='" & gstrUNITID & "' and FUNCTIONALITY='CC_GLOBALFLAG' AND ACTIVE=1", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
        If Not objFinconfRecordset.EOF Then
            mblnCCFlag = True
        Else
            mblnCCFlag = False
        End If
        If objFinconfRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objFinconfRecordset.Close() : objFinconfRecordset = Nothing
        ' ISSUE ID 10151004  END 
        mblnZEROASNCUMMS_REJECTIONINVOICE = CBool(Find_Value("SELECT ZEROASNCUMMS_REJECTION FROM SALES_PARAMETER WHERE UNIT_CODE='" + gstrUNITID + "'"))
        mintmaxnoofitems_Annexture = CInt(Find_Value("SELECT MAXIMUMITEMS_FORANNEXTURE FROM SALES_PARAMETER WHERE UNIT_CODE='" + gstrUNITID + "'"))
        mintmaxnoofitems_barcodeToyota = CInt(Find_Value("SELECT MAXIMUMITEMS_FORBARCODETOYOTA FROM SALES_PARAMETER WHERE UNIT_CODE='" + gstrUNITID + "'"))
        '10688760 
        mbln_SHIPPING_ADDRESS = CBool(Find_Value("SELECT REQD_SHIPPING_ADDRESS FROM SALES_PARAMETER WHERE UNIT_CODE='" + gstrUNITID + "'"))
        '10688760 
        '10869290
        mblnServiceInvoicemate = CBool(Find_Value("SELECT ServiceInvoice_MATE FROM SALES_PARAMETER WHERE UNIT_CODE='" + gstrUNITID + "'"))
        '10869290
        Me.txtlorryno.Text = ""
        chkprintreprint.Enabled = False
        chkprintreprint.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        '102027599
        If CBool(Find_Value("SELECT ISNULL(MAX(CAST(EWAY_BILL_FUNCTIONALITY AS INT)),0) EWAY_BILL_FUNCTIONALITY FROM SALECONF (NOLOCK) WHERE  UNIT_CODE = '" & gstrUNITID & "' AND DATEDIFF(DD,GETDATE(),FIN_START_DATE)<=0  AND DATEDIFF(DD,FIN_END_DATE,GETDATE())<=0 ")) Then
            btnExceptionInvoices.Enabled = True
        Else
            btnExceptionInvoices.Enabled = False
        End If
        '' PRAVEEN DIGITAL SIGN
        mblnISTrueSignRequired = CBool(Find_Value("SELECT ISNULL(IS_TRUE_SIGN_REQUIRED,0) FROM gen_unitmaster (NOLOCK) WHERE Unt_CodeID='" + gstrUNITID + "'"))
        mblnAPIUrl = Find_Value("Select API_Url from gen_unitmaster (NOLOCK) where Unt_CodeID = '" & gstrUNITID & "'")
        mblnPFX_ID = Find_Value("Select PFX_ID from gen_unitmaster (NOLOCK) where Unt_CodeID = '" & gstrUNITID & "'")
        mblnPFX_Pass = Find_Value("Select PFX_password from gen_unitmaster (NOLOCK) where Unt_CodeID = '" & gstrUNITID & "'")
        mblnAPI_Key = Find_Value("Select API_Key from gen_unitmaster (NOLOCK) where Unt_CodeID = '" & gstrUNITID & "'")
        Exit Sub
Err_Handler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub selectDataFromSaleConf(ByRef LocationCode As String, ByRef combo As System.Windows.Forms.ComboBox, ByRef feild As String, ByRef invoicetype As String, ByRef pstrCondition As String)
        Dim strSql As String
        Dim rsSaleConf As ClsResultSetDB_Invoice
        Dim intRowCount As Short
        Dim intloopcount As Short
        On Error GoTo Err_Handler
        strSql = "select Distinct(" & feild & ") from Saleconf (nolock) where  UNIT_CODE = '" & gstrUNITID & "' and Location_Code='" & LocationCode & "' and Invoice_Type in(" & invoicetype & ") and " & pstrCondition
        rsSaleConf = New ClsResultSetDB_Invoice
        rsSaleConf.GetResult(strSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsSaleConf.GetNoRows > 0 Then
            combo.Items.Clear()
            intRowCount = rsSaleConf.GetNoRows
            VB6.SetItemString(combo, 0, "-None-")
            rsSaleConf.MoveFirst()
            For intloopcount = 1 To intRowCount
                VB6.SetItemString(combo, intloopcount, rsSaleConf.GetValue(feild))
                rsSaleConf.MoveNext()
            Next intloopcount
            rsSaleConf.ResultSetClose()
            rsSaleConf = Nothing
        End If
        Exit Sub
Err_Handler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0008_SOUTH_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error GoTo Err_Handler
        mdifrmMain.RemoveFormNameFromWindowList = mintFormIndex
        frmModules.NodeFontBold(Tag) = False
        Me.Dispose()
        Exit Sub
Err_Handler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Function ValuetoVariables() As Boolean
        Dim strSql As String
        Dim rsSalesChallan As ClsResultSetDB_Invoice
        Dim strInvoiceDate As String
        Dim strCustomerCode As String
        On Error GoTo Err_Handler
        strSql = "select INVOICE_DATE,Account_Code from Saleschallan_Dtl where Doc_No =" & Me.Ctlinvoice.Text & "  and Location_Code='" & Trim(txtUnitCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "'"
        rsSalesChallan = New ClsResultSetDB_Invoice
        rsSalesChallan.GetResult(strSql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        strInvoiceDate = VB6.Format(rsSalesChallan.GetValue("Invoice_Date"), gstrDateFormat)
        strCustomerCode = rsSalesChallan.GetValue("Account_Code")
        rsSalesChallan.ResultSetClose()
        mInvType = Me.lbldescription.Text
        mSubCat = Me.lblcategory.Text
        mInvNo = CDbl(Val(GenerateInvoiceNo(mInvType, mSubCat, strInvoiceDate, strCustomerCode)))
        strSql = " Select Asseccable= ISNULL(SUM(Accessible_amount),0) from sales_dtl "
        strSql = strSql & " where Doc_No =" & Me.Ctlinvoice.Text & " and Location_Code='" & Trim(txtUnitCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "'"
        mresult = New ClsResultSetDB_Invoice
        mresult.GetResult(strSql)
        mAssessableValue = mresult.GetValue("Asseccable")
        mresult.ResultSetClose()
        ValuetoVariables = True
        Exit Function
Err_Handler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
        ValuetoVariables = False
    End Function
    Public Function updatesalesconfandsaleschallan() As Boolean
        Dim strSql As String
        Dim rsSalesChallan As ClsResultSetDB_Invoice
        Dim dblInvoiceAmt As Double
        Dim strInvoiceDate As String
        On Error GoTo Err_Handler
        strSql = "select * from Saleschallan_dtl where Doc_No = " & Me.Ctlinvoice.Text
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
                        salesconf = "update saleconf set CURRENT_NO_TRF_SAMEGSTIN = " & mSaleConfNo & ", OpenningBal = openningBal - " & mAssessableValue & " where  UNIT_CODE = '" & gstrUNITID & "' and Invoice_type <> 'EXP' and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
                    Else
                        salesconf = "update saleconf set current_No = " & mSaleConfNo & ", OpenningBal = openningBal - " & mAssessableValue & " where  UNIT_CODE = '" & gstrUNITID & "' and Invoice_type <> 'EXP' and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
                    End If

                Else
                    If IsGSTINSAME(mAccount_Code) And (lbldescription.Text.ToUpper = "TRF" Or lbldescription.Text.ToUpper = "ITD") Then
                        salesconf = "update saleconf set CURRENT_NO_TRF_SAMEGSTIN = " & mSaleConfNo & " where  UNIT_CODE = '" & gstrUNITID & "' and Single_Series = 1 and Invoice_type <> 'EXP' and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0" & vbCrLf
                        salesconf = salesconf & "update saleconf set OpenningBal = openningBal - " & mAssessableValue & " where UNIT_CODE = '" & gstrUNITID & "' and  Invoice_type <> 'EXP' and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
                    Else
                        salesconf = "update saleconf set current_No = " & mSaleConfNo & " where  UNIT_CODE = '" & gstrUNITID & "' and Single_Series = 1 and Invoice_type <> 'EXP' and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0" & vbCrLf
                        salesconf = salesconf & "update saleconf set OpenningBal = openningBal - " & mAssessableValue & " where UNIT_CODE = '" & gstrUNITID & "' and  Invoice_type <> 'EXP' and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
                    End If

                End If
            Else
                salesconf = "update saleconf set current_No = " & mSaleConfNo & " where  UNIT_CODE = '" & gstrUNITID & "' and Invoice_type = 'EXP' and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
            End If
        Else

            If DataExist("SELECT TOP 1 1 FROM SALES_PARAMETER  WHERE SINGLE_INVOICE_SERIES= 1 and UNIT_CODE='" + gstrUNITID + "'") Then
                If Not mblnSameSeries Then
                    '10869291
                    If lbldescription.Text.ToUpper = "TRF" Or lbldescription.Text.ToUpper = "ITD" Then
                        If IsGSTINSAME(mAccount_Code) And (lbldescription.Text.ToUpper = "TRF" Or lbldescription.Text.ToUpper = "ITD") Then
                            salesconf = "update saleconf set CURRENT_NO_TRF_SAMEGSTIN = " & mSaleConfNo & " WHERE UNIT_CODE IN( SELECT UNIT_CODE FROM SALES_PARAMETER WHERE SINGLE_INVOICE_SERIES=1)  AND  Invoice_type IN('TRF','ITD') and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
                        Else
                            salesconf = "update saleconf set current_No = " & mSaleConfNo & " WHERE UNIT_CODE IN( SELECT UNIT_CODE FROM SALES_PARAMETER WHERE SINGLE_INVOICE_SERIES=1)  AND  Invoice_type IN('TRF','ITD') and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
                        End If
                    Else
                        If IsGSTINSAME(mAccount_Code) And (lbldescription.Text.ToUpper = "TRF" Or lbldescription.Text.ToUpper = "ITD") Then
                            salesconf = "update saleconf set CURRENT_NO_TRF_SAMEGSTIN = " & mSaleConfNo & " WHERE UNIT_CODE IN( SELECT UNIT_CODE FROM SALES_PARAMETER WHERE SINGLE_INVOICE_SERIES=1)  AND  Invoice_type = '" & Me.lbldescription.Text & "' " & " and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
                        Else
                            salesconf = "update saleconf set current_No = " & mSaleConfNo & " WHERE UNIT_CODE IN( SELECT UNIT_CODE FROM SALES_PARAMETER WHERE SINGLE_INVOICE_SERIES=1)  AND  Invoice_type = '" & Me.lbldescription.Text & "' " & " and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
                        End If

                    End If
                Else
                    salesconf = "update saleconf set current_No = " & mSaleConfNo & " WHERE  UNIT_CODE IN( SELECT UNIT_CODE FROM SALES_PARAMETER WHERE SINGLE_INVOICE_SERIES=1)  AND Single_Series = 1 " & " and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
                End If
            Else
                If Not mblnSameSeries Then
                    '10869291
                    If lbldescription.Text.ToUpper = "TRF" Or lbldescription.Text.ToUpper = "ITD" Then
                        If IsGSTINSAME(mAccount_Code) And (lbldescription.Text.ToUpper = "TRF" Or lbldescription.Text.ToUpper = "ITD") Then
                            salesconf = "update saleconf set CURRENT_NO_TRF_SAMEGSTIN = " & mSaleConfNo & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_type in ('ITD','TRF') and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
                        Else
                            salesconf = "update saleconf set current_No = " & mSaleConfNo & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_type in ('ITD','TRF') and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
                        End If
                    Else
                        If IsGSTINSAME(mAccount_Code) And (lbldescription.Text.ToUpper = "TRF" Or lbldescription.Text.ToUpper = "ITD") Then
                            salesconf = "update saleconf set CURRENT_NO_TRF_SAMEGSTIN = " & mSaleConfNo & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_type = '" & Me.lbldescription.Text & "' and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
                        Else
                            salesconf = "update saleconf set current_No = " & mSaleConfNo & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_type = '" & Me.lbldescription.Text & "' and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
                        End If
                    End If
                Else
                    If IsGSTINSAME(mAccount_Code) And (lbldescription.Text.ToUpper = "TRF" Or lbldescription.Text.ToUpper = "ITD") Then
                        salesconf = "update saleconf set CURRENT_NO_TRF_SAMEGSTIN = " & mSaleConfNo & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Single_Series = 1 and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
                    Else
                        salesconf = "update saleconf set current_No = " & mSaleConfNo & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Single_Series = 1 and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
                    End If

                End If
            End If
        End If
        mP_Connection.Execute("INSERT INTO INV_ERROR_DTL(QUERY,UNIT_CODE) VALUES('" & Replace(salesconf, "'", "") & "','" & gstrUNITID & "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        mP_Connection.Execute(salesconf, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

        '24 aug 2016'
        '06 july 2017'
        Dim rsSalesParameter1 As New ADODB.Recordset

        If rsSalesParameter1.State = ADODB.ObjectStateEnum.adStateOpen Then rsSalesParameter1.Close()
        rsSalesParameter1.Open("SELECT TotalInvoiceAmount_RoundOff, TotalInvoiceAmountRoundOff_Decimal FROM SALES_PARAMETER WHERE UNIT_CODE='" + gstrUNITID + "'", mP_Connection)
        If Not rsSalesParameter1.EOF Then
            mblnTotalInvoiceAmountRoundOff = rsSalesParameter1.Fields("TotalInvoiceAmount_RoundOff").Value
            mintTotalInvoiceAmountRoundOff = rsSalesParameter1.Fields("TotalInvoiceAmountRoundOff_Decimal").Value
        End If
        If rsSalesParameter1.State = ADODB.ObjectStateEnum.adStateOpen Then
            rsSalesParameter1.Close()
            rsSalesParameter1 = Nothing
        End If

        '06 july 2017
        ''
        If mblnTotalInvoiceAmountRoundOff = True Then
            If mblnlorryno = True Then
                saleschallan = "UPDATE SalesChallan_Dtl SET doc_no=" & mInvNo & ", total_amount = " & System.Math.Round(dblInvoiceAmt, 0, MidpointRounding.AwayFromZero) & ", Bill_Flag=1,print_flag = 1 ,LORRYNO_DATE= '" & txtlorryno.Text.Trim & "'  WHERE Doc_No=" & Ctlinvoice.Text & " and UNIT_CODE = '" & gstrUNITID & "'  and Location_Code='" & Trim(txtUnitCode.Text) & "' " & vbCrLf
            Else
                saleschallan = "UPDATE SalesChallan_Dtl SET doc_no=" & mInvNo & ", total_amount = " & System.Math.Round(dblInvoiceAmt, 0, MidpointRounding.AwayFromZero) & ", Bill_Flag=1,print_flag = 1 WHERE Doc_No=" & Ctlinvoice.Text & " and UNIT_CODE = '" & gstrUNITID & "'  and Location_Code='" & Trim(txtUnitCode.Text) & "' " & vbCrLf
            End If
        Else
            If mblnlorryno = True Then
                saleschallan = "UPDATE SalesChallan_Dtl SET doc_no=" & mInvNo & ", total_amount = " & System.Math.Round(dblInvoiceAmt, mintTotalInvoiceAmountRoundOff) & ", Bill_Flag=1,print_flag = 1 ,LORRYNO_DATE= '" & txtlorryno.Text.Trim & "' WHERE Doc_No=" & Ctlinvoice.Text & " and UNIT_CODE = '" & gstrUNITID & "'  and Location_Code='" & Trim(txtUnitCode.Text) & "' " & vbCrLf
            Else
                saleschallan = "UPDATE SalesChallan_Dtl SET doc_no=" & mInvNo & ", total_amount = " & System.Math.Round(dblInvoiceAmt, mintTotalInvoiceAmountRoundOff) & ", Bill_Flag=1,print_flag = 1 WHERE Doc_No=" & Ctlinvoice.Text & " and UNIT_CODE = '" & gstrUNITID & "'  and Location_Code='" & Trim(txtUnitCode.Text) & "' " & vbCrLf
            End If
        End If
        'saleschallan = "UPDATE SalesChallan_Dtl SET doc_no=" & mInvNo & ", total_amount = " & System.Math.Round(dblInvoiceAmt, 0) & ", Bill_Flag=1,print_flag = 1 WHERE Doc_No=" & Ctlinvoice.Text & " and UNIT_CODE = '" & gstrUNITID & "'  and Location_Code='" & Trim(txtUnitCode.Text) & "' " & vbCrLf
        '24 aug 2016'
        saleschallan = saleschallan & "UPDATE Sales_Dtl SET doc_no=" & mInvNo & " WHERE Doc_No=" & Ctlinvoice.Text & " and Location_Code='" & Trim(txtUnitCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "'"

        saleschallan = saleschallan & "UPDATE PCA_SCHEDULE_INVOICE_KNOCKOFF SET INVOICE_NO=" & mInvNo & ", Upd_Dt = getdate() ,Upd_Userid = '" & mP_User & "'   WHERE INVOICE_NO=" & Ctlinvoice.Text & " and UNIT_CODE = '" & gstrUNITID & "'"

        updatesalesconfandsaleschallan = True
        Exit Function
Err_Handler:
        SqlConnectionclass.ExecuteNonQuery("INSERT INTO INV_ERR_PRIMARYKEY(ERR_NUMBER,ERR_DESCRIPTION,TMPINVOICENO,INVOICE_NO,UNIT_CODE,FUNCTIONNAME ) VALUES('" & Err.Number & "','" & Replace(Err.Description, "'", "") & "','" & Ctlinvoice.Text & "','" & mInvNo & "','" & gstrUNITID & "','updatesaleconfandsaleschallan')")
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)

        updatesalesconfandsaleschallan = False
    End Function
    Public Function ValidSelection() As Boolean
        Dim blnInvalidData As Boolean
        Dim strErrMsg As String
        Dim ctlBlank As System.Windows.Forms.Control
        Dim lNo As Integer
        On Error GoTo Err_Handler
        ValidRecord = False
        lNo = 1
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
            Call Msgbox(strErrMsg, MsgBoxStyle.Information, "Error")
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
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Public Function InvoiceGeneration(ByRef RdAddSold As ReportDocument, ByRef Frm As Object) As Boolean

        Call Logging_Starting_End_Time("Invoice locking: InvoiceGeneration Started from Inside: Trans Count :" + GetTransCountIfAvailable().ToString(), DateTime.Now.ToString(), "Saved", mInvNo)

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
        Dim strInvoicestatus As String
        Dim STRCUSTOMERCODE As String
        Dim oCmd As ADODB.Command
        Dim strIPAddress As String
        Dim strStartingPDSMultipleQry As String
        Dim dblTCStaxAmt As Double
        Dim ShipGstinid, Shipgststatecode As String

        On Error GoTo Err_Handler
        'tcs changes
        dblTCStaxAmt = Val(Trim(Find_Value("SELECT isnull(SUM(TCSTAXAMOUNT),0) AS TCSTAXAMOUNT FROM SALES_DTL WHERE UNIT_CODE='" + gstrUNITID + "'AND DOC_NO='" & Trim(Ctlinvoice.Text) & "'")))

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
                Msgbox("Error encountered while Calculating TCS Item level .Please try Again.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, My.Resources.resEmpower.STR100)
                oCmd = Nothing
                Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                Exit Function
            End If
            oCmd = Nothing

        End If

        'tcs changes

        strIPAddress = gstrIpaddressWinSck
        rsCompMst = New ClsResultSetDB_Invoice

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

        Call Logging_Starting_End_Time(mCheckValARana + " Invoice locking: Going to Initialize Values: Trans Count :" + GetTransCountIfAvailable().ToString(), DateTime.Now.ToString(), "Saved", mInvNo)

        If optInvYes(1).Checked = False And mblnLock_Clicked Then
            Call InitializeValues()
            If Not ValuetoVariables() Then
                InvoiceGeneration = False
                Exit Function
            End If
            If mblnEOUUnit = True Then
                If lbldescription.Text <> "EXP" Then
                    If mOpeeningBalance < mAssessableValue Then
                        Msgbox("Opening Balance is Less then Invoice Assessable Value", MsgBoxStyle.Information, "empower")
                        InvoiceGeneration = False
                        Exit Function
                    End If
                End If
            End If
            '26 sep 2016'
            If mblnlorryno = True Then
                If txtlorryno.Enabled = True And txtlorryno.Text.ToString.Length = 0 Then
                    Msgbox("Lorry No .Can't be Empty ", MsgBoxStyle.Information, ResolveResString(100))
                    InvoiceGeneration = False
                    Exit Function
                End If
            End If
            '26 sep 2016'

            If mblnpostinfin = True Then
                If Not CreateStringForAccounts() Then
                    InvoiceGeneration = False
                    Exit Function
                End If
            Else

            End If


            Call Logging_Starting_End_Time("Invoice locking: CreateStringForAccounts Completed: Trans Count :" + GetTransCountIfAvailable().ToString(), DateTime.Now.ToString(), "Saved", mInvNo)

            If Not updatesalesconfandsaleschallan() Then
                InvoiceGeneration = False
                Exit Function
            End If

            Call Logging_Starting_End_Time("Invoice locking: updatesalesconfandsaleschallan Completed: Trans Count :" + GetTransCountIfAvailable().ToString(), DateTime.Now.ToString(), "Saved", mInvNo)

            If Not UpdateinSale_Dtl() Then
                InvoiceGeneration = False
                Exit Function
            End If

            Call Logging_Starting_End_Time("Invoice locking: UpdateinSale_Dtl Completed: Trans Count :" + GetTransCountIfAvailable().ToString(), DateTime.Now.ToString(), "Saved", mInvNo)

            If gstrUNITID = "STH" And optInvYes(0).Checked = True Then
                mP_Connection.Execute("UPDATE DELIVERY_MKT_ACKN_HISTORY set doc_no='" & mInvNo & "' where doc_no= " & Val(Ctlinvoice.Text) & " AND UNIT_CODE = '" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End If

            If UCase(lbldescription.Text) = "REJ" Then
                If Len(Trim(mCust_Ref)) > 0 Then
                    If Not UpdateGrnHdr(CDbl(Val(mCust_Ref)), mInvNo) Then
                        InvoiceGeneration = False
                        Exit Function
                    End If
                End If
            End If
        End If

        If UCase(lbldescription.Text) = "JOB" Then
            mP_Connection.Execute("DELETE FROM  tempCustAnnex WHERE UNIT_CODE = '" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords) ' to delete all the records from table before inserting new one for selected invoice
            If mblnBOMcheck_flag = True Then
                If BomCheck() = False Then
                    InvoiceGeneration = False
                    Exit Function
                End If
            End If
        End If
        If optInvYes(0).Checked = True Then
            If DataExist("SELECT TOP 1 1 FROM SALES_PARAMETER WHERE INVOICE_LOCKING_ENTRY_SAMEDATE=1  and UNIT_CODE = '" & gstrUNITID & "'") Then
                mP_Connection.Execute("update Saleschallan_dtl set invoice_date= Convert(varchar(12), getdate(), 106) WHERE UNIT_CODE='" + gstrUNITID + "' AND  Doc_no = '" & Ctlinvoice.Text & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End If
        End If

        Call Logging_Starting_End_Time("Invoice locking: INVOICE_LOCKING_ENTRY_SAMEDATE Checked: Trans Count :" + GetTransCountIfAvailable().ToString(), DateTime.Now.ToString(), "Saved", mInvNo)


        Address = gstr_WRK_ADDRESS1 & gstr_WRK_ADDRESS2
        rsCompMst = New ClsResultSetDB_Invoice
        rsCompMst.GetResult("Select a.* from Customer_Mst a, saleschallan_dtl b where a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND a.Customer_code = b.Account_code and b.Doc_No = " & Ctlinvoice.Text & " and b.Location_Code='" & Trim(txtUnitCode.Text) & "'")
        If rsCompMst.GetNoRows > 0 Then
            blnCSIEX_Inc = rsCompMst.GetValue("CSIEX_Inc")
            DeliveredAdd = Trim(rsCompMst.GetValue("Ship_address1"))
            If Len(Trim(DeliveredAdd)) Then
                DeliveredAdd = Trim(DeliveredAdd) & "," & Trim(rsCompMst.GetValue("Ship_address2"))
            Else
                DeliveredAdd = Trim(rsCompMst.GetValue("Ship_address2"))
            End If
        End If
        rsCompMst.ResultSetClose()


        Dim strInvoiceDate As String
        Dim dblExistingInvNo As Double
        Dim strSql1 As String


        'If CUST_REJECTION_FLAG = True Then
        'If gstrUNITID = "M03" Then
        'RdAddSold.Load(My.Application.Info.DirectoryPath & "\Reports\rptInvoicerejcustomerMateCh.rpt")
        'ElseIf gstrUNITID = "M02" Then
        'RdAddSold.Load(My.Application.Info.DirectoryPath & "\Reports\rptInvoicerejcustomerMateCh_mateunit1.rpt")
        'Else
        'RdAddSold.Load(My.Application.Info.DirectoryPath & "\Reports\rptInvoicerejcustomerMateCh_FSP.rpt")
        'End If
        'Else
        'RdAddSold.Load(My.Application.Info.DirectoryPath & "\Reports\" & mstrReportFilename & ".rpt")
        'End If

        Call Logging_Starting_End_Time("Invoice locking: Going to Initialize Report: Trans Count :" + GetTransCountIfAvailable().ToString(), DateTime.Now.ToString(), "Saved", mInvNo)

        RdAddSold = Frm.GetReportDocument()
        STRCUSTOMERCODE = Trim(Find_Value("SELECT ACCOUNT_CODE FROM SALESCHALLAN_DTL WHERE UNIT_CODE='" + gstrUNITID + "'AND DOC_NO='" & Trim(Ctlinvoice.Text) & "'"))
        If gstrUNITID = "STH" Then
            oCmd = New ADODB.Command
            With oCmd
                .ActiveConnection = mP_Connection
                .CommandTimeout = 0
                .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                .CommandText = "PRC_INVOICEPRINTING_SMRC"
                .Parameters.Append(.CreateParameter("@UnitCode", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                .Parameters.Append(.CreateParameter("@LOC_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 8, Trim(txtUnitCode.Text)))
                .Parameters.Append(.CreateParameter("@DOC_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, Trim(Ctlinvoice.Text)))
                .Parameters.Append(.CreateParameter("@INV_TYPE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 3, Trim(Me.lbldescription.Text)))
                .Parameters.Append(.CreateParameter("@INV_SUBTYPE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Me.lblcategory.Text)))
                .Parameters.Append(.CreateParameter("@IP_ADDRESS", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, strIPAddress.Trim()))
                .Parameters.Append(.CreateParameter("@ERRCODE", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInputOutput))
                .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End With
            If oCmd.Parameters("@ERRCODE").Value <> 0 Then
                Msgbox("Error encountered while generating data for report.Please try Again.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, My.Resources.resEmpower.STR100)
                oCmd = Nothing
                Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                Exit Function
            End If
            oCmd = Nothing
            RdAddSold.Load(My.Application.Info.DirectoryPath & "\Reports\" & mstrReportFilename & ".rpt")

        ElseIf IsGSTINSAME(STRCUSTOMERCODE) = True And (UCase(Trim(cmbInvType.Text)) = "TRANSFER INVOICE" Or UCase(Trim(cmbInvType.Text)) = "INTER-DIVISION") Then
            oCmd = New ADODB.Command
            With oCmd
                .ActiveConnection = mP_Connection
                .CommandTimeout = 0
                .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                .CommandText = "PRC_INVOICEPRINTING_MATE"
                .Parameters.Append(.CreateParameter("@UnitCode", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                .Parameters.Append(.CreateParameter("@LOC_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 8, Trim(txtUnitCode.Text)))
                .Parameters.Append(.CreateParameter("@DOC_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, Trim(Ctlinvoice.Text)))
                .Parameters.Append(.CreateParameter("@INV_TYPE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 3, Trim(Me.lbldescription.Text)))
                .Parameters.Append(.CreateParameter("@INV_SUBTYPE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Me.lblcategory.Text)))
                .Parameters.Append(.CreateParameter("@IP_ADDRESS", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, strIPAddress.Trim()))
                .Parameters.Append(.CreateParameter("@ERRCODE", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInputOutput))
                .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End With
            If oCmd.Parameters("@ERRCODE").Value <> 0 Then
                Msgbox("Error encountered while generating data for report.Please try Again.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, My.Resources.resEmpower.STR100)
                oCmd = Nothing
                Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                Exit Function
            End If
            oCmd = Nothing
            RdAddSold.Load(My.Application.Info.DirectoryPath & "\Reports\Delivery_Challan_GST_A4REPORTS.rpt")
        ElseIf CBool(Find_Value("select REJINVOICE_GRIN_PVKNOCKING from sales_parameter WHERE UNIT_CODE='" + gstrUNITID + "'")) = True And UCase(Trim(cmbInvType.Text)) = "REJECTION" Then
            oCmd = New ADODB.Command
            With oCmd
                .ActiveConnection = mP_Connection
                .CommandTimeout = 0
                .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                .CommandText = "PRC_INVOICEPRINTING_MATE"
                .Parameters.Append(.CreateParameter("@UnitCode", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                .Parameters.Append(.CreateParameter("@LOC_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 8, Trim(txtUnitCode.Text)))
                .Parameters.Append(.CreateParameter("@DOC_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, Trim(Ctlinvoice.Text)))
                .Parameters.Append(.CreateParameter("@INV_TYPE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 3, Trim(Me.lbldescription.Text)))
                .Parameters.Append(.CreateParameter("@INV_SUBTYPE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Me.lblcategory.Text)))
                .Parameters.Append(.CreateParameter("@IP_ADDRESS", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, strIPAddress.Trim()))
                .Parameters.Append(.CreateParameter("@ERRCODE", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInputOutput))
                .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End With
            If oCmd.Parameters("@ERRCODE").Value <> 0 Then
                Msgbox("Error encountered while generating data for report.Please try Again.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, My.Resources.resEmpower.STR100)
                oCmd = Nothing
                Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                Exit Function
            End If
            oCmd = Nothing
            RdAddSold.Load(My.Application.Info.DirectoryPath & "\Reports\Delivery_Challan_GST_REJECTION.rpt")
            '01 jan 2024
        ElseIf (UCase(Trim(gstrUNITID)) = "MSD" And GetPrintMethod(STRCUSTOMERCODE).ToUpper() = "TATA") Then
            oCmd = New ADODB.Command
            With oCmd
                .ActiveConnection = mP_Connection
                .CommandTimeout = 0
                .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                .CommandText = "PRC_INVOICEPRINTING_MATE"
                .Parameters.Append(.CreateParameter("@UnitCode", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                .Parameters.Append(.CreateParameter("@LOC_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 8, Trim(txtUnitCode.Text)))
                .Parameters.Append(.CreateParameter("@DOC_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, Trim(Ctlinvoice.Text)))
                .Parameters.Append(.CreateParameter("@INV_TYPE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 3, Trim(Me.lbldescription.Text)))
                .Parameters.Append(.CreateParameter("@INV_SUBTYPE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Me.lblcategory.Text)))
                .Parameters.Append(.CreateParameter("@IP_ADDRESS", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, strIPAddress.Trim()))
                .Parameters.Append(.CreateParameter("@ERRCODE", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInputOutput))
                .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End With
            If oCmd.Parameters("@ERRCODE").Value <> 0 Then
                Msgbox("Error encountered while generating data for report.Please try Again.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, My.Resources.resEmpower.STR100)
                oCmd = Nothing
                Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                Exit Function
            End If
            oCmd = Nothing
            'RdAddSold.Load(My.Application.Info.DirectoryPath & "\Reports\Delivery_Challan_GST_REJECTION.rpt")
            RdAddSold.Load(My.Application.Info.DirectoryPath & "\Reports\" & mstrReportFilename & ".rpt")
            '01 jan 2024
        ElseIf gstrUNITID = "STH" Then
            oCmd = New ADODB.Command
            With oCmd
                .ActiveConnection = mP_Connection
                .CommandTimeout = 0
                .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                .CommandText = "PRC_INVOICEPRINTING_SMRC"
                .Parameters.Append(.CreateParameter("@UnitCode", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                .Parameters.Append(.CreateParameter("@LOC_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 8, Trim(txtUnitCode.Text)))
                .Parameters.Append(.CreateParameter("@DOC_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, Trim(Ctlinvoice.Text)))
                .Parameters.Append(.CreateParameter("@INV_TYPE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 3, Trim(Me.lbldescription.Text)))
                .Parameters.Append(.CreateParameter("@INV_SUBTYPE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Me.lblcategory.Text)))
                .Parameters.Append(.CreateParameter("@IP_ADDRESS", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, strIPAddress.Trim()))
                .Parameters.Append(.CreateParameter("@ERRCODE", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInputOutput))
                .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End With
            If oCmd.Parameters("@ERRCODE").Value <> 0 Then
                Msgbox("Error encountered while generating data for report.Please try Again.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, My.Resources.resEmpower.STR100)
                oCmd = Nothing
                Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                Exit Function
            End If
            oCmd = Nothing
            'RdAddSold.Load(My.Application.Info.DirectoryPath & "\Reports\Delivery_Challan_GST_REJECTION.rpt")
            RdAddSold.Load(My.Application.Info.DirectoryPath & "\Reports\" & mstrReportFilename & ".rpt")
            '01 jan 2024

        Else

            If CUST_REJECTION_FLAG = True Then
                If UCase(Trim(cmbInvType.Text)) = "REJECTION" And CBool(Find_Value("select REJINVOICE_GRIN_PVKNOCKING from sales_parameter WHERE UNIT_CODE='" + gstrUNITID + "'")) = True Then
                    RdAddSold.Load(My.Application.Info.DirectoryPath & "\Reports\Delivery_Challan_GST_REJECTION.rpt")
                Else
                    If gstrUNITID = "M03" Then
                        RdAddSold.Load(My.Application.Info.DirectoryPath & "\Reports\rptInvoicerejcustomerMateCh.rpt")
                    ElseIf gstrUNITID = "M02" Then
                        RdAddSold.Load(My.Application.Info.DirectoryPath & "\Reports\rptInvoicerejcustomerMateCh_mateunit1.rpt")
                    Else
                        If mblncustomerlevel_A4report_functionlity = True And mblnA4reports_invoicewise = True Then
                            RdAddSold.Load(My.Application.Info.DirectoryPath & "\Reports\rptInvoicerejcustomerMateCh_FSP_A4reports.rpt")
                        Else
                            RdAddSold.Load(My.Application.Info.DirectoryPath & "\Reports\rptInvoicerejcustomerMateCh_FSP.rpt")
                        End If
                    End If
                End If
            Else
                strStartingPDSMultipleQry = "set dateformat 'dmy' SELECT * FROM  Saleschallan_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND invoice_type='" & Trim(Me.lbldescription.Text) & "' and sub_category='" & Trim(Me.lblcategory.Text) & "' and bill_flag = '1' and cancel_flag = '0' and Location_Code='" & Trim(txtUnitCode.Text) & "'"
                strStartingPDSMultipleQry += " AND (( INVOICE_DATE >= '" & mblnTOYOTA_MULTIPLESO_ONEPDS_STDATE & "')) "
                If optInvYes(1).Checked = True Then
                    strStartingPDSMultipleQry += " and doc_no=" & Ctlinvoice.Text
                End If

                If UCase(Trim(cmbInvType.Text)) = "REJECTION" And CBool(Find_Value("select REJINVOICE_GRIN_PVKNOCKING from sales_parameter WHERE UNIT_CODE='" + gstrUNITID + "'")) = True Then
                    RdAddSold.Load(My.Application.Info.DirectoryPath & "\Reports\Delivery_Challan_GST_REJECTION.rpt")
                Else
                    If ((CBool(Find_Value("SELECT ISNULL(MULTIPLE_SO_PDS_TOYOTA,0)as MULTIPLE_SO_PDS_TOYOTA FROM SALES_PARAMETER where unit_code = '" & gstrUNITID & "'")) = True)) Then
                        If DataExist(strStartingPDSMultipleQry) = True And ((CBool(Find_Value("SELECT ISNULL(PDS_TOYOTA_CUSTOMER,0)as PDS_TOYOTA_CUSTOMER  FROM CUSTOMER_MST WHERE UNIT_CODE = '" & gstrUNITID & "'AND CUSTOMER_CODE ='" & STRCUSTOMERCODE & "'")) = True)) Then
                            If mblncustomerlevel_A4report_functionlity = True And mblnA4reports_invoicewise = True Then
                                RdAddSold.Load(My.Application.Info.DirectoryPath & "\Reports\" & mstrReportFilename & "_MultipleSO_A4reports.rpt")
                            End If
                        Else
                            If mblncustomerlevel_A4report_functionlity = True And mblnA4reports_invoicewise = True Then
                                RdAddSold.Load(My.Application.Info.DirectoryPath & "\Reports\" & mstrReportFilename & "_A4reports.rpt")
                            Else
                                RdAddSold.Load(My.Application.Info.DirectoryPath & "\Reports\" & mstrReportFilename & ".rpt")
                            End If
                        End If
                    End If
                End If
            End If

        End If

        '---------------------------------------------
        If gstrUNITID <> "MST" Then
            Frm.EnablePrintButton = True
        End If

        '--------------------------------------------
        If IsGSTINSAME(STRCUSTOMERCODE) = True And (UCase(Trim(cmbInvType.Text)) = "TRANSFER INVOICE" Or UCase(Trim(cmbInvType.Text)) = "INTER-DIVISION") Then
            strSql = "{TMP_INVOICEPRINT.IP_ADDRESS}='" & strIPAddress.Trim() & "'  and {TMP_INVOICEPRINT.Unit_Code}='" & gstrUNITID & "'"
        ElseIf UCase(Trim(gstrUNITID)) = "MSD" And GetPrintMethod(STRCUSTOMERCODE).ToUpper() = "TATA" Then
            strSql = "{TMP_INVOICEPRINT.IP_ADDRESS}='" & strIPAddress.Trim() & "'  and {TMP_INVOICEPRINT.Unit_Code}='" & gstrUNITID & "'"
        ElseIf UCase(Trim(gstrUNITID)) = "STH" Then
            strSql = "{TMP_INVOICEPRINT.IP_ADDRESS}='" & strIPAddress.Trim() & "'  and {TMP_INVOICEPRINT.Unit_Code}='" & gstrUNITID & "'"
        Else
            If CBool(Find_Value("select REJINVOICE_GRIN_PVKNOCKING from sales_parameter WHERE UNIT_CODE='" + gstrUNITID + "'")) And UCase(Trim(cmbInvType.Text)) = "REJECTION" Then
                strSql = "{TMP_INVOICEPRINT.IP_ADDRESS}='" & strIPAddress.Trim() & "'  and {TMP_INVOICEPRINT.Unit_Code}='" & gstrUNITID & "'"
            Else
                strSql = "{SalesChallan_Dtl.Location_Code}='" & Trim(txtUnitCode.Text) & "' and {SalesChallan_Dtl.Doc_No} =" & Trim(Ctlinvoice.Text) & " and {SalesChallan_Dtl.UNIT_CODE} ='" & gstrUNITID & "' "
            End If

        End If
        RdAddSold.DataDefinition.RecordSelectionFormula = strSql
        Dim RSCNT As ADODB.Recordset
        RSCNT = mP_Connection.Execute("select isnull(MultipleSO,0) from  saleschallan_dtl (NOLOCK) where doc_no = " & Val(Ctlinvoice.Text) & " AND UNIT_CODE = '" & gstrUNITID & "'")
        If RSCNT.Fields(0).Value = True Then
            'RdAddSold.Load(My.Application.Info.DirectoryPath & "\Reports\rptInvoiceMateCh_MUL_SO.rpt")
            mP_Connection.Execute("update SalesChallan_Dtl set multipleso=0 where doc_no= " & Val(Ctlinvoice.Text) & " AND UNIT_CODE = '" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            'SqlConnectionclass.ExecuteNonQuery("update SalesChallan_Dtl set multipleso=0 where doc_no= " & Val(Ctlinvoice.Text) & " AND UNIT_CODE = '" & gstrUNITID & "'")
        End If
        If UCase(cmbInvType.Text) <> "JOBWORK INVOICE" Then
            RdAddSold.DataDefinition.FormulaFields("Category").Text = "'" & Me.lblcategory.Text & "'"
        End If
        If gstrUNITID = "STH" Then
            If mInvNo = 0 Then
                mInvNo = Ctlinvoice.Text
            End If

        End If
        If optInvYes(0).Checked = True Then
            RdAddSold.DataDefinition.FormulaFields("InvoiceNo").Text = "'" & mInvNo & "'"
        Else
            RdAddSold.DataDefinition.FormulaFields("InvoiceNo").Text = "'" & CDbl(Val(Ctlinvoice.Text)) & "'"
        End If
        RdAddSold.DataDefinition.FormulaFields("Registration").Text = "'" & RegNo & "'"
        RdAddSold.DataDefinition.FormulaFields("Range").Text = "'" & Range & "'"
        RdAddSold.DataDefinition.FormulaFields("CompanyName").Text = "'" & gstrCOMPANY & "'"
        RdAddSold.DataDefinition.FormulaFields("CompanyAddress").Text = "'" & Address & "'"
        RdAddSold.DataDefinition.FormulaFields("Phone").Text = "'" & Phone & "'"
        RdAddSold.DataDefinition.FormulaFields("Fax").Text = "'" & Fax & "'"
        If UCase(cmbInvType.Text) <> "JOBWORK INVOICE" Then
            RdAddSold.DataDefinition.FormulaFields("EMail").Text = "'" & EMail & "'"
        End If
        RdAddSold.DataDefinition.FormulaFields("PLA").Text = "'" & PLA & "'"
        RdAddSold.DataDefinition.FormulaFields("UPST").Text = "'" & UPST & "'"
        RdAddSold.DataDefinition.FormulaFields("CST").Text = "'" & CST & "'"
        RdAddSold.DataDefinition.FormulaFields("Division").Text = "'" & Division & "'"
        RdAddSold.DataDefinition.FormulaFields("commissionerate").Text = "'" & Commissionerate & "'"
        RdAddSold.DataDefinition.FormulaFields("InvoiceRule").Text = "'" & Invoice_Rule & "'"
        RdAddSold.DataDefinition.FormulaFields("EOUFlag").Text = "'" & mblnEOUUnit & "'"
        If optYes(0).Checked = True Then
            If (gstrUNITID = "MST" Or gstrUNITID = "MSB") Then
                DeliveredAdd = ""
                rsCompMst = New ClsResultSetDB_Invoice
                rsCompMst.GetResult("Select * from VW_SHIPPINGCODE_DESC  where UNIT_CODE = '" & gstrUNITID & "'and Doc_No = '" & Ctlinvoice.Text & "'")
                If rsCompMst.GetNoRows > 0 Then
                    ShipGstinid = Trim(rsCompMst.GetValue("GSTIN_ID"))
                    Shipgststatecode = Trim(rsCompMst.GetValue("GST_STATE_CODE"))

                    DeliveredAdd = Trim(rsCompMst.GetValue("Ship_address1"))
                    If Len(Trim(DeliveredAdd)) Then
                        DeliveredAdd = Trim(DeliveredAdd) & "," & Trim(rsCompMst.GetValue("Ship_address2"))
                    Else
                        DeliveredAdd = Trim(rsCompMst.GetValue("Ship_address2"))
                    End If
                End If
                rsCompMst.ResultSetClose()

                RdAddSold.DataDefinition.FormulaFields("DeliveredAt").Text = "' Delivered At '"
                RdAddSold.DataDefinition.FormulaFields("Address2").Text = "'" & DeliveredAdd & "'"
                If mstrReportFilename = "rptInvoiceMateCh_BANG_GST" Then
                    RdAddSold.DataDefinition.FormulaFields("ShipGstinid").Text = "'" & ShipGstinid & "'"
                    RdAddSold.DataDefinition.FormulaFields("Shipgststatecode").Text = "'" & Shipgststatecode & "'"
                End If

            Else
                RdAddSold.DataDefinition.FormulaFields("DeliveredAt").Text = "' Delivered At '"
                RdAddSold.DataDefinition.FormulaFields("Address2").Text = "'" & DeliveredAdd & "'"
                'RdAddSold.DataDefinition.FormulaFields("ShipGstinid).Text = "''"
                'RdAddSold.DataDefinition.FormulaFields("Shipgststatecode").Text = "''"

            End If

        Else
            RdAddSold.DataDefinition.FormulaFields("DeliveredAt").Text = "''"
            RdAddSold.DataDefinition.FormulaFields("Address2").Text = "''" 'to pass blanck Address in this case will overwrite this Formula written in Crystal Report for else case
        End If
        '10688760
        If mbln_SHIPPING_ADDRESS = True And mbln_SHIPPING_ADDRESS_INVOICEWISE = True And optYes(0).Checked = True Then
            RdAddSold.DataDefinition.FormulaFields("Address2").Text = "'" & DeliveredAdd & "'"
        Else
            RdAddSold.DataDefinition.FormulaFields("Address2").Text = "''"
        End If
        '10688760
        RdAddSold.DataDefinition.FormulaFields("PLADuty").Text = "'" & Trim(txtPLA.Text) & "'"
        RdAddSold.DataDefinition.FormulaFields("InsuranceFlag").Text = "'" & mblnInsuranceFlag & "'"
        RdAddSold.DataDefinition.FormulaFields("StringYear").Text = "'" & Year(GetServerDate) & "'"
        RdAddSold.DataDefinition.FormulaFields("DateOfRemoval").Text = "'" & dtpRemoval.Text & "'"
        '13 dec 2017'
        If mlbnprintforexcurrency = True Then
            If ChkPrintForex.CheckState = System.Windows.Forms.CheckState.Checked Then
                RdAddSold.DataDefinition.FormulaFields("printforex").Text = "'BASE'"
            Else
                RdAddSold.DataDefinition.FormulaFields("printforex").Text = "''"
            End If
        End If

        '13 dec 2017'
        If UCase(cmbInvType.Text) = "REJECTION" Then
            strGRNDate = "" : strVendorInvDate = "" : strVendorInvNo = "" : strCustRefForGrn = ""
            rsGrnHdr = New ClsResultSetDB_Invoice
            rsGrnHdr.GetResult("Select Cust_ref from salesChallan_dtl where doc_No = " & Ctlinvoice.Text & " and UNIT_CODE = '" & gstrUNITID & "'")
            If rsGrnHdr.GetNoRows > 0 Then
                rsGrnHdr.MoveFirst()
                strCustRefForGrn = rsGrnHdr.GetValue("Cust_ref")
            End If
            rsGrnHdr.ResultSetClose()

            'added by priti on 20 April to add Supplier invoice no and date in case of LRN rejection
            If Len(Trim(strCustRefForGrn)) = 0 Then
                Dim strLRNNo = SqlConnectionclass.ExecuteScalar("SELECT top 1 Ref_Doc_No FROM MKT_INVREJ_DTL WHERE UNIT_CODE ='" + gstrUNITID + "' and invoice_no=" & Ctlinvoice.Text & " ")
                If Len(Trim(strLRNNo)) > 0 Then
                    strCustRefForGrn = SqlConnectionclass.ExecuteScalar("select top 1 grn_hdr.doc_no  from itembatch_dtl a join itembatch_dtl b " &
                                    "on a.unit_code = b.unit_code and a.item_code = b.item_code and a.batch_no = b.batch_no and b.doc_type=10  " &
                                    "join grn_hdr  on a.batch_no=convert(varchar,grn_hdr.doc_no,25) where a.unit_code='" + gstrUNITID + "' and a.doc_no=" & strLRNNo & " " &
                                    "and a.doc_type=12 and grn_hdr.grn_date  < = a.doc_date order by grn_date desc ")
                End If
            End If
            'Code ends by priti on 20 April to add Supplier invoice no and date in case of LRN rejection

            If Len(Trim(strCustRefForGrn)) > 0 Then
                rsGrnHdr = New ClsResultSetDB_Invoice
                rsGrnHdr.GetResult("select grn_date,Invoice_no,Invoice_date from grn_hdr where  UNIT_CODE = '" & gstrUNITID & "' and From_Location ='01R1' and doc_No = " & strCustRefForGrn)
                If rsGrnHdr.GetNoRows > 0 Then
                    rsGrnHdr.MoveFirst()
                    'strGRNDate = IIf(IsDBNull(rsGrnHdr.GetValue("grn_date")), "", VB.Format(rsGrnHdr.GetValue("grn_date"), gstrDateFormat))
                    strGRNDate = IIf(IsDBNull(rsGrnHdr.GetValue("grn_date")), "", Convert.ToDateTime(rsGrnHdr.GetValue("grn_date")).ToString(gstrDateFormat))
                    'strVendorInvDate = IIf(IsDBNull(rsGrnHdr.GetValue("invoice_date")), "", VB.Format(rsGrnHdr.GetValue("invoice_date"), gstrDateFormat))
                    strVendorInvDate = IIf(IsDBNull(rsGrnHdr.GetValue("invoice_date")), "", Convert.ToDateTime(rsGrnHdr.GetValue("invoice_date")).ToString(gstrDateFormat))
                    strVendorInvNo = rsGrnHdr.GetValue("Invoice_No")
                End If

                rsGrnHdr.ResultSetClose()
            End If

            Dim blnShowGrinNo As Boolean = CBool(Find_Value("SELECT isnull(ShowGrinno,0) as ShowGrinno  FROM SALES_PARAMETER WHERE UNIT_CODE='" + gstrUNITID + "'"))

            If blnShowGrinNo = True Then
                RdAddSold.DataDefinition.FormulaFields("GrinNo").Text = "'" + strCustRefForGrn + "'"   'added by priti on 20 April to add Supplier invoice no and date in case of LRN rejection
            End If

            RdAddSold.DataDefinition.FormulaFields("GrinDate").Text = "'" & strGRNDate & "'"
            RdAddSold.DataDefinition.FormulaFields("GrinInvoiceNo").Text = "'" & strVendorInvNo & "'"
            RdAddSold.DataDefinition.FormulaFields("GrinInvoiceDate").Text = "'" & strVendorInvDate & "'"
            MSTRREJECTIONNOTE = ""
            If CBool(Find_Value("select REJINVOICE_GRIN_PVKNOCKING from sales_parameter WHERE UNIT_CODE='" + gstrUNITID + "'")) = True Then
                If DataExist("SELECT TOP 1 1 FROM MKT_INVREJ_DTL WHERE UNIT_CODE = '" & gstrUNITID & "' AND REJ_TYPE=1 AND CANCEL_FLAG=0 AND INVOICE_NO=" & Ctlinvoice.Text) Then 'GRIN RELATED QUERY
                    MSTRREJECTIONNOTE = Find_Value("SELECT TOP 1 REJECTED_NOTE  FROM AP_DOCMASTER WHERE APDOCM_UNIT ='" & gstrUNITID & "' AND APDOCM_VONO IN(select PV_No from grn_hdr where UNIT_CODE='" & gstrUNITID & "' AND Doc_No ='" & strCustRefForGrn & "')")
                End If
                RdAddSold.DataDefinition.FormulaFields("REJECTIONNOTE").Text = "'" + MSTRREJECTIONNOTE + "'"
            End If

        End If

        If CBool(Find_Value("Select AllowTempoaryInvoiceTag from Sales_Parameter WHERE UNIT_CODE='" + gstrUNITID + "'")) = True And ((UCase(cmbInvType.Text) = "NORMAL INVOICE" And UCase(CmbCategory.Text) <> "SCRAP" And mblnLock_Clicked = False)) Then
            If optInvYes(0).Checked = True Then
                strInvoicestatus = "TEMPORARY INVOICE"
                mInvNo = "0"
                RdAddSold.DataDefinition.FormulaFields("InvoiceNo").Text = "'" & mInvNo & "'"
                RdAddSold.DataDefinition.FormulaFields("Invoicestatus").Text = "'" & strInvoicestatus & "'"
            Else
                strInvoicestatus = ""
                RdAddSold.DataDefinition.FormulaFields("Invoicestatus").Text = "'" & strInvoicestatus & "'"
            End If
        End If
        '17 apr 2015
        If mblncustomerlevel_A4report_functionlity = True And mblnA4reports_invoicewise = True And mblncustomerlevel_Annexture_printing = True And mblncustomerspecificreport = True Then
            Dim STRTOTALCOUNT As String
            Dim strTOTNOOFANNEXPAGES As String
            STRTOTALCOUNT = Find_Value("SELECT COUNT(*) FROM INVOICE_QRIMAGE WHERE  UNIT_CODE = '" & gstrUNITID & "' AND  INVOICE_NO=" & Trim(Ctlinvoice.Text))
            RdAddSold.DataDefinition.FormulaFields("TOTALBARCODEIMAGES").Text = STRTOTALCOUNT
            strTOTNOOFANNEXPAGES = Find_Value("SELECT CEILING(COUNT(*)/50.0) FROM SALES_DTL WHERE  UNIT_CODE = '" & gstrUNITID & "' AND  doc_no=" & Trim(Ctlinvoice.Text))
            RdAddSold.DataDefinition.FormulaFields("TOTALANNEXTUREPAGES").Text = strTOTNOOFANNEXPAGES
        End If
        '17 APR 2015
        If mblncustomerlevel_A4report_functionlity = True And mblnA4reports_invoicewise = True And mblncustomerlevel_Annexture_printing = True And mblncustomerspecificreport = True Then
            RdAddSold.DataDefinition.FormulaFields("ANNEXTURE").Text = False
            RdAddSold.DataDefinition.FormulaFields("TOYOTABARCODEPRINT").Text = False
            If DataExist("SELECT TOP 1 1 FROM SALES_DTL  WHERE UNIT_CODE = '" & gstrUNITID & "' and doc_no= " & Trim(Ctlinvoice.Text) & " HAVING COUNT(*)> " & mintmaxnoofitems_barcodeToyota) = True Then
                RdAddSold.DataDefinition.FormulaFields("TOYOTABARCODEPRINT").Text = True
            End If
            If DataExist("SELECT TOP 1 1 FROM SALES_DTL  WHERE UNIT_CODE = '" & gstrUNITID & "' and doc_no= " & Trim(Ctlinvoice.Text) & " HAVING COUNT(*)> " & mintmaxnoofitems_Annexture) = True Then
                RdAddSold.DataDefinition.FormulaFields("ANNEXTURE").Text = True
                Dim strTOTALQty As String
                Dim STRTOTALACCESSIBLE As String
                strTOTALQty = Find_Value("select sum(SALES_QUANTITY) from sales_dtl where  UNIT_CODE = '" & gstrUNITID & "' and  doc_no=" & Trim(Ctlinvoice.Text))
                RdAddSold.DataDefinition.FormulaFields("TOTALQTY").Text = strTOTALQty
                STRTOTALACCESSIBLE = Find_Value("select sum(accessible_amount) from sales_dtl where  UNIT_CODE = '" & gstrUNITID & "' and  doc_no=" & Trim(Ctlinvoice.Text))
                RdAddSold.DataDefinition.FormulaFields("TOTALACCESSIBLE").Text = STRTOTALACCESSIBLE
            End If
        End If
        If mblnA4reports_invoicewise = True And mblncustomerlevel_Annexture_printing = True And gstrUNITID = "MST" Then

            RdAddSold.DataDefinition.FormulaFields("ANNEXTURE").Text = True
            Dim strTOTALQty As String
            Dim STRTOTALACCESSIBLE As String
            strTOTALQty = Find_Value("select sum(SALES_QUANTITY) from sales_dtl where  UNIT_CODE = '" & gstrUNITID & "' and  doc_no=" & Trim(Ctlinvoice.Text))
            RdAddSold.DataDefinition.FormulaFields("TOTALQTY").Text = strTOTALQty
            STRTOTALACCESSIBLE = Find_Value("select sum(accessible_amount) from sales_dtl where  UNIT_CODE = '" & gstrUNITID & "' and  doc_no=" & Trim(Ctlinvoice.Text))
            RdAddSold.DataDefinition.FormulaFields("TOTALACCESSIBLE").Text = STRTOTALACCESSIBLE
        End If

        If mstrReportFilename = "" Then
            Msgbox("No Report filename selected for the invoice. Invoice cannot be printed", MsgBoxStyle.Information, "empower")
            Exit Function
        End If
        RSCNT.Close()
        InvoiceGeneration = True
        Call Logging_Starting_End_Time("Invoice locking: InvoiceGeneration Completed with Return Value True: Trans Count :" + GetTransCountIfAvailable().ToString(), DateTime.Now.ToString(), "Saved", mInvNo)

        Exit Function
Err_Handler:
        InvoiceGeneration = False ' AMIT 05APR2023
        SqlConnectionclass.ExecuteNonQuery("INSERT INTO INV_ERR_PRIMARYKEY(ERR_NUMBER,ERR_DESCRIPTION,TMPINVOICENO,INVOICE_NO,UNIT_CODE,FUNCTIONNAME ) VALUES('" & Err.Number & "','" & Replace(Err.Description, "'", "") & "','" & Ctlinvoice.Text & "','" & mInvNo & "','" & gstrUNITID & "','Invoicegeneration')")
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Public Sub InitializeValues()
        On Error GoTo ErrHandler
        mExDuty = 0 : mInvNo = 0 : mBasicAmt = 0 : msubTotal = 0 : mOtherAmt = 0 : mGrTotal = 0 : mStAmt = 0 : mFrAmt = 0
        mDoc_No = 0 : mCustmtrl = 0 : mAmortization = 0 : mstrAnnex = "" : strupdateGrinhdr = "" : mblnCustSupp = False
        Exit Sub
ErrHandler:
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Public Function ParentQty(ByRef pstrItemCode As String, ByRef pstrfinished As Object) As Double
        Dim rsParentQty As ClsResultSetDB_Invoice
        Dim strParentQty As String
        On Error GoTo ErrHandler
        strParentQty = "select sum(required_qty + waste_Qty) as TotalQty from Bom_Mst where finished_Product_code ='"
        strParentQty = strParentQty & pstrfinished & "' and rawMaterial_Code ='" & pstrItemCode & "'  and UNIT_CODE = '" & gstrUNITID & "'"
        rsParentQty = New ClsResultSetDB_Invoice
        rsParentQty.GetResult(strParentQty, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        ParentQty = rsParentQty.GetValue("TotalQty")
        rsParentQty.ResultSetClose()
        Exit Function
ErrHandler:
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Public Function InsertUpdateAnnex(ByRef parrCustAnnex As Object, ByRef pstrFinishedItem As Object, ByRef intMaxCount As Short) As Object
        Dim intloopcount As Short
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
        For intloopcount = 0 To intMaxCount
            rsVandBom = New ClsResultSetDB_Invoice
            rsVandBom.GetResult("Select RawMaterial_Code from Vendor_bom where Finish_Product_code = '" & pstrFinishedItem & "' and Vendor_code = '" & strCustCode & "' and rawMaterial_code ='" & parrCustAnnex(0, intloopcount) & "' and UNIT_CODE = '" & gstrUNITID & "'")
            If rsVandBom.GetNoRows > 0 Then
                strRef57F4 = Replace(ref57f4, "§", "','")
                strRef57F4 = "'" & strRef57F4 & "'"
                strannex = "Select Balance_qty,Ref57f4_No,ref57f4_Date from CustAnnex_HDR "
                strannex = strannex & " WHERE  UNIT_CODE = '" & gstrUNITID & "' and Item_code ='" & parrCustAnnex(0, intloopcount) & "' and Customer_code ='"
                strannex = strannex & strCustCode & "'"
                If blnFIFOFlag = False Then
                    strannex = strannex & " and Ref57f4_No in (" & strRef57F4 & ") "
                End If
                strannex = strannex & " order by ref57f4_Date"
                rsCustAnnex = New ClsResultSetDB_Invoice
                rsCustAnnex.GetResult(strannex)
                intMaxLoop = rsCustAnnex.GetNoRows
                rsCustAnnex.MoveFirst()
                blnValue = True
                For intLoopcount1 = 1 To intMaxLoop
                    If blnValue = True Then
                        strRef57F4 = rsCustAnnex.GetValue("Ref57f4_No")
                        dblbalanceqty = rsCustAnnex.GetValue("Balance_Qty")
                        str57f4Date = rsCustAnnex.GetValue("ref57f4_Date")
                        mstrAnnex = Trim(mstrAnnex) & " Update CustAnnex_HDR "
                        If dblbalanceqty < parrCustAnnex(1, intloopcount) Then
                            mstrAnnex = Trim(mstrAnnex) & " Set Balance_Qty = 0 "
                        Else
                            mstrAnnex = Trim(mstrAnnex) & " Set Balance_Qty = Balance_Qty - " & parrCustAnnex(1, intloopcount)
                        End If
                        mstrAnnex = mstrAnnex & " WHERE  UNIT_CODE = '" & gstrUNITID & "' and Item_code ='" & parrCustAnnex(0, intloopcount) & "' and Customer_code ='"
                        mstrAnnex = mstrAnnex & strCustCode & "' and Ref57f4_No ='" & strRef57F4 & "' "

                        mstrAnnex = mstrAnnex & "Insert into CustAnnex_dtl (Doc_Ty,"
                        mstrAnnex = mstrAnnex & "Invoice_No,Invoice_Date,ref57f4_Date,Ref57f4_No,"
                        mstrAnnex = mstrAnnex & "Item_Code,Quantity,"
                        mstrAnnex = mstrAnnex & "Customer_Code,"
                        mstrAnnex = mstrAnnex & "Location_Code,Product_Code,Ent_Userid,Ent_dt,"
                        mstrAnnex = mstrAnnex & "Upd_Userid,Upd_dt,Unit_Code) values ('O'," & mInvNo & ",GetDate(),'" & str57f4Date & "','"
                        mstrAnnex = mstrAnnex & ref57f4 & "','" & parrCustAnnex(0, intloopcount) & "'," & parrCustAnnex(1, intloopcount) & ","
                        mstrAnnex = mstrAnnex & "'" & strCustCode & "',"
                        mstrAnnex = mstrAnnex & "'" & Trim(txtUnitCode.Text) & "','" & pstrFinishedItem & "','" & mP_User & "',GETDATE(),'"
                        mstrAnnex = mstrAnnex & mP_User & "',GETDATE(),'" & gstrUNITID & "')"
                        If dblbalanceqty < parrCustAnnex(1, intloopcount) Then
                            mP_Connection.Execute(" insert into tempCustAnnex (Ref57f4_No,Annex_No,ref57f4_date,Item_code,Quantity,Balance_qty,finishedItem,UNIT_CODE) values ('" & strRef57F4 & "',0,'" & str57f4Date & "','" & parrCustAnnex(0, intloopcount) & "'," & dblbalanceqty & ",0,'" & pstrFinishedItem & "','" & gstrUNITID & "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        Else
                            mP_Connection.Execute(" insert into tempCustAnnex (Ref57f4_No,Annex_No,ref57f4_date,Item_code,Quantity,Balance_qty,finishedItem,UNIT_CODE) values ('" & strRef57F4 & "',0,'" & str57f4Date & "','" & parrCustAnnex(0, intloopcount) & "'," & parrCustAnnex(1, intloopcount) & "," & dblbalanceqty - parrCustAnnex(1, intloopcount) & ",'" & pstrFinishedItem & "','" & gstrUNITID & "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            blnValue = False
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
ErrHandler:
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Public Function BomCheck() As Boolean
        'input Variable -  Item Code to Found, reqquantity of Finished Product,row in spread
        Dim intChallanMax As Short
        Dim intSpCurrentRow As Short
        Dim intCurrentItem As Short
        Dim VarFinishedItem As Object
        Dim strRef57F4 As String
        Dim strBomMst As String
        Dim strCustAnnexDtl As String
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
        inti = 0
        intAnnexMaxCount = 0
        ReDim arrCustAnnex(3, intAnnexMaxCount)
        strchallan = " select a.Account_code,a.ref_Doc_No,a.Fifo_Flag,b.Item_Code,b.Sales_Quantity from "
        strchallan = strchallan & "salesChallan_dtl a,Sales_dtl b where a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND a.Doc_No = " & Ctlinvoice.Text
        strchallan = strchallan & " and a.Doc_No = b.Doc_no"
        rsSalesChallan = New ClsResultSetDB_Invoice
        rsSalesChallan.GetResult(strchallan, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        intChallanMax = rsSalesChallan.GetNoRows
        rsSalesChallan.MoveFirst()
        Dim intArrCount As Short
        Dim blnItemFoundinArray As Boolean ' to be used to check if item already exist in Array arrItem where we are storing all item we found in Cust annex
        If intChallanMax >= 1 Then
            For intSpCurrentRow = 1 To intChallanMax
                VarFinishedItem = rsSalesChallan.GetValue("Item_Code")
                strCustCode = rsSalesChallan.GetValue("Account_code")
                dblFinishedQty = rsSalesChallan.GetValue("Sales_quantity")
                ref57f4 = rsSalesChallan.GetValue("ref_doc_no")
                strRef57F4 = Replace(ref57f4, "§", "','", 1)
                strRef57F4 = "'" & strRef57F4 & "'"
                blnFIFOFlag = rsSalesChallan.GetValue("FIFO_Flag")

                strBomMst = "Select RawMaterial_Code,Process_type,Required_qty + Waste_qty "
                strBomMst = strBomMst & " As TotalReqQty"
                strBomMst = strBomMst & " from Bom_Mst where  UNIT_CODE = '" & gstrUNITID & "' and Finished_Product_code ='"
                strBomMst = strBomMst & VarFinishedItem & "' Order By Bom_Level"
                rsBomMst = New ClsResultSetDB_Invoice
                rsBomMst.GetResult(strBomMst, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                intBomMaxItem = rsBomMst.GetNoRows
                rsBomMst.MoveFirst()
                If intBomMaxItem > 0 Then ' Item Found in Bom_Mst
                    rsVandorBom = New ClsResultSetDB_Invoice

                    rsVandorBom.GetResult("Select RawMaterial_Code from Vendor_bom where Finish_Product_code = '" & VarFinishedItem & "' and Vendor_code = '" & strCustCode & "' and UNIT_CODE = '" & gstrUNITID & "'")
                    If rsVandorBom.GetNoRows > 0 Then
                        rsVandorBom.ResultSetClose()
                        For intCurrentItem = 1 To intBomMaxItem
                            strBomItem = ""
                            strBomItem = rsBomMst.GetValue("RawMaterial_Code")
                            strCustAnnexDtl = "Select Item_Code,Balance_qty = sum(Balance_qty) from CustAnnex_hdr where Customer_code ='"
                            strCustAnnexDtl = strCustAnnexDtl & Trim(strCustCode) & "' and UNIT_CODE = '" & gstrUNITID & "'"
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
                                rsVandorBom = New ClsResultSetDB_Invoice
                                rsVandorBom.GetResult("Select RawMaterial_Code from Vendor_bom where Finish_Product_code = '" & VarFinishedItem & "'and RawMaterial_code = '" & strBomItem & "' and Vendor_Code = '" & strCustCode & "' and UNIT_CODE = '" & gstrUNITID & "'")
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
                                            If UCase(Trim(arrItem(intArrCount))) = UCase(rsCustAnnexDtl.GetValue("Item_code")) Then
                                                blnItemFoundinArray = True
                                                arrReqQty(intArrCount) = arrReqQty(intArrCount) + (dblTotalReqQty * dblFinishedQty)
                                                If arrQty(intArrCount) < arrReqQty(intArrCount) Then 'in case if sum up is less then Quantity suplied in cust annex
                                                    Msgbox("Customer Supplied Materail for Item " & arrItem(inti) & "is" & arrQty(inti) & ".", MsgBoxStyle.OkOnly, "empower")
                                                    Cmdinvoice.Focus()
                                                    BomCheck = False
                                                    Exit Function
                                                Else
                                                    Call ToGetIteminAcustannex(arrCustAnnex, strBomItem, intAnnexMaxCount, dblTotalReqQty * dblFinishedQty)
                                                End If
                                            End If
                                        Next
                                        If blnItemFoundinArray = False Then
                                            arrItem(inti) = rsCustAnnexDtl.GetValue("Item_code")
                                            arrQty(inti) = rsCustAnnexDtl.GetValue("Balance_qty")
                                            arrReqQty(inti) = dblTotalReqQty * dblFinishedQty
                                            If arrQty(inti) < arrReqQty(inti) Then 'again  check for Quantity requird as compare to supplied in CustAnnex
                                                Msgbox("Customer Supplied Materail for Item " & arrItem(inti) & "is" & arrQty(inti) & ".", MsgBoxStyle.OkOnly, "empower")
                                                Cmdinvoice.Focus()
                                                BomCheck = False
                                                Exit Function
                                            Else
                                                If Len(Trim(arrCustAnnex(0, intAnnexMaxCount))) > 0 Then
                                                    intAnnexMaxCount = intAnnexMaxCount + 1
                                                    ReDim Preserve arrCustAnnex(3, intAnnexMaxCount)
                                                End If
                                                arrCustAnnex(0, intAnnexMaxCount) = rsCustAnnexDtl.GetValue("Item_code")
                                                arrCustAnnex(1, intAnnexMaxCount) = dblTotalReqQty * dblFinishedQty
                                                arrCustAnnex(2, intAnnexMaxCount) = (rsCustAnnexDtl.GetValue("Balance_qty") - (dblTotalReqQty * dblFinishedQty))
                                            End If
                                        End If
                                    Else ' if inti=0 then to add values
                                        arrItem(inti) = rsCustAnnexDtl.GetValue("Item_code")
                                        arrQty(inti) = rsCustAnnexDtl.GetValue("Balance_qty")
                                        arrReqQty(inti) = dblTotalReqQty * dblFinishedQty
                                        If arrQty(inti) < arrReqQty(inti) Then 'Again Same Check
                                            Msgbox("Customer Supplied Material for Item " & arrItem(inti) & "is" & arrQty(inti) & ".", MsgBoxStyle.OkOnly, "empower")
                                            Cmdinvoice.Focus()
                                            BomCheck = False
                                            Exit Function
                                        Else
                                            If Len(Trim(arrCustAnnex(0, intAnnexMaxCount))) > 0 Then
                                                intAnnexMaxCount = intAnnexMaxCount + 1
                                                ReDim Preserve arrCustAnnex(3, intAnnexMaxCount)
                                            End If
                                            arrCustAnnex(0, intAnnexMaxCount) = rsCustAnnexDtl.GetValue("Item_code")
                                            arrCustAnnex(1, intAnnexMaxCount) = dblTotalReqQty * dblFinishedQty
                                            arrCustAnnex(2, intAnnexMaxCount) = (rsCustAnnexDtl.GetValue("Balance_qty") - (dblTotalReqQty * dblFinishedQty))
                                        End If
                                    End If
                                Else
                                    rsVandorBom.ResultSetClose()
                                End If
                            Else ' if Item Not Found in Cust Annex
                                rsVandorBom = New ClsResultSetDB_Invoice
                                rsVandorBom.GetResult("Select RawMaterial_Code from Vendor_bom where Finish_Product_code = '" & VarFinishedItem & "'and RawMaterial_code = '" & strBomItem & "' and Vendor_Code = '" & strCustCode & "' and UNIT_CODE = '" & gstrUNITID & "'")
                                If rsVandorBom.GetNoRows > 0 Then
                                    rsVandorBom.ResultSetClose()
                                    Msgbox("Item " & strBomItem & " is not supplied.", MsgBoxStyle.OkOnly, "empower")
                                    Cmdinvoice.Focus()
                                    BomCheck = False
                                    Exit Function
                                Else ' if it'Process type is not I then Explore it Again in BOM_Mst
                                    rsVandorBom.ResultSetClose()
                                    rsItemMst = New ClsResultSetDB_Invoice
                                    rsItemMst.GetResult("Select Item_Main_grp from Item_Mst (NOLOCK) Where Item_code = '" & strBomItem & "' and UNIT_CODE = '" & gstrUNITID & "'")
                                    If (UCase(rsItemMst.GetValue("Item_Main_grp")) = "R") Or (UCase(rsItemMst.GetValue("Item_Main_grp")) = "C") Then
                                        rsItemMst.ResultSetClose()
                                        BomCheck = True
                                    Else
                                        rsItemMst.ResultSetClose()
                                        dblFinishedQty = dblFinishedQty * rsBomMst.GetValue("TotalReqQty")
                                        If ExploreBom(strBomItem, dblFinishedQty, intSpCurrentRow, strCustCode, ref57f4, intAnnexMaxCount, CStr(VarFinishedItem)) = False Then
                                            BomCheck = False
                                            Exit Function
                                        End If
                                    End If
                                End If
                            End If
                            rsCustAnnexDtl.ResultSetClose()
                            rsBomMst.MoveNext()
                            inti = inti + 1
                        Next
                        rsSalesChallan.MoveNext()
                    Else
                        Msgbox("No BOM Defind for the Item (" & VarFinishedItem & ") defined in challan", MsgBoxStyle.Information, "empower")
                        BomCheck = False
                        rsVandorBom.ResultSetClose()
                        Exit Function
                    End If
                Else
                    Msgbox("No Customer BOM Defind for Item (" & VarFinishedItem & ") defined in challan", MsgBoxStyle.Information, "empower")
                    BomCheck = False
                    rsBomMst.ResultSetClose()
                    Exit Function
                End If
                rsBomMst.ResultSetClose()
                Call InsertUpdateAnnex(arrCustAnnex, VarFinishedItem, intAnnexMaxCount)
                inti = 0
                intAnnexMaxCount = 0
                ReDim arrCustAnnex(3, intAnnexMaxCount)
                ReDim arrItem(inti)
                ReDim arrQty(inti)
                ReDim arrReqQty(inti)
            Next
        End If
        rsSalesChallan.ResultSetClose()
        BomCheck = True
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
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
        Dim strCustAnnexDtl As String
        Dim strref As String
        On Error GoTo ErrHandler
        strBomMstRaw = "Select RawMaterial_Code,Required_qty + Waste_qty "
        strBomMstRaw = strBomMstRaw & " As TotalReqQty,Process_Type from Bom_Mst where "
        strBomMstRaw = strBomMstRaw & " item_Code ='" & strBomItem
        strBomMstRaw = strBomMstRaw & "'and finished_product_code ='"
        strBomMstRaw = strBomMstRaw & pstrItemCode & "' and UNIT_CODE = '" & gstrUNITID & "'"
        rsBomMstRaw = New ClsResultSetDB_Invoice
        rsBomMstRaw.GetResult(strBomMstRaw, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        Dim intArrCount As Short
        Dim blnArrItemFound As Boolean
        If rsBomMstRaw.GetNoRows > 0 Then ' If Item Found in Bom Mst
            intBomMaxRaw = rsBomMstRaw.GetNoRows
            rsBomMstRaw.MoveFirst()
            For intCurrentRaw = 1 To intBomMaxRaw
                strBomItem = rsBomMstRaw.GetValue("RawMaterial_code")
                dblTotalReqQty = rsBomMstRaw.GetValue("TotalReqQty")

                strCustAnnexDtl = "Select Item_Code,Balance_qty,REF57F4_DATE from CustAnnex_hdr where Customer_code ='"
                strCustAnnexDtl = strCustAnnexDtl & Trim(pstrCustCode) & "' and UNIT_CODE = '" & gstrUNITID & "'"
                If blnFIFOFlag = False Then
                    strref = Replace(pstrRef, "§", "','", 1)
                    strref = "'" & strref & "'"
                    strCustAnnexDtl = strCustAnnexDtl & " and ref57f4_no IN ("
                    strCustAnnexDtl = strCustAnnexDtl & Trim(strref) & ")"
                End If
                strCustAnnexDtl = strCustAnnexDtl & "  and getdate() <= "
                strCustAnnexDtl = strCustAnnexDtl & " DateAdd(d, 180, ref57f4_date)"
                strCustAnnexDtl = strCustAnnexDtl & " and Item_code ='" & strBomItem & "'"
                rsCustAnnexDtl = New ClsResultSetDB_Invoice
                rsCustAnnexDtl.GetResult(strCustAnnexDtl, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                If rsCustAnnexDtl.GetNoRows >= 1 Then 'if item Found in CustAnnex then replace that item from Parant string
                    rsVandorBom = New ClsResultSetDB_Invoice

                    rsVandorBom.GetResult("Select RawMaterial_Code from Vendor_bom where Finish_Product_code = '" & pstrFinishedProduct & "'and RawMaterial_code = '" & strBomItem & "' and Vendor_code = '" & pstrCustCode & "' and UNIT_CODE = '" & gstrUNITID & "'")
                    If rsVandorBom.GetNoRows > 0 Then
                        rsVandorBom.ResultSetClose()
                        rsCustAnnexDtl.MoveFirst()
                        inti = inti + 1
                        ReDim Preserve arrItem(inti)
                        ReDim Preserve arrQty(inti)
                        ReDim Preserve arrReqQty(inti)
                        blnArrItemFound = False
                        For intArrCount = 0 To UBound(arrItem) - 1 'to check if ITem Already there in ArrItem Array
                            If UCase(Trim(arrItem(intArrCount))) = UCase(Trim(rsCustAnnexDtl.GetValue("Item_code"))) Then
                                blnArrItemFound = True
                                arrReqQty(intArrCount) = arrReqQty(intArrCount) + (dblTotalReqQty * pstrFinishedQty)
                                If arrQty(intArrCount) < arrReqQty(intArrCount) Then ' to Check with Quantity supplieded in Cust Annex
                                    Msgbox("Customer Supplied Materail for Item " & arrItem(inti) & " is " & arrQty(inti) & " .", MsgBoxStyle.OkOnly, "empower")
                                    Cmdinvoice.Focus()
                                    ExploreBom = False
                                    Exit Function
                                Else
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
                            arrItem(inti) = rsCustAnnexDtl.GetValue("Item_code")
                            arrQty(inti) = rsCustAnnexDtl.GetValue("Balance_qty")
                            arrReqQty(inti) = dblTotalReqQty * pstrFinishedQty
                            If arrQty(inti) < arrReqQty(inti) Then
                                Msgbox("Customer Supplied Materail for Item " & arrItem(inti) & " is " & arrQty(inti) & " .", MsgBoxStyle.OkOnly, "empower")
                                Cmdinvoice.Focus()
                                ExploreBom = False
                                Exit Function
                            Else
                                If Len(Trim(arrCustAnnex(0, pintAnnexMaxCount))) > 0 Then
                                    pintAnnexMaxCount = pintAnnexMaxCount + 1
                                    ReDim Preserve arrCustAnnex(3, pintAnnexMaxCount)
                                End If
                                arrCustAnnex(0, pintAnnexMaxCount) = rsCustAnnexDtl.GetValue("Item_code")
                                arrCustAnnex(1, pintAnnexMaxCount) = (dblTotalReqQty * pstrFinishedQty)
                                arrCustAnnex(2, pintAnnexMaxCount) = (rsCustAnnexDtl.GetValue("Balance_qty") - (dblTotalReqQty * pstrFinishedQty))
                                ExploreBom = True
                            End If
                        Else
                        End If
                    Else
                        rsVandorBom.ResultSetClose()
                    End If
                Else
                    rsCustAnnexDtl.ResultSetClose()
                    '            If strProcessType = "I" Then
                    rsVandorBom = New ClsResultSetDB_Invoice
                    rsVandorBom.GetResult("Select RawMaterial_Code from Vendor_bom where Finish_Product_code = '" & pstrItemCode & "'and RawMaterial_code = '" & strBomItem & "' and Vendor_code ='" & pstrCustCode & "' and UNIT_CODE = '" & gstrUNITID & "'")
                    If rsVandorBom.GetNoRows > 0 Then
                        rsVandorBom.ResultSetClose()
                        Msgbox("Item " & strBomItem & " is not supplied.", MsgBoxStyle.OkOnly, "empower")
                        Cmdinvoice.Focus()
                        ExploreBom = False
                        Exit Function
                    Else 'if not of Process type I then again Explore
                        rsVandorBom.ResultSetClose()
                        rsItemMst = New ClsResultSetDB_Invoice
                        rsItemMst.GetResult("Select Item_Main_grp from Item_Mst (NOLOCK) Where Item_code = '" & strBomItem & "' and UNIT_CODE = '" & gstrUNITID & "'")
                        If (UCase(rsItemMst.GetValue("Item_Main_grp")) = "R") Or (UCase(rsItemMst.GetValue("Item_Main_grp")) = "C") Then
                            ExploreBom = True
                        Else
                            pstrFinishedQty = pstrFinishedQty * dblTotalReqQty
                            Call ExploreBom(strBomItem, pstrFinishedQty, pstrSPCurrentRow, pstrCustCode, pstrRef, pintAnnexMaxCount, pstrFinishedProduct)
                        End If
                        rsItemMst.ResultSetClose()
                    End If
                End If
                rsCustAnnexDtl.ResultSetClose()
                rsBomMstRaw.MoveNext()
            Next
        Else
            rsBomMstRaw.ResultSetClose()
            Msgbox("No BOM Defind for Item (" & strBomItem & ") defined in challan", MsgBoxStyle.Information, "empower")
            ExploreBom = False
            Exit Function
        End If
        rsBomMstRaw.ResultSetClose()
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Public Function ToGetIteminAcustannex(ByRef pvarArray(,) As Object, ByRef pstrItemCode As Object, ByRef pintArrMaxCount As Short, ByRef pdblReqQuantity As Double) As Object
        Dim intLoopCounter As Short
        On Error GoTo ErrHandler
        For intLoopCounter = 0 To pintArrMaxCount
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
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
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
                Cmdinvoice.Focus()
                dtpRemoval.Enabled = True
                dtpRemoval.Focus()
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
        rsSalesDtl = New ClsResultSetDB_Invoice
        rsSalesDtl.GetResult("Select Item_Code,Sales_Quantity from Sales_dtl where doc_No =" & Ctlinvoice.Text & " and Location_Code='" & Trim(txtUnitCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "'")
        intMaxLoop = rsSalesDtl.GetNoRows : rsSalesDtl.MoveFirst()
        CheckDataFromGrin = False
        For intLoopCounter = 1 To intMaxLoop
            StrItemCode = rsSalesDtl.GetValue("Item_code")
            dblItemQty = rsSalesDtl.GetValue("Sales_quantity")

            strSql = "select a.Doc_No,a.Item_code,a.Rejected_Quantity,"
            strSql = strSql & "Despatch_Quantity = isnull(a.Despatch_Quantity,0),"
            strSql = strSql & " Inspected_Quantity = isnull(Inspected_Quantity,0),"
            strSql = strSql & "RGP_Quantity = isnull(RGP_Quantity,0) from grn_Dtl a,grn_hdr b Where a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND "
            strSql = strSql & "a.Doc_type = b.Doc_type And a.Doc_No = b.Doc_No and "
            strSql = strSql & "a.From_Location = b.From_Location and a.From_Location ='01R1'"
            strSql = strSql & "and a.Rejected_quantity > 0 and b.Vendor_code = '" & pstrCustCode
            strSql = strSql & "' and a.Doc_No = " & pdblDocNo & " and a.Item_code = '" & StrItemCode & "'"
            rsGrnDtl = New ClsResultSetDB_Invoice
            rsGrnDtl.GetResult(strSql)
            dblRejQty = rsGrnDtl.GetValue("Rejected_Quantity") - rsGrnDtl.GetValue("Despatch_Quantity") - rsGrnDtl.GetValue("Inspected_Quantity") - rsGrnDtl.GetValue("RGP_Quantity")
            If rsGrnDtl.GetNoRows > 0 Then
                If dblItemQty > (dblRejQty) Then
                    Msgbox("Max. Quantity Allowed For Item " & StrItemCode & " is " & dblRejQty & ", Quantity Entered in Invoice is : " & dblItemQty)
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
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Public Function UpdateGrnHdr(ByRef pdblGrinNo As Double, ByRef pdblinvoiceNo As Double) As Boolean
        Dim rsSalesDtl As ClsResultSetDB_Invoice
        Dim intMaxLoop As Short
        Dim StrItemCode As String
        Dim dblQty As Double
        Dim intloopcount As Short
        rsSalesDtl = New ClsResultSetDB_Invoice
        On Error GoTo ErrHandler
        rsSalesDtl.GetResult("select * from sales_dtl where Doc_No = " & Ctlinvoice.Text & " and Location_Code='" & Trim(txtUnitCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "'")
        If rsSalesDtl.GetNoRows > 0 Then
            intMaxLoop = rsSalesDtl.GetNoRows
            rsSalesDtl.MoveFirst()
            strupdateGrinhdr = ""
            For intloopcount = 1 To intMaxLoop
                StrItemCode = rsSalesDtl.GetValue("ITem_code")
                dblQty = rsSalesDtl.GetValue("Sales_Quantity")
                If Len(Trim(strupdateGrinhdr)) = 0 Then
                    strupdateGrinhdr = "Update Grn_Dtl Set Despatch_Quantity = isnull(Despatch_Quantity,0) +" & dblQty
                    strupdateGrinhdr = strupdateGrinhdr & " Where  UNIT_CODE = '" & gstrUNITID & "' and ITem_Code = '" & StrItemCode & "' and Doc_No = " & pdblGrinNo
                Else
                    strupdateGrinhdr = strupdateGrinhdr & vbCrLf & "Update Grn_Dtl Set Despatch_Quantity = isnull(Despatch_Quantity,0) + " & dblQty
                    strupdateGrinhdr = strupdateGrinhdr & " Where ITem_Code = '" & StrItemCode & "' and UNIT_CODE = '" & gstrUNITID & "' and Doc_No = " & pdblGrinNo
                End If
                rsSalesDtl.MoveNext()
            Next
        Else
            Msgbox("No Items Available in Invoice " & Ctlinvoice.Text)
        End If
        rsSalesDtl.ResultSetClose()
        UpdateGrnHdr = True
        Exit Function
ErrHandler:
        SqlConnectionclass.ExecuteNonQuery("INSERT INTO INV_ERR_PRIMARYKEY(ERR_NUMBER,ERR_DESCRIPTION,TMPINVOICENO,INVOICE_NO,UNIT_CODE,FUNCTIONNAME ) VALUES('" & Err.Number & "','" & Replace(Err.Description, "'", "") & "','" & Ctlinvoice.Text & "','" & mInvNo & "','" & gstrUNITID & "','UpdateGrnHdr')")
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        rsSalesDtl.ResultSetClose()
        UpdateGrnHdr = False
    End Function
    Private Function GetTaxGlSl(ByVal TaxType As String) As String
        Dim objRecordSet As New ADODB.Recordset
        On Error GoTo ErrHandler
        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close()
        ' Added by priti on 13 Mar 2024 to add rejection tax type APRC 
        If UCase(Trim(cmbInvType.Text)) = "REJECTION" And DataExist("SELECT  TOP 1 1 FROM MKT_INVREJ_DTL WHERE  invoice_no=" & Ctlinvoice.Text & "  and rej_Type=2 and unit_code='" & gstrUNITID & "'") = True Then
            Dim strTaxType = SqlConnectionclass.ExecuteScalar("select isnull(RejectionTaxType,'') from sales_parameter where UNIT_CODE = '" & gstrUNITID & "'")
            If Len(strTaxType) = 0 Then
                objRecordSet.Open("SELECT tx_glCode, tx_slCode FROM fin_TaxGlRel WHERE UNIT_CODE='" + gstrUNITID + "' AND  tx_rowType = 'ARTAX' AND tx_taxId ='" & TaxType & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
            Else
                objRecordSet.Open("SELECT tx_glCode, tx_slCode FROM fin_TaxGlRel WHERE UNIT_CODE='" + gstrUNITID + "' AND  tx_rowType = '" & strTaxType & "' AND tx_taxId ='" & TaxType & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
            End If
        Else
            objRecordSet.Open("SELECT tx_glCode, tx_slCode FROM fin_TaxGlRel WHERE UNIT_CODE='" + gstrUNITID + "' AND  tx_rowType = 'ARTAX' AND tx_taxId ='" & TaxType & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
        End If
        'objRecordSet.Open("SELECT tx_glCode, tx_slCode FROM fin_TaxGlRel where tx_rowType = 'ARTAX' AND tx_taxId ='" & TaxType & "' and UNIT_CODE = '" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
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
    Private Function GetTaxGlSl_IGST_RECOVERY(ByVal TaxType As String) As String
        Dim objRecordSet As New ADODB.Recordset
        On Error GoTo ErrHandler
        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close()
        objRecordSet.Open("SELECT tx_glCode, tx_slCode FROM fin_TaxGlRel WHERE UNIT_CODE='" + gstrUNITID + "' AND  tx_rowType = 'IGSTR' AND tx_taxId ='" & TaxType & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
        If objRecordSet.EOF Then
            GetTaxGlSl_IGST_RECOVERY = "N"
            If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                objRecordSet.Close()
                objRecordSet = Nothing
            End If
            Exit Function
        End If
        GetTaxGlSl_IGST_RECOVERY = Trim(IIf(IsDBNull(objRecordSet.Fields("tx_glCode").Value), "", objRecordSet.Fields("tx_glCode").Value)) & "»" & Trim(IIf(IsDBNull(objRecordSet.Fields("tx_slCode").Value), "", objRecordSet.Fields("tx_slCode").Value))
        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
            objRecordSet.Close()
            objRecordSet = Nothing
        End If
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        GetTaxGlSl_IGST_RECOVERY = "N"
        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
            objRecordSet.Close()
            objRecordSet = Nothing
        End If
    End Function

    Private Function CreateStringForAccounts() As Boolean
        Dim objRecordSet As New ADODB.Recordset
        Dim objTmpRecordset As New ADODB.Recordset
        Dim rstoolglslrecordset As New ADODB.Recordset
        Dim strRetval As String
        Dim strInvoiceNo As String
        Dim strInvoiceDate As String
        Dim strCurrencyCode As String
        Dim dblInvoiceAmt As Double
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
        Dim rsObjItemcode As New ADODB.Recordset
        Dim strToolGLCr As String
        Dim strToolSLCr As String
        Dim strToolGLDr As String
        Dim strToolSLDr As String
        Dim dblToolCost As Double
        Dim TOOL_AMOR_FLAG As Boolean
        Dim RSTOOLPOST As ADODB.Recordset
        Dim RSTOOLGET As ADODB.Recordset
        Dim STRCSIEX_GL As String
        Dim STRCSIEX_SL As String
        Dim DBLCSIEDCESS As Double
        Dim strTaxCCCode As String
        Dim dblTCStaxAmt As Double
        Dim dblInvoiceAmtRoundOff_diff As Double

        On Error GoTo ErrHandler
        objRecordSet.Open("SELECT * FROM  saleschallan_dtl WHERE Doc_No='" & Trim(Ctlinvoice.Text) & "' and Location_Code='" & Trim(txtUnitCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
        If objRecordSet.EOF Then
            Msgbox("Invoice details not found", MsgBoxStyle.Information, "empower")
            CreateStringForAccounts = False
            If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                objRecordSet.Close()
                objRecordSet = Nothing
            End If
            Exit Function
        End If
        strInvoiceNo = CStr(mInvNo)
        'strInvoiceDate = VB6.Format(objRecordSet.Fields("Invoice_Date").Value, "dd-MMM-yyyy")
        If DataExist("SELECT TOP 1 1 FROM SALES_PARAMETER WHERE INVOICE_LOCKING_ENTRY_SAMEDATE=1  and UNIT_CODE = '" & gstrUNITID & "'") Then
            strInvoiceDate = VB6.Format(GetServerDateTime(), "dd-MMM-yyyy")
        Else
            strInvoiceDate = VB6.Format(objRecordSet.Fields("Invoice_Date").Value, "dd-MMM-yyyy")
        End If

        strCurrencyCode = Trim(IIf(IsDBNull(objRecordSet.Fields("Currency_Code").Value), "", objRecordSet.Fields("Currency_Code").Value))
        dblInvoiceAmt = IIf(IsDBNull(objRecordSet.Fields("total_amount").Value), 0, objRecordSet.Fields("total_amount").Value)
        dblInvoiceAmtRoundOff_diff = IIf(IsDBNull(objRecordSet.Fields("TotalInvoiceAmtRoundOff_diff").Value), 0, objRecordSet.Fields("TotalInvoiceAmtRoundOff_diff").Value)
        dblExchangeRate = IIf(IsDBNull(objRecordSet.Fields("Exchange_Rate").Value), 1, objRecordSet.Fields("Exchange_Rate").Value)
        strCustCode = Trim(objRecordSet.Fields("Account_Code").Value)
        dblTCStaxAmt = IIf(IsDBNull(objRecordSet.Fields("TCSTaxAmount").Value), 1, objRecordSet.Fields("TCSTaxAmount").Value)
        strCreditTermsID = Trim(IIf(IsDBNull(objRecordSet.Fields("payment_terms").Value), "", objRecordSet.Fields("payment_terms").Value))
        mblnSEZ_Knockingoff_req = DataExist("SELECT TOP 1 1 FROM CUSTOMER_MST (NOLOCK) WHERE SEZ_CUSTOMER = 1 AND CUSTOMER_CODE='" & strCustCode & "' and Unit_code='" & gstrUNITID & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))")

        If UCase(lbldescription.Text) <> "SMP" Then 'if invoice type is not sample sales then
            If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
            If UCase(Trim(cmbInvType.Text)) = "REJECTION" And CUST_REJECTION_FLAG = False Then
                If Len(Trim(strCustRef)) <> 0 Then

                    objTmpRecordset.Open("SELECT ISNULL(SUM(Basic_Amount),0) AS Basic_Amt FROM sales_dtl WHERE  UNIT_CODE = '" & gstrUNITID & "' and doc_no =" & Ctlinvoice.Text, mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
                    If Not objTmpRecordset.EOF Then
                        dblBasicAmount = Val(objTmpRecordset.Fields("Basic_Amt").Value)
                    End If
                    If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
                    dblInvoiceAmt = System.Math.Round(dblInvoiceAmt - dblBasicAmount, 4)
                    dblBasicAmount = 0
                    objTmpRecordset.Open("SELECT GL_AccountID, Ven_slCode, CrTrm_Termid FROM Pur_VendorMaster where Prty_PartyID='" & strCustCode & "' and UNIT_CODE = '" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
                Else
                    objTmpRecordset.Open("SELECT GL_AccountID, Ven_slCode, CrTrm_Termid FROM Pur_VendorMaster where Prty_PartyID='" & strCustCode & "' and UNIT_CODE = '" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
                End If
            ElseIf UCase(Trim(cmbInvType.Text)) = "REJECTION" And CUST_REJECTION_FLAG = True Then
                objTmpRecordset.Open("SELECT ISNULL(SUM(Basic_Amount),0) AS Basic_Amt FROM sales_dtl WHERE  UNIT_CODE = '" & gstrUNITID & "' and doc_no =" & Ctlinvoice.Text, mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
                If Not objTmpRecordset.EOF Then
                    dblBasicAmount = Val(objTmpRecordset.Fields("Basic_Amt").Value)
                End If
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
                dblInvoiceAmt = dblInvoiceAmt - dblBasicAmount
                dblBasicAmount = 0
                objTmpRecordset.Open("SELECT Cst_ArCode, Cst_slCode, Cst_CreditTerm FROM Sal_CustomerMaster where Prty_PartyID='" & strCustCode & "' and UNIT_CODE = '" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
            Else
                objTmpRecordset.Open("SELECT Cst_ArCode, Cst_slCode, Cst_CreditTerm FROM Sal_CustomerMaster where Prty_PartyID='" & strCustCode & "' and UNIT_CODE = '" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
            End If
            If objTmpRecordset.EOF Then
                If UCase(Trim(cmbInvType.Text)) = "REJECTION" And CUST_REJECTION_FLAG = False Then
                    Msgbox("Vendor details not found", MsgBoxStyle.Information, "empower")
                Else
                    Msgbox("Customer details not found", MsgBoxStyle.Information, "empower")
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
            If UCase(Trim(cmbInvType.Text)) = "REJECTION" And CUST_REJECTION_FLAG = False Then
                strCustomerGL = Trim(IIf(IsDBNull(objTmpRecordset.Fields("GL_AccountID").Value), "", objTmpRecordset.Fields("GL_AccountID").Value))
                strCustomerSL = Trim(IIf(IsDBNull(objTmpRecordset.Fields("Ven_slCode").Value), "", objTmpRecordset.Fields("Ven_slCode").Value))
                If strCreditTermsID.Length = 0 Then
                    strCreditTermsID = Trim(IIf(IsDBNull(objTmpRecordset.Fields("CrTrm_Termid").Value), "", objTmpRecordset.Fields("CrTrm_Termid").Value))
                End If
            Else
                strCustomerGL = Trim(IIf(IsDBNull(objTmpRecordset.Fields("Cst_ArCode").Value), "", objTmpRecordset.Fields("Cst_ArCode").Value))
                strCustomerSL = Trim(IIf(IsDBNull(objTmpRecordset.Fields("Cst_slCode").Value), "", objTmpRecordset.Fields("Cst_slCode").Value))
                If strCreditTermsID.Length = 0 Then
                    strCreditTermsID = Trim(IIf(IsDBNull(objTmpRecordset.Fields("Cst_CreditTerm").Value), "", objTmpRecordset.Fields("Cst_CreditTerm").Value))
                End If
            End If
            If strCreditTermsID = "" Then
                Msgbox("Credit Terms not found", MsgBoxStyle.Information, "empower")
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
            strRetval = objCreditTerms.RetCR_Term_Dates("", "INV", strCreditTermsID, strInvoiceDate, gstrUNITID, "", "", gstrCONNECTIONSTRING)
            objCreditTerms = Nothing
            If CheckString(strRetval) = "Y" Then
                strRetval = Mid(strRetval, 3)
                varTmp = Split(strRetval, "»")
                strBasicDueDate = VB6.Format(varTmp(0), "dd-MMM-yyyy")
                strPaymentDueDate = VB6.Format(varTmp(1), "dd-MMM-yyyy")
                strExpectedDueDate = VB6.Format(varTmp(1), "dd-MMM-yyyy")
            Else
                Msgbox(CheckString(strRetval), MsgBoxStyle.Information, "empower")
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
            strRetval = GetItemGLSL("", "Sample_Expences")
            If strRetval = "N" Then
                CreateStringForAccounts = False
                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSet.Close()
                    objRecordSet = Nothing
                End If
                Exit Function
            End If
            varTmp = Split(strRetval, "»")
            strCustomerGL = varTmp(0)
            strCustomerSL = varTmp(1)
        End If
        mstrMasterString = ""
        mstrDetailString = ""
        ''24 aug 2016'
        Dim rsSalesParameter As New ADODB.Recordset

        If rsSalesParameter.State = ADODB.ObjectStateEnum.adStateOpen Then rsSalesParameter.Close()
        rsSalesParameter.Open("SELECT TotalInvoiceAmount_RoundOff, TotalInvoiceAmountRoundOff_Decimal FROM SALES_PARAMETER WHERE UNIT_CODE='" + gstrUNITID + "'", mP_Connection)
        If Not rsSalesParameter.EOF Then
            mblnTotalInvoiceAmountRoundOff = rsSalesParameter.Fields("TotalInvoiceAmount_RoundOff").Value
            mintTotalInvoiceAmountRoundOff = rsSalesParameter.Fields("TotalInvoiceAmountRoundOff_Decimal").Value
        End If
        If rsSalesParameter.State = ADODB.ObjectStateEnum.adStateOpen Then
            rsSalesParameter.Close()
            rsSalesParameter = Nothing
        End If
        '24 aug 2016'
        If UCase(Trim(cmbInvType.Text)) = "REJECTION" And CUST_REJECTION_FLAG = False Then
            '24 aug 2016'
            If mblnTotalInvoiceAmountRoundOff Then
                'mstrMasterString = "M»»" & VB6.Format(GetServerDate(), "dd-MMM-yyyy") & "»0»»" & gstrUNITID & "»" & Trim(strCustCode) & "»" & strInvoiceNo & "»" & strInvoiceDate & "»" & strBasicDueDate & "»" & strPaymentDueDate & "»" & strExpectedDueDate & "»" & strCurrencyCode & "»" & dblExchangeRate & "»" & System.Math.Round(dblInvoiceAmt) & "»0»»»Rej. Inv. " & strInvoiceNo & "»" & strCustomerGL & "»" & strCustomerSL & "»DR»" & strCustomerGL & "»" & strCustomerSL & "»»" & gstrCURRENCYCODE & "»" & mP_User & "»getdate()»0»AP»»»»0»»¦"
                mstrMasterString = "M»»" & VB6.Format(GetServerDate(), "dd-MMM-yyyy") & "»0»»" & gstrUNITID & "»" & Trim(strCustCode) & "»" & strInvoiceNo & "»" & strInvoiceDate & "»" & strInvoiceDate & "»" & strInvoiceDate & "»" & strInvoiceDate & "»" & strCurrencyCode & "»" & dblExchangeRate & "»" & System.Math.Round(dblInvoiceAmt) & "»0»»»Rej. Inv. " & strInvoiceNo & "»" & strCustomerGL & "»" & strCustomerSL & "»DR»" & strCustomerGL & "»" & strCustomerSL & "»»" & gstrCURRENCYCODE & "»" & mP_User & "»getdate()»0»AP»»»»0»»¦"
            Else
                'mstrMasterString = "M»»" & VB6.Format(GetServerDate(), "dd-MMM-yyyy") & "»0»»" & gstrUNITID & "»" & Trim(strCustCode) & "»" & strInvoiceNo & "»" & strInvoiceDate & "»" & strBasicDueDate & "»" & strPaymentDueDate & "»" & strExpectedDueDate & "»" & strCurrencyCode & "»" & dblExchangeRate & "»" & System.Math.Round(dblInvoiceAmt, mintTotalInvoiceAmountRoundOff) & "»0»»»Rej. Inv. " & strInvoiceNo & "»" & strCustomerGL & "»" & strCustomerSL & "»DR»" & strCustomerGL & "»" & strCustomerSL & "»»" & gstrCURRENCYCODE & "»" & mP_User & "»getdate()»0»AP»»»»0»»¦"
                mstrMasterString = "M»»" & VB6.Format(GetServerDate(), "dd-MMM-yyyy") & "»0»»" & gstrUNITID & "»" & Trim(strCustCode) & "»" & strInvoiceNo & "»" & strInvoiceDate & "»" & strInvoiceDate & "»" & strInvoiceDate & "»" & strInvoiceDate & "»" & strCurrencyCode & "»" & dblExchangeRate & "»" & System.Math.Round(dblInvoiceAmt, mintTotalInvoiceAmountRoundOff) & "»0»»»Rej. Inv. " & strInvoiceNo & "»" & strCustomerGL & "»" & strCustomerSL & "»DR»" & strCustomerGL & "»" & strCustomerSL & "»»" & gstrCURRENCYCODE & "»" & mP_User & "»getdate()»0»AP»»»»0»»¦"
            End If


            '24 aug 2016'
        Else
            mstrMasterString = "I»" & strInvoiceNo & "»Dr»»" & strInvoiceDate & "»»»»»SAL»I»" & strInvoiceNo & "»" & strInvoiceDate & "»"
            If UCase(lbldescription.Text) <> "SMP" Then
                mstrMasterString = mstrMasterString & Trim(strCustCode) & "»" & gstrUNITID & "»" & strCurrencyCode & "»"
            Else
                mstrMasterString = mstrMasterString & "»" & gstrUNITID & "»" & strCurrencyCode & "»"
            End If
            '24 aug 2016 
            If mblnTotalInvoiceAmountRoundOff Then
                mstrMasterString = mstrMasterString & System.Math.Round(dblInvoiceAmt, 0) & "»" & System.Math.Round(dblInvoiceAmt * dblExchangeRate, 0) & "»" & dblExchangeRate & "»" & strCreditTermsID & "»" & strBasicDueDate & "»" & strPaymentDueDate & "»" & strExpectedDueDate & "»" & strCustomerGL & "»" & strCustomerSL & "»" & mP_User & "»getdate()»»"
            Else
                mstrMasterString = mstrMasterString & System.Math.Round(dblInvoiceAmt, mintTotalInvoiceAmountRoundOff) & "»" & System.Math.Round(dblInvoiceAmt * dblExchangeRate, mintTotalInvoiceAmountRoundOff) & "»" & dblExchangeRate & "»" & strCreditTermsID & "»" & strBasicDueDate & "»" & strPaymentDueDate & "»" & strExpectedDueDate & "»" & strCustomerGL & "»" & strCustomerSL & "»" & mP_User & "»getdate()»»"
            End If
            'mstrMasterString = mstrMasterString & System.Math.Round(dblInvoiceAmt, 0) & "»" & System.Math.Round(System.Math.Round(dblInvoiceAmt, 0) * dblExchangeRate, 2) & "»" & dblExchangeRate & "»" & strCreditTermsID & "»" & strBasicDueDate & "»" & strPaymentDueDate & "»" & strExpectedDueDate & "»" & strCustomerGL & "»" & strCustomerSL & "»" & mP_User & "»getdate()»»"
            '24 aug 2016

        End If
        iCtr = 1

        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
        If Trim(IIf(IsDBNull(objRecordSet.Fields("SalesTax_Type").Value), "", objRecordSet.Fields("SalesTax_Type").Value)) <> "" Then
            objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE  UNIT_CODE = '" & gstrUNITID & "' and TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("SalesTax_Type").Value), "", objRecordSet.Fields("SalesTax_Type").Value)) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
            If Not objTmpRecordset.EOF Then
                strTaxType = Trim(UCase(objTmpRecordset.Fields("Tx_TaxeID").Value))
            Else
                Msgbox("Tax type not found", MsgBoxStyle.Information, "empower")
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
            If strTaxType = "LST" Or strTaxType = "CST" Or strTaxType = "VAT" Then
                dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("Sales_Tax_Amount").Value), 0, objRecordSet.Fields("Sales_Tax_Amount").Value)
                dblBaseCurrencyAmount = dblTaxAmt
                dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("SalesTax_Per").Value), 0, objRecordSet.Fields("SalesTax_Per").Value)
                If dblBaseCurrencyAmount > 0 Then
                    'initializing the tax gl and sl here
                    strRetval = GetTaxGlSl(strTaxType)
                    If strRetval = "N" Then
                        Msgbox("GL for ARTAX is not defined for " & strTaxType, MsgBoxStyle.Information, "empower")
                        CreateStringForAccounts = False
                        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                            objRecordSet.Close()
                            objRecordSet = Nothing
                        End If
                        Exit Function
                    End If
                    varTmp = Split(strRetval, "»")
                    strTaxGL = varTmp(0)
                    strTaxSL = varTmp(1)
                    If UCase(Trim(cmbInvType.Text)) = "REJECTION" And CUST_REJECTION_FLAG = False Then
                        mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & dblTaxAmt & "»»CST/LST for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                    Else
                        mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»" & strTaxType & "»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                    End If
                    iCtr = iCtr + 1
                End If
            End If
        End If
        'service invoice
        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
        If Trim(IIf(IsDBNull(objRecordSet.Fields("ServiceTax_Type").Value), "", objRecordSet.Fields("ServiceTax_Type").Value)) <> "" Then
            objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("ServiceTax_Type").Value), "", objRecordSet.Fields("ServiceTax_Type").Value)) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
            If Not objTmpRecordset.EOF Then
                strTaxType = Trim(UCase(objTmpRecordset.Fields("Tx_TaxeID").Value))
            Else
                Msgbox("SRT Tax type not found", MsgBoxStyle.Information, "eMPro")
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
                    strRetval = GetTaxGlSl(strTaxType)
                    If strRetval = "N" Then
                        Msgbox("GL for ARTAX is not defined for " & strTaxType, MsgBoxStyle.Information, "eMPro")
                        CreateStringForAccounts = False
                        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                            objRecordSet.Close()
                            objRecordSet = Nothing
                        End If
                        Exit Function
                    End If
                    varTmp = Split(strRetval, "»")
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

        'service invoice
        '10706455  
        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
        If Trim(IIf(IsDBNull(objRecordSet.Fields("ADDVAT_Type").Value), "", objRecordSet.Fields("ADDVAT_Type").Value)) <> "" Then
            objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("ADDVAT_Type").Value), "", objRecordSet.Fields("ADDVAT_Type").Value)) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
            If Not objTmpRecordset.EOF Then
                strTaxType = Trim(UCase(objTmpRecordset.Fields("Tx_TaxeID").Value))
            Else
                Msgbox("Tax type not found", MsgBoxStyle.Information, "eMPro")
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
                    strRetval = GetTaxGlSl(strTaxType)
                    If strRetval = "N" Then
                        Msgbox("GL for ARTAX is not defined for " & strTaxType, MsgBoxStyle.Information, "eMPro")
                        CreateStringForAccounts = False
                        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                            objRecordSet.Close()
                            objRecordSet = Nothing
                        End If
                        Exit Function
                    End If
                    varTmp = Split(strRetval, "»")
                    strTaxGL = varTmp(0)
                    strTaxSL = varTmp(1)
                    If UCase(Trim(cmbInvType.Text)) <> "REJECTION" Then
                        mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»" & strTaxType & "»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                    Else
                        mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & dblTaxAmt & "»»ADDVAT for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                    End If
                    iCtr = iCtr + 1
                End If
            End If
        End If

        '10706455  
        'SBC TAX
        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
        If Trim(IIf(IsDBNull(objRecordSet.Fields("SBCTAX_TYPE").Value), "", objRecordSet.Fields("SBCTAX_TYPE").Value)) <> "" Then
            objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("SBCTAX_TYPE").Value), "", objRecordSet.Fields("SBCTAX_TYPE").Value)) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
            If Not objTmpRecordset.EOF Then
                strTaxType = Trim(UCase(objTmpRecordset.Fields("Tx_TaxeID").Value))
            Else
                Msgbox("Tax type not found", MsgBoxStyle.Information, "eMPro")
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
            If strTaxType = "SBC" Then
                dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("SBCTAX_TYPE_Amount").Value), 0, objRecordSet.Fields("SBCTAX_TYPE_Amount").Value)
                dblBaseCurrencyAmount = dblTaxAmt
                dblTaxRate = Find_Value("select SBCTAX_TYPE_Amount from saleschallan_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  doc_no='" & Ctlinvoice.Text & "'")
                If dblBaseCurrencyAmount > 0 Then
                    strRetval = GetTaxGlSl(strTaxType)
                    If strRetval = "N" Then
                        Msgbox("GL for ARTAX is not defined for " & strTaxType, MsgBoxStyle.Information, "eMPro")
                        CreateStringForAccounts = False
                        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                            objRecordSet.Close()
                            objRecordSet = Nothing
                        End If
                        Exit Function
                    End If
                    varTmp = Split(strRetval, "»")
                    strTaxGL = varTmp(0)
                    strTaxSL = varTmp(1)
                    If UCase(Trim(cmbInvType.Text)) <> "REJECTION" Then
                        mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»" & strTaxType & "»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                    Else
                        mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & dblTaxAmt & "»»SBCTAX for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                    End If
                    iCtr = iCtr + 1
                End If
            End If
        End If

        'SBC TAX

        'KKC TAX

        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
        If Trim(IIf(IsDBNull(objRecordSet.Fields("KKCTAX_TYPE").Value), "", objRecordSet.Fields("KKCTAX_TYPE").Value)) <> "" Then
            objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("KKCTAX_TYPE").Value), "", objRecordSet.Fields("KKCTAX_TYPE").Value)) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
            If Not objTmpRecordset.EOF Then
                strTaxType = Trim(UCase(objTmpRecordset.Fields("Tx_TaxeID").Value))
            Else
                Msgbox("Tax type not found", MsgBoxStyle.Information, "eMPro")
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
            If strTaxType = "KKC" Then
                dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("KKCTAX_TYPE_Amount").Value), 0, objRecordSet.Fields("KKCTAX_TYPE_Amount").Value)
                dblBaseCurrencyAmount = dblTaxAmt
                dblTaxRate = Find_Value("select KKCTAX_TYPE_Amount from saleschallan_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  doc_no='" & Ctlinvoice.Text & "'")
                If dblBaseCurrencyAmount > 0 Then
                    strRetval = GetTaxGlSl(strTaxType)
                    If strRetval = "N" Then
                        Msgbox("GL for ARTAX is not defined for " & strTaxType, MsgBoxStyle.Information, "eMPro")
                        CreateStringForAccounts = False
                        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                            objRecordSet.Close()
                            objRecordSet = Nothing
                        End If
                        Exit Function
                    End If
                    varTmp = Split(strRetval, "»")
                    strTaxGL = varTmp(0)
                    strTaxSL = varTmp(1)

                    If UCase(Trim(cmbInvType.Text)) <> "REJECTION" Then
                        mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»" & strTaxType & "»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                    Else
                        mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & dblTaxAmt & "»»KKC TAX for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                    End If
                    iCtr = iCtr + 1
                End If
            End If
        End If

        'KKC TAX

        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
        If Trim(IIf(IsDBNull(objRecordSet.Fields("ECESS_Type").Value), "", objRecordSet.Fields("ECESS_Type").Value)) <> "" Then
            objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE  UNIT_CODE = '" & gstrUNITID & "' and TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("ECESS_Type").Value), "", objRecordSet.Fields("ECESS_Type").Value)) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
            If Not objTmpRecordset.EOF Then
                strTaxType = Trim(UCase(objTmpRecordset.Fields("Tx_TaxeID").Value))
            Else
                Msgbox("Tax type not found", MsgBoxStyle.Information, "empower")
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
                    strRetval = GetTaxGlSl(strTaxType)
                    If strRetval = "N" Then
                        Msgbox("GL for ARTAX is not defined for " & strTaxType, MsgBoxStyle.Information, "empower")
                        CreateStringForAccounts = False
                        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                            objRecordSet.Close()
                            objRecordSet = Nothing
                        End If
                        Exit Function
                    End If
                    varTmp = Split(strRetval, "»")
                    strTaxGL = varTmp(0)
                    strTaxSL = varTmp(1)
                    If UCase(Trim(cmbInvType.Text)) = "REJECTION" And CUST_REJECTION_FLAG = False Then
                        mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & dblTaxAmt & "»»ECS for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                    Else
                        mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»" & strTaxType & "»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                    End If
                    iCtr = iCtr + 1
                End If
            End If
        End If
        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
        If Trim(IIf(IsDBNull(objRecordSet.Fields("SECESS_Type").Value), "", objRecordSet.Fields("SECESS_Type").Value)) <> "" Then
            objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE  UNIT_CODE = '" & gstrUNITID & "' and TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("SECESS_Type").Value), "", objRecordSet.Fields("SECESS_Type").Value)) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
            If Not objTmpRecordset.EOF Then
                strTaxType = Trim(UCase(objTmpRecordset.Fields("Tx_TaxeID").Value))
            Else
                Msgbox("Tax type not found", MsgBoxStyle.Information, ResolveResString(100))
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
                    strRetval = GetTaxGlSl(strTaxType)
                    If strRetval = "N" Then
                        Msgbox("GL for ARTAX is not defined for " & strTaxType, MsgBoxStyle.Information, ResolveResString(100))
                        CreateStringForAccounts = False
                        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                            objRecordSet.Close()
                            objRecordSet = Nothing
                        End If
                        Exit Function
                    End If
                    varTmp = Split(strRetval, "»")
                    strTaxGL = varTmp(0)
                    strTaxSL = varTmp(1)
                    If UCase(Trim(cmbInvType.Text)) = "REJECTION" And CUST_REJECTION_FLAG = False Then
                        mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & dblTaxAmt & "»»ECSSH for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                    Else
                        mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»" & strTaxType & "»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                    End If
                    iCtr = iCtr + 1
                End If
            End If
        End If

        'If (UCase(Trim(lbldescription.Text)) = "INV") And (UCase(Trim(lblcategory.Text)) = "L") Then
        If Not (UCase(Trim(lbldescription.Text)) = "EXP") Then
            dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("TCSTaxAmount").Value), 0, objRecordSet.Fields("TCSTaxAmount").Value)
            dblBaseCurrencyAmount = dblTaxAmt
            If dblBaseCurrencyAmount > 0 Then
                'initializing the tax gl and sl here
                strRetval = GetTaxGlSl("TCS")
                If strRetval = "N" Then
                    Msgbox("GL For Purpose Code TCS Tax is not defined. ", MsgBoxStyle.Information, "eMPro")
                    CreateStringForAccounts = False
                    If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                        objRecordSet.Close()
                        objRecordSet = Nothing
                    End If
                    Exit Function
                End If
                varTmp = Split(strRetval, "»")
                strTaxGL = varTmp(0)
                strTaxSL = varTmp(1)
                If System.Math.Abs(dblBaseCurrencyAmount) > 0 Then
                    If UCase(Trim(cmbInvType.Text)) = "REJECTION" And CUST_REJECTION_FLAG = False Then
                        mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & dblTaxAmt & "»»TCS for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                    Else
                        mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»»TCS»0»" & "»»0»" & strTaxGL & "»" & strTaxSL & "»" & System.Math.Abs(dblBaseCurrencyAmount) & "»" & "Cr»»»»»»0»0»0»0»0" & "¦"
                    End If




                End If
                iCtr = iCtr + 1
            End If
        End If

        dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("Surcharge_Sales_Tax_Amount").Value), 0, objRecordSet.Fields("Surcharge_Sales_Tax_Amount").Value)
        dblBaseCurrencyAmount = dblTaxAmt
        dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("Surcharge_SalesTax_Per").Value), 0, objRecordSet.Fields("Surcharge_SalesTax_Per").Value)
        If dblBaseCurrencyAmount > 0 Then
            strRetval = GetTaxGlSl("SST")
            If strRetval = "N" Then
                Msgbox("GL for ARTAX is not defined for SST", MsgBoxStyle.Information, "empower")
                CreateStringForAccounts = False
                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSet.Close()
                    objRecordSet = Nothing
                End If
                Exit Function
            End If
            varTmp = Split(strRetval, "»")
            strTaxGL = varTmp(0)
            strTaxSL = varTmp(1)
            If rsObjItemcode.State = ADODB.ObjectStateEnum.adStateOpen Then rsObjItemcode.Close()
            rsObjItemcode.Open("SELECT sales_dtl.item_code FROM sales_dtl, item_mst (NOLOCK) WHERE sales_dtl.UNIT_CODE=item_mst.UNIT_CODE AND sales_dtl.UNIT_CODE = '" & gstrUNITID & "' and sales_dtl.Doc_No='" & Trim(Ctlinvoice.Text) & "' and sales_dtl.Item_Code=item_mst.Item_Code and sales_dtl.Location_Code='" & Trim(txtUnitCode.Text) & "' ", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If UCase(Trim(cmbInvType.Text)) = "REJECTION" And CUST_REJECTION_FLAG = False Then
                mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & dblTaxAmt & "»»Surcharge for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
            Else
                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»SST»0»" & Trim(rsObjItemcode.Fields(0).Value) & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
            End If
            iCtr = iCtr + 1
        End If
        dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("Insurance").Value), 0, objRecordSet.Fields("Insurance").Value)
        dblBaseCurrencyAmount = dblTaxAmt
        If dblBaseCurrencyAmount > 0 Then
            strRetval = GetTaxGlSl("INS")
            If strRetval = "N" Then
                Msgbox("GL for ARTAX is not defined for INS(Insurance)", MsgBoxStyle.Information, "empower")
                CreateStringForAccounts = False
                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSet.Close()
                    objRecordSet = Nothing
                End If
                Exit Function
            End If
            varTmp = Split(strRetval, "»")
            strTaxGL = varTmp(0)
            strTaxSL = varTmp(1)
            If UCase(Trim(cmbInvType.Text)) = "REJECTION" And CUST_REJECTION_FLAG = False Then
                mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & dblTaxAmt & "»»Insurance for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
            Else
                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»INS»0»" & "»»0»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
            End If
            iCtr = iCtr + 1
        End If
        dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("Frieght_Amount").Value), 0, objRecordSet.Fields("Frieght_Amount").Value)
        dblBaseCurrencyAmount = dblTaxAmt
        If dblBaseCurrencyAmount > 0 Then
            strRetval = GetTaxGlSl("FRT")
            If strRetval = "N" Then
                Msgbox("GL for ARTAX is not defined for FRT(Packing & Forwarding)", MsgBoxStyle.Information, "empower")
                CreateStringForAccounts = False
                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSet.Close()
                    objRecordSet = Nothing
                End If
                Exit Function
            End If
            varTmp = Split(strRetval, "»")
            strTaxGL = varTmp(0)
            strTaxSL = varTmp(1)
            If UCase(Trim(cmbInvType.Text)) = "REJECTION" And CUST_REJECTION_FLAG = False Then
                mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & dblTaxAmt & "»»Freight for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
            Else
                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»FRT»0»" & "»»0»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
            End If
            iCtr = iCtr + 1
        End If
        dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("Tot_Add_Excise_Amt").Value), 0, objRecordSet.Fields("Tot_Add_Excise_Amt").Value)
        dblBaseCurrencyAmount = dblTaxAmt
        If dblBaseCurrencyAmount > 0 Then
            strRetval = GetTaxGlSl("AED")
            If strRetval = "N" Then
                Msgbox("GL for ARTAX is not defined for AED(Additional excise duty)", MsgBoxStyle.Information, "empower")
                CreateStringForAccounts = False
                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSet.Close()
                    objRecordSet = Nothing
                End If
                Exit Function
            End If
            varTmp = Split(strRetval, "»")
            strTaxGL = varTmp(0)
            strTaxSL = varTmp(1)
            If UCase(Trim(cmbInvType.Text)) = "REJECTION" And CUST_REJECTION_FLAG = False Then
                mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & dblTaxAmt & "»»Add Ex Duty for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
            Else
                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»AED»0»" & "»»0»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
            End If
            iCtr = iCtr + 1
        End If
        Dim RSCSIEXGLSL As ADODB.Recordset
        Dim RSCSIEXAMT As ADODB.Recordset
        Dim objRS As New ClsResultSetDB_Invoice
        objRS.GetResult("Select a.CSIEX_Inc from Customer_Mst a (NOLOCK), saleschallan_dtl b where a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND a.Customer_code = b.Account_code and b.Doc_No = " & Ctlinvoice.Text)
        If objRS.GetNoRows > 0 Then
            blnCSIEX_Inc = objRS.GetValue("CSIEX_Inc")
        End If
        objRecordSet.Close()
        If blnCSIEX_Inc = True Then
            RSCSIEXGLSL = mP_Connection.Execute("SELECT a.CSIEX_GL,a.CSIEX_SL from Customer_Mst a, saleschallan_dtl b where a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND a.Customer_code = b.Account_code and b.Doc_No = " & Ctlinvoice.Text & " and b.Location_Code='" & Trim(txtUnitCode.Text) & "'")
            If RSCSIEXGLSL.EOF <> True Then
                STRCSIEX_GL = Trim(IIf(IsDBNull(RSCSIEXGLSL.Fields("CSIEX_GL").Value), "", RSCSIEXGLSL.Fields("CSIEX_GL").Value))
                STRCSIEX_SL = Trim(IIf(IsDBNull(RSCSIEXGLSL.Fields("CSIEX_SL").Value), "", RSCSIEXGLSL.Fields("CSIEX_SL").Value))
                If mblnCSM_Knockingoff_req Then
                    RSCSIEXAMT = mP_Connection.Execute("SELECT ISNULL(SUM(A.CSIEXCISE_AMOUNT) ,0) CSIEXCISE_AMOUNT FROM SALES_DTL A,SALESCHALLAN_DTL B WHERE A.UNIT_CODE = B.UNIT_CODE and A.UNIT_CODE = '" & gstrUNITID & "' AND A.DOC_NO = '" & Trim(Ctlinvoice.Text) & "' AND A.DOC_NO = B.DOC_NO AND A.LOCATION_CODE='" & Trim(txtUnitCode.Text) & "'")
                Else
                    RSCSIEXAMT = mP_Connection.Execute("SELECT ISNULL(SUM(A.CSIEXCISE_AMOUNT + A.CSIEXCISE_AMOUNT * B.ECESS_PER /100+A.CSIEXCISE_AMOUNT * B.SECESS_PER /100),0) CSIEXCISE_AMOUNT FROM SALES_DTL A,SALESCHALLAN_DTL B WHERE A.UNIT_CODE = B.UNIT_CODE and A.UNIT_CODE = '" & gstrUNITID & "' AND A.DOC_NO = '" & Trim(Ctlinvoice.Text) & "' AND A.DOC_NO = B.DOC_NO AND A.LOCATION_CODE='" & Trim(txtUnitCode.Text) & "' GROUP BY ECESS_PER")
                End If
                If RSCSIEXAMT.EOF <> True Then
                    DBLCSIEDCESS = System.Math.Round(RSCSIEXAMT.Fields("CSIEXCISE_AMOUNT").Value * dblExchangeRate, 2)
                End If
                If STRCSIEX_GL <> "" And STRCSIEX_SL <> "" Then
                    If DBLCSIEDCESS > 0 Then
                        mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»CSIX»0»»»»" & STRCSIEX_GL & "»" & STRCSIEX_SL & "»" & DBLCSIEDCESS & "»Dr»»»»»»0»0»0»0»0" & "¦"
                        iCtr = iCtr + 1
                    End If
                End If
            End If
            RSCSIEXGLSL = Nothing
            RSCSIEXAMT = Nothing
        End If
        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close()
        objRecordSet.Open("SELECT sales_dtl.*, item_mst.GlGrp_code FROM sales_dtl, item_mst (NOLOCK) WHERE sales_dtl.UNIT_CODE = item_mst.UNIT_CODE and sales_dtl.UNIT_CODE = '" & gstrUNITID & "' AND sales_dtl.Doc_No='" & Trim(Ctlinvoice.Text) & "' and sales_dtl.Item_Code=item_mst.Item_Code and sales_dtl.Location_Code='" & Trim(txtUnitCode.Text) & "'")
        If objRecordSet.EOF Then
            Msgbox("Item details not found.", MsgBoxStyle.Information, "empower")
            CreateStringForAccounts = False
            If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                objRecordSet.Close()
                objRecordSet = Nothing
            End If
            Exit Function
        End If
        While Not objRecordSet.EOF
            strGlGroupId = Trim(IIf(IsDBNull(objRecordSet.Fields("GlGrp_code").Value), "", objRecordSet.Fields("GlGrp_code").Value))
            dblToolCost = Val(IIf(IsDBNull(objRecordSet.Fields("Toolcost_Amount").Value), "", objRecordSet.Fields("Toolcost_amount").Value))
            If ((UCase(Trim(cmbInvType.Text)) = "REJECTION" And Trim(strCustRef) = "") Or UCase(Trim(cmbInvType.Text)) <> "REJECTION") And CUST_REJECTION_FLAG = False Then
                dblBasicAmount = IIf(IsDBNull(objRecordSet.Fields("Basic_Amount").Value), 0, objRecordSet.Fields("Basic_Amount").Value)
                If mblnAddCustomerMaterial Then
                    dblBaseCurrencyAmount = dblBasicAmount + IIf(IsDBNull(objRecordSet.Fields("CustMtrl_Amount").Value), 0, objRecordSet.Fields("CustMtrl_Amount").Value)
                Else
                    dblBaseCurrencyAmount = dblBasicAmount
                End If
                If dblBaseCurrencyAmount > 0 Then
                    strRetval = GetItemGLSL(strGlGroupId, mstrPurposeCode)
                    If strRetval = "N" Then
                        CreateStringForAccounts = False
                        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                            objRecordSet.Close()
                            objRecordSet = Nothing
                        End If
                        Exit Function
                    End If
                    varTmp = Split(strRetval, "»")
                    strItemGL = varTmp(0)
                    strItemSL = varTmp(1)
                    If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
                    objTmpRecordset.Open("SELECT * FROM invcc_dtl WHERE  UNIT_CODE = '" & gstrUNITID & "' and Invoice_Type='" & lbldescription.Text & "' AND Sub_Type = '" & lblcategory.Text & "' AND Location_Code ='" & Trim(txtUnitCode.Text) & "' AND ccM_cc_Percentage > 0", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
                    If Not objTmpRecordset.EOF Then
                        While Not objTmpRecordset.EOF
                            dblCCShare = (dblBaseCurrencyAmount / 100) * objTmpRecordset.Fields("ccM_cc_Percentage").Value
                            If UCase(Trim(cmbInvType.Text)) = "REJECTION" And CUST_REJECTION_FLAG = False Then
                                mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strItemGL & "»" & strItemSL & "»" & Trim(objTmpRecordset.Fields("ccM_ccCode").Value) & "»»»CR»" & dblCCShare & "»»Basic for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                            Else
                                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»ITM»SAL»" & iCtr & "»" & Trim(objRecordSet.Fields("item_code").Value) & "»" & strGlGroupId & "»0»" & strItemGL & "»" & strItemSL & "»" & dblCCShare & "»Cr»»" & Trim(objTmpRecordset.Fields("ccM_ccCode").Value) & "»»»»0»0»0»0»0¦"
                            End If
                            objTmpRecordset.MoveNext()
                            iCtr = iCtr + 1
                        End While
                    Else
                        If (((UCase(Trim(cmbInvType.Text))) = "TRANSFER INVOICE") Or ((UCase(Trim(cmbInvType.Text))) = "REJECTION") Or ((UCase(Trim(cmbInvType.Text))) = "INTER-DIVISION")) Then
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
                        If UCase(Trim(cmbInvType.Text)) = "REJECTION" And CUST_REJECTION_FLAG = False Then
                            mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strItemGL & "»" & strItemSL & "»" & strTaxCCCode & "»»»CR»" & dblBaseCurrencyAmount & "»»Basic for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                        Else
                            mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»ITM»SAL»" & iCtr & "»" & Trim(objRecordSet.Fields("item_code").Value) & "»" & strGlGroupId & "»0»" & strItemGL & "»" & strItemSL & "»" & dblBaseCurrencyAmount & "»Cr»» " & strTaxCCCode & "»»»»0»0»0»0»0" & "¦"
                        End If
                        iCtr = iCtr + 1
                    End If
                End If
            End If
            If mblnEOUUnit = False Then
                dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("Excise_Tax").Value), 0, objRecordSet.Fields("Excise_Tax").Value)
            Else
                dblTaxAmt = ((IIf(IsDBNull(objRecordSet.Fields("Excise_Tax").Value), 0, objRecordSet.Fields("Excise_Tax").Value) + IIf(IsDBNull(objRecordSet.Fields("CVD_amount").Value), 0, objRecordSet.Fields("CVD_amount").Value) + IIf(IsDBNull(objRecordSet.Fields("SVD_amount").Value), 0, objRecordSet.Fields("SVD_amount").Value)) / 2)
            End If
            If mblnExciseRoundOFFFlag Then dblTaxAmt = System.Math.Round(dblTaxAmt, 0)
            dblBaseCurrencyAmount = dblTaxAmt
            If dblBaseCurrencyAmount > 0 Then
                strRetval = GetTaxGlSl("EXC")
                If strRetval = "N" Then
                    Msgbox("GL for ARTAX is not defined for EXC", MsgBoxStyle.Information, "empower")
                    CreateStringForAccounts = False
                    If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                        objRecordSet.Close()
                        objRecordSet = Nothing
                    End If
                    Exit Function
                End If
                varTmp = Split(strRetval, "»")
                strTaxGL = varTmp(0)
                strTaxSL = varTmp(1)
                If UCase(Trim(cmbInvType.Text)) = "REJECTION" And CUST_REJECTION_FLAG = False Then
                    mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & dblTaxAmt & "»»Excise for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                Else
                    mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»EXC»0»" & Trim(objRecordSet.Fields("item_code").Value) & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                End If
                iCtr = iCtr + 1
            End If
            dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("pkg_Amount").Value), 0, objRecordSet.Fields("Pkg_Amount").Value)
            dblBaseCurrencyAmount = dblTaxAmt
            If dblBaseCurrencyAmount > 0 Then
                strRetval = GetTaxGlSl("PKT")
                If strRetval = "N" Then
                    Msgbox("GL for ARTAX is not defined for PACKING", MsgBoxStyle.Information, "empower")
                    CreateStringForAccounts = False
                    If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                        objRecordSet.Close()
                        objRecordSet = Nothing
                    End If
                    Exit Function
                End If
                varTmp = Split(strRetval, "»")
                strTaxGL = varTmp(0)
                strTaxSL = varTmp(1)
                If UCase(Trim(cmbInvType.Text)) = "REJECTION" And CUST_REJECTION_FLAG = False Then
                    mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & dblTaxAmt & "»»Packing Charges for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                Else
                    mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»PKT»0»" & "»»0»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                End If
                iCtr = iCtr + 1
            End If
            RSTOOLGET = mP_Connection.Execute("SELECT ACCOUNT_CODE,A.CUST_REF,A.AMENDMENT_NO,ITEM_CODE,CUST_ITEM_CODE FROM SALESCHALLAN_DTL A (NOLOCK), SALES_DTL B (NOLOCK) WHERE A.UNIT_CODE = B.UNIT_CODE and A.UNIT_CODE = '" & gstrUNITID & "' AND A.DOC_NO = B.DOC_NO AND A.DOC_NO = '" & Trim(Ctlinvoice.Text) & "'  AND B.ITEM_CODE = '" & Trim(objRecordSet.Fields("item_code").Value) & "'")
            If RSTOOLGET.EOF <> True Then
                RSTOOLPOST = mP_Connection.Execute("select isnull(TOOL_AMOR_FLAG,0) as TOOL_AMOR_FLAG From CUST_ORD_DTL (NOLOCK) where  UNIT_CODE = '" & gstrUNITID & "' and Account_Code = '" & Trim(RSTOOLGET.Fields("Account_Code").Value) & "' and Cust_Ref = '" & Trim(RSTOOLGET.Fields("Cust_Ref").Value) & "' and Amendment_No = '" & Trim(RSTOOLGET.Fields("Amendment_no").Value) & "' and Item_Code = '" & Trim(RSTOOLGET.Fields("Item_code").Value) & "' and Cust_DrgNo = '" & Trim(RSTOOLGET.Fields("Cust_Item_Code").Value) & "'")
                If RSTOOLPOST.EOF <> True Then
                    TOOL_AMOR_FLAG = RSTOOLPOST.Fields("TOOL_AMOR_FLAG").Value
                End If
            End If
            RSTOOLPOST = Nothing
            RSTOOLGET = Nothing
            If rstoolglslrecordset.State = 1 Then rstoolglslrecordset.Close()
            rstoolglslrecordset.CursorLocation = ADODB.CursorLocationEnum.adUseClient
            rstoolglslrecordset.Open("SELECT a.tool_gl,a.tool_sl from Customer_Mst a, saleschallan_dtl b where a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND  a.Customer_code = b.Account_code and b.Doc_No = " & Ctlinvoice.Text & " and b.Location_Code='" & Trim(txtUnitCode.Text) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If Not rstoolglslrecordset.EOF Then
                strToolGLCr = IIf(IsDBNull(Trim(rstoolglslrecordset.Fields(0).Value.ToString)), "", Trim(rstoolglslrecordset.Fields(0).Value.ToString))
                strToolSLCr = Trim(IIf(IsDBNull(rstoolglslrecordset.Fields(1).Value), "", rstoolglslrecordset.Fields(1).Value))
                dblBaseCurrencyAmount = dblToolCost * dblExchangeRate
                If strToolGLCr <> "" Then
                    If dblToolCost > 0 And TOOL_AMOR_FLAG = True Then
                        mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»TOOL»0»»»»" & strToolGLCr & "»" & strToolSLCr & "»" & dblBaseCurrencyAmount & "»Dr»»»»»»0»0»0»0»0" & "¦"
                        iCtr = iCtr + 1
                    End If
                End If
            End If
            If rstoolglslrecordset.State = 1 Then rstoolglslrecordset.Close()
            dblBaseCurrencyAmount = dblToolCost * dblExchangeRate
            If strToolGLCr <> "" Then ''if Tool Gl for customer is not defined then it is not accounted
                If dblToolCost > 0 And TOOL_AMOR_FLAG = True Then
                    rstoolglslrecordset.CursorLocation = ADODB.CursorLocationEnum.adUseClient
                    rstoolglslrecordset.Open("SELECT isnull(gbl_glcode,0),isnull(gbl_slcode,0) FROM Fin_globalgl WHERE  UNIT_CODE = '" & gstrUNITID & "' and gbl_prpscode='TOOLAMOR'", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                    If Not rstoolglslrecordset.EOF Then
                        strToolGLDr = Trim(rstoolglslrecordset.Fields(0).Value)
                        strToolSLDr = Trim(rstoolglslrecordset.Fields(1).Value)
                    Else
                        Msgbox("Purpose code not defined for Tool Amortization", MsgBoxStyle.Information, "empower")
                        CreateStringForAccounts = False
                        Exit Function
                    End If
                    If rstoolglslrecordset.State = 1 Then rstoolglslrecordset.Close()
                    mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»TOOL»0»»»»" & strToolGLDr & "»" & strToolSLDr & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                    iCtr = iCtr + 1
                End If
            End If
            dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("Others").Value), 0, objRecordSet.Fields("Others").Value)
            dblBaseCurrencyAmount = dblTaxAmt
            strRetval = GetTaxGlSl("OTH")
            If strRetval = "N" Then
                Msgbox("GL for ARTAX is not defined for OTHERS", MsgBoxStyle.Information, "eMPro")
                CreateStringForAccounts = False
                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                    objRecordSet.Close()
                    objRecordSet = Nothing
                End If
                Exit Function
            End If
            varTmp = Split(strRetval, "»")
            strTaxGL = varTmp(0)
            strTaxSL = varTmp(1)
            If dblBaseCurrencyAmount > 0 Then
                If UCase(Trim(cmbInvType.Text)) = "REJECTION" And CUST_REJECTION_FLAG = False Then
                    mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" & dblTaxAmt & "»»Other Charges for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                Else
                    mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»OTH»0»" & Trim(objRecordSet.Fields("item_code").Value) & "»»0»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
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
                        Msgbox("Tax type not found", vbInformation, ResolveResString(100))
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
                            strRetval = GetTaxGlSl(strTaxType)
                            If strRetval = "N" Then
                                Msgbox("GL for ARTAX is not defined for " & strTaxType, vbInformation, ResolveResString(100))
                                CreateStringForAccounts = False
                                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close() : objRecordSet = Nothing
                                Exit Function
                            End If
                            varTmp = Split(strRetval, "»")
                            strTaxGL = varTmp(0)
                            strTaxSL = varTmp(1)
                            If UCase$(Trim(cmbInvType.Text)) <> "REJECTION" Then
                                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" &
                                iCtr & "»TAX»CGST»0»" & Trim(objRecordSet.Fields("item_code").Value) &
                               "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                            Else
                                mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" &
                                         dblTaxAmt & "»»CGST for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
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
                        Msgbox("Tax type not found", vbInformation, ResolveResString(100))
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
                            strRetval = GetTaxGlSl(strTaxType)
                            If strRetval = "N" Then
                                Msgbox("GL for ARTAX is not defined for " & strTaxType, vbInformation, ResolveResString(100))
                                CreateStringForAccounts = False
                                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close() : objRecordSet = Nothing
                                Exit Function
                            End If
                            varTmp = Split(strRetval, "»")
                            strTaxGL = varTmp(0)
                            strTaxSL = varTmp(1)
                            If UCase$(Trim(cmbInvType.Text)) <> "REJECTION" Then
                                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" &
                                iCtr & "»TAX»SGST»0»" & Trim(objRecordSet.Fields("item_code").Value) &
                               "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                            Else
                                mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" &
                                         dblTaxAmt & "»»SGST for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
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
                        Msgbox("Tax type not found", vbInformation, ResolveResString(100))
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
                            strRetval = GetTaxGlSl(strTaxType)
                            If strRetval = "N" Then
                                Msgbox("GL for ARTAX is not defined for " & strTaxType, vbInformation, ResolveResString(100))
                                CreateStringForAccounts = False
                                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close() : objRecordSet = Nothing
                                Exit Function
                            End If
                            varTmp = Split(strRetval, "»")
                            strTaxGL = varTmp(0)
                            strTaxSL = varTmp(1)
                            If UCase$(Trim(cmbInvType.Text)) <> "REJECTION" Then
                                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" &
                                iCtr & "»TAX»IGST»0»" & Trim(objRecordSet.Fields("item_code").Value) &
                               "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                            Else
                                mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" &
                                         dblTaxAmt & "»»IGST for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                            End If
                            iCtr = iCtr + 1
                        End If
                    End If
                End If
                'IGSTR

                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
                If Trim(IIf(IsDBNull(objRecordSet.Fields("IGSTTXRT_TYPE").Value), "", objRecordSet.Fields("IGSTTXRT_TYPE").Value)) <> "" Then
                    objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("IGSTTXRT_TYPE").Value), "", objRecordSet.Fields("IGSTTXRT_TYPE").Value)) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
                    If Not objTmpRecordset.EOF Then
                        strTaxType = Trim(UCase$(objTmpRecordset.Fields("Tx_TaxeID").Value))
                    Else
                        Msgbox("Tax type not found", vbInformation, ResolveResString(100))
                        CreateStringForAccounts = False
                        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close() : objRecordSet = Nothing
                        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close() : objTmpRecordset = Nothing
                        Exit Function
                    End If
                    If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
                    If strTaxType = "IGST" And mblnSEZ_Knockingoff_req = True And cmbInvType.Text = "EXPORT INVOICE" Then
                        dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("IGST_AMT").Value), 0, objRecordSet.Fields("IGST_AMT").Value)
                        dblBaseCurrencyAmount = dblTaxAmt
                        dblTaxRate = IIf(IsDBNull(objRecordSet.Fields("IGST_PERCENT").Value), 0, objRecordSet.Fields("IGST_PERCENT").Value)
                        If dblBaseCurrencyAmount > 0 Then
                            strRetval = GetTaxGlSl_IGST_RECOVERY(strTaxType)
                            If strRetval = "N" Then
                                Msgbox("GL for ARTAX is not defined for " & strTaxType, vbInformation, ResolveResString(100))
                                CreateStringForAccounts = False
                                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close() : objRecordSet = Nothing
                                Exit Function
                            End If
                            varTmp = Split(strRetval, "»")
                            strTaxGL = varTmp(0)
                            strTaxSL = varTmp(1)
                            If UCase$(Trim(cmbInvType.Text)) <> "REJECTION" Then
                                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" &
                                iCtr & "»TAX»IGST»0»" & Trim(objRecordSet.Fields("item_code").Value) &
                               "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Dr»»»»»»0»0»0»0»0" & "¦"
                            Else
                                mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»DR»" &
                                         dblTaxAmt & "»»IGST for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                            End If
                            iCtr = iCtr + 1
                        End If
                    End If
                End If
                'IGSTR
                'UTGST
                If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
                If Trim(IIf(IsDBNull(objRecordSet.Fields("UTGSTTXRT_TYPE").Value), "", objRecordSet.Fields("UTGSTTXRT_TYPE").Value)) <> "" Then
                    objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE UNIT_CODE='" + gstrUNITID + "' AND  TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("UTGSTTXRT_TYPE").Value), "", objRecordSet.Fields("UTGSTTXRT_TYPE").Value)) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
                    If Not objTmpRecordset.EOF Then
                        strTaxType = Trim(UCase$(objTmpRecordset.Fields("Tx_TaxeID").Value))
                    Else
                        Msgbox("Tax type not found", vbInformation, ResolveResString(100))
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
                            strRetval = GetTaxGlSl(strTaxType)
                            If strRetval = "N" Then
                                Msgbox("GL for ARTAX is not defined for " & strTaxType, vbInformation, ResolveResString(100))
                                CreateStringForAccounts = False
                                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close() : objRecordSet = Nothing
                                Exit Function
                            End If
                            varTmp = Split(strRetval, "»")
                            strTaxGL = varTmp(0)
                            strTaxSL = varTmp(1)
                            If UCase$(Trim(cmbInvType.Text)) <> "REJECTION" Then
                                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" &
                                iCtr & "»TAX»UTGST»0»" & Trim(objRecordSet.Fields("item_code").Value) &
                               "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                            Else
                                mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" &
                                         dblTaxAmt & "»»UTGST for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
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
                        Msgbox("Tax type not found", vbInformation, ResolveResString(100))
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
                            strRetval = GetTaxGlSl(strTaxType)
                            If strRetval = "N" Then
                                Msgbox("GL for ARTAX is not defined for " & strTaxType, vbInformation, ResolveResString(100))
                                CreateStringForAccounts = False
                                If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close() : objRecordSet = Nothing
                                Exit Function
                            End If
                            varTmp = Split(strRetval, "»")
                            strTaxGL = varTmp(0)
                            strTaxSL = varTmp(1)
                            If UCase$(Trim(cmbInvType.Text)) <> "REJECTION" Then
                                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" &
                                iCtr & "»TAX»CCESS»0»" & Trim(objRecordSet.Fields("item_code").Value) &
                               "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                            Else
                                mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strTaxGL & "»" & strTaxSL & "»»»»CR»" &
                                         dblTaxAmt & "»»CCESS for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                            End If
                            iCtr = iCtr + 1
                        End If
                    End If
                End If
            End If
            '101188073 End
            objRecordSet.MoveNext()
        End While
        strRetval = GetItemGLSL("", "Rounded_Amt")
        If strRetval = "N" Then
            CreateStringForAccounts = False
            If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                objRecordSet.Close()
                objRecordSet = Nothing
            End If
            Exit Function
        End If
        varTmp = Split(strRetval, "»")
        strItemGL = varTmp(0)
        strItemSL = varTmp(1)
        If mblnTotalInvoiceAmountRoundOff = True Then
            dblBaseCurrencyAmount = System.Math.Round(dblInvoiceAmt - System.Math.Round(dblInvoiceAmt, 0), 4)
        Else
            dblBaseCurrencyAmount = dblInvoiceAmt - System.Math.Round(dblInvoiceAmt, mintTotalInvoiceAmountRoundOff)
        End If
        If (UCase(gstrUNITID) = "MSD" Or UCase(gstrUNITID) = "MBJ") And dblBaseCurrencyAmount = 0 Then
            dblBaseCurrencyAmount = dblInvoiceAmtRoundOff_diff
            dblBaseCurrencyAmount = System.Math.Round(dblBaseCurrencyAmount, 4)
        End If
        '        dblBaseCurrencyAmount = System.Math.Round(dblInvoiceAmt - System.Math.Round(dblInvoiceAmt, 0), 4)
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
        If System.Math.Abs(dblBaseCurrencyAmount) > 0 Then
            If UCase(Trim(cmbInvType.Text)) = "REJECTION" And CUST_REJECTION_FLAG = False Then
                If dblBaseCurrencyAmount < 0 Then
                    mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strItemGL & "»" & strItemSL & "»" & strTaxCCCode & "»»»CR»" & System.Math.Abs(dblBaseCurrencyAmount) & "»Rounding off amount for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                Else
                    mstrDetailString = mstrDetailString & "M»»" & iCtr & "»»»" & strItemGL & "»" & strItemSL & "»" & strTaxCCCode & "»»»DR»" & System.Math.Abs(dblBaseCurrencyAmount) & "»Rounding off amount for Rej. Inv. " & strInvoiceNo & "»0»0»0»0»0»0»0¦"
                End If
            Else
                mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»»RND»0»" & "»»0»" & strItemGL & "»" & strItemSL & "»" & System.Math.Abs(dblBaseCurrencyAmount) & "»"
                If dblBaseCurrencyAmount < 0 Then
                    mstrDetailString = mstrDetailString & "Cr»»" & strTaxCCCode & "»»»»0»0»0»0»0" & "¦"
                Else
                    mstrDetailString = mstrDetailString & "Dr»»" & strTaxCCCode & "»»»»0»0»0»0»0" & "¦"
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
        objRecordSet.Open("SELECT invGld_glcode, invGld_slcode FROM fin_InvGLGrpDtl WHERE  UNIT_CODE = '" & gstrUNITID & "' and invGld_prpsCode = '" & PurposeCode & "' AND invGld_invGLGrpId = '" & InventoryGlGroup & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
        If objRecordSet.EOF Then
            objRecordSet.Close()
            objRecordSet.Open("SELECT gbl_glCode, gbl_slCode FROM fin_globalGL WHERE  UNIT_CODE = '" & gstrUNITID & "' and gbl_prpsCode = '" & PurposeCode & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
            If objRecordSet.EOF Then
                GetItemGLSL = "N"
                Msgbox("GL and SL not defined for Purpose Code: " & PurposeCode, MsgBoxStyle.Information, "empower")
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
            Msgbox("GL and SL not defined for Purpose Code:" & PurposeCode, MsgBoxStyle.Information, "empower")
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

    Public Function UpdateinSale_Dtl() As Boolean
        Dim rssaledtl As ClsResultSetDB_Invoice
        Dim rsSaleConf As ClsResultSetDB_Invoice
        Dim rsSalesChallan As ClsResultSetDB_Invoice
        Dim strSql As String
        Dim strInvoiceDate As String
        Dim strStockLocCode As String
        Dim intRow, intloopcount As Short
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
        Dim rsMktSchedule As ClsResultSetDB_Invoice
        Dim rsSalesParameter As New ClsResultSetDB_Invoice
        Dim strToolCode As String
        strupdateitbalmst = ""
        strupdatecustodtdtl = ""
        strUpdateAmorDtl = ""
        strupdateamordtlbom = ""
        On Error GoTo Err_Handler
        strSql = "select * from Saleschallan_Dtl where Doc_No =" & Me.Ctlinvoice.Text & "  and Location_Code='" & Trim(txtUnitCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "'"
        rsSalesChallan = New ClsResultSetDB_Invoice
        rsSalesChallan.GetResult(strSql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        strInvoiceDate = VB6.Format(rsSalesChallan.GetValue("Invoice_Date"), gstrDateFormat)
        rsSaleConf = New ClsResultSetDB_Invoice

        rsSaleConf.GetResult("Select Stock_Location from saleconf (nolock) where  UNIT_CODE = '" & gstrUNITID & "' and Description = '" & Me.cmbInvType.Text & "' and Sub_Type_Description ='" & Me.CmbCategory.Text & "' and Location_Code='" & Trim(txtUnitCode.Text) & "'and datediff(dd,'" & getDateForDB(strInvoiceDate) & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0")
        strStockLocCode = rsSaleConf.GetValue("Stock_Location")
        rsSaleConf.ResultSetClose()

        strSql = "Select * from sales_Dtl where Doc_No = " & Me.Ctlinvoice.Text & " and Location_Code='" & Trim(txtUnitCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "'"
        rsSalesParameter.GetResult("Select CheckToolAmortisation from Sales_Parameter WHERE UNIT_CODE = '" & gstrUNITID & "'")
        If rsSalesParameter.GetNoRows > 0 Then
            rsSalesParameter.MoveFirst()
            If Len(Trim(rsSalesParameter.GetValue("CheckToolAmortisation"))) = 0 Then
                Msgbox("First define Check Tool Amortisation in Sales Parameter", MsgBoxStyle.Information, "eMPro")
                rsSalesParameter.ResultSetClose()
                UpdateinSale_Dtl = False
                Exit Function
            End If
            blnCheckToolCost = rsSalesParameter.GetValue("CheckToolAmortisation")
        End If

        rsSalesParameter.ResultSetClose()
        'sattu
        If rsSalesChallan.GetValue("MultipleSo") = "True" Then
            'Commented by shabbir : Wrong Query need to be replaced
            strupdatecustodtdtl = "Update B Set B.Despatch_Qty = B.Despatch_Qty + a.Sales_Quantity"
            strupdatecustodtdtl = strupdatecustodtdtl & " from Sales_dtl A,Cust_ord_dtl B,Emp_InvoiceSOLinkage C "
            strupdatecustodtdtl = strupdatecustodtdtl & " Where B.Cust_Ref = C.Cust_Ref And B.Amendment_No = C.Amendment_No "
            strupdatecustodtdtl = strupdatecustodtdtl & " And A.item_code = B.Item_code And B.Doc_no = C.Doc_no And A.Cust_Item_Code = B.Cust_DrgNo"
            strupdatecustodtdtl = strupdatecustodtdtl & " And B.Active_Flag = 'A' and B.Authorized_flag = 1 and A.unit_code = B.Unit_code AND B.UNIT_CODE = C.UNIT_CODE AND A.UNIT_CODE = '" & gstrUNITID & "' "
            strupdatecustodtdtl = strupdatecustodtdtl & " And A.doc_no = " & mInvNo & ""
        Else
            'Changed by Shabbir/Rajpal on 01st Apr 2013
            strupdatecustodtdtl = "Update B "
            strupdatecustodtdtl = strupdatecustodtdtl & "set Despatch_Qty = Despatch_Qty + a.Sales_Quantity "
            strupdatecustodtdtl = strupdatecustodtdtl & "FROM SALES_DTL A,CUST_ORD_DTL B,SALESCHALLAN_DTL C "
            strupdatecustodtdtl = strupdatecustodtdtl & "WHERE	A.UNIT_CODE=C.UNIT_CODE "
            strupdatecustodtdtl = strupdatecustodtdtl & "       AND A.DOC_NO=C.DOC_NO "
            strupdatecustodtdtl = strupdatecustodtdtl & "       AND A.LOCATION_CODE=C.LOCATION_CODE "
            strupdatecustodtdtl = strupdatecustodtdtl & "       AND A.UNIT_CODE=B.UNIT_CODE "
            strupdatecustodtdtl = strupdatecustodtdtl & "       AND A.CUST_ITEM_CODE = B.CUST_DRGNO "
            strupdatecustodtdtl = strupdatecustodtdtl & "       AND A.ITEM_CODE = B.ITEM_CODE "
            strupdatecustodtdtl = strupdatecustodtdtl & "       AND B.ACCOUNT_CODE=C.ACCOUNT_CODE "
            strupdatecustodtdtl = strupdatecustodtdtl & "       AND B.CUST_REF =C.CUST_REF "
            strupdatecustodtdtl = strupdatecustodtdtl & "       AND B.AMENDMENT_NO =C.AMENDMENT_NO "
            strupdatecustodtdtl = strupdatecustodtdtl & "       AND B.ACTIVE_FLAG = 'A' "
            strupdatecustodtdtl = strupdatecustodtdtl & "       AND B.AUTHORIZED_FLAG = 1 "
            strupdatecustodtdtl = strupdatecustodtdtl & "       AND C.ACCOUNT_CODE ='" & mAccount_Code & "' "
            strupdatecustodtdtl = strupdatecustodtdtl & "       AND B.CUST_REF ='" & mCust_Ref & "' "
            strupdatecustodtdtl = strupdatecustodtdtl & "       AND B.AMENDMENT_NO ='" & mAmendment_No & "' "
            strupdatecustodtdtl = strupdatecustodtdtl & "       AND B.ACTIVE_FLAG ='A' "
            strupdatecustodtdtl = strupdatecustodtdtl & "       AND A.UNIT_CODE ='" & gstrUNITID & "' "
            strupdatecustodtdtl = strupdatecustodtdtl & "       AND A.DOC_NO=" & mInvNo & ""
        End If

        rssaledtl = New ClsResultSetDB_Invoice
        rssaledtl.GetResult(strSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rssaledtl.GetNoRows > 0 Then
            intRow = rssaledtl.GetNoRows
            rssaledtl.MoveFirst()
            For intloopcount = 1 To intRow
                If Not rssaledtl.EOFRecord Then
                    mItem_Code = rssaledtl.GetValue("Item_Code")
                    mCust_Item_Code = rssaledtl.GetValue("Cust_Item_Code")
                    mSales_Quantity = IIf(rssaledtl.GetValue("Sales_Quantity") = "", 0, rssaledtl.GetValue("Sales_Quantity"))
                    mToolCost = Val(rssaledtl.GetValue("toolCost_amount"))

                    strupdateitbalmst = Trim(strupdateitbalmst) & "Update Itembal_mst set cur_bal= cur_bal-"
                    strupdateitbalmst = strupdateitbalmst & mSales_Quantity & " where Location_code = '" & strStockLocation
                    strupdateitbalmst = strupdateitbalmst & "' and item_code = '" & mItem_Code & "' and unit_code = '" & gstrUNITID & "'"

                    'If rsSalesChallan.GetValue("MultipleSo") = "True" Then
                    '    strupdatecustodtdtl = "Update B Set B.Despatch_Qty = B.Despatch_Qty + a.Sales_Quantity"
                    '    strupdatecustodtdtl = strupdatecustodtdtl & " from Sales_dtl A,Cust_ord_dtl B,Emp_InvoiceSOLinkage C "
                    '    strupdatecustodtdtl = strupdatecustodtdtl & " Where B.Cust_Ref = C.Cust_Ref And B.Amendment_No = C.Amendment_No "
                    '    strupdatecustodtdtl = strupdatecustodtdtl & " And A.item_code = B.Item_code And B.Doc_no = C.Doc_no And A.Cust_Item_Code = B.Cust_DrgNo"
                    '    strupdatecustodtdtl = strupdatecustodtdtl & " And B.Active_Flag = 'A' and B.Authorized_flag = 1 and A.unit_code = B.Unit_code AND B.UNIT_CODE = C.UNIT_CODE AND A.UNIT_CODE = '" & gstrUNITID & "' "
                    '    strupdatecustodtdtl = strupdatecustodtdtl & " And A.doc_no = " & mInvNo & ""
                    'Else
                    '    strupdatecustodtdtl = Trim(strupdatecustodtdtl) & "Update Cust_ord_dtl set Despatch_Qty = Despatch_Qty + "
                    '    strupdatecustodtdtl = strupdatecustodtdtl & mSales_Quantity & " where Account_code ='"
                    '    strupdatecustodtdtl = strupdatecustodtdtl & mAccount_Code & "'and Cust_DrgNo = '"
                    '    strupdatecustodtdtl = strupdatecustodtdtl & mCust_Item_Code & "' and Cust_ref = '" & mCust_Ref
                    '    strupdatecustodtdtl = strupdatecustodtdtl & "'and amendment_no = '" & mAmendment_No & "' and active_Flag ='A' AND UNIT_CODE = '" & gstrUNITID & "'"
                    'End If

                    If blnCheckToolCost = True Then
                        strItembal = "select BalanceQty = isnull(a.proj_qty,0) - isnull(a.ClosingValueSMIEL,0),a.Tool_C from Amor_dtl a,Tool_Mst b"
                        strItembal = strItembal & " where a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND  account_code = '" & strAccountCode & "'"
                        strItembal = strItembal & " and Item_code = '" & mItem_Code & "' and a.Tool_c = b.tool_c and a.Item_code = b.Product_No order by a.tool_c"
                        rsMktSchedule = New ClsResultSetDB_Invoice
                        rsMktSchedule.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        If rsMktSchedule.GetNoRows > 0 Then
                            rsMktSchedule.MoveFirst()
                            strQuantity = CStr(Val(rsMktSchedule.GetValue("BalanceQty")))
                            strToolCode = rsMktSchedule.GetValue("tool_c")
                            rsMktSchedule.ResultSetClose()
                            strItembal = "select BalanceQty = sum(isnull(usedProjQty,0)) from Amor_dtl a "
                            strItembal = strItembal & " Where Item_code = '" & mItem_Code & "' and a.Tool_c = '" & strToolCode & "' and a.UNIT_CODE = '" & gstrUNITID & "'"
                            rsMktSchedule = New ClsResultSetDB_Invoice
                            rsMktSchedule.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                            strQuantity = CStr(CDbl(Val(strQuantity)) - Val(rsMktSchedule.GetValue("BalanceQty")))
                            rsMktSchedule.ResultSetClose()
                            If Val(CStr(mSales_Quantity)) > Val(strQuantity) Then
                                If Val(strQuantity) = 0 Then
                                    Msgbox("No Balance Available for Item (" & mItem_Code & ") and customer Part Code (" & mCust_Item_Code & ") For Amortisation Calculations. ", MsgBoxStyle.OkOnly, "eMPro")
                                Else
                                    Msgbox("Quantity should not be Greater then available Balance Quantity for Amortisarion " & strQuantity, MsgBoxStyle.OkOnly, "eMPro")
                                End If
                                UpdateinSale_Dtl = False
                                Exit Function
                            Else
                                strUpdateAmorDtl = Trim(strUpdateAmorDtl) & " Update Amor_dtl set usedProjQty = "
                                strUpdateAmorDtl = Trim(strUpdateAmorDtl) & " isnull(usedProjQty,0) + " & mSales_Quantity
                                strUpdateAmorDtl = Trim(strUpdateAmorDtl) & " where account_code = '" & strAccountCode
                                strUpdateAmorDtl = Trim(strUpdateAmorDtl) & "' and tool_c = '" & strToolCode & "'"
                                strUpdateAmorDtl = Trim(strUpdateAmorDtl) & " and item_code = '" & mItem_Code & "' and UNIT_CODE = '" & gstrUNITID & "'"
                            End If
                        Else
                            strItembal = "select BalanceQty = isnull(proj_qty,0) - isnull(ClosingValueSMIEL,0) from Amor_dtl "
                            strItembal = strItembal & " where account_code = '" & strAccountCode & "' and UNIT_CODE = '" & gstrUNITID & "' "
                            strItembal = strItembal & " and Item_code = '" & mItem_Code & "' "
                            rsMktSchedule = New ClsResultSetDB_Invoice
                            rsMktSchedule.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                            If rsMktSchedule.GetNoRows > 0 Then
                                rsMktSchedule.MoveFirst()
                                strQuantity = CStr(Val(rsMktSchedule.GetValue("BalanceQty")))
                                rsMktSchedule.ResultSetClose()
                                strItembal = "select BalanceQty = sum(isnull(usedProjQty,0)) from Amor_dtl "
                                strItembal = strItembal & " Where Item_code = '" & mItem_Code & "' and UNIT_CODE = '" & gstrUNITID & "'"
                                rsMktSchedule = New ClsResultSetDB_Invoice
                                rsMktSchedule.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                                strQuantity = CStr(Val(strQuantity) - Val(rsMktSchedule.GetValue("BalanceQty")))
                                rsMktSchedule.ResultSetClose()
                                If Val(CStr(mSales_Quantity)) > Val(strQuantity) Then
                                    If Val(strQuantity) = 0 Then
                                        Msgbox("No Balance Available for Item (" & mItem_Code & ") and customer Part Code (" & mCust_Item_Code & ") For Amortisation Calculations. ", MsgBoxStyle.OkOnly, "eMPro")
                                    Else
                                        Msgbox("Quantity should not be Greater then available Balance Quantity for Amortisarion " & strQuantity, MsgBoxStyle.OkOnly, "eMPro")
                                    End If
                                    UpdateinSale_Dtl = False
                                    Exit Function
                                Else
                                    strUpdateAmorDtl = Trim(strUpdateAmorDtl) & " Update Amor_dtl set usedProjQty = "
                                    strUpdateAmorDtl = Trim(strUpdateAmorDtl) & " isnull(usedProjQty,0) + " & mSales_Quantity
                                    strUpdateAmorDtl = Trim(strUpdateAmorDtl) & " where account_code = '" & strAccountCode & "'"
                                    strUpdateAmorDtl = Trim(strUpdateAmorDtl) & " and item_code = '" & mItem_Code & "' and UNIT_CODE = '" & gstrUNITID & "'"
                                End If
                            End If
                        End If
                        With mP_Connection
                            .Execute("DELETE FROM tmpBOM WHERE UNIT_CODE ='" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            .Execute("BOMExplosion '" & Trim(mItem_Code) & "','" & Trim(mItem_Code) & "',1,0,'" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        End With
                        rsbom.GetResult("select * from tmpBOM WHERE UNIT_CODE = '" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        If rsbom.GetNoRows > 0 Then
                            irowcount = rsbom.GetNoRows
                            rsbom.MoveFirst()
                            For intRwCount1 = 1 To irowcount
                                strItembal = "select BalanceQty = isnull(a.proj_qty,0) - isnull(a.ClosingValueSMIEL,0),a.tool_C from Amor_dtl a, tool_mst b "
                                strItembal = strItembal & " where  a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND account_code = '" & Trim(strAccountCode) & "'"
                                strItembal = strItembal & " and Item_code = '" & rsbom.GetValue("item_code") & "' and a.Tool_c = b.Tool_c and a.ITem_code = b.Product_no order by a.tool_c"
                                rsMktSchedule = New ClsResultSetDB_Invoice
                                rsMktSchedule.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                                If rsMktSchedule.GetNoRows > 0 Then
                                    rsMktSchedule.MoveFirst()
                                    strQuantity = CStr(Val(rsMktSchedule.GetValue("BalanceQty")))
                                    strToolCode = rsMktSchedule.GetValue("tool_c")
                                    varItemQty1 = CStr(mSales_Quantity * Val(rsbom.GetValue("grossweight")))
                                    strItembal = "select BalanceQty = sum(isnull(usedProjQty,0)) from Amor_dtl a "
                                    strItembal = strItembal & " where account_code = '" & Trim(strAccountCode) & "' and a.UNIT_CODE = '" & gstrUNITID & "'"
                                    strItembal = strItembal & " and Item_code = '" & rsbom.GetValue("item_code") & "' and a.Tool_c '" & strToolCode & "'"
                                    strQuantity = CStr(Val(strQuantity) - Val(rsMktSchedule.GetValue("BalanceQty")))
                                    rsMktSchedule.ResultSetClose()
                                    rsMktSchedule = New ClsResultSetDB_Invoice
                                    rsMktSchedule.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                                    rsMktSchedule.ResultSetClose()
                                    If Val(varItemQty1) > Val(strQuantity) Then
                                        If Val(strQuantity) = 0 Then
                                            Msgbox("No Balance Available for Item (" & rsbom.GetValue("item_code") & ") and customer Part Code (" & mCust_Item_Code & ") For Amortisation Calculations. ", MsgBoxStyle.OkOnly, "eMPro")
                                        Else
                                            Msgbox("Quantity should not be Greater then available Balance Quantity for Amortisarion of this Item (" & rsbom.GetValue("item_code") & ")" & mSales_Quantity, MsgBoxStyle.OkOnly, "eMPro")
                                        End If
                                        UpdateinSale_Dtl = False
                                        Exit Function
                                    Else
                                        strupdateamordtlbom = Trim(strupdateamordtlbom) & " Update Amor_dtl set usedProjQty = "
                                        strupdateamordtlbom = Trim(strupdateamordtlbom) & " isnull(usedProjQty,0) + " & mSales_Quantity * Val(rsbom.GetValue("grossweight"))
                                        strupdateamordtlbom = Trim(strupdateamordtlbom) & " where account_code = '" & strAccountCode
                                        strupdateamordtlbom = Trim(strupdateamordtlbom) & "' and Item_code = '" & rsbom.GetValue("item_code")
                                        strupdateamordtlbom = Trim(strupdateamordtlbom) & "' and tool_c = '" & strToolCode & "' and UNIT_CODE = '" & gstrUNITID & "'"
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
        rsSalesChallan.ResultSetClose()
        rssaledtl.ResultSetClose()
        rssaledtl = Nothing
        UpdateinSale_Dtl = True
        Exit Function
Err_Handler:
        SqlConnectionclass.ExecuteNonQuery("INSERT INTO INV_ERR_PRIMARYKEY(ERR_NUMBER,ERR_DESCRIPTION,TMPINVOICENO,INVOICE_NO,UNIT_CODE,FUNCTIONNAME ) VALUES('" & Err.Number & "','" & Replace(Err.Description, "'", "") & "','" & Ctlinvoice.Text & "','" & mInvNo & "','" & gstrUNITID & "','Updateinsales_dtl')")
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
        UpdateinSale_Dtl = False
    End Function
    Private Sub ShowCode_Desc(ByVal pstrQuery As String, ByRef pctlCode As System.Windows.Forms.TextBox, Optional ByRef pctlDesc As System.Windows.Forms.Label = Nothing)
        '--------------------------------------------------------------------------------------
        'Name       :   ShowCode_Desc
        'Type       :   Sub
        'Author     :   tapanjain
        'Arguments  :   Query(string),Code(Text Box),Description(Label)
        'Return     :   None
        'Purpose    :   Show Code and Description window and set focus on code
        '---------------------------------------------------------------------------------------
        Dim varHelp() As String
        On Error GoTo ErrHandler
        With ctlHelp
            .CreateDSN(gstrCONNECTIONSERVER, gstrCONNECTIONDESCRIPTION, gstrCONNECTIONDSN, gstrCONNECTIONDATABASE)
            .ConnectAsUser = gstrCONNECTIONUSER
            .ConnectThroughDSN = gstrCONNECTIONDSN
            .ConnectWithPWD = gstrCONNECTIONPASSWORD
        End With
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        varHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, pstrQuery)
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        If (UBound(varHelp)) <> -1 Then
            If varHelp(0) <> "0" Then
                pctlCode.Text = Trim(varHelp(0))
                If Not (pctlDesc Is Nothing) Then
                    pctlDesc.Text = Trim(varHelp(1))
                End If
                pctlCode.Focus()
            Else
                Msgbox("No Record Available", MsgBoxStyle.Information, "empower")
            End If
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Function GenerateInvoiceNo(ByVal pstrInvoiceType As String, ByRef pstrInvoiceSubType As String, ByVal pstrRequiredDate As String, ByVal strCustomerCode As String) As String
        On Error GoTo ErrHandler
        Dim strCheckDOcNo As String 'Gets the Doc Number from Back End
        Dim strTempSeries As String 'Find the Numeric series in Doc No
        Dim strSuffix As String 'Generate a NEW Series
        Dim strZeroSuffix As String
        Dim strFin_Start_Date As String
        Dim strFin_End_Date As String
        Dim strSql As String 'String SQL Query
        Dim strUPDATESql As String 'String SQL Query
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        Dim objRs As New ClsResultSetDB_Invoice
        Dim objRs1 As New ClsResultSetDB_Invoice
        Dim strInvoiceDate As String
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


            strSql = "Select Current_No,Suffix,Fin_start_date,Fin_end_Date,ISNULL(CURRENT_NO_TRF_SAMEGSTIN,0) CURRENT_NO_TRF From saleConf Where "
            strSql = strSql & "Invoice_Type ='" & pstrInvoiceType & "' and UNIT_CODE = '" & gstrUNITID & "' and  sub_type='" & pstrInvoiceSubType & "' " &
                     " AND Location_Code ='" & Trim(txtUnitCode.Text) & "' and '" &
                     getDateForDB(pstrRequiredDate) & "' between fin_start_date and fin_end_date"

            objRs.GetResult(strSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If objRs.GetNoRows > 0 Then
                If pstrInvoiceType.ToUpper() = "TRF" Or pstrInvoiceType.ToUpper() = "ITD" Then
                    If IsGSTINSAME(strCustomerCode) Then
                        strCheckDOcNo = objRs.GetValue("CURRENT_NO_TRF")
                    Else
                        strCheckDOcNo = objRs.GetValue("Current_No")
                    End If
                Else
                    strCheckDOcNo = objRs.GetValue("Current_No")
                End If
                strSuffix = objRs.GetValue("suffix")
                strFin_Start_Date = VB6.Format(objRs.GetValue("Fin_Start_Date"), gstrDateFormat)
                strFin_End_Date = VB6.Format(objRs.GetValue("Fin_End_Date"), gstrDateFormat)
            Else
                Err.Raise(vbObjectError + 20008, "[GenerateDocNo]", "Incorrect Parameters Passed Invoice Number cannot be Generated.")
            End If
            objRs.ResultSetClose()
            objRs = Nothing
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

            '    strSql = "select Invoice_Date from Saleschallan_dtl where Doc_No = " & Me.Ctlinvoice.Text
            '    strSql = strSql & " and Invoice_type = '" & mInvType & "'  and  sub_category =  '" & mSubCat & "' and Location_Code='" & Trim(txtUnitCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "'"
            '    objRs1.GetResult(strSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            '    If objRs1.GetNoRows > 0 Then
            '        strInvoiceDate = VB6.Format(objRs1.GetValue("Invoice_Date"), gstrDateFormat)
            '    End If
            '    objRs1.ResultSetClose()
            '    objRs1 = Nothing
            '    If mblnEOUUnit = True Then
            '        If UCase(lbldescription.Text) <> "EXP" Then
            '            If Not mblnSameSeries Then
            '                salesconf = "update saleconf set current_No = " & mSaleConfNo & ", OpenningBal = openningBal - " & mAssessableValue & " where  UNIT_CODE = '" & gstrUNITID & "' and Invoice_type <> 'EXP' and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
            '            Else
            '                salesconf = "update saleconf set current_No = " & mSaleConfNo & " where  UNIT_CODE = '" & gstrUNITID & "' and Single_Series = 1 and Invoice_type <> 'EXP' and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0" & vbCrLf
            '                salesconf = salesconf & "update saleconf set OpenningBal = openningBal - " & mAssessableValue & " where UNIT_CODE = '" & gstrUNITID & "' and  Invoice_type <> 'EXP' and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
            '            End If
            '        Else
            '            salesconf = "update saleconf set current_No = " & mSaleConfNo & " where  UNIT_CODE = '" & gstrUNITID & "' and Invoice_type = 'EXP' and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
            '        End If
            '    Else

            '        If DataExist("SELECT TOP 1 1 FROM SALES_PARAMETER  WHERE SINGLE_INVOICE_SERIES= 1 and UNIT_CODE='" + gstrUNITID + "'") Then
            '            If Not mblnSameSeries Then
            '                salesconf = "update saleconf set current_No = " & mSaleConfNo & " WHERE UNIT_CODE IN( SELECT UNIT_CODE FROM SALES_PARAMETER WHERE SINGLE_INVOICE_SERIES=1)  AND  Invoice_type = '" & Me.lbldescription.Text & "' " & " and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
            '            Else
            '                salesconf = "update saleconf set current_No = " & mSaleConfNo & " WHERE  UNIT_CODE IN( SELECT UNIT_CODE FROM SALES_PARAMETER WHERE SINGLE_INVOICE_SERIES=1)  AND Single_Series = 1 " & " and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
            '            End If
            '        Else
            '            If Not mblnSameSeries Then
            '                salesconf = "update saleconf set current_No = " & mSaleConfNo & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_type = '" & Me.lbldescription.Text & "' and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
            '            Else
            '                salesconf = "update saleconf set current_No = " & mSaleConfNo & " WHERE UNIT_CODE='" + gstrUNITID + "' AND  Single_Series = 1 and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
            '            End If
            '        End If
            '    End If
            'mP_Connection.Execute("INSERT INTO INV_ERROR_DTL(QUERY,UNIT_CODE) VALUES('" & Replace(salesconf, "'", "") & "','" & gstrUNITID & "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            'mP_Connection.Execute(salesconf, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            '101188073 Start
            If gblnGSTUnit Then
                If Len(GSTUnitPrefixCode) > 0 Then
                    strTempSeries = GSTUnitPrefixCode & strTempSeries
                End If
            End If
            '101188073 End
            GenerateInvoiceNo = strTempSeries
        End If
        Exit Function
ErrHandler:
        objRs = Nothing
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function InvAgstBarCode() As Boolean
        '----------------------------------------------------------------------------
        'Author         :   Manoj Kr. Vaish
        'Function       :   Get the BarCodefor Invoice from sales_parameter
        'Comments       :   Date: 04 Feb 2008 ,Issue Id: 22303
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim strQry As String
        Dim Rs As ClsResultSetDB_Invoice
        InvAgstBarCode = False
        strQry = "Select isnull(BarCodeTrackingInInvoice,0) as BarCodeTrackingInInvoice from sales_parameter WHERE  UNIT_CODE = '" & gstrUNITID & "'"
        Rs = New ClsResultSetDB_Invoice
        If Rs.GetResult(strQry) = False Then GoTo ErrHandler
        If Rs.GetValue("BarCodeTrackingInInvoice") = "True" Then
            strQry = "Select isnull(a.BarcodeTrackingAllowed,0) as BarcodeTrackingAllowed"
            strQry = strQry & " from SaleConf (nolock) a,SalesChallan_Dtl b where a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND Doc_No ='" & Trim(Ctlinvoice.Text) & "'"
            strQry = strQry & " and a.Invoice_Type = b.Invoice_type and a.Sub_type = b.Sub_Category and"
            strQry = strQry & " a.Location_Code = b.Location_Code And (Fin_Start_Date <= getDate() And Fin_End_Date >= getDate())"
            Call Rs.GetResult(strQry, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If Rs.GetNoRows > 0 Then
                If Rs.GetValue("BarcodeTrackingAllowed") = "True" Then
                    InvAgstBarCode = True
                Else
                    InvAgstBarCode = False
                End If
            End If
        End If
        Rs.ResultSetClose()
        Rs = Nothing
        Exit Function
ErrHandler:
        Rs = Nothing
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Public Function BarCodeTracking(ByVal pstrInvNo As String, ByVal pstrMode As String) As Boolean
        '----------------------------------------------------------------------------
        'Author         :   Manoj Kr. Vaish
        'Argument       :   Invoice Numbers.
        'Return Value   :   True or False
        'Function       :   Update Bar_BondedStock while invoice editing,deleting & Locking
        'Comments       :   Date: 04 Feb 2008 ,Issue Id: 22303
        'Revised By     :   Manoj Kr Vaish
        'Revision Date  :   28 Nov 2008 Issue ID : eMpro-20090209-27201
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
        Select Case pstrMode
            Case "LOCK"
                If UCase(Trim(CmbCategory.Text)) = "RAW MATERIAL" Or UCase(Trim(CmbCategory.Text)) = "INPUTS" Or UCase(Trim(CmbCategory.Text)) = "COMPONENTS" Then

                    strSql = "select A.item_code,B.Itm_itemalias,isnull(Sales_Quantity,0) as Sales_Qty"
                    strSql = strSql & " from sales_dtl A Inner join Item_mst B (NOLOCK) on a.UNIT_CODE = b.UNIT_CODE  AND A.item_code=B.item_code "
                    strSql = strSql & " where B.Barcode_tracking=1 and A.doc_no='" & Trim(pstrInvNo) & "' and a.UNIT_CODE = '" & gstrUNITID & "'"
                    rsGetQty.GetResult(strSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                    If rsGetQty.GetNoRows > 0 Then
                        rsGetQty.MoveFirst()
                        Do While Not rsGetQty.EOFRecord

                            strSql = "select Isnull(Sum(Convert(numeric(16,4),Issue_Qty)),0) as Issue_Qty from Bar_Invoice_Issue "
                            strSql = strSql & " where  UNIT_CODE = '" & gstrUNITID & "' and Issue_misno='" & Trim(pstrInvNo) & "' and invoice_status is null and substring(Issue_partBarCode,1,8)='" & rsGetQty.GetValue("Itm_itemalias") & "'"
                            rsGetBondedQty.GetResult(strSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                            If rsGetBondedQty.GetNoRows > 0 Then
                                If rsGetQty.GetValue("Sales_Qty") = rsGetBondedQty.GetValue("Issue_Qty") Then
                                    blnQuantitymatch = True
                                Else
                                    Msgbox("Issued Quantity is less than Invoice Quantity.", vbInformation, ResolveResString(100))
                                    mblnQuantityCheck = False
                                    Exit Function
                                End If
                                rsGetQty.MoveNext()
                            Else
                                Msgbox("No Items are issued for this invoice.", vbInformation, ResolveResString(100))
                                mblnQuantityCheck = False
                                Exit Function
                            End If
                        Loop
                    End If
                    mstrupdateBarBondedStockQty = ""
                    mstrupdateBarBondedStockFlag = ""
                    If blnQuantitymatch = True Then
                        'mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & "Insert into Bar_Issue(Issue_No,Issue_UserId,Issue_Unit,Issue_CurrDtTm,Issue_PTId,Issue_UId,"
                        'mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & "Issue_PartBarcode,Issue_MISNo,Issue_LocNo,Issue_Qty,Issue_Date,Issue_Time,Doc_Type,Location_code,UNIT_CODE)"

                        'mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & "Select Issue_No,Issue_UserId,Issue_Unit,Issue_CurrDtTm,Issue_PTId,Issue_UId,Issue_PartBarcode,"
                        'mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & "Issue_MISNo,Issue_LocNo,Issue_Qty,Issue_Date,Issue_Time,Doc_Type,Location_code,UNIT_CODE "
                        'mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & "from Bar_Invoice_Issue where  UNIT_CODE = '" & gstrUNITID & "' and Issue_MISNo='" & Trim(pstrInvNo) & "' and Invoice_status is null" & vbCrLf

                        'strSql = "select A.CRef_PacketNo,isnull(sum(A.CRef_BalQty),0)as BarQuantity,Isnull(sum(Convert(numeric(16,4),Issue_Qty)),0)as SalesQuantity "
                        'strSql = strSql & "from Bar_CrossReference A,Bar_Invoice_Issue B where A.UNIT_CODE = B.UNIT_CODE and A.UNIT_CODE = '" & gstrUNITID & "' AND A.CRef_PacketNo=substring(B.Issue_PartbarCode,9,len(CRef_PacketNo)) and "
                        'strSql = strSql & "A.CRef_PartCode=substring(B.Issue_partBarCode,1,8) and A.CRef_Stage='B' and B.Issue_MISNO='" & Trim(pstrInvNo) & "' and Invoice_status is null group by A.CRef_PacketNo"
                        'rsGetQty.GetResult(strSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        'If rsGetQty.GetNoRows > 0 Then
                        '    rsGetQty.MoveFirst()
                        '    Do While Not rsGetQty.EOFRecord
                        '        If rsGetQty.GetValue("BarQuantity") - rsGetQty.GetValue("SalesQuantity") > 0 Then
                        '            mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & "update A set A.CRef_BalQty=A.CRef_BalQty-" & rsGetQty.GetValue("SalesQuantity") & ""
                        '            mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & " from Bar_CrossReference A,Bar_Invoice_Issue B"
                        '            mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & " Where A.UNIT_CODE = B.UNIT_CODE and A.UNIT_CODE = '" & gstrUNITID & "' AND A.CRef_PartCode=substring(B.Issue_partBarCode,1,8) and A.CRef_Stage='B'"
                        '            mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & " and B.Issue_MISNO='" & Trim(pstrInvNo) & "' and A.CRef_PacketNo='" & rsGetQty.GetValue("CRef_PacketNo") & "'" & vbCrLf
                        '        ElseIf rsGetQty.GetValue("BarQuantity") - rsGetQty.GetValue("SalesQuantity") = 0 Then

                        '            mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & "update A set A.CRef_BalQty=A.CRef_BalQty-" & rsGetQty.GetValue("SalesQuantity") & ",A.CRef_Stage='I'"
                        '            mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & " from Bar_CrossReference A,Bar_Invoice_Issue B"
                        '            mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & " Where A.UNIT_CODE = B.UNIT_CODE and A.UNIT_CODE = '" & gstrUNITID & "' AND  A.CRef_PartCode=substring(B.Issue_partBarCode,1,8) and A.CRef_Stage='B'"
                        '            mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & " and B.Issue_MISNO='" & Trim(pstrInvNo) & "' and A.CRef_PacketNo='" & rsGetQty.GetValue("CRef_PacketNo") & "'" & vbCrLf
                        '        ElseIf rsGetQty.GetValue("BarQuantity") - rsGetQty.GetValue("SalesQuantity") < 0 Then
                        '            MsgBox("Quantity is not available in this Packet [" & rsGetQty.GetValue("CRef_PacketNo") & "] against issued quantity.", vbInformation, ResolveResString(100))
                        '            mblnQuantityCheck = False
                        '            Exit Function
                        '        End If
                        '        rsGetQty.MoveNext()
                        '    Loop
                        '    mstrupdateBarBondedStockFlag = mstrupdateBarBondedStockFlag & "Update Bar_Invoice_Issue Set Issue_Misno='" & mInvNo & "',Invoice_Status=1 where Issue_MisNo='" & Trim(pstrInvNo) & "' and UNIT_CODE = '" & gstrUNITID & "'" & vbCrLf
                        '    mstrupdateBarBondedStockFlag = mstrupdateBarBondedStockFlag & "Update Bar_Issue Set Issue_Misno='" & mInvNo & "' where Issue_MisNo='" & Trim(pstrInvNo) & "' and UNIT_CODE = '" & gstrUNITID & "'" & vbCrLf
                        '    BarCodeTracking = True
                        'End If

                        If checkBARcrossrefence_Invoicequantity(Ctlinvoice.Text) = False Then
                            mblnQuantityCheck = False
                            Exit Function
                        Else
                            BarCodeTracking = True
                        End If

                    Else
                        BarCodeTracking = False
                    End If
                    rsGetBondedQty = Nothing
                    rsGetQty = Nothing
                Else
                    '**************************Check Picked Quantity Against Invocie Quantity********************************
                    strSql = "select A.item_code,B.Itm_itemalias,isnull(Sales_Quantity,0) as Sales_Qty"
                    strSql = strSql & " from sales_dtl A Inner join Item_mst B (NOLOCK) on A.UNIT_CODE = B.UNIT_CODE  AND A.item_code=B.item_code"
                    strSql = strSql & " where B.Barcode_tracking=1 and A.doc_no='" & Trim(pstrInvNo) & "' and A.UNIT_CODE = '" & gstrUNITID & "'"
                    rsGetQty.GetResult(strSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                    If rsGetQty.GetNoRows > 0 Then
                        rsGetQty.MoveFirst()
                        Do While Not rsGetQty.EOFRecord
                            strSql = "select isnull(sum(Quantity),0) as BondedStock_Qty from bar_BondedStock_Dtl "
                            strSql = strSql & " where  UNIT_CODE = '" & gstrUNITID & "' and invoice_no='" & Trim(pstrInvNo) & "' and Status_Flag='W' and item_alias='" & rsGetQty.GetValue("Itm_itemalias") & "'"
                            rsGetBondedQty.GetResult(strSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                            If CDbl(rsGetQty.GetValue("Sales_Qty")) = CDbl(rsGetBondedQty.GetValue("BondedStock_Qty")) Then
                                blnQuantitymatch = True
                            Else
                                Msgbox("Picked Quantity is less than Invoice Quantity.", MsgBoxStyle.Information, ResolveResString(100))
                                mblnQuantityCheck = False
                                Exit Function
                            End If
                            rsGetQty.MoveNext()
                        Loop
                    End If
                    '******************************Update bar Bonded Stock**********************************
                    mstrupdateBarBondedStockQty = ""
                    If blnQuantitymatch = True Then
                        strSql = "select B.Box_label,isnull(sum(A.Quantity),0)as BarQuantity,isnull(sum(B.Quantity),0)as SalesQuantity from Bar_BondedStock A,Bar_BondedStock_Dtl B where A.UNIT_CODE = B.UNIT_CODE and A.UNIT_CODE = '" & gstrUNITID & "' AND"
                        strSql = strSql & " A.Box_Label=B.Box_label and A.Status='B' and B.Status_Flag='W' and B.Invoice_No='" & Trim(pstrInvNo) & "' Group By B.Box_label"
                        rsGetQty.GetResult(strSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        If rsGetQty.GetNoRows > 0 Then
                            rsGetQty.MoveFirst()
                            Do While Not rsGetQty.EOFRecord
                                If rsGetQty.GetValue("BarQuantity") - rsGetQty.GetValue("SalesQuantity") > 0 Then
                                    mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & "update A set A.Quantity=A.Quantity-" & rsGetQty.GetValue("SalesQuantity") & ""
                                    mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & " from bar_BondedStock A,bar_BondedStock_Dtl B"
                                    mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & " Where A.UNIT_CODE = B.UNIT_CODE and A.UNIT_CODE = '" & gstrUNITID & "' AND  A.Box_label = B.Box_label and A.Status='B' and B.Status_Flag='W'"
                                    mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & " and B.Invoice_no='" & Trim(pstrInvNo) & "' and A.Box_label='" & rsGetQty.GetValue("Box_label") & "'" & vbCrLf

                                ElseIf rsGetQty.GetValue("BarQuantity") - rsGetQty.GetValue("SalesQuantity") = 0 Then
                                    mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & "update A set A.Quantity=A.Quantity-" & rsGetQty.GetValue("SalesQuantity") & ",A.Status='I'"
                                    mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & " from bar_BondedStock A,bar_BondedStock_Dtl B"
                                    mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & " Where A.UNIT_CODE = B.UNIT_CODE and A.UNIT_CODE = '" & gstrUNITID & "' AND  A.Box_label = B.Box_label and A.Status='B' and B.Status_Flag='W'"
                                    mstrupdateBarBondedStockQty = mstrupdateBarBondedStockQty & " and B.Invoice_no='" & Trim(pstrInvNo) & "' and A.Box_label='" & rsGetQty.GetValue("Box_label") & "'" & vbCrLf
                                ElseIf rsGetQty.GetValue("BarQuantity") - rsGetQty.GetValue("SalesQuantity") < 0 Then
                                    Msgbox("Quantity is not available in this Box [" & rsGetQty.GetValue("Box_label") & "] against picked quantity.", MsgBoxStyle.Information, ResolveResString(100))
                                    mblnQuantityCheck = False
                                    Exit Function
                                End If
                                rsGetQty.MoveNext()
                            Loop
                            mstrupdateBarBondedStockFlag = "Update bar_BondedStock_Dtl set Status_Flag='L',Invoice_no='" & Trim(CStr(mInvNo)) & "' where Invoice_no='" & Trim(pstrInvNo) & "' and Status_Flag='W' and UNIT_CODE = '" & gstrUNITID & "'"
                            BarCodeTracking = True
                        End If
                    Else
                        BarCodeTracking = False
                    End If
                    rsGetBondedQty = Nothing
                    rsGetQty = Nothing
                End If
        End Select
        Exit Function
ErrHandler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
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
        ' Rs.Open(strField, mP_Connection, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText)
        Rs.Open(strField, mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
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
                Msgbox(oCmd.Parameters(oCmd.Parameters.Count - 1).Value, MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100))
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
                .Parameters.Append(.CreateParameter("@TEMP_INV_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, pstrInvNo))
                .Parameters.Append(.CreateParameter("@INV_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, Trim(mInvNo)))
                .Parameters.Append(.CreateParameter("@Msg", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamOutput, 1000))
                .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End With

            If oCmd.Parameters(oCmd.Parameters.Count - 1).Value <> "" Then
                updateBARcrossrefence_Invoicequantity = False
                Msgbox(oCmd.Parameters(oCmd.Parameters.Count - 1).Value, MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100))
                oCmd = Nothing
                Exit Function
            End If
            oCmd = Nothing

            updateBARcrossrefence_Invoicequantity = True

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Sub Cmdinvoice_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.UCfraRepCmd.ButtonClickEventArgs) Handles Cmdinvoice.ButtonClick

        '-----------------KILL PDF PRINTING PROCESS
        Try
            Dim aProcess As System.Diagnostics.Process
            aProcess = System.Diagnostics.Process.GetProcessById(pdfPrintProcID)
            If aProcess.HasExited = False Then
                aProcess.Kill()
            End If
        Catch ex As Exception
        End Try
        '-----------------KILL PDF PRINTING PROCESS

        Dim rsSalesConf As ClsResultSetDB_Invoice
        Dim rssaledtl As ClsResultSetDB_Invoice
        Dim rsItembal As ClsResultSetDB_Invoice
        Dim rsSalesParameter As ClsResultSetDB_Invoice
        Dim rsbom As ClsResultSetDB_Invoice
        Dim strSalesconf As String
        Dim ItemCode As String
        Dim strDrgNo As String
        Dim strAccountCode As String
        Dim StrAmendmentNo As String
        Dim SALEDTL As String
        Dim intRow As Short
        Dim intloopcount As Short
        Dim salesQuantity As Double
        Dim intNoCopies As Short
        Dim strRetval As String
        Dim objDrCr As prj_DrCrNote.cls_DrCrNote
        Dim strInvoiceDate As String
        Dim irowcount As Short
        Dim intRwCount1 As Short
        Dim varItemQty1 As Double
        Dim strToolCode As String
        Dim blnBatchTrack As Boolean
        Dim strBatchQuery As String
        Dim dblToolCost As Double
        Dim blnCheckToolCost As Boolean
        Dim strItembal As String
        Dim strtoolQuantity As String
        Dim ObjBarcodeHMI As New Prj_BCHMI.cls_BCHMI(gstrUNITID)
        'Dim ObjBarcodeHMI As New prj_bchm cls_BCHMI(gstrUNITID)
        Dim strBarcodeMsg As String
        Dim strpath As String
        Dim rsBatchMst As ClsResultSetDB_Invoice
        Dim rsBatch As ClsResultSetDB_Invoice
        Dim intTotalAmountValue As Double
        Dim intmaxsubreportloop, intsubreportloopcounter As Integer
        Dim objCom As New SqlCommand()
        'Dim RdAddSold As New ReportDocument
        'Dim Frm As New eMProCrystalReportViewer
        Dim Frm As Object = Nothing
        Dim RdAddSold As ReportDocument
        Dim mstrInvRejSQL As String
        Dim blnInvoiceAgainstMultipleSO As Boolean
        Dim blnDSTracking As Boolean
        Dim intTotalNoofitemsinInvoices As Integer
        Dim rsGENERATEBARCODE As ClsResultSetDB_Invoice
        Dim strBARCODEtype As String
        '10564189
        Dim blnAllow_ASNFlag As Boolean

        'sattu
        Dim strPrintMethod As String = ""
        Dim strSQL As String = ""
        Dim intTotalNoofSlabs As Integer = 0
        '10812364
        Dim strBarcodeMsg_paratemeter As String
        '10812364
        '10564189
        'RdAddSold = Frm.GetReportDocument()
        Dim ts As Object
        Dim strGateEntryBarCode As String
        Dim strBinningBarCode As String
        Dim strGateentrypath As String
        Dim strbinpath As String
        Dim fso As New Scripting.FileSystemObject
        Dim blnIsPDFExported As Boolean = False
        Dim STRASNTYPE As String
        Dim STRCUSTTYPE As String
        Dim strASNFILEPATH As String

        strBarcodeMsg = ""
        strBarcodeMsg_paratemeter = ""
        '10564189 value initalize 
        blnAllow_ASNFlag = False
        '10564189

        Try

            If UCase(Trim(gstrUNITID)) = "MST" Then
                Frm = New eMProCrystalReportViewer
            Else
                Frm = New eMProCrystalReportViewer_Inv
            End If

            If e.Button = UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE Then
                Me.Close()
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
                'Call CmbCategory_Validating(CmbCategory, New System.ComponentModel.CancelEventArgs(False))
            End If
            SALEDTL = "select * from Saleschallan_Dtl where Doc_No =" & Me.Ctlinvoice.Text & "  and Location_Code='" & Trim(txtUnitCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "'"
            rssaledtl = New ClsResultSetDB_Invoice
            rssaledtl.GetResult(SALEDTL, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            strAccountCode = rssaledtl.GetValue("Account_code")
            mstracountcode = strAccountCode
            strCustRef = rssaledtl.GetValue("Cust_ref")
            StrAmendmentNo = rssaledtl.GetValue("Amendment_No")
            intTotalAmountValue = rssaledtl.GetValue("Total_Amount")
            strInvoiceDate = VB6.Format(rssaledtl.GetValue("Invoice_Date"), gstrDateFormat)
            blnInvoiceAgainstMultipleSO = rssaledtl.GetValue("InvoiceAgainstMultipleSO")
            rssaledtl.ResultSetClose()
            '10617093
            mblncustomerlevel_A4report_functionlity = False
            mblncustomerlevel_Annexture_printing = False

            If AllowA4Reports(strAccountCode) = True Then
                mblncustomerlevel_A4report_functionlity = True
            End If
            If ALLOW_ANNEXTUREPRINTING(strAccountCode) = True Then
                mblncustomerlevel_Annexture_printing = True
            End If
            mblncustomerspecificreport = False
            mblnAllowCustomerSpecificReport_COMP = False
            mblnAllowCustomerSpecificReport_COMP = SqlConnectionclass.ExecuteScalar("Select isnull(AllowCustomerSpecificReport_COMP,0) from customer_mst (Nolock) where customer_code='" & strAccountCode & "' and unit_code='" + gstrUNITID + "'")

            mblnAllowCustomerSpecificReport_RAW = False
            mblnAllowCustomerSpecificReport_RAW = SqlConnectionclass.ExecuteScalar("Select isnull(AllowCustomerSpecificReport_RAW,0) from customer_mst (Nolock) where customer_code='" & strAccountCode & "' and unit_code='" + gstrUNITID + "'")

            If (UCase(cmbInvType.Text) = "NORMAL INVOICE" And UCase(CmbCategory.Text) = "FINISHED GOODS") Or (UCase(cmbInvType.Text) = "EXPORT INVOICE") Then
                If AllowCustomerspecificreport(strAccountCode) = True Then
                    mblncustomerspecificreport = True
                End If
            End If

            '10617093
            If AllowASNPrinting(strAccountCode) = True And ValidASNFilePath() = False Then
                Msgbox("Invalid ASN Path in Mind.cfg ", vbInformation, ResolveResString(100))
                Exit Sub
            End If

            '01 jan 2024
            If UCase(cmbInvType.Text) = "NORMAL INVOICE" And (UCase(CmbCategory.Text) = "ASSETS" Or UCase(CmbCategory.Text) = "COMPONENTS" Or UCase(CmbCategory.Text) = "RAW MATERIAL") Then
                If GetPrintMethod(strAccountCode).ToUpper() = "TATA" Then
                    If AllowCustomerspecificreport(strAccountCode) = True Then
                        mblncustomerspecificreport = True
                    End If
                End If
                If mblnAllowCustomerSpecificReport_COMP = True And AllowCustomerspecificreport(strAccountCode) = True Then
                    mblncustomerspecificreport = True
                End If
                If mblnAllowCustomerSpecificReport_RAW = True And AllowCustomerspecificreport(strAccountCode) = True Then
                    mblncustomerspecificreport = True
                End If

            End If

            '01 jan 2024
            strSalesconf = "Select UpdatePO_Flag,UpdateStock_Flag,Stock_Location,OpenningBal,Preprinted_Flag,NoCopies,isnull(BatchTrackingAllowed,0) as BatchTrackingAllowed , ALLOW_ASNFLAG , AllowA4Reports , Noofcopies_A4report ,Reqd_Shipping_Address  from saleconf (nolock) where  UNIT_CODE = '" & gstrUNITID & "' and "
            strSalesconf = strSalesconf & "Invoice_type = '" & Me.lbldescription.Text & "' and sub_type = '"
            strSalesconf = strSalesconf & Me.lblcategory.Text & "' and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
            rsSalesConf = New ClsResultSetDB_Invoice
            rsSalesConf.GetResult(strSalesconf, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            If rsSalesConf.GetNoRows > 0 Then

                updatePOflag = rsSalesConf.GetValue("UpdatePO_Flag")
                updatestockflag = rsSalesConf.GetValue("UpdateStock_Flag")
                strStockLocation = rsSalesConf.GetValue("Stock_Location")
                mOpeeningBalance = Val(rsSalesConf.GetValue("OpenningBal"))
                intNoCopies = Val(rsSalesConf.GetValue("NoCopies"))
                gstrIntNoCopies = intNoCopies
                blnBatchTrack = rsSalesConf.GetValue("BatchTrackingAllowed")
                '10564189 
                blnAllow_ASNFlag = rsSalesConf.GetValue("Allow_ASNFlag")
                '10564189 
                mblnA4reports_invoicewise = rsSalesConf.GetValue("AllowA4Reports")
                intNoCopies_A4reports = Val(rsSalesConf.GetValue("Noofcopies_A4report"))
                '10688760
                mbln_SHIPPING_ADDRESS_INVOICEWISE = rsSalesConf.GetValue("Reqd_Shipping_Address")
                '10688760
            Else
                Msgbox("Please Define Stock Location in Sales Configuration. ")
                Exit Sub
            End If
            rsSalesConf.ResultSetClose()
            rsSalesConf = Nothing
            If Len(Trim(strStockLocation)) = 0 Then
                Msgbox("Please Define Stock Location in Sales Configuration. ")
                Exit Sub
            End If
            rsSalesParameter = New ClsResultSetDB_Invoice
            rsSalesParameter.GetResult("Select Batch_Tracking = Isnull(Batch_Tracking,0),CheckToolAmortisation,isnull(HMIBarcodePath,'') as HMIBarcodePath from Sales_Parameter (NOLOCK) WHERE UNIT_CODE = '" & gstrUNITID & "' ")
            If rsSalesParameter.GetNoRows > 0 Then
                rsSalesParameter.MoveFirst()
                If Len(Trim(rsSalesParameter.GetValue("CheckToolAmortisation"))) = 0 Then
                    Msgbox("First define Check Tool Amortisation in Sales Parameter", MsgBoxStyle.Information, "eMPro")
                    rsSalesParameter.ResultSetClose()
                    Exit Sub
                End If
                blnCheckToolCost = rsSalesParameter.GetValue("CheckToolAmortisation")
                strpath = rsSalesParameter.GetValue("HMIBarcodePath")
            Else
                Msgbox("No Data Defined in Sales Parameter", MsgBoxStyle.Information, "eMPro")
                rsSalesParameter.ResultSetClose()
                Exit Sub
            End If
            '15 jan 2024
            If optInvYes(0).Checked = True Then

                SALEDTL = "Select sum(Sales_Quantity)SALES_QUANTITY,Item_code from sales_Dtl where Doc_No = " & Me.Ctlinvoice.Text & " and Location_Code='" & Trim(txtUnitCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "' GROUP BY ITEM_CODE "
                rssaledtl = New ClsResultSetDB_Invoice
                rssaledtl.GetResult(SALEDTL, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                intRow = rssaledtl.GetNoRows
                rssaledtl.MoveFirst()

                For intloopcount = 1 To intRow
                    ItemCode = rssaledtl.GetValue("Item_code")
                    salesQuantity = rssaledtl.GetValue("SALES_QUANTITY")
                    If Not (((UCase(Trim(cmbInvType.Text)) = "SERVICE INVOICE") And mblnServiceInvoicemate = True) Or (gstrUNITID = "STH")) Then
                        rsItembal = New ClsResultSetDB_Invoice
                        rsItembal.GetResult("Select Cur_bal from Itembal_Mst where Item_code = '" & ItemCode & "'and Location_code ='" & strStockLocation & "' and UNIT_CODE = '" & gstrUNITID & "' ", ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                        If rsItembal.GetNoRows > 0 Then
                            If salesQuantity > rsItembal.GetValue("Cur_Bal") Then
                                Msgbox("Balance for item " & ItemCode & " at Location " & strStockLocation & " not available. ", MsgBoxStyle.Information, "Empower")
                                rsItembal.ResultSetClose()
                                Exit Sub
                            End If
                            rsItembal.ResultSetClose()
                        Else
                            Msgbox("No Item in ItemMaster for Location " & strStockLocation & ".", MsgBoxStyle.OkOnly, "empower")
                            rsItembal.ResultSetClose()
                            Exit Sub
                        End If
                    End If
                    rssaledtl.MoveNext()
                Next
            End If
            '15 jan 2024
            SALEDTL = "Select Sales_Quantity,Item_code,Cust_Item_Code,toolcost_amount from sales_Dtl where Doc_No = " & Me.Ctlinvoice.Text & " and Location_Code='" & Trim(txtUnitCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "'"
            rssaledtl = New ClsResultSetDB_Invoice
            rssaledtl.GetResult(SALEDTL, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            intRow = rssaledtl.GetNoRows
            rssaledtl.MoveFirst()
            If optInvYes(0).Checked = True Then
                For intloopcount = 1 To intRow
                    ItemCode = rssaledtl.GetValue("Item_code")
                    salesQuantity = rssaledtl.GetValue("Sales_quantity")
                    strDrgNo = rssaledtl.GetValue("Cust_Item_code")
                    dblToolCost = IIf(IsDBNull(rssaledtl.GetValue("ToolCost_amount")), 0, Val(rssaledtl.GetValue("ToolCost_amount")))
                    '10869290
                    If Not (((UCase(Trim(cmbInvType.Text)) = "SERVICE INVOICE") And mblnServiceInvoicemate = True) Or (gstrUNITID = "STH")) Then
                        '10869290
                        rsItembal = New ClsResultSetDB_Invoice
                        rsItembal.GetResult("Select Cur_bal from Itembal_Mst where Item_code = '" & ItemCode & "'and Location_code ='" & strStockLocation & "' and UNIT_CODE = '" & gstrUNITID & "' ", ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                        If rsItembal.GetNoRows > 0 Then
                            If salesQuantity > rsItembal.GetValue("Cur_Bal") Then
                                Msgbox("Balance for item " & ItemCode & " at Location " & strStockLocation & " not available. ", MsgBoxStyle.Information, "Empower")
                                rsItembal.ResultSetClose()
                                Exit Sub
                            End If
                            rsItembal.ResultSetClose()
                        Else
                            Msgbox("No Item in ItemMaster for Location " & strStockLocation & ".", MsgBoxStyle.OkOnly, "empower")
                            rsItembal.ResultSetClose()
                            Exit Sub
                        End If
                    End If

                    If blnBatchTrack = True And updatestockflag = True And UCase(Trim(cmbInvType.Text)) <> "REJECTION" And UCase(Trim(cmbInvType.Text)) <> "JOBWORK INVOICE" Then
                        rsBatch = New ClsResultSetDB_Invoice
                        Dim strquery As String
                        strquery = "Select Batch_No,Batch_Qty from ItemBatch_Dtl where Doc_Type = 9999 and Doc_no = '" & Trim(Me.Ctlinvoice.Text) & "' and Item_Code = '" & ItemCode & "' and UNIT_CODE = '" & gstrUNITID & "' "
                        Call rsBatch.GetResult(strquery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        If rsBatch.RowCount <= 0 Then
                            Msgbox(" Batch Details are Not Available ", MsgBoxStyle.Information, "eMPro")
                            Exit Sub
                        End If
                        rsBatch.MoveFirst()
                        While Not rsBatch.EOFRecord
                            rsBatchMst = New ClsResultSetDB_Invoice
                            Call rsBatchMst.GetResult("Select Current_batch_Qty = Isnull(Current_batch_Qty,0) From ItemBatch_Mst where Batch_No = '" & rsBatch.GetValue("Batch_No") & "' and Location_Code = '" & strStockLocation & "' and Item_Code = '" & Trim(ItemCode) & "' and UNIT_CODE = '" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                            If rsBatchMst.RowCount > 0 Then
                                If Val(rsBatchMst.GetValue("Current_Batch_Qty")) < rsBatch.GetValue("Batch_Qty") Then
                                    Msgbox("Balance for item " & ItemCode & " at Location " & strStockLocation & " is " & Val(rsBatchMst.GetValue("Current_Batch_Qty")) & " at Batch Master")
                                    rsBatch.ResultSetClose()
                                    rsBatchMst.ResultSetClose()
                                    Exit Sub
                                Else
                                    strBatchQuery = strBatchQuery & "  Update ItemBatch_Mst Set Current_batch_Qty = Current_batch_Qty - " & Val(rsBatch.GetValue("Batch_Qty")) & ",Upd_Userid = '" & mP_User & "' ,Upd_Dt = getdate()  where Batch_No = '" & rsBatch.GetValue("Batch_No") & "' and Location_Code = '" & strStockLocation & "' and Item_Code = '" & Trim(ItemCode) & "' and UNIT_CODE = '" & gstrUNITID & "'"
                                End If
                                rsBatchMst.ResultSetClose()
                            Else
                                Msgbox("Balance for item " & ItemCode & " at Location " & strStockLocation & " not available in Batch Master. ")
                                rsBatch.ResultSetClose()
                                rsBatchMst.ResultSetClose()
                                Exit Sub
                            End If
                            rsBatch.MoveNext()
                        End While
                        rsBatch.ResultSetClose()
                    End If
                    If Len(Trim(strCustRef)) > 0 Then
                        If UCase(cmbInvType.Text) <> "REJECTION" Then
                            rsItembal = New ClsResultSetDB_Invoice
                            rsItembal.GetResult("Select balanceQty = order_qty - despatch_Qty,OpenSO from Cust_ord_dtl where account_code ='" & strAccountCode & "' and Cust_ref ='" & strCustRef & "' and Amendment_No = '" & StrAmendmentNo & "' and Item_code ='" & ItemCode & "' and Cust_drgNo ='" & strDrgNo & "' and Active_flag ='A' and UNIT_CODE = '" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                            If rsItembal.GetNoRows > 0 Then
                                If rsItembal.GetValue("OpenSO") = False Then
                                    If salesQuantity > rsItembal.GetValue("BalanceQty") Then
                                        Msgbox("Balance Quantity in SO for item " & ItemCode & " is " & rsItembal.GetValue("BalanceQty") & ".Check Quantity of Item in Challan.", MsgBoxStyle.Information, "Empower")
                                        rsItembal.ResultSetClose()
                                        Exit Sub
                                    End If
                                    rsItembal.ResultSetClose()
                                End If
                            Else
                                Msgbox("No Item (" & StrItemCode & ") exist in SO - " & strCustRef & ".", MsgBoxStyle.Information, "empower")
                                rsItembal.ResultSetClose()
                                Exit Sub
                            End If
                        End If
                    End If
                    If blnCheckToolCost = True Then
                        strItembal = "select BalanceQty = isnull(a.proj_qty,0) - isnull(a.ClosingValueSMIEL,0),a.Tool_C from Amor_dtl a,Tool_Mst b"
                        strItembal = strItembal & " where a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND  account_code = '" & strAccountCode & "'"
                        strItembal = strItembal & " and Item_code = '" & ItemCode & "' and a.Tool_c = b.tool_c and a.Item_code = b.Product_No order by a.tool_c"
                        rsItembal = New ClsResultSetDB_Invoice
                        rsItembal.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        If rsItembal.GetNoRows > 0 Then
                            rsItembal.MoveFirst()
                            strtoolQuantity = CStr(Val(rsItembal.GetValue("BalanceQty")))
                            strToolCode = rsItembal.GetValue("Tool_c")
                            rsItembal.ResultSetClose()
                            strItembal = "select BalanceQty = sum(isnull(UsedProjQty,0)) from Amor_dtl "
                            strItembal = strItembal & " where "
                            strItembal = strItembal & " Item_code = '" & ItemCode & "' and Tool_c = '" & strToolCode & "' and UNIT_CODE = '" & gstrUNITID & "'"
                            rsItembal = New ClsResultSetDB_Invoice
                            rsItembal.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                            rsItembal.MoveFirst()
                            strtoolQuantity = CStr(CDbl(Val(strtoolQuantity)) - Val(rsItembal.GetValue("BalanceQty")))
                            rsItembal.ResultSetClose()
                            If Val(CStr(salesQuantity)) > Val(strtoolQuantity) Then
                                If Val(strtoolQuantity) = 0 Then
                                    Msgbox("No Balance Available for Item (" & ItemCode & ") and customer Part Code (" & strDrgNo & ") For Amortisation Calculations. ", MsgBoxStyle.OkOnly, "eMPro")
                                Else
                                    Msgbox("Quantity should not be Greater then available Balance Quantity for Amortisarion " & strtoolQuantity, MsgBoxStyle.OkOnly, "eMPro")
                                End If
                                Exit Sub
                            End If
                        Else
                            rsItembal.ResultSetClose()
                            strItembal = "select BalanceQty = isnull(proj_qty,0) - isnull(ClosingValueSMIEL,0) from Amor_dtl"
                            strItembal = strItembal & " where account_code = '" & strAccountCode & "'"
                            strItembal = strItembal & " and Item_code = '" & ItemCode & "' and UNIT_CODE = '" & gstrUNITID & "'"
                            rsItembal = New ClsResultSetDB_Invoice
                            rsItembal.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                            If rsItembal.GetNoRows > 0 Then
                                rsItembal.MoveFirst()
                                strtoolQuantity = CStr(Val(rsItembal.GetValue("BalanceQty")))
                                rsItembal.ResultSetClose()
                                strItembal = "select BalanceQty = sum(isnull(UsedProjQty,0)) from Amor_dtl"
                                strItembal = strItembal & " Where Item_code = '" & ItemCode & "' and UNIT_CODE = '" & gstrUNITID & "'"
                                rsItembal = New ClsResultSetDB_Invoice
                                rsItembal.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                                strtoolQuantity = CStr(CDbl(Val(strtoolQuantity)) - Val(rsItembal.GetValue("BalanceQty")))
                                rsItembal.ResultSetClose()
                                If Val(CStr(salesQuantity)) > Val(strtoolQuantity) Then
                                    If CDbl(Val(strtoolQuantity)) = 0 Then
                                        Msgbox("No Balance Available for Item (" & ItemCode & ") and customer Part Code (" & strDrgNo & ") For Amortisation Calculations. ", MsgBoxStyle.OkOnly, "eMPro")
                                    Else
                                        Msgbox("Quantity should not be Greater then available Balance Quantity for Amortisarion " & strtoolQuantity, MsgBoxStyle.OkOnly, "eMPro")
                                    End If
                                    Exit Sub
                                End If
                            Else
                                rsItembal.ResultSetClose()
                            End If
                        End If
                        With mP_Connection
                            .Execute("DELETE FROM tmpBOM WHERE UNIT_CODE = '" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            .Execute("BOMExplosion '" & Trim(ItemCode) & "','" & Trim(ItemCode) & "',1,0,'" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        End With
                        rsbom = New ClsResultSetDB_Invoice
                        rsbom.GetResult("select * from tmpBOM WHERE UNIT_CODE = '" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        If rsbom.GetNoRows > 0 Then
                            irowcount = rsbom.GetNoRows
                            rsbom.MoveFirst()
                            For intRwCount1 = 1 To irowcount
                                strItembal = "select BalanceQty = isnull(a.proj_qty,0) - isnull(a.ClosingValueSMIEL,0),a.tool_C from Amor_dtl a, tool_mst b "
                                strItembal = strItembal & " where a.UNIT_CODE = b.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' AND account_code = '" & Trim(strAccountCode) & "'"
                                strItembal = strItembal & " and Item_code = '" & rsbom.GetValue("item_code") & "' and a.Tool_c = b.Tool_c and a.ITem_code = b.Product_no order by a.tool_c"
                                rsItembal = New ClsResultSetDB_Invoice
                                rsItembal.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                                If rsItembal.GetNoRows > 0 Then
                                    rsItembal.MoveFirst()
                                    strtoolQuantity = CStr(Val(rsItembal.GetValue("BalanceQty")))
                                    strToolCode = rsItembal.GetValue("Tool_c")
                                    rsItembal.ResultSetClose()
                                    strItembal = "select BalanceQty = sum(isnull(UsedProjQty,0)) from Amor_dtl a"
                                    strItembal = strItembal & " where account_code = '" & Trim(strAccountCode) & "'"
                                    strItembal = strItembal & " and Item_code = '" & rsbom.GetValue("item_code") & "' and a.Tool_c = '" & strToolCode & "' and a.UNIT_CODE = '" & gstrUNITID & "'"
                                    rsItembal = New ClsResultSetDB_Invoice
                                    rsItembal.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                                    varItemQty1 = (salesQuantity * Val(rsbom.GetValue("grossweight")))
                                    strtoolQuantity = CStr(CDbl(Val(strtoolQuantity)) - Val(rsItembal.GetValue("BalanceQty")))
                                    rsItembal.ResultSetClose()
                                    If Val(CStr(varItemQty1)) > Val(strtoolQuantity) Then
                                        If CDbl(Val(strtoolQuantity)) = 0 Then
                                            Msgbox("No Balance Available for Item (" & rsbom.GetValue("item_code") & ") and customer Part Code (" & strDrgNo & ") For Amortisation Calculations. ", MsgBoxStyle.OkOnly, "eMPro")
                                        Else
                                            Msgbox("Quantity should not be Greater then available Balance Quantity for Amortisarion of this Item (" & rsbom.GetValue("item_code") & ")" & strtoolQuantity, MsgBoxStyle.OkOnly, "eMPro")
                                        End If
                                        rsbom.ResultSetClose()
                                        Exit Sub
                                    End If
                                Else
                                    rsItembal.ResultSetClose()
                                End If
                                rsbom.MoveNext()
                            Next
                        Else
                            rsbom.ResultSetClose()
                        End If
                    End If
                    rssaledtl.MoveNext()
                Next
                rssaledtl.ResultSetClose()
                rssaledtl = Nothing
                If UCase(cmbInvType.Text) = "REJECTION" Then
                    If Len(Trim(strCustRef)) < 0 Then
                        If CheckDataFromGrin(CDbl(Val(Trim(strCustRef))), strAccountCode) = False Then
                            Exit Sub
                        End If
                    End If
                End If
                '****
            End If
            Dim intLoopCounter As Short
            Dim intMaxLoop As Short
            '10825102 
            Dim intNoCopies_A4reports_orignial As Short
            Dim intNoCopies_A4reports_REPRINT As Short
            Dim COPYNAME_New(1) As String
            COPYNAME_New(0) = String.Empty
            COPYNAME_New(1) = "Y"

            Dim DT_A4CUSTOMER_INVOICEPRINTINGTAG As DataTable = SqlConnectionclass.GetDataTable("SELECT RTRIM(LTRIM(TEXTHEADING)) TEXTHEADING, HARDCOPYPRINTREQUIRED,SERIALNO, ORIGINAL_REPRINT FROM  A4CUSTOMER_INVOICEPRINTINGTAG (Nolock) WHERE UNIT_CODE='" + gstrUNITID + "' AND CUSTOMER_CODE='" + strAccountCode + "' ORDER BY SERIALNO")
            If mblnISTrueSignRequired = True Then
                If DT_A4CUSTOMER_INVOICEPRINTINGTAG.Rows.Count = 0 Then
                    CustomRollbackTrans()
                    Msgbox("Please Define A4 Customer Invoice Printing Tags for the customer : " + strAccountCode)
                    Exit Sub
                End If
            End If

            Select Case e.Button
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                    mCheckValARana = "CLOSECLICKED"
                    Me.Close()
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT_TO_WINDOW
                    mCheckValARana = "PRINTTOWINDOW"
                    mblnLock_Clicked = False
                    ''Praveen Digital Sign
                    If optInvYes(0).Checked = True Then
                        mblnISCrystalReportRequired = CBool(Find_Value("select dbo.ISCrystalReportRequired_InvoicePrint ( '" + gstrUNITID + "','" & strAccountCode & "','" & lbldescription.Text & "','" & lblcategory.Text & "','O')"))
                    Else
                        mblnISCrystalReportRequired = CBool(Find_Value("select dbo.ISCrystalReportRequired_InvoicePrint ( '" + gstrUNITID + "','" & strAccountCode & "','" & lbldescription.Text & "','" & lblcategory.Text & "','R')"))
                    End If
                    If mblnlorryno = True Then
                        Dim strlorryquery As String
                        strlorryquery = "UPDATE SalesChallan_Dtl SET LORRYNO_DATE= '" & txtlorryno.Text.Trim & "'  WHERE Doc_No=" & Ctlinvoice.Text & " and UNIT_CODE = '" & gstrUNITID & "'  and Location_Code='" & Trim(txtUnitCode.Text) & "'"
                        mP_Connection.Execute(strlorryquery, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End If
                    '102027599
                    If optInvYes(0).Checked = False Then
                        If mblnEwaybill_Print Then
                            Call IRN_QRBarcode()
                        End If
                    End If
                    If InvoiceGeneration(RdAddSold, Frm) = True Then
                        'If AllowASNTextFile(strAccountCode) = True Then
                        '    If (UCase(cmbInvType.Text) <> "REJECTION") Then
                        '        mInvNo = Ctlinvoice.Text
                        '        If ASNTEXTFILE_DETAILS(mInvNo, strAccountCode) = False Then
                        '            CustomRollbackTrans()
                        '            Exit Sub
                        '        Else
                        '            Exit Sub
                        '        End If
                        '    End If
                        'End If

                        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                        If gstrUNITID <> "MST" Then
                            Frm.EnablePrintButton = False
                        End If
                        '---------------------------------------------
                        'Frm.EnablePrintButton = False
                        '--------------------------------------------
                        'If mblncustomerlevel_Annexture_printing = True And mblncustomerlevel_A4report_functionlity = True Then
                        '    If optInvYes(0).Checked = True Then
                        '        Printbarcode_TOYOTA(Ctlinvoice.Text.Trim)
                        '    End If
                        'End If

                        If AllowBarCodePrinting(strAccountCode) = True Then
                            If optInvYes(0).Checked = False And mblnEwaybill_Print = True Then
                                '------------------------------------------------------------------------------------
                                rsGENERATEBARCODE = New ClsResultSetDB_Invoice
                                rsGENERATEBARCODE.GetResult("SELECT PRINT_METHOD FROM CUSTOMER_MST C WHERE C.UNIT_CODE='" & gstrUNITID & "' AND C.CUSTOMER_CODE='" & strAccountCode & "'")
                                strPrintMethod = UCase(rsGENERATEBARCODE.GetValue("PRINT_METHOD").ToString)
                                rsGENERATEBARCODE.ResultSetClose()
                                rsGENERATEBARCODE = Nothing

                                If optInvYes(0).Checked = False And mblnEwaybill_Print = True Then
                                    If strPrintMethod = "MOBIS" Then

                                        strBarcodeMsg = ObjBarcodeHMI.GenerateBarCode_LINELEVEL_SALESORDER_2dbarcode_Mobis(gstrUserMyDocPath, Trim(Ctlinvoice.Text), "MOBIS", "", "", True, Trim(Ctlinvoice.Text), gstrCONNECTIONSTRING)


                                        If VB.Left(strBarcodeMsg, 1) <> "Y" Then
                                            Msgbox("Problem While Generating Barcode Image.", vbInformation, ResolveResString(100))
                                            Exit Sub
                                        Else
                                            If SaveBarCodeImage_singlelevelso_2DBARCODE(Ctlinvoice.Text, gstrUserMyDocPath, Mid(strBarcodeMsg, 3)) = False Then
                                                Msgbox("Problem While saving Barcode Image.", vbInformation, ResolveResString(100))
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

                        Call ReprintQRbarcode()



                        If AllowBarCodePrinting(strAccountCode) = True And blnlinelevelcustomer = False Then
                            rsGENERATEBARCODE = New ClsResultSetDB_Invoice
                            rsGENERATEBARCODE.GetResult("SELECT PRINT_METHOD FROM CUSTOMER_MST C WHERE C.UNIT_CODE='" & gstrUNITID & "' AND C.CUSTOMER_CODE='" & strAccountCode & "'")
                            strPrintMethod = UCase(rsGENERATEBARCODE.GetValue("PRINT_METHOD").ToString)
                            rsGENERATEBARCODE.ResultSetClose()
                            rsGENERATEBARCODE = Nothing

                            If optInvYes(0).Checked = True Then
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
                                                Msgbox("Problem While Generating Barcode Image.", vbInformation, ResolveResString(100))
                                                rsGENERATEBARCODE.ResultSetClose()
                                                rsGENERATEBARCODE = Nothing
                                                Exit Sub
                                            Else
                                                '10812364
                                                strBarcodeMsg_paratemeter = Mid(strBarcodeMsg, 3)
                                                '10812364
                                                If SaveQRBarCodeImage(Trim(Ctlinvoice.Text), Val(rsGENERATEBARCODE.GetValue("FromLineNo").ToString), Val(rsGENERATEBARCODE.GetValue("ToLineNo").ToString), strBarcodeMsg_paratemeter, intRow, intTotalNoofSlabs, "") = False Then
                                                    Msgbox("Problem While saving Barcode Image.", vbInformation, ResolveResString(100))
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
                                    '12 oct 2017'
                                ElseIf strPrintMethod = "TOYOTA_NEW" Then
                                    strSQL = "select * from dbo.UFN_INV_BARCODE('" & gstrUNITID & "'," & Ctlinvoice.Text.Trim & ")"
                                    rsGENERATEBARCODE = New ClsResultSetDB_Invoice
                                    rsGENERATEBARCODE.GetResult(strSQL)
                                    intTotalNoofSlabs = rsGENERATEBARCODE.GetNoRows
                                    If intTotalNoofSlabs > 0 Then
                                        rsGENERATEBARCODE.MoveFirst()
                                        For intRow = 1 To intTotalNoofSlabs
                                            strBarcodeMsg = ObjBarcodeHMI.GenerateQRBarCode_NEWTOYOTA(gstrUserMyDocPath, True, Trim(Ctlinvoice.Text), Trim(Ctlinvoice.Text), Val(rsGENERATEBARCODE.GetValue("FromLineNo").ToString), Val(rsGENERATEBARCODE.GetValue("ToLineNo").ToString), intRow, intTotalNoofSlabs, gstrCONNECTIONSTRING)
                                            If VB.Left(strBarcodeMsg, 1) <> "Y" Then
                                                CustomRollbackTrans()
                                                Msgbox("Problem While Generating Barcode Image.", vbInformation, ResolveResString(100))

                                                rsGENERATEBARCODE.ResultSetClose()
                                                rsGENERATEBARCODE = Nothing
                                                Exit Sub
                                            Else
                                                '10812364
                                                strBarcodeMsg_paratemeter = Mid(strBarcodeMsg, 3)
                                                '10812364
                                                If SaveQRBarCodeImage(Trim(Ctlinvoice.Text), Val(rsGENERATEBARCODE.GetValue("FromLineNo").ToString), Val(rsGENERATEBARCODE.GetValue("ToLineNo").ToString), strBarcodeMsg_paratemeter, intRow, intTotalNoofSlabs, "") = False Then
                                                    CustomRollbackTrans()
                                                    Msgbox("Problem While saving Barcode Image.", vbInformation, ResolveResString(100))

                                                    rsGENERATEBARCODE.ResultSetClose()
                                                    rsGENERATEBARCODE = Nothing
                                                    Exit Sub
                                                End If
                                            End If
                                        Next
                                    End If
                                    rsGENERATEBARCODE.ResultSetClose()
                                    rsGENERATEBARCODE = Nothing
                                    '------------------------------------------------------------------------------------
                                    '12 oct 2017

                                Else

                                    strBarcodeMsg = ObjBarcodeHMI.GenerateBarCode(gstrUserMyDocPath, mInvNo, True, Trim(Ctlinvoice.Text), gstrCONNECTIONSTRING)
                                    If VB.Left(strBarcodeMsg, 1) <> "Y" Then
                                        Msgbox("Problem While Generating Barcode Image.", vbInformation, ResolveResString(100))
                                        Exit Sub
                                    End If
                                    If SaveBarCodeImage(Me.Ctlinvoice.Text, gstrUserMyDocPath) = False Then
                                        Msgbox("Problem While saving Barcode Image.", vbInformation, ResolveResString(100))
                                        Exit Sub
                                    End If
                                End If
                            End If
                        End If

                        '01 jan 2024
                        If AllowBarCodePrinting(strAccountCode) Then
                            If GetPrintMethod(strAccountCode).ToUpper() = "TATA" Then
                                Dim StrTATAsuffix As String
                                StrTATAsuffix = gstrUNITID & mInvNo.ToString.Trim & DateTime.Now.ToString("ddMMyyHHmmssfff")

                                strBarcodeMsg = ObjBarcodeHMI.GenerateQRBarCodeForTATAMotors(gstrUserMyDocPath, True, Trim(Ctlinvoice.Text), mInvNo, gstrCONNECTIONSTRING, StrTATAsuffix)
                                If VB.Left(strBarcodeMsg, 1) <> "Y" Then
                                    CustomRollbackTrans()
                                    Msgbox("Problem While Generating Barcode Image.", vbInformation, ResolveResString(100))
                                    Exit Sub
                                Else
                                    strBarcodeMsg_paratemeter = Mid(strBarcodeMsg, 3)
                                    If Not SaveQRBarCodeImageTATA(Trim(Ctlinvoice.Text), mInvNo, strBarcodeMsg_paratemeter, StrTATAsuffix) Then
                                        CustomRollbackTrans()
                                        Msgbox("Problem While Saving Barcode Image.", vbInformation, ResolveResString(100))
                                        Exit Sub
                                    End If
                                End If
                            End If
                        End If

                        '01 jan 2024


                        If AllowBarCodePrinting(strAccountCode) = True And blnlinelevelcustomer = True Then
                            'sattu
                            '------------------------------------------------------------------------------------
                            rsGENERATEBARCODE = New ClsResultSetDB_Invoice
                            rsGENERATEBARCODE.GetResult("SELECT PRINT_METHOD FROM CUSTOMER_MST C WHERE C.UNIT_CODE='" & gstrUNITID & "' AND C.CUSTOMER_CODE='" & strAccountCode & "'")
                            strPrintMethod = UCase(rsGENERATEBARCODE.GetValue("PRINT_METHOD").ToString)
                            rsGENERATEBARCODE.ResultSetClose()
                            rsGENERATEBARCODE = Nothing
                            '01 MAR 2021
                            blnlinelevelcustomer = Find_Value("SELECT SOUPLD_LINE_LEVEL_SALESORDER FROM CUSTOMER_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & strAccountCode & "'")
                            If optInvYes(1).Checked = True And blnlinelevelcustomer = True And gstrUNITID = "MST" Then
                                If strPrintMethod = "NORMAL" Then
                                    strBarcodeMsg = ObjBarcodeHMI.GenerateBarCode_LINELEVEL_SALESORDER_2dbarcode(gstrUserMyDocPath, Trim(Ctlinvoice.Text), "NORMAL", "", "", True, Trim(Ctlinvoice.Text), gstrCONNECTIONSTRING)
                                    Dim strQuery As String

                                    If VB.Left(strBarcodeMsg, 1) <> "Y" Then
                                        Msgbox("Problem While Generating Barcode Image.", vbInformation, ResolveResString(100))
                                        Exit Sub
                                    Else
                                        If SaveBarCodeImage_singlelevelso_2DBARCODE(Ctlinvoice.Text, gstrUserMyDocPath, Mid(strBarcodeMsg, 3)) = False Then
                                            Msgbox("Problem While saving Barcode Image.", vbInformation, ResolveResString(100))
                                            Exit Sub
                                        End If
                                    End If
                                End If
                            End If

                            '01 MAR 2021
                            '03 apr 2023
                            '
                            If optInvYes(1).Checked = True Then
                                If (strPrintMethod = "NORMAL" And gstrUNITID = "M03") Or mblnAllowCustomerSpecificReport_COMP = True Then
                                    If (UCase(cmbInvType.Text) = "NORMAL INVOICE" And UCase(CmbCategory.Text) = "FINISHED GOODS") Or mblnAllowCustomerSpecificReport_COMP = True Then

                                        If DataExist("select TOP 1 1 from saleschallan_Dtl where UNIT_CODE='" & gstrUNITID & "' AND doc_no=" & Ctlinvoice.Text & " AND BARCODEIMAGE IS NULL") Then
                                            strBarcodeMsg = ObjBarcodeHMI.GenerateBarCode_LINELEVEL_SALESORDER_2dbarcode(gstrUserMyDocPath, Ctlinvoice.Text, "NORMAL", "", "", True, Trim(Ctlinvoice.Text), gstrCONNECTIONSTRING)
                                            Dim strQuery As String

                                            If VB.Left(strBarcodeMsg, 1) <> "Y" Then
                                                Msgbox("Problem While Generating Barcode Image, Can't take Invoice Print/Reprint", vbInformation, ResolveResString(100))
                                                Exit Sub
                                            Else
                                                If SaveBarCodeImage_singlelevelso_2DBARCODE(Ctlinvoice.Text, gstrUserMyDocPath, Mid(strBarcodeMsg, 3)) = False Then
                                                    Msgbox("Problem While saving Barcode Image.", vbInformation, ResolveResString(100))
                                                    Exit Sub
                                                End If
                                            End If

                                        End If
                                    End If
                                End If

                            End If
                            '03 apr 2023 changes ended

                            If optInvYes(0).Checked = True Then
                                If strPrintMethod = "NORMAL" Then
                                    strBarcodeMsg = ObjBarcodeHMI.GenerateBarCode_LINELEVEL_SALESORDER_2dbarcode(gstrUserMyDocPath, mInvNo, "NORMAL", "", "", True, Trim(Ctlinvoice.Text), gstrCONNECTIONSTRING)
                                    Dim strQuery As String

                                    If VB.Left(strBarcodeMsg, 1) <> "Y" Then
                                        Msgbox("Problem While Generating Barcode Image.", vbInformation, ResolveResString(100))
                                        Exit Sub
                                    Else
                                        If SaveBarCodeImage_singlelevelso_2DBARCODE(Ctlinvoice.Text, gstrUserMyDocPath, Mid(strBarcodeMsg, 3)) = False Then
                                            Msgbox("Problem While saving Barcode Image.", vbInformation, ResolveResString(100))
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
                                            Msgbox("Problem While Generating Barcode Image.", vbInformation, ResolveResString(100))
                                            Exit Sub
                                        End If
                                        If SaveBarCodeImage_singlelevelso(Ctlinvoice.Text, rsGENERATEBARCODE.GetValue("Cust_Item_Code").ToString, rsGENERATEBARCODE.GetValue("Item_Code").ToString, gstrUserMyDocPath, intRow) = False Then
                                            Msgbox("Problem While saving Barcode Image.", vbInformation, ResolveResString(100))
                                            Exit Sub
                                        End If
                                        rsGENERATEBARCODE.MoveNext()
                                    Next

                                    rsGENERATEBARCODE.ResultSetClose()
                                    rsGENERATEBARCODE = Nothing
                                End If
                            End If
                        End If

                        If optInvYes(1).Checked = True And AllowASNPrinting(strAccountCode) = True Then
                            If mblnASNExist = True Then
                                mP_Connection.Execute("Update CreatedASN Set ASN_NO='" & Trim$(txtASNNumber.Text) & "',Updatedon=getdate() where doc_no='" & Trim$(Me.Ctlinvoice.Text) & "' and UNIT_CODE = '" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            Else
                                mP_Connection.Execute("Insert into CreatedASN values('" & Trim$(Me.Ctlinvoice.Text) & "','" & Trim$(txtASNNumber.Text) & "',getdate(),'" & mP_User & "',getdate(),'" & gstrUNITID & "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            End If
                        End If
                        Frm.EnableDrillDown = True
                        ''Praveen Digital Sign Changes 
                        If optInvYes(1).Checked Then 'ÓNLY REPRINT
                            If mblnISCrystalReportRequired Then
                                Frm.Show()
                            Else
                                Frm.eMProCrystalReportViewer_Load(Me, New System.EventArgs)
                            End If
                        Else
                            Frm.Show()
                        End If

                        If optInvYes(1).Checked = True Then 'ÓNLY REPRINT
                            If (blnIsPDFExported = False) Then
                                EXPORTINVOICETOPDF_ONPRINTREPRINT(strAccountCode, lbldescription.Text, lblcategory.Text, RdAddSold, DT_A4CUSTOMER_INVOICEPRINTINGTAG)
                                blnIsPDFExported = True
                            End If

                        End If
                        'SATISH KESHARWANI CHANGE
                        'Dim strsql As String
                        'RdAddSold.Load(My.Application.Info.DirectoryPath & "\Reports\Annextureprinting_A4reports.rpt")
                        'strSql = "{SalesChallan_Dtl.Location_Code}='" & Trim(txtUnitCode.Text) & "' and {SalesChallan_Dtl.Doc_No} =" & Trim(Ctlinvoice.Text) & " and {SalesChallan_Dtl.UNIT_CODE} ='" & gstrUNITID & "' "
                        'RdAddSold.DataDefinition.RecordSelectionFormula = strSql
                        'Frm.show()
                        'If chktkmlbarcode.Checked = True And optInvYes(1).Checked = True Then
                        'Print_barcodelabel(Me.Ctlinvoice.Text)
                        'End If
                        'SATISH KESHARWANI CHANGE
                        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                        'If AllowBarCodePrinting(strAccountCode) = True And optInvYes(0).Checked = True Then
                        '    If DeleteBarCodeImage(Me.Ctlinvoice.Text) = False Then
                        '        MsgBox("Problem While deleting Barcode Image.", vbInformation, ResolveResString(100))
                        '        Exit Sub
                        '    End If
                        'End If
                    Else
                        Exit Sub
                    End If
                    If cmbInvType.Text.Trim.ToUpper = "REJECTION" Then
                        If CBool(Find_Value("select REJINVOICE_GRIN_PVKNOCKING from sales_parameter WHERE UNIT_CODE='" + gstrUNITID + "'")) = True Then
                            If DataExist("SELECT TOP 1 1 FROM MKT_INVREJ_DTL WHERE UNIT_CODE = '" & gstrUNITID & "' AND REJ_TYPE=1 AND CANCEL_FLAG=0 AND INVOICE_NO=" & Ctlinvoice.Text) Then 'GRIN RELATED QUERY
                                Call PrintDebitnote(MSTRREJECTIONNOTE)
                            End If
                        End If

                    End If
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT_TO_PRINTER
                    mCheckValARana = "PRINTTOPRINTER"
                    mblnLock_Clicked = True
                    ''Praveen Digital Sign
                    If optInvYes(0).Checked = True Then
                        mblnISCrystalReportRequired = CBool(Find_Value("select dbo.ISCrystalReportRequired_InvoicePrint ( '" + gstrUNITID + "','" & strAccountCode & "','" & lbldescription.Text & "','" & lblcategory.Text & "','O')"))
                    Else
                        mblnISCrystalReportRequired = CBool(Find_Value("select dbo.ISCrystalReportRequired_InvoicePrint ( '" + gstrUNITID + "','" & strAccountCode & "','" & lbldescription.Text & "','" & lblcategory.Text & "','R')"))
                    End If

                    If optInvYes(1).Checked = True Then
                        '102027599
                        If optInvYes(0).Checked = False Then
                            If mblnEwaybill_Print Then
                                Call IRN_QRBarcode()
                            End If
                        End If

                        If InvoiceGeneration(RdAddSold, Frm) = True Then
                            If gstrUNITID <> "MST" Then
                                Frm.EnablePrintButton = False
                            End If


                            If optInvYes(0).Checked = True Then
                                '10825102 
                                '  If mblnA4reports_invoicewise = True And mblncustomerlevel_A4report_functionlity = True Then
                                If mblncustomerlevel_A4report_functionlity = True Then
                                    'intNoCopies_A4reports_orignial = CInt(Find_Value("select isnull(MAX(SERIALNO),0) SERIALNO from A4CUSTOMER_INVOICEPRINTINGTAG  WHERE UNIT_CODE='" + gstrUNITID + "'AND CUSTOMER_CODE='" & strAccountCode & "' AND ORIGINAL_REPRINT='O'"))
                                    'intMaxLoop = intNoCopies_A4reports_orignial

                                    Dim DataRowFiltered_O() As DataRow = DT_A4CUSTOMER_INVOICEPRINTINGTAG.Select("ORIGINAL_REPRINT='O'")
                                    intNoCopies_A4reports_orignial = DataRowFiltered_O.Length  ' CInt(Find_Value("select isnull(MAX(SERIALNO),0) SERIALNO from A4CUSTOMER_INVOICEPRINTINGTAG  WHERE UNIT_CODE='" + gstrUNITID + "'AND CUSTOMER_CODE='" & strAccountCode & "' AND ORIGINAL_REPRINT='O'"))
                                    intMaxLoop = intNoCopies_A4reports_orignial
                                Else
                                    intMaxLoop = intNoCopies
                                End If
                            Else
                                'If mblnA4reports_invoicewise = True And mblncustomerlevel_A4report_functionlity = True Then
                                If mblncustomerlevel_A4report_functionlity = True Then
                                    '10825102 
                                    'intNoCopies_A4reports_REPRINT = CInt(Find_Value("select isnull(MAX(SERIALNO),0) SERIALNO  from A4CUSTOMER_INVOICEPRINTINGTAG  WHERE UNIT_CODE='" + gstrUNITID + "'AND CUSTOMER_CODE='" & strAccountCode & "' AND ORIGINAL_REPRINT='R'"))
                                    'intMaxLoop = intNoCopies_A4reports_REPRINT

                                    Dim DataRowFiltered_R() As DataRow = DT_A4CUSTOMER_INVOICEPRINTINGTAG.Select("ORIGINAL_REPRINT='R'")
                                    intNoCopies_A4reports_REPRINT = DataRowFiltered_R.Length  ' CInt(Find_Value("select isnull(MAX(SERIALNO),0) SERIALNO from A4CUSTOMER_INVOICEPRINTINGTAG  WHERE UNIT_CODE='" + gstrUNITID + "'AND CUSTOMER_CODE='" & strAccountCode & "' AND ORIGINAL_REPRINT='R'"))
                                    intMaxLoop = intNoCopies_A4reports_REPRINT
                                Else
                                    If intNoCopies > 1 Then
                                        intMaxLoop = intNoCopies - 1
                                    Else
                                        intMaxLoop = intNoCopies
                                    End If
                                End If
                            End If

                        End If
                    End If
                    If AllowBarCodePrinting(strAccountCode) = True Then
                        If optInvYes(0).Checked = False And mblnEwaybill_Print = True Then
                            '------------------------------------------------------------------------------------
                            rsGENERATEBARCODE = New ClsResultSetDB_Invoice
                            rsGENERATEBARCODE.GetResult("SELECT PRINT_METHOD FROM CUSTOMER_MST C WHERE C.UNIT_CODE='" & gstrUNITID & "' AND C.CUSTOMER_CODE='" & strAccountCode & "'")
                            strPrintMethod = UCase(rsGENERATEBARCODE.GetValue("PRINT_METHOD").ToString)
                            rsGENERATEBARCODE.ResultSetClose()
                            rsGENERATEBARCODE = Nothing

                            If optInvYes(0).Checked = False And mblnEwaybill_Print = True Then
                                If strPrintMethod = "MOBIS" Then

                                    strBarcodeMsg = ObjBarcodeHMI.GenerateBarCode_LINELEVEL_SALESORDER_2dbarcode_Mobis(gstrUserMyDocPath, Trim(Ctlinvoice.Text), "MOBIS", "", "", True, Trim(Ctlinvoice.Text), gstrCONNECTIONSTRING)


                                    If VB.Left(strBarcodeMsg, 1) <> "Y" Then
                                        Msgbox("Problem While Generating Barcode Image.", vbInformation, ResolveResString(100))
                                        Exit Sub
                                    Else
                                        If SaveBarCodeImage_singlelevelso_2DBARCODE(Ctlinvoice.Text, gstrUserMyDocPath, Mid(strBarcodeMsg, 3)) = False Then
                                            Msgbox("Problem While saving Barcode Image.", vbInformation, ResolveResString(100))
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

                    '03 apr 2023
                    If optInvYes(1).Checked = True Then
                        If (strPrintMethod = "NORMAL" And gstrUNITID = "M03") Or mblnAllowCustomerSpecificReport_COMP = True Then
                            If (UCase(cmbInvType.Text) = "NORMAL INVOICE" And UCase(CmbCategory.Text) = "FINISHED GOODS") Or mblnAllowCustomerSpecificReport_COMP = True Then

                                If DataExist("select TOP 1 1 from saleschallan_Dtl where UNIT_CODE='" & gstrUNITID & "' AND doc_no=" & Ctlinvoice.Text & " AND BARCODEIMAGE IS NULL") Then
                                    strBarcodeMsg = ObjBarcodeHMI.GenerateBarCode_LINELEVEL_SALESORDER_2dbarcode(gstrUserMyDocPath, Ctlinvoice.Text, "NORMAL", "", "", True, Trim(Ctlinvoice.Text), gstrCONNECTIONSTRING)
                                    Dim strQuery As String

                                    If VB.Left(strBarcodeMsg, 1) <> "Y" Then
                                        Msgbox("Problem While Generating Barcode Image, Can't take Invoice Print/Reprint", vbInformation, ResolveResString(100))
                                        Exit Sub
                                    Else
                                        If SaveBarCodeImage_singlelevelso_2DBARCODE(Ctlinvoice.Text, gstrUserMyDocPath, Mid(strBarcodeMsg, 3)) = False Then
                                            Msgbox("Problem While saving Barcode Image.", vbInformation, ResolveResString(100))
                                            Exit Sub
                                        End If
                                    End If

                                End If
                            End If
                        End If

                    End If
                    '03 apr 2023 changes ended
                    Call ReprintQRbarcode()

                    If Trim(UCase(cmbInvType.Text)) <> "REJECTION" Then  'Added by priti on 25 Jul 2025 to check customer outstanding Limit
                        blnisCreditLimitMandatory = Convert.ToBoolean(SqlConnectionclass.ExecuteScalar("Select isnull(isCreditLimitMandatory,0) from Sales_parameter where  unit_code='" & gstrUNITID & "'"))
                        If blnisCreditLimitMandatory = True Then
                            Dim dblOutstandingLimit As Double = SqlConnectionclass.ExecuteScalar("SELECT dbo.FUNC_FIN_GET_CL_BALANCE('" & gstrUNITID + "','" & strAccountCode & "','C')")
                            dblCreditLimit = SqlConnectionclass.ExecuteScalar("SELECT CreditLimit FROM customer_mst where unit_code='" & gstrUNITID + "' and Customer_Code ='" & strAccountCode & "'")
                            If (intTotalAmountValue + dblOutstandingLimit) > dblCreditLimit Then
                                Msgbox("Customer Out Standing Balance Amount Is greater than Customer Credit Limit.", vbInformation + vbOKOnly, ResolveResString(100))
                                Exit Sub
                            End If
                        End If
                    End If

                    If chkLockPrintingFlag.CheckState = CheckState.Checked Then
                        If optInvYes(0).Checked = True Then
                            If ConfirmWindow(10344, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then

                                ''Related Party validation start here
                                'Dim blnIsRelatedPartyUnit As Boolean = SqlConnectionclass.ExecuteScalar("Select isnull(IsRelatedParty,0) from sales_parameter (Nolock) where unit_code='" + gstrUNITID + "'")
                                'Dim blnIsRelatedPartyCust As Boolean = SqlConnectionclass.ExecuteScalar("Select IsRelatedPartyCust from saleschallan_dtl SC (Nolock),Customer_mst CM where SC.account_code=CM.customer_code and SC.Unit_code=CM.unit_code and  SC.unit_code='" + gstrUNITID + "'  and doc_no='" & Me.Ctlinvoice.Text & "'")
                                If (UCase(Trim(cmbInvType.Text)) <> "REJECTION") Then

                                    Dim dtSO As DataTable
                                    Dim strRetMsg As String = String.Empty
                                    Dim strCustomer As String = ""
                                    Dim strCustReference As String = ""
                                    Dim strAmendment As String = ""
                                    Dim blnIsOpenPO As Boolean
                                    strSQL = "Select OpenSO,SC.account_code,SC.cust_ref,SC.amendment_no from saleschallan_dtl SC (Nolock),Cust_ord_hdr SO where SC.account_code=SO.account_code  and SC.Cust_ref=SO.Cust_ref and SC.amendment_no=SO.amendment_no and  SC.unit_code='" + gstrUNITID + "'  and doc_no='" & Me.Ctlinvoice.Text & "'"
                                    dtSO = SqlConnectionclass.GetDataTable(strSQL)
                                    If dtSO.Rows.Count > 0 Then
                                        blnIsOpenPO = dtSO.Rows(0)("OpenSO")
                                        strCustomer = dtSO.Rows(0)("account_code")
                                        strCustReference = dtSO.Rows(0)("cust_ref")
                                        strAmendment = dtSO.Rows(0)("amendment_no")
                                    End If

                                    Dim CmdRP As New ADODB.Command
                                    With CmdRP
                                        .CommandText = "USP_VALIDATE_RELATED_PARTY_BUDGET_VALUE_SALESORDER"
                                        .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                                        .ActiveConnection = mP_Connection
                                        .CommandTimeout = 0

                                        .Parameters.Append(.CreateParameter("@UNIT_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                                        .Parameters.Append(.CreateParameter("@SO_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 25, strCustReference))
                                        .Parameters.Append(.CreateParameter("@AMENDMENT_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 25, strAmendment))
                                        .Parameters.Append(.CreateParameter("@CUSTOMER_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 8, strCustomer))
                                        .Parameters.Append(.CreateParameter("@IP_ADDRESS", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, mP_User))
                                        .Parameters.Append(.CreateParameter("@SOURCE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 4, "INV"))
                                        .Parameters.Append(.CreateParameter("@MSG_OUT", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInputOutput, 8000, ""))
                                        .Parameters.Append(.CreateParameter("@INVOICENO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, Trim(Ctlinvoice.Text)))

                                        .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

                                        strRetMsg = Convert.ToString(.Parameters("@MSG_OUT").Value)
                                        If String.IsNullOrEmpty(strRetMsg) = False Then

                                            strRetMsg += vbCrLf + "Transaction cannot save !"
                                            Msgbox(strRetMsg, MsgBoxStyle.Exclamation, "RELATED PARTY BUDGET VALIDATION")
                                            Exit Sub
                                        End If

                                    End With
                                    CmdRP = Nothing


                                End If
                                ''Related Party validation ends here

                                Dim strtime As String = GetServerDateTime()

                                Call Logging_Starting_End_Time("Invoice locking: Going To Start Transaction With Option " + optInvYes(0).Checked.ToString() + " Reprint Status " + chkprintreprint.Checked.ToString + " : Trans Count :" + GetTransCountIfAvailable().ToString(), strtime, "Saved", mInvNo)
                                CustomRollbackTrans()
                                mP_Connection.BeginTrans()
                                Call Logging_Starting_End_Time("Invoice locking: Transaction Started BUTTON_PRINT_TO_PRINTER: Trans Count :" + GetTransCountIfAvailable().ToString(), strtime, "Saved", mInvNo)

                                mP_Connection.Execute("set Dateformat 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)


                                If Not InvoiceGeneration(RdAddSold, Frm) Then
                                    CustomRollbackTrans()
                                    Exit Sub
                                End If
                                Call Logging_Starting_End_Time("Invoice locking: InvoiceGeneration Function Completed: Trans Count :" + GetTransCountIfAvailable().ToString(), strtime, "Saved", mInvNo)

                                If optInvYes(0).Checked = True Then

                                    'If mblnA4reports_invoicewise = True And mblncustomerlevel_A4report_functionlity = True Then
                                    If mblncustomerlevel_A4report_functionlity = True Then
                                        '10825102 
                                        'intNoCopies_A4reports_orignial = CInt(Find_Value("select isnull(MAX(SERIALNO),0) SERIALNO from A4CUSTOMER_INVOICEPRINTINGTAG  WHERE UNIT_CODE='" + gstrUNITID + "'AND CUSTOMER_CODE='" & strAccountCode & "' AND ORIGINAL_REPRINT='O'"))
                                        'intMaxLoop = intNoCopies_A4reports_orignial

                                        Dim DataRowFiltered_O() As DataRow = DT_A4CUSTOMER_INVOICEPRINTINGTAG.Select("ORIGINAL_REPRINT='O'")
                                        intNoCopies_A4reports_orignial = DataRowFiltered_O.Length  ' CInt(Find_Value("select isnull(MAX(SERIALNO),0) SERIALNO from A4CUSTOMER_INVOICEPRINTINGTAG  WHERE UNIT_CODE='" + gstrUNITID + "'AND CUSTOMER_CODE='" & strAccountCode & "' AND ORIGINAL_REPRINT='O'"))
                                        intMaxLoop = intNoCopies_A4reports_orignial
                                    Else
                                        intMaxLoop = intNoCopies
                                    End If
                                Else
                                    'If mblnA4reports_invoicewise = True And mblncustomerlevel_A4report_functionlity = True Then
                                    If mblncustomerlevel_A4report_functionlity = True Then
                                        '10825102 
                                        'intNoCopies_A4reports_REPRINT = CInt(Find_Value("select isnull(MAX(SERIALNO),0) SERIALNO  from A4CUSTOMER_INVOICEPRINTINGTAG  WHERE UNIT_CODE='" + gstrUNITID + "'AND CUSTOMER_CODE='" & strAccountCode & "' AND ORIGINAL_REPRINT='R'"))
                                        'intMaxLoop = intNoCopies_A4reports_REPRINT

                                        Dim DataRowFiltered_R() As DataRow = DT_A4CUSTOMER_INVOICEPRINTINGTAG.Select("ORIGINAL_REPRINT='R'")
                                        intNoCopies_A4reports_REPRINT = DataRowFiltered_R.Length  ' CInt(Find_Value("select isnull(MAX(SERIALNO),0) SERIALNO from A4CUSTOMER_INVOICEPRINTINGTAG  WHERE UNIT_CODE='" + gstrUNITID + "'AND CUSTOMER_CODE='" & strAccountCode & "' AND ORIGINAL_REPRINT='R'"))
                                        intMaxLoop = intNoCopies_A4reports_REPRINT
                                    Else
                                        If intNoCopies > 1 Then
                                            intMaxLoop = intNoCopies - 1
                                        Else
                                            intMaxLoop = intNoCopies
                                        End If
                                    End If
                                End If
                                '10617093
                                'If mblncustomerlevel_Annexture_printing = True And mblncustomerlevel_A4report_functionlity = True Then
                                '    If optInvYes(0).Checked = True Then
                                '        Printbarcode_TOYOTA(Ctlinvoice.Text.Trim)
                                '    End If
                                'End If
                                '10617093

                                If AllowBarCodePrinting(strAccountCode) = True And blnlinelevelcustomer = False Then
                                    rsGENERATEBARCODE = New ClsResultSetDB_Invoice
                                    rsGENERATEBARCODE.GetResult("SELECT PRINT_METHOD FROM CUSTOMER_MST C WHERE C.UNIT_CODE='" & gstrUNITID & "' AND C.CUSTOMER_CODE='" & strAccountCode & "'")
                                    strPrintMethod = UCase(rsGENERATEBARCODE.GetValue("PRINT_METHOD").ToString)
                                    rsGENERATEBARCODE.ResultSetClose()
                                    rsGENERATEBARCODE = Nothing

                                    strBarcodeMsg = ""
                                    If optInvYes(0).Checked = True Then
                                        If strPrintMethod = "TOYOTA" Then
                                            strSQL = "select  * from dbo.UFN_INV_BARCODE('" & gstrUNITID & "'," & Ctlinvoice.Text.Trim & ")"
                                            rsGENERATEBARCODE = New ClsResultSetDB_Invoice
                                            rsGENERATEBARCODE.GetResult(strSQL)
                                            intTotalNoofSlabs = rsGENERATEBARCODE.GetNoRows
                                            If intTotalNoofSlabs > 0 Then
                                                rsGENERATEBARCODE.MoveFirst()
                                                For intRow = 1 To intTotalNoofSlabs
                                                    strBarcodeMsg = ObjBarcodeHMI.GenerateQRBarCode(gstrUserMyDocPath, False, Trim(Ctlinvoice.Text), mInvNo, Val(rsGENERATEBARCODE.GetValue("FromLineNo").ToString), Val(rsGENERATEBARCODE.GetValue("ToLineNo").ToString), intRow, intTotalNoofSlabs, gstrCONNECTIONSTRING)
                                                    If VB.Left(strBarcodeMsg, 1) <> "Y" Then
                                                        CustomRollbackTrans()
                                                        Msgbox("Problem While Generating Barcode Image.", vbInformation, ResolveResString(100))
                                                        rsGENERATEBARCODE.ResultSetClose()
                                                        rsGENERATEBARCODE = Nothing
                                                        Exit Sub
                                                    Else
                                                        '10812364
                                                        strBarcodeMsg_paratemeter = Mid(strBarcodeMsg, 3)
                                                        '10812364
                                                        If SaveQRBarCodeImage(Trim(Ctlinvoice.Text), Val(rsGENERATEBARCODE.GetValue("FromLineNo").ToString), Val(rsGENERATEBARCODE.GetValue("ToLineNo").ToString), strBarcodeMsg_paratemeter, intRow, intTotalNoofSlabs, mInvNo) = False Then
                                                            CustomRollbackTrans()
                                                            Msgbox("Problem While saving Barcode Image.", vbInformation, ResolveResString(100))
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
                                            '12 oct 2017'
                                            '03 JAN 2024'
                                        ElseIf AllowBarCodePrinting(strAccountCode) Then
                                            If GetPrintMethod(strAccountCode).ToUpper() = "TATA" Then
                                                Dim StrTATAsuffix As String
                                                StrTATAsuffix = gstrUNITID & mInvNo.ToString.Trim & DateTime.Now.ToString("ddMMyyHHmmssfff")

                                                strBarcodeMsg = ObjBarcodeHMI.GenerateQRBarCodeForTATAMotors(gstrUserMyDocPath, True, Trim(Ctlinvoice.Text), mInvNo, gstrCONNECTIONSTRING, StrTATAsuffix)
                                                If VB.Left(strBarcodeMsg, 1) <> "Y" Then
                                                    CustomRollbackTrans()
                                                    Msgbox("Problem While Generating Barcode Image.", vbInformation, ResolveResString(100))
                                                    Exit Sub
                                                Else
                                                    strBarcodeMsg_paratemeter = Mid(strBarcodeMsg, 3)
                                                    If Not SaveQRBarCodeImageTATA(Trim(Ctlinvoice.Text), mInvNo, strBarcodeMsg_paratemeter, StrTATAsuffix) Then
                                                        CustomRollbackTrans()
                                                        Msgbox("Problem While Saving Barcode Image.", vbInformation, ResolveResString(100))
                                                        Exit Sub
                                                    End If
                                                End If
                                            End If


                                            '03 JAN 2024
                                        ElseIf strPrintMethod = "TOYOTA_NEW" Then
                                            strSQL = "select * from dbo.UFN_INV_BARCODE('" & gstrUNITID & "'," & Ctlinvoice.Text.Trim & ")"
                                            rsGENERATEBARCODE = New ClsResultSetDB_Invoice
                                            rsGENERATEBARCODE.GetResult(strSQL)
                                            intTotalNoofSlabs = rsGENERATEBARCODE.GetNoRows
                                            If intTotalNoofSlabs > 0 Then
                                                rsGENERATEBARCODE.MoveFirst()
                                                For intRow = 1 To intTotalNoofSlabs
                                                    strBarcodeMsg = ObjBarcodeHMI.GenerateQRBarCode_NEWTOYOTA(gstrUserMyDocPath, False, Trim(Ctlinvoice.Text), mInvNo, Val(rsGENERATEBARCODE.GetValue("FromLineNo").ToString), Val(rsGENERATEBARCODE.GetValue("ToLineNo").ToString), intRow, intTotalNoofSlabs, gstrCONNECTIONSTRING)
                                                    If VB.Left(strBarcodeMsg, 1) <> "Y" Then
                                                        CustomRollbackTrans()
                                                        Msgbox("Problem While Generating Barcode Image.", vbInformation, ResolveResString(100))

                                                        rsGENERATEBARCODE.ResultSetClose()
                                                        rsGENERATEBARCODE = Nothing
                                                        Exit Sub
                                                    Else
                                                        '10812364
                                                        strBarcodeMsg_paratemeter = Mid(strBarcodeMsg, 3)
                                                        '10812364
                                                        If SaveQRBarCodeImage(Trim(Ctlinvoice.Text), Val(rsGENERATEBARCODE.GetValue("FromLineNo").ToString), Val(rsGENERATEBARCODE.GetValue("ToLineNo").ToString), strBarcodeMsg_paratemeter, intRow, intTotalNoofSlabs, mInvNo) = False Then
                                                            CustomRollbackTrans()
                                                            Msgbox("Problem While saving Barcode Image.", vbInformation, ResolveResString(100))

                                                            rsGENERATEBARCODE.ResultSetClose()
                                                            rsGENERATEBARCODE = Nothing
                                                            Exit Sub
                                                        End If
                                                    End If
                                                Next
                                            End If
                                            rsGENERATEBARCODE.ResultSetClose()
                                            rsGENERATEBARCODE = Nothing
                                            '------------------------------------------------------------------------------------
                                            '12 oct 2017

                                        Else
                                            strBarcodeMsg = ObjBarcodeHMI.GenerateBarCode(gstrUserMyDocPath, mInvNo, True, Trim(Ctlinvoice.Text), gstrCONNECTIONSTRING)
                                            If VB.Left(strBarcodeMsg, 1) <> "Y" Then
                                                CustomRollbackTrans()
                                                Msgbox("Problem While Generating Barcode Image.", vbInformation, ResolveResString(100))

                                                Exit Sub
                                            End If
                                            If SaveBarCodeImage(Me.Ctlinvoice.Text, gstrUserMyDocPath) = False Then
                                                CustomRollbackTrans()
                                                Msgbox("Problem While saving Barcode Image.", vbInformation, ResolveResString(100))

                                                Exit Sub
                                            End If
                                        End If
                                    End If
                                End If

                                Call Logging_Starting_End_Time("Invoice locking: Going To Execute ReprintQRbarcode: Trans Count :" + GetTransCountIfAvailable().ToString(), strtime, "Saved", mInvNo)


                                Call ReprintQRbarcode()
                                '16 feb 2017 Vendor Portal changes
                                If optInvYes(0).Checked = True Then
                                    Dim blnCustunitcodeReqd As Boolean = DataExist("SELECT TOP 1 1 UNIT_CODE  FROM PUR_SCH_TRF_MAPPING WHERE SALE_UNIT = '" & gstrUNITID & "' AND PUR_UNIT_CUSTOMER_CODE ='" & strAccountCode & "' AND ACTIVE=1")
                                    ''13 FEB 2025 --Auto Gate Inward Functionality For Intergroup Changes
                                    rsGENERATEBARCODE = New ClsResultSetDB_Invoice
                                    rsGENERATEBARCODE.GetResult("SELECT TOP 1 SERVER_IP,EMPRO_DB_NAME  FROM PUR_SCH_TRF_MAPPING WHERE SALE_UNIT = '" & gstrUNITID & "' AND PUR_UNIT_CUSTOMER_CODE ='" & strAccountCode & "'")
                                    Dim strSERVER_IP = UCase(rsGENERATEBARCODE.GetValue("SERVER_IP").ToString)
                                    Dim strEMPRO_DB_NAME = UCase(rsGENERATEBARCODE.GetValue("EMPRO_DB_NAME").ToString)
                                    rsGENERATEBARCODE.ResultSetClose()
                                    rsGENERATEBARCODE = Nothing
                                    If strSERVER_IP <> "0" And strEMPRO_DB_NAME <> "0" Then
                                        ''Dim blnBinningBarcodeReqd As Boolean = DataExist("SELECT TOP 1 1 FROM POCONFIG_MST WHERE AUTO_PO_SCHEDULE_TRF = 1 and UNIT_CODE in( SELECT TOP  1 UNIT_CODE  FROM PUR_SCH_TRF_MAPPING WHERE SALE_UNIT = '" & gstrUNITID & "' AND PUR_UNIT_CUSTOMER_CODE ='" & strAccountCode & "')")
                                        Dim blnBinningBarcodeReqd As Boolean = DataExist("SELECT TOP 1 1 FROM [" + strSERVER_IP + "].[" + strEMPRO_DB_NAME + "].dbo.POCONFIG_MST WHERE AUTO_PO_SCHEDULE_TRF = 1 and UNIT_CODE in( SELECT TOP  1 UNIT_CODE  FROM PUR_SCH_TRF_MAPPING WHERE SALE_UNIT = '" & gstrUNITID & "' AND PUR_UNIT_CUSTOMER_CODE ='" & strAccountCode & "' AND ACTIVE=1)")
                                        If blnCustunitcodeReqd = True And blnBinningBarcodeReqd = True And (UCase(Trim(cmbInvType.Text)) = "TRANSFER INVOICE" Or UCase(Trim(cmbInvType.Text)) = "NORMAL INVOICE" Or UCase(Trim(cmbInvType.Text)) = "INTER-DIVISION") Then
                                            Dim cmd As SqlCommand = Nothing
                                            cmd = New System.Data.SqlClient.SqlCommand()
                                            cmd.Connection = SqlConnectionclass.GetConnection

                                            With cmd
                                                .CommandText = String.Empty
                                                .Parameters.Clear()
                                                .CommandType = CommandType.StoredProcedure
                                                .CommandText = "AUTO_VP_ASN_GEN"
                                                .CommandTimeout = 0

                                                .Parameters.Add("@VendorCode", SqlDbType.VarChar, 10).Value = gstrUNITID
                                                .Parameters.Add("@TEMPInvoiceNo", SqlDbType.VarChar, 10).Value = Ctlinvoice.Text.Trim
                                                .Parameters.Add("@InvoiceNo", SqlDbType.VarChar, 10).Value = Trim(mInvNo)
                                                .Parameters.Add("@IPADDRESS", SqlDbType.VarChar, 20).Value = gstrIpaddressWinSck
                                                .Parameters.Add("@MessageOut", SqlDbType.VarChar, 5000).Direction = ParameterDirection.Output

                                                cmd.ExecuteNonQuery()
                                                If .Parameters(.Parameters.Count - 1).Value.ToString().Trim <> "SUCCESS" Then
                                                    CustomRollbackTrans()
                                                    MessageBox.Show(.Parameters(.Parameters.Count - 1).Value.ToString(), ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error)

                                                    Exit Sub
                                                End If

                                            End With
                                            cmd = Nothing

                                            'Dim objComm As New ADODB.Command
                                            'With objComm
                                            '    .ActiveConnection = mP_Connection

                                            '    .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                                            '    .CommandText = "AUTO_VP_ASN_GEN"
                                            '    .CommandTimeout = 0
                                            '    .Parameters.Append(.CreateParameter("@VendorCode", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                                            '    .Parameters.Append(.CreateParameter("@TEMPInvoiceNo", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, Ctlinvoice.Text.Trim))
                                            '    .Parameters.Append(.CreateParameter("@InvoiceNo", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, Trim(mInvNo)))
                                            '    .Parameters.Append(.CreateParameter("@IPADDRESS", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, gstrIpaddressWinSck))
                                            '    .Parameters.Append(.CreateParameter("@MessageOut", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInputOutput, 1000, 0))

                                            '    .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                            '    If .Parameters(.Parameters.Count - 1).Value.ToString().Trim <> "SUCCESS" Then
                                            '        CustomRollbackTrans()
                                            '        MessageBox.Show(.Parameters(.Parameters.Count - 1).Value.ToString(), ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error)

                                            '        Exit Sub
                                            '    End If
                                            'End With
                                            'objComm = Nothing
                                        End If
                                    End If
                                    ''13 FEB 2025 --Auto Gate Inward Functionality For Intergroup Changes
                                End If
                                '16 feb 2017 Vendor Portal changes

                                Call Logging_Starting_End_Time("Invoice locking: AllowBarCodePrinting : Trans Count :" + GetTransCountIfAvailable().ToString(), strtime, "Saved", mInvNo)

                                If AllowBarCodePrinting(strAccountCode) = True And blnlinelevelcustomer = True Then

                                    'sattu
                                    '------------------------------------------------------------------------------------
                                    rsGENERATEBARCODE = New ClsResultSetDB_Invoice
                                    rsGENERATEBARCODE.GetResult("SELECT PRINT_METHOD FROM CUSTOMER_MST C WHERE C.UNIT_CODE='" & gstrUNITID & "' AND C.CUSTOMER_CODE='" & strAccountCode & "'")
                                    strPrintMethod = UCase(rsGENERATEBARCODE.GetValue("PRINT_METHOD").ToString)
                                    rsGENERATEBARCODE.ResultSetClose()
                                    rsGENERATEBARCODE = Nothing

                                    If optInvYes(0).Checked = True Then
                                        If strPrintMethod = "TOYOTA" Then
                                            strSQL = "select  * from dbo.UFN_INV_BARCODE('" & gstrUNITID & "'," & Ctlinvoice.Text.Trim & ")"
                                            rsGENERATEBARCODE = New ClsResultSetDB_Invoice
                                            rsGENERATEBARCODE.GetResult(strSQL)
                                            intTotalNoofSlabs = rsGENERATEBARCODE.GetNoRows
                                            If intTotalNoofSlabs > 0 Then
                                                rsGENERATEBARCODE.MoveFirst()
                                                For intRow = 1 To intTotalNoofSlabs
                                                    strBarcodeMsg = ObjBarcodeHMI.GenerateQRBarCode(gstrUserMyDocPath, True, Trim(Ctlinvoice.Text), mInvNo, Val(rsGENERATEBARCODE.GetValue("FromLineNo").ToString), Val(rsGENERATEBARCODE.GetValue("ToLineNo").ToString), intRow, intTotalNoofSlabs, gstrCONNECTIONSTRING)
                                                    If VB.Left(strBarcodeMsg, 1) <> "Y" Then
                                                        CustomRollbackTrans()
                                                        Msgbox("Problem While Generating Barcode Image.", vbInformation, ResolveResString(100))

                                                        rsGENERATEBARCODE.ResultSetClose()
                                                        rsGENERATEBARCODE = Nothing
                                                        Exit Sub
                                                    Else
                                                        '10812364
                                                        strBarcodeMsg_paratemeter = Mid(strBarcodeMsg, 3)
                                                        '10812364
                                                        If SaveQRBarCodeImage(Trim(Ctlinvoice.Text), Val(rsGENERATEBARCODE.GetValue("FromLineNo").ToString), Val(rsGENERATEBARCODE.GetValue("ToLineNo").ToString), strBarcodeMsg_paratemeter, intRow, intTotalNoofSlabs, mInvNo) = False Then
                                                            CustomRollbackTrans()
                                                            Msgbox("Problem While saving Barcode Image.", vbInformation, ResolveResString(100))

                                                            rsGENERATEBARCODE.ResultSetClose()
                                                            rsGENERATEBARCODE = Nothing
                                                            Exit Sub
                                                        End If
                                                    End If
                                                Next
                                            End If
                                            rsGENERATEBARCODE.ResultSetClose()
                                            rsGENERATEBARCODE = Nothing
                                            '------------------------------------------------------------------------------------
                                            '12 oct 2017'
                                        ElseIf strPrintMethod = "TOYOTA_NEW" Then
                                            strSQL = "select * from dbo.UFN_INV_BARCODE('" & gstrUNITID & "'," & Ctlinvoice.Text.Trim & ")"
                                            rsGENERATEBARCODE = New ClsResultSetDB_Invoice
                                            rsGENERATEBARCODE.GetResult(strSQL)
                                            intTotalNoofSlabs = rsGENERATEBARCODE.GetNoRows
                                            If intTotalNoofSlabs > 0 Then
                                                rsGENERATEBARCODE.MoveFirst()
                                                For intRow = 1 To intTotalNoofSlabs
                                                    strBarcodeMsg = ObjBarcodeHMI.GenerateQRBarCode_NEWTOYOTA(gstrUserMyDocPath, True, Trim(Ctlinvoice.Text), mInvNo, Val(rsGENERATEBARCODE.GetValue("FromLineNo").ToString), Val(rsGENERATEBARCODE.GetValue("ToLineNo").ToString), intRow, intTotalNoofSlabs, gstrCONNECTIONSTRING)
                                                    If VB.Left(strBarcodeMsg, 1) <> "Y" Then
                                                        CustomRollbackTrans()
                                                        Msgbox("Problem While Generating Barcode Image.", vbInformation, ResolveResString(100))

                                                        rsGENERATEBARCODE.ResultSetClose()
                                                        rsGENERATEBARCODE = Nothing
                                                        Exit Sub
                                                    Else
                                                        '10812364
                                                        strBarcodeMsg_paratemeter = Mid(strBarcodeMsg, 3)
                                                        '10812364
                                                        If SaveQRBarCodeImage(Trim(Ctlinvoice.Text), Val(rsGENERATEBARCODE.GetValue("FromLineNo").ToString), Val(rsGENERATEBARCODE.GetValue("ToLineNo").ToString), strBarcodeMsg_paratemeter, intRow, intTotalNoofSlabs, mInvNo) = False Then
                                                            CustomRollbackTrans()
                                                            Msgbox("Problem While saving Barcode Image.", vbInformation, ResolveResString(100))
                                                            rsGENERATEBARCODE.ResultSetClose()
                                                            rsGENERATEBARCODE = Nothing
                                                            Exit Sub
                                                        End If
                                                    End If
                                                Next
                                            End If
                                            rsGENERATEBARCODE.ResultSetClose()
                                            rsGENERATEBARCODE = Nothing
                                            '------------------------------------------------------------------------------------
                                            '12 oct 2017

                                        ElseIf strPrintMethod = "NORMAL" Then
                                            strBarcodeMsg = ObjBarcodeHMI.GenerateBarCode_LINELEVEL_SALESORDER_2dbarcode(gstrUserMyDocPath, mInvNo, "NORMAL", "", "", True, Trim(Ctlinvoice.Text), gstrCONNECTIONSTRING)
                                            If VB.Left(strBarcodeMsg, 1) <> "Y" Then
                                                CustomRollbackTrans()
                                                Msgbox("Problem While Generating Barcode Image.", vbInformation, ResolveResString(100))
                                                Exit Sub
                                            End If

                                            If SaveBarCodeImage_singlelevelso_2DBARCODE(Ctlinvoice.Text, gstrUserMyDocPath, strBarcodeMsg) = False Then
                                                CustomRollbackTrans()
                                                Msgbox("Problem While saving Barcode Image.", vbInformation, ResolveResString(100))
                                                Exit Sub
                                            End If
                                        Else
                                            'sattu
                                            '---------------------------------------------------------------------------------------------------
                                            rsGENERATEBARCODE = New ClsResultSetDB_Invoice
                                            rsGENERATEBARCODE.GetResult("SELECT PRINT_METHOD,ITEM_CODE,CUST_ITEM_CODE FROM SALES_DTL SD ,SALESCHALLAN_DTL SC ,CUSTOMER_MST C WHERE C.UNIT_CODE=SC.UNIT_CODE AND C.CUSTOMER_CODE=SC.ACCOUNT_CODE AND SC.UNIT_CODE=SD.UNIT_CODE AND " &
                                                " SC.DOC_NO=SD.DOC_NO  AND SC.UNIT_CODE='" & gstrUNITID & "' AND  " &
                                                    " SC.DOC_NO= " & Ctlinvoice.Text & " ORDER BY CUST_ITEM_CODE ")

                                            intTotalNoofitemsinInvoices = rsGENERATEBARCODE.GetNoRows
                                            rsGENERATEBARCODE.MoveFirst()

                                            For intRow = 1 To intTotalNoofitemsinInvoices
                                                strBarcodeMsg = ObjBarcodeHMI.GenerateBarCode_LINELEVEL_SALESORDER(gstrUserMyDocPath, mInvNo, rsGENERATEBARCODE.GetValue("PRINT_METHOD").ToString, rsGENERATEBARCODE.GetValue("Cust_Item_Code").ToString, rsGENERATEBARCODE.GetValue("Item_Code").ToString, True, Trim(Ctlinvoice.Text), gstrCONNECTIONSTRING)
                                                If VB.Left(strBarcodeMsg, 1) <> "Y" Then
                                                    CustomRollbackTrans()
                                                    Msgbox("Problem While Generating Barcode Image.", vbInformation, ResolveResString(100))

                                                    Exit Sub
                                                End If

                                                If SaveBarCodeImage_singlelevelso(Ctlinvoice.Text, rsGENERATEBARCODE.GetValue("Cust_Item_Code").ToString, rsGENERATEBARCODE.GetValue("Item_Code").ToString, gstrUserMyDocPath, intRow) = False Then
                                                    CustomRollbackTrans()
                                                    Msgbox("Problem While saving Barcode Image.", vbInformation, ResolveResString(100))
                                                    Exit Sub
                                                End If
                                                rsGENERATEBARCODE.MoveNext()
                                            Next

                                            rsGENERATEBARCODE.ResultSetClose()
                                            rsGENERATEBARCODE = Nothing
                                            '---------------------------------------------------------------------------------------------------
                                        End If
                                    End If
                                End If


                                Call Logging_Starting_End_Time("Invoice locking: AllowBarCodePrinting Done : Trans Count :" + GetTransCountIfAvailable().ToString(), strtime, "Saved", mInvNo)
                                ''10910711
                                'If DataExist("SELECT TOP 1 1 FROM SALECONF WHERE UNIT_CODE='" + gstrUNITID + "'  AND FIN_START_DATE <= GETDATE() AND FIN_END_DATE >= GETDATE() AND ATN_ENABLED =1  and INVOICE_TYPE='" & lbldescription.Text & "' AND SUB_TYPE='" & lblcategory.Text & "'") Then
                                '    mP_Connection.Execute("Exec FA_AUTO_ATN_POSTING  '" & mInvNo & "', '" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

                                '    Dim objComm As New ADODB.Command
                                '    With objComm
                                '        .ActiveConnection = mP_Connection
                                '        .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                                '        .CommandText = "USP_FIN_JV"
                                '        .CommandTimeout = 0
                                '        .Parameters.Append(.CreateParameter("@IP_ADDRESS", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, gstrIpaddressWinSck))
                                '        .Parameters.Append(.CreateParameter("@UNIT_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                                '        .Parameters.Append(.CreateParameter("@TMP_INVOICE_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Ctlinvoice.Text))
                                '        .Parameters.Append(.CreateParameter("@INVOICE_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, mInvNo.ToString))
                                '        .Parameters.Append(.CreateParameter("@MODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 1, "A"))
                                '        .Parameters.Append(.CreateParameter("@ERR", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInputOutput, 800, 0))

                                '        .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                '        If .Parameters(.Parameters.Count - 1).Value.ToString().Trim.Length <> 0 Then
                                '            MsgBox("Unable To Do ATN .", MsgBoxStyle.Information, ResolveResString(100))
                                '            CustomRollbackTrans()
                                '            Exit Sub
                                '        End If
                                '    End With
                                '    objComm = Nothing

                                'End If

                                ''10910711

                                'Added for Issue ID 22303 Starts
                                'If InvAgstBarCode() = True And mstrFGDomestic = "1" Then
                                If InvAgstBarCode() = True And (mstrFGDomestic = "1" Or mstrFGDomestic = "2") Then
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

                                Call Logging_Starting_End_Time("Invoice locking: CSM Checking Start : Trans Count :" + GetTransCountIfAvailable().ToString(), strtime, "Saved", mInvNo)

                                mP_Connection.Execute("UPDATE CSM_INVOICE_DTL SET DOC_NO = " & mInvNo & ", INVOICE_LOCK = 1, INVOICE_LOCK_DT = GETDATE() WHERE  UNIT_CODE = '" & gstrUNITID & "'  and DOC_NO = " & Ctlinvoice.Text & " AND INVOICE_LOCK = 0", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                mP_Connection.Execute(saleschallan, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                If UCase(cmbInvType.Text) = "NORMAL INVOICE" And UCase(CmbCategory.Text) = "FINISHED GOODS" Then
                                    Dim blnCSM_Knockingoff_req As Boolean = DataExist("SELECT TOP 1 1 FROM SALES_PARAMETER WHERE CSM_KNOCKINGOFF_REQ = 1 and UNIT_CODE = '" & gstrUNITID & "'")
                                    If blnCSM_Knockingoff_req Then
                                        Dim objComm As New ADODB.Command
                                        With objComm
                                            .ActiveConnection = mP_Connection
                                            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                                            .CommandText = "USP_LOCK_CSM_DETAILS"
                                            .CommandTimeout = 0
                                            .Parameters.Append(.CreateParameter("@UNIT_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                                            .Parameters.Append(.CreateParameter("@OLD_INV_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, CInt(Ctlinvoice.Text)))
                                            .Parameters.Append(.CreateParameter("@NEW_INV_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, mInvNo))
                                            .Parameters.Append(.CreateParameter("@RETURN", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInputOutput, , 0))

                                            .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                            If .Parameters(.Parameters.Count - 1).Value <> 0 Then
                                                CustomRollbackTrans()
                                                Msgbox("Unable To Lock CSM Knocking Off Details.", MsgBoxStyle.Information, ResolveResString(100))

                                                Exit Sub
                                            End If
                                        End With
                                        objComm = Nothing
                                    End If
                                End If
                                'smrc thailand changes started
                                If gstrUNITID = "STH" Then
                                    STRCUSTTYPE = Find_Value("SELECT THAI_CUST_TYPE FROM customer_mst WHERE  UNIT_CODE = '" & gstrUNITID & "' AND  CUSTOMER_CODE='" & strAccountCode & "'")
                                    If UCase(cmbInvType.Text) = "NORMAL INVOICE" Or UCase(cmbInvType.Text) = "EXPORT INVOICE" Then
                                        Dim objComm As New ADODB.Command
                                        With objComm
                                            .ActiveConnection = mP_Connection
                                            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                                            .CommandText = "USP_SMRC_THAI_UPDATE_SAN_STOCK_LOCATIONS"
                                            .CommandTimeout = 0
                                            .Parameters.Append(.CreateParameter("@UNIT_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                                            .Parameters.Append(.CreateParameter("@DOC_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, mInvNo))
                                            .Parameters.Append(.CreateParameter("@IPADDRESS", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 50, gstrIpaddressWinSck))
                                            .Parameters.Append(.CreateParameter("@USERID", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 50, mP_User))
                                            .Parameters.Append(.CreateParameter("@CUSTTYPE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 2, STRCUSTTYPE))
                                            .Parameters.Append(.CreateParameter("@ERRMSG", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInputOutput, 8000, ""))

                                            .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                            If .Parameters(.Parameters.Count - 1).Value <> "" Then
                                                CustomRollbackTrans()
                                                Msgbox(.Parameters(.Parameters.Count - 1).Value.ToString(), MsgBoxStyle.Information, ResolveResString(100))

                                                Exit Sub
                                            End If
                                        End With
                                        objComm = Nothing
                                    End If
                                End If

                                'smrc thailand changes ended
                                If updatePOflag = True Then
                                    mP_Connection.Execute(strupdatecustodtdtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                End If
                                If updatestockflag = True Then
                                    mP_Connection.Execute(strupdateitbalmst, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                End If
                                If UCase(cmbInvType.Text) = "JOBWORK INVOICE" Then
                                    If Len(mstrAnnex) <> 0 Then
                                        mP_Connection.Execute("SET DATEFORMAT 'DMY'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                        mP_Connection.Execute(mstrAnnex, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                    End If
                                End If
                                If UCase(Me.lbldescription.Text) = "REJ" Then
                                    If Len(Trim(mCust_Ref)) > 0 Then
                                        mP_Connection.Execute(strupdateGrinhdr, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                    End If
                                End If
                                If Len(Trim(strBatchQuery)) > 0 Then
                                    mP_Connection.Execute(strBatchQuery, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                    mP_Connection.Execute("Update ItemBatch_dtl Set Doc_no = '" & Trim(CStr(mInvNo)) & "' Where Doc_no = '" & Trim(Me.Ctlinvoice.Text) & "' and Doc_Type = 9999 and UNIT_CODE = '" & gstrUNITID & "' ", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                End If
                                blnDSTracking = CBool(Find_Value("SELECT isnull(DSWiseTracking,0)as DSWiseTracking FROM sales_parameter WHERE UNIT_CODE='" + gstrUNITID + "'"))
                                'If ((UCase(Trim(cmbInvType.Text)) = "NORMAL INVOICE") And (UCase(CStr((Trim(CmbCategory.Text)) = "FINISHED GOODS")))) Then
                                If (((UCase(Trim(cmbInvType.Text))) = "TRANSFER INVOICE")) Or (((UCase(Trim(cmbInvType.Text))) = "INTER-DIVISION")) Or ((UCase(Trim(cmbInvType.Text)) = "NORMAL INVOICE") And (UCase(CStr((Trim(CmbCategory.Text)) = "FINISHED GOODS")))) Then
                                    If UCase(Trim(cmbInvType.Text)) <> "SERVICE INVOICE" Then
                                        If blnInvoiceAgainstMultipleSO = False Then
                                            '10665764
                                            'If blnDSTracking = True And AllowTextFileGeneration_SMIIEL(strAccountCode) = True And CBool(Find_Value("SELECT SMIEL_FTP_INVOICE FROM SALES_PARAMETER WHERE UNIT_CODE='" + gstrUNITID + "'")) = True Then
                                            If blnDSTracking = True And CBool(Find_Value("SELECT SMIEL_FTP_INVOICE FROM SALES_PARAMETER WHERE UNIT_CODE='" + gstrUNITID + "'")) = True Then
                                                '10665764
                                                If Not UpdateMktSchedule() Then
                                                    CustomRollbackTrans()
                                                    Exit Sub
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                                'If AllowASNTextFileGeneration(strAccountCode) = True Then
                                '    mP_Connection.Execute("UPDATE MKT_ASN_INVDTL SET DOC_NO=" & Trim(mInvNo) & " where dOC_NO=" & Trim(Ctlinvoice.Text) & " and UNIT_CODE = '" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                '    If FordASNFileGeneration(mInvNo, strAccountCode) = False Then
                                '        CustomRollbackTrans()
                                '        Exit Sub
                                '    Else
                                '        If Len(mstrupdateASNdtl) > 0 Then
                                '            mP_Connection.Execute(mstrupdateASNdtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                '            mP_Connection.Execute(mstrupdateASNCumFig, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                '        End If
                                '    End If
                                'End If
                                Call Logging_Starting_End_Time("Invoice locking: Fin Post Start : Trans Count :" + GetTransCountIfAvailable().ToString(), strtime, "Saved", mInvNo)

                                If mblnpostinfin = True Then
                                    objDrCr = New prj_DrCrNote.cls_DrCrNote(GetServerDate())
                                    Try
                                        If UCase(Trim(cmbInvType.Text)) = "REJECTION" And CUST_REJECTION_FLAG = False Then
                                            If RejInvOptionalPostingFlag() = True And DataExist("SELECT TOP 1 1 FROM MKT_INVREJ_DTL WHERE UNIT_CODE = '" & gstrUNITID & "' AND REJ_TYPE=1 AND CANCEL_FLAG=0 AND INVOICE_NO=" & Ctlinvoice.Text) Then
                                                strRetval = "Y"
                                            Else
                                                prj_DocGenerator.cls_DocumentGenerator.gbln_AR_AP_Dr_Cr_Doc_Sub_Category = "DR"
                                                strRetval = objDrCr.SetAPDocument(gstrUNITID, mstrMasterString, mstrDetailString, prj_GLTransactions.cls_GLTransactions.udtOperationType.optInsert, gstrCONNECTIONSTRING, "", mP_Connection)
                                                prj_DocGenerator.cls_DocumentGenerator.gbln_AR_AP_Dr_Cr_Doc_Sub_Category = ""
                                            End If
                                        Else
                                            If UCase(Trim(cmbInvType.Text)) = "REJECTION" And RejInvOptionalPostingFlag() = True And Not DataExist("SELECT TOP 1 1 FROM MKT_INVREJ_DTL WHERE UNIT_CODE = '" & gstrUNITID & "' AND REJ_TYPE=2 AND CANCEL_FLAG=0 AND INVOICE_NO=" & mInvNo) Then
                                                strRetval = "Y"
                                            Else
                                                strRetval = objDrCr.SetARInvoiceDocument(gstrUNITID, mstrMasterString, mstrDetailString, prj_GLTransactions.cls_GLTransactions.udtOperationType.optInsert, gstrCONNECTIONSTRING, "", mP_Connection)
                                            End If

                                        End If
                                        strRetval = CheckString(strRetval)
                                        objDrCr = Nothing
                                    Catch ex As Exception
                                        Throw New ApplicationException("Finance DLL:" + ex.Message.ToString())
                                    Finally
                                        objDrCr = Nothing
                                    End Try
                                Else
                                    strRetval = "Y"
                                End If

                                '================================================
                                'Dim intTransactionCount As Double
                                'Dim rstTransactions As New ADODB.Recordset
                                'Call rstTransactions.Open("SELECT @@trancount", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
                                'intTransactionCount = IIf(Not IsDBNull(rstTransactions.Fields(0).Value), rstTransactions.Fields(0).Value, 0)
                                'If rstTransactions.State = ADODB.ObjectStateEnum.adStateOpen Then rstTransactions.Close()
                                'rstTransactions = Nothing

                                'If strRetval = "Y" And intTransactionCount = 0 Then
                                '    ' APNA PROCEDURE
                                'End If
                                '================================================

                                If DataExist("select total_amount from SalesChallan_dtl where location_code='" & Trim(txtUnitCode.Text) & "' and doc_no='" & mInvNo & "' And total_amount = 0  and UNIT_CODE = '" & gstrUNITID & "'") Then
                                    CustomRollbackTrans()
                                    Msgbox("Kindly EDIT/UPDATE the Invoice Again", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "eMPro")

                                    Exit Sub
                                End If

                                If DataExist("select doc_no from Supplementaryinv_hdr where location_code='" & Trim(txtUnitCode.Text) & "' and doc_no='" & mInvNo & "' and UNIT_CODE = '" & gstrUNITID & "'") Then
                                    CustomRollbackTrans()
                                    Msgbox("Already Exist with the same Number , Please Try Again  ", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "eMPro")
                                    Exit Sub
                                End If
                                If Not strRetval = "Y" Then 'jimmycarter
                                    Try 'BECAUSE FINANCE DLL IS ROLLBACKING THE TRANSACTION
                                        CustomRollbackTrans()
                                    Catch ex As Exception
                                    End Try
                                    Msgbox(strRetval, MsgBoxStyle.Information, "empower")
                                    Exit Sub
                                Else
                                    Call Logging_Starting_End_Time("Invoice locking: Fin Post Done : Trans Count :" + GetTransCountIfAvailable().ToString(), strtime, "Saved", mInvNo)

                                    '10910711
                                    If DataExist("SELECT TOP 1 1 FROM SALECONF  WHERE UNIT_CODE='" + gstrUNITID + "'  AND FIN_START_DATE <= GETDATE() AND FIN_END_DATE >= GETDATE() AND ATN_ENABLED =1  and INVOICE_TYPE='" & lbldescription.Text & "' AND SUB_TYPE='" & lblcategory.Text & "'") Then
                                        mP_Connection.Execute("Exec FA_AUTO_ATN_POSTING  '" & mInvNo & "', '" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                        Dim objComm As New ADODB.Command
                                        With objComm
                                            .ActiveConnection = mP_Connection
                                            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                                            .CommandText = "USP_FIN_JV"
                                            .CommandTimeout = 0
                                            .Parameters.Append(.CreateParameter("@IP_ADDRESS", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, gstrIpaddressWinSck))
                                            .Parameters.Append(.CreateParameter("@UNIT_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                                            .Parameters.Append(.CreateParameter("@TMP_INVOICE_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Ctlinvoice.Text))
                                            .Parameters.Append(.CreateParameter("@INVOICE_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, mInvNo.ToString))
                                            .Parameters.Append(.CreateParameter("@MODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 1, "A"))
                                            .Parameters.Append(.CreateParameter("@ERR", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInputOutput, 800, 0))

                                            .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                            If .Parameters(.Parameters.Count - 1).Value.ToString().Trim.Length <> 0 Then
                                                CustomRollbackTrans()
                                                Msgbox("Unable To Do ATN .", MsgBoxStyle.Information, ResolveResString(100))
                                                Exit Sub
                                            End If
                                        End With
                                        objComm = Nothing
                                    End If
                                    '10910711

                                    Call Logging_Starting_End_Time("Invoice locking: FA_AUTO_ATN_POSTING : Trans Count :" + GetTransCountIfAvailable().ToString(), strtime, "Saved", mInvNo)

                                    '10736222
                                    If DataExist("SELECT TOP 1 1 FROM CT2_INVOICE_KNOCKOFF WHERE UNIT_CODE='" + gstrUNITID + "' and TMP_INVOICE_NO='" & Ctlinvoice.Text & "'") Then
                                        mP_Connection.Execute("UPDATE CT2_INVOICE_KNOCKOFF SET ACT_INV_NO= '" & mInvNo & "' WHERE UNIT_CODE='" + gstrUNITID + "' AND   TMP_INVOICE_NO = " & Trim(Ctlinvoice.Text), , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                    End If
                                    '10736222

                                    ''''10277476 
                                    blnDSTracking = CBool(Find_Value("SELECT isnull(DSWiseTracking,0)as DSWiseTracking FROM sales_parameter (nolock) WHERE UNIT_CODE='" + gstrUNITID + "'"))
                                    'If (((UCase(Trim(cmbInvType.Text))) = "TRANSFER INVOICE") Or ((UCase(Trim(cmbInvType.Text)) = "NORMAL INVOICE") And (UCase(CStr((Trim(CmbCategory.Text)) = "FINISHED GOODS")))) Then
                                    If (((UCase(Trim(cmbInvType.Text))) = "TRANSFER INVOICE")) Or ((UCase(Trim(cmbInvType.Text)) = "NORMAL INVOICE") And (UCase(CStr((Trim(CmbCategory.Text)) = "FINISHED GOODS")))) Or (((UCase(Trim(cmbInvType.Text))) = "INTER-DIVISION")) Then
                                        If blnDSTracking = True And AllowTextFileGeneration_SMIIEL(strAccountCode) = True And CBool(Find_Value("SELECT SMIEL_FTP_INVOICE FROM SALES_PARAMETER WHERE UNIT_CODE='" + gstrUNITID + "'")) = True Then
                                            If CheckInvoices(mInvNo, strAccountCode) = False Then
                                                CustomRollbackTrans()
                                                Exit Sub
                                            End If
                                        End If
                                    End If
                                    ''''10277476 


                                    If AllowASNTextFileGeneration(strAccountCode) = True And blnAllow_ASNFlag = True Then
                                        mP_Connection.Execute("UPDATE MKT_ASN_INVDTL SET DOC_NO=" & Trim(mInvNo) & " where dOC_NO=" & Trim(Ctlinvoice.Text) & " and UNIT_CODE = '" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                        'NORMAL INVOICE 
                                        If (UCase(cmbInvType.Text) = "NORMAL INVOICE" And UCase(CmbCategory.Text) = "FINISHED GOODS") Then
                                            If FordASNFileGeneration(mInvNo, strAccountCode) = False Then
                                                CustomRollbackTrans()
                                                Exit Sub
                                            Else
                                                If Len(mstrupdateASNdtl) > 0 Then
                                                    mP_Connection.Execute(mstrupdateASNdtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                                    mP_Connection.Execute(mstrupdateASNCumFig, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                                End If
                                            End If
                                        End If
                                        'REJECTION INVOICE 
                                        If (UCase(cmbInvType.Text) = "REJECTION") Then
                                            If FordASNFileGeneration_Rejection(mInvNo, strAccountCode) = False Then
                                                CustomRollbackTrans()
                                                Exit Sub
                                            Else
                                                If Len(mstrupdateASNdtl) > 0 Then
                                                    mP_Connection.Execute(mstrupdateASNdtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                                    mP_Connection.Execute(mstrupdateASNCumFig, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                                End If
                                            End If
                                        End If
                                    End If
                                    '****101041360

                                    Call Logging_Starting_End_Time("Invoice locking: ASN Started : Trans Count :" + GetTransCountIfAvailable().ToString(), strtime, "Saved", mInvNo)

                                    If AllowASNTextFile(strAccountCode) = True And blnAllow_ASNFlag = True Then
                                        If (UCase(cmbInvType.Text) = "NORMAL INVOICE" And UCase(CmbCategory.Text) = "FINISHED GOODS") Then
                                            If ASNTEXTFILE_DETAILS(mInvNo, strAccountCode) = False Then
                                                CustomRollbackTrans()
                                                Exit Sub
                                            Else

                                            End If
                                        End If
                                    End If
                                    '****101041360
                                    If gstrUNITID = "STH" Then
                                        STRASNTYPE = Find_Value("SELECT ASN_TYPE FROM customer_mst WHERE  UNIT_CODE = '" & gstrUNITID & "' AND  CUSTOMER_CODE='" & strAccountCode & "'")
                                        If AllowASNTextFile(strAccountCode) = True And UCase(STRASNTYPE) = "GEN_MOTORS" Then
                                            If (UCase(cmbInvType.Text) = "NORMAL INVOICE" And UCase(CmbCategory.Text) = "FINISHED GOODS") Then
                                                If GENMOTORSASNFileGeneration_THAI(mInvNo, strAccountCode) = False Then
                                                    CustomRollbackTrans()
                                                    Exit Sub
                                                Else
                                                    If Len(mstrupdateASNdtl) > 0 Then
                                                        mP_Connection.Execute(mstrupdateASNdtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                                        mP_Connection.Execute(mstrupdateASNCumFig, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                                    End If
                                                End If
                                            End If

                                        End If
                                    End If

                                    Call Logging_Starting_End_Time("Invoice locking: Global Tool Updation : Trans Count :" + GetTransCountIfAvailable().ToString(), strtime, "Saved", mInvNo)

                                    ''ADDED BY VINOD FOR GLOBAL TOOL CHANGES
                                    If (UCase(cmbInvType.Text) = "NORMAL INVOICE" Or UCase(cmbInvType.Text) = "TRANSFER INVOICE") And UCase(CmbCategory.Text) = "ASSETS" Then
                                        If UpdateGlobalTool(mInvNo, Val(Ctlinvoice.Text)) = False Then
                                            CustomRollbackTrans()
                                            Exit Sub
                                        End If
                                    End If
                                    ''
                                    mstrInvRejSQL = ""
                                    If CBool(Find_Value("Select REJINV_Tracking from Sales_Parameter WHERE UNIT_CODE='" + gstrUNITID + "'")) = True Then
                                        mstrInvRejSQL = "Update MKT_INVREJ_DTL Set Invoice_No='" & mInvNo & "' WHERE UNIT_CODE='" + gstrUNITID + "' AND  Invoice_No='" & Trim(Ctlinvoice.Text) & "'"
                                    End If

                                    If Len(Trim(mstrInvRejSQL)) <> 0 Then
                                        mP_Connection.Execute(mstrInvRejSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                    End If
                                    'mP_Connection.CommitTrans()
                                    If DataExist("select top 1 1 from sales_parameter (nolock) where  unit_code='" & gstrUNITID & "' and bln_Trfinv_GateEntry_barcodeimg =1") Then
                                        If optInvYes(0).Checked = True And (UCase(Trim(cmbInvType.Text)) = "TRANSFER INVOICE" Or UCase(Trim(cmbInvType.Text)) = "NORMAL INVOICE" Or UCase(Trim(cmbInvType.Text)) = "INTER-DIVISION") Then
                                            'strSQL = "select top  1 1 from ASN_Lebels_temp_2D where vendorcode ='" & gstrUNITID & "' and invoiceno ='" & mInvNo & "'"
                                            '' mayur''If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSQL)) = True Then
                                            If DataExist("select top  1 1 from ASN_Lebels_temp_2D where vendorcode ='" & gstrUNITID & "' and invoiceno ='" & mInvNo & "'") Then
                                                strGateEntryBarCode = Find_Value("select gestring  from ASN_Lebels_temp_2D where vendorcode ='" & gstrUNITID & "' and invoiceno ='" & mInvNo & "'")
                                                'strGateentrypath = "c:\GateentryBarcodeimg"

                                                strGateentrypath = gstrUserMyDocPath & "\GateentryBarcodeimg"
                                                Dim BarcodeImage As Bitmap = create2DImage(GetEncodedString(strGateEntryBarCode.Trim(), ""))
                                                BarcodeImage.Save(strGateentrypath & ".JPEG")
                                                If Dir(strGateentrypath & ".txt") <> "" Then Kill(strGateentrypath & ".txt")
                                                ts = fso.OpenTextFile(strGateentrypath & ".txt", Scripting.IOMode.ForWriting, True)
                                                ts.Write(strGateEntryBarCode)
                                                ts.Close()
                                                ts = Nothing

                                                If SaveGateEntryBarcodeImage_singlelevelso_2DBARCODE(mInvNo, strGateentrypath, Mid(strGateEntryBarCode, 3)) = False Then
                                                    CustomRollbackTrans()
                                                    Msgbox("Problem While saving Gate Entry Barcode Image.", vbInformation, ResolveResString(100))

                                                    Exit Sub
                                                End If
                                                'strbinpath = "c:\BinningBarcodeImg"
                                                strbinpath = gstrUserMyDocPath & "\BinningBarcodeImg"
                                                strBinningBarCode = Find_Value("select BINNString  from ASN_Lebels_temp_2D where vendorcode ='" & gstrUNITID & "' and invoiceno ='" & mInvNo & "'")
                                                If strBinningBarCode <> "" Then
                                                    Dim BarcodeImage1 As Bitmap = create2DImage(GetEncodedString(strBinningBarCode.Trim(), ""))
                                                    BarcodeImage1.Save(strbinpath & ".JPEG")
                                                    If Dir(strbinpath & ".txt") <> "" Then Kill(strbinpath & ".txt")
                                                    ts = fso.OpenTextFile(strbinpath & ".txt", Scripting.IOMode.ForWriting, True)
                                                    ts.Write(strBinningBarCode)
                                                    ts.Close()
                                                    ts = Nothing

                                                    If SaveBinningBarcodeImage_singlelevelso_2DBARCODE(mInvNo, strbinpath, Mid(strBinningBarCode, 3)) = False Then
                                                        CustomRollbackTrans()
                                                        Msgbox("Problem While saving Gate Entry Barcode Image.", vbInformation, ResolveResString(100))

                                                        Exit Sub
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                    '101375632

                                    Call Logging_Starting_End_Time("Invoice locking: Checking Pallet Status : Trans Count :" + GetTransCountIfAvailable().ToString(), strtime, "Saved", mInvNo)

                                    If GetPaletteStatus(cmbInvType.Text, CmbCategory.Text, strAccountCode) Then
                                        Dim objCommand As New ADODB.Command
                                        With objCommand
                                            .ActiveConnection = mP_Connection
                                            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                                            .CommandText = "USP_NORMAL_FG_INVOICE_BARCODE_LOCK"
                                            .CommandTimeout = 0
                                            .Parameters.Append(.CreateParameter("@UNIT_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                                            .Parameters.Append(.CreateParameter("@USER_ID", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 50, mP_User))
                                            .Parameters.Append(.CreateParameter("@TEMP_INVOICE_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Ctlinvoice.Text))
                                            .Parameters.Append(.CreateParameter("@DOC_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, mInvNo.ToString))
                                            .Parameters.Append(.CreateParameter("@MESSAGE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInputOutput, 500, 0))

                                            .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                            If .Parameters(.Parameters.Count - 1).Value.ToString().Trim.Length <> 0 Then
                                                CustomRollbackTrans()
                                                Msgbox(.Parameters(.Parameters.Count - 1).Value.ToString(), MsgBoxStyle.Information, ResolveResString(100))
                                                Exit Sub
                                            End If
                                        End With
                                        objCommand = Nothing
                                    End If

                                    Call Logging_Starting_End_Time("Invoice locking: Going To Commit : Trans Count :" + GetTransCountIfAvailable().ToString(), strtime, "Saved", mInvNo)
                                    mP_Connection.CommitTrans()
                                    Call Logging_Starting_End_Time("Invoice locking: Commited : Trans Count :" + GetTransCountIfAvailable().ToString(), strtime, "Saved", mInvNo)

                                    'praveen on 08/08/2017 for DC invoice to have GE barcode----Start
                                    If optInvYes(0).Checked = True And (UCase(Trim(cmbInvType.Text)) = "TRANSFER INVOICE") Then
                                        Dim oCmd1 As ADODB.Command
                                        oCmd1 = New ADODB.Command
                                        With oCmd1
                                            .ActiveConnection = mP_Connection
                                            .CommandTimeout = 0
                                            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                                            .CommandText = "PRC_INVOICEPRINTING_MATE"
                                            .Parameters.Append(.CreateParameter("@UnitCode", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                                            .Parameters.Append(.CreateParameter("@LOC_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 8, Trim(txtUnitCode.Text)))
                                            .Parameters.Append(.CreateParameter("@DOC_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, mInvNo))
                                            .Parameters.Append(.CreateParameter("@INV_TYPE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 3, Trim(Me.lbldescription.Text)))
                                            .Parameters.Append(.CreateParameter("@INV_SUBTYPE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 1, Trim(Me.lblcategory.Text)))
                                            .Parameters.Append(.CreateParameter("@IP_ADDRESS", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, gstrIpaddressWinSck.Trim()))
                                            .Parameters.Append(.CreateParameter("@ERRCODE", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInputOutput))
                                            .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                        End With
                                        If oCmd1.Parameters("@ERRCODE").Value <> 0 Then
                                            Msgbox("Invoice generated.Please try Reprint.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, My.Resources.resEmpower.STR100)
                                            oCmd1 = Nothing
                                            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                                            Exit Sub
                                        End If
                                        oCmd1 = Nothing
                                    End If
                                    'praveen on 08/08/2017 for DC invoice to have GE barcode----End

                                    Call Logging_Starting_End_Time("Invoice locking", strtime, "Saved", mInvNo)
                                    Msgbox("Invoice has been locked successfully with number " & mInvNo, MsgBoxStyle.Information, "empower")
                                    txtlorryno.Text = ""
                                    '10736222
                                    'If DataExist("SELECT TOP 1 1 FROM CT2_INVOICE_KNOCKOFF WHERE UNIT_CODE='" & gstrUNITID & "' and ACT_INV_NO='" & mInvNo & "'") Then
                                    '    With objCom
                                    '        .Parameters.Clear()
                                    '        .CommandText = "USP_SENDAUTOMAILER_CT2_INVOICE"
                                    '        .CommandType = CommandType.StoredProcedure
                                    '        .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                                    '        .Parameters.Add("@INVOICE_NO", SqlDbType.VarChar, 12).Value = Trim(mInvNo)
                                    '        .Parameters.Add("@ERRMSG", SqlDbType.VarChar, 1000).Direction = ParameterDirection.Output
                                    '        SqlConnectionclass.ExecuteNonQuery(objCom)
                                    '    End With
                                    'End If
                                    ' Mayur Transfer Invoice Auto Mailer Start 20 Nov 2017

                                    If optInvYes(0).Checked = True And (UCase(Trim(cmbInvType.Text)) = "TRANSFER INVOICE") Then
                                        If DataExist("select top  1 1 from ASN_Lebels_temp_2D where vendorcode ='" & gstrUNITID & "' and invoiceno ='" & mInvNo & "'") Then

                                            Using sqlCmd As New SqlCommand
                                                With sqlCmd
                                                    .CommandText = "USP_TRANSFERINVOICE_NOTIFICATION_MAIL"
                                                    .CommandType = CommandType.StoredProcedure
                                                    .CommandTimeout = 0
                                                    .Parameters.Add("@UNITCODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                                                    .Parameters.Add("@INVOICENO", SqlDbType.VarChar, 12).Value = Trim(mInvNo)
                                                    SqlConnectionclass.ExecuteNonQuery(sqlCmd)
                                                End With
                                            End Using

                                        End If
                                    End If

                                    ' Mayur Transfer Invoice Auto Mailer End 20 Nov 2017


                                    If DataExist("SELECT TOP 1 1 FROM CT2_INVOICE_KNOCKOFF WHERE UNIT_CODE='" & gstrUNITID & "' and ACT_INV_NO='" & mInvNo & "'") Then
                                        Using sqlCmd As New SqlCommand
                                            With sqlCmd
                                                .CommandText = "USP_SENDAUTOMAILER_CT2_INVOICE"
                                                .CommandType = CommandType.StoredProcedure
                                                .CommandTimeout = 0
                                                .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                                                .Parameters.Add("@INVOICE_NO", SqlDbType.VarChar, 12).Value = Trim(mInvNo)
                                                .Parameters.Add("@ERRMSG", SqlDbType.VarChar, 1000).Direction = ParameterDirection.Output
                                                SqlConnectionclass.ExecuteNonQuery(sqlCmd)
                                            End With
                                        End Using
                                    End If

                                    ''10736222

                                    'SATISH KESHAERWANI CHANGE
                                    If chktkmlbarcode.Checked = True Then
                                        Print_barcodelabel(mInvNo)
                                    End If
                                    'SATISH KESHAERWANI CHANGE
                                    ChkCustDetails.CheckState = System.Windows.Forms.CheckState.Unchecked
                                    Ctlinvoice.Text = ""
                                    If gstrUNITID = "STH" Then
                                        strSQL = "{TMP_INVOICEPRINT.IP_ADDRESS}='" & gstrIpaddressWinSck & "'  and {TMP_INVOICEPRINT.Unit_Code}='" & gstrUNITID & "'"
                                    ElseIf IsGSTINSAME(strAccountCode) And (lbldescription.Text.Trim.ToUpper = "TRF" Or lbldescription.Text.Trim.ToUpper = "ITD") Then
                                        strSQL = "{TMP_INVOICEPRINT.IP_ADDRESS}='" & gstrIpaddressWinSck & "'  and {TMP_INVOICEPRINT.Unit_Code}='" & gstrUNITID & "'"
                                    Else
                                        If CBool(Find_Value("select REJINVOICE_GRIN_PVKNOCKING from sales_parameter WHERE UNIT_CODE='" + gstrUNITID + "'")) And UCase(Trim(cmbInvType.Text)) = "REJECTION" Then
                                            strSQL = "{TMP_INVOICEPRINT.IP_ADDRESS}='" & gstrIpaddressWinSck & "'  and {TMP_INVOICEPRINT.Unit_Code}='" & gstrUNITID & "'"
                                        Else
                                            RdAddSold.DataDefinition.RecordSelectionFormula = "{SalesChallan_Dtl.Location_Code}='" & Trim(txtUnitCode.Text) & "' and {SalesChallan_Dtl.Doc_No} =" & mInvNo & " and {SalesChallan_Dtl.UNIT_CODE} ='" & gstrUNITID & "' and {SalesChallan_Dtl.Invoice_Type}= '" & Trim(Me.lbldescription.Text) & "'  and {SalesChallan_Dtl.Sub_Category} = '" & Trim(Me.lblcategory.Text) & "'"
                                        End If
                                    End If
                                End If
                                Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.WaitCursor)
                                For intLoopCounter = 1 To intMaxLoop
                                    Select Case intLoopCounter
                                        Case 1
                                            If optInvYes(0).Checked = True Then
                                                RdAddSold.DataDefinition.FormulaFields("CopyName").Text = "'ORIGINAL FOR BUYER'"
                                            Else
                                                RdAddSold.DataDefinition.FormulaFields("CopyName").Text = "'ORIGINAL FOR BUYER (REPRINT)'"
                                            End If
                                        Case 2
                                            If optInvYes(0).Checked = True Then
                                                RdAddSold.DataDefinition.FormulaFields("CopyName").Text = "'DUPLICATE FOR TRANSPORTER'"
                                            Else
                                                RdAddSold.DataDefinition.FormulaFields("CopyName").Text = "'DUPLICATE FOR TRANSPORTER (REPRINT)'"
                                            End If
                                        Case 3
                                            If optInvYes(0).Checked = True Then
                                                RdAddSold.DataDefinition.FormulaFields("CopyName").Text = "'TRIPLICATE FOR ASSESSEE'"
                                            Else
                                                RdAddSold.DataDefinition.FormulaFields("CopyName").Text = "'TRIPLICATE FOR ASSESSEE (REPRINT)'"
                                            End If
                                        Case 4
                                            If optInvYes(0).Checked = True Then
                                                RdAddSold.DataDefinition.FormulaFields("CopyName").Text = "'EXTRA COPY '"
                                            Else
                                                RdAddSold.DataDefinition.FormulaFields("CopyName").Text = "'EXTRA COPY (REPRINT)'"
                                            End If

                                    End Select
                                    COPYNAME_New(0) = RdAddSold.DataDefinition.FormulaFields("CopyName").Text
                                    COPYNAME_New(1) = "Y"
                                    RdAddSold.DataDefinition.FormulaFields("InsuranceFlag").Text = "'" & mblnInsuranceFlag & "'"
                                    If AllowASNPrinting(strAccountCode) = False Then
                                        Frm.SetReportDocument()
                                        If mblnEwaybill_Print = False Then
                                            ''Praveen Digital Sign Changes 
                                            If mblnISCrystalReportRequired Then
                                                If COPYNAME_New(1) = "Y" Then
                                                    RdAddSold.PrintToPrinter(1, False, 0, 0)
                                                End If
                                            End If

                                            ''AMIT RANA 20 Jun 2019    
                                            If (blnIsPDFExported = False) Then
                                                EXPORTINVOICETOPDF_ONPRINTREPRINT(strAccountCode, lbldescription.Text, lblcategory.Text, RdAddSold, DT_A4CUSTOMER_INVOICEPRINTINGTAG)
                                                blnIsPDFExported = True
                                            End If
                                            ''AMIT RANA 20 Jun 2019 

                                        Else
                                            If chkprintreprint.Checked = True And optInvYes(1).Checked = True Then
                                                ''Praveen Digital Sign Changes 
                                                If mblnISCrystalReportRequired Then
                                                    If COPYNAME_New(1) = "Y" Then
                                                        RdAddSold.PrintToPrinter(1, False, 0, 0)
                                                    End If
                                                End If


                                                ''AMIT RANA 20 Jun 2019    
                                                If (blnIsPDFExported = False) Then
                                                    EXPORTINVOICETOPDF_ONPRINTREPRINT(strAccountCode, lbldescription.Text, lblcategory.Text, RdAddSold, DT_A4CUSTOMER_INVOICEPRINTINGTAG)
                                                    blnIsPDFExported = True
                                                End If
                                                ''AMIT RANA 20 Jun 2019 
                                            Else
                                                If optInvYes(1).Checked = True Then
                                                    If Not DataExist("SELECT TOP 1 1 FROM FIRSTTIME_INVOICEPRINTING  WHERE UNIT_CODE = '" & gstrUNITID & "' and doc_no= " & Trim(Ctlinvoice.Text) & "") = True Then
                                                        mP_Connection.Execute("Insert into FIRSTTIME_INVOICEPRINTING(unit_code,doc_no,ent_dt,ent_userid) values('" & gstrUNITID & "','" & Trim$(Ctlinvoice.Text) & "',getdate(),'" & mP_User & "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                                        ''Praveen Digital Sign Changes 
                                                        If mblnISCrystalReportRequired Then
                                                            If COPYNAME_New(1) = "Y" Then
                                                                RdAddSold.PrintToPrinter(1, False, 0, 0)
                                                            End If
                                                        End If
                                                    Else
                                                        If mblnISCrystalReportRequired Then   ''Praveen Digital Sign Changes 
                                                            If COPYNAME_New(1) = "Y" Then
                                                                RdAddSold.PrintToPrinter(1, False, 0, 0)
                                                            End If
                                                        End If
                                                    End If


                                                    ''AMIT RANA 20 Jun 2019    
                                                    If (blnIsPDFExported = False) Then
                                                        EXPORTINVOICETOPDF_ONPRINTREPRINT(strAccountCode, lbldescription.Text, lblcategory.Text, RdAddSold, DT_A4CUSTOMER_INVOICEPRINTINGTAG)
                                                        blnIsPDFExported = True
                                                    End If
                                                    ''AMIT RANA 20 Jun 2019 
                                                End If
                                            End If
                                        End If

                                    End If

                                Next
                                'ONLY FOR TOYOTA CUSTOMER FOR BATE BGLORE
                                If mblncustomerlevel_A4report_functionlity = True And mblnA4reports_invoicewise = True And mblncustomerlevel_Annexture_printing = True And mblncustomerspecificreport = True Then
                                    Dim strTOTALQty As String
                                    Dim Strquery As String
                                    Dim strpdsno As String
                                    Dim intLoopCounters As Short
                                    RdAddSold.DataDefinition.FormulaFields("TOYOTABARCODEPRINT").Text = False
                                    If DataExist("SELECT TOP 1 1 FROM SALES_DTL  WHERE UNIT_CODE = '" & gstrUNITID & "' and doc_no= " & Trim(mInvNo) & " HAVING COUNT(*)> " & mintmaxnoofitems_barcodeToyota) = True Then
                                        RdAddSold.DataDefinition.FormulaFields("TOYOTABARCODEPRINT").Text = True
                                    End If
                                    RdAddSold.DataDefinition.FormulaFields("ANNEXTURE").Text = False
                                    Strquery = "SELECT TOP 1 1 FROM SALES_DTL  WHERE UNIT_CODE = '" & gstrUNITID & "' and doc_no= " & Trim(mInvNo) & " HAVING COUNT(*)> " & mintmaxnoofitems_Annexture
                                    RdAddSold = Frm.GetReportDocument()
                                    RdAddSold.Load(My.Application.Info.DirectoryPath & "\Reports\Annextureprinting_A4reports.rpt")
                                    RdAddSold.DataDefinition.FormulaFields("Invoiceno").Text = "'" & mInvNo & "'"
                                    strpdsno = Find_Value("SELECT LORRYNO_DATE FROM SALESCHALLAN_DTL WHERE  UNIT_CODE = '" & gstrUNITID & "' AND  DOC_NO=" & Trim(mInvNo))
                                    RdAddSold.DataDefinition.FormulaFields("pdsno").Text = "'" & strpdsno & "'"
                                    strSQL = "{SalesChallan_Dtl.Location_Code}='" & Trim(txtUnitCode.Text) & "' and {SalesChallan_Dtl.Doc_No} =" & Trim(mInvNo) & " and {SalesChallan_Dtl.UNIT_CODE} ='" & gstrUNITID & "' "
                                    RdAddSold.DataDefinition.RecordSelectionFormula = strSQL


                                    If DataExist(Strquery) = True Then
                                        Frm.SetReportDocument()
                                        'For intLoopCounters = 0 To 1
                                        For intLoopCounters = 1 To intMaxLoop
                                            If mblnEwaybill_Print = False Then
                                                ''Praveen Digital Sign Changes 
                                                If mblnISCrystalReportRequired Then
                                                    If COPYNAME_New(1) = "Y" Then
                                                        RdAddSold.PrintToPrinter(1, False, 0, 0)
                                                    End If
                                                End If

                                                ''AMIT RANA 20 Jun 2019    
                                                'If (blnIsPDFExported = False) Then
                                                '    EXPORTINVOICETOPDF_ONPRINTREPRINT(strAccountCode, lbldescription.Text, lblcategory.Text, RdAddSold)
                                                '    blnIsPDFExported = True
                                                'End If
                                                ''AMIT RANA 20 Jun 2019 

                                            Else
                                                If chkprintreprint.Checked = True And optInvYes(1).Checked = True Then
                                                    ''Praveen Digital Sign Changes 
                                                    If mblnISCrystalReportRequired Then
                                                        If COPYNAME_New(1) = "Y" Then
                                                            RdAddSold.PrintToPrinter(1, False, 0, 0)
                                                        End If
                                                    End If

                                                    ''AMIT RANA 20 Jun 2019    
                                                    'If (blnIsPDFExported = False) Then
                                                    '    EXPORTINVOICETOPDF_ONPRINTREPRINT(strAccountCode, lbldescription.Text, lblcategory.Text, RdAddSold)
                                                    '    blnIsPDFExported = True
                                                    'End If
                                                    ''AMIT RANA 20 Jun 2019 

                                                Else
                                                    If optInvYes(1).Checked = True Then
                                                        If Not DataExist("SELECT TOP 1 1 FROM FIRSTTIME_INVOICEPRINTING  WHERE UNIT_CODE = '" & gstrUNITID & "' and doc_no= " & Trim(Ctlinvoice.Text) & "") = True Then
                                                            mP_Connection.Execute("Insert into FIRSTTIME_INVOICEPRINTING(unit_code,doc_no,ent_dt,ent_userid) values('" & gstrUNITID & "','" & Trim$(Ctlinvoice.Text) & "',getdate(),'" & mP_User & "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                                            If mblnISCrystalReportRequired Then  ''Praveen Digital Sign Changes 
                                                                If COPYNAME_New(1) = "Y" Then
                                                                    RdAddSold.PrintToPrinter(1, False, 0, 0)
                                                                End If
                                                            End If
                                                        Else
                                                            If mblnISCrystalReportRequired Then  ''Praveen Digital Sign Changes 
                                                                If COPYNAME_New(1) = "Y" Then
                                                                    RdAddSold.PrintToPrinter(1, False, 0, 0)
                                                                End If
                                                            End If
                                                        End If


                                                        ''AMIT RANA 20 Jun 2019    
                                                        'If (blnIsPDFExported = False) Then
                                                        '    EXPORTINVOICETOPDF_ONPRINTREPRINT(strAccountCode, lbldescription.Text, lblcategory.Text, RdAddSold)
                                                        '    blnIsPDFExported = True
                                                        'End If
                                                        ''AMIT RANA 20 Jun 2019 
                                                    End If


                                                End If
                                            End If
                                        Next
                                    End If

                                End If
                                '30 aug 2022 maruti toyota anexture changes
                                If mblncustomerspecificreport = False And mblnA4reports_invoicewise = True And mblncustomerlevel_Annexture_printing = True And gstrUNITID = "MST" Then
                                    Dim strTOTALQty As String
                                    Dim Strquery As String
                                    Dim strpdsno As String
                                    Dim intLoopCounters As Short
                                    RdAddSold.DataDefinition.FormulaFields("ANNEXTURE").Text = False
                                    Strquery = "SELECT TOP 1 1 FROM SALES_DTL  WHERE UNIT_CODE = '" & gstrUNITID & "' and doc_no= " & Trim(mInvNo) & " HAVING COUNT(*)> " & mintmaxnoofitems_Annexture
                                    RdAddSold = Frm.GetReportDocument()
                                    RdAddSold.Load(My.Application.Info.DirectoryPath & "\Reports\Annextureprinting_A4reports.rpt")
                                    RdAddSold.DataDefinition.FormulaFields("Invoiceno").Text = "'" & mInvNo & "'"
                                    strpdsno = Find_Value("SELECT LORRYNO_DATE FROM SALESCHALLAN_DTL WHERE  UNIT_CODE = '" & gstrUNITID & "' AND  DOC_NO=" & Trim(mInvNo))
                                    RdAddSold.DataDefinition.FormulaFields("pdsno").Text = "'" & strpdsno & "'"
                                    strSQL = "{SalesChallan_Dtl.Location_Code}='" & Trim(txtUnitCode.Text) & "' and {SalesChallan_Dtl.Doc_No} =" & Trim(mInvNo) & " and {SalesChallan_Dtl.UNIT_CODE} ='" & gstrUNITID & "' "
                                    RdAddSold.DataDefinition.RecordSelectionFormula = strSQL


                                    If DataExist(Strquery) = True Then
                                        Frm.SetReportDocument()
                                        'For intLoopCounters = 0 To 1
                                        For intLoopCounters = 1 To intMaxLoop
                                            If mblnEwaybill_Print = False Then
                                                If mblnISCrystalReportRequired Then   ''Praveen Digital Sign Changes 
                                                    If COPYNAME_New(1) = "Y" Then
                                                        RdAddSold.PrintToPrinter(1, False, 0, 0)
                                                    End If
                                                End If

                                                ''AMIT RANA 20 Jun 2019    
                                                'If (blnIsPDFExported = False) Then
                                                '    EXPORTINVOICETOPDF_ONPRINTREPRINT(strAccountCode, lbldescription.Text, lblcategory.Text, RdAddSold)
                                                '    blnIsPDFExported = True
                                                'End If
                                                ''AMIT RANA 20 Jun 2019 

                                            Else
                                                If chkprintreprint.Checked = True And optInvYes(1).Checked = True Then
                                                    If mblnISCrystalReportRequired Then   ''Praveen Digital Sign Changes 
                                                        If COPYNAME_New(1) = "Y" Then
                                                            RdAddSold.PrintToPrinter(1, False, 0, 0)
                                                        End If
                                                    End If

                                                    ''AMIT RANA 20 Jun 2019    
                                                    'If (blnIsPDFExported = False) Then
                                                    '    EXPORTINVOICETOPDF_ONPRINTREPRINT(strAccountCode, lbldescription.Text, lblcategory.Text, RdAddSold)
                                                    '    blnIsPDFExported = True
                                                    'End If
                                                    ''AMIT RANA 20 Jun 2019 

                                                Else
                                                    If optInvYes(1).Checked = True Then
                                                        If Not DataExist("SELECT TOP 1 1 FROM FIRSTTIME_INVOICEPRINTING  WHERE UNIT_CODE = '" & gstrUNITID & "' and doc_no= " & Trim(Ctlinvoice.Text) & "") = True Then
                                                            mP_Connection.Execute("Insert into FIRSTTIME_INVOICEPRINTING(unit_code,doc_no,ent_dt,ent_userid) values('" & gstrUNITID & "','" & Trim$(Ctlinvoice.Text) & "',getdate(),'" & mP_User & "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                                            If mblnISCrystalReportRequired Then   ''Praveen Digital Sign Changes 
                                                                If COPYNAME_New(1) = "Y" Then
                                                                    RdAddSold.PrintToPrinter(1, False, 0, 0)
                                                                End If
                                                            End If
                                                        Else
                                                            If mblnISCrystalReportRequired Then   ''Praveen Digital Sign Changes 
                                                                If COPYNAME_New(1) = "Y" Then
                                                                    RdAddSold.PrintToPrinter(1, False, 0, 0)
                                                                End If
                                                            End If
                                                        End If


                                                        ''AMIT RANA 20 Jun 2019    
                                                        'If (blnIsPDFExported = False) Then
                                                        '    EXPORTINVOICETOPDF_ONPRINTREPRINT(strAccountCode, lbldescription.Text, lblcategory.Text, RdAddSold)
                                                        '    blnIsPDFExported = True
                                                        'End If
                                                        ''AMIT RANA 20 Jun 2019 
                                                    End If


                                                End If
                                            End If
                                        Next
                                    End If

                                End If
                                '30 aug 2022 maruti toyota annexture changes
                                'TOYOTA BARCODE 
                                'If mblncustomerlevel_A4report_functionlity = True And mblnA4reports_invoicewise = True And mblncustomerlevel_Annexture_printing = True Then

                                '    Dim strTOTALQty As String
                                '    Dim Strquery As String
                                '    Dim strpdsno As String
                                '    Dim intLoopCounters As Short
                                '    Strquery = "SELECT TOP 1 1 FROM SALES_DTL  WHERE UNIT_CODE = '" & gstrUNITID & "' and doc_no= " & Trim(mInvNo) & " HAVING COUNT(*)> " & mintmaxnoofitems_barcodeToyota

                                '    If DataExist(Strquery) = True Then
                                '        RdAddSold = Frm.GetReportDocument()
                                '        RdAddSold.Load(My.Application.Info.DirectoryPath & "\Reports\BarcodeToyotaprinting.rpt")
                                '        RdAddSold.DataDefinition.FormulaFields("Invoiceno").Text = "'" & mInvNo & "'"
                                '        strpdsno = Find_Value("SELECT LORRYNO_DATE FROM SALESCHALLAN_DTL WHERE  UNIT_CODE = '" & gstrUNITID & "' AND  DOC_NO=" & Trim(mInvNo))
                                '        RdAddSold.DataDefinition.FormulaFields("pdsno").Text = "'" & strpdsno & "'"
                                '        strSQL = "{SalesChallan_Dtl.Location_Code}='" & Trim(txtUnitCode.Text) & "' and {SalesChallan_Dtl.Doc_No} =" & Trim(mInvNo) & " and {SalesChallan_Dtl.UNIT_CODE} ='" & gstrUNITID & "' "
                                '        RdAddSold.DataDefinition.RecordSelectionFormula = strSQL

                                '        Frm.SetReportDocument()

                                '        For intLoopCounters = 0 To 1
                                '            RdAddSold.PrintToPrinter(1, False, 0, 0)
                                '        Next
                                '    End If

                                'End If

                                ''TOYOTA BARCODE 
                                Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
                            Else
                                If AllowBarCodePrinting(strAccountCode) = True And optInvYes(0).Checked = True Then
                                    If DeleteBarCodeImage(Me.Ctlinvoice.Text) = False Then
                                        Msgbox("Problem While deleting Barcode Image.", vbInformation, ResolveResString(100))
                                        Exit Sub
                                    End If
                                End If
                            End If
                        Else
                            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.WaitCursor)
                            If AllowASNPrinting(strAccountCode) = True Then
                                If mblnASNExist = True Then
                                    RdAddSold.DataDefinition.FormulaFields("CopyName").Text = "'EXTRA COPY'"
                                    COPYNAME_New(0) = RdAddSold.DataDefinition.FormulaFields("CopyName").Text
                                    COPYNAME_New(1) = "Y"
                                Else
                                    mP_Connection.Execute("Insert into CreatedASN values('" & Trim$(Me.Ctlinvoice.Text) & "','" & Trim$(txtASNNumber.Text) & "',getdate(),'" & mP_User & "',getdate(),'" & gstrUNITID & "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                End If
                            End If
                            '*************
                            ''If mblnA4reports_invoicewise = True And mblncustomerlevel_A4report_functionlity = True Then
                            If mblncustomerlevel_A4report_functionlity = True Then

                                For intLoopCounter = 1 To intMaxLoop
                                    If mblnEwaybill_Print = False Then
                                        If optInvYes(0).Checked = True Then
                                            ' COPYNAME = Find_Value("Select TEXTHEADING FROM  A4CUSTOMER_INVOICEPRINTINGTAG WHERE UNIT_CODE='" + gstrUNITID + "' AND CUSTOMER_CODE='" & strAccountCode & "' AND ORIGINAL_REPRINT='O' AND SERIALNO=" & intLoopCounter)
                                            Dim DataRowFiltered() As DataRow = DT_A4CUSTOMER_INVOICEPRINTINGTAG.Select("ORIGINAL_REPRINT='O' and SERIALNO=" + intLoopCounter.ToString())
                                            COPYNAME_New(0) = DataRowFiltered(0)("TEXTHEADING").ToString()
                                            COPYNAME_New(1) = DataRowFiltered(0)("HARDCOPYPRINTREQUIRED").ToString()
                                        Else
                                            'COPYNAME = Find_Value("Select TEXTHEADING FROM  A4CUSTOMER_INVOICEPRINTINGTAG WHERE UNIT_CODE='" + gstrUNITID + "' AND CUSTOMER_CODE='" & strAccountCode & "' AND ORIGINAL_REPRINT='R' AND SERIALNO=" & intLoopCounter)
                                            Dim DataRowFiltered() As DataRow = DT_A4CUSTOMER_INVOICEPRINTINGTAG.Select("ORIGINAL_REPRINT='R' and SERIALNO=" + intLoopCounter.ToString())
                                            COPYNAME_New(0) = DataRowFiltered(0)("TEXTHEADING").ToString()
                                            COPYNAME_New(1) = DataRowFiltered(0)("HARDCOPYPRINTREQUIRED").ToString()
                                        End If
                                    Else
                                        If chkprintreprint.Checked = True And optInvYes(1).Checked = True Then
                                            'COPYNAME = Find_Value("Select TEXTHEADING FROM  A4CUSTOMER_INVOICEPRINTINGTAG WHERE UNIT_CODE='" + gstrUNITID + "' AND CUSTOMER_CODE='" & strAccountCode & "' AND ORIGINAL_REPRINT='R' AND SERIALNO=" & intLoopCounter)
                                            Dim DataRowFiltered() As DataRow = DT_A4CUSTOMER_INVOICEPRINTINGTAG.Select("ORIGINAL_REPRINT='R' and SERIALNO=" + intLoopCounter.ToString())
                                            COPYNAME_New(0) = DataRowFiltered(0)("TEXTHEADING").ToString()
                                            COPYNAME_New(1) = DataRowFiltered(0)("HARDCOPYPRINTREQUIRED").ToString()
                                        ElseIf chkprintreprint.Checked = False And optInvYes(1).Checked = True Then
                                            'COPYNAME = Find_Value("Select TEXTHEADING FROM  A4CUSTOMER_INVOICEPRINTINGTAG WHERE UNIT_CODE='" + gstrUNITID + "' AND CUSTOMER_CODE='" & strAccountCode & "' AND ORIGINAL_REPRINT='O' AND SERIALNO=" & intLoopCounter)
                                            Dim DataRowFiltered() As DataRow = DT_A4CUSTOMER_INVOICEPRINTINGTAG.Select("ORIGINAL_REPRINT='O' and SERIALNO=" + intLoopCounter.ToString())
                                            COPYNAME_New(0) = DataRowFiltered(0)("TEXTHEADING").ToString()
                                            COPYNAME_New(1) = DataRowFiltered(0)("HARDCOPYPRINTREQUIRED").ToString()
                                        Else
                                            'COPYNAME = Find_Value("Select TEXTHEADING FROM  A4CUSTOMER_INVOICEPRINTINGTAG WHERE UNIT_CODE='" + gstrUNITID + "' AND CUSTOMER_CODE='" & strAccountCode & "' AND ORIGINAL_REPRINT='R' AND SERIALNO=" & intLoopCounter)
                                            Dim DataRowFiltered() As DataRow = DT_A4CUSTOMER_INVOICEPRINTINGTAG.Select("ORIGINAL_REPRINT='R' and SERIALNO=" + intLoopCounter.ToString())
                                            COPYNAME_New(0) = DataRowFiltered(0)("TEXTHEADING").ToString()
                                            COPYNAME_New(1) = DataRowFiltered(0)("HARDCOPYPRINTREQUIRED").ToString()
                                        End If
                                    End If
                                    RdAddSold.DataDefinition.FormulaFields("CopyName").Text = "'" & COPYNAME_New(0) & "'"
                                    RdAddSold.DataDefinition.FormulaFields("InsuranceFlag").Text = "'" + CStr(mblnInsuranceFlag) + "'"
                                    If gstrUNITID = "STH" Then
                                        If intLoopCounter <= 2 Then
                                            RdAddSold.DataDefinition.FormulaFields("CUST_SIGN").Text = "'CUSTOMER'"
                                        Else
                                            RdAddSold.DataDefinition.FormulaFields("CUST_SIGN").Text = "'ACCOUNTING'"
                                        End If
                                    End If
                                    Frm.SetReportDocument()
                                    'Dim dblewaymaxvalue As Double
                                    'If optInvYes(0).Checked = True Then
                                    '    dblewaymaxvalue = Find_Value("select total_amount from saleschallan_Dtl where unit_code='" + gstrUNITID + "' and doc_no =" & mInvNo)
                                    'Else
                                    '    dblewaymaxvalue = Find_Value("select total_amount from saleschallan_Dtl where unit_code='" + gstrUNITID + "' and doc_no =" & Ctlinvoice.Text)
                                    'End If

                                    If mblnEwaybill_Print = False Then
                                        If mblnISCrystalReportRequired Then   ''Praveen Digital Sign Changes 
                                            If COPYNAME_New(1) = "Y" Then
                                                RdAddSold.PrintToPrinter(1, False, 0, 0)
                                            End If
                                        End If

                                        ''AMIT RANA 20 Jun 2019    
                                        If (blnIsPDFExported = False) Then
                                            EXPORTINVOICETOPDF_ONPRINTREPRINT(strAccountCode, lbldescription.Text, lblcategory.Text, RdAddSold, DT_A4CUSTOMER_INVOICEPRINTINGTAG)
                                            blnIsPDFExported = True
                                        End If
                                        ''AMIT RANA 20 Jun 2019 
                                    Else
                                        'If dblewaymaxvalue <= mdblewaymaximumvalue Then
                                        '    RdAddSold.PrintToPrinter(1, False, 0, 0)

                                        '    ''AMIT RANA 20 Jun 2019    
                                        '    If (blnIsPDFExported = False) Then
                                        '        EXPORTINVOICETOPDF_ONPRINTREPRINT(strAccountCode, lbldescription.Text, lblcategory.Text, RdAddSold)
                                        '        blnIsPDFExported = True
                                        '    End If
                                        ''AMIT RANA 20 Jun 2019 
                                        If chkprintreprint.Checked = True And optInvYes(1).Checked = True Then
                                            If mblnISCrystalReportRequired Then   ''Praveen Digital Sign Changes 
                                                If COPYNAME_New(1) = "Y" Then
                                                    RdAddSold.PrintToPrinter(1, False, 0, 0)
                                                End If
                                            End If


                                            ''AMIT RANA 20 Jun 2019    
                                            If (blnIsPDFExported = False) Then
                                                EXPORTINVOICETOPDF_ONPRINTREPRINT(strAccountCode, lbldescription.Text, lblcategory.Text, RdAddSold, DT_A4CUSTOMER_INVOICEPRINTINGTAG)
                                                blnIsPDFExported = True
                                            End If
                                            ''AMIT RANA 20 Jun 2019 
                                        Else
                                            If optInvYes(1).Checked = True Then
                                                If Not DataExist("SELECT TOP 1 1 FROM FIRSTTIME_INVOICEPRINTING  WHERE UNIT_CODE = '" & gstrUNITID & "' and doc_no= " & Trim(Ctlinvoice.Text) & "") = True Then
                                                    mP_Connection.Execute("Insert into FIRSTTIME_INVOICEPRINTING(unit_code,doc_no,ent_dt,ent_userid) values('" & gstrUNITID & "','" & Trim$(Ctlinvoice.Text) & "',getdate(),'" & mP_User & "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                                    If mblnISCrystalReportRequired Then  ''Praveen Digital Sign Changes 
                                                        If COPYNAME_New(1) = "Y" Then
                                                            RdAddSold.PrintToPrinter(1, False, 0, 0)
                                                        End If
                                                    End If
                                                Else
                                                    If mblnISCrystalReportRequired Then   ''Praveen Digital Sign Changes 
                                                        If COPYNAME_New(1) = "Y" Then
                                                            RdAddSold.PrintToPrinter(1, False, 0, 0)
                                                        End If
                                                    End If
                                                End If


                                                ''AMIT RANA 20 Jun 2019    
                                                If (blnIsPDFExported = False) Then
                                                    EXPORTINVOICETOPDF_ONPRINTREPRINT(strAccountCode, lbldescription.Text, lblcategory.Text, RdAddSold, DT_A4CUSTOMER_INVOICEPRINTINGTAG)
                                                    blnIsPDFExported = True
                                                End If
                                                ''AMIT RANA 20 Jun 2019 
                                            End If
                                        End If
                                    End If

                                Next
                            Else
                                For intLoopCounter = 1 To intMaxLoop
                                    Select Case intLoopCounter
                                        Case 1
                                            If optInvYes(0).Checked = True Then
                                                RdAddSold.DataDefinition.FormulaFields("CopyName").Text = "'ORIGINAL FOR BUYER'"
                                            Else
                                                If chkprintreprint.Checked = True Then
                                                    RdAddSold.DataDefinition.FormulaFields("CopyName").Text = "'ORIGINAL FOR BUYER (REPRINT)'"
                                                Else
                                                    RdAddSold.DataDefinition.FormulaFields("CopyName").Text = "'ORIGINAL FOR BUYER'"
                                                End If
                                            End If
                                        Case 2
                                            If optInvYes(0).Checked = True Then
                                                RdAddSold.DataDefinition.FormulaFields("CopyName").Text = "'DUPLICATE FOR TRANSPORTER'"
                                            Else
                                                If chkprintreprint.Checked = True Then
                                                    RdAddSold.DataDefinition.FormulaFields("CopyName").Text = "'DUPLICATE FOR TRANSPORTER (REPRINT)'"
                                                Else
                                                    RdAddSold.DataDefinition.FormulaFields("CopyName").Text = "'DUPLICATE FOR TRANSPORTER'"
                                                End If
                                            End If
                                        Case 3
                                            If optInvYes(0).Checked = True Then
                                                RdAddSold.DataDefinition.FormulaFields("CopyName").Text = "'TRIPLICATE FOR ASSESSEE'"
                                            Else
                                                If chkprintreprint.Checked = True Then
                                                    RdAddSold.DataDefinition.FormulaFields("CopyName").Text = "'TRIPLICATE FOR ASSESSEE (REPRINT)'"
                                                Else
                                                    RdAddSold.DataDefinition.FormulaFields("CopyName").Text = "'TRIPLICATE FOR ASSESSEE'"
                                                End If
                                                'RdAddSold.DataDefinition.FormulaFields("CopyName").Text = "'TRIPLICATE FOR ASSESSEE (REPRINT)'"

                                            End If
                                        Case 4
                                            If optInvYes(0).Checked = True Then
                                                RdAddSold.DataDefinition.FormulaFields("CopyName").Text = "'EXTRA COPY'"
                                            Else
                                                If chkprintreprint.Checked = True Then
                                                    RdAddSold.DataDefinition.FormulaFields("CopyName").Text = "'EXTRA COPY (REPRINT)'"
                                                Else
                                                    RdAddSold.DataDefinition.FormulaFields("CopyName").Text = "'EXTRA COPY'"
                                                End If
                                                'RdAddSold.DataDefinition.FormulaFields("CopyName").Text = "'EXTRA COPY (REPRINT)'"
                                            End If


                                    End Select
                                    COPYNAME_New(0) = RdAddSold.DataDefinition.FormulaFields("CopyName").Text
                                    COPYNAME_New(1) = "Y"
                                    RdAddSold.DataDefinition.FormulaFields("InsuranceFlag").Text = "'" & mblnInsuranceFlag & "'"
                                    Frm.SetReportDocument()
                                    If mblnEwaybill_Print = False Then
                                        If mblnISCrystalReportRequired Then   ''Praveen Digital Sign Changes 
                                            If COPYNAME_New(1) = "Y" Then
                                                RdAddSold.PrintToPrinter(1, False, 0, 0)
                                            End If
                                        End If

                                        ''AMIT RANA 20 Jun 2019    
                                        If (blnIsPDFExported = False) Then
                                            EXPORTINVOICETOPDF_ONPRINTREPRINT(strAccountCode, lbldescription.Text, lblcategory.Text, RdAddSold, DT_A4CUSTOMER_INVOICEPRINTINGTAG)
                                            blnIsPDFExported = True
                                        End If
                                        ''AMIT RANA 20 Jun 2019 
                                    Else
                                        If chkprintreprint.Checked = True And optInvYes(1).Checked = True Then
                                            If mblnISCrystalReportRequired Then   ''Praveen Digital Sign Changes 
                                                If COPYNAME_New(1) = "Y" Then
                                                    RdAddSold.PrintToPrinter(1, False, 0, 0)
                                                End If
                                            End If

                                            ''AMIT RANA 20 Jun 2019    
                                            If (blnIsPDFExported = False) Then
                                                EXPORTINVOICETOPDF_ONPRINTREPRINT(strAccountCode, lbldescription.Text, lblcategory.Text, RdAddSold, DT_A4CUSTOMER_INVOICEPRINTINGTAG)
                                                blnIsPDFExported = True
                                            End If
                                            ''AMIT RANA 20 Jun 2019 
                                        Else
                                            If optInvYes(1).Checked = True Then
                                                If Not DataExist("SELECT TOP 1 1 FROM FIRSTTIME_INVOICEPRINTING  WHERE UNIT_CODE = '" & gstrUNITID & "' and doc_no= " & Trim(Ctlinvoice.Text) & "") = True Then
                                                    mP_Connection.Execute("Insert into FIRSTTIME_INVOICEPRINTING(unit_code,doc_no,ent_dt,ent_userid) values('" & gstrUNITID & "','" & Trim$(Ctlinvoice.Text) & "',getdate(),'" & mP_User & "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                                    If mblnISCrystalReportRequired Then  ''Praveen Digital Sign Changes 
                                                        If COPYNAME_New(1) = "Y" Then
                                                            RdAddSold.PrintToPrinter(1, False, 0, 0)
                                                        End If
                                                    End If
                                                Else
                                                    If mblnISCrystalReportRequired Then  ''Praveen Digital Sign Changes 
                                                        If COPYNAME_New(1) = "Y" Then
                                                            RdAddSold.PrintToPrinter(1, False, 0, 0)
                                                        End If
                                                    End If
                                                End If


                                                ''AMIT RANA 20 Jun 2019    
                                                If (blnIsPDFExported = False) Then
                                                    EXPORTINVOICETOPDF_ONPRINTREPRINT(strAccountCode, lbldescription.Text, lblcategory.Text, RdAddSold, DT_A4CUSTOMER_INVOICEPRINTINGTAG)
                                                    blnIsPDFExported = True
                                                End If
                                                ''AMIT RANA 20 Jun 2019 
                                            End If
                                        End If
                                    End If

                                Next
                            End If

                            '*******************************
                            If mblncustomerlevel_A4report_functionlity = True And mblnA4reports_invoicewise = True And mblncustomerlevel_Annexture_printing = True And mblncustomerspecificreport = True Then
                                Dim strTOTALQty As String
                                Dim Strquery As String
                                Dim strpdsno As String
                                Dim intLoopCounters As Short

                                RdAddSold.DataDefinition.FormulaFields("ANNEXTURE").Text = False
                                Strquery = "SELECT TOP 1 1 FROM SALES_DTL  WHERE UNIT_CODE = '" & gstrUNITID & "' and doc_no= " & Trim(Ctlinvoice.Text) & " HAVING COUNT(*)> " & mintmaxnoofitems_Annexture
                                RdAddSold = Frm.GetReportDocument()
                                RdAddSold.Load(My.Application.Info.DirectoryPath & "\Reports\Annextureprinting_A4reports.rpt")
                                RdAddSold.DataDefinition.FormulaFields("Invoiceno").Text = "'" & Ctlinvoice.Text & "'"
                                strpdsno = Find_Value("SELECT LORRYNO_DATE FROM SALESCHALLAN_DTL WHERE  UNIT_CODE = '" & gstrUNITID & "' AND  DOC_NO=" & Trim(Ctlinvoice.Text))
                                RdAddSold.DataDefinition.FormulaFields("pdsno").Text = "'" & strpdsno & "'"
                                strSQL = "{SalesChallan_Dtl.Location_Code}='" & Trim(txtUnitCode.Text) & "' and {SalesChallan_Dtl.Doc_No} =" & Trim(Ctlinvoice.Text) & " and {SalesChallan_Dtl.UNIT_CODE} ='" & gstrUNITID & "' "
                                RdAddSold.DataDefinition.RecordSelectionFormula = strSQL

                                If DataExist(Strquery) = True Then
                                    Frm.SetReportDocument()
                                    'For intLoopCounters = 0 To 1

                                    For intLoopCounters = 1 To intMaxLoop
                                        If chkprintreprint.Checked = True And optInvYes(1).Checked = True Then
                                            'COPYNAME = Find_Value("Select TEXTHEADING FROM  A4CUSTOMER_INVOICEPRINTINGTAG WHERE UNIT_CODE='" + gstrUNITID + "' AND CUSTOMER_CODE='" & strAccountCode & "' AND ORIGINAL_REPRINT='R' AND SERIALNO=" & intLoopCounter)
                                            Dim DataRowFiltered() As DataRow = DT_A4CUSTOMER_INVOICEPRINTINGTAG.Select("ORIGINAL_REPRINT='R' and SERIALNO=" + intLoopCounters.ToString())
                                            COPYNAME_New(0) = DataRowFiltered(0)("TEXTHEADING").ToString()
                                            COPYNAME_New(1) = DataRowFiltered(0)("HARDCOPYPRINTREQUIRED").ToString()
                                        ElseIf chkprintreprint.Checked = False And optInvYes(1).Checked = True Then
                                            'COPYNAME = Find_Value("Select TEXTHEADING FROM  A4CUSTOMER_INVOICEPRINTINGTAG WHERE UNIT_CODE='" + gstrUNITID + "' AND CUSTOMER_CODE='" & strAccountCode & "' AND ORIGINAL_REPRINT='O' AND SERIALNO=" & intLoopCounter)
                                            Dim DataRowFiltered() As DataRow = DT_A4CUSTOMER_INVOICEPRINTINGTAG.Select("ORIGINAL_REPRINT='O' and SERIALNO=" + intLoopCounters.ToString())
                                            COPYNAME_New(0) = DataRowFiltered(0)("TEXTHEADING").ToString()
                                            COPYNAME_New(1) = DataRowFiltered(0)("HARDCOPYPRINTREQUIRED").ToString()
                                        Else
                                            'COPYNAME = Find_Value("Select TEXTHEADING FROM  A4CUSTOMER_INVOICEPRINTINGTAG WHERE UNIT_CODE='" + gstrUNITID + "' AND CUSTOMER_CODE='" & strAccountCode & "' AND ORIGINAL_REPRINT='R' AND SERIALNO=" & intLoopCounter)
                                            Dim DataRowFiltered() As DataRow = DT_A4CUSTOMER_INVOICEPRINTINGTAG.Select("ORIGINAL_REPRINT='R' and SERIALNO=" + intLoopCounters.ToString())
                                            COPYNAME_New(0) = DataRowFiltered(0)("TEXTHEADING").ToString()
                                            COPYNAME_New(1) = DataRowFiltered(0)("HARDCOPYPRINTREQUIRED").ToString()
                                        End If
                                        If mblnEwaybill_Print = False Then
                                            If mblnISCrystalReportRequired Then    ''Praveen Digital Sign Changes 
                                                If COPYNAME_New(1) = "Y" Then
                                                    RdAddSold.PrintToPrinter(1, False, 0, 0)
                                                End If
                                            End If

                                            ''AMIT RANA 20 Jun 2019    
                                            'If (blnIsPDFExported = False) Then
                                            '    EXPORTINVOICETOPDF_ONPRINTREPRINT(strAccountCode, lbldescription.Text, lblcategory.Text, RdAddSold)
                                            '    blnIsPDFExported = True
                                            'End If
                                            ''AMIT RANA 20 Jun 2019 
                                        Else
                                            If chkprintreprint.Checked = True And optInvYes(1).Checked = True Then
                                                If mblnISCrystalReportRequired Then   ''Praveen Digital Sign Changes 
                                                    If COPYNAME_New(1) = "Y" Then
                                                        RdAddSold.PrintToPrinter(1, False, 0, 0)
                                                    End If
                                                End If

                                                ''AMIT RANA 20 Jun 2019    
                                                'If (blnIsPDFExported = False) Then
                                                '    EXPORTINVOICETOPDF_ONPRINTREPRINT(strAccountCode, lbldescription.Text, lblcategory.Text, RdAddSold)
                                                '    blnIsPDFExported = True
                                                'End If
                                                ''AMIT RANA 20 Jun 2019 
                                            Else
                                                If optInvYes(1).Checked = True Then
                                                    If Not DataExist("SELECT TOP 1 1 FROM FIRSTTIME_INVOICEPRINTING  WHERE UNIT_CODE = '" & gstrUNITID & "' and doc_no= " & Trim(Ctlinvoice.Text) & "") = True Then
                                                        mP_Connection.Execute("Insert into FIRSTTIME_INVOICEPRINTING(unit_code,doc_no,ent_dt,ent_userid) values('" & gstrUNITID & "','" & Trim$(Ctlinvoice.Text) & "',getdate(),'" & mP_User & "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                                        If mblnISCrystalReportRequired Then   ''Praveen Digital Sign Changes 
                                                            If COPYNAME_New(1) = "Y" Then
                                                                RdAddSold.PrintToPrinter(1, False, 0, 0)
                                                            End If
                                                        End If
                                                    Else
                                                        If mblnISCrystalReportRequired Then   ''Praveen Digital Sign Changes 
                                                            If COPYNAME_New(1) = "Y" Then
                                                                RdAddSold.PrintToPrinter(1, False, 0, 0)
                                                            End If
                                                        End If
                                                    End If


                                                    ''AMIT RANA 20 Jun 2019    
                                                    'If (blnIsPDFExported = False) Then
                                                    '    EXPORTINVOICETOPDF_ONPRINTREPRINT(strAccountCode, lbldescription.Text, lblcategory.Text, RdAddSold)
                                                    '    blnIsPDFExported = True
                                                    'End If
                                                    ''AMIT RANA 20 Jun 2019 
                                                End If
                                            End If
                                        End If
                                    Next
                                End If

                            End If

                            'If mblncustomerlevel_A4report_functionlity = False And mblnA4reports_invoicewise = True And mblncustomerlevel_Annexture_printing = True And gstrUNITID = "MST" Then
                            If mblncustomerspecificreport = False And mblncustomerlevel_A4report_functionlity = True And mblnA4reports_invoicewise = True And mblncustomerlevel_Annexture_printing = True And gstrUNITID = "MST" Then
                                Dim strTOTALQty As String
                                Dim Strquery As String
                                Dim strpdsno As String
                                Dim intLoopCounters As Short

                                RdAddSold.DataDefinition.FormulaFields("ANNEXTURE").Text = False
                                Strquery = "SELECT TOP 1 1 FROM SALES_DTL  WHERE UNIT_CODE = '" & gstrUNITID & "' and doc_no= " & Trim(Ctlinvoice.Text) & " HAVING COUNT(*)> " & mintmaxnoofitems_Annexture
                                RdAddSold = Frm.GetReportDocument()
                                RdAddSold.Load(My.Application.Info.DirectoryPath & "\Reports\Annextureprinting_A4reports.rpt")
                                RdAddSold.DataDefinition.FormulaFields("Invoiceno").Text = "'" & Ctlinvoice.Text & "'"
                                strpdsno = Find_Value("SELECT LORRYNO_DATE FROM SALESCHALLAN_DTL WHERE  UNIT_CODE = '" & gstrUNITID & "' AND  DOC_NO=" & Trim(Ctlinvoice.Text))
                                RdAddSold.DataDefinition.FormulaFields("pdsno").Text = "'" & strpdsno & "'"
                                strSQL = "{SalesChallan_Dtl.Location_Code}='" & Trim(txtUnitCode.Text) & "' and {SalesChallan_Dtl.Doc_No} =" & Trim(Ctlinvoice.Text) & " and {SalesChallan_Dtl.UNIT_CODE} ='" & gstrUNITID & "' "
                                RdAddSold.DataDefinition.RecordSelectionFormula = strSQL

                                If DataExist(Strquery) = True Then
                                    Frm.SetReportDocument()
                                    'For intLoopCounters = 0 To 1
                                    For intLoopCounters = 1 To intMaxLoop
                                        If mblnEwaybill_Print = False Then
                                            If mblnISCrystalReportRequired Then  ''Praveen Digital Sign Changes 
                                                If COPYNAME_New(1) = "Y" Then
                                                    RdAddSold.PrintToPrinter(1, False, 0, 0)
                                                End If
                                            End If

                                            ''AMIT RANA 20 Jun 2019    
                                            'If (blnIsPDFExported = False) Then
                                            '    EXPORTINVOICETOPDF_ONPRINTREPRINT(strAccountCode, lbldescription.Text, lblcategory.Text, RdAddSold)
                                            '    blnIsPDFExported = True
                                            'End If
                                            ''AMIT RANA 20 Jun 2019 
                                        Else
                                            If chkprintreprint.Checked = True And optInvYes(1).Checked = True Then
                                                If mblnISCrystalReportRequired Then   ''Praveen Digital Sign Changes 
                                                    If COPYNAME_New(1) = "Y" Then
                                                        RdAddSold.PrintToPrinter(1, False, 0, 0)
                                                    End If
                                                End If

                                                ''AMIT RANA 20 Jun 2019    
                                                'If (blnIsPDFExported = False) Then
                                                '    EXPORTINVOICETOPDF_ONPRINTREPRINT(strAccountCode, lbldescription.Text, lblcategory.Text, RdAddSold)
                                                '    blnIsPDFExported = True
                                                'End If
                                                ''AMIT RANA 20 Jun 2019 
                                            Else
                                                If optInvYes(1).Checked = True Then
                                                    If Not DataExist("SELECT TOP 1 1 FROM FIRSTTIME_INVOICEPRINTING  WHERE UNIT_CODE = '" & gstrUNITID & "' and doc_no= " & Trim(Ctlinvoice.Text) & "") = True Then
                                                        mP_Connection.Execute("Insert into FIRSTTIME_INVOICEPRINTING(unit_code,doc_no,ent_dt,ent_userid) values('" & gstrUNITID & "','" & Trim$(Ctlinvoice.Text) & "',getdate(),'" & mP_User & "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                                        If mblnISCrystalReportRequired Then  ''Praveen Digital Sign Changes 
                                                            If COPYNAME_New(1) = "Y" Then
                                                                RdAddSold.PrintToPrinter(1, False, 0, 0)
                                                            End If
                                                        End If
                                                    Else
                                                        If mblnISCrystalReportRequired Then  ''Praveen Digital Sign Changes 
                                                            If COPYNAME_New(1) = "Y" Then
                                                                RdAddSold.PrintToPrinter(1, False, 0, 0)
                                                            End If
                                                        End If
                                                    End If


                                                    ''AMIT RANA 20 Jun 2019    
                                                    'If (blnIsPDFExported = False) Then
                                                    '    EXPORTINVOICETOPDF_ONPRINTREPRINT(strAccountCode, lbldescription.Text, lblcategory.Text, RdAddSold)
                                                    '    blnIsPDFExported = True
                                                    'End If
                                                    ''AMIT RANA 20 Jun 2019 
                                                End If
                                            End If
                                        End If
                                    Next
                                End If

                            End If

                            'TOYOTA BARCODE 

                            'If mblncustomerlevel_A4report_functionlity = True And mblnA4reports_invoicewise = True And mblncustomerlevel_Annexture_printing = True Then

                            '    Dim strTOTALQty As String
                            '    Dim Strquery As String
                            '    Dim strpdsno As String
                            '    Dim intLoopCounters As Short
                            '    Strquery = "SELECT TOP 1 1 FROM SALES_DTL  WHERE UNIT_CODE = '" & gstrUNITID & "' and doc_no= " & Trim(mInvNo) & " HAVING COUNT(*)> " & mintmaxnoofitems_barcodeToyota
                            '    If DataExist(Strquery) = True Then
                            '        RdAddSold = Frm.GetReportDocument()
                            '        RdAddSold.Load(My.Application.Info.DirectoryPath & "\Reports\BarcodeToyotaprinting.rpt")
                            '        RdAddSold.DataDefinition.FormulaFields("Invoiceno").Text = "'" & mInvNo & "'"
                            '        strpdsno = Find_Value("SELECT LORRYNO_DATE FROM SALESCHALLAN_DTL WHERE  UNIT_CODE = '" & gstrUNITID & "' AND  DOC_NO=" & Trim(mInvNo))
                            '        RdAddSold.DataDefinition.FormulaFields("pdsno").Text = "'" & strpdsno & "'"
                            '        strSQL = "{SalesChallan_Dtl.Location_Code}='" & Trim(txtUnitCode.Text) & "' and {SalesChallan_Dtl.Doc_No} =" & Trim(mInvNo) & " and {SalesChallan_Dtl.UNIT_CODE} ='" & gstrUNITID & "' "
                            '        RdAddSold.DataDefinition.RecordSelectionFormula = strSQL

                            '        Frm.SetReportDocument()
                            '        For intLoopCounters = 0 To 1
                            '            RdAddSold.PrintToPrinter(1, False, 0, 0)
                            '        Next
                            '    End If

                            'End If
                            ''TOYOTA BARCODE 


                            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
                        End If


                    ElseIf AllowBarCodePrinting(strAccountCode) = True Then
                        If chkLockPrintingFlag.CheckState = CheckState.Unchecked And optInvYes(0).Checked = True Then
                            If DeleteBarCodeImage(Me.Ctlinvoice.Text) = False Then
                                Msgbox("Problem While deleting Barcode Image.", vbInformation, ResolveResString(100))
                                Exit Sub
                            End If
                        End If
                    End If
                    ''SATISH KESHAERWANI CHANGE
                    'If chktkmlbarcode.Checked = True Then
                    '    Print_barcodelabel(mInvNo)
                    'End If
                    ''SATISH KESHAERWANI CHANGE
                    If cmbInvType.Text.Trim.ToUpper = "REJECTION" Then
                        If CBool(Find_Value("select REJINVOICE_GRIN_PVKNOCKING from sales_parameter WHERE UNIT_CODE='" + gstrUNITID + "'")) = True Then
                            If DataExist("SELECT TOP 1 1 FROM MKT_INVREJ_DTL WHERE UNIT_CODE = '" & gstrUNITID & "' AND REJ_TYPE=1 AND CANCEL_FLAG=0 AND INVOICE_NO=" & mInvNo) Then
                                Call PrintDebitnote(MSTRREJECTIONNOTE)
                            End If
                        End If

                    End If
                    Ctlinvoice.Text = ""
                Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT_TO_FILE
                    mCheckValARana = "PRINTTOFILE"
                    mblnLock_Clicked = False
                    If mblnlorryno = True Then
                        Dim strlorryquery As String
                        strlorryquery = "UPDATE SalesChallan_Dtl SET LORRYNO_DATE= '" & txtlorryno.Text.Trim & "'  WHERE Doc_No=" & Ctlinvoice.Text & " and UNIT_CODE = '" & gstrUNITID & "'  and Location_Code='" & Trim(txtUnitCode.Text) & "'"
                        mP_Connection.Execute(strlorryquery, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End If
                    If InvoiceGeneration(RdAddSold, Frm) = True Then
                        frmExport.ShowDialog()
                        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
                        If gblnCancelExport Then Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default) : Exit Sub
                        If AllowBarCodePrinting(strAccountCode) = True Then
                            strBarcodeMsg = ""
                            If optInvYes(0).Checked = True Then
                                strBarcodeMsg = ObjBarcodeHMI.GenerateBarCode(gstrUserMyDocPath, mInvNo, True, Trim(Ctlinvoice.Text), gstrCONNECTIONSTRING)
                                If VB.Left(strBarcodeMsg, 1) <> "Y" Then
                                    Msgbox("Problem While Generating Barcode Image.", vbInformation, ResolveResString(100))
                                    Exit Sub
                                End If
                                If SaveBarCodeImage(Me.Ctlinvoice.Text, gstrUserMyDocPath) = False Then
                                    Msgbox("Problem While saving Barcode Image.", vbInformation, ResolveResString(100))
                                    Exit Sub
                                End If
                                If chkLockPrintingFlag.CheckState = CheckState.Unchecked And optInvYes(0).Checked = True Then
                                    If DeleteBarCodeImage(Me.Ctlinvoice.Text) = False Then
                                        Msgbox("Problem While deleting Barcode Image.", vbInformation, ResolveResString(100))
                                        Exit Sub
                                    End If
                                End If
                            End If
                        End If
                        Frm.ExportToFile()
                        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                    End If
            End Select
            Me.ChkCustDetails.CheckState = System.Windows.Forms.CheckState.Unchecked
            Exit Sub

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
                Call Logging_Starting_End_Time("Invoice locking PK ERROR: Saleconf Update Inv No " + mSaleConfNo.ToString + " : Current Inv No :" + mInvNo.ToString, DateTime.Now, "PK Issue", mInvNo)

                Msgbox(Ex.Message + " :Attemped to correct Internal PK Issue. Please Try Again!!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical, My.Resources.resEmpower.STR100)
            End If

            If Err.Number = 20545 Then
                'Resume Next
            Else
                SqlConnectionclass.ExecuteNonQuery("INSERT INTO INV_ERR_PRIMARYKEY(ERR_NUMBER,ERR_DESCRIPTION,TMPINVOICENO,INVOICE_NO,UNIT_CODE,FUNCTIONNAME ) VALUES('" & Err.Number & "','" & Replace(Ex.Message, "'", "") & "','" & Ctlinvoice.Text & "','" & mInvNo & "','" & gstrUNITID & "','cmdinvoice')")
                Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Ex.Message, mP_Connection, "CMDINVOICE")
            End If
        Finally
            '-- changes started by prashant rajpal on 09th Mar 2023
            If optInvYes(0).Checked = True And Val(mInvNo) <> 0 Then
                If Not strRetval = "Y" Then
                    If Not DataExist("select doc_no from Saleschallan_Dtl where doc_no='" & mInvNo & "' and UNIT_CODE = '" & gstrUNITID & "'") Then
                        CustomRollbackTrans()
                        mP_Connection.BeginTrans()
                        mP_Connection.Execute("delete from ar_docmaster WHERE docM_unit='" + gstrUNITID + "' AND  docm_vono='" & mInvNo & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        mP_Connection.Execute("delete from ar_docdtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  docd_vono='" & mInvNo & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        mP_Connection.Execute("delete from fin_gltrans WHERE glt_UntCodeID='" + gstrUNITID + "' AND  glt_srcdocno='" & mInvNo & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        mP_Connection.CommitTrans()
                        Call Logging_Starting_End_Time("Deletion of Finance Existance Tables : " + gstrUNITID + ":" + mInvNo.ToString(), DateTime.Now.ToString(), "Saved", mInvNo.ToString())
                        Msgbox("Internal Issue Occured !! ,  Please try Again.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, My.Resources.resEmpower.STR100)
                    End If
                End If
            End If
            '-- changes done by prashant rajpal on 09th Mar 2023
        End Try

    End Sub


    Private Sub ctlFormHeader1_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlFormHeader1.Click
        Call ShowHelp("HLPMKTTRN0008.htm")
    End Sub
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

    Private Sub PrintDebitnote(ByVal debitnno As String)
        On Error GoTo ErrHandler
        Dim strSelectionFormula As String = String.Empty
        Dim strReportName As String = String.Empty
        Dim strRptTitle As String = String.Empty
        Dim strRptName As String = String.Empty
        Dim rdCDS As ReportDocument
        Dim strRepPath As String = String.Empty
        Dim RepViewer As New eMProCrystalReportViewer
        Dim reptitle As String = String.Empty

        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)

        rdCDS = RepViewer.GetReportDocument
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        strRptName = GetPlantName()

        strSelectionFormula = "{Gen_UnitMaster.Unt_CodeID} = '" & gstrUNITID.Trim & "' and {Gen_PartyMaster.Prty_PartyID} = '" & mstracountcode & "' and {ap_docMaster.apDocM_voNo} = '" & Trim(debitnno) & "' and {ap_docMaster.apDocM_drcr}='DR' and {ap_docDtl.apdocD_DrCr} in ['DR','CR'] and {ap_docMaster.apDocM_voType}='M'"
        strReportName = "AP_DrCr_RSA"

        strRptName = "\Reports\" & strReportName & "_" & strRptName & ".rpt"

        If Not CheckFile(strRptName) Then
            strRptName = "\Reports\" & strReportName & ".rpt"
        End If
        strRptTitle = "AP Debit Note"

        strRepPath = My.Application.Info.DirectoryPath & strRptName
        rdCDS.Load(strRepPath)

        With rdCDS
            .RecordSelectionFormula = strSelectionFormula
            .SetParameterValue("reptitle", strRptTitle)
            .SetParameterValue("uid", gstrUserIDSelected.Trim)
            .SetParameterValue("cname", Trim(gstrCOMPANY))
            .SetParameterValue("unit", gstrUNITID.Trim)

        End With
        RepViewer.Show()
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Exit Sub
ErrHandler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
    End Sub
    Public Function SaveBarCodeImage_singlelevelso_2DBARCODE_In_Checking(ByVal pstrInvNo As String, ByVal pstrPath As String, ByVal strbarcodestring As String) As Boolean

        Dim strQuery As String
        Dim pstrpathKIA As String
        Dim blnResizing As Boolean
        Dim img As Image

        Try
            SaveBarCodeImage_singlelevelso_2DBARCODE_In_Checking = True

            pstrpathKIA = pstrPath & "BarcodeImgKIA.JPEG"
            pstrPath = pstrPath & "BarcodeImg.JPEG"

            blnResizing = CBool(Find_Value("SELECT isnull(Resize_barcode,0)as Resize_barcode FROM customer_mst (nolock) WHERE UNIT_CODE='" + gstrUNITID + "' and customer_code='" & mstracountcode & "'"))
            If blnResizing = True Then
                Call Resizeimage(pstrPath, pstrpathKIA, 400, 400)
            End If


            If blnResizing = True Then
                img = Image.FromFile(pstrpathKIA)
                'stimage.LoadFromFile(pstrpathKIA)
            Else
                'stimage.LoadFromFile(pstrPath)
                img = Image.FromFile(pstrPath)
            End If


            Dim bytes As Byte() = CType((New ImageConverter()).ConvertTo(img, GetType(Byte())), Byte())
            Dim cmd As New SqlCommand
            Dim sql As String = "UPDATE SALESCHALLAN_DTL SET BARCODEIMAGE=@BARCODEIMAGE WHERE DOC_NO=@DOC_NO AND UNIT_CODE= @UNIT_CODE"
            cmd.CommandType = CommandType.Text
            cmd.CommandText = sql
            cmd.Parameters.AddWithValue("@BARCODEIMAGE", bytes)
            cmd.Parameters.AddWithValue("@DOC_NO", Trim(pstrInvNo))
            cmd.Parameters.AddWithValue("@UNIT_CODE", gstrUNITID)
            SqlConnectionclass.ExecuteNonQuery(cmd)
            sql = String.Empty
            Exit Function
        Catch Ex As Exception
            Msgbox(Ex.Message)
            SaveBarCodeImage_singlelevelso_2DBARCODE_In_Checking = False
        End Try


    End Function


    Private Function GenerateASNFileForGen_Motors(ByVal pstrFileLocation As String, ByVal pstrInvoiceNo As String) As Boolean
        'Created By     : Tek Chand
        'Created On     : 04 Mar 2025
        'Reason         : To Generate ASN File for selected Invoices
        '--------------------------------------------------------------------------------------
        Dim strsql As String
        Dim rsGetASNData As ClsResultSetDB
        Dim Obj_FSO As Scripting.FileSystemObject
        Dim strLocation As String
        Dim strFileName As String
        Dim strRecord As String
        Dim strsplit As String()
        Dim strsplitstr As String
        Dim strsplitstrFinal As String

        Dim Sqlcmd As New SqlCommand
        Dim SqlAdp As SqlDataAdapter
        Dim DS As DataSet

        Dim intLineNo As Short
        On Error GoTo Err_Handler
        GenerateASNFileForGen_Motors = True
        '----------------------------------------
        Obj_FSO = New Scripting.FileSystemObject
        If Not Obj_FSO.FolderExists(pstrFileLocation) Then
            Obj_FSO.CreateFolder(pstrFileLocation)
        End If
        If Mid(Trim(pstrFileLocation), Len(Trim(pstrFileLocation))) <> "\" Then
            strLocation = pstrFileLocation & "\"
        End If

        strFileName = "ASN" & VB6.Format(GetServerDateTime(), "ddMMyyyyhhmmss") & ".txt"
        strFileName = strLocation & strFileName
        'Kill(strLocation & "*.csv")
        strsplitstr = pstrInvoiceNo
        strsplit = pstrInvoiceNo.Split(",")



        strsplitstrFinal = String.Join(", ", strsplit.Select(Function(w) w.Replace("'", "")).ToArray())
        FileClose(1)
        FileOpen(1, strFileName, OpenMode.Append)
        Obj_FSO = Nothing
        rsGetASNData = New ClsResultSetDB
        'strsql = "Exec USP_ASNDataForGen_Motors '" & gstrUNITID & "','" & strsplitstrFinal & "'"
        strsql = "Select '7C1' as SuplierEDICode,'BFT' as CustomerEDICode,SCL.Doc_No,DBO.UFN_GET_YYYYMMDDHHMM(INVOICE_DATE,INVOICE_TIME) AS ASNDate,
        DBO.UFN_GET_YYYYMMDDHHMM_1MonthAdd(INVOICE_DATE,INVOICE_TIME) AS ASNDateAddMonth,I.Weight*SD.Sales_Quantity as Gross_wt,I.Weight*SD.Sales_Quantity as Net_Wt,I.cons_measure_code,SCL.Doc_No,'72668' as STPC,'659974166' as SPC,
        'D37' as PD,'88120' as MRI,'SS' as MT,'A084' as CSC,'TE' as EQ,'TN' as CN,'4' as CPS,'' as blank,'' as blank,SD.Item_Code,'4' as MY,SD.Sales_Quantity,SD.Sales_Quantity as CQ,
        SCL.Cust_Ref as SONumber
        from SalesChallan_Dtl SCL With(nolock)
        left join Sales_Dtl SD With(nolock) on SCL.Doc_No=sd.Doc_No and scl.UNIT_CODE=sd.UNIT_CODE
        left join Item_Mst I With(nolock) on I.Item_Code=SD.Item_Code and I.UNIT_CODE=SD.UNIT_CODE
        WHERE SCL.BILL_FLAG=1 AND SCL.CANCEL_FLAG=0 AND SCL.DOC_NO in(" & pstrInvoiceNo & ") AND SCL.UNIT_CODE = '" & gstrUNITID & "'"
        rsGetASNData.GetResult(strsql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsGetASNData.GetNoRows > 0 Then
            rsGetASNData.MoveFirst()
            Do While Not rsGetASNData.EOFRecord
                strRecord = ""
                strRecord = IIf(IsDBNull(rsGetASNData.GetValue("SuplierEDICode")), "", rsGetASNData.GetValue("SuplierEDICode"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsGetASNData.GetValue("CustomerEDICode")), "", rsGetASNData.GetValue("CustomerEDICode"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsGetASNData.GetValue("Doc_No")), "", rsGetASNData.GetValue("Doc_No"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsGetASNData.GetValue("ASNDate")), "", rsGetASNData.GetValue("ASNDate"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsGetASNData.GetValue("ASNDate")), "", rsGetASNData.GetValue("ASNDate"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsGetASNData.GetValue("ASNDateAddMonth")), "", rsGetASNData.GetValue("ASNDateAddMonth"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsGetASNData.GetValue("Gross_wt")), "", rsGetASNData.GetValue("Gross_wt"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsGetASNData.GetValue("Net_Wt")), "", rsGetASNData.GetValue("Net_Wt"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsGetASNData.GetValue("cons_measure_code")), "", rsGetASNData.GetValue("cons_measure_code"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsGetASNData.GetValue("Doc_No")), "", rsGetASNData.GetValue("Doc_No"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsGetASNData.GetValue("STPC")), "", rsGetASNData.GetValue("STPC"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsGetASNData.GetValue("SPC")), "", rsGetASNData.GetValue("SPC"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsGetASNData.GetValue("PD")), "", rsGetASNData.GetValue("PD"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsGetASNData.GetValue("MRI")), "", rsGetASNData.GetValue("MRI"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsGetASNData.GetValue("MT")), "", rsGetASNData.GetValue("MT"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsGetASNData.GetValue("CSC")), "", rsGetASNData.GetValue("CSC"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsGetASNData.GetValue("EQ")), "", rsGetASNData.GetValue("EQ"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsGetASNData.GetValue("CN")), "", rsGetASNData.GetValue("CN"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsGetASNData.GetValue("CPS")), "", rsGetASNData.GetValue("CPS"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsGetASNData.GetValue("blank")), "", rsGetASNData.GetValue("blank"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsGetASNData.GetValue("blank")), "", rsGetASNData.GetValue("blank"))

                strRecord = strRecord & "," & IIf(IsDBNull(rsGetASNData.GetValue("Item_Code")), "", rsGetASNData.GetValue("Item_Code"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsGetASNData.GetValue("MY")), "", rsGetASNData.GetValue("MY"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsGetASNData.GetValue("Sales_Quantity")), "", rsGetASNData.GetValue("Sales_Quantity"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsGetASNData.GetValue("CQ")), "", rsGetASNData.GetValue("CQ"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsGetASNData.GetValue("cons_measure_code")), "", rsGetASNData.GetValue("cons_measure_code"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsGetASNData.GetValue("SONumber")), "", rsGetASNData.GetValue("SONumber"))

                PrintLine(1, strRecord) : intLineNo = intLineNo + 1
                rsGetASNData.MoveNext()
            Loop
            rsGetASNData.ResultSetClose()
            rsGetASNData = Nothing
        Else
            Msgbox("No Invoice Records found to generate the File.", MsgBoxStyle.Information, ResolveResString(100))
            FileClose(1)
            Kill(strFileName)
            rsGetASNData.ResultSetClose()
            rsGetASNData = Nothing
            GenerateASNFileForGen_Motors = False
            Exit Function
        End If
        FileClose(1)
        Exit Function
Err_Handler:
        If Err.Number = 55 Then
            Msgbox("File Already Open, Cann't Generate the ASN File.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
            GenerateASNFileForGen_Motors = False
            Exit Function
        End If
        GenerateASNFileForGen_Motors = False
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

    Public Function SaveBarCodeImage_singlelevelso_2DBARCODE(ByVal pstrInvNo As String, ByVal pstrPath As String, ByVal strbarcodestring As String) As Boolean
        On Error GoTo ErrHandler
        Dim stimage As ADODB.Stream
        Dim strQuery As String
        Dim Rs As ADODB.Recordset
        Dim pstrpathKIA As String
        Dim blnResizing As Boolean

        SaveBarCodeImage_singlelevelso_2DBARCODE = True
        stimage = New ADODB.Stream
        stimage.Type = ADODB.StreamTypeEnum.adTypeBinary
        pstrpathKIA = pstrPath & "BarcodeImgKIA.JPEG"
        pstrPath = pstrPath & "BarcodeImg.JPEG"

        blnResizing = CBool(Find_Value("SELECT isnull(Resize_barcode,0)as Resize_barcode FROM customer_mst WHERE UNIT_CODE='" + gstrUNITID + "' and customer_code='" & mstracountcode & "'"))
        If blnResizing = True Then
            Call Resizeimage(pstrPath, pstrpathKIA, 400, 400)
        End If

        stimage.Open()
        If blnResizing = True Then
            stimage.LoadFromFile(pstrpathKIA)
        Else
            stimage.LoadFromFile(pstrPath)
        End If

        strQuery = "select  barcodeimage  from saleschallan_dtl where doc_no='" & Trim(pstrInvNo) & "' and UNIT_CODE = '" & gstrUNITID & "'"
        Rs = Nothing
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
    Private Sub Resizeimage(ByVal strsourceimage As String, ByVal strtargetimage As String, ByVal IntHeight As Integer, ByVal IntWidth As Integer)

        Dim widthC As Integer = 200
        Dim heightC As Integer = 200
        Using imageC As Bitmap = New Bitmap(strsourceimage)
            Using targetC As Bitmap = New Bitmap(widthC, heightC, PixelFormat.Format24bppRgb)
                Using graphicC As Graphics = Graphics.FromImage(targetC)
                    graphicC.SmoothingMode = SmoothingMode.AntiAlias
                    graphicC.InterpolationMode = InterpolationMode.HighQualityBicubic
                    graphicC.PixelOffsetMode = PixelOffsetMode.HighQuality
                    graphicC.CompositingQuality = CompositingQuality.HighSpeed
                    graphicC.CompositingMode = CompositingMode.SourceCopy
                    graphicC.DrawImage(imageC, 0, 0, widthC, heightC)
                    targetC.Save(strtargetimage)
                End Using
            End Using
        End Using

    End Sub
    Public Function SaveBarCodeImage(ByVal pstrInvNo As String, ByVal pstrPath As String) As Boolean
        '----------------------------------------------------------------------------
        'Author         :   Manoj Vaish
        'Argument       :   Invoice No.
        'Return Value   :   True if successful,False if fails
        'Function       :   Save the barCode Image Into database
        'Issue ID       :   22486
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim stimage As ADODB.Stream
        Dim strQuery As String
        Dim Rs As ADODB.Recordset
        SaveBarCodeImage = True
        stimage = New ADODB.Stream
        stimage.Type = ADODB.StreamTypeEnum.adTypeBinary
        stimage.Open()
        pstrPath = pstrPath & "\BarcodeImg.bmp"
        stimage.LoadFromFile(pstrPath)
        strQuery = "select  barCodeImage from saleschallan_dtl where doc_no='" & Trim(pstrInvNo) & "' and UNIT_CODE = '" & gstrUNITID & "'"
        Rs = New ADODB.Recordset
        Rs.Open(strQuery, mP_Connection, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
        Rs.Fields("barCodeImage").Value = stimage.Read
        Rs.Update()
        Rs.Close()
        Rs = Nothing
        Exit Function
ErrHandler:
        SaveBarCodeImage = False
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

    Public Function DeleteBarCodeImage(ByVal pstrInvNo As String) As Boolean
        '----------------------------------------------------------------------------
        'Author         :   Manoj Vaish
        'Argument       :   Invoice No.
        'Return Value   :   True if successful,False if fails
        'Function       :   Set the barCode Image field Null Into database if invocie is not locked
        'Issue ID       :   22486
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim strQuery As String
        DeleteBarCodeImage = True
        strQuery = "update saleschallan_dtl set barCodeImage=NULL from saleschallan_dtl where doc_no='" & Trim(pstrInvNo) & "' and UNIT_CODE = '" & gstrUNITID & "'"
        mP_Connection.Execute(strQuery, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        Exit Function
ErrHandler:
        DeleteBarCodeImage = False
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function AllowBarCodePrinting(ByVal pstraccoutncode As String) As Boolean
        '----------------------------------------------------------------------------
        'Author         :   Manoj Kr. Vaish
        'Function       :   Check BarCodePrinting from sales_parameter
        'Comments       :   Date: 26 Feb 2008 ,Issue Id: 22486
        '----------------------------------------------------------------------------
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
    Private Sub txtASNNumber_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtASNNumber.KeyPress
        'Created  By     : Manoj Kr Vaish
        'Creation Date   : 09 Mar 2009
        'Issue ID        : eMpro-20090204-27027
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
    End Sub
    Private Function AllowASNPrinting(ByVal pstraccountcode As String) As Boolean
        'Revised By     : Manoj Kr. Vaish
        'Revised On     : 24 Mar  2009
        'Arguments      : Account Code
        'Return Value   : True/False
        'Issue ID       : eMpro-20090204-27027
        'Reason         : Check ASNPrinting from Customer Master
        '--------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim strQry As String
        Dim Rs As ClsResultSetDB_Invoice
        AllowASNPrinting = False
        strQry = "Select isnull(AllowASNPrinting,0) as AllowASNPrinting from customer_mst where Customer_Code='" & Trim(pstraccountcode) & "' and UNIT_CODE = '" & gstrUNITID & "'"
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
        'Issue ID       : eMpro-20090204-27027
        'Reason         : Check ASN already exist for invoice
        '--------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim rsgetASNNumber As ClsResultSetDB_Invoice
        Dim strsql As String
        rsgetASNNumber = New ClsResultSetDB_Invoice
        strsql = "select ASN_NO from CreatedASN where doc_no='" & Trim(pstrInvoiceNo) & "' and UNIT_CODE = '" & gstrUNITID & "'"
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

            Me.txtlorryno.Visible = False
            Me.txtlorryno.Enabled = False
            Me.txtlorryno.Text = ""
            Me.txtlorryno.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            Me.lblLorry.Visible = False
        End If
ErrHandler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Function AllowASNTextFileGeneration(ByVal pstraccountcode As String) As Boolean
        'Revised By     : Manoj Kr. Vaish
        'Revised On     : 14 May 2009
        'Arguments      : Account Code
        'Return Value   : True/False
        'Issue ID       : eMpro-20090513-31282
        'Reason         : Check ASNTextFileGeneration from Customer Master
        '--------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim strQry As String
        Dim Rs As ClsResultSetDB_Invoice
        AllowASNTextFileGeneration = False
        If (UCase(Trim(cmbInvType.Text)) = "NORMAL INVOICE" And UCase(Trim(CmbCategory.Text)) = "FINISHED GOODS") Then
            strQry = "Select isnull(AllowASNTextGeneration,0) as AllowASNTextGeneration from customer_mst where Customer_Code='" & Trim(pstraccountcode) & "' and UNIT_CODE = '" & gstrUNITID & "'"
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
        '10126648 
        If (UCase(Trim(cmbInvType.Text)) = "REJECTION" And UCase(Trim(CmbCategory.Text)) = "REJECTION") Then
            strQry = "Select isnull(AllowASNTextGeneration,0) as AllowASNTextGeneration from Vendor_mst where Vendor_code ='" & Trim(pstraccountcode) & "' and UNIT_CODE = '" & gstrUNITID & "'"
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
        '10126648: End
        Exit Function
ErrHandler:
        Rs = Nothing
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function FordASNFileGeneration(ByVal pintdocno As Double, ByVal pstraccountcode As String) As Boolean
        'Revised By     : Manoj Kr. Vaish
        'Revised On     : 14 May 2009
        'Arguments      : INvoice No
        'Issue ID       : eMpro-20090513-31282
        'Reason         : Generate ASN File for FORD
        '--------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim rsgetData As New ClsResultSetDB_Invoice
        Dim strquery As String
        Dim fs As FileStream
        Dim sw As StreamWriter
        Dim strASNdata As String
        Dim Dcount As Integer
        Dim TotalQty As Double
        Dim dblSalesQty As Double
        Dim strcontainerdespQty As String
        Dim dblcummulativeQty As Double
        Dim dblContainerQty As Double
        Dim strTotalQty As String
        Dim strASNFilepath As String
        Dim strASNFilepathforEDI As String
        Dim strSQL As String
        Dim strLorryNo As String
        Dim strtotalquantity As String
        Dim strnoofItems As String

        strASNdata = ""
        strquery = "select * from dbo.FN_GETASNDETAIL(" & pintdocno & ",'" & gstrUNITID & "')"
        rsgetData.GetResult(strquery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsgetData.GetNoRows > 0 Then
            If rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Length = 0 Then
                MessageBox.Show("Customer vendor code is not defined for Customer : " & pstraccountcode, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                FordASNFileGeneration = False
                Exit Function
            Else
                If rsgetData.GetValue("CUST_PLANTCODE").ToString.Trim().Length = 0 Then
                    MessageBox.Show("Unable To Get Plant Code For The Customer: " & pstraccountcode & " While Generating ASN File." & vbCrLf &
                                    "Invoice Can't Be Locked", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                    FordASNFileGeneration = False
                    Exit Function
                End If
                '10856126
                strSQL = "select dbo.UDF_ISDOCKCODE_INVOICE( '" & gstrUNITID & "','" & pstraccountcode & "','NORMAL INVOICE','FINISHED GOODS' )"
                If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSQL)) = True Then
                    If rsgetData.GetValue("CARRIER_CODE").ToString.Trim().Length = 0 Then
                        MessageBox.Show("CARRIER CODE is not defined for Customer : " & pstraccountcode, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                        FordASNFileGeneration = False
                        Exit Function
                    End If
                    strASNdata = "856HD20200000000000000000" & Space(5 - rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim.Length) + rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim & rsgetData.GetValue("CUST_PLANTCODE").ToString.Trim & Space(5 - rsgetData.GetValue("CUST_PLANTCODE").ToString.Trim.Length) & ",     ,     *" & vbCrLf
                    strSQL = "select single_lorryno_reqd from sales_parameter where unit_code = '" & gstrUNITID & "' "
                    If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSQL)) = False Then
                        strASNdata = strASNdata & "856A M" & mInvNo.ToString.Trim() & Space(10 - mInvNo.ToString.Trim().Length) & Space(5 - rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim.Length) & rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim & "09" & VB6.Format(rsgetData.GetValue("INVOICE_DATE"), "yymmdd") & VB6.Format(rsgetData.GetValue("INVOICE_DATE"), "hhmm") & VB6.Format(rsgetData.GetValue("INVOICE_DATE"), "yymmdd") & VB6.Format(rsgetData.GetValue("INVOICE_DATE"), "hhmm") & Space(10) & "+00000000100KG+00000000080KG" & "AE0N" & Space(8) & rsgetData.GetValue("TRANSPORT_TYPE").ToString() & Space("12") & "00" & VB.Right(rsgetData.GetValue("LORRYNO_DATE").ToString, 4) & Space(4) & Space(35) & "M" & VB.Right(mInvNo, 5) & Space(5) & rsgetData.GetValue("CUST_PLANTCODE").ToString.Trim() & Space(5 - rsgetData.GetValue("CUST_PLANTCODE").ToString.Trim.Length) & rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim & Space(6) & Space(5) & Space(4 - rsgetData.GetValue("CARRIER_CODE").ToString.Trim.Length) & rsgetData.GetValue("CARRIER_CODE").ToString.Trim + Space(11 - rsgetData.GetValue("LORRYNO_DATE").ToString.Trim.Length) + rsgetData.GetValue("LORRYNO_DATE").ToString.Trim & Space(7 - rsgetData.GetValue("DOCKCODE").ToString.Trim.Length) + rsgetData.GetValue("DOCKCODE").ToString.Trim & vbCrLf
                    Else
                        strLorryNo = rsgetData.GetValue("LORRYNO_DATE").ToString.Trim
                        If Len(strLorryNo) >= 5 Then
                            strLorryNo = VB.Right(strLorryNo, 5)
                        Else
                            While Len(strLorryNo) < 5
                                strLorryNo = "0" + strLorryNo
                            End While
                        End If
                        strASNdata = strASNdata & "856A M" & mInvNo.ToString.Trim() & Space(10 - mInvNo.ToString.Trim().Length) & Space(5 - rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim.Length) & rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim & "09" & VB6.Format(rsgetData.GetValue("INVOICE_DATE"), "yymmdd") & VB6.Format(rsgetData.GetValue("INVOICE_DATE"), "hhmm") & VB6.Format(rsgetData.GetValue("INVOICE_DATE"), "yymmdd") & VB6.Format(rsgetData.GetValue("INVOICE_DATE"), "hhmm") & Space(10) & "+00000000100KG+00000000080KG" & "AE0N" & Space(8) & rsgetData.GetValue("TRANSPORT_TYPE").ToString() & Space("12") & "0" & strLorryNo & Space(4) & Space(35) & "M" & VB.Right(mInvNo, 5) & Space(5) & rsgetData.GetValue("CUST_PLANTCODE").ToString.Trim() & Space(5 - rsgetData.GetValue("CUST_PLANTCODE").ToString.Trim.Length) & rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim & Space(6) & Space(5) & Space(4 - rsgetData.GetValue("CARRIER_CODE").ToString.Trim.Length) & rsgetData.GetValue("CARRIER_CODE").ToString.Trim + Space(11 - rsgetData.GetValue("LORRYNO_DATE").ToString.Trim.Length) + rsgetData.GetValue("LORRYNO_DATE").ToString.Trim & Space(7 - rsgetData.GetValue("DOCKCODE").ToString.Trim.Length) + rsgetData.GetValue("DOCKCODE").ToString.Trim & vbCrLf
                        'strSQL = "select dbo.IsLorryNoExists_For_ASN( '" & gstrUNITID & "','" & rsgetData.GetValue("LORRYNO_DATE").ToString.Trim & "'," & pintdocno & " )"
                        'If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSQL)) = True Then
                        '    strASNdata = strASNdata & "856A M" & mInvNo.ToString.Trim() & Space(10 - mInvNo.ToString.Trim().Length) & Space(5 - rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim.Length) & rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim & "09" & VB6.Format(rsgetData.GetValue("INVOICE_DATE"), "yymmdd") & VB6.Format(rsgetData.GetValue("INVOICE_DATE"), "hhmm") & VB6.Format(rsgetData.GetValue("INVOICE_DATE"), "yymmdd") & VB6.Format(rsgetData.GetValue("INVOICE_DATE"), "hhmm") & Space(10) & "+00000000100KG+00000000080KG" & "AE0N" & Space(8) & rsgetData.GetValue("TRANSPORT_TYPE").ToString() & Space("12") & "0" & VB.Right(rsgetData.GetValue("LORRYNO_DATE").ToString, 4) & "A" & Space(4) & Space(35) & "M" & VB.Right(mInvNo, 5) & Space(5) & rsgetData.GetValue("CUST_PLANTCODE").ToString.Trim() & Space(5 - rsgetData.GetValue("CUST_PLANTCODE").ToString.Trim.Length) & rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim & Space(6) & Space(5) & Space(4 - rsgetData.GetValue("CARRIER_CODE").ToString.Trim.Length) & rsgetData.GetValue("CARRIER_CODE").ToString.Trim + Space(11 - rsgetData.GetValue("LORRYNO_DATE").ToString.Trim.Length) + rsgetData.GetValue("LORRYNO_DATE").ToString.Trim & Space(7 - rsgetData.GetValue("DOCKCODE").ToString.Trim.Length) + rsgetData.GetValue("DOCKCODE").ToString.Trim & vbCrLf
                        'Else
                        '    strASNdata = strASNdata & "856A M" & mInvNo.ToString.Trim() & Space(10 - mInvNo.ToString.Trim().Length) & Space(5 - rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim.Length) & rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim & "09" & VB6.Format(rsgetData.GetValue("INVOICE_DATE"), "yymmdd") & VB6.Format(rsgetData.GetValue("INVOICE_DATE"), "hhmm") & VB6.Format(rsgetData.GetValue("INVOICE_DATE"), "yymmdd") & VB6.Format(rsgetData.GetValue("INVOICE_DATE"), "hhmm") & Space(10) & "+00000000100KG+00000000080KG" & "AE0N" & Space(8) & rsgetData.GetValue("TRANSPORT_TYPE").ToString() & Space("12") & "0" & VB.Right(rsgetData.GetValue("LORRYNO_DATE").ToString, 4) & "B" & Space(4) & Space(35) & "M" & VB.Right(mInvNo, 5) & Space(5) & rsgetData.GetValue("CUST_PLANTCODE").ToString.Trim() & Space(5 - rsgetData.GetValue("CUST_PLANTCODE").ToString.Trim.Length) & rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim & Space(6) & Space(5) & Space(4 - rsgetData.GetValue("CARRIER_CODE").ToString.Trim.Length) & rsgetData.GetValue("CARRIER_CODE").ToString.Trim + Space(11 - rsgetData.GetValue("LORRYNO_DATE").ToString.Trim.Length) + rsgetData.GetValue("LORRYNO_DATE").ToString.Trim & Space(7 - rsgetData.GetValue("DOCKCODE").ToString.Trim.Length) + rsgetData.GetValue("DOCKCODE").ToString.Trim & vbCrLf
                        'End If


                    End If
                Else
                    '20 nov 2017
                    'strASNdata = "856HD20200000000000000000" & Space(5 - rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim.Length) + rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim & rsgetData.GetValue("CUST_PLANTCODE").ToString.Trim & Space(5 - rsgetData.GetValue("CUST_PLANTCODE").ToString.Trim.Length) & ",     ,     *" & vbCrLf
                    'strASNdata = strASNdata & "856A M" & mInvNo.ToString.Trim() & Space(10 - mInvNo.ToString.Trim().Length) & Space(5 - rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim.Length) & rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim & "09" & VB6.Format(rsgetData.GetValue("INVOICE_DATE"), "yymmdd") & VB6.Format(rsgetData.GetValue("INVOICE_DATE"), "hhmm") & VB6.Format(rsgetData.GetValue("INVOICE_DATE"), "yymmdd") & VB6.Format(rsgetData.GetValue("INVOICE_DATE"), "hhmm") & Space(10) & "+00000000100KG+00000000080KG" & "AE0N" & Space(8) & rsgetData.GetValue("TRANSPORT_TYPE").ToString() & Space("12") & "M" & VB.Right(mInvNo, 5) & Space(4) & Space(35) & "M" & VB.Right(mInvNo, 5) & Space(5) & rsgetData.GetValue("CUST_PLANTCODE").ToString.Trim() & Space(5 - rsgetData.GetValue("CUST_PLANTCODE").ToString.Trim.Length) & rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim & Space(6) & rsgetData.GetValue("ARL_CODE").ToString.Trim() & Space(5 - rsgetData.GetValue("ARL_CODE").ToString.Trim().Length) & VB6.Format(GetServerDateTime, "mmddhhmm") & Space(3) & "0000000.00" & vbCrLf
                    strSQL = "select single_lorryno_reqd from sales_parameter where unit_code = '" & gstrUNITID & "' "
                    If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSQL)) = False Then
                        strASNdata = "856HD20200000000000000000" & Space(5 - rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim.Length) + rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim & rsgetData.GetValue("CUST_PLANTCODE").ToString.Trim & Space(5 - rsgetData.GetValue("CUST_PLANTCODE").ToString.Trim.Length) & ",     ,     *" & vbCrLf
                        If gstrUNITID = "MSD" And rsgetData.GetValue("INTERMEDIATE_CONSIGNEE_CODE").ToString.Trim().Length > 0 Then
                            strASNdata = strASNdata & "856A M" & mInvNo.ToString.Trim() & Space(10 - mInvNo.ToString.Trim().Length) & Space(5 - rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim.Length) & rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim & "09" & VB6.Format(rsgetData.GetValue("INVOICE_DATE"), "yymmdd") & VB6.Format(rsgetData.GetValue("INVOICE_DATE"), "hhmm") & VB6.Format(rsgetData.GetValue("INVOICE_DATE"), "yymmdd") & VB6.Format(rsgetData.GetValue("INVOICE_DATE"), "hhmm") & Space(10) & "+00000000100KG+00000000080KG" & "AE0N" & Space(8) & rsgetData.GetValue("TRANSPORT_TYPE").ToString() & Space("12") & "M" & VB.Right(mInvNo, 5) & Space(4) & Space(35) & "M" & VB.Right(mInvNo, 5) & Space(5) & rsgetData.GetValue("CUST_PLANTCODE").ToString.Trim() & Space(5 - rsgetData.GetValue("CUST_PLANTCODE").ToString.Trim.Length) & rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim & Space(6) & rsgetData.GetValue("INTERMEDIATE_CONSIGNEE_CODE").ToString.Trim() & Space(5 - rsgetData.GetValue("INTERMEDIATE_CONSIGNEE_CODE").ToString.Trim().Length) & VB6.Format(GetServerDateTime, "mmddhhmm") & Space(3) & "0000000.00" & vbCrLf
                        Else
                            strASNdata = strASNdata & "856A M" & mInvNo.ToString.Trim() & Space(10 - mInvNo.ToString.Trim().Length) & Space(5 - rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim.Length) & rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim & "09" & VB6.Format(rsgetData.GetValue("INVOICE_DATE"), "yymmdd") & VB6.Format(rsgetData.GetValue("INVOICE_DATE"), "hhmm") & VB6.Format(rsgetData.GetValue("INVOICE_DATE"), "yymmdd") & VB6.Format(rsgetData.GetValue("INVOICE_DATE"), "hhmm") & Space(10) & "+00000000100KG+00000000080KG" & "AE0N" & Space(8) & rsgetData.GetValue("TRANSPORT_TYPE").ToString() & Space("12") & "M" & VB.Right(mInvNo, 5) & Space(4) & Space(35) & "M" & VB.Right(mInvNo, 5) & Space(5) & rsgetData.GetValue("CUST_PLANTCODE").ToString.Trim() & Space(5 - rsgetData.GetValue("CUST_PLANTCODE").ToString.Trim.Length) & rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim & Space(6) & rsgetData.GetValue("ARL_CODE").ToString.Trim() & Space(5 - rsgetData.GetValue("ARL_CODE").ToString.Trim().Length) & VB6.Format(GetServerDateTime, "mmddhhmm") & Space(3) & "0000000.00" & vbCrLf
                        End If

                    Else
                        strLorryNo = rsgetData.GetValue("LORRYNO_DATE").ToString.Trim
                        If Len(strLorryNo) >= 5 Then
                            strLorryNo = VB.Right(strLorryNo, 5)
                        Else
                            While Len(strLorryNo) < 5
                                strLorryNo = "0" + strLorryNo
                            End While
                        End If
                        strASNdata = "856HD20200000000000000000" & Space(5 - rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim.Length) + rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim & rsgetData.GetValue("CUST_PLANTCODE").ToString.Trim & Space(5 - rsgetData.GetValue("CUST_PLANTCODE").ToString.Trim.Length) & ",     ,     *" & vbCrLf
                        strASNdata = strASNdata & "856A M" & mInvNo.ToString.Trim() & Space(10 - mInvNo.ToString.Trim().Length) & Space(5 - rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim.Length) & rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim & "09" & VB6.Format(rsgetData.GetValue("INVOICE_DATE"), "yymmdd") & VB6.Format(rsgetData.GetValue("INVOICE_DATE"), "hhmm") & VB6.Format(rsgetData.GetValue("INVOICE_DATE"), "yymmdd") & VB6.Format(rsgetData.GetValue("INVOICE_DATE"), "hhmm") & Space(10) & "+00000000100KG+00000000080KG" & "AE0N" & Space(8) & rsgetData.GetValue("TRANSPORT_TYPE").ToString() & Space("12") & "0" & strLorryNo & Space(4) & Space(35) & "M" & VB.Right(mInvNo, 5) & Space(5) & rsgetData.GetValue("CUST_PLANTCODE").ToString.Trim() & Space(5 - rsgetData.GetValue("CUST_PLANTCODE").ToString.Trim.Length) & rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim & Space(6) & rsgetData.GetValue("ARL_CODE").ToString.Trim() & Space(5 - rsgetData.GetValue("ARL_CODE").ToString.Trim().Length) & VB6.Format(GetServerDateTime, "mmddhhmm") & Space(3) & "0000000.00" & vbCrLf
                    End If
                    '20 nov 2017
                End If
                '10856126
                Dcount = 2
                strcontainerdespQty = Find_Value("select sum(isnull(to_box,0)-isnull(from_box,0)+1) as Desp_Qty from sales_dtl where  UNIT_CODE = '" & gstrUNITID & "' and  doc_no=" & pintdocno)
                strASNdata = strASNdata & "856TD"
                Select Case rsgetData.GetValue("CONTAINER").ToString.Trim.Length()
                    Case 3
                        strASNdata = strASNdata & rsgetData.GetValue("CONTAINER").ToString.Trim() & "90+" & Mid("000000", strcontainerdespQty.Length(), 6) & strcontainerdespQty.ToString()
                    Case 4
                        strASNdata = strASNdata & rsgetData.GetValue("CONTAINER").ToString.Trim() & " +" & Mid("000000", strcontainerdespQty.Length(), 6) & strcontainerdespQty.ToString()
                    Case 5
                        strASNdata = strASNdata & rsgetData.GetValue("CONTAINER").ToString.Trim() & "+" & Mid("000000", strcontainerdespQty.Length(), 6) & strcontainerdespQty.ToString()
                    Case 1, 2
                        strASNdata = strASNdata & rsgetData.GetValue("CONTAINER").ToString.Trim() & Space(3 - rsgetData.GetValue("CONTAINER").ToString.Trim.Length()) & "  +" & Mid("000000", strcontainerdespQty.Length(), 6) & strcontainerdespQty.ToString()
                    Case Else
                        strASNdata = strASNdata & VB.Left(rsgetData.GetValue("CONTAINER").ToString.Trim(), 5) & "+" & Mid("000000", strcontainerdespQty.Length(), 6) & strcontainerdespQty.ToString()
                End Select
                strtotalquantity = CInt(Find_Value("select sum(isnull(sales_quantity,0)) from sales_dtl where  UNIT_CODE = '" & gstrUNITID & "' and  doc_no=" & pintdocno))
                'strtotalquantity = Val(strtotalquantity)

                While Len(strtotalquantity) < 8
                    strtotalquantity = "0" + strtotalquantity
                End While
                strASNdata = strASNdata & strtotalquantity
                strnoofItems = CInt(Find_Value("select COUNT(*) NOOFITEMS  from sales_dtl where  UNIT_CODE = '" & gstrUNITID & "' and  doc_no=" & pintdocno))
                While Len(strnoofItems) < 3
                    strnoofItems = "0" + strnoofItems
                End While
                strASNdata = strASNdata & strnoofItems & vbCrLf

                Dcount = Dcount + 1
                rsgetData.MoveFirst()
                Do While Not rsgetData.EOFRecord
                    dblcummulativeQty = 0
                    dblSalesQty = 0
                    dblContainerQty = 0
                    dblcummulativeQty = Find_Value("SELECT DBO.UDF_GET_CUMMULATIVEQTY('" & gstrUNITID & "','" & rsgetData.GetValue("CUST_PLANTCODE").ToString() & "','" & rsgetData.GetValue("CUST_PART_CODE").ToString() & "'," & pintdocno & ")")
                    dblSalesQty = rsgetData.GetValue("SALES_QUANTITY")
                    dblcummulativeQty = dblcummulativeQty + dblSalesQty
                    dblContainerQty = rsgetData.GetValue("CONTAINER_QTY")
                    strASNdata = strASNdata & "856P "
                    'strASNdata = strASNdata & rsgetData.GetValue("CUST_PART_CODE").ToString().Trim & Space(30 - rsgetData.GetValue("CUST_PART_CODE").ToString().Length())
                    strASNdata = strASNdata & rsgetData.GetValue("CUST_PART_CODE").ToString().Trim & Space(30 - rsgetData.GetValue("CUST_PART_CODE").ToString().Trim.Length())
                    dblSalesQty = rsgetData.GetValue("Sales_Quantity")
                    strASNdata = strASNdata & "BP+" & Mid("0000000", dblSalesQty.ToString.Length(), 8) & dblSalesQty & "EA+"
                    strTotalQty = Val(strTotalQty) + Val(rsgetData.GetValue("SALES_QUANTITY"))
                    Dcount = Dcount + 1
                    strASNdata = strASNdata & Mid("000000000", dblcummulativeQty.ToString().Length(), 10) & dblcummulativeQty
                    strASNdata = strASNdata & "+0000000000" & Space(10) & mInvNo.ToString & Space(11 - mInvNo.ToString.Length()) & rsgetData.GetValue("CUST_VENDOR_CODE").ToString & VB6.Format(rsgetData.GetValue("INVOICE_DATE").ToString(), "yymmdd") & VB6.Format(rsgetData.GetValue("INVOICE_DATE").ToString(), "hhmm") & vbCrLf
                    strASNdata = strASNdata & "856PA" & Space(30) & "+00000000000  +00000000000  " & vbCrLf
                    strASNdata = strASNdata & "856V " & "+000000000000000" & vbCrLf
                    Dcount = Dcount + 2
                    strASNdata = strASNdata & "856C +" & Mid("0000000", dblContainerQty.ToString.Length(), 8) & dblContainerQty & "+" & Mid("0000", rsgetData.GetValue("CONTAINER_DESP_QTY").ToString.Length, 5) & rsgetData.GetValue("CONTAINER_DESP_QTY").ToString & rsgetData.GetValue("CONTAINER").ToString & "90" & vbCrLf
                    Dcount = Dcount + 1
                    mstrupdateASNdtl = Trim(mstrupdateASNdtl) & "UPDATE MKT_ASN_INVDTL SET ASN_STATUS=1,CUMMULATIVE_QTY=" & dblcummulativeQty & " WHERE UNIT_CODE = '" & gstrUNITID & "' AND DOC_NO=" & pintdocno & " AND CUST_PART_CODE='" & rsgetData.GetValue("CUST_PART_CODE").ToString().Trim() & "' AND CUST_PLANTCODE='" & rsgetData.GetValue("CUST_PLANTCODE").ToString().Trim & "'" & vbCrLf
                    mstrupdateASNCumFig = Trim(mstrupdateASNCumFig) & "UPDATE MKT_ASN_CUMFIG SET CUMMULATIVE_QTY=" & dblcummulativeQty & " WHERE UNIT_CODE = '" & gstrUNITID & "' AND CUST_PART_CODE='" & rsgetData.GetValue("CUST_PART_CODE").ToString().Trim() & "' AND CUST_PLANTCODE='" & rsgetData.GetValue("CUST_PLANTCODE").ToString().Trim & "'" & vbCrLf
                    rsgetData.MoveNext()
                Loop
                Dcount = Dcount + 1
                strASNdata = strASNdata & "856T " & Mid("0000", Dcount.ToString.Length, 5) & Dcount & Mid("00000000", strTotalQty.ToString.Length(), 9) & strTotalQty
                'gstrASNPath = ReadValueFromINI(Application.StartupPath & "\mind.cfg", "ASNPATH-" & gstrUNITID, "Filepath")
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
                rsgetData.ResultSetClose()
                rsgetData = Nothing
                FordASNFileGeneration = True
            End If
        Else
            MessageBox.Show("Unable To Generate ASN File. Invoice Can't Be Locked", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
            FordASNFileGeneration = False
        End If
        Exit Function
ErrHandler:
        FordASNFileGeneration = False
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function FordASNFileGeneration_Rejection(ByVal pintdocno As Double, ByVal pstraccountcode As String) As Boolean
        'Revised By     : Prashant Rajpal
        'Revised On     : 20 nov 2012
        'Reason         : Generate ASN File for Rejection Invoice : FORD
        '--------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim rsgetData As New ClsResultSetDB_Invoice
        Dim strquery As String
        Dim fs As FileStream
        Dim sw As StreamWriter
        Dim strASNdata As String
        Dim Dcount As Integer
        Dim TotalQty As Double
        Dim dblSalesQty As Double
        Dim strcontainerdespQty As String
        Dim dblcummulativeQty As Double
        Dim dblContainerQty As Double
        Dim strTotalQty As String
        Dim strASNFilepath As String
        Dim strASNFilepathforEDI As String
        Dim strtotalquantity As String
        Dim strnoofItems As String

        strASNdata = ""
        strquery = "select * from dbo.FN_GETASNDETAIL_REJECTION(" & pintdocno & ",'" & gstrUNITID & "')"
        rsgetData.GetResult(strquery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsgetData.GetNoRows > 0 Then
            If rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Length = 0 Then
                MessageBox.Show("Customer vendor code is not defined for Customer : " & pstraccountcode, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                FordASNFileGeneration_Rejection = False
                Exit Function
            Else
                If rsgetData.GetValue("CUST_PLANTCODE").ToString.Trim().Length = 0 Then
                    MessageBox.Show("Unable To Get Plant Code For The Customer: " & pstraccountcode & " While Generating ASN File." & vbCrLf &
                                    "Invoice Can't Be Locked", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                    FordASNFileGeneration_Rejection = False
                    Exit Function
                End If
                strASNdata = "856HD20200000000000000000" & Space(5 - rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim.Length) + rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim & rsgetData.GetValue("CUST_PLANTCODE").ToString.Trim & Space(5 - rsgetData.GetValue("CUST_PLANTCODE").ToString.Trim.Length) & ",     ,     *" & vbCrLf
                strASNdata = strASNdata & "856A M" & mInvNo.ToString.Trim() & Space(10 - mInvNo.ToString.Trim().Length) & Space(5 - rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim.Length) & rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim & "09" & VB6.Format(rsgetData.GetValue("INVOICE_DATE"), "yymmdd") & VB6.Format(rsgetData.GetValue("INVOICE_DATE"), "hhmm") & VB6.Format(rsgetData.GetValue("INVOICE_DATE"), "yymmdd") & VB6.Format(rsgetData.GetValue("INVOICE_DATE"), "hhmm") & Space(10) & "+00000000100KG+00000000080KG" & "AE0N" & Space(8) & rsgetData.GetValue("TRANSPORT_TYPE").ToString() & Space("12") & "M" & VB.Right(mInvNo, 5) & Space(4) & Space(35) & "M" & VB.Right(mInvNo, 5) & Space(5) & rsgetData.GetValue("CUST_PLANTCODE").ToString.Trim() & Space(5 - rsgetData.GetValue("CUST_PLANTCODE").ToString.Trim.Length) & rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim & Space(6) & rsgetData.GetValue("ARL_CODE").ToString.Trim() & Space(5 - rsgetData.GetValue("ARL_CODE").ToString.Trim().Length) & VB6.Format(GetServerDateTime, "mmddhhmm") & Space(3) & "0000000.00" & vbCrLf
                Dcount = 2
                strcontainerdespQty = Find_Value("select sum(isnull(to_box,0)-isnull(from_box,0)+1) as Desp_Qty from sales_dtl where  UNIT_CODE = '" & gstrUNITID & "' and  doc_no=" & pintdocno)
                strASNdata = strASNdata & "856TD"
                Select Case rsgetData.GetValue("CONTAINER").ToString.Trim.Length()
                    Case 3
                        strASNdata = strASNdata & rsgetData.GetValue("CONTAINER").ToString.Trim() & "90+" & Mid("000000", strcontainerdespQty.Length(), 6) & strcontainerdespQty.ToString()
                    Case 4
                        strASNdata = strASNdata & rsgetData.GetValue("CONTAINER").ToString.Trim() & " +" & Mid("000000", strcontainerdespQty.Length(), 6) & strcontainerdespQty.ToString()
                    Case 5
                        strASNdata = strASNdata & rsgetData.GetValue("CONTAINER").ToString.Trim() & "+" & Mid("000000", strcontainerdespQty.Length(), 6) & strcontainerdespQty.ToString()
                    Case 1, 2
                        strASNdata = strASNdata & rsgetData.GetValue("CONTAINER").ToString.Trim() & Space(3 - rsgetData.GetValue("CONTAINER").ToString.Trim.Length()) & "  +" & Mid("000000", strcontainerdespQty.Length(), 6) & strcontainerdespQty.ToString()
                    Case Else
                        strASNdata = strASNdata & VB.Left(rsgetData.GetValue("CONTAINER").ToString.Trim(), 5) & "+" & Mid("000000", strcontainerdespQty.Length(), 6) & strcontainerdespQty.ToString()
                End Select

                strtotalquantity = CInt(Find_Value("select sum(isnull(sales_quantity,0)) from sales_dtl where  UNIT_CODE = '" & gstrUNITID & "' and  doc_no=" & pintdocno))
                While Len(strtotalquantity) < 8
                    strtotalquantity = "0" + strtotalquantity
                End While
                strASNdata = strASNdata & strtotalquantity
                strnoofItems = CInt(Find_Value("select COUNT(*) NOOFITEMS  from sales_dtl where  UNIT_CODE = '" & gstrUNITID & "' and  doc_no=" & pintdocno))
                While Len(strnoofItems) < 3
                    strnoofItems = "0" + strnoofItems
                End While
                strASNdata = strASNdata & strnoofItems & vbCrLf


                Dcount = Dcount + 1
                rsgetData.MoveFirst()
                Do While Not rsgetData.EOFRecord
                    dblcummulativeQty = 0
                    dblSalesQty = 0
                    dblContainerQty = 0
                    dblcummulativeQty = Find_Value("SELECT DBO.UDF_GET_CUMMULATIVEQTY_REJECTION('" & gstrUNITID & "','" & rsgetData.GetValue("CUST_PLANTCODE").ToString() & "','" & rsgetData.GetValue("CUST_ITEM_CODE").ToString() & "'," & pintdocno & ")")
                    dblSalesQty = rsgetData.GetValue("SALES_QUANTITY")
                    dblcummulativeQty = dblcummulativeQty + dblSalesQty
                    If mblnZEROASNCUMMS_REJECTIONINVOICE = True Then
                        dblcummulativeQty = 0
                    End If
                    dblContainerQty = rsgetData.GetValue("CONTAINER_QTY")
                    strASNdata = strASNdata & "856P "
                    strASNdata = strASNdata & rsgetData.GetValue("CUST_PART_CODE").ToString().Trim & Space(30 - rsgetData.GetValue("CUST_PART_CODE").ToString().Length())
                    dblSalesQty = rsgetData.GetValue("Sales_Quantity")
                    strASNdata = strASNdata & "BP+" & Mid("0000000", dblSalesQty.ToString.Length(), 8) & dblSalesQty & "EA+"
                    strTotalQty = Val(strTotalQty) + Val(rsgetData.GetValue("SALES_QUANTITY"))
                    Dcount = Dcount + 1
                    strASNdata = strASNdata & Mid("000000000", dblcummulativeQty.ToString().Length(), 10) & dblcummulativeQty
                    strASNdata = strASNdata & "+0000000000" & Space(10) & mInvNo.ToString & Space(11 - mInvNo.ToString.Length()) & rsgetData.GetValue("CUST_VENDOR_CODE").ToString & VB6.Format(rsgetData.GetValue("INVOICE_DATE").ToString(), "yymmdd") & VB6.Format(rsgetData.GetValue("INVOICE_DATE").ToString(), "hhmm") & vbCrLf
                    strASNdata = strASNdata & "856PA" & Space(30) & "+00000000000  +00000000000  " & vbCrLf
                    strASNdata = strASNdata & "856V " & "+000000000000000" & vbCrLf
                    Dcount = Dcount + 2
                    strASNdata = strASNdata & "856C +" & Mid("0000000", dblContainerQty.ToString.Length(), 8) & dblContainerQty & "+" & Mid("0000", rsgetData.GetValue("CONTAINER_DESP_QTY").ToString.Length, 5) & rsgetData.GetValue("CONTAINER_DESP_QTY").ToString & rsgetData.GetValue("CONTAINER").ToString & "90" & vbCrLf
                    Dcount = Dcount + 1
                    mstrupdateASNdtl = Trim(mstrupdateASNdtl) & "UPDATE MKT_ASN_INVDTL SET ASN_STATUS=1,CUMMULATIVE_QTY=" & dblcummulativeQty & " WHERE UNIT_CODE = '" & gstrUNITID & "' AND DOC_NO=" & pintdocno & " AND CUST_PART_CODE='" & rsgetData.GetValue("CUST_PART_CODE").ToString().Trim() & "' AND CUST_PLANTCODE='" & rsgetData.GetValue("CUST_PLANTCODE").ToString().Trim & "'" & vbCrLf
                    mstrupdateASNCumFig = Trim(mstrupdateASNCumFig) & "UPDATE MKT_ASN_CUMFIG SET CUMMULATIVE_QTY=" & dblcummulativeQty & " WHERE UNIT_CODE = '" & gstrUNITID & "' AND CUST_PART_CODE='" & rsgetData.GetValue("CUST_PART_CODE").ToString().Trim() & "' AND CUST_PLANTCODE='" & rsgetData.GetValue("CUST_PLANTCODE").ToString().Trim & "'" & vbCrLf
                    rsgetData.MoveNext()
                Loop
                Dcount = Dcount + 1
                strASNdata = strASNdata & "856T " & Mid("0000", Dcount.ToString.Length, 5) & Dcount & Mid("00000000", strTotalQty.ToString.Length(), 9) & strTotalQty
                'gstrASNPath = ReadValueFromINI(Application.StartupPath & "\mind.cfg", "ASNPATH-" & gstrUNITID, "Filepath")
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
                rsgetData.ResultSetClose()
                rsgetData = Nothing
                FordASNFileGeneration_Rejection = True
            End If
        Else
            MessageBox.Show("Unable To Generate ASN File. Invoice Can't Be Locked", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
            FordASNFileGeneration_Rejection = False
        End If
        Exit Function
ErrHandler:
        FordASNFileGeneration_Rejection = False
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function GENMOTORSASNFileGeneration_THAI(ByVal pintdocno As Double, ByVal pstraccountcode As String) As Boolean
        On Error GoTo ErrHandler
        Dim rsgetData As New ClsResultSetDB_Invoice
        Dim strquery As String
        Dim fs As FileStream
        Dim sw As StreamWriter
        Dim strRecord As String
        Dim Dcount As Integer
        Dim TotalQty As Double
        Dim dblSalesQty As Double
        Dim strcontainerdespQty As String
        Dim dblcummulativeQty As Double
        Dim dblContainerQty As Double
        Dim strTotalQty As String
        Dim strASNFilepath As String
        Dim strASNFilepathforEDI As String
        Dim strtotalquantity As String
        Dim strnoofItems As String


        strquery = "select * from dbo.FN_GETASNDETAIL_GENMOTORS(" & pintdocno & ",'" & gstrUNITID & "')"
        rsgetData.GetResult(strquery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsgetData.GetNoRows > 0 Then
            rsgetData.MoveFirst()
            Do While Not rsgetData.EOFRecord
                dblcummulativeQty = 0
                strRecord = ""
                strRecord = IIf(IsDBNull(rsgetData.GetValue("SuplierEDICode")), "", rsgetData.GetValue("SuplierEDICode"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsgetData.GetValue("CustomerEDICode")), "", rsgetData.GetValue("CustomerEDICode"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsgetData.GetValue("Doc_No")), "", rsgetData.GetValue("Doc_No"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsgetData.GetValue("ASNDate")), "", rsgetData.GetValue("ASNDate"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsgetData.GetValue("DESPDate")), "", rsgetData.GetValue("DESPDate"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsgetData.GetValue("NextMonthDate")), "", rsgetData.GetValue("ASNDateAddMonth"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsgetData.GetValue("Gross_wt")), "", rsgetData.GetValue("Gross_wt"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsgetData.GetValue("Net_Wt")), "", rsgetData.GetValue("Net_Wt"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsgetData.GetValue("MEASURE_UNIT")), "", rsgetData.GetValue("Measure_code"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsgetData.GetValue("LANDING_NO")), "", rsgetData.GetValue("LANDING_NO"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsgetData.GetValue("STPC")), "", rsgetData.GetValue("STPC"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsgetData.GetValue("SPC")), "", rsgetData.GetValue("SPC"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsgetData.GetValue("PD")), "", rsgetData.GetValue("PD"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsgetData.GetValue("MRI")), "", rsgetData.GetValue("MRI"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsgetData.GetValue("MT")), "", rsgetData.GetValue("MT"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsgetData.GetValue("CSC")), "", rsgetData.GetValue("CSC"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsgetData.GetValue("EQ")), "", rsgetData.GetValue("EQ"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsgetData.GetValue("CONVEYANCE_NUMBER")), "", rsgetData.GetValue("CONVEYANCE_NUMBER"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsgetData.GetValue("PACK_SEQ")), "", rsgetData.GetValue("PACK_SEQ"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsgetData.GetValue("PACKING_CODE")), "", rsgetData.GetValue("PACKING_CODE"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsgetData.GetValue("pkg_qty")), "", rsgetData.GetValue("PKG_QTY"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsgetData.GetValue("CUST_ITEM_CODE")), "", rsgetData.GetValue("CUST_ITEM_CODE"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsgetData.GetValue("MODELYEAR")), "", rsgetData.GetValue("MODELYEAR"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsgetData.GetValue("Sales_Quantity")), "", rsgetData.GetValue("Sales_Quantity"))
                dblcummulativeQty = Find_Value("SELECT DBO.UDF_GET_CUMMULATIVEQTY_THAI('" & gstrUNITID & "','" & rsgetData.GetValue("CUST_ITEM_CODE").ToString() & "','" & pstraccountcode & "'," & pintdocno & ")")
                strRecord = strRecord & "," & dblcummulativeQty
                strRecord = strRecord & "," & IIf(IsDBNull(rsgetData.GetValue("cons_measure_code")), "", rsgetData.GetValue("cons_measure_code"))
                strRecord = strRecord & "," & IIf(IsDBNull(rsgetData.GetValue("SONumber")), "", rsgetData.GetValue("SONumber"))
                rsgetData.MoveNext()
            Loop


            gstrASNPath = gstrUserMyDocPath
            gstrASNPathForEDI = Find_Value("SELECT ASNFILEPATH FROM CUSTOMER_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & pstraccountcode & "'")

            If Directory.Exists(gstrASNPath) = False Then
                Directory.CreateDirectory(gstrASNPath)
            End If
            If Directory.Exists(gstrASNPathForEDI) = False Then
                Directory.CreateDirectory(gstrASNPathForEDI)
            End If
            strASNFilepath = gstrASNPath & "\" & pintdocno.ToString() & ".txt"
            strASNFilepathforEDI = gstrASNPathForEDI & "\" & pintdocno.ToString() & ".txt"

            fs = File.Create(strASNFilepath)
            sw = New StreamWriter(fs)
            sw.WriteLine(strRecord)
            sw.Close()
            fs.Close()
            If File.Exists(strASNFilepathforEDI) = False Then
                File.Copy(strASNFilepath, strASNFilepathforEDI)
            End If

            rsgetData.ResultSetClose()
            rsgetData = Nothing
            GENMOTORSASNFileGeneration_THAI = True

        End If

        Exit Function
ErrHandler:
        GENMOTORSASNFileGeneration_THAI = False
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function VerifyGLCCMappingFlag(ByVal Glcode As String) As Boolean
        On Error GoTo ErrHandler
        Dim rsVerifyGLCC_CODEFlag As ClsResultSetDB_Invoice
        rsVerifyGLCC_CODEFlag = New ClsResultSetDB_Invoice
        rsVerifyGLCC_CODEFlag.GetResult("select isnull(GLM_GLCODE,0) as GLM_GLCODE from FIN_GLMASTER WHERE GLM_GLCODE='" & Glcode & "' and UNIT_CODE = '" & gstrUNITID & "'")
        If rsVerifyGLCC_CODEFlag.GetNoRows > 0 Then
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
        objRecordSet.Open("SELECT gbl_ccCode FROM fin_globalgl WHERE gbl_prpsCode = '" & PurposeCode & "' and UNIT_CODE = '" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
        If objRecordSet.EOF Then
            GetCommonCCCode = "N"
            Msgbox("CC Code not defined for  : " & PurposeCode, MsgBoxStyle.Information, "eMPro")
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
            Msgbox("CC Code not defined  for Purpose Code:" & PurposeCode, MsgBoxStyle.Information, "eMPro")
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
    Public Function ValidASNFilePath() As Boolean
        'Issue id : 10160094 
        On Error GoTo Err_Handler
        ValidASNFilePath = True
        '        gstrASNPath = ReadValueFromINI(Application.StartupPath & "\mind.cfg", "ASNPATH-" & gstrUNITID, "Filepath")
        gstrASNPath = gstrUserMyDocPath
        gstrASNPathForEDI = ReadValueFromINI(Application.StartupPath & "\mind.cfg", "ASNPATH-" & gstrUNITID, "FilepathforEDI")
        If gstrASNPath = "" Or gstrASNPathForEDI = "" Then
            ValidASNFilePath = False
            Exit Function
        End If
        'Issue id : 10160094 
        Exit Function
Err_Handler:
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function RejInvOptionalPostingFlag() As Boolean

        On Error GoTo ErrHandler
        Dim rsRejInvPostFlag As ClsResultSetDB_Invoice
        rsRejInvPostFlag = New ClsResultSetDB_Invoice
        rsRejInvPostFlag.GetResult("select isnull(RejInvOptionalPostingFlag,0) as RejInvOptionalPostingFlag from sales_parameter WHERE UNIT_CODE='" + gstrUNITID + "'")
        If rsRejInvPostFlag.GetNoRows > 0 Then
            RejInvOptionalPostingFlag = rsRejInvPostFlag.GetValue("RejInvOptionalPostingFlag")
        Else
            RejInvOptionalPostingFlag = False
        End If
        rsRejInvPostFlag.ResultSetClose()
        Exit Function
ErrHandler:
        RejInvOptionalPostingFlag = False
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

        rsTKMLBARCODE = New ClsResultSetDB_Invoice
        rsSALESCHALLANDTL = New ClsResultSetDB_Invoice
        rsSALESDTL = New ClsResultSetDB_Invoice
        rsTKMLBARCODE.GetResult("SELECT ISNULL(TKML_BARCODELABEL_FORMAT,'') TKML_BARCODELABEL_FORMAT FROM SALES_PARAMETER WHERE UNIT_CODE='" & gstrUNITID & "' ")
        If rsTKMLBARCODE.GetNoRows > 0 Then
            StrBarcodelabelFormat = rsTKMLBARCODE.GetValue("TKML_BARCODELABEL_FORMAT").ToString
            SW = File.CreateText(gstrUserMyDocPath + "TKML_BARCODELABEL.TXT")
            StrNewlabel = StrBarcodelabelFormat
            rsSALESCHALLANDTL.GetResult("SELECT Distinct Lorryno_date,SD.doc_no,invoice_date,total_amount,Ecess_amount,Secess_amount, Sales_tax_amount FROM SALES_DTL SD ,SALESCHALLAN_DTL SC WHERE SC.UNIT_CODE=SD.UNIT_CODE AND " &
                                    " SC.DOC_NO=SD.DOC_NO  AND SC.UNIT_CODE='" & gstrUNITID & "' AND  " &
                                    " SC.DOC_NO= " & pstrInvNo & "")
            intmaxRows = rsSALESCHALLANDTL.GetNoRows
            If intmaxRows > 0 Then
                strBarcode = rsSALESCHALLANDTL.GetValue("Lorryno_date").ToString & ","
                strBarcode = strBarcode + rsSALESCHALLANDTL.GetValue("doc_no").ToString & ","
                strBarcode = strBarcode + VB6.Format(rsSALESCHALLANDTL.GetValue("invoice_date"), "ddmmyy") & ","
                strtotalBasicAmount = Find_Value("SELECT SUM(ISNULL(BASIC_AMOUNT,0)) AS TOTALBASIC_AMOUNT FROM SALES_DTL WHERE  UNIT_CODE = '" & gstrUNITID & "' AND  DOC_NO=" & pstrInvNo)
                strBarcode = strBarcode + VB6.Format(strtotalBasicAmount, "###.00").ToString & ","

                strtotalExciseAmount = Find_Value("SELECT SUM(ISNULL(EXCISE_TAX,0)) AS TOTALEXCISE_AMOUNT FROM SALES_DTL WHERE  UNIT_CODE = '" & gstrUNITID & "' AND  DOC_NO=" & pstrInvNo)
                strBarcode = strBarcode + VB6.Format(strtotalExciseAmount, "###.00").ToString & ","
                strBarcode = strBarcode + rsSALESCHALLANDTL.GetValue("ecess_amount").ToString & ","
                strBarcode = strBarcode + rsSALESCHALLANDTL.GetValue("secess_amount").ToString & ","
                strBarcode = strBarcode + VB6.Format(rsSALESCHALLANDTL.GetValue("Sales_tax_amount").ToString, "###.00") & ","
                strBarcode = strBarcode + VB6.Format(rsSALESCHALLANDTL.GetValue("total_amount").ToString, "###.00") & ","
                strBarcode = strBarcode + "1/1~"

                rsSALESDTL.GetResult("SELECT CUST_ITEM_CODE,SALES_QUANTITY  FROM SALES_DTL SD ,SALESCHALLAN_DTL SC WHERE SC.UNIT_CODE=SD.UNIT_CODE AND " &
                                    " SC.DOC_NO=SD.DOC_NO  AND SC.UNIT_CODE='" & gstrUNITID & "' AND  " &
                                    " SC.DOC_NO= " & pstrInvNo & "")

                intmaxRows = rsSALESDTL.GetNoRows

                rsSALESDTL.MoveFirst()
                For intRow = 1 To intmaxRows
                    strBarcode = strBarcode + rsSALESDTL.GetValue("Cust_Item_Code").ToString & "," + Val(rsSALESDTL.GetValue("sales_quantity")).ToString & "~"
                    rsSALESDTL.MoveNext()
                Next

            End If

            StrNewlabel = StrNewlabel.Replace("V_SRVDINO", rsSALESCHALLANDTL.GetValue("Lorryno_date").ToString)
            StrNewlabel = StrNewlabel.Replace("V_Barcode", strBarcode)
            StrNewlabel = StrNewlabel.Replace("V_Invno", rsSALESCHALLANDTL.GetValue("doc_no").ToString)
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
        SW = File.CreateText(gstrUserMyDocPath + "TKML_BARCODELABEL.BAT")
        SW.WriteLine("CD\")
        SW.WriteLine("C:")
        SW.WriteLine("MODE:LPT1")
        SW.WriteLine("COPY """ & gstrUserMyDocPath & "TKML_BARCODELABEL.TXT"" LPT1")
        SW.Close()

        Shell("cmd.exe /c """ & gstrUserMyDocPath & "TKML_BARCODELABEL.BAT""", AppWinStyle.MinimizedNoFocus)
        Msgbox("Labels printed successfully.", MsgBoxStyle.Information, ResolveResString(100))
        Exit Function
ErrHandler:
        Print_barcodelabel = False
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Public Function CheckInvoices(ByVal strInvoiceno As String, ByVal straccountcode As String) As Boolean

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
        Dim strFTPPath As String = String.Empty
        Dim strFTPEDIPATH As String = String.Empty
        Dim strFTPPATHforlog As String = String.Empty
        Dim blnFTPwithEDI As Boolean
        Dim blnNewData As Boolean
        Dim strBuffer(14) As String
        Dim nInFile As Short 'File Handle of the arguments file

        If gstrUNITID = "SML" Then
            VCode = "S073"
        ElseIf gstrUNITID = "M03" Then
            VCode = "M582"
        ElseIf gstrUNITID = "MST" Then
            VCode = "M581"
        End If

        nInFile = FreeFile()
        'mstrins = ""

        'FileOpen(nInFile, My.Application.Info.DirectoryPath & "\" & "FTParguments.cfg", OpenMode.Input)
        'Dim counter As Short
        'counter = 1


        'While Not EOF(nInFile)
        '    strBuffer(counter) = LineInput(nInFile)
        '    counter = counter + 1
        'End While

        strFTPPath = ReadValueFromINI(Application.StartupPath & "\MultipleFTpArgument.cfg", "FTPPATH-" & gstrUNITID, "FilepathFTP")
        strFTPEDIPATH = ReadValueFromINI(Application.StartupPath & "\MultipleFTpArgument.cfg", "FTPPATH-" & gstrUNITID, "FilepathforFTPEDI")
        strFTPPATHforlog = ReadValueFromINI(Application.StartupPath & "\MultipleFTpArgument.cfg", "FTPPATH-" & gstrUNITID, "FilepathforFTPlog")
        'gstrASNPathForEDI = ReadValueFromINI(Application.StartupPath & "\mind.cfg", "ASNPATH-" & gstrUNITID, "FilepathforEDI")
        'blnFTPwithEDI = GetArgumentValue(strBuffer(13))
        'pstrTempPathForEDI = GetArgumentValue(strBuffer(14)) ' Path of temp. files for EDI
        'pstrTempPath = GetArgumentValue(strBuffer(3)) ' Path of temp. files for FTP

        'FileClose(nInFile)
        'If Dir(pstrTempPath & "\Invoices", FileAttribute.Directory) = "" Then MkDir(pstrTempPath & "\Invoices\")

        'Check whether folder c:\temp exist or not for EDI
        If Not Directory.Exists(strFTPPath & "\Invoices") Then
            Directory.CreateDirectory(strFTPPath & "\Invoices")
        End If
        'If Dir(pstrTempPathForEDI & "\Invoices\", FileAttribute.Directory) = "" Then MkDir(pstrTempPathForEDI & "\Invoices\")
        If Not Directory.Exists(strFTPEDIPATH & "\Invoices") Then
            Directory.CreateDirectory(strFTPEDIPATH & "\Invoices")
        End If

        strSql = "select * from dbo.FN_GET_FTP_FILEDATA(" & strInvoiceno & ",'" & gstrUNITID & "')"
        rs_hdr = New ClsResultSetDB_Invoice
        rs_hdr.GetResult(strSql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)

        Dim rsGetAmount As ClsResultSetDB_Invoice
        Dim stramt As String
        Do While Not rs_hdr.EOFRecord
            strSBUFolder = strFTPPath & "\Invoices\"
            strEDIFolder = strFTPEDIPATH & "\Invoices\"

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
            'strSql = "SELECT CUST_PART_CODE, DSNO,QUANTITYKNOCKEDOFF " & " FROM MKT_INVDSHISTORY H " & " INNER JOIN SALESCHALLAN_DTL SC ON SC.UNIT_CODE=H.UNIT_CODE AND SC.DOC_NO = H.DOC_NO " & " AND SC.LOCATION_CODE = H.LOCATION_CODE " & " AND SC.ACCOUNT_CODE = H.CUSTOMER_CODE" & " INNER JOIN SALES_DTL SD ON SC.UNIT_CODE=SD.UNIT_CODE AND SC.DOC_NO = SD.DOC_NO" & " AND SC.LOCATION_CODE = SD.LOCATION_CODE" & " AND SD.ITEM_CODE = H.ITEM_CODE" & " AND SC.UNIT_CODE='" & gstrUNITID & "' AND SC.DOC_NO = " & minv
            'objGetDSData = New ClsResultSetDB_Invoice
            'objGetDSData.GetResult(strSql)
            'strDSFileData = ""
            'Do While Not objGetDSData.EOFRecord
            '    strDSFileData = strDSFileData & objGetDSData.GetValue("CUST_PART_CODE").ToString & "|" & objGetDSData.GetValue("DSNO").ToString & "|" & objGetDSData.GetValue("QUANTITYKNOCKEDOFF").ToString & "^"
            '    objGetDSData.MoveNext()
            'Loop
            'objGetDSData.ResultSetClose()
            strSql = "SELECT CUST_PART_CODE, CASE WHEN DSNO='ECSS' THEN left(SC.CUST_REF,10) ELSE DSNO END DSNO ,QUANTITYKNOCKEDOFF ,SD.RATE,SD.HSNSACCODE,SD.CGSTTXRT_TYPE,SD.HSNSACCODE,SD.CGST_PERCENT,SD.CGST_AMT,SD.SGSTTXRT_TYPE,SD.SGST_PERCENT,SD.SGST_AMT, SD.IGSTTXRT_TYPE,SD.IGST_PERCENT,SD.IGST_AMT   FROM MKT_INVDSHISTORY H  INNER JOIN SALESCHALLAN_DTL SC ON SC.UNIT_CODE=H.UNIT_CODE AND SC.DOC_NO = H.DOC_NO  AND SC.LOCATION_CODE = H.LOCATION_CODE  AND SC.ACCOUNT_CODE = H.CUSTOMER_CODE INNER JOIN SALES_DTL SD ON SC.UNIT_CODE=SD.UNIT_CODE AND SC.DOC_NO = SD.DOC_NO AND SC.LOCATION_CODE = SD.LOCATION_CODE AND SD.ITEM_CODE = H.ITEM_CODE AND SC.UNIT_CODE='" & gstrUNITID & "' AND SC.DOC_NO = " & minv
            objGetDSData = New ClsResultSetDB_Invoice
            objGetDSData.GetResult(strSql)
            strDSFileData = ""
            Do While Not objGetDSData.EOFRecord
                strDSFileData = strDSFileData & objGetDSData.GetValue("CUST_PART_CODE").ToString & "|" & objGetDSData.GetValue("DSNO").ToString & "|" & objGetDSData.GetValue("QUANTITYKNOCKEDOFF").ToString & "|"
                strDSFileData = strDSFileData & "0|" & objGetDSData.GetValue("HSNSACCODE").ToString & "|" & objGetDSData.GetValue("RATE").ToString & "|"
                strDSFileData = strDSFileData & "CGST|" & objGetDSData.GetValue("CGST_PERCENT").ToString & "|" & objGetDSData.GetValue("CGST_AMT").ToString & "|"
                strDSFileData = strDSFileData & "SGST|" & objGetDSData.GetValue("SGST_PERCENT").ToString & "|" & objGetDSData.GetValue("SGST_AMT").ToString & "|"
                strDSFileData = strDSFileData & "IGST|" & objGetDSData.GetValue("IGST_PERCENT").ToString & "|" & objGetDSData.GetValue("IGST_AMT").ToString & "|"
                strDSFileData = strDSFileData & "TAX4|0|0|TAX5|0|0^"
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
    Public Function UpdateMktSchedule() As Boolean
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
        blnDSTracking = CBool(Find_Value("SELECT isnull(DSWiseTracking,0)as DSWiseTracking FROM sales_parameter WHERE UNIT_CODE='" + gstrUNITID + "'"))
        straccountcode = Find_Value("select account_code from saleschallan_dtl WHERE UNIT_CODE='" + gstrUNITID + "' AND  doc_no='" & mInvNo & "'")
        If blnDSTracking = True Then
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
                        .Parameters.Append(.CreateParameter("@DOC_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, mInvNo))
                        .Parameters.Append(.CreateParameter("@ACCOUNT_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 8, Trim(rsGetData.GetValue("Account_code"))))
                        .Parameters.Append(.CreateParameter("@ITEM_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 30, Trim(rsGetData.GetValue("Item_Code"))))
                        .Parameters.Append(.CreateParameter("@CUSTITEM_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 30, Trim(rsGetData.GetValue("Cust_item_Code"))))
                        .Parameters.Append(.CreateParameter("@REQ_QTY", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, ADODB.DataTypeEnum.adCurrency, Val(rsGetData.GetValue("sales_quantity"))))
                        .Parameters.Append(.CreateParameter("@DATE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 11, getDateForDB(Trim(rsGetData.GetValue("INvoice_date")))))
                        .Parameters.Append(.CreateParameter("@USERID", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 16, gstrUserIDSelected))
                        .Parameters.Append(.CreateParameter("@MSG", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamOutput, 500))
                        .Parameters.Append(.CreateParameter("@ERR", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamOutput, 100))
                        .ActiveConnection = mP_Connection
                        .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        If Len(.Parameters(10).Value) > 0 Then
                            Msgbox(.Parameters(11).Value, vbInformation + vbOKOnly, ResolveResString(100))
                            UpdateMktSchedule = False
                            Com = Nothing
                            Exit Function
                        End If
                        If Len(.Parameters(9).Value) > 0 Then
                            Msgbox(.Parameters(9).Value, vbInformation + vbOKOnly, ResolveResString(100))
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
    Private Function AllowA4Reports(ByVal pstraccountcode As String) As Boolean
        On Error GoTo ErrHandler
        '10617093
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
    Private Function ALLOW_ANNEXTUREPRINTING(ByVal pstraccountcode As String) As Boolean
        On Error GoTo ErrHandler
        '10617093
        'PURPOSE: THIS FUNCTION IS TO KNOW WHETHER ANNEXTURE WILL BE PRINTING FOR PARTICULAR CUSTOMER OR NOT

        Dim strQry As String
        Dim Rs As ClsResultSetDB_Invoice

        ALLOW_ANNEXTUREPRINTING = False
        strQry = "Select isnull(ALLOW_ANNEXTUREPRINTING,0) as ALLOW_ANNEXTUREPRINTING from customer_mst where UNIT_CODE='" + gstrUNITID + "' AND  Customer_Code='" & Trim(pstraccountcode) & "'"
        Rs = New ClsResultSetDB_Invoice
        If Rs.GetResult(strQry) = False Then GoTo ErrHandler
        If Rs.GetValue("ALLOW_ANNEXTUREPRINTING") = "True" Then
            ALLOW_ANNEXTUREPRINTING = True
        Else
            ALLOW_ANNEXTUREPRINTING = False
        End If

        Rs.ResultSetClose()
        Rs = Nothing

        Exit Function
ErrHandler:
        Rs = Nothing
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function Printbarcode_TOYOTA(ByVal pstrInvNo As String) As Boolean

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

        rsTKMLBARCODE = New ClsResultSetDB_Invoice
        rsSALESCHALLANDTL = New ClsResultSetDB_Invoice
        rsSALESDTL = New ClsResultSetDB_Invoice
        rsTKMLBARCODE.GetResult("SELECT ISNULL(TKML_BARCODELABEL_FORMAT,'') TKML_BARCODELABEL_FORMAT FROM SALES_PARAMETER WHERE UNIT_CODE='" & gstrUNITID & "' ")
        If rsTKMLBARCODE.GetNoRows > 0 Then

            rsSALESCHALLANDTL.GetResult("SELECT Distinct Lorryno_date,SD.doc_no,invoice_date,total_amount,Ecess_amount,Secess_amount, Sales_tax_amount FROM SALES_DTL SD ,SALESCHALLAN_DTL SC WHERE SC.UNIT_CODE=SD.UNIT_CODE AND " &
                                    " SC.DOC_NO=SD.DOC_NO  AND SC.UNIT_CODE='" & gstrUNITID & "' AND  " &
                                    " SC.DOC_NO= " & pstrInvNo & "")
            intmaxRows = rsSALESCHALLANDTL.GetNoRows
            If intmaxRows > 0 Then
                strBarcode = rsSALESCHALLANDTL.GetValue("Lorryno_date").ToString & ","
                strBarcode = strBarcode + rsSALESCHALLANDTL.GetValue("doc_no").ToString & ","
                strBarcode = strBarcode + VB6.Format(rsSALESCHALLANDTL.GetValue("invoice_date"), "ddmmyy") & ","
                strtotalBasicAmount = Find_Value("SELECT SUM(ISNULL(BASIC_AMOUNT,0)) AS TOTALBASIC_AMOUNT FROM SALES_DTL WHERE  UNIT_CODE = '" & gstrUNITID & "' AND  DOC_NO=" & pstrInvNo)
                strBarcode = strBarcode + VB6.Format(strtotalBasicAmount, "###.00").ToString & ","

                strtotalExciseAmount = Find_Value("SELECT SUM(ISNULL(EXCISE_TAX,0)) AS TOTALEXCISE_AMOUNT FROM SALES_DTL WHERE  UNIT_CODE = '" & gstrUNITID & "' AND  DOC_NO=" & pstrInvNo)
                strBarcode = strBarcode + VB6.Format(strtotalExciseAmount, "###.00").ToString & ","
                strBarcode = strBarcode + rsSALESCHALLANDTL.GetValue("ecess_amount").ToString & ","
                strBarcode = strBarcode + rsSALESCHALLANDTL.GetValue("secess_amount").ToString & ","
                strBarcode = strBarcode + VB6.Format(rsSALESCHALLANDTL.GetValue("Sales_tax_amount").ToString, "###.00") & ","
                strBarcode = strBarcode + VB6.Format(rsSALESCHALLANDTL.GetValue("total_amount").ToString, "###.00") & ","
                strBarcode = strBarcode + "1/1~"

                rsSALESDTL.GetResult("SELECT CUST_ITEM_CODE,SALES_QUANTITY  FROM SALES_DTL SD ,SALESCHALLAN_DTL SC WHERE SC.UNIT_CODE=SD.UNIT_CODE AND " &
                                    " SC.DOC_NO=SD.DOC_NO  AND SC.UNIT_CODE='" & gstrUNITID & "' AND  " &
                                    " SC.DOC_NO= " & pstrInvNo & "")

                intmaxRows = rsSALESDTL.GetNoRows

                rsSALESDTL.MoveFirst()
                For intRow = 1 To intmaxRows
                    strBarcode = strBarcode + rsSALESDTL.GetValue("Cust_Item_Code").ToString & "," + Val(rsSALESDTL.GetValue("sales_quantity")).ToString & "~"
                    rsSALESDTL.MoveNext()
                Next
            End If

        Else
            Printbarcode_TOYOTA = False
        End If
        Dim strPath As String = gstrUserMyDocPath & "ToyotaBarcodeImg.JPEG"
        Dim ts As Object
        Dim fso As New Scripting.FileSystemObject
        Dim stimage As ADODB.Stream
        Dim strQuery As String
        Dim Rs As ADODB.Recordset
        Dim BarcodeImage As Bitmap = create2DImage(GetEncodedString(strBarcode.Trim(), ""))
        BarcodeImage.Save(strPath)
        'stimage.LoadFromFile(strPath)
        strQuery = "select  barcodeimage  from saleschallan_dtl where doc_no='" & Trim(pstrInvNo) & "' and UNIT_CODE = '" & gstrUNITID & "'"
        Rs = New ADODB.Recordset
        Rs.Open(strQuery, mP_Connection, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
        Rs.Fields("barcodeimage").Value = GetFileBytes(strPath)
        Rs.Update()
        Rs.Close()
        Rs = Nothing

        rsTKMLBARCODE.ResultSetClose()
        rsSALESDTL.ResultSetClose()
        rsSALESCHALLANDTL.ResultSetClose()
        'SW.Close()
        Exit Function
ErrHandler:
        Printbarcode_TOYOTA = False
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Public Function GetFileBytes(ByVal strFilePath As String) As Byte()

        Try
            If Not File.Exists(strFilePath) Then Return Nothing

            Dim oFs As New FileStream(strFilePath, FileMode.Open, FileAccess.Read)
            Dim oBinaryReader As New BinaryReader(oFs)
            Dim FileBytes As Byte() = oBinaryReader.ReadBytes(CInt(oFs.Length))

            oBinaryReader.Close()
            oFs.Close()
            Return FileBytes
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Function

    Private Function EncodedString(ByVal DataString As String) As String
        Dim objCRUFLQRC As New CRUFLQRC.UFL()
        Return objCRUFLQRC.MW6Encoder(DataString, 0, 0, 0)
    End Function
    Private Function create2DImage(ByVal data As String) As Bitmap
        Dim barcode As New Bitmap(1, 1)

        Dim PFC As New PrivateFontCollection
        PFC.AddFontFile("c:\windows\fonts\MW6Matrix.TTF")
        Dim FF As FontFamily = PFC.Families(0)
        Dim fontName As New Font(FF, 30)

        Dim graphics__1 As Graphics = Graphics.FromImage(barcode)
        Dim dataSize As SizeF = graphics__1.MeasureString(data, fontName)

        barcode = New Bitmap(barcode, dataSize.ToSize())
        graphics__1 = Graphics.FromImage(barcode)
        graphics__1.Clear(Color.White)
        graphics__1.TextRenderingHint = TextRenderingHint.SingleBitPerPixel

        graphics__1.DrawString(data, fontName, New SolidBrush(Color.Black), 0, 0)
        graphics__1.Flush()
        fontName.Dispose()
        graphics__1.Dispose()

        Return barcode
    End Function
    Private Function GetEncodedString(ByVal DataString As String, ByVal DataType As String) As String
        Dim Encoded_1 As String = EncodedString(DataString)
        System.GC.Collect()
        Dim Encoded_2 As String = EncodedString(DataString)

        If Encoded_1 = Encoded_2 Then
            Return Encoded_1
        Else
            For count As Integer = 0 To 4
                Encoded_1 = String.Empty
                Encoded_2 = String.Empty

                Encoded_1 = EncodedString(DataString)
                Encoded_2 = EncodedString(DataString)

                If Encoded_1 = Encoded_2 Then
                    Return Encoded_1
                End If
            Next
        End If

        If Encoded_1 = Encoded_2 Then
            Return Encoded_1
        Else
            Throw New Exception(DataType & " String encoding mismatch.")
        End If
    End Function
    Public Function SaveBinningBarcodeImage_singlelevelso_2DBARCODE(ByVal pstrInvNo As String, ByVal pstrPath As String, ByVal strbarcodestring As String) As Boolean
        On Error GoTo ErrHandler
        Dim stimage As ADODB.Stream
        Dim strQuery As String
        Dim Rs As ADODB.Recordset
        SaveBinningBarcodeImage_singlelevelso_2DBARCODE = True
        stimage = New ADODB.Stream
        stimage.Type = ADODB.StreamTypeEnum.adTypeBinary
        stimage.Open()
        pstrPath = pstrPath & ".JPEG"
        stimage.LoadFromFile(pstrPath)
        strQuery = "select  Bin_Image  from ASN_Lebels_temp_2D where invoiceno='" & Trim(pstrInvNo) & "' and vendorcode = '" & gstrUNITID & "'"
        Rs = New ADODB.Recordset
        Rs.Open(strQuery, mP_Connection, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
        Rs.Fields("Bin_Image").Value = stimage.Read
        Rs.Update()
        Rs.Close()
        Rs = Nothing
        Exit Function

ErrHandler:
        SaveBinningBarcodeImage_singlelevelso_2DBARCODE = False
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)

    End Function


#Region "GLOBAL TOOL CHANAGES"
    'ADDED BY VINOD FOR GLOBAL TOOL CHANGES
    Private Function UpdateGlobalTool(ByVal intInvNo As String, ByVal intTempInvNo As String) As Boolean
        Dim Cmd As New ADODB.Command
        Try
            With Cmd
                .CommandText = "USP_UPDATE_GLOBAL_TOOL_INVOICE_LOCK"
                .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                .CommandTimeout = 0
                .let_ActiveConnection(mP_Connection)
                .Parameters.Append(.CreateParameter("@UNIT_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                .Parameters.Append(.CreateParameter("@TEMP_INVOICE_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, intTempInvNo))
                .Parameters.Append(.CreateParameter("@INVOICE_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, intInvNo))
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

    Public Function SaveQRBarCodeImage(ByVal pstrTempInvNo As String, ByVal intFromLineNo As Integer, ByVal intToLineNo As Integer, ByVal strbarcodestring As String, ByVal intRow As Integer, ByVal intTotalNoofSlabs As Integer,
                                        Optional ByVal pstrActualInvNo As String = "") As Boolean

        On Error GoTo ErrHandler

        Dim stimage As ADODB.Stream
        Dim strQuery As String
        Dim Rs As ADODB.Recordset
        Dim pstrPath As String = ""
        Dim strSql As String = ""
        Dim strserial As String = ""
        Dim blnCROP_QRIMAGE As Boolean = False
        pstrPath = gstrUserMyDocPath
        SaveQRBarCodeImage = True

        'If pstrActualInvNo.Trim.Length > 0 Then
        'strSql = "Select Top 1 1 from INVOICE_QRIMAGE WHERE unit_Code='" & gstrUNITID & "' And INVOICE_NO='" & pstrActualInvNo & "' And FROMLINENo=" & intFromLineNo & " AND TOLINENO=" & intToLineNo & ""
        'Else
        strSql = "Select Top 1 1 from INVOICE_QRIMAGE WHERE unit_Code='" & gstrUNITID & "' And TMP_INVOICENO='" & pstrTempInvNo & "' And FROMLINENo=" & intFromLineNo & " AND TOLINENO=" & intToLineNo & ""
        'End If
        strserial = CStr(intRow) + "/" + CStr(intTotalNoofSlabs)

        If DataExist(strSql) = False Then
            strSql = "INSERT INTO INVOICE_QRIMAGE (INVOICE_NO,TMP_INVOICENO,FROMLINENO,TOLINENO,UNIT_CODE ,SERIALHEADING,BARCODESTRING ) VALUES('" & pstrTempInvNo & "','" & pstrTempInvNo & "'," & intFromLineNo & "," & intToLineNo & ",'" & gstrUNITID & "','" & strserial & "','" & strbarcodestring & "')"
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

        strQuery = "select  BARCODEIMG,INVOICE_NO  from INVOICE_QRIMAGE where UNIT_CODE = '" & gstrUNITID & "' AND TMP_INVOICENO='" & pstrTempInvNo & "' And FROMLINENO =" & intFromLineNo & " and TOLINENO =" & intToLineNo & ""
        Rs = New ADODB.Recordset
        Rs.Open(strQuery, mP_Connection, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)

        Rs.Fields("BARCODEIMG").Value = stimage.Read
        If pstrActualInvNo.Trim.Length > 0 Then
            Rs.Fields("INVOICE_NO").Value = mInvNo
        End If

        Rs.Update()
        Rs.Close()
        Rs = Nothing

        Exit Function
ErrHandler:
        SaveQRBarCodeImage = False
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
    Private Function AllowASNTextFile(ByVal pstraccountcode As String) As Boolean
        '**********************************************************************************************
        'ISSUE ID 101041360  - EDI INVOICE FOR FSP
        'PURPOSE :TO GENERATE THE .TXT FILE 
        '**********************************************************************************************

        On Error GoTo ErrHandler
        Dim strQry As String
        Dim Rs As ClsResultSetDB_Invoice

        AllowASNTextFile = False
        If (UCase(Trim(cmbInvType.Text)) <> "REJECTION") Then
            strQry = "Select isnull(AllowASNTextFILE,0) as AllowASNTextFILE from customer_mst where Customer_Code='" & Trim(pstraccountcode) & "' and UNIT_CODE = '" & gstrUNITID & "'"
            Rs = New ClsResultSetDB_Invoice
            If Rs.GetResult(strQry) = False Then GoTo ErrHandler
            If Rs.GetValue("AllowASNTextFILE") = "True" Then
                AllowASNTextFile = True
            Else
                AllowASNTextFile = False
            End If
            Rs.ResultSetClose()
            Rs = Nothing
        End If
        Exit Function

ErrHandler:
        Rs = Nothing
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

    Private Function ASNTEXTFILE_DETAILS(ByVal pstrInvNo As String, ByVal straccountcode As String) As Boolean
        '**********************************************************************************************
        'ISSUE ID 101041360  - EDI INVOICE FOR FSP
        'PURPOSE :TO GENERATE THE .TXT FILE 
        '**********************************************************************************************
        On Error GoTo ErrHandler
        Dim rsgetData As New ClsResultSetDB_Invoice
        Dim strquery As String
        Dim fs As FileStream
        Dim sw As StreamWriter
        Dim strASNdata As String
        Dim Dcount As Integer
        Dim TotalQty As Double
        Dim dblSalesQty As Double
        Dim strcontainerdespQty As String
        Dim dblcummulativeQty As Double
        Dim dblContainerQty As Double
        Dim strTotalQty As String
        Dim strASNFilepath As String
        Dim strASNFilepathforEDI As String
        Dim strSQL As String
        Dim ASNPath As String
        Dim ASNPathForEDI As String
        Dim intloopcounter As Integer
        Dim intmaxloop As Integer
        strASNdata = ""

        strquery = "select * from dbo.FN_FORDREQD_INVOICEDETAILS(" & pstrInvNo & ",'" & gstrUNITID & "')"
        rsgetData.GetResult(strquery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsgetData.GetNoRows > 0 Then
            If rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Length = 0 Then
                MessageBox.Show("Customer vendor code is not defined for Customer : " & straccountcode, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                ASNTEXTFILE_DETAILS = False
                Exit Function
            Else
                If rsgetData.GetValue("PLANT_CODE").ToString.Trim().Length = 0 Then
                    MessageBox.Show("Unable To Get Plant Code For The Customer: " & straccountcode & " While Generating ASN File." & vbCrLf &
                                    "Invoice Can't Be Locked", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                    ASNTEXTFILE_DETAILS = False
                    Exit Function
                End If
                strASNdata = ""

                intmaxloop = rsgetData.GetNoRows
                rsgetData.MoveFirst()
                For intloopcounter = 1 To intmaxloop
                    'strASNdata = strASNdata & rsgetData.GetValue("CUST_VENDOR_CODE").ToString().Trim & Space(5 - rsgetData.GetValue("CUST_VENDOR_CODE").ToString().Length()) & ","
                    'strASNdata = strASNdata & rsgetData.GetValue("PLANT_CODE").ToString().Trim & Space(5 - rsgetData.GetValue("PLANT_CODE").ToString().Length()) & ","
                    'strASNdata = strASNdata & pstrInvNo & "," & rsgetData.GetValue("INVOICE_DATE").ToString.Trim & "," & rsgetData.GetValue("CUST_REF").ToString.Trim & ","
                    'strASNdata = strASNdata & rsgetData.GetValue("PACKING_LIST").ToString.Trim & "," & rsgetData.GetValue("SUPPLIERID").ToString.Trim & ","
                    ''strASNdata = strASNdata & rsgetData.GetValue("SHIPTOFORDPLANT").ToString.Trim & "," & rsgetData.GetValue("CUST_ITEM_CODE ").ToString.Trim & ","
                    'strASNdata = strASNdata & rsgetData.GetValue("SHIPTOFORDPLANT").ToString.Trim & "," & rsgetData.GetValue("CURRENCY_CODE") & "," & rsgetData.GetValue("CUST_ITEM_CODE ").ToString.Trim & ","
                    'strASNdata = strASNdata & Val(rsgetData.GetValue("SALES_QUANTITY")) & "," & rsgetData.GetValue("CONS_MEASURE_CODE").ToString.Trim & ","
                    'strASNdata = strASNdata & rsgetData.GetValue("RATE").ToString.Trim & "," & rsgetData.GetValue("TOOL_COST").ToString.Trim & ","
                    'strASNdata = strASNdata & rsgetData.GetValue("CHILDPARTCOST").ToString.Trim & "," & rsgetData.GetValue("BASIC_AMOUNT").ToString.Trim & ","
                    'strASNdata = strASNdata & rsgetData.GetValue("ACCESSIBLE_AMOUNT").ToString.Trim & "," & rsgetData.GetValue("EXCISE_AMOUNT").ToString.Trim & ","
                    'strASNdata = strASNdata & intmaxloop & "," & rsgetData.GetValue("VALUEONCALCULATEVAT").ToString.Trim & "," & rsgetData.GetValue("CUSTOMDUTY").ToString.Trim & ","
                    ''strASNdata = strASNdata & rsgetData.GetValue("BASICTOOLCOST").ToString.Trim & ",0,0,0,0,"
                    'strASNdata = strASNdata & rsgetData.GetValue("BASICTOOLCOST").ToString.Trim & ",0,0,0,"
                    'strASNdata = strASNdata & rsgetData.GetValue("VALUEONCALCULATEADDVAT").ToString.Trim & ","
                    'strASNdata = strASNdata & rsgetData.GetValue("TOTAL_AMOUNT").ToString.Trim & "," & rsgetData.GetValue("TOTALBASICAMOUNT").ToString.Trim & ","
                    'strASNdata = strASNdata & rsgetData.GetValue("SALESTAXTYPE").ToString.Trim & "," & Val(rsgetData.GetValue("SALESTAX_PER")) & ","
                    'strASNdata = strASNdata & rsgetData.GetValue("SALES_TAX_AMOUNT").ToString.Trim & ",LOC,0,0,EXC,"
                    'strASNdata = strASNdata & rsgetData.GetValue("EXCISE_PER").ToString.Trim & "," & rsgetData.GetValue("TOTALEXCISEAMOUNT").ToString.Trim
                    'strASNdata = strASNdata & ",CUD,0,0,CVD,0,0,ADD,0,0,OTH,0,0"
                    strASNdata = strASNdata & rsgetData.GetValue("CUST_VENDOR_CODE").ToString().Trim & Space(5 - rsgetData.GetValue("CUST_VENDOR_CODE").ToString().Length()) & ","
                    strASNdata = strASNdata & rsgetData.GetValue("PLANT_CODE").ToString().Trim & Space(5 - rsgetData.GetValue("PLANT_CODE").ToString().Length()) & ","
                    strASNdata = strASNdata & pstrInvNo & "," & rsgetData.GetValue("INVOICE_DATE").ToString.Trim & "," & rsgetData.GetValue("CUST_REF").ToString.Trim & ","
                    strASNdata = strASNdata & rsgetData.GetValue("PACKING_LIST").ToString.Trim & "," & rsgetData.GetValue("SUPPLIERID").ToString.Trim & ","
                    'strASNdata = strASNdata & rsgetData.GetValue("SHIPTOFORDPLANT").ToString.Trim & "," & rsgetData.GetValue("CUST_ITEM_CODE ").ToString.Trim & ","
                    strASNdata = strASNdata & rsgetData.GetValue("SHIPTOFORDPLANT").ToString.Trim & "," & rsgetData.GetValue("CURRENCY_CODE") & "," & rsgetData.GetValue("CUST_ITEM_CODE ").ToString.Trim & ","
                    strASNdata = strASNdata & Val(rsgetData.GetValue("SALES_QUANTITY")) & "," & rsgetData.GetValue("CONS_MEASURE_CODE").ToString.Trim & ","
                    strASNdata = strASNdata & rsgetData.GetValue("RATE").ToString.Trim & "," & rsgetData.GetValue("TOOL_COST").ToString.Trim & ","
                    strASNdata = strASNdata & rsgetData.GetValue("CHILDPARTCOST").ToString.Trim & "," & rsgetData.GetValue("BASIC_AMOUNT").ToString.Trim & ","
                    strASNdata = strASNdata & rsgetData.GetValue("ACCESSIBLE_AMOUNT").ToString.Trim & "," & rsgetData.GetValue("EXCISE_AMOUNT").ToString.Trim & ","
                    strASNdata = strASNdata & intmaxloop & "," & rsgetData.GetValue("VALUEONCALCULATEVAT").ToString.Trim & "," & rsgetData.GetValue("CUSTOMDUTY").ToString.Trim & ","
                    'strASNdata = strASNdata & rsgetData.GetValue("BASICTOOLCOST").ToString.Trim & ",0,0,0,0,"
                    'strASNdata = strASNdata & rsgetData.GetValue("BASICTOOLCOST").ToString.Trim & ",0,0,0,"
                    strASNdata = strASNdata & rsgetData.GetValue("BASICTOOLCOST").ToString.Trim & ",0,0,"
                    strASNdata = strASNdata & rsgetData.GetValue("TCSEXCLUDE_TOTALAMT").ToString & ","
                    strASNdata = strASNdata & rsgetData.GetValue("VALUEONCALCULATEADDVAT").ToString.Trim & ","
                    strASNdata = strASNdata & rsgetData.GetValue("TOTAL_AMOUNT").ToString.Trim & "," & rsgetData.GetValue("TOTALBASICAMOUNT").ToString.Trim & ","
                    strASNdata = strASNdata & rsgetData.GetValue("SALESTAXTYPE").ToString.Trim & "," & Val(rsgetData.GetValue("SALESTAX_PER")) & ","
                    strASNdata = strASNdata & rsgetData.GetValue("SALES_TAX_AMOUNT").ToString.Trim & ",LOC,0,0,EXC,"
                    strASNdata = strASNdata & rsgetData.GetValue("EXCISE_PER").ToString.Trim & "," & rsgetData.GetValue("TOTALEXCISEAMOUNT").ToString.Trim
                    '16 JUNE 2017
                    'strASNdata = strASNdata & ",CUD,0,0,CVD,0,0,ADD,0,0,OTH,0,0"
                    'strASNdata = strASNdata & ",CUD,0,0,CVD,0,0,ADD,0,0,OTH," & rsgetData.GetValue("ADDVAT_PER").ToString.Trim & "," & rsgetData.GetValue("ADDVAT_AMOUNT").ToString.Trim
                    strASNdata = strASNdata & ",CUD,0,0,CVD,0,0,ADD,0,0,OTH," & rsgetData.GetValue("TCSTAX_per").ToString.Trim & "," & rsgetData.GetValue("TCSTAXAMOUNT").ToString.Trim
                    'GST CHANGES
                    If gblnGSTUnit = True Then
                        strASNdata = strASNdata & "," & rsgetData.GetValue("CALCULATEON_CGSTAMT").ToString.Trim & "," & rsgetData.GetValue("CALCULATEON_CGSTREVSERALAMT").ToString.Trim & ","
                        strASNdata = strASNdata & rsgetData.GetValue("CALCULATEON_SGSTAMT").ToString.Trim & "," & rsgetData.GetValue("CALCULATEON_SGSTREVSERALAMT").ToString.Trim & ","
                        strASNdata = strASNdata & rsgetData.GetValue("CALCULATEON_IGSTAMT").ToString.Trim & "," & rsgetData.GetValue("CALCULATEON_IGSTREVSERALAMT").ToString.Trim & ","
                        strASNdata = strASNdata & rsgetData.GetValue("GSTBCD").ToString.Trim & ","
                        strASNdata = strASNdata & rsgetData.GetValue("GSTCOMPECC").ToString.Trim & "," & rsgetData.GetValue("GSTRCDCOMPCC").ToString.Trim & ","
                        strASNdata = strASNdata & rsgetData.GetValue("CGSTCODE").ToString.Trim & "," & rsgetData.GetValue("CGSTPREFIX").ToString.Trim & ","
                        strASNdata = strASNdata & rsgetData.GetValue("CGSTPERCENT").ToString.Trim & "," & rsgetData.GetValue("CGST_AMT").ToString.Trim & ","
                        strASNdata = strASNdata & rsgetData.GetValue("IGSTCODE").ToString.Trim & "," & rsgetData.GetValue("IGSTPREFIX").ToString.Trim & ","
                        strASNdata = strASNdata & rsgetData.GetValue("IGSTPERCENT").ToString.Trim & "," & rsgetData.GetValue("IGST_AMT").ToString.Trim & ","
                        strASNdata = strASNdata & rsgetData.GetValue("SGSTCODE").ToString.Trim & "," & rsgetData.GetValue("SGSTPREFIX").ToString.Trim & ","
                        strASNdata = strASNdata & rsgetData.GetValue("SGSTPERCENT").ToString.Trim & "," & rsgetData.GetValue("SGST_AMT").ToString.Trim & "," & rsgetData.GetValue("GSTCUD").ToString.Trim & ","
                        strASNdata = strASNdata & rsgetData.GetValue("IGSTRPREFIX").ToString.Trim & "," & rsgetData.GetValue("GSTCUD_PER").ToString.Trim & ","
                        strASNdata = strASNdata & rsgetData.GetValue("GSTCUD_AMT").ToString.Trim & "," & rsgetData.GetValue("GST_COMPECC").ToString.Trim & ","
                        strASNdata = strASNdata & rsgetData.GetValue("GST_COMPECC_PERFIX").ToString.Trim & "," & rsgetData.GetValue("SUR_PER").ToString.Trim & ","
                        strASNdata = strASNdata & rsgetData.GetValue("SUR_AMT").ToString.Trim
                    End If
                    If intloopcounter < intmaxloop Then
                        strASNdata = strASNdata & vbCrLf
                    End If

                    rsgetData.MoveNext()
                Next
                ASNPath = gstrUserMyDocPath
                ASNPathForEDI = Find_Value("SELECT ASNFILEPATH FROM CUSTOMER_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & straccountcode & "'")

                If Directory.Exists(ASNPath) = False Then
                    Directory.CreateDirectory(ASNPath)
                End If
                If Directory.Exists(ASNPathForEDI) = False Then
                    Directory.CreateDirectory(ASNPathForEDI)
                End If
                strASNFilepath = ASNPath & "\" & mInvNo.ToString() & ".txt"
                strASNFilepathforEDI = ASNPathForEDI & "\" & mInvNo.ToString() & ".txt"
                fs = File.Create(strASNFilepath)
                sw = New StreamWriter(fs)
                sw.WriteLine(strASNdata)
                sw.Close()
                fs.Close()
                If File.Exists(strASNFilepathforEDI) = False Then
                    File.Copy(strASNFilepath, strASNFilepathforEDI)
                End If
                rsgetData.ResultSetClose()
                rsgetData = Nothing
                ASNTEXTFILE_DETAILS = True
            End If
        Else
            MessageBox.Show("Unable To Generate ASN Text File. Invoice Can't Be Locked", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
            ASNTEXTFILE_DETAILS = False
        End If
        Exit Function
ErrHandler:
        ASNTEXTFILE_DETAILS = False
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
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

        If AllowBarCodePrinting(straccountcode) = True And blnlinelevelcustomer = False And ChkQrbarcodereprint.Checked = True Then
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
                                Msgbox("Problem While Generating Barcode Image.", vbInformation, ResolveResString(100))
                                rsGENERATEBARCODE.ResultSetClose()
                                rsGENERATEBARCODE = Nothing
                                Exit Sub
                            Else
                                '10812364
                                strBarcodeMsg_paratemeter = Mid(strBarcodeMsg, 3)
                                '10812364
                                If UPDATEQRBarCodeImage(Trim(Ctlinvoice.Text), Val(rsGENERATEBARCODE.GetValue("FromLineNo").ToString), Val(rsGENERATEBARCODE.GetValue("ToLineNo").ToString), strBarcodeMsg_paratemeter, intRow, intTotalNoofSlabs, "") = False Then
                                    Msgbox("Problem While saving Barcode Image.", vbInformation, ResolveResString(100))
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
                ElseIf AllowBarCodePrinting(straccountcode) = True And ChkQrbarcodereprint.Checked = True Then
                    rsGENERATEBARCODE = New ClsResultSetDB_Invoice
                    rsGENERATEBARCODE.GetResult("SELECT PRINT_METHOD FROM CUSTOMER_MST C WHERE C.UNIT_CODE='" & gstrUNITID & "' AND C.CUSTOMER_CODE='" & straccountcode & "'")
                    strPrintMethod = UCase(rsGENERATEBARCODE.GetValue("PRINT_METHOD").ToString)
                    rsGENERATEBARCODE.ResultSetClose()
                    rsGENERATEBARCODE = Nothing
                    If optInvYes(0).Checked = False Then 'only reprint 
                        If strPrintMethod = "TATA" Then
                            mP_Connection.Execute("Exec USP_INCR_TRUECOPY_DESCRIPTOR  '" & gstrUNITID & "', '" & Ctlinvoice.Text & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        End If
                    End If



                Else

                    strBarcodeMsg = ObjBarcodeHMI.GenerateBarCode(gstrUserMyDocPath, mInvNo, True, Trim(Ctlinvoice.Text), gstrCONNECTIONSTRING)
                    If VB.Left(strBarcodeMsg, 1) <> "Y" Then
                        Msgbox("Problem While Generating Barcode Image.", vbInformation, ResolveResString(100))
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
    Public Function SaveGateEntryBarcodeImage_singlelevelso_2DBARCODE(ByVal pstrInvNo As String, ByVal pstrPath As String, ByVal strbarcodestring As String) As Boolean
        On Error GoTo ErrHandler
        Dim stimage As ADODB.Stream
        Dim strQuery As String
        Dim Rs As ADODB.Recordset
        SaveGateEntryBarcodeImage_singlelevelso_2DBARCODE = True
        stimage = New ADODB.Stream
        stimage.Type = ADODB.StreamTypeEnum.adTypeBinary
        stimage.Open()
        pstrPath = pstrPath & ".JPEG"
        stimage.LoadFromFile(pstrPath)
        strQuery = "select  GE_Image  from ASN_Lebels_temp_2D where invoiceno='" & Trim(pstrInvNo) & "' and vendorcode = '" & gstrUNITID & "'"
        Rs = New ADODB.Recordset
        Rs.Open(strQuery, mP_Connection, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
        Rs.Fields("GE_Image").Value = stimage.Read
        Rs.Update()
        Rs.Close()
        Rs = Nothing
        Exit Function
ErrHandler:
        SaveGateEntryBarcodeImage_singlelevelso_2DBARCODE = False
        Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function IsGSTINSAME(ByVal strCustomerCode As String) As Boolean
        If SqlConnectionclass.ExecuteScalar("Select ISNULL(GSTIN_Id,'') GSTIN_Id From Customer_Mst Where UNIT_CODE='" & gstrUNITID & "' And Customer_Code='" & strCustomerCode & "'") = SqlConnectionclass.ExecuteScalar("Select ISNULL(GSTIN_ID,'') GSTIN_ID From Gen_UnitMaster Where Unt_CodeId='" & gstrUNITID & "'") Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Sub lblprintinBasecurrency_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblprintinBasecurrency.Click

    End Sub

    Private Sub chkPrintReprint_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkprintreprint.CheckedChanged
        If chkprintreprint.Checked = True Then
            Ctlinvoice.Text = ""
            _optInvYes_1.Text = "Reprint"
        Else
            Ctlinvoice.Text = ""
            _optInvYes_1.Text = "Print"
        End If
    End Sub
    Private Function GetExportOptions(ByVal strInvoiceNoForFileName As String, ByRef strCreatedPDFPath As String) As ExportOptions
        If (System.IO.Directory.Exists(My.Application.Info.DirectoryPath + "\InvoicePDF") = False) Then
            System.IO.Directory.CreateDirectory(My.Application.Info.DirectoryPath + "\InvoicePDF")
        End If
        Dim fileDestinationOptions As New DiskFileDestinationOptions
        Dim exportOptions As New ExportOptions()
        strCreatedPDFPath = My.Application.Info.DirectoryPath + "\InvoicePDF\" + strInvoiceNoForFileName + ".pdf"
        fileDestinationOptions.DiskFileName = strCreatedPDFPath 'eInvoicingFileName
        exportOptions.ExportDestinationOptions = fileDestinationOptions
        exportOptions.ExportDestinationType = ExportDestinationType.DiskFile
        exportOptions.ExportFormatType = ExportFormatType.PortableDocFormat
        Return exportOptions
    End Function
    Private Sub EXPORTINVOICETOPDF_ONPRINTREPRINT(ByVal strAccountCode As String, ByVal strInvoiceType As String, ByVal strInvoiceSubType As String, ByRef RPTDoc As ReportDocument, ByVal DTA4InvoicePrintingTag As DataTable)
        ''AMIT RANA 20 Jun 2019
        Try
            mblnAnnextureExport = False
            Dim frmpdfViewer As New frmeMProPDfViewer
            Dim strCreatedPDFPath As String = String.Empty
            Dim strRESULT As Collections.ArrayList
            Dim STRINVOICENO_PDF As String = String.Empty
            Dim intLoopCounter As Integer
            Dim intMaxLoop As Integer
            Dim Frm As Object = Nothing
            Dim RPTDocTemp As ReportDocument
            'Dim strCOPYNAME As String = String.Empty
            Dim strCOPYNAME(1) As String
            Dim strSQL_Annexture As String


            If optInvYes(0).Checked = True Then 'GENERATE THE INVOICE
                STRINVOICENO_PDF = mInvNo
            Else
                STRINVOICENO_PDF = Ctlinvoice.Text
            End If

            Dim OBJPdfConfig As Object = SqlConnectionclass.ExecuteScalar("SELECT COUNT(*) FROM INVOICE_PDF_CONFIG (NOLOCK) WHERE UNIT_CODE='" + gstrUNITID + "' AND CUSTOMER_CODE='" + strAccountCode.ToString() + "' AND INVOICE_TYPE='" & strInvoiceType & "' AND INVOICE_SUB_TYPE='" & strInvoiceSubType & "' And IS_ACTIVE=1")
            Dim OBJCommonDigital_EINVOICING_CONFIG As Object = SqlConnectionclass.ExecuteScalar("select Count(*) from CommonDigital_EINVOICING_CONFIG (NOLOCK) where UNIT_CODE='" + gstrUNITID + "' and CUSTOMER_CODE='" + strAccountCode.ToString() + "' AND INVOICE_TYPE='" & strInvoiceType & "' and SUB_TYPE='" & strInvoiceSubType & "' AND IS_ACTIVE=1")
            If (mblncustomerlevel_A4report_functionlity = True) Then
                If optInvYes(0).Checked = True Then

                    'If mblnA4reports_invoicewise = True And mblncustomerlevel_A4report_functionlity = True Then
                    If mblncustomerlevel_A4report_functionlity = True Then
                        '10825102 
                        Dim DataRowFiltered_O() As DataRow = DTA4InvoicePrintingTag.Select("ORIGINAL_REPRINT='O'")
                        intMaxLoop = DataRowFiltered_O.Length 'intMaxLoop = CInt(Find_Value("select isnull(MAX(SERIALNO),0) SERIALNO from A4CUSTOMER_INVOICEPRINTINGTAG  WHERE UNIT_CODE='" + gstrUNITID + "'AND CUSTOMER_CODE='" & strAccountCode & "' AND ORIGINAL_REPRINT='O'"))
                    Else
                        intMaxLoop = gstrIntNoCopies
                    End If
                Else
                    'If mblnA4reports_invoicewise = True And mblncustomerlevel_A4report_functionlity = True Then
                    If mblncustomerlevel_A4report_functionlity = True Then
                        '10825102 
                        Dim DataRowFiltered_R() As DataRow = DTA4InvoicePrintingTag.Select("ORIGINAL_REPRINT='R'")
                        intMaxLoop = DataRowFiltered_R.Length ' intMaxLoop = CInt(Find_Value("select isnull(MAX(SERIALNO),0) SERIALNO  from A4CUSTOMER_INVOICEPRINTINGTAG  WHERE UNIT_CODE='" + gstrUNITID + "'AND CUSTOMER_CODE='" & strAccountCode & "' AND ORIGINAL_REPRINT='R'"))
                    Else
                        If gstrIntNoCopies > 1 Then
                            intMaxLoop = gstrIntNoCopies - 1
                        Else
                            intMaxLoop = gstrIntNoCopies
                        End If
                    End If
                End If
            Else
                intMaxLoop = 2
                'If mblnISTrueSignRequired Then
                '    intMaxLoop = 2
                'End If
            End If


            For intLoopCounter = 1 To intMaxLoop
                If (mblncustomerlevel_A4report_functionlity = True) Then
                    If mblnEwaybill_Print = False Then
                        If optInvYes(0).Checked = True Then
                            'strCOPYNAME = Find_Value("Select TEXTHEADING FROM  A4CUSTOMER_INVOICEPRINTINGTAG WHERE UNIT_CODE='" + gstrUNITID + "' AND CUSTOMER_CODE='" & strAccountCode & "' AND ORIGINAL_REPRINT='O' AND SERIALNO=" & intLoopCounter)
                            Dim DataRowFiltered() As DataRow = DTA4InvoicePrintingTag.Select("ORIGINAL_REPRINT='O' and SERIALNO=" + intLoopCounter.ToString())
                            strCOPYNAME(0) = DataRowFiltered(0)("TEXTHEADING").ToString()
                            strCOPYNAME(1) = DataRowFiltered(0)("HARDCOPYPRINTREQUIRED").ToString()
                        Else
                            'strCOPYNAME = Find_Value("Select TEXTHEADING FROM  A4CUSTOMER_INVOICEPRINTINGTAG WHERE UNIT_CODE='" + gstrUNITID + "' AND CUSTOMER_CODE='" & strAccountCode & "' AND ORIGINAL_REPRINT='R' AND SERIALNO=" & intLoopCounter)
                            Dim DataRowFiltered() As DataRow = DTA4InvoicePrintingTag.Select("ORIGINAL_REPRINT='R' and SERIALNO=" + intLoopCounter.ToString())
                            strCOPYNAME(0) = DataRowFiltered(0)("TEXTHEADING").ToString()
                            strCOPYNAME(1) = DataRowFiltered(0)("HARDCOPYPRINTREQUIRED").ToString()
                        End If
                    Else
                        If chkprintreprint.Checked = True And optInvYes(1).Checked = True Then
                            'strCOPYNAME = Find_Value("Select TEXTHEADING FROM  A4CUSTOMER_INVOICEPRINTINGTAG WHERE UNIT_CODE='" + gstrUNITID + "' AND CUSTOMER_CODE='" & strAccountCode & "' AND ORIGINAL_REPRINT='R' AND SERIALNO=" & intLoopCounter)
                            Dim DataRowFiltered() As DataRow = DTA4InvoicePrintingTag.Select("ORIGINAL_REPRINT='R' and SERIALNO=" + intLoopCounter.ToString())
                            strCOPYNAME(0) = DataRowFiltered(0)("TEXTHEADING").ToString()
                            strCOPYNAME(1) = DataRowFiltered(0)("HARDCOPYPRINTREQUIRED").ToString()
                        ElseIf chkprintreprint.Checked = False And optInvYes(1).Checked = True Then
                            'strCOPYNAME = Find_Value("Select TEXTHEADING FROM  A4CUSTOMER_INVOICEPRINTINGTAG WHERE UNIT_CODE='" + gstrUNITID + "' AND CUSTOMER_CODE='" & strAccountCode & "' AND ORIGINAL_REPRINT='O' AND SERIALNO=" & intLoopCounter)
                            Dim DataRowFiltered() As DataRow = DTA4InvoicePrintingTag.Select("ORIGINAL_REPRINT='O' and SERIALNO=" + intLoopCounter.ToString())
                            strCOPYNAME(0) = DataRowFiltered(0)("TEXTHEADING").ToString()
                            strCOPYNAME(1) = DataRowFiltered(0)("HARDCOPYPRINTREQUIRED").ToString()
                        Else
                            'strCOPYNAME = Find_Value("Select TEXTHEADING FROM  A4CUSTOMER_INVOICEPRINTINGTAG WHERE UNIT_CODE='" + gstrUNITID + "' AND CUSTOMER_CODE='" & strAccountCode & "' AND ORIGINAL_REPRINT='R' AND SERIALNO=" & intLoopCounter)
                            Dim DataRowFiltered() As DataRow = DTA4InvoicePrintingTag.Select("ORIGINAL_REPRINT='R' and SERIALNO=" + intLoopCounter.ToString())
                            strCOPYNAME(0) = DataRowFiltered(0)("TEXTHEADING").ToString()
                            strCOPYNAME(1) = DataRowFiltered(0)("HARDCOPYPRINTREQUIRED").ToString()
                        End If
                    End If
                Else
                    If intLoopCounter = 1 Then
                        strCOPYNAME(0) = "ORIGINAL FOR BUYER"
                        strCOPYNAME(1) = "Y"
                    End If
                    If intLoopCounter = 2 Then
                        strCOPYNAME(0) = "DUPLICATE FOR TRANSPORTER"
                        strCOPYNAME(1) = "Y"
                    End If
                    If intLoopCounter = 3 Then
                        strCOPYNAME(0) = "TRIPLICATE FOR ASSESSEE"
                        strCOPYNAME(1) = "Y"
                    End If
                    If intLoopCounter = 4 Then
                        strCOPYNAME(0) = "EXTRA COPY"
                        strCOPYNAME(1) = "Y"
                    End If
                    If intLoopCounter = 5 Then
                        strCOPYNAME(0) = "EXTRA COPY"
                        strCOPYNAME(1) = "Y"
                    End If
                End If
                Select Case intLoopCounter
                    Case 1
                        RPTDoc.DataDefinition.FormulaFields("CopyName").Text = "'" & strCOPYNAME(0) & "'"

                        'Start code to merge original invoice and annexture
                        If mblnA4reports_invoicewise = True And mblncustomerlevel_Annexture_printing = True Then
                            If DataExist("SELECT TOP 1 1 FROM SALES_DTL WHERE UNIT_CODE = '" & gstrUNITID & "' and doc_no= " & Trim(STRINVOICENO_PDF) & " HAVING COUNT(*)> " & mintmaxnoofitems_Annexture) = True Then
                                mblnAnnextureExport = True
                            End If
                        End If

                        If mblnAnnextureExport Then
                            RPTDoc.Export(GetExportOptions(STRINVOICENO_PDF & "_First", String.Empty))
                            If UCase(Trim(gstrUNITID)) = "MST" Then
                                Frm = New eMProCrystalReportViewer
                            Else
                                Frm = New eMProCrystalReportViewer_Inv
                            End If
                            Dim strSQL As String
                            Dim strpdsno As String
                            If mblncustomerspecificreport Then
                                RPTDoc.DataDefinition.FormulaFields("TOYOTABARCODEPRINT").Text = False
                                If DataExist("SELECT TOP 1 1 FROM SALES_DTL  WHERE UNIT_CODE = '" & gstrUNITID & "' and doc_no= " & Trim(STRINVOICENO_PDF) & " HAVING COUNT(*)> " & mintmaxnoofitems_barcodeToyota) = True Then
                                    RPTDoc.DataDefinition.FormulaFields("TOYOTABARCODEPRINT").Text = True
                                End If
                            End If
                            RPTDoc.DataDefinition.FormulaFields("ANNEXTURE").Text = True
                            RPTDocTemp = Frm.GetReportDocument()
                            RPTDocTemp.Load(My.Application.Info.DirectoryPath & "\Reports\Annextureprinting_A4reports.rpt")
                            RPTDocTemp.DataDefinition.FormulaFields("Invoiceno").Text = "'" & STRINVOICENO_PDF & "'"
                            strpdsno = Find_Value("SELECT LORRYNO_DATE FROM SALESCHALLAN_DTL WHERE  UNIT_CODE = '" & gstrUNITID & "' AND  DOC_NO=" & Trim(STRINVOICENO_PDF))
                            RPTDocTemp.DataDefinition.FormulaFields("pdsno").Text = "'" & strpdsno & "'"
                            strSQL = "{SalesChallan_Dtl.Location_Code}='" & Trim(txtUnitCode.Text) & "' and {SalesChallan_Dtl.Doc_No} =" & Trim(STRINVOICENO_PDF) & " and {SalesChallan_Dtl.UNIT_CODE} ='" & gstrUNITID & "' "
                            RPTDocTemp.DataDefinition.RecordSelectionFormula = strSQL
                            Frm.SetReportDocument()
                            RPTDocTemp.Export(GetExportOptions(STRINVOICENO_PDF & "_Annexture", String.Empty))

                            Dim invoiceFilePath As String = My.Application.Info.DirectoryPath + "\InvoicePDF\"
                            Dim fileNames(2) As String
                            fileNames(0) = invoiceFilePath & STRINVOICENO_PDF & "_SignedInvoice.pdf"
                            fileNames(1) = invoiceFilePath & STRINVOICENO_PDF & "_First.pdf"
                            fileNames(2) = invoiceFilePath & STRINVOICENO_PDF & "_Annexture.pdf"

                            System.IO.File.Copy(fileNames(1), invoiceFilePath & STRINVOICENO_PDF & "_SignFirstCopy.pdf", True)

                            MergingInvoiceWithAnnexture(STRINVOICENO_PDF, strCreatedPDFPath)

                            Dim strCheckDataExistsSALESCHALLAN_SIGNED_PDFS As String = String.Empty
                            strCheckDataExistsSALESCHALLAN_SIGNED_PDFS = Find_Value("SELECT TOP 1 ISNULL(DATALENGTH(AnnextureInvoiceFirstPage_Binary), -1) FROM SALESCHALLAN_SIGNED_PDFS WHERE UNIT_CODE='" + gstrUNITID + "' AND ACCOUNT_CODE='" + strAccountCode.ToString() + "' AND DOC_NO=" + STRINVOICENO_PDF + " And Upper(Doc_Type_Text)='" + strCOPYNAME(0) + "' ")
                            Dim signedDocFILE As Byte()

                            If strCheckDataExistsSALESCHALLAN_SIGNED_PDFS = "" Or strCheckDataExistsSALESCHALLAN_SIGNED_PDFS = "-1" Then
                                strRESULT = SavePDFInvoicesInDB(strAccountCode.ToString(), STRINVOICENO_PDF, strCOPYNAME(0), gstrUNITID, strCreatedPDFPath, mblnAnnextureExport, OBJPdfConfig, 1)
                                If strRESULT.Item(0).ToString().Trim().ToUpper() = "SUCCESS" Then
                                    File.WriteAllBytes(fileNames(0), strRESULT.Item(1))
                                    strSQL = "UPDATE SALESCHALLAN_DTL SET TRUECOPY_DESCRIPTOR=TRUECOPY_DESCRIPTOR+1    WHERE DOC_NO=" + STRINVOICENO_PDF + " AND UNIT_CODE='" + gstrUNITID + "' "
                                    SqlConnectionclass.ExecuteNonQuery(strSQL)
                                    Dim strTrueCopyDescriptor As Integer
                                    strTrueCopyDescriptor = Find_Value("select TrueCopy_Descriptor from SalesChallan_Dtl where Unit_COde='" & gstrUNITID & "' and Doc_no='" + STRINVOICENO_PDF + "' ")
                                    Dim signedBase64 As String
                                    Dim DocFILE As Byte()

                                    Dim objApi As clsTrueCopyAPI.TrueCopyAPI = New clsTrueCopyAPI.TrueCopyAPI()
                                    DocFILE = GetFileBytes(invoiceFilePath & STRINVOICENO_PDF & "_SignFirstCopy.pdf")
                                    Dim filedata As String = objApi.GetBase64StringFormByte(DocFILE)
                                    signedBase64 = objApi.GetSignedDocumentBase64(mblnAPIUrl, filedata, mblnPFX_ID, mblnPFX_Pass, mblnAPI_Key, STRINVOICENO_PDF + ".pdf", "1[" + doubleCOORDINATE_LLX_copy.ToString + ":" + doubleCOORDINATE_LLY_copy.ToString + "]", strSignAuthorityName_copy & vbCrLf & "Reason: " & strReason_copy & vbCrLf & "Location: " & strLocation_copy, gstrUNITID + STRINVOICENO_PDF + strCOPYNAME(0) + strTrueCopyDescriptor.ToString())
                                    signedDocFILE = Convert.FromBase64String(signedBase64)
                                    File.WriteAllBytes(invoiceFilePath & STRINVOICENO_PDF & "_SignFirstCopy.pdf", signedDocFILE)

                                    Dim oSqlCmd As New SqlCommand()
                                    strSQL = "Update SALESCHALLAN_SIGNED_PDFS SET   AnnextureInvoiceFirstPage_Binary=@docFILE_SIGNED    WHERE UNIT_CODE='" + gstrUNITID + "' AND ACCOUNT_CODE='" + strAccountCode.ToString() + "' AND DOC_NO=" + STRINVOICENO_PDF + " And Upper(Doc_Type_Text)='" + strCOPYNAME(0) + "' "
                                    oSqlCmd.Parameters.Add("@docFILE_SIGNED", SqlDbType.VarBinary).Value = signedDocFILE
                                    If strSQL <> "" Then
                                        oSqlCmd.CommandText = strSQL
                                        SqlConnectionclass.ExecuteNonQuery(oSqlCmd)
                                    End If
                                Else
                                    If mblnISTrueSignRequired And Not mblnISCrystalReportRequired Then
                                        Exit Sub
                                    End If
                                End If
                            Else
                                Dim imageData As Byte() = DirectCast(SqlConnectionclass.ExecuteScalar("select AnnextureInvoiceFirstPage_Binary from SALESCHALLAN_SIGNED_PDFS  WHERE UNIT_CODE='" + gstrUNITID + "' AND ACCOUNT_CODE='" + strAccountCode.ToString() + "' AND DOC_NO=" + STRINVOICENO_PDF + " And Upper(Doc_Type_Text)='" + strCOPYNAME(0) + "'"), Byte())
                                File.WriteAllBytes(invoiceFilePath & STRINVOICENO_PDF & "_SignFirstCopy.pdf", imageData)
                                signedDocFILE = DirectCast(SqlConnectionclass.ExecuteScalar("select PDF_Binary from SALESCHALLAN_SIGNED_PDFS  WHERE UNIT_CODE='" + gstrUNITID + "' AND ACCOUNT_CODE='" + strAccountCode.ToString() + "' AND DOC_NO=" + STRINVOICENO_PDF + " And Upper(Doc_Type_Text)='" + strCOPYNAME(0) + "'"), Byte())
                            End If
                            If mblnISTrueSignRequired And Not mblnISCrystalReportRequired Then
                                If File.Exists(invoiceFilePath & STRINVOICENO_PDF & "_SignFirstCopy.pdf") Then
                                    If mCheckValARana = "PRINTTOPRINTER" And strCOPYNAME(1) = "Y" Then
                                        Dim MyProcess As New Process
                                        Try
                                            MyProcess.StartInfo.FileName = """" + mPdfReaderPath + """" ' + " ""C:\Users\amitrana\AppData\Local\Temp\7510093120231003191510.pdf""" ' + strFilePath
                                            MyProcess.StartInfo.Arguments = String.Format("{0}", " /t   " + """" + invoiceFilePath & STRINVOICENO_PDF & "_SignFirstCopy.pdf" + """")
                                            MyProcess.StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                                            MyProcess.Start()
                                            pdfPrintProcID = MyProcess.Id

                                            RPTDocTemp.PrintToPrinter(1, False, 0, 0)
                                        Catch Ex As Exception
                                            Msgbox(Ex.Message)
                                        Finally
                                        End Try
                                    End If
                                    DownloadPDFs(STRINVOICENO_PDF, strCOPYNAME(0), "N", signedDocFILE, frmpdfViewer, RPTDoc)
                                End If
                            End If
                            Try

                                If File.Exists(invoiceFilePath & STRINVOICENO_PDF & ".pdf") Then
                                    File.Delete(invoiceFilePath & STRINVOICENO_PDF & ".pdf")
                                End If
                                If File.Exists(invoiceFilePath & STRINVOICENO_PDF & "_SignedInvoice.pdf") Then
                                    File.Delete(invoiceFilePath & STRINVOICENO_PDF & "_SignedInvoice.pdf")
                                End If
                                If File.Exists(invoiceFilePath & STRINVOICENO_PDF & "_SignFirstCopy.pdf") Then
                                    File.Delete(invoiceFilePath & STRINVOICENO_PDF & "_SignFirstCopy.pdf")
                                End If

                            Catch ex As Exception

                            End Try
                        Else
                            RPTDoc.Export(GetExportOptions(STRINVOICENO_PDF, strCreatedPDFPath))
                            strRESULT = SavePDFInvoicesInDB(strAccountCode.ToString(), STRINVOICENO_PDF, strCOPYNAME(0), gstrUNITID, strCreatedPDFPath, mblnAnnextureExport, OBJPdfConfig, OBJCommonDigital_EINVOICING_CONFIG)
                            If strRESULT.Item(0).ToString().Trim().ToUpper() = "SUCCESS" Then
                                DownloadPDFs(STRINVOICENO_PDF, strCOPYNAME(0), strCOPYNAME(1), strRESULT.Item(1), frmpdfViewer, RPTDoc)
                            Else
                                If mblnISTrueSignRequired And Not mblnISCrystalReportRequired Then
                                    Exit Sub
                                End If
                            End If
                        End If
                        'End code to merge original invoice and annexture
                    Case 2
                        RPTDoc.DataDefinition.FormulaFields("CopyName").Text = "'" & strCOPYNAME(0) & "'"
                        'Start code to merge duplicate invoice and annexture
                        If mblnAnnextureExport Then
                            RPTDoc.Export(GetExportOptions(STRINVOICENO_PDF & "_First", String.Empty))
                            If UCase(Trim(gstrUNITID)) = "MST" Then
                                Frm = New eMProCrystalReportViewer
                            Else
                                Frm = New eMProCrystalReportViewer_Inv
                            End If
                            Dim strSQL As String
                            Dim strpdsno As String
                            If mblncustomerspecificreport Then
                                RPTDoc.DataDefinition.FormulaFields("TOYOTABARCODEPRINT").Text = False
                                If DataExist("SELECT TOP 1 1 FROM SALES_DTL  WHERE UNIT_CODE = '" & gstrUNITID & "' and doc_no= " & Trim(STRINVOICENO_PDF) & " HAVING COUNT(*)> " & mintmaxnoofitems_barcodeToyota) = True Then
                                    RPTDoc.DataDefinition.FormulaFields("TOYOTABARCODEPRINT").Text = True
                                End If
                            End If
                            RPTDoc.DataDefinition.FormulaFields("ANNEXTURE").Text = True
                            RPTDocTemp = Frm.GetReportDocument()
                            RPTDocTemp.Load(My.Application.Info.DirectoryPath & "\Reports\Annextureprinting_A4reports.rpt")
                            RPTDocTemp.DataDefinition.FormulaFields("Invoiceno").Text = "'" & STRINVOICENO_PDF & "'"
                            strpdsno = Find_Value("SELECT LORRYNO_DATE FROM SALESCHALLAN_DTL WHERE  UNIT_CODE = '" & gstrUNITID & "' AND  DOC_NO=" & Trim(STRINVOICENO_PDF))
                            RPTDocTemp.DataDefinition.FormulaFields("pdsno").Text = "'" & strpdsno & "'"
                            strSQL = "{SalesChallan_Dtl.Location_Code}='" & Trim(txtUnitCode.Text) & "' and {SalesChallan_Dtl.Doc_No} =" & Trim(STRINVOICENO_PDF) & " and {SalesChallan_Dtl.UNIT_CODE} ='" & gstrUNITID & "' "
                            RPTDocTemp.DataDefinition.RecordSelectionFormula = strSQL
                            Frm.SetReportDocument()
                            RPTDocTemp.Export(GetExportOptions(STRINVOICENO_PDF & "_Annexture", String.Empty))

                            Dim invoiceFilePath As String = My.Application.Info.DirectoryPath + "\InvoicePDF\"
                            Dim fileNames(2) As String
                            fileNames(0) = invoiceFilePath & STRINVOICENO_PDF & "_SignedInvoice.pdf"
                            fileNames(1) = invoiceFilePath & STRINVOICENO_PDF & "_First.pdf"
                            fileNames(2) = invoiceFilePath & STRINVOICENO_PDF & "_Annexture.pdf"

                            System.IO.File.Copy(fileNames(1), invoiceFilePath & STRINVOICENO_PDF & "_SignFirstCopy.pdf", True)

                            MergingInvoiceWithAnnexture(STRINVOICENO_PDF, strCreatedPDFPath)

                            Dim strCheckDataExistsSALESCHALLAN_SIGNED_PDFS As String = String.Empty
                            strCheckDataExistsSALESCHALLAN_SIGNED_PDFS = Find_Value("SELECT TOP 1 ISNULL(DATALENGTH(AnnextureInvoiceFirstPage_Binary), -1) FROM SALESCHALLAN_SIGNED_PDFS WHERE UNIT_CODE='" + gstrUNITID + "' AND ACCOUNT_CODE='" + strAccountCode.ToString() + "' AND DOC_NO=" + STRINVOICENO_PDF + " And Upper(Doc_Type_Text)='" + strCOPYNAME(0) + "' ")
                            Dim signedDocFILE As Byte()

                            If strCheckDataExistsSALESCHALLAN_SIGNED_PDFS = "" Or strCheckDataExistsSALESCHALLAN_SIGNED_PDFS = "-1" Then
                                strRESULT = SavePDFInvoicesInDB(strAccountCode.ToString(), STRINVOICENO_PDF, strCOPYNAME(0), gstrUNITID, strCreatedPDFPath, mblnAnnextureExport, OBJPdfConfig, OBJCommonDigital_EINVOICING_CONFIG)
                                If strRESULT.Item(0).ToString().Trim().ToUpper() = "SUCCESS" Then
                                    File.WriteAllBytes(fileNames(0), strRESULT.Item(1))
                                    strSQL = "UPDATE SALESCHALLAN_DTL SET TRUECOPY_DESCRIPTOR=TRUECOPY_DESCRIPTOR+1    WHERE DOC_NO=" + STRINVOICENO_PDF + " AND UNIT_CODE='" + gstrUNITID + "' "
                                    SqlConnectionclass.ExecuteNonQuery(strSQL)
                                    Dim strTrueCopyDescriptor As Integer
                                    strTrueCopyDescriptor = Find_Value("select TrueCopy_Descriptor from SalesChallan_Dtl where Unit_COde='" & gstrUNITID & "' and Doc_no='" + STRINVOICENO_PDF + "' ")
                                    Dim signedBase64 As String
                                    Dim DocFILE As Byte()

                                    Dim objApi As clsTrueCopyAPI.TrueCopyAPI = New clsTrueCopyAPI.TrueCopyAPI()
                                    DocFILE = GetFileBytes(invoiceFilePath & STRINVOICENO_PDF & "_SignFirstCopy.pdf")
                                    Dim filedata As String = objApi.GetBase64StringFormByte(DocFILE)
                                    signedBase64 = objApi.GetSignedDocumentBase64(mblnAPIUrl, filedata, mblnPFX_ID, mblnPFX_Pass, mblnAPI_Key, STRINVOICENO_PDF + ".pdf", "1[" + doubleCOORDINATE_LLX_copy.ToString + ":" + doubleCOORDINATE_LLY_copy.ToString + "]", strSignAuthorityName_copy & vbCrLf & "Reason: " & strReason_copy & vbCrLf & "Location: " & strLocation_copy, gstrUNITID + STRINVOICENO_PDF + strCOPYNAME(0) + strTrueCopyDescriptor.ToString())
                                    signedDocFILE = Convert.FromBase64String(signedBase64)
                                    File.WriteAllBytes(invoiceFilePath & STRINVOICENO_PDF & "_SignFirstCopy.pdf", signedDocFILE)

                                    Dim oSqlCmd As New SqlCommand()
                                    strSQL = "Update SALESCHALLAN_SIGNED_PDFS SET   AnnextureInvoiceFirstPage_Binary=@docFILE_SIGNED    WHERE UNIT_CODE='" + gstrUNITID + "' AND ACCOUNT_CODE='" + strAccountCode.ToString() + "' AND DOC_NO=" + STRINVOICENO_PDF + " And Upper(Doc_Type_Text)='" + strCOPYNAME(0) + "' "
                                    oSqlCmd.Parameters.Add("@docFILE_SIGNED", SqlDbType.VarBinary).Value = signedDocFILE
                                    If strSQL <> "" Then
                                        oSqlCmd.CommandText = strSQL
                                        SqlConnectionclass.ExecuteNonQuery(oSqlCmd)
                                    End If
                                Else
                                    If mblnISTrueSignRequired And Not mblnISCrystalReportRequired Then
                                        Exit Sub
                                    End If
                                End If
                            Else
                                Dim imageData As Byte() = DirectCast(SqlConnectionclass.ExecuteScalar("select AnnextureInvoiceFirstPage_Binary from SALESCHALLAN_SIGNED_PDFS  WHERE UNIT_CODE='" + gstrUNITID + "' AND ACCOUNT_CODE='" + strAccountCode.ToString() + "' AND DOC_NO=" + STRINVOICENO_PDF + " And Upper(Doc_Type_Text)='" + strCOPYNAME(0) + "'"), Byte())
                                File.WriteAllBytes(invoiceFilePath & STRINVOICENO_PDF & "_SignFirstCopy.pdf", imageData)
                                signedDocFILE = DirectCast(SqlConnectionclass.ExecuteScalar("select PDF_Binary from SALESCHALLAN_SIGNED_PDFS  WHERE UNIT_CODE='" + gstrUNITID + "' AND ACCOUNT_CODE='" + strAccountCode.ToString() + "' AND DOC_NO=" + STRINVOICENO_PDF + " And Upper(Doc_Type_Text)='" + strCOPYNAME(0) + "'"), Byte())
                            End If
                            If mblnISTrueSignRequired And Not mblnISCrystalReportRequired Then

                                If File.Exists(invoiceFilePath & STRINVOICENO_PDF & "_SignFirstCopy.pdf") Then
                                    If mCheckValARana = "PRINTTOPRINTER" And strCOPYNAME(1) = "Y" Then
                                        Dim MyProcess As New Process
                                        Try
                                            MyProcess.StartInfo.FileName = """" + mPdfReaderPath + """" ' + " ""C:\Users\amitrana\AppData\Local\Temp\7510093120231003191510.pdf""" ' + strFilePath
                                            MyProcess.StartInfo.Arguments = String.Format("{0}", " /t   " + """" + invoiceFilePath & STRINVOICENO_PDF & "_SignFirstCopy.pdf" + """")
                                            MyProcess.StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                                            MyProcess.Start()
                                            pdfPrintProcID = MyProcess.Id

                                            RPTDocTemp.PrintToPrinter(1, False, 0, 0)
                                        Catch Ex As Exception
                                            Msgbox(Ex.Message)
                                        Finally
                                        End Try
                                    End If
                                    DownloadPDFs(STRINVOICENO_PDF, strCOPYNAME(0), "N", signedDocFILE, frmpdfViewer, RPTDoc)
                                End If
                            End If
                            Try
                                If File.Exists(invoiceFilePath & STRINVOICENO_PDF & "_SignFirstCopy.pdf") Then
                                    File.Delete(invoiceFilePath & STRINVOICENO_PDF & "_SignFirstCopy.pdf")
                                End If
                                If File.Exists(invoiceFilePath & STRINVOICENO_PDF & ".pdf") Then
                                    File.Delete(invoiceFilePath & STRINVOICENO_PDF & ".pdf")
                                End If
                                If File.Exists(invoiceFilePath & STRINVOICENO_PDF & "_SignedInvoice.pdf") Then
                                    File.Delete(invoiceFilePath & STRINVOICENO_PDF & "_SignedInvoice.pdf")
                                End If

                            Catch ex As Exception

                            End Try
                        Else
                            RPTDoc.Export(GetExportOptions(STRINVOICENO_PDF, strCreatedPDFPath))
                            strRESULT = SavePDFInvoicesInDB(strAccountCode.ToString(), STRINVOICENO_PDF, strCOPYNAME(0), gstrUNITID, strCreatedPDFPath, mblnAnnextureExport, OBJPdfConfig, OBJCommonDigital_EINVOICING_CONFIG)
                            If strRESULT.Item(0).ToString().Trim().ToUpper() = "SUCCESS" Then
                                DownloadPDFs(STRINVOICENO_PDF, strCOPYNAME(0), strCOPYNAME(1), strRESULT.Item(1), frmpdfViewer, RPTDoc)
                            Else
                                If mblnISTrueSignRequired And Not mblnISCrystalReportRequired Then
                                    Exit Sub
                                End If
                            End If
                        End If
                        'End code to merge duplicate invoice and annexture
                    Case 3
                        RPTDoc.DataDefinition.FormulaFields("CopyName").Text = "'" & strCOPYNAME(0) & "'"
                        'Start code to merge duplicate invoice and annexture
                        If mblnAnnextureExport Then
                            RPTDoc.Export(GetExportOptions(STRINVOICENO_PDF & "_First", String.Empty))
                            If UCase(Trim(gstrUNITID)) = "MST" Then
                                Frm = New eMProCrystalReportViewer
                            Else
                                Frm = New eMProCrystalReportViewer_Inv
                            End If
                            Dim strSQL As String
                            Dim strpdsno As String
                            If mblncustomerspecificreport Then
                                RPTDoc.DataDefinition.FormulaFields("TOYOTABARCODEPRINT").Text = False
                                If DataExist("SELECT TOP 1 1 FROM SALES_DTL  WHERE UNIT_CODE = '" & gstrUNITID & "' and doc_no= " & Trim(STRINVOICENO_PDF) & " HAVING COUNT(*)> " & mintmaxnoofitems_barcodeToyota) = True Then
                                    RPTDoc.DataDefinition.FormulaFields("TOYOTABARCODEPRINT").Text = True
                                End If
                            End If
                            RPTDoc.DataDefinition.FormulaFields("ANNEXTURE").Text = True
                            RPTDocTemp = Frm.GetReportDocument()
                            RPTDocTemp.Load(My.Application.Info.DirectoryPath & "\Reports\Annextureprinting_A4reports.rpt")
                            RPTDocTemp.DataDefinition.FormulaFields("Invoiceno").Text = "'" & STRINVOICENO_PDF & "'"
                            strpdsno = Find_Value("SELECT LORRYNO_DATE FROM SALESCHALLAN_DTL WHERE  UNIT_CODE = '" & gstrUNITID & "' AND  DOC_NO=" & Trim(STRINVOICENO_PDF))
                            RPTDocTemp.DataDefinition.FormulaFields("pdsno").Text = "'" & strpdsno & "'"
                            strSQL = "{SalesChallan_Dtl.Location_Code}='" & Trim(txtUnitCode.Text) & "' and {SalesChallan_Dtl.Doc_No} =" & Trim(STRINVOICENO_PDF) & " and {SalesChallan_Dtl.UNIT_CODE} ='" & gstrUNITID & "' "
                            RPTDocTemp.DataDefinition.RecordSelectionFormula = strSQL
                            Frm.SetReportDocument()
                            RPTDocTemp.Export(GetExportOptions(STRINVOICENO_PDF & "_Annexture", String.Empty))

                            Dim invoiceFilePath As String = My.Application.Info.DirectoryPath + "\InvoicePDF\"
                            Dim fileNames(2) As String
                            fileNames(0) = invoiceFilePath & STRINVOICENO_PDF & "_SignedInvoice.pdf"
                            fileNames(1) = invoiceFilePath & STRINVOICENO_PDF & "_First.pdf"
                            fileNames(2) = invoiceFilePath & STRINVOICENO_PDF & "_Annexture.pdf"

                            System.IO.File.Copy(fileNames(1), invoiceFilePath & STRINVOICENO_PDF & "_SignFirstCopy.pdf", True)

                            MergingInvoiceWithAnnexture(STRINVOICENO_PDF, strCreatedPDFPath)

                            Dim strCheckDataExistsSALESCHALLAN_SIGNED_PDFS As String = String.Empty
                            strCheckDataExistsSALESCHALLAN_SIGNED_PDFS = Find_Value("SELECT TOP 1 ISNULL(DATALENGTH(AnnextureInvoiceFirstPage_Binary), -1) FROM SALESCHALLAN_SIGNED_PDFS WHERE UNIT_CODE='" + gstrUNITID + "' AND ACCOUNT_CODE='" + strAccountCode.ToString() + "' AND DOC_NO=" + STRINVOICENO_PDF + " And Upper(Doc_Type_Text)='" + strCOPYNAME(0) + "' ")
                            Dim signedDocFILE As Byte()

                            If strCheckDataExistsSALESCHALLAN_SIGNED_PDFS = "" Or strCheckDataExistsSALESCHALLAN_SIGNED_PDFS = "-1" Then
                                strRESULT = SavePDFInvoicesInDB(strAccountCode.ToString(), STRINVOICENO_PDF, strCOPYNAME(0), gstrUNITID, strCreatedPDFPath, mblnAnnextureExport, OBJPdfConfig, OBJCommonDigital_EINVOICING_CONFIG)
                                If strRESULT.Item(0).ToString().Trim().ToUpper() = "SUCCESS" Then
                                    File.WriteAllBytes(fileNames(0), strRESULT.Item(1))
                                    strSQL = "UPDATE SALESCHALLAN_DTL SET TRUECOPY_DESCRIPTOR=TRUECOPY_DESCRIPTOR+1    WHERE DOC_NO=" + STRINVOICENO_PDF + " AND UNIT_CODE='" + gstrUNITID + "' "
                                    SqlConnectionclass.ExecuteNonQuery(strSQL)
                                    Dim strTrueCopyDescriptor As Integer
                                    strTrueCopyDescriptor = Find_Value("select TrueCopy_Descriptor from SalesChallan_Dtl where Unit_COde='" & gstrUNITID & "' and Doc_no='" + STRINVOICENO_PDF + "' ")
                                    Dim signedBase64 As String
                                    Dim DocFILE As Byte()

                                    Dim objApi As clsTrueCopyAPI.TrueCopyAPI = New clsTrueCopyAPI.TrueCopyAPI()
                                    DocFILE = GetFileBytes(invoiceFilePath & STRINVOICENO_PDF & "_SignFirstCopy.pdf")
                                    Dim filedata As String = objApi.GetBase64StringFormByte(DocFILE)
                                    signedBase64 = objApi.GetSignedDocumentBase64(mblnAPIUrl, filedata, mblnPFX_ID, mblnPFX_Pass, mblnAPI_Key, STRINVOICENO_PDF + ".pdf", "1[" + doubleCOORDINATE_LLX_copy.ToString + ":" + doubleCOORDINATE_LLY_copy.ToString + "]", strSignAuthorityName_copy & vbCrLf & "Reason: " & strReason_copy & vbCrLf & "Location: " & strLocation_copy, gstrUNITID + STRINVOICENO_PDF + strCOPYNAME(0) + strTrueCopyDescriptor.ToString())
                                    signedDocFILE = Convert.FromBase64String(signedBase64)
                                    File.WriteAllBytes(invoiceFilePath & STRINVOICENO_PDF & "_SignFirstCopy.pdf", signedDocFILE)

                                    Dim oSqlCmd As New SqlCommand()
                                    strSQL = "Update SALESCHALLAN_SIGNED_PDFS SET   AnnextureInvoiceFirstPage_Binary=@docFILE_SIGNED    WHERE UNIT_CODE='" + gstrUNITID + "' AND ACCOUNT_CODE='" + strAccountCode.ToString() + "' AND DOC_NO=" + STRINVOICENO_PDF + " And Upper(Doc_Type_Text)='" + strCOPYNAME(0) + "' "
                                    oSqlCmd.Parameters.Add("@docFILE_SIGNED", SqlDbType.VarBinary).Value = signedDocFILE
                                    If strSQL <> "" Then
                                        oSqlCmd.CommandText = strSQL
                                        SqlConnectionclass.ExecuteNonQuery(oSqlCmd)
                                    End If
                                Else
                                    If mblnISTrueSignRequired And Not mblnISCrystalReportRequired Then
                                        Exit Sub
                                    End If
                                End If
                            Else

                                Dim imageData As Byte() = DirectCast(SqlConnectionclass.ExecuteScalar("select AnnextureInvoiceFirstPage_Binary from SALESCHALLAN_SIGNED_PDFS  WHERE UNIT_CODE='" + gstrUNITID + "' AND ACCOUNT_CODE='" + strAccountCode.ToString() + "' AND DOC_NO=" + STRINVOICENO_PDF + " And Upper(Doc_Type_Text)='" + strCOPYNAME(0) + "'"), Byte())
                                File.WriteAllBytes(invoiceFilePath & STRINVOICENO_PDF & "_SignFirstCopy.pdf", imageData)
                                signedDocFILE = DirectCast(SqlConnectionclass.ExecuteScalar("select PDF_Binary from SALESCHALLAN_SIGNED_PDFS  WHERE UNIT_CODE='" + gstrUNITID + "' AND ACCOUNT_CODE='" + strAccountCode.ToString() + "' AND DOC_NO=" + STRINVOICENO_PDF + " And Upper(Doc_Type_Text)='" + strCOPYNAME(0) + "'"), Byte())
                            End If
                            If mblnISTrueSignRequired And Not mblnISCrystalReportRequired Then

                                If File.Exists(invoiceFilePath & STRINVOICENO_PDF & "_SignFirstCopy.pdf") Then
                                    If mCheckValARana = "PRINTTOPRINTER" And strCOPYNAME(1) = "Y" Then
                                        Dim MyProcess As New Process
                                        Try
                                            MyProcess.StartInfo.FileName = """" + mPdfReaderPath + """" ' + " ""C:\Users\amitrana\AppData\Local\Temp\7510093120231003191510.pdf""" ' + strFilePath
                                            MyProcess.StartInfo.Arguments = String.Format("{0}", " /t   " + """" + invoiceFilePath & STRINVOICENO_PDF & "_SignFirstCopy.pdf" + """")
                                            MyProcess.StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                                            MyProcess.Start()
                                            pdfPrintProcID = MyProcess.Id

                                            RPTDocTemp.PrintToPrinter(1, False, 0, 0)
                                        Catch Ex As Exception
                                            Msgbox(Ex.Message)
                                        Finally
                                        End Try
                                    End If
                                    DownloadPDFs(STRINVOICENO_PDF, strCOPYNAME(0), "N", signedDocFILE, frmpdfViewer, RPTDoc)
                                End If
                            End If
                            Try
                                If File.Exists(invoiceFilePath & STRINVOICENO_PDF & "_SignFirstCopy.pdf") Then
                                    File.Delete(invoiceFilePath & STRINVOICENO_PDF & "_SignFirstCopy.pdf")
                                End If
                                If File.Exists(invoiceFilePath & STRINVOICENO_PDF & ".pdf") Then
                                    File.Delete(invoiceFilePath & STRINVOICENO_PDF & ".pdf")
                                End If
                                If File.Exists(invoiceFilePath & STRINVOICENO_PDF & "_SignedInvoice.pdf") Then
                                    File.Delete(invoiceFilePath & STRINVOICENO_PDF & "_SignedInvoice.pdf")
                                End If

                            Catch ex As Exception

                            End Try
                        Else
                            RPTDoc.Export(GetExportOptions(STRINVOICENO_PDF, strCreatedPDFPath))
                            strRESULT = SavePDFInvoicesInDB(strAccountCode.ToString(), STRINVOICENO_PDF, strCOPYNAME(0), gstrUNITID, strCreatedPDFPath, mblnAnnextureExport, OBJPdfConfig, OBJCommonDigital_EINVOICING_CONFIG)
                            If strRESULT.Item(0).ToString().Trim().ToUpper() = "SUCCESS" Then
                                DownloadPDFs(STRINVOICENO_PDF, strCOPYNAME(0), strCOPYNAME(1), strRESULT.Item(1), frmpdfViewer, RPTDoc)
                            Else
                                If mblnISTrueSignRequired And Not mblnISCrystalReportRequired Then
                                    Exit Sub
                                End If
                            End If
                        End If
                        'End code to merge duplicate invoice and annexture
                    Case 4
                        RPTDoc.DataDefinition.FormulaFields("CopyName").Text = "'" & strCOPYNAME(0) & "'"
                        'Start code to merge duplicate invoice and annexture
                        If mblnAnnextureExport Then
                            RPTDoc.Export(GetExportOptions(STRINVOICENO_PDF & "_First", String.Empty))
                            If UCase(Trim(gstrUNITID)) = "MST" Then
                                Frm = New eMProCrystalReportViewer
                            Else
                                Frm = New eMProCrystalReportViewer_Inv
                            End If
                            Dim strSQL As String
                            Dim strpdsno As String
                            If mblncustomerspecificreport Then
                                RPTDoc.DataDefinition.FormulaFields("TOYOTABARCODEPRINT").Text = False
                                If DataExist("SELECT TOP 1 1 FROM SALES_DTL  WHERE UNIT_CODE = '" & gstrUNITID & "' and doc_no= " & Trim(STRINVOICENO_PDF) & " HAVING COUNT(*)> " & mintmaxnoofitems_barcodeToyota) = True Then
                                    RPTDoc.DataDefinition.FormulaFields("TOYOTABARCODEPRINT").Text = True
                                End If
                            End If
                            RPTDoc.DataDefinition.FormulaFields("ANNEXTURE").Text = True
                            RPTDocTemp = Frm.GetReportDocument()
                            RPTDocTemp.Load(My.Application.Info.DirectoryPath & "\Reports\Annextureprinting_A4reports.rpt")
                            RPTDocTemp.DataDefinition.FormulaFields("Invoiceno").Text = "'" & STRINVOICENO_PDF & "'"
                            strpdsno = Find_Value("SELECT LORRYNO_DATE FROM SALESCHALLAN_DTL WHERE  UNIT_CODE = '" & gstrUNITID & "' AND  DOC_NO=" & Trim(STRINVOICENO_PDF))
                            RPTDocTemp.DataDefinition.FormulaFields("pdsno").Text = "'" & strpdsno & "'"
                            strSQL = "{SalesChallan_Dtl.Location_Code}='" & Trim(txtUnitCode.Text) & "' and {SalesChallan_Dtl.Doc_No} =" & Trim(STRINVOICENO_PDF) & " and {SalesChallan_Dtl.UNIT_CODE} ='" & gstrUNITID & "' "
                            RPTDocTemp.DataDefinition.RecordSelectionFormula = strSQL
                            Frm.SetReportDocument()
                            RPTDocTemp.Export(GetExportOptions(STRINVOICENO_PDF & "_Annexture", String.Empty))

                            Dim invoiceFilePath As String = My.Application.Info.DirectoryPath + "\InvoicePDF\"
                            Dim fileNames(2) As String
                            fileNames(0) = invoiceFilePath & STRINVOICENO_PDF & "_SignedInvoice.pdf"
                            fileNames(1) = invoiceFilePath & STRINVOICENO_PDF & "_First.pdf"
                            fileNames(2) = invoiceFilePath & STRINVOICENO_PDF & "_Annexture.pdf"

                            System.IO.File.Copy(fileNames(1), invoiceFilePath & STRINVOICENO_PDF & "_SignFirstCopy.pdf", True)

                            MergingInvoiceWithAnnexture(STRINVOICENO_PDF, strCreatedPDFPath)

                            Dim strCheckDataExistsSALESCHALLAN_SIGNED_PDFS As String = String.Empty
                            strCheckDataExistsSALESCHALLAN_SIGNED_PDFS = Find_Value("SELECT TOP 1 ISNULL(DATALENGTH(AnnextureInvoiceFirstPage_Binary), -1) FROM SALESCHALLAN_SIGNED_PDFS WHERE UNIT_CODE='" + gstrUNITID + "' AND ACCOUNT_CODE='" + strAccountCode.ToString() + "' AND DOC_NO=" + STRINVOICENO_PDF + " And Upper(Doc_Type_Text)='" + strCOPYNAME(0) + "' ")
                            Dim signedDocFILE As Byte()

                            If strCheckDataExistsSALESCHALLAN_SIGNED_PDFS = "" Or strCheckDataExistsSALESCHALLAN_SIGNED_PDFS = "-1" Then
                                strRESULT = SavePDFInvoicesInDB(strAccountCode.ToString(), STRINVOICENO_PDF, strCOPYNAME(0), gstrUNITID, strCreatedPDFPath, mblnAnnextureExport, OBJPdfConfig, OBJCommonDigital_EINVOICING_CONFIG)
                                If strRESULT.Item(0).ToString().Trim().ToUpper() = "SUCCESS" Then
                                    File.WriteAllBytes(fileNames(0), strRESULT.Item(1))
                                    strSQL = "UPDATE SALESCHALLAN_DTL SET TRUECOPY_DESCRIPTOR=TRUECOPY_DESCRIPTOR+1    WHERE DOC_NO=" + STRINVOICENO_PDF + " AND UNIT_CODE='" + gstrUNITID + "' "
                                    SqlConnectionclass.ExecuteNonQuery(strSQL)
                                    Dim strTrueCopyDescriptor As Integer
                                    strTrueCopyDescriptor = Find_Value("select TrueCopy_Descriptor from SalesChallan_Dtl where Unit_COde='" & gstrUNITID & "' and Doc_no='" + STRINVOICENO_PDF + "' ")
                                    Dim signedBase64 As String
                                    Dim DocFILE As Byte()

                                    Dim objApi As clsTrueCopyAPI.TrueCopyAPI = New clsTrueCopyAPI.TrueCopyAPI()
                                    DocFILE = GetFileBytes(invoiceFilePath & STRINVOICENO_PDF & "_SignFirstCopy.pdf")
                                    Dim filedata As String = objApi.GetBase64StringFormByte(DocFILE)
                                    signedBase64 = objApi.GetSignedDocumentBase64(mblnAPIUrl, filedata, mblnPFX_ID, mblnPFX_Pass, mblnAPI_Key, STRINVOICENO_PDF + ".pdf", "1[" + doubleCOORDINATE_LLX_copy.ToString + ":" + doubleCOORDINATE_LLY_copy.ToString + "]", strSignAuthorityName_copy & vbCrLf & "Reason: " & strReason_copy & vbCrLf & "Location: " & strLocation_copy, gstrUNITID + STRINVOICENO_PDF + strCOPYNAME(0) + strTrueCopyDescriptor.ToString())
                                    signedDocFILE = Convert.FromBase64String(signedBase64)
                                    File.WriteAllBytes(invoiceFilePath & STRINVOICENO_PDF & "_SignFirstCopy.pdf", signedDocFILE)

                                    Dim oSqlCmd As New SqlCommand()
                                    strSQL = "Update SALESCHALLAN_SIGNED_PDFS SET   AnnextureInvoiceFirstPage_Binary=@docFILE_SIGNED    WHERE UNIT_CODE='" + gstrUNITID + "' AND ACCOUNT_CODE='" + strAccountCode.ToString() + "' AND DOC_NO=" + STRINVOICENO_PDF + " And Upper(Doc_Type_Text)='" + strCOPYNAME(0) + "' "
                                    oSqlCmd.Parameters.Add("@docFILE_SIGNED", SqlDbType.VarBinary).Value = signedDocFILE
                                    If strSQL <> "" Then
                                        oSqlCmd.CommandText = strSQL
                                        SqlConnectionclass.ExecuteNonQuery(oSqlCmd)
                                    End If
                                Else
                                    If mblnISTrueSignRequired And Not mblnISCrystalReportRequired Then
                                        Exit Sub
                                    End If
                                End If
                            Else
                                Dim imageData As Byte() = DirectCast(SqlConnectionclass.ExecuteScalar("select AnnextureInvoiceFirstPage_Binary from SALESCHALLAN_SIGNED_PDFS  WHERE UNIT_CODE='" + gstrUNITID + "' AND ACCOUNT_CODE='" + strAccountCode.ToString() + "' AND DOC_NO=" + STRINVOICENO_PDF + " And Upper(Doc_Type_Text)='" + strCOPYNAME(0) + "'"), Byte())
                                File.WriteAllBytes(invoiceFilePath & STRINVOICENO_PDF & "_SignFirstCopy.pdf", imageData)
                                signedDocFILE = DirectCast(SqlConnectionclass.ExecuteScalar("select PDF_Binary from SALESCHALLAN_SIGNED_PDFS  WHERE UNIT_CODE='" + gstrUNITID + "' AND ACCOUNT_CODE='" + strAccountCode.ToString() + "' AND DOC_NO=" + STRINVOICENO_PDF + " And Upper(Doc_Type_Text)='" + strCOPYNAME(0) + "'"), Byte())
                            End If
                            If mblnISTrueSignRequired And Not mblnISCrystalReportRequired Then

                                If File.Exists(invoiceFilePath & STRINVOICENO_PDF & "_SignFirstCopy.pdf") Then
                                    If mCheckValARana = "PRINTTOPRINTER" And strCOPYNAME(1) = "Y" Then
                                        Dim MyProcess As New Process
                                        Try
                                            MyProcess.StartInfo.FileName = """" + mPdfReaderPath + """" ' + " ""C:\Users\amitrana\AppData\Local\Temp\7510093120231003191510.pdf""" ' + strFilePath
                                            MyProcess.StartInfo.Arguments = String.Format("{0}", " /t   " + """" + invoiceFilePath & STRINVOICENO_PDF & "_SignFirstCopy.pdf" + """")
                                            MyProcess.StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                                            MyProcess.Start()
                                            pdfPrintProcID = MyProcess.Id

                                            RPTDocTemp.PrintToPrinter(1, False, 0, 0)
                                        Catch Ex As Exception
                                            Msgbox(Ex.Message)
                                        Finally
                                        End Try
                                    End If
                                    DownloadPDFs(STRINVOICENO_PDF, strCOPYNAME(0), "N", signedDocFILE, frmpdfViewer, RPTDoc)
                                End If

                            End If
                            Try
                                If File.Exists(invoiceFilePath & STRINVOICENO_PDF & "_SignFirstCopy.pdf") Then
                                    File.Delete(invoiceFilePath & STRINVOICENO_PDF & "_SignFirstCopy.pdf")
                                End If
                                If File.Exists(invoiceFilePath & STRINVOICENO_PDF & ".pdf") Then
                                    File.Delete(invoiceFilePath & STRINVOICENO_PDF & ".pdf")
                                End If
                                If File.Exists(invoiceFilePath & STRINVOICENO_PDF & "_SignedInvoice.pdf") Then
                                    File.Delete(invoiceFilePath & STRINVOICENO_PDF & "_SignedInvoice.pdf")
                                End If

                            Catch ex As Exception

                            End Try
                        Else
                            RPTDoc.Export(GetExportOptions(STRINVOICENO_PDF, strCreatedPDFPath))
                            strRESULT = SavePDFInvoicesInDB(strAccountCode.ToString(), STRINVOICENO_PDF, strCOPYNAME(0), gstrUNITID, strCreatedPDFPath, mblnAnnextureExport, OBJPdfConfig, OBJCommonDigital_EINVOICING_CONFIG)
                            If strRESULT.Item(0).ToString().Trim().ToUpper() = "SUCCESS" Then
                                DownloadPDFs(STRINVOICENO_PDF, strCOPYNAME(0), strCOPYNAME(1), strRESULT.Item(1), frmpdfViewer, RPTDoc)
                            Else
                                If mblnISTrueSignRequired And Not mblnISCrystalReportRequired Then
                                    Exit Sub
                                End If
                            End If
                        End If
                        'End code to merge duplicate invoice and annexture
                    Case 5
                        RPTDoc.DataDefinition.FormulaFields("CopyName").Text = "'" & strCOPYNAME(0) & "'"
                        'Start code to merge duplicate invoice and annexture
                        If mblnAnnextureExport Then
                            RPTDoc.Export(GetExportOptions(STRINVOICENO_PDF & "_First", String.Empty))
                            If UCase(Trim(gstrUNITID)) = "MST" Then
                                Frm = New eMProCrystalReportViewer
                            Else
                                Frm = New eMProCrystalReportViewer_Inv
                            End If
                            Dim strSQL As String
                            Dim strpdsno As String
                            If mblncustomerspecificreport Then
                                RPTDoc.DataDefinition.FormulaFields("TOYOTABARCODEPRINT").Text = False
                                If DataExist("SELECT TOP 1 1 FROM SALES_DTL  WHERE UNIT_CODE = '" & gstrUNITID & "' and doc_no= " & Trim(STRINVOICENO_PDF) & " HAVING COUNT(*)> " & mintmaxnoofitems_barcodeToyota) = True Then
                                    RPTDoc.DataDefinition.FormulaFields("TOYOTABARCODEPRINT").Text = True
                                End If
                            End If
                            RPTDoc.DataDefinition.FormulaFields("ANNEXTURE").Text = True
                            RPTDocTemp = Frm.GetReportDocument()
                            RPTDocTemp.Load(My.Application.Info.DirectoryPath & "\Reports\Annextureprinting_A4reports.rpt")
                            RPTDocTemp.DataDefinition.FormulaFields("Invoiceno").Text = "'" & STRINVOICENO_PDF & "'"
                            strpdsno = Find_Value("SELECT LORRYNO_DATE FROM SALESCHALLAN_DTL WHERE  UNIT_CODE = '" & gstrUNITID & "' AND  DOC_NO=" & Trim(STRINVOICENO_PDF))
                            RPTDocTemp.DataDefinition.FormulaFields("pdsno").Text = "'" & strpdsno & "'"
                            strSQL = "{SalesChallan_Dtl.Location_Code}='" & Trim(txtUnitCode.Text) & "' and {SalesChallan_Dtl.Doc_No} =" & Trim(STRINVOICENO_PDF) & " and {SalesChallan_Dtl.UNIT_CODE} ='" & gstrUNITID & "' "
                            RPTDocTemp.DataDefinition.RecordSelectionFormula = strSQL
                            Frm.SetReportDocument()
                            RPTDocTemp.Export(GetExportOptions(STRINVOICENO_PDF & "_Annexture", String.Empty))

                            Dim invoiceFilePath As String = My.Application.Info.DirectoryPath + "\InvoicePDF\"
                            Dim fileNames(2) As String
                            fileNames(0) = invoiceFilePath & STRINVOICENO_PDF & "_SignedInvoice.pdf"
                            fileNames(1) = invoiceFilePath & STRINVOICENO_PDF & "_First.pdf"
                            fileNames(2) = invoiceFilePath & STRINVOICENO_PDF & "_Annexture.pdf"

                            System.IO.File.Copy(fileNames(1), invoiceFilePath & STRINVOICENO_PDF & "_SignFirstCopy.pdf", True)

                            MergingInvoiceWithAnnexture(STRINVOICENO_PDF, strCreatedPDFPath)

                            Dim strCheckDataExistsSALESCHALLAN_SIGNED_PDFS As String = String.Empty
                            strCheckDataExistsSALESCHALLAN_SIGNED_PDFS = Find_Value("SELECT TOP 1 ISNULL(DATALENGTH(AnnextureInvoiceFirstPage_Binary), -1) FROM SALESCHALLAN_SIGNED_PDFS WHERE UNIT_CODE='" + gstrUNITID + "' AND ACCOUNT_CODE='" + strAccountCode.ToString() + "' AND DOC_NO=" + STRINVOICENO_PDF + " And Upper(Doc_Type_Text)='" + strCOPYNAME(0) + "' ")
                            Dim signedDocFILE As Byte()

                            If strCheckDataExistsSALESCHALLAN_SIGNED_PDFS = "" Or strCheckDataExistsSALESCHALLAN_SIGNED_PDFS = "-1" Then
                                strRESULT = SavePDFInvoicesInDB(strAccountCode.ToString(), STRINVOICENO_PDF, strCOPYNAME(0), gstrUNITID, strCreatedPDFPath, mblnAnnextureExport, OBJPdfConfig, OBJCommonDigital_EINVOICING_CONFIG)
                                If strRESULT.Item(0).ToString().Trim().ToUpper() = "SUCCESS" Then
                                    File.WriteAllBytes(fileNames(0), strRESULT.Item(1))
                                    strSQL = "UPDATE SALESCHALLAN_DTL SET TRUECOPY_DESCRIPTOR=TRUECOPY_DESCRIPTOR+1    WHERE DOC_NO=" + STRINVOICENO_PDF + " AND UNIT_CODE='" + gstrUNITID + "' "
                                    SqlConnectionclass.ExecuteNonQuery(strSQL)
                                    Dim strTrueCopyDescriptor As Integer
                                    strTrueCopyDescriptor = Find_Value("select TrueCopy_Descriptor from SalesChallan_Dtl where Unit_COde='" & gstrUNITID & "' and Doc_no='" + STRINVOICENO_PDF + "' ")
                                    Dim signedBase64 As String
                                    Dim DocFILE As Byte()

                                    Dim objApi As clsTrueCopyAPI.TrueCopyAPI = New clsTrueCopyAPI.TrueCopyAPI()
                                    DocFILE = GetFileBytes(invoiceFilePath & STRINVOICENO_PDF & "_SignFirstCopy.pdf")
                                    Dim filedata As String = objApi.GetBase64StringFormByte(DocFILE)
                                    signedBase64 = objApi.GetSignedDocumentBase64(mblnAPIUrl, filedata, mblnPFX_ID, mblnPFX_Pass, mblnAPI_Key, STRINVOICENO_PDF + ".pdf", "1[" + doubleCOORDINATE_LLX_copy.ToString + ":" + doubleCOORDINATE_LLY_copy.ToString + "]", strSignAuthorityName_copy & vbCrLf & "Reason: " & strReason_copy & vbCrLf & "Location: " & strLocation_copy, gstrUNITID + STRINVOICENO_PDF + strCOPYNAME(0) + strTrueCopyDescriptor.ToString())
                                    signedDocFILE = Convert.FromBase64String(signedBase64)
                                    File.WriteAllBytes(invoiceFilePath & STRINVOICENO_PDF & "_SignFirstCopy.pdf", signedDocFILE)

                                    Dim oSqlCmd As New SqlCommand()
                                    strSQL = "Update SALESCHALLAN_SIGNED_PDFS SET   AnnextureInvoiceFirstPage_Binary=@docFILE_SIGNED    WHERE UNIT_CODE='" + gstrUNITID + "' AND ACCOUNT_CODE='" + strAccountCode.ToString() + "' AND DOC_NO=" + STRINVOICENO_PDF + " And Upper(Doc_Type_Text)='" + strCOPYNAME(0) + "' "
                                    oSqlCmd.Parameters.Add("@docFILE_SIGNED", SqlDbType.VarBinary).Value = signedDocFILE
                                    If strSQL <> "" Then
                                        oSqlCmd.CommandText = strSQL
                                        SqlConnectionclass.ExecuteNonQuery(oSqlCmd)
                                    End If
                                Else
                                    If mblnISTrueSignRequired And Not mblnISCrystalReportRequired Then
                                        Exit Sub
                                    End If
                                End If
                            Else
                                Dim imageData As Byte() = DirectCast(SqlConnectionclass.ExecuteScalar("select AnnextureInvoiceFirstPage_Binary from SALESCHALLAN_SIGNED_PDFS  WHERE UNIT_CODE='" + gstrUNITID + "' AND ACCOUNT_CODE='" + strAccountCode.ToString() + "' AND DOC_NO=" + STRINVOICENO_PDF + " And Upper(Doc_Type_Text)='" + strCOPYNAME(0) + "'"), Byte())
                                File.WriteAllBytes(invoiceFilePath & STRINVOICENO_PDF & "_SignFirstCopy.pdf", imageData)
                                signedDocFILE = DirectCast(SqlConnectionclass.ExecuteScalar("select PDF_Binary from SALESCHALLAN_SIGNED_PDFS  WHERE UNIT_CODE='" + gstrUNITID + "' AND ACCOUNT_CODE='" + strAccountCode.ToString() + "' AND DOC_NO=" + STRINVOICENO_PDF + " And Upper(Doc_Type_Text)='" + strCOPYNAME(0) + "'"), Byte())
                            End If
                            If mblnISTrueSignRequired And Not mblnISCrystalReportRequired Then

                                If File.Exists(invoiceFilePath & STRINVOICENO_PDF & "_SignFirstCopy.pdf") Then
                                    If mCheckValARana = "PRINTTOPRINTER" And strCOPYNAME(1) = "Y" Then
                                        Dim MyProcess As New Process
                                        Try
                                            MyProcess.StartInfo.FileName = """" + mPdfReaderPath + """" ' + " ""C:\Users\amitrana\AppData\Local\Temp\7510093120231003191510.pdf""" ' + strFilePath
                                            MyProcess.StartInfo.Arguments = String.Format("{0}", " /t   " + """" + invoiceFilePath & STRINVOICENO_PDF & "_SignFirstCopy.pdf" + """")
                                            MyProcess.StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                                            MyProcess.Start()
                                            pdfPrintProcID = MyProcess.Id

                                            RPTDocTemp.PrintToPrinter(1, False, 0, 0)
                                        Catch Ex As Exception
                                            Msgbox(Ex.Message)
                                        Finally
                                        End Try
                                    End If
                                    DownloadPDFs(STRINVOICENO_PDF, strCOPYNAME(0), "N", signedDocFILE, frmpdfViewer, RPTDoc)
                                End If
                            End If
                            Try
                                If File.Exists(invoiceFilePath & STRINVOICENO_PDF & "_SignFirstCopy.pdf") Then
                                    File.Delete(invoiceFilePath & STRINVOICENO_PDF & "_SignFirstCopy.pdf")
                                End If
                                If File.Exists(invoiceFilePath & STRINVOICENO_PDF & ".pdf") Then
                                    File.Delete(invoiceFilePath & STRINVOICENO_PDF & ".pdf")
                                End If
                                If File.Exists(invoiceFilePath & STRINVOICENO_PDF & "_SignedInvoice.pdf") Then
                                    File.Delete(invoiceFilePath & STRINVOICENO_PDF & "_SignedInvoice.pdf")
                                End If

                            Catch ex As Exception

                            End Try
                        Else
                            RPTDoc.Export(GetExportOptions(STRINVOICENO_PDF, strCreatedPDFPath))
                            strRESULT = SavePDFInvoicesInDB(strAccountCode.ToString(), STRINVOICENO_PDF, strCOPYNAME(0), gstrUNITID, strCreatedPDFPath, mblnAnnextureExport, OBJPdfConfig, OBJCommonDigital_EINVOICING_CONFIG)
                            If strRESULT.Item(0).ToString().Trim().ToUpper() = "SUCCESS" Then
                                DownloadPDFs(STRINVOICENO_PDF, strCOPYNAME(0), strCOPYNAME(1), strRESULT.Item(1), frmpdfViewer, RPTDoc)
                            Else
                                If mblnISTrueSignRequired And Not mblnISCrystalReportRequired Then
                                    Exit Sub
                                End If
                            End If
                        End If
                        'End code to merge duplicate invoice and annexture

                End Select
            Next
            If Not mblnISCrystalReportRequired And frmpdfViewer.PDF_Path_1 <> "" Then
                frmpdfViewer.ShowDialog()
            End If



            ''AMIT RANA 20 Jun 2019
        Catch ex As Exception
            Msgbox(ex.Message.ToString())
        End Try

    End Sub

    Private Sub DownloadPDFs(ByVal strDocNo As String, ByVal strDocType As String, ByVal strPrintRequired As String, ByVal bytebuffer As Byte(), ByVal frmpdfViewer As frmeMProPDfViewer, ByRef RPTDoc As ReportDocument)
        Try
            If bytebuffer Is Nothing Then
                Exit Sub
            End If
            Dim strFilePath, strQry As String
            Dim buffer As Byte()
            Dim strReportName() As String

            'strQry = "Select top 1 PDF_BINARY from SALESCHALLAN_SIGNED_PDFS  where UNIT_CODE = '" + gstrUNITID + "' and DOC_NO = '" + strDocNo + "' and doc_type_text = '" + cmbInvoiceType.SelectedValue + "'"
            'buffer = SqlConnectionclass.ExecuteScalar(strQry)
            Dim strdatetime As String
            strdatetime = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")
            strdatetime = strdatetime.Replace("/", "").Replace(":", "").Replace(" ", "")
            buffer = bytebuffer
            If strDocType = "ORIGINAL FOR BUYER" Then
                strFilePath = System.IO.Path.GetTempPath() + strDocNo + strdatetime + "_org.pdf"
            ElseIf strDocType = "DUPLICATE FOR TRANSPORTER" Then
                strFilePath = System.IO.Path.GetTempPath() + strDocNo + strdatetime + "_dup.pdf"
            ElseIf strDocType = "TRIPLICATE FOR ASSESSEE" Then
                strFilePath = System.IO.Path.GetTempPath() + strDocNo + strdatetime + "_tri.pdf"
            ElseIf strDocType = "EXTRA COPY" Then
                strFilePath = System.IO.Path.GetTempPath() + strDocNo + strdatetime + "_ext.pdf"
            Else
                strFilePath = System.IO.Path.GetTempPath() + strDocNo + strdatetime + ".pdf"
            End If

            Try
                For Each file As IO.FileInfo In New IO.DirectoryInfo(System.IO.Path.GetTempPath()).GetFiles("*.pdf*")
                    If (Now - file.CreationTime).Days > 2 Then
                        Try
                            file.Delete()
                        Catch ex As Exception           ' log exception or ignore '     
                        End Try
                    End If
                Next
            Catch ex As Exception

            End Try


            'Try
            '    If System.IO.File.Exists(strFilePath) Then
            '        System.IO.File.Delete(strFilePath)
            '    End If
            'Catch ex As Exception

            'End Try
            System.IO.File.WriteAllBytes(strFilePath, buffer)
            If mCheckValARana = "PRINTTOPRINTER" And strPrintRequired = "Y" Then
                'Dim MyProcess As New Process
                'MyProcess.StartInfo.FileName = """C:\Program Files (x86)\Foxit Software\Foxit Reader\FoxitReader.exe """ + " ""C:\Users\amitrana\AppData\Local\Temp\7510093120231003191510.pdf""" ' + strFilePath
                'MyProcess.StartInfo.Verb = "Print"
                'MyProcess.StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                'MyProcess.Start()
                Dim MyProcess As New Process
                Try
                    MyProcess.StartInfo.FileName = """" + mPdfReaderPath + """" ' + " ""C:\Users\amitrana\AppData\Local\Temp\7510093120231003191510.pdf""" ' + strFilePath
                    MyProcess.StartInfo.Arguments = String.Format("{0}", " /t   " + """" + strFilePath + """")
                    'MyProcess.StartInfo.Verb = "Print"
                    MyProcess.StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                    MyProcess.Start()
                    pdfPrintProcID = MyProcess.Id
                Catch Ex As Exception
                    Msgbox(Ex.Message)
                Finally
                End Try
            ElseIf mCheckValARana = "PRINTTOWINDOW" Then
                'Dim act As Action(Of String) = New Action(Of String)(AddressOf openPDFFile)
                'act.BeginInvoke(strFilePath, Nothing, Nothing)
                strReportName = RPTDoc.FilePath.ToUpper.Split("\")
                frmpdfViewer.Text = "       REPORT NAME : " + strReportName(strReportName.Length - 1)

                If frmpdfViewer.PDF_Path_1 = "" Then
                    frmpdfViewer.PDF_Path_1 = strFilePath
                    frmpdfViewer.PDF_Path_1_TabName = strDocType
                    Exit Sub
                End If
                If frmpdfViewer.PDF_Path_2 = "" Then
                    frmpdfViewer.PDF_Path_2 = strFilePath
                    frmpdfViewer.PDF_Path_2_TabName = strDocType
                    Exit Sub
                End If
                If frmpdfViewer.PDF_Path_3 = "" Then
                    frmpdfViewer.PDF_Path_3 = strFilePath
                    frmpdfViewer.PDF_Path_3_TabName = strDocType
                    Exit Sub
                End If
                If frmpdfViewer.PDF_Path_4 = "" Then
                    frmpdfViewer.PDF_Path_4 = strFilePath
                    frmpdfViewer.PDF_Path_4_TabName = strDocType
                    Exit Sub
                End If
                If frmpdfViewer.PDF_Path_5 = "" Then
                    frmpdfViewer.PDF_Path_5 = strFilePath
                    frmpdfViewer.PDF_Path_5_TabName = strDocType
                    Exit Sub
                End If
            End If

            'Try
            '    System.IO.File.Delete(System.IO.Path.GetTempPath())
            'Catch ex As Exception

            'End Try
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub
    Private Shared Sub openPDFFile(ByVal strFilePath As String)
        Try
            Using p As New System.Diagnostics.Process
                p.StartInfo = New System.Diagnostics.ProcessStartInfo(strFilePath)
                p.Start()
                p.WaitForExit()
                Try
                    System.IO.File.Delete(strFilePath)
                Catch ex As Exception
                    '   MessageBox.Show("error in openPDFFile function", ResolveResString(100), MessageBoxButtons.OK)
                End Try
            End Using
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK)
        End Try
    End Sub
    Dim doubleCOORDINATE_LLX_copy As Double = 0
    Dim doubleCOORDINATE_LLY_copy As Double = 0
    Dim strSignAuthorityName_copy As String
    Dim strReason_copy As String
    Dim strLocation_copy As String
    Public Function SavePDFInvoicesInDB(ByVal strCustomer As String, ByVal strInvoiceNo As String, ByVal strDocTypeText As String, ByVal strUnit As String, ByVal strInvoicePDFPath As String, ByVal boolAnnextureRequired As Boolean, ByVal OBJPdfConfig As Object, ByVal OBJCommonDigital_EINVOICING_CONFIG As Object) As Collections.ArrayList
        ''Dim RESULT As String = "FAIL"
        Dim RESULT As Collections.ArrayList
        RESULT = New Collections.ArrayList
        Try
            Dim strSQL As String = String.Empty
            Dim rows As Integer = 0
            Dim oSqlCmd As New SqlCommand()
            RESULT.Add("FAIL")
            SavePDFInvoicesInDB = RESULT
            Dim DocFILE As Byte()
            If (strCustomer = "" Or strInvoiceNo = "" Or strUnit = "" Or strInvoicePDFPath = "") Then
                ''RESULT = "Customer/InvoiceNo/Unit/InvoicePath Can't be Blank!"
                RESULT.Clear()
                RESULT.Add("Customer/InvoiceNo/Unit/InvoicePath Can't be Blank!")
                SavePDFInvoicesInDB = RESULT
                Exit Function
            End If
            If (System.IO.File.Exists(strInvoicePDFPath)) Then

            Else
                RESULT.Clear()
                RESULT.Add("Invoice File Path Not Exist!")
                SavePDFInvoicesInDB = RESULT
                Exit Function
            End If

            '' PRAVEEN DIGITAL SIGN --GET DATA FOR SIGNED PDF
            ''102853899 
            Dim strTrueCopyDescriptor As String = String.Empty
            Dim strSignAuthorityName As String = String.Empty
            Dim strReason As String = String.Empty
            Dim strLocation As String = String.Empty
            Dim doubleCOORDINATE_LLX As Double = 0
            Dim doubleCOORDINATE_LLY As Double = 0
            Dim boolisSaveSignedInvoicePDF As Boolean = False

            Dim dtCommonDigital_EINVOICING_CONFIG As DataTable

            strTrueCopyDescriptor = Find_Value("select TrueCopy_Descriptor from SalesChallan_Dtl where Unit_COde='" & gstrUNITID & "' and Doc_no='" + strInvoiceNo.ToString() + "' ")

            If (Val(OBJCommonDigital_EINVOICING_CONFIG.ToString()) > 0 And mblnISTrueSignRequired And Not mblnISCrystalReportRequired) Then
                strSQL = "select ISNULL(SignAuthorityName,'') SignAuthorityName,isnull(Reason,'') Reason,isnull(Location,'') Location,isnull(IsSaveSignedInvoicePDF,0) IsSaveSignedInvoicePDF from CommonDigital_EINVOICING_CONFIG (NOLOCK) where UNIT_CODE='" & gstrUNITID & "' and CUSTOMER_CODE='" + strCustomer + "' AND INVOICE_TYPE='" + lbldescription.Text + "' and SUB_TYPE='" + lblcategory.Text + "' AND IS_ACTIVE=1"
                dtCommonDigital_EINVOICING_CONFIG = SqlConnectionclass.GetDataTable(strSQL)
                If dtCommonDigital_EINVOICING_CONFIG.Rows.Count > 0 Then
                    strSignAuthorityName = dtCommonDigital_EINVOICING_CONFIG.Rows(0)("SignAuthorityName")
                    strReason = dtCommonDigital_EINVOICING_CONFIG.Rows(0)("Reason")
                    strLocation = dtCommonDigital_EINVOICING_CONFIG.Rows(0)("Location")
                    boolisSaveSignedInvoicePDF = CBool(dtCommonDigital_EINVOICING_CONFIG.Rows(0)("IsSaveSignedInvoicePDF"))
                End If
                If strSignAuthorityName <> "" Then
                    doubleCOORDINATE_LLX = Convert.ToDecimal(Find_Value("select ISNULL(COORDINATE_LLX,0) from CommonDigital_EINVOICING_CONFIG (NOLOCK) where UNIT_CODE='" & gstrUNITID & "' and CUSTOMER_CODE='" + strCustomer + "' AND INVOICE_TYPE='" + lbldescription.Text + "' and SUB_TYPE='" + lblcategory.Text + "' AND IS_ACTIVE=1"))
                    doubleCOORDINATE_LLY = Convert.ToDecimal(Find_Value("select ISNULL(COORDINATE_LLY,0) from CommonDigital_EINVOICING_CONFIG (NOLOCK) where UNIT_CODE='" & gstrUNITID & "' and CUSTOMER_CODE='" + strCustomer + "' AND INVOICE_TYPE='" + lbldescription.Text + "' and SUB_TYPE='" + lblcategory.Text + "' AND IS_ACTIVE=1"))
                End If
            End If

            '' PRAVEEN DIGITAL SIGN --GET DATA FOR SIGNED PDF

            ''DELETE THE FILE BEFORE NEW INSERT
            If strInvoicePDFPath <> "" Then
                If (strDocTypeText.Trim.ToString().ToUpper() = "SUPPLEMENTARY") Then
                    strSQL = "DELETE FROM SALESCHALLAN_SUPP_PDFS WHERE UNIT_CODE='" + strUnit + "' AND ACCOUNT_CODE='" + strCustomer + "' AND DOC_NO='" + strInvoiceNo.ToString() + "' And Upper(Doc_Type_Text)='" + strDocTypeText.ToUpper() + "'"
                Else
                    If Val(OBJPdfConfig.ToString()) > 0 Then
                        strSQL = "DELETE FROM SALESCHALLAN_PDFS WHERE UNIT_CODE='" + strUnit + "' AND ACCOUNT_CODE='" + strCustomer + "' AND DOC_NO=" + strInvoiceNo + " And Upper(Doc_Type_Text)='" + strDocTypeText.ToUpper() + "'"
                    End If
                    If (boolisSaveSignedInvoicePDF And Val(OBJCommonDigital_EINVOICING_CONFIG.ToString()) > 0 And mblnISTrueSignRequired And Not mblnISCrystalReportRequired) Then
                        strSQL += "DELETE FROM SALESCHALLAN_SIGNED_PDFS WHERE UNIT_CODE='" + strUnit + "' AND ACCOUNT_CODE='" + strCustomer + "' AND DOC_NO=" + strInvoiceNo + " And Upper(Doc_Type_Text)='" + strDocTypeText.ToUpper() + "'"
                        'strSQL += "DELETE FROM COMMONDIGITAL_EINVOICING_LOG WHERE UNIT_CODE='" + strUnit + "' AND CUSTOMER_CODE='" + strCustomer + "' AND DOC_NO=" + strInvoiceNo + " And Upper(Doc_Type_Text)='" + strDocTypeText.ToUpper() + "'"
                    End If
                End If

            End If
            If strSQL <> "" Then
                SqlConnectionclass.ExecuteNonQuery(strSQL)
            End If
            ''INSERT NEW FILE

            strSQL = String.Empty
            DocFILE = GetFileBytes(strInvoicePDFPath)
            If (strDocTypeText.Trim.ToString().ToUpper() = "SUPPLEMENTARY") Then
                strSQL = "INSERT INTO SALESCHALLAN_SUPP_PDFS (DOC_NO,UNIT_CODE,ACCOUNT_CODE,DOC_TYPE_TEXT,PDF_BINARY) VALUES ('" + strInvoiceNo.ToString() + "','" + strUnit + "','" + strCustomer + "','" + strDocTypeText + "',@docFILE )"
            Else
                If Val(OBJPdfConfig.ToString()) > 0 Then
                    strSQL = "INSERT INTO SALESCHALLAN_PDFS (DOC_NO,UNIT_CODE,ACCOUNT_CODE,DOC_TYPE_TEXT,PDF_BINARY) VALUES (" + strInvoiceNo + ",'" + strUnit + "','" + strCustomer + "','" + strDocTypeText + "',@docFILE )"
                End If
            End If


            '' PRAVEEN DIGITAL SIGN --GET DATA FOR SIGNED PDF
            ''102853899
            Dim signedDocFILE As Byte()
            Dim signedBase64 As String
            If (strDocTypeText.Trim.ToString().ToUpper() = "SUPPLEMENTARY") Then
                '' Not Required
            Else
                doubleCOORDINATE_LLX_copy = doubleCOORDINATE_LLX
                doubleCOORDINATE_LLY_copy = doubleCOORDINATE_LLY
                strSignAuthorityName_copy = strSignAuthorityName
                strReason_copy = strReason
                strLocation_copy = strLocation

                If (Val(OBJCommonDigital_EINVOICING_CONFIG.ToString()) > 0 And mblnISTrueSignRequired And Not mblnISCrystalReportRequired) Then
                    If boolisSaveSignedInvoicePDF Then
                        strSQL += "INSERT INTO SALESCHALLAN_SIGNED_PDFS (DOC_NO,UNIT_CODE,ACCOUNT_CODE,DOC_TYPE_TEXT,PDF_BINARY,IS_SIGNED,IS_SEND,ENT_DT) VALUES (" + strInvoiceNo + ",'" + strUnit + "','" + strCustomer + "','" + strDocTypeText + "',@docFILE_SIGNED,1,0,GETDATE() )"
                        'strSQL += "INSERT INTO COMMONDIGITAL_EINVOICING_LOG (DOC_NO,UNIT_CODE,CUSTOMER_CODE,DOC_TYPE_TEXT,IS_SIGNED,IS_SEND,IS_QR_SEND,SIGNED_DT,SEND_DT,QR_SEND_DT,ENT_DT) VALUES (" + strInvoiceNo + ",'" + strUnit + "','" + strCustomer + "','" + strDocTypeText + "',1,1,0,GETDATE(),GETDATE(),NULL,GETDATE() )"
                        strSQL += "EXEC USP_EINVOICING_CONFIG_LOG_INSERT '" + strInvoiceNo + "','" + strUnit + "','" + strCustomer + "','" + lbldescription.Text + "','" + lblcategory.Text + "','" + strDocTypeText + "'"
                        Dim objApi As clsTrueCopyAPI.TrueCopyAPI = New clsTrueCopyAPI.TrueCopyAPI()
                        Dim filedata As String = objApi.GetBase64StringFormByte(DocFILE)
                        If boolAnnextureRequired Then
                            signedBase64 = objApi.GetSignedDocumentBase64(mblnAPIUrl, filedata, mblnPFX_ID, mblnPFX_Pass, mblnAPI_Key, strInvoiceNo + ".pdf", "1[" + doubleCOORDINATE_LLX.ToString + ":" + doubleCOORDINATE_LLY.ToString + "]", strSignAuthorityName & vbCrLf & "Reason: " & strReason & vbCrLf & "Location: " & strLocation, strUnit + strInvoiceNo + strDocTypeText + strTrueCopyDescriptor)
                        Else
                            signedBase64 = objApi.GetSignedDocumentBase64(mblnAPIUrl, filedata, mblnPFX_ID, mblnPFX_Pass, mblnAPI_Key, strInvoiceNo + ".pdf", "[" + doubleCOORDINATE_LLX.ToString + ":" + doubleCOORDINATE_LLY.ToString + "]", strSignAuthorityName & vbCrLf & "Reason: " & strReason & vbCrLf & "Location: " & strLocation, strUnit + strInvoiceNo + strDocTypeText + strTrueCopyDescriptor)
                        End If
                        If signedBase64 <> "" Then
                            If signedBase64.Contains("ERROR ON TRUE COPY SERVER") Then
                                MessageBox.Show(signedBase64)
                                signedBase64 = ""
                            Else
                                signedDocFILE = Convert.FromBase64String(signedBase64)
                            End If
                        End If
                    Else
                        Dim objApi As clsTrueCopyAPI.TrueCopyAPI = New clsTrueCopyAPI.TrueCopyAPI()
                        Dim filedata As String = objApi.GetBase64StringFormByte(DocFILE)
                        If boolAnnextureRequired Then
                            signedBase64 = objApi.GetSignedDocumentBase64(mblnAPIUrl, filedata, mblnPFX_ID, mblnPFX_Pass, mblnAPI_Key, strInvoiceNo + ".pdf", "1[" + doubleCOORDINATE_LLX.ToString + ":" + doubleCOORDINATE_LLY.ToString + "]", strSignAuthorityName & vbCrLf & "Reason: " & strReason & vbCrLf & "Location: " & strLocation, strUnit + strInvoiceNo + strDocTypeText + strTrueCopyDescriptor)
                        Else
                            signedBase64 = objApi.GetSignedDocumentBase64(mblnAPIUrl, filedata, mblnPFX_ID, mblnPFX_Pass, mblnAPI_Key, strInvoiceNo + ".pdf", "[" + doubleCOORDINATE_LLX.ToString + ":" + doubleCOORDINATE_LLY.ToString + "]", strSignAuthorityName & vbCrLf & "Reason: " & strReason & vbCrLf & "Location: " & strLocation, strUnit + strInvoiceNo + strDocTypeText + strTrueCopyDescriptor)
                        End If
                        If signedBase64 <> "" Then
                            If signedBase64.Contains("ERROR ON TRUE COPY SERVER") Then
                                MessageBox.Show(signedBase64)
                                signedBase64 = ""
                            Else
                                signedDocFILE = Convert.FromBase64String(signedBase64)
                            End If
                        End If
                    End If
                ElseIf (mblnISTrueSignRequired And Not mblnISCrystalReportRequired) Then
                    strSQL = "select top 1 ISNULL(SignAuthorityName,'') SignAuthorityName,isnull(Reason,'') Reason,isnull(Location,'') Location from CommonDigital_EINVOICING_CONFIG (NOLOCK) where UNIT_CODE='" & gstrUNITID & "'  AND  IS_ACTIVE=1 order by ENT_DT DESC"
                    dtCommonDigital_EINVOICING_CONFIG = SqlConnectionclass.GetDataTable(strSQL)
                    If dtCommonDigital_EINVOICING_CONFIG.Rows.Count > 0 Then
                        strSignAuthorityName = dtCommonDigital_EINVOICING_CONFIG.Rows(0)("SignAuthorityName")
                        strReason = dtCommonDigital_EINVOICING_CONFIG.Rows(0)("Reason")
                        strLocation = dtCommonDigital_EINVOICING_CONFIG.Rows(0)("Location")
                    End If
                    If strSignAuthorityName <> "" Then
                        doubleCOORDINATE_LLX = Convert.ToDecimal(Find_Value("select TOP 1 ISNULL(COORDINATE_LLX,0) from CommonDigital_EINVOICING_CONFIG (NOLOCK) where UNIT_CODE='" & gstrUNITID & "'  AND  IS_ACTIVE=1 order by ENT_DT DESC"))
                        doubleCOORDINATE_LLY = Convert.ToDecimal(Find_Value("select TOP 1 ISNULL(COORDINATE_LLY,0) from CommonDigital_EINVOICING_CONFIG (NOLOCK) where UNIT_CODE='" & gstrUNITID & "'  AND  IS_ACTIVE=1 order by ENT_DT DESC"))
                    End If

                    Dim objApi As clsTrueCopyAPI.TrueCopyAPI = New clsTrueCopyAPI.TrueCopyAPI()
                    Dim filedata As String = objApi.GetBase64StringFormByte(DocFILE)
                    'If boolAnnextureRequired Then
                    '    signedBase64 = objApi.GetSignedDocumentBase64(mblnAPIUrl, filedata, mblnPFX_ID, mblnPFX_Pass, mblnAPI_Key, strInvoiceNo + ".pdf", "1[" + "0" + ":" + "0" + "]", "----------" & vbCrLf & "Reason: " & "INVOICE" & vbCrLf & "Location: " & "----------", strUnit + strInvoiceNo + strDocTypeText + strTrueCopyDescriptor)
                    'Else
                    '    signedBase64 = objApi.GetSignedDocumentBase64(mblnAPIUrl, filedata, mblnPFX_ID, mblnPFX_Pass, mblnAPI_Key, strInvoiceNo + ".pdf", "[" + "0" + ":" + "0" + "]", "----------" & vbCrLf & "Reason: " & "INVOICE" & vbCrLf & "Location: " & "----------", strUnit + strInvoiceNo + strDocTypeText + strTrueCopyDescriptor)
                    'End If
                    If boolAnnextureRequired Then
                        signedBase64 = objApi.GetSignedDocumentBase64(mblnAPIUrl, filedata, mblnPFX_ID, mblnPFX_Pass, mblnAPI_Key, strInvoiceNo + ".pdf", "1[" + doubleCOORDINATE_LLX.ToString + ":" + doubleCOORDINATE_LLY.ToString + "]", strSignAuthorityName & vbCrLf & "Reason: " & strReason & vbCrLf & "Location: " & strLocation, strUnit + strInvoiceNo + strDocTypeText + strTrueCopyDescriptor)
                    Else
                        signedBase64 = objApi.GetSignedDocumentBase64(mblnAPIUrl, filedata, mblnPFX_ID, mblnPFX_Pass, mblnAPI_Key, strInvoiceNo + ".pdf", "[" + doubleCOORDINATE_LLX.ToString + ":" + doubleCOORDINATE_LLY.ToString + "]", strSignAuthorityName & vbCrLf & "Reason: " & strReason & vbCrLf & "Location: " & strLocation, strUnit + strInvoiceNo + strDocTypeText + strTrueCopyDescriptor)
                    End If

                    If signedBase64 <> "" Then
                        If signedBase64.Contains("ERROR ON TRUE COPY SERVER") Then
                            MessageBox.Show(signedBase64)
                            signedBase64 = ""
                        Else
                            signedDocFILE = Convert.FromBase64String(signedBase64)
                        End If
                    End If
                End If
            End If
            '' PRAVEEN DIGITAL SIGN --GET DATA FOR SIGNED PDF

            oSqlCmd.Parameters.Clear()
            If Val(OBJPdfConfig.ToString()) > 0 Then
                oSqlCmd.Parameters.Add("@docFILE", SqlDbType.VarBinary).Value = DocFILE
            End If

            '' PRAVEEN DIGITAL SIGN
            ''102853899
            If (strDocTypeText.Trim.ToString().ToUpper() = "SUPPLEMENTARY") Then
                '' Not Required
            Else
                If (boolisSaveSignedInvoicePDF And Val(OBJCommonDigital_EINVOICING_CONFIG.ToString()) > 0 And mblnISTrueSignRequired And Not mblnISCrystalReportRequired) Then
                    oSqlCmd.Parameters.Add("@docFILE_SIGNED", SqlDbType.VarBinary).Value = signedDocFILE
                End If
            End If
            If strSQL <> "" Then
                oSqlCmd.CommandText = strSQL
                rows = SqlConnectionclass.ExecuteNonQuery(oSqlCmd)
            End If
            DocFILE = Nothing
            If signedBase64 <> "" Then
                'RESULT = "SUCCESS"
                'SavePDFInvoicesInDB = RESULT
                RESULT.Clear()
                RESULT.Add("SUCCESS")
                RESULT.Add(signedDocFILE)
                SavePDFInvoicesInDB = RESULT
            End If
            Try
                File.Delete(strInvoicePDFPath)
            Catch ex As Exception
            End Try
        Catch ex As Exception
            RESULT.Clear()
            RESULT.Add(ex.Message.ToString())
            SavePDFInvoicesInDB = RESULT
        End Try

    End Function
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
                    Msgbox("Problem While Generating Barcode Image.", vbInformation, ResolveResString(100))
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
    Private Function Msgbox(ByVal Prompt As Object, Optional ByVal Buttons As Microsoft.VisualBasic.MsgBoxStyle = MsgBoxStyle.DefaultButton1, Optional ByVal Title As Object = Nothing) As Microsoft.VisualBasic.MsgBoxResult
        Try
            mP_Connection.RollbackTrans()
        Catch ex As Exception
        End Try
        Msgbox = Microsoft.VisualBasic.Interaction.MsgBox(Prompt, Buttons, Title)
    End Function
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
    Public Function DataExist(ByVal pstrQry As String) As Boolean
        Dim oRs As ADODB.Recordset
        DataExist = False
        If Len(Trim(pstrQry)) = 0 Then Exit Function
        Try
            oRs = mP_Connection.Execute(pstrQry)
            If Not (oRs.BOF And oRs.EOF) Then
                DataExist = True
            End If
            oRs.Close()
            oRs = Nothing
        Catch ex As Exception
            oRs = Nothing
            DataExist = False
            Call gobjError.RAISEERROR_INVOICE(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Function
    Private Function GetPrintMethod(ByVal pstraccoutncode As String) As String
        On Error GoTo ErrHandler
        Dim strQry As String
        Dim Rs As ClsResultSetDB
        GetPrintMethod = String.Empty
        strQry = "Select isnull(PRINT_METHOD,'') as PRINT_METHOD from customer_mst (nolock) where Customer_Code='" & Trim(pstraccoutncode) & "' and UNIT_CODE = '" & gstrUNITID & "'"
        Rs = New ClsResultSetDB
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
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
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
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

    'Added by priti to update new digital sign print
    Private Sub chkNewDigitalSign_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkNewDigitalSign.CheckedChanged
        Try
            If Ctlinvoice.Text > 0 Then
                If chkNewDigitalSign.Checked Then
                    If Msgbox("Do you want to Regenerate New Digital Sign With Current Time ?", MsgBoxStyle.YesNo, "Empro") = MsgBoxResult.Yes Then
                        SqlConnectionclass.ExecuteNonQuery("Update Saleschallan_dtl set TrueCopy_Descriptor=TrueCopy_Descriptor + 1 where Doc_no='" & Ctlinvoice.Text & "' and bill_flag=1 and Unit_code='" & gstrUNITID & "'")

                        SqlConnectionclass.ExecuteNonQuery("Delete From SALESCHALLAN_SIGNED_PDFS where Doc_no='" & Ctlinvoice.Text & "' and Unit_code='" & gstrUNITID & "'") '' Added on 25JUNE2024


                        chkNewDigitalSign.Checked = False
                        chkNewDigitalSign.Enabled = False
                    Else
                        chkNewDigitalSign.Checked = False
                    End If
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    'End by priti to update new digital sign print

 
End Class
