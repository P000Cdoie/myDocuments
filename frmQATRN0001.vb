Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports CrystalDecisions
Imports CrystalDecisions.CrystalReports.Engine
Imports System.IO
Imports System.Data.SqlClient
'*10857384
Imports System.Linq
Imports System.Collections.Generic
'*

Friend Class frmQATRN0001
    Inherits System.Windows.Forms.Form

#Region "REVISION HISTORY"
    '----------------------------------------------------
    'Copyright(c)       - MIND
    'Name of module     - frmQATRN0001.FRM
    'Created by         - Manav Bakshi
    'Modified By        - Geeta Dahra
    'Created Date       - 23April 2001
    'description        - QA GRIN Inspection
    'Revised date       -   1.    12/10/2001
    '                   -   2.    16/01/2002
    '                   -   3.    20/03/2002
    '                   -   4.    25/04/2002(Nitin Sood)
    '                   -   5.    06/05/2002
    '                   -   6.    14/05/2002
    '                   -   7.    27/05/2002
    '                   -   8.    28/05/2002
    '                   -   9.    04/06/2002
    '                   -   10.   05/06/2002
    '                   -   11.   05/06/2002
    '                   -   12.   22/06/2002
    '                   -   13.   09/07/2002
    '                   -   14.   31/07/2002
    '                   -   15.   12/08/2002
    '                   -   16.   20/08/2002
    '                   -   17.   27/08/2002
    '                   -   18. 16/09/2002, Check Out Version No - 33, PMS Issue Log No - 584.
    '                   -   19. 23/10/2002
    '*********************************************************************************************************************
    '                   -   20. 12th March 2003, Check Out Version No - 3, PIMS No - 1900
    '                   -   21. 28th April 2003, Check Out Version No - 4
    '                   -   22. 02nd May 2003, Check Out Version No -5
    '
    'Revision History   - Made rectifications agst Internal Issue Log error nos - 111,113,114,116
    '                   - agst Check out version no. 6
    '                   - AGST FORM NO 6073
    '                   - Issues Reported by Geeta Dahra
    '---------------------------------------------------------------------------------------------------------------------
    '                   -   1.  TAB Sequencing and Enter Keypress Order has been syncronized., Focussing has been corrected.
    '                   -   2.  Decimal Allowed Flag Checking According to Flag and according Message Prompt.
    '                   -   3.  Description amd UOM have been displayed in the GRID.
    '                   -   4.  Quantities in the grid will be displayed as per Decimal Allowed Flag.
    '                   -   5.  Purchase Order No is not displayed and other cosmetic changes. Other filelds are also added.
    '                   -   6.  Back Date Inspection Date is allowed.
    '----------------------------------------------------------------------------------------------------------------------
    '                   -   Design and Coding has been almost Changed to make Compatible with supplementry GRIN
    '                   -   Check Out Version No. - 14, Check Out Date  -   13/05/2002
    '----------------------------------------------------------------------------------------------------------------------
    '                   -   In Actual Quantity, Challan Qty was Displayed.
    '                   -   Check Out Version No. - 18, Check Out Date  -   27/05/2002
    '----------------------------------------------------------------------------------------------------------------------
    '                   -   In Rejected Quantity, Challan Qty was Displayed. And Vend Item Received Qty was not updating correctly.
    '                   -   Check Out Version No. - 19, Check Out Date  -   28/05/2002
    '----------------------------------------------------------------------------------------------------------------------
    '                   -   In QA date, Current Date was saved even if QA Date is not current Date.
    '                   -   Check Out Version No. - 20, Check Out Date  -   04/06/2002
    '----------------------------------------------------------------------------------------------------------------------
    '                   -   For Refreshing, 2 options are given instead of 1 - YES/NO Instead of OK
    '                   -   Cursor Movement is made in The spread but have to use mouse in last row and column.
    '                   -   Check Out Version No. - 21, Check Out Date  -   05/06/2002
    '----------------------------------------------------------------------------------------------------------------------
    '                   -   Invalid Proc Call or Argument
    '                   -   Check Out Version No. - 22, Check Out Date  -   05/06/2002
    '----------------------------------------------------------------------------------------------------------------------
    '                   -   Queries changed as per changes in GRIN form.
    '                   -   Check Out Version No. - 22, Check Out Date  -   05/06/2002
    '----------------------------------------------------------------------------------------------------------------------
    '                   -   Those items having 0 actual qty shall not be displayed., Default Rec Location is dispalyed.
    '                   -   Check Out Version No. - 27, Check Out Date  -   09/07/2002
    '----------------------------------------------------------------------------------------------------------------------
    '                   -   First 2 columns are Frozen.
    '                   -   Check Out Version No. - 28, Check Out Date  -   31/07/2002
    '----------------------------------------------------------------------------------------------------------------------
    '                   -   1.  Remarks Field is saved.
    '                   -   2.  Supp. GRIN Autho. shall update Inspected_Quantity field in GRN_Dtl Table.
    '                   -   Check Out Version No. - 29, Check Out Date  -   12/08/2002
    '----------------------------------------------------------------------------------------------------------------------
    '                   -   1.  GRIN entered Sequence of Itemd.
    '                   -   2.  Material Specification is displayed with description.
    '                   -   3.  Code Review Changes Made.
    '                   -   Check Out Version No. - 30, Check Out Date  -   20/08/2002
    '----------------------------------------------------------------------------------------------------------------------
    '                   -   1.  For Rejected GRIN, updation of Daily Purchase Schedule is checked thoroughly and corrected.
    '                   -   2.  PMS Issue Log No - 447, Check Out Version No. - 32, Check Out Date  -   27/08/2002
    '----------------------------------------------------------------------------------------------------------------------
    '                   -   18. Columns Resizing Allowed and After authorization, grid shall scroll back.
    '----------------------------------------------------------------------------------------------------------------------
    '                   -   19. Excess to PO details displayed in GRID. Also Excess PO Qty is added to Rejection Store.
    '                           Rejection Analysis and Customer Rejection 's Actual Analysis.
    '**********************************************************************************************************************
    '                   -   20. RollBackTrans been Added, For Customer Rejection GRIN, Defect Details capturing as per actual even if all items are rejected.
    '                   -   21. Problem of Zero Accepted & Zero Rejected Quantities Removed.
    '                   -   22. Check of Original/Supp GRIN Removal.
    '                   -   23. 5761:Modified By Preety Jain: Batch Details Entry should be Enabled in Transactions.
    '                   -   24 :Creation of Gate Entry - validation at incoming gate till challan quantities
    'change by Preety Jain on 13-oct-2004 - If Actual Qty is 0 then Accepted and Rejected Qty can be zero.
    'Revision Date  - 09/12/2004
    'Revised By     - Sunil Tiwari
    'Description    - 1. Introduce a Checkbox control for Inspection Entry in Grid.
    '                 2. modify the functionality of the locking of grid buttons according to the value of inspection entry.
    '***************************************************************************
    'Revision Date  - 01-June-2005
    'Revised By     - Sandeep Chadha -
    'Issue No       - 14119
    'Description    - Correct the Issue log of RGP Reconcilation Developed by Sunita.
    '***************************************************************************
    'Revision Date  - 13-June-2005
    'Revised By     - Sourabh Khatri
    'Issue No       - 14799
    'Description    - Generate Incomin material inpection report for APPL.
    '***************************************************************************
    'Revision Date  - 08-July-2005
    'Revised By     - Sourabh/Prashant
    'Issue No       - 14679
    'Description    - Show lot No
    '***************************************************************************
    'Revision Date  - 24-Oct-2005
    'Revised By     - Shabbir Hussain
    'Issue No       - 16066 (Hilex-Both Accepted And rejected Quantity is coming 0 randomly.
    'Description    - Add a function "CheckForvalidAcceptance" to check in case if during
    '                 authorization accepted and rejected quantity of an item in grn_dtl is 0.
    '                 If it occurs then the transaction is rolled back and user is asked
    '                 to authorize the grin again.
    '**************************************************************************************************************************************
    'Revision Date  - 25-Jan-2006
    'Revised By     - Davinder Singh
    'Issue No       - 16968
    'Description    - A vbmodel form is opened when user click on the Print button on the form.
    '                 This vbmodel form gives an error when user click the help to select Doc_no.
    '**************************************************************************************************************************************
    'Revision Date  - 09-Oct-2006
    'Revised By     - Davinder Singh
    'Issue No       - 17893
    'Description    - To implement the functionality of multiple PO's in the same Grin
    '**************************************************************************************************************************************
    'Revision Date  - 29-June-2007
    'Revised By     - Rajeev Gupta
    'Issue No       - 20407
    'Description    - In sales_parameter Table if DirectInspectionGRINPrint = 1 then Inspection Detail Entry(GRIN) Frame is not show (if There is no value in ctlItem_c1 then Inspection Detail Entry(GRIN) Frame will show)
    '               - and if DirectInspectionGRINPrint = 0 then Inspection Detail Entry(GRIN) Frame is Show
    '**************************************************************************************************************************************
    'Revision Date  - 13-Aug-2007
    'Revised By     - Davinder Singh
    'Issue No       - 20852
    'Description    - During Making Supplementary Grin Invoice No field alongwith
    '                 other information is not editable and it takes the same
    '                 invoice No as the parent Grin has. It works fine for all
    '                 type of Grins except Customer Supplied type of Grin (Z)
    '                 as Marketting table Custannex_Hdr table has primary key
    '                 on columns (CustomerCode,InvoiceNo,ItemCode) And in case of
    '                 supplementary grin it gives primary key error.This case may also
    '                 happen if user makes a normal Grin with same Invoice No. for same customer
    '                 So in case of Customer supplied type of Grin we need to
    '                 add the qty received through Supplementary Grin in the Qty
    '                 Received through Parent Grin
    '**************************************************************************************************************************************
    ' Revised By        -   Davinder Singh
    ' Revision Date     -   14 Sep 2007
    ' Issue Id          -   21102
    ' Revision History  -   1) Expired Batches will be displayed in Red colour and
    '                          whole batch Qty. will be Rejected
    '                       2) Concept of shelf-life is added and MFG. and EXP. dates
    '                          are computed and saved in ItemBatch_Mst based on Shelf_Life days defined in Item_Mst
    '                       3) To Check if an earlier Batch already received for same item and same Vendor
    '                       4) Barcode Functionality on this form will work only if both flags (BARCODE_PROCESS,BARCODE_GRN)
    '                          of table BARCODE_CONFIG_MST are set to 1
    '                       5) Barcode related code is added to restrict the user from Authorizing the GRIN
    '                          if Binning of material is not done (This validation will be applicable to
    '                          those items only for which BarcodeTracking is on in Item_Mst)
    '**************************************************************************************************************************************
    ' Revised By                 -   Davinder Singh
    ' Revision Date              -   22 Oct 2007
    ' Issue Id                   -   21345 (Merger)
    ' Revision History           -   1) To Restict the Authorization if previous
    '                                   pending grins exist for Authorization with same items (Flag based)
    '                                2) SCAR
    '**************************************************************************************************************************************
    ' Revised By                 -   Siddharth Ranjan
    ' Revision Date              -   21 Nov 2007
    ' Revision History           -   1) Changes in SP PRC_GRIN_AUTH to stop validation of Grin qty with binned qty
    '**************************************************************************************************************************************
    ' Revised By                 -   neeraj yadav
    ' Revision Date              -   05 dec 2007
    ' Issue Id                   -   21681
    ' Revision History           -   grin edit for barcoded items
    '*********************************************************************************************************************
    ' Revised By                 -   Davinder Singh
    ' Revision Date              -   16 Oct 2008
    ' Issue Id                   -   eMpro -20080924 - 21936
    ' Revision History           -   To Knockoff Purchase Schedule Schedule No. wise
    '*********************************************************************************************************************
    ' Revised By                 -   Siddharth Ranjan
    ' Revision Date              -   17 Nov 2008
    ' Issue Id                   -   eMpro-20081117-23670
    ' Revision History           -   To replace existing function GetQryOutput with DataExist where return type is boolean variable
    '*********************************************************************************************************************
    ' Revised By                 -   Davinder Singh
    ' Revision Date              -   02 FEB 2009
    ' Issue Id                   -   eMpro-20090202-26873 
    ' Revision History           -   Object not closed error Occuring in function Validate_Data()
    '*********************************************************************************************************************
    ' Revised By                 -   Davinder Singh
    ' Revision Date              -   03 Mar 2009
    ' Issue Id                   -   eMpro-20090303-28111 
    ' Revision History           -   QA Date can't be less than current Date
    '*********************************************************************************************************************
    ' Revised By                 -   Davinder Singh
    ' Revision Date              -   25 Jun 2009
    ' Issue Id                   -   eMpro-20090625-32885 
    ' History                    -   Tab Index corrected
    '*********************************************************************************************************************
    ' Revised By                 -   Vinod Singh
    ' Revision Date              -   14/09/2009
    ' Description                -   Validation for Difference in Rate Flag added for 
    '                                Customer Supplied Grin made against Customer Supplied PO
    '*********************************************************************************************************************
    ' Revised By                 -   Vinod Singh
    ' Revision Date              -   25/02/2010
    ' Description                -   Changes for showing missing pakets details if binned qty is less than accepted qty
    '                                for barcode items
    '*********************************************************************************************************************
    ' Revised By                 -   Vinod Singh
    ' Revision Date              -   08/03/2011
    ' Description                -   Changes for Auto Report View 
    '*********************************************************************************************************************
    ' Revised By                 -   SHUBHRA VERMA
    ' Revision Date              -   24 MAY 2011
    ' Description                -   MULTI UNIT CHANGES
    '*********************************************************************************************************************
    ' Revised By                 -   Rajeev Gupta
    ' Revision Date              -   17 Oct 2012
    ' Description                -   10268417 - Exchange rate issue in PV
    '*********************************************************************************************************************
    'Modified By       -  Neha Ghai
    'Modified On       -  09 July 2013
    'Issue Id          -  10384524  
    'Issue Description -  Option to view Item Document added.
    '*********************************************************************************************************************
    ' Revision Date     : 19/08/2013
    ' Revised By        : Vinod Singh
    ' Issue Id          : 10378778
    ' Revision History  : Global Tool Transfer Changes
    '*********************************************************************************************************************
    ' Revision Date     : 18/12/2013
    ' Revised By        : Vinod Singh
    ' Issue Id          : 10504051
    ' Revision History  : Partial Packet Cancellation
    '*********************************************************************************************************************
    'Revised On         : 13 JAN 2014
    'Revised By         : Saurav Kumar
    'Purpose            : Update query was executing inside loop (which was not required).
    'Issue Id           : 10517328
    '***********************************************************************************************************************************
    'Revised On         : 19 MAR 2014
    'Revised By         : Prachi Jain
    'Purpose            : SCAR -Query 
    'Issue Id           : 10557858 
    '***********************************************************************************************************************************
    ' Revision Date     : 24 Jul 2014
    ' Revised By        : Vinod Singh
    ' Issue Id          : 10639761 — Gate entry changes for PO type - Misc
    '***********************************************************************************************************************************
    ' Revision Date     : 31 Aug 2015
    ' Revised By        : Vinod Singh
    ' Issue Id          : 10857384-Service PO Changes
    '***********************************************************************************************************************************
    ' Revised By        : Ekta Uniyal
    ' Revision Date     : 8 Jan 2016
    ' Issue Id          : 10847531-FW GLOBAL DEFECT MASTER CHANGES
    ' Desc              : "Incoming QA" Defect to be show only into GRIN Authorization transactions.
    '***********************************************************************************************************************************
    ' Revised By        : Vinod Singh
    ' Revision Date     : 02 June 2016
    ' Issue Id          : 10949734 — SCAR Mandatory at the time rejection
    '***********************************************************************************************************************************
    ' Revised By        : Ashish Sharma
    ' Revision Date     : 21 AUG 2018
    ' Issue Id          : 101161852 - Quality process IQA
    '***********************************************************************************************************************************
    ' Revised By        : Anand Yadav
    ' Revision Date     : 20 NOV 2018
    ' Issue Id          : 101661477  - Increase  txtinv size
    '***********************************************************************************************************************************
    ' Revised By        : Vinod Singh
    ' Revision Date     : 11 Nov 2019
    ' Revision          : Vend. Inv. Date older than PO date Changes 
    '                       1. if above is approved then grin qty can be fully accepted
    '                       2. if above is rejected then grin qty can only be fully rejected
    '***********************************************************************************************************************************
    ' Revised By        : Vinod Singh
    ' Revision Date     : 22 Nov 2019
    ' Revision          : Auto MTN for JOB Work GRIN
    '***********************************************************************************************************************************
    ' Revised By        : Anand Yadav
    ' Revision Date     : 06 Jan 2021
    ' Revision          : 102278094 — eMPro MATE PUNE QA: GRN auth inspected by 
    '***********************************************************************************************************************************


#End Region

    Dim mintIndex As Short
    Dim strDoc As Double
    Dim intCount As Short
    Dim dblActualQty As Double
    Dim mstrLocationCode As String
    Dim mstrErrorCaption As String
    Dim mstrErrorDesc As String
    Dim mlngErrorNo As Short
    Dim mctlError As System.Windows.Forms.Control
    Dim mstrMasterString As String
    Dim mstrDetailString As String
    Dim mstrChallanNo As String
    Dim mstrCurrencyID As String
    Dim mstrGrinType As String
    Dim mstrGateEntryID As String
    Const Col_Item_Code As Short = 1
    Const Col_Item_Description As Short = 2
    Const Col_Item_UOM As Short = 3
    Const Col_Item_Rate As Short = 4
    Const Col_Challan_Qty As Short = 5
    Const Col_Actual_Qty As Short = 6
    Const Col_Excess_PO_Qty As Short = 7
    Const Col_Accepted_Qty As Short = 8
    Const Col_Rejected_Qty As Short = 9
    Const Col_Rejection_Reason As Short = 10
    Const Col_Inspection_Entry As Short = 11
    Const Col_Asseccible_Rate As Short = 12
    Const Col_GL_Group As Short = 13
    Const Col_Project_Code As Short = 14
    Const Col_Discount_Per As Short = 15
    Const Col_Batch_Details As Short = 17
    Const Col_Rejection_Details As Short = 16
    Const Col_Receipt_Qty As Short = 18
    Const col_Inspection_Control_Details As Short = 19
    Const Col_ViewDoc As Short = 20

    '*10857384
    Const Col_ESI As Short = 21
    '*
    Const Col_IQA As Short = 22

    Const Col_Defect_Code As Short = 1
    Const Col_Defect_Help As Short = 2
    Const Col_Defect_Description As Short = 3
    Const Col_Defect_qty As Short = 4

    '10847531
    Const Col_Defect_Category As Short = 5
    'End Here

    Dim marrDefectDetails(,) As String
    Dim mstrItemCode As String
    Dim mstrGRType As New VB6.FixedLengthString(1)
    Dim mblnBatchTracking As Boolean
    Dim mbln_Inspection_Control_Details As Boolean
    Dim mblnMultiplePO_Allowed As Boolean
    Dim mblnGateEntryRequired As Boolean
    Dim mstrMeasureCode As String
    Public strCertificate_No As String
    Public strCertificate_Desc As String
    Dim strInspectionEntryValue As String
    Dim mblnBarcodeGRN As Boolean
    Dim mblnIQARequired As Boolean = False
    Dim mblnGRINDeviationForm As Boolean = False
    Dim REPDOC As ReportDocument
    Dim REPVIEWER As eMProCrystalReportViewer
    Dim STRREPPATH As String
    'VID CHANGES'
    Dim mblnVendInvDateOlderThanPODate As Boolean = False
    '101161852
    Public Sub New(ByVal paramBlnGRINDeviation As Boolean)
        MyBase.New()
        'This call is required by the Windows Form Designer.
        InitializeComponent()
        mblnGRINDeviationForm = paramBlnGRINDeviation
        Me.MdiParent = prjMPower.mdifrmMain
        prjMPower.mdifrmMain.Show()
        Form_Initialize_Renamed()
        Me.Name = "FRMQATRN0001_GRIN_DEVIATION"
        CheckIQARequired()
        If Not mblnIQARequired Then
            MsgBox("Please first ON the flag [IQA_REQUIRED] in [STORES_CONFIGMST].", MsgBoxStyle.Information, ResolveResString(100))
            Me.Close()
        End If
    End Sub
    Private Enum Enm_fsPOWiseRejectionDtls
        Col_PO_No = 1
        Col_Challan_Quantity = 2
        Col_Actual_Quantity = 3
        Col_Excess_PO_Quantity = 4
        Col_Accepted_Quantity = 5
        Col_Rejected_Quantity = 6
        Col_TmpRejected_Quantity = 7
    End Enum
    Private Structure Item_PO_Dtls
        Dim ItemCode As String
        Dim PONO As Integer
        Dim ChallanQty As Double
        Dim ReceiptQty As Double
        Dim ActualQty As Double
        Dim ExcessPOQty As Double
        Dim AcceptedQty As Double
        Dim RejectedQty As Double
        Dim Rate_Renamed As Double
        Dim AssesableRate As Double
        Dim DiscountAmt As Double
        Dim DiscountPer As Double
        Dim DiffInRates As Boolean
        Dim GLGrpCode As String
        Dim ProjectCode As String
        Dim UOM As String
        Dim Item_Desc As String
        Dim Rej_reason As String
    End Structure
    Dim mArrPOdtl() As Item_PO_Dtls
    Dim mstrPOstatus As String
    Dim mstrUOM As String
    Dim mstrDescription As String
    Dim mstrRejReason As String
    Private mblnShowInpectionGRINFRame As Boolean
    Dim mblnSaveScar As Boolean
    '10857384
    Dim ListESIDetail As New List(Of cls_GRIN_ESI_Detail)
    Private Sub CmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
        On Error GoTo ErrHandler
        With Me
            .fraDefectDetails.Visible = False
            If mblnBatchTracking = False And mbln_Inspection_Control_Details = False Then
                With Me.sprdata
                    If .Row <> .MaxRows Then
                        .Row = .Row + 1
                        .Col = Col_Accepted_Qty
                        .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus() : Exit Sub
                    Else
                        Me.cmdGrpAuthorise1.Focus()
                    End If
                End With
            End If
        End With
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdcansel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdcansel.Click
        '****************************************************
        'Created By     -  Sourabh Khatri
        'Description    -  To Show help for GRin Number
        'Arguments      -  None
        'Return Value   -  None
        '****************************************************
        On Error GoTo Errorhandler
        InspectionDtlEntryEnable(False)
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdDocHlp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdDocHlp.Click
        '****************************************************
        'Created By     -  Sourabh Khatri
        'Description    -  To Show help for GRin Number
        'Arguments      -  None
        'Return Value   -  None
        '****************************************************
        On Error GoTo Errorhandler
        Dim strDocNo() As String
        strDocNo = ctlGRINHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select doc_no," & DateColumnNameInShowList("grn_date") & " As grn_date,vendor_name" & _
            " from grn_hdr Left join vendor_mst on" & _
            " grn_hdr.UNIT_code = vendor_mst.UNIT_code AND " & _
            " grn_hdr.vendor_code = vendor_mst.vendor_code where grn_hdr.UNIT_CODE = '" & gstrUNITID & "' AND" & _
            " GRN_Cancelled = 0", "eMPro")
        If UBound(strDocNo) = -1 Then
            Exit Sub
        ElseIf strDocNo(0) = "0" Then
            Call MsgBox("No Grin exist to show", MsgBoxStyle.OkOnly, ResolveResString(100))
            Exit Sub
        Else
            txtDocNo.Text = strDocNo(0)
            cmdPrint.Focus()
        End If
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
        On Error GoTo ErrHandler
        Dim strSql As String
        Dim strHelp() As String
        If txtrecloc.Text.Trim.Length = 0 Then
            MsgBox("First Enter From Location Code.", MsgBoxStyle.Information, ResolveResString(100))
            txtrecloc.Focus()
            Exit Sub
        End If
        '*COMMENTED AND REWRITTEN , ISSUE ID : 10857384
        'strSql = "SELECT DOC_NO," & DateColumnNameInShowList("GRN_DATE") & "As GRN_DATE FROM GRN_HDR" & _
        '        " WHERE UNIT_CODE = '" & gstrUNITID & "' AND DOC_TYPE = 10" & _
        '        " AND FROM_LOCATION = '" & txtrecloc.Text.Trim & "'" & _
        '        " AND ISNULL(LEN(QA_Authorized_Code),0) = 0" & _
        '        " AND inspection_auth='" & mP_User & "'"
        '101161852
        If mblnGRINDeviationForm Then
            strSql = "SELECT DOC_NO,GRN_DATE FROM DBO.UFN_PENDING_GRIN_FOR_AUTHORIZATION_DEVIATION('" & gstrUnitId & "','" & txtrecloc.Text & "','" & mP_User & "')"
        Else
            strSql = "SELECT DOC_NO,GRN_DATE FROM UFN_PENDING_GRIN_FOR_AUTHORIZATION('" & gstrUnitId & "','" & txtrecloc.Text & "','" & mP_User & "')"
        End If

        '*
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        strHelp = ctlGRINHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSql, "Grin No(s) Help", 1)
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)

        If UBound(strHelp) = -1 Then Exit Sub
        If strHelp(0) = "0" Then
            Call MsgBox("No GRIN is Pending for Authorization", MsgBoxStyle.Information, "Information")
        Else
            ctlitem_c1.Text = strHelp(0)
            ctlitem_c1.Focus()
            System.Windows.Forms.Application.DoEvents()
            Call ctlItem_c1_Validating(ctlitem_c1, New System.ComponentModel.CancelEventArgs(False))
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub SetInspectionEvironment()
        Dim strRejectedQuantity As String
        With Me.sprdata
            .Row = .ActiveRow
            .Col = Col_Inspection_Entry
            strInspectionEntryValue = Trim(.Text)
            .Row = .ActiveRow
            .Col = Col_Rejected_Qty
            strRejectedQuantity = Trim(.Text)
            If strInspectionEntryValue = "0" Or strInspectionEntryValue = "" Then
                If Len(strRejectedQuantity) > 0 And Val(strRejectedQuantity) <> Val("0") Then '' rejected_qty > 0
                    .Col = Col_Rejection_Details
                    .Lock = False
                    .Col = col_Inspection_Control_Details
                    .Lock = True
                Else
                    .Col = Col_Rejection_Details
                    .Lock = True
                    .Col = col_Inspection_Control_Details
                    .Lock = True
                End If
            Else
                If Len(strRejectedQuantity) > 0 And Val(strRejectedQuantity) <> Val("0") Then '' rejected_qty > 0
                    .Col = Col_Rejection_Details
                    .Lock = True
                    .Col = col_Inspection_Control_Details
                    .Lock = False
                Else
                    .Col = Col_Rejection_Details
                    .Lock = True
                    .Col = col_Inspection_Control_Details
                    .Lock = False
                End If
            End If
        End With
    End Sub
    Private Sub CmdHelpEmpQA_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdHelpEmpQA.Click
        On Error GoTo ErrHandler
        Dim strGrnumberHelp As String
        If Len(Me.txtEmpQA.Text) = 0 Then 'To check if There is No Text Then Show All Help
            'strGrnumberHelp = ShowList(1, (txtEmpQA.MaxLength), "", "user_id", "fullname", "user_mst", , "LIst of User(s)")'102278094
            strGrnumberHelp = ShowList(1, (txtEmpQA.MaxLength), "", "user_id", "fullname", "user_mst", " and InActive='N' ")
            If (strGrnumberHelp = "-1") Then
                Call MsgBox("No User(s) are Defined. First Define User(s).", MsgBoxStyle.Information, "Information")
            Else
                txtEmpQA.Text = strGrnumberHelp
                System.Windows.Forms.Application.DoEvents()
                Call txtEmpQA_Validating(txtEmpQA, New System.ComponentModel.CancelEventArgs(False))
            End If
        Else
            'strGrnumberHelp = ShowList(1, (txtEmpQA.MaxLength), "", "user_id", "fullName", "user_mst", , "List of User(s)")'102278094
            strGrnumberHelp = ShowList(1, (txtEmpQA.MaxLength), "", "user_id", "fullName", "user_mst", " and InActive='N' ")
            If (strGrnumberHelp = "-1") Then
                Call MsgBox("No User(s) are Defined. First Define User(s).", MsgBoxStyle.Information, "Information")
            Else
                txtEmpQA.Text = strGrnumberHelp
                System.Windows.Forms.Application.DoEvents()
                Call txtEmpQA_Validating(txtEmpQA, New System.ComponentModel.CancelEventArgs(False))
                Me.dtpQADate.Focus() : Exit Sub
            End If
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
        On Error GoTo ErrHandler
        Dim intTotalRows As Short
        Dim intCounter As Short
        Dim intCounter1 As Short
        Dim dblDefectQty As Double
        Dim strDefectCode As New VB6.FixedLengthString(3)
        Dim intRow As Short
        Dim varinspent As Object
        With Me
            If .lblDefRejection.Text = "0" Then ' If Total rejection agst Defects is 0
                MsgBox("Total Rejection against Defects Can't be Zero.", MsgBoxStyle.Information, ResolveResString(100))
                With Me.vaDefects
                    .Row = 1
                    .Col = Col_Defect_qty
                    .EditModePermanent = True
                    .EditModeReplace = True
                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                    Exit Sub
                End With
            ElseIf Val(.lblDefRejection.Text) <> Val(lblTotalRej.Text) Then  ' If Total rejection agst Defects<> Total rejection for Item
                MsgBox("Total Rejection against Defects Should be Equal to Total Rejected Qty. " & Me.lblTotalRej.Text, MsgBoxStyle.Information, ResolveResString(100))
                With vaDefects
                    .Row = 1
                    .Col = Col_Defect_qty
                    .EditModePermanent = True
                    .EditModeReplace = True
                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell : Exit Sub
                End With
            Else ' If Correct information of Defects with quantities is entered.
                With vaDefects
                    intTotalRows = .MaxRows
                    For intCounter1 = 1 To UBound(marrDefectDetails, 2) ' For Each Row in the Array
                        If StrComp(marrDefectDetails(0, intCounter1), mstrItemCode) = 0 Then ' If Item already exists in the Array
                            marrDefectDetails(0, intCounter1) = ""
                            marrDefectDetails(1, intCounter1) = ""
                            marrDefectDetails(2, intCounter1) = "0" 'Their Defect Quantities are Set to 0
                        End If
                    Next
                    For intCounter = 1 To intTotalRows
                        .Row = intCounter
                        .Col = Col_Defect_Code
                        strDefectCode.Value = .Text.Trim
                        .Row = intCounter
                        .Col = Col_Defect_qty
                        dblDefectQty = Val(.Text)
                        If dblDefectQty > 0 Then
                            If UBound(marrDefectDetails) = 0 Then ' If Nothing exists in the array.
                                ReDim marrDefectDetails(3, 1)
                                marrDefectDetails(0, UBound(marrDefectDetails, 2)) = mstrItemCode
                                marrDefectDetails(1, UBound(marrDefectDetails, 2)) = strDefectCode.Value
                                marrDefectDetails(2, UBound(marrDefectDetails, 2)) = CStr(dblDefectQty)
                            Else ' If some entries are there in the Array.
                                For intCounter1 = 1 To UBound(marrDefectDetails, 2) ' For Each Row in the Array
                                    If StrComp(marrDefectDetails(0, intCounter1), mstrItemCode) = 0 Then ' If Item already exists in the Array
                                        If StrComp(marrDefectDetails(1, intCounter1), strDefectCode.Value) = 0 Then ' If Defect already exists in the array.
                                            marrDefectDetails(0, intCounter1) = mstrItemCode
                                            marrDefectDetails(1, intCounter1) = strDefectCode.Value
                                            marrDefectDetails(2, intCounter1) = CStr(dblDefectQty)
                                            Exit For
                                        Else ' Item is in the array, but Defect is Not in the Array.
                                            If dblDefectQty > 0 Then
                                                ReDim Preserve marrDefectDetails(3, UBound(marrDefectDetails, 2) + 1)
                                                marrDefectDetails(0, UBound(marrDefectDetails, 2)) = mstrItemCode
                                                marrDefectDetails(1, UBound(marrDefectDetails, 2)) = strDefectCode.Value
                                                marrDefectDetails(2, UBound(marrDefectDetails, 2)) = CStr(dblDefectQty)
                                                Exit For
                                            End If
                                        End If
                                    Else ' If Item does not exist in the Array Already.
                                        ReDim Preserve marrDefectDetails(3, UBound(marrDefectDetails, 2) + 1)
                                        marrDefectDetails(0, UBound(marrDefectDetails, 2)) = mstrItemCode
                                        marrDefectDetails(1, UBound(marrDefectDetails, 2)) = strDefectCode.Value
                                        marrDefectDetails(2, UBound(marrDefectDetails, 2)) = CStr(dblDefectQty)
                                        Exit For
                                    End If
                                Next
                            End If
                        End If
                    Next
                End With
                fraDefectDetails.Visible = False
                With sprdata
                    intRow = .Row
                    If intRow = .MaxRows Then
                        If mblnBatchTracking = True Then
                            .Focus()
                            .Row = intRow
                            .Col = Col_Batch_Details
                            .EditModePermanent = True
                            .EditModeReplace = True
                            .Action = FPSpreadADO.ActionConstants.ActionActiveCell : Exit Sub
                        ElseIf mbln_Inspection_Control_Details = True Then
                            Call .SetText(Col_Inspection_Entry, intRow, varinspent)
                            If varinspent = "1" Then
                                .Focus()
                                .Row = intRow
                                .Col = col_Inspection_Control_Details
                                .EditModePermanent = True
                                .EditModeReplace = True
                                .Action = FPSpreadADO.ActionConstants.ActionActiveCell : Exit Sub
                            Else
                                cmdGrpAuthorise1.Focus() : Exit Sub
                            End If
                        Else
                            cmdGrpAuthorise1.Focus() : Exit Sub
                        End If
                    Else
                        If mblnBatchTracking = True Then
                            .Focus()
                            .Row = intRow
                            .Col = Col_Batch_Details
                            .EditModePermanent = True
                            .EditModeReplace = True
                            .Action = FPSpreadADO.ActionConstants.ActionActiveCell : Exit Sub
                        ElseIf mbln_Inspection_Control_Details = True Then
                            Call .SetText(Col_Inspection_Entry, intRow, varinspent)
                            If varinspent = "1" Then
                                .Focus()
                                .Row = intRow
                                .Col = col_Inspection_Control_Details
                                .EditModePermanent = True
                                .EditModeReplace = True
                                .Action = FPSpreadADO.ActionConstants.ActionActiveCell : Exit Sub
                            Else
                                .Focus()
                                .Row = intRow + 1
                                .Col = Col_Accepted_Qty
                                .EditModePermanent = True
                                .EditModeReplace = True
                                .Action = FPSpreadADO.ActionConstants.ActionActiveCell : Exit Sub
                            End If
                        Else
                            .Focus()
                            .Row = intRow + 1
                            .Col = Col_Accepted_Qty
                            .EditModePermanent = True
                            .EditModeReplace = True
                            .Action = FPSpreadADO.ActionConstants.ActionActiveCell : Exit Sub
                        End If
                    End If
                End With
            End If
        End With
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdPORejCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdPORejCancel.Click
        On Error GoTo ErrHandler
        fraPoWiseRejectionDtl.Visible = False
        InspectionDtlEntryEnable(False)
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdPORejOk_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdPORejOk.Click
        '***************************************************************************
        'Revision Date  - 07-Oct-2006
        'Revised By     - Davinder Singh
        'Issue No       - 17893
        'Description    - To Fill the PO wise rejection details in array
        '***************************************************************************
        On Error GoTo ErrHandler
        Dim intCtrArr As Short
        Dim intCtrGrid As Short
        Dim lngPONO As Integer
        With fsPOWiseRejectionDtls
            If Val(lblRemainingRejectedQty.Text) < 0 Then
                MsgBox("Total Rejection against PO(s) can not exceed total Rejection Qty: " & lblTotalRejectedQty.Text, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                .Row = 1
                .Col = Enm_fsPOWiseRejectionDtls.Col_Rejected_Quantity
                .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                .Focus()
            ElseIf Val(lblRemainingRejectedQty.Text) > 0 Then
                MsgBox("Total Rejection against PO(s) can not less than total Rejection Qty: " & lblTotalRejectedQty.Text, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                .Row = 1
                .Col = Enm_fsPOWiseRejectionDtls.Col_Rejected_Quantity
                .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                .Focus()
            Else
                For intCtrGrid = 1 To .MaxRows Step 1
                    For intCtrArr = 1 To UBound(mArrPOdtl) Step 1
                        If StrComp(mstrItemCode, mArrPOdtl(intCtrArr).ItemCode, CompareMethod.Text) = 0 Then
                            .Row = intCtrGrid
                            .Col = Enm_fsPOWiseRejectionDtls.Col_PO_No
                            lngPONO = CInt(Trim(.Text))
                            If lngPONO = mArrPOdtl(intCtrArr).PONO Then
                                .Col = Enm_fsPOWiseRejectionDtls.Col_Accepted_Quantity
                                mArrPOdtl(intCtrArr).AcceptedQty = CDbl(Trim(.Text))
                                .Col = Enm_fsPOWiseRejectionDtls.Col_Rejected_Quantity
                                mArrPOdtl(intCtrArr).RejectedQty = CDbl(Trim(.Text))
                                mArrPOdtl(intCtrArr).UOM = mstrUOM
                                mArrPOdtl(intCtrArr).Item_Desc = mstrDescription
                                mArrPOdtl(intCtrArr).Rej_reason = mstrRejReason
                                Exit For
                            End If
                        End If
                    Next intCtrArr
                Next intCtrGrid
                fraPoWiseRejectionDtl.Visible = False
                InspectionDtlEntryEnable(False)
            End If
        End With
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdPrint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdPrint.Click
        '****************************************************
        'Created By     -  Sourabh Khatri
        'Description    -  To Print Grin/ispection Report
        'Arguments      -  None
        'Return Value   -  None
        '****************************************************
        On Error GoTo Errorhandler
        Dim strSql As String
        Dim blnStatus As String
        strSql = "SELECT TOP 1 1 FROM GRN_HDR" & _
                " where UNIT_CODE = '" & gstrUNITID & "' AND doc_no = '" & txtDocNo.Text.Trim & "'" & _
                " and GRN_Cancelled = 0"
        blnStatus = DataExist(strSql)
        If txtDocNo.Text.Trim.Length = 0 Then
            MsgBox("GRIN number can not be blank", MsgBoxStyle.Information, ResolveResString(100))
            txtDocNo.Focus()
            Exit Sub
        ElseIf Not IsNumeric(txtDocNo.Text.Trim) Then
            MsgBox("GRIN number should be numeric ", MsgBoxStyle.Information, ResolveResString(100))
            txtDocNo.SelectionStart = 0
            txtDocNo.SelectionLength = txtDocNo.Text.Length
            txtDocNo.Focus()
            Exit Sub
        ElseIf blnStatus = False Then
            MsgBox("Invalid GRIN number ", MsgBoxStyle.Information, ResolveResString(100))
            txtDocNo.SelectionStart = 0
            txtDocNo.SelectionLength = Len(txtDocNo.Text)
            txtDocNo.Focus()
            Exit Sub
        Else
            Call ShowGrinReport()
        End If
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdRecLocList_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdRecLocList.Click
        On Error GoTo ErrHandler
        Dim strSql As String
        Dim strHelp() As String
        If optGRIN.Checked = True Then
            strSql = "SELECT LOCATION_CODE,DESCRIPTION FROM LOCATION_MST" & _
                    " WHERE UNIT_CODE = '" & gstrUNITID & "' AND upper(loc_type) ='R'"
        Else
            strSql = "SELECT LOCATION_CODE,DESCRIPTION FROM LOCATION_MST" & _
                    " WHERE UNIT_CODE = '" & gstrUNITID & "' AND upper(loc_type) ='J'"
        End If
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        strHelp = ctlGRINHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSql, "List of Location(s)", 1)
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        If UBound(strHelp) = -1 Then Exit Sub
        If strHelp(0) = "0" Then
            MsgBox("No From Location(s) are defined. First Define Receipt Location(s).", MsgBoxStyle.Information, ResolveResString(100))
            txtrecloc.Focus()
        Else
            txtrecloc.Text = strHelp(0)
            lblRecLoc.Text = strHelp(1)
            Call txtRecLoc_Validating(txtrecloc, New System.ComponentModel.CancelEventArgs(False))
        End If
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub ctlFormHeader1_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        '--------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Arguments     : Nil
        ' Return Value  : Nil
        ' Function      : To call the appropriate Help page
        ' Datetime      : 07-Nov-2006
        '--------------------------------------------------------------------
        On Error GoTo ErrHandler
        Call ShowHelp("HLP" & Mid(Me.Name, 4, Len(Me.Name)) & ".htm")
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub ctlItem_c1_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ctlitem_c1.TextChanged
        On Error GoTo ErrHandler
        If Len(Me.ctlitem_c1.Text) = 0 Then
            mstrLocationCode = Me.txtrecloc.Text
            Call RefreshFrm() : ctlitem_c1.Focus() : Me.txtrecloc.Text = mstrLocationCode : Call txtRecLoc_Validating(txtrecloc, New System.ComponentModel.CancelEventArgs(False)) : Exit Sub
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub ctlItem_c1_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles ctlitem_c1.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 13 And KeyAscii <> 8 Then KeyAscii = 0
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                ctlItem_c1_Validating(ctlitem_c1, New System.ComponentModel.CancelEventArgs(False))
        End Select
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub ctlItem_c1_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles ctlitem_c1.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F1 Then Call cmdHelp_Click(cmdHelp, New System.EventArgs())
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub ctlItem_c1_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles ctlitem_c1.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        '*****************************************************************************************
        'Revised By     - Davinder Singh
        'Revision Date  - 15 SEP 2007
        'Description    - To Check if Complete Barcode Process Followed or Not
        '                 by calling Stored Procedure PRC_GRIN_AUTH
        '                 This checking will be made if BarCode Process for GRIN is on at site
        '*****************************************************************************************
        On Error GoTo ErrHandler
        Dim strsql As String
        Dim blnStatus As Boolean
        Dim cmd As ADODB.Command
        ''--- If Lenght of GRIN Number IS Zero then Clear all the Fields
        If txtrecloc.Text.Trim.Length = 0 Then
            MsgBox("First Enter From Location.", MsgBoxStyle.Information, ResolveResString(100))
            txtrecloc.Enabled = True
            txtrecloc.Focus()
            GoTo EventExitSub
        End If
        If ctlitem_c1.Text.Length = 0 Then
            mstrLocationCode = txtrecloc.Text
            Call RefreshFrm()
            txtrecloc.Text = mstrLocationCode
            GoTo EventExitSub
        Else
            ReDim garrBatchDetails(7, 0)
            strDoc = Val(ctlitem_c1.Text)
            '*10857384
            'strsql = "SELECT TOP 1 1 FROM Grn_Hdr" & _
            '        " WHERE UNIT_CODE = '" & gstrUNITID & "' AND" & _
            '        " from_location='" & txtrecloc.Text.Trim & "'" & _
            '        " AND doc_type=10" & _
            '        " AND doc_no=" & strDoc & _
            '        " AND ISNULL(LEN(qa_authorized_code),0)=0" & _
            '        " AND inspection_auth='" & mP_User & "'"
            If mblnGRINDeviationForm Then
                strsql = "SELECT TOP 1 1 FROM DBO.UFN_PENDING_GRIN_FOR_AUTHORIZATION_DEVIATION('" & gstrUNITID & "','" & txtrecloc.Text & "','" & mP_User & "') WHERE DOC_NO=" & strDoc & ""
            Else
                strsql = "SELECT TOP 1 1 FROM UFN_PENDING_GRIN_FOR_AUTHORIZATION('" & gstrUNITID & "','" & txtrecloc.Text & "','" & mP_User & "') WHERE DOC_NO=" & strDoc & " "
            End If
            '*
            blnStatus = DataExist(strsql)
            If blnStatus = True Then
                txtEmpQA.Enabled = True
                CmdHelpEmpQA.Enabled = True
                cmdGrpAuthorise1.Enabled(1) = True
                CmdCertificates.Enabled = True
                Call display_details()
                txtEmpQA.Focus()
            Else
                strsql = "SELECT TOP 1 1 FROM Grn_Hdr" & _
                        " WHERE UNIT_CODE = '" & gstrUNITID & "' AND from_location='" & txtrecloc.Text.Trim & "'" & _
                        " AND doc_type=10" & _
                        " AND doc_no=" & strDoc & _
                        " AND ISNULL(LEN(qa_authorized_code),0)>0"
                blnStatus = DataExist(strsql)
                If blnStatus = True Then
                    MsgBox("Entered GRIN is Already Authorized.", MsgBoxStyle.Information, ResolveResString(100))
                End If
                strsql = "SELECT TOP 1 1 FROM Grn_Hdr" & _
                        " WHERE UNIT_CODE = '" & gstrUNITID & "' AND" & _
                        " from_location='" & txtrecloc.Text.Trim & "'" & _
                        " AND doc_type=10" & _
                        " AND doc_no=" & strDoc & _
                        " AND ISNULL(LEN(qa_authorized_code),0)=0"
                blnStatus = DataExist(strsql)
                If blnStatus = True Then
                    MsgBox("You are not Authorized to inspect selected GRIN", MsgBoxStyle.Information, ResolveResString(100))
                End If
                ctlitem_c1.Text = String.Empty
                ctlitem_c1.Focus()
                GoTo EventExitSub
            End If
        End If
        If mblnVendInvDateOlderThanPODate = False OrElse txtReqApprovalStatus.Text.Trim.ToUpper <> "REJECTED" Then
            If mblnBarcodeGRN = True And Barcode_Location(txtrecloc.Text, txtdesloc.Text) Then
                cmd = New ADODB.Command
                With cmd
                    .let_ActiveConnection(mP_Connection)
                    .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                    .CommandText = "PRC_GRIN_AUTH"
                    .Parameters.Append(.CreateParameter("@UNITCODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                    .Parameters.Append(.CreateParameter("@GRN_NO", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, ADODB.DataTypeEnum.adInteger, Val(ctlitem_c1.Text)))
                    .Parameters.Append(.CreateParameter("@LOCATION_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, txtdesloc.Text))
                    .Parameters.Append(.CreateParameter("@MSG", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamOutput, 500))
                    .Parameters.Append(.CreateParameter("@ERR", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamOutput, ADODB.DataTypeEnum.adInteger))
                    .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                End With
                If cmd.Parameters(cmd.Parameters.Count - 1).Value <> 0 Then
                    MsgBox("Error while Validating Barcode Process For Items", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100))
                    cmd = Nothing
                    GoTo ErrHandler
                End If
                If Len(cmd.Parameters(cmd.Parameters.Count - 2).Value) <> 0 Then
                    MsgBox(cmd.Parameters(cmd.Parameters.Count - 2).Value, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    Me.ctlitem_c1.Text = ""
                    Me.ctlitem_c1.Focus()
                End If
                cmd = Nothing
            End If
        End If
        GoTo EventExitSub

ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub dtpQADate_Change(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        On Error GoTo ErrHandler
        dtpQADate.MinDate = GetServerDate() : dtpQADate.MaxDate = GetServerDate()
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub dtpQADate_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSComCtl2.DDTPickerEvents_KeyDownEvent)
        On Error GoTo ErrHandler
        If eventArgs.keyCode = System.Windows.Forms.Keys.Return Then
            With Me.sprdata
                If .MaxRows > 0 Then
                    .Row = 1
                    .Col = Col_Accepted_Qty
                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus() : Exit Sub
                End If
            End With
        End If
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub dtpQADate_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs)
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        If Len(Me.ctlitem_c1.Text) > 0 Then
            With Me.sprdata
                .Row = 1
                .Col = Col_Accepted_Qty
                .EditModePermanent = True
                .EditModeReplace = True
                .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus() : GoTo EventExitSub
            End With
        End If
        GoTo EventExitSub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub frmQATRN0001_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Me.Tag) = False
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub Form_Initialize_Renamed()
        On Error GoTo ErrHandler
        Dim RS As ClsResultSetDB
        RS = New ClsResultSetDB
        RS.GetResult("SELECT DirectInspectionGRINPrint FROM sales_parameter where UNIT_CODE = '" & gstrUNITID & "'")
        If RS.GetValue("DirectInspectionGRINPrint") = True Then
            mblnShowInpectionGRINFRame = False
        Else
            mblnShowInpectionGRINFRame = True
        End If
        RS.ResultSetClose()
        RS = Nothing
        Exit Sub 'To prevent the execution of errhandler
ErrHandler:
        RS = Nothing
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmQATRN0001_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '------------------------------------------------------------------------------------------------------'
        ' Author        : Davinder singh
        ' Arguments     : Keycode and shift
        ' Return Value  : Nil
        ' Function      : To invoke the onlinehelp associated with form
        ' Datetime      : 26 April 2005
        '------------------------------------------------------------------------------------------------------'
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0 Then
            Call ctlFormHeader1_ClickEvent(ctlFormHeader1, New System.EventArgs())
        End If
        Exit Sub 'To prevent the execution of errhandler
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmQATRN0001_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo ErrHandler
        Dim strsql As String
        Call Initialize_controls()

        '*10857384
        sprdata.MaxCols = Col_IQA
        '*
        Call SetGridHdrs()

        With Me.sprdata
            .Row = .MaxRows
            .Row2 = .MaxRows
            .Col = Col_Item_Code
            .Col2 = Col_Accepted_Qty
            .BlockMode = True
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .BlockMode = False
            .ColsFrozen = 2
            .UserResizeCol = FPSpreadADO.UserResizeConstants2.UserResizeOn
            mblnBatchTracking = FormLevelBatch_Tracking(Me.Name)
            mblnGateEntryRequired = GateEntryGrn_Req()
            mbln_Inspection_Control_Details = Inspection_Control_Detail()
            mblnMultiplePO_Allowed = MultiplePO_Allowed()
            mblnBarcodeGRN = GetBarcodeProcessGRN()
            CheckIQARequired()
            .Col = Col_Asseccible_Rate
            .Col2 = Col_Rejection_Details
            .Row = 1
            .Row2 = .MaxRows
            .BlockMode = True
            .ColHidden = True
            .BlockMode = False

            '*10857384
            '.MaxCols = Col_ViewDoc
            '.MaxCols = Col_ESI
            '*

            .set_ColWidth(18, 800)

            If mbln_Inspection_Control_Details = True Then
                '10384524
                '.MaxCols = 20
                .set_ColWidth(19, 1400)
                .Row = 0
                .Row2 = .MaxRows
                .Col = Col_Inspection_Entry
                .Col2 = Col_Inspection_Entry
                .ColHidden = False
                .Row = 0
                .Col = col_Inspection_Control_Details : .Text = "Inspection Details"
                .ColHidden = True
                .Row = 1
                .Row2 = .MaxRows
                .Col = col_Inspection_Control_Details
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonText = "Inspection Details"
                strsql = "select * into #inspection_entry_hdr from inspection_entry_hdr WHERE UNIT_CODE = '" & gstrUNITID & "'"
                mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                strsql = "select * into #inspection_entry_dtl from inspection_entry_dtl WHERE UNIT_CODE = '" & gstrUNITID & "'"
                mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            Else
                .Row = 0
                .Row2 = .MaxRows
                .Col = Col_Inspection_Entry
                .Col2 = Col_Inspection_Entry
                .ColHidden = True
            End If
            .Col = Col_Receipt_Qty
            .Col2 = Col_Receipt_Qty
            .Row = 1
            .Row2 = .MaxRows
            .BlockMode = True
            .ColHidden = True
            .BlockMode = False
            '.Row = 1
            '.Row2 = .MaxRows
            '.Col = Col_Rejection_Details
            '.Col2 = Col_Rejection_Details
            '.BlockMode = True
            '.ColHidden = False
            '.BlockMode = False
            If mblnBatchTracking = True Then ReDim garrBatchDetails(7, 0)
            If mblnBatchTracking = True Then
                .Row = 1
                .Row2 = .MaxRows
                .Col = Col_Batch_Details
                .Col2 = Col_Batch_Details
                .BlockMode = True
                .ColHidden = False
                .BlockMode = False
            Else
                .Row = 1
                .Row2 = .MaxRows
                .Col = Col_Batch_Details
                .Col2 = Col_Batch_Details
                .BlockMode = True
                .ColHidden = True
                .BlockMode = False
            End If
            '101161852
            If mblnIQARequired Then
                .Row = 1
                .Row2 = .MaxRows
                .Col = Col_IQA
                .Col2 = Col_IQA
                .BlockMode = True
                .ColHidden = False
                .BlockMode = False
                .Row = 1
                .Row2 = .MaxRows
                .Col = Col_Rejection_Details
                .Col2 = Col_Rejection_Details
                .BlockMode = True
                .ColHidden = True
                .BlockMode = False
            Else
                .Row = 1
                .Row2 = .MaxRows
                .Col = Col_IQA
                .Col2 = Col_IQA
                .BlockMode = True
                .ColHidden = True
                .BlockMode = False
                .Row = 1
                .Row2 = .MaxRows
                .Col = Col_Rejection_Details
                .Col2 = Col_Rejection_Details
                .BlockMode = True
                .ColHidden = False
                .BlockMode = False
            End If
        End With
        '101161852
        If mblnGRINDeviationForm Then
            Me.ctlFormHeader1.Tag = "FRMQATRN0001_GRIN_DEVIATION"
            ctlFormHeader1.HeaderString = "Inspection Entry (GRIN Deviation)"
        End If
        mintIndex = mdifrmMain.AddFormNameToWindowList(Me.ctlFormHeader1.Tag)
        Call InspectionDtlEntryIconsLoad()
        btnIQCReport.Enabled = True

        'VID CHANGES
        CheckFlags()
        grpInvDtPODtApproval.Visible = mblnVendInvDateOlderThanPODate

        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub Initialize_controls()
        On Error GoTo ErrHandler
        ReDim marrDefectDetails(0, 0)
        Call MakeRowHeight()
        Call FillLabelFromResFile(Me)
        mlngErrorNo = 1
        mstrErrorDesc = String.Empty
        mstrErrorCaption = ResolveResString(10059)
        Call FitToClient(Me, Frame1, ctlFormHeader1, (Me.cmdGrpAuthorise1), 600)
        Call EnableControls(False, Me)
        With Me
            .optGRIN.Enabled = True : .optGRIN.Checked = True : .optSuppGRIN.Enabled = True
            .txtrecloc.Enabled = True : .txtrecloc.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : .cmdRecLocList.Enabled = True
            .ctlitem_c1.Enabled = True : .ctlitem_c1.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : .cmdHelp.Enabled = True
            .txtEmpQA.Enabled = True : .txtEmpQA.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : .CmdHelpEmpQA.Enabled = True
            .dtpQADate.Enabled = True : .dtpQADate.Format = DateTimePickerFormat.Custom : .dtpQADate.CustomFormat = gstrDateFormat
            .dtpQADate.MinDate = GetServerDate() : .dtpQADate.MaxDate = GetServerDate() : .dtpQADate.Value = GetServerDate()
            .txtRefGRINDate.Text = String.Empty
            .fraDefectDetails.Visible = False
            strCertificate_No = String.Empty : strCertificate_Desc = String.Empty
            .txtrecloc.Text = GetQryOutput("select LOCATION_CODE from location_mst where UNIT_CODE = '" & gstrUNITID & "' AND upper(loc_type)='R'")
            .lblRecLoc.Text = GetQryOutput("select DESCRIPTION from location_mst where UNIT_CODE = '" & gstrUNITID & "' AND location_code='" & .txtrecloc.Text.Trim & "'")
            '.vaDefects.CursorStyle = FPSpreadADO.CursorStyleConstants.CursorStyleArrow
            cmdGrpAuthorise1.Enabled(2) = True
        End With
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub RefreshFrm()
        On Error GoTo ErrHandler
        With Me
            ReDim marrDefectDetails(0, 0)
            ReDim garrBatchDetails(7, 0)
            If mbln_Inspection_Control_Details = True Then
                mP_Connection.Execute("delete from  #inspection_entry_hdr WHERE UNIT_CODE = '" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                mP_Connection.Execute("delete from #inspection_entry_dtl WHERE UNIT_CODE = '" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End If
            .sprdata.MaxRows = 0
            .cmdGrpAuthorise1.Enabled(0) = False
            .cmdGrpAuthorise1.Enabled(1) = False
            .cmdGrpAuthorise1.Enabled(2) = True
            .txtrecloc.Enabled = True : .ctlitem_c1.Enabled = True : ctlitem_c1.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : cmdHelp.Enabled = True
            .ctlitem_c1.Text = String.Empty : .lblRecLoc.Text = String.Empty
            .txtdesloc.Text = String.Empty : .lbldescLoc.Text = String.Empty
            .txtgrtype.Text = String.Empty : .txtInwRegNo.Text = String.Empty : .lblAuthSign.Text = String.Empty
            .txtInvValue.Text = String.Empty
            .txtgrdt.Text = String.Empty
            .dtpQADate.Enabled = True : .dtpQADate.Format = DateTimePickerFormat.Custom : .dtpQADate.CustomFormat = gstrDateFormat
            .dtpQADate.MinDate = GetServerDate() : .dtpQADate.MaxDate = GetServerDate() : .dtpQADate.Value = GetServerDate()
            .txtEmpQA.Enabled = True : .txtEmpQA.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : .CmdHelpEmpQA.Enabled = True : .txtEmpQA.Text = String.Empty
            .txtvend.Text = String.Empty
            .txtvencd.Text = String.Empty : .txtpo.Text = String.Empty : .txtinv.Text = String.Empty : .txtinvdt.Text = String.Empty
            .lblEmpCodeDesc.Text = String.Empty : .txtRefGRINNo.Text = String.Empty : .txtRefGRINDate.Text = String.Empty : .txtRemarks.Text = String.Empty
        End With
        With sprdata
            .MaxRows = 1
            .Row = 1
            .Col = Col_Item_UOM
            .Action = FPSpreadADO.ActionConstants.ActionActiveCell
            .EditModePermanent = True
            .EditModeReplace = True
        End With
        '*10857384
        ListESIDetail.Clear()
        '*

        txtEasyReqNo.Text = ""
        txtReqApprovalStatus.Text = ""
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmQATRN0001_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        On Error GoTo ErrHandler
        Dim enmValue As eMPowerFunctions.ConfirmWindowReturnEnum
        If UnloadMode >= 0 And UnloadMode <= 5 Then
            If Me.cmdGrpAuthorise1.Button = UCActXCtl.clsDeclares.ButtonEnum.BUTTON_AUTHORIZE Then
                enmValue = ConfirmWindow(10064, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION)
                If enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                    Cancel = 1
                Else
                    Cancel = 0
                End If
            End If
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub frmQATRN0001_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.Escape Then
            Call cmdGrpAuthorise1_ButtonClick(cmdGrpAuthorise1, New UCActXCtl.cmdGrpAuthorise.ButtonClickEventArgs(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE))
        ElseIf KeyCode = System.Windows.Forms.Keys.Return Then  'Enter key pressed
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmQATRN0001_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo ErrHandler
        mdifrmMain.CheckFormName = mintIndex
        frmModules.NodeFontBold(Me.Tag) = True
        optGRIN.Focus()
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub display_details()
        '***************************************************************************
        'Revision Date    - 09-Oct-2006
        'Revised By       - Davinder Singh
        'Issue No         - 17893
        'Revision History - Changes made to hadle the case of Grin against Multiple PO(s)
        'Description      - To populate the data on the form
        '***************************************************************************
        On Error GoTo ErrHandler
        Dim oRS As ADODB.Recordset
        Dim strSql As String
        Dim strDocCategory As String = String.Empty
        '*10857384
        ListESIDetail.Clear()
        '*

        strSql = " SELECT * FROM DBO.UFN_GET_GRIN_HDR_DATA(" & _
               "'" & gstrUNITID & "','" & txtrecloc.Text.Trim & "'," & _
               ctlitem_c1.Text & ")"
        oRS = mP_Connection.Execute(strSql)
        If Not (oRS.BOF And oRS.EOF) Then
            mstrGRType.Value = oRS.Fields("Doc_category").Value
            mstrPOstatus = oRS.Fields("PO_Status").Value
            mstrChallanNo = oRS.Fields("Challan_No").Value
            mstrCurrencyID = oRS.Fields("Org_Category").Value
            mstrGateEntryID = oRS.Fields("Gate_no").Value
            txtrecloc.Text = oRS.Fields("from_location").Value
            txtEmpQA.Text = oRS.Fields("inspection_auth").Value
            lblEmpCodeDesc.Text = oRS.Fields("INSPECTION_NAME").Value
            txtInwRegNo.Text = oRS.Fields("gate_no").Value
            txtInvValue.Text = oRS.Fields("bill_amount").Value
            lblRecLoc.Text = oRS.Fields("FROM_LOC_DESC").Value
            txtdesloc.Text = oRS.Fields("to_location").Value
            lbldescLoc.Text = oRS.Fields("TO_LOC_DESC").Value
            lblAuthSign.Text = oRS.Fields("AUTH_NAME").Value
            txtRefGRINNo.Text = IIf(oRS.Fields("ref_grin_no").Value = 0, "", oRS.Fields("ref_grin_no").Value)
            If txtRefGRINNo.Text.Trim.Length > 0 Then
                txtRefGRINDate.Text = Format(oRS.Fields("REF_GRIN_DATE").Value, gstrDateFormat)
            End If
            txtgrdt.Text = Format(oRS.Fields("grn_date").Value, gstrDateFormat)
            txtgrtype.Text = oRS.Fields("doc_category").Value
            strDocCategory = txtgrtype.Text.Trim.ToUpper
            Call txtgrtype_Validate(False)
            txtvencd.Text = oRS.Fields("vendor_code").Value
            txtvend.Text = oRS.Fields("vendor_NAME").Value
            txtinv.Text = oRS.Fields("invoice_no").Value
            txtinvdt.Text = Format(oRS.Fields("invoice_date").Value, gstrDateFormat)
            txtRemarks.Text = oRS.Fields("remarks").Value
        End If
        oRS.Close()
        oRS = Nothing
        ''---- Fill data in Array if Grin is not Without PO
        If mstrPOstatus <> "W" Then
            txtpo.Text = GetPOList()
            Call FillArrayWithPOdtls()
        Else
            txtpo.Clear()
        End If
        ''---- Fill data in grid

        '*10857384
        With sprdata
            .Row = 0
            .Col = Col_ESI
            If strDocCategory = "W" Then
                .ColHidden = False
            Else
                .ColHidden = True
            End If
        End With
        '*
        '101161852
        CheckIQARequired()
        If mblnIQARequired Then
            Dim dtIQAMainGrp As DataTable = SqlConnectionclass.GetDataTable("SELECT DISTINCT DESCR FROM LISTS (NOLOCK) WHERE KEY1='IQA' AND KEY2='ITEM_MAIN_GRP' AND UNIT_CODE='" & gstrUnitId & "'")
            If dtIQAMainGrp IsNot Nothing AndAlso dtIQAMainGrp.Rows.Count > 0 Then
                For Each dr As DataRow In dtIQAMainGrp.Rows
                    sprdata.Row = 0
                    sprdata.Col = Col_IQA
                    If Convert.ToString(dr("DESCR")).ToUpper() = strDocCategory.ToUpper() Then
                        mblnIQARequired = True
                        sprdata.ColHidden = False
                        Exit For
                    Else
                        mblnIQARequired = False
                        sprdata.ColHidden = True
                    End If
                Next
                dtIQAMainGrp.Dispose()
            End If
        End If
        If mblnIQARequired Then
            sprdata.Row = 0
            sprdata.Col = Col_Rejection_Details
            sprdata.ColHidden = True
        Else
            sprdata.Row = 0
            sprdata.Col = Col_Rejection_Details
            sprdata.ColHidden = False
        End If

        Call LoadData()
        'VID CHANGES
        GetVendorInvoiceDateOlderThanPODateApprovalDetail()
        Exit Sub
ErrHandler:
        oRS = Nothing
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub LoadData()
        On Error GoTo ErrHandler
        Dim strsql As String
        Dim oRS As ADODB.Recordset
        Dim intUOMdecimalPlaces As Short
        Dim strItemCode As String
        Dim strMeasureCode As String
        Dim intCounter As Integer
        Dim oRSDocExist As New ClsResultSetDB

        strsql = "SELECT * FROM DBO.UFN_GET_GRIN_DTL_DATA(" & _
                 "'" & gstrUNITID & "', '" & txtrecloc.Text.Trim & "'," & _
                 ctlitem_c1.Text & ") ORDER BY SERIAL_NO"
        oRS = mP_Connection.Execute(strsql)
        If Not (oRS.BOF And oRS.EOF) Then
            With sprdata
                .MaxRows = 0
                .Enabled = True
                .CursorStyle = FPSpreadADO.CursorStyleConstants.CursorStyleArrow
                .SetText(Col_ViewDoc, 0, "Item Doc")

                '*10857384
                .SetText(Col_ESI, 0, "ESI Detail")
                '*

                While Not oRS.EOF
                    .MaxRows = .MaxRows + 1
                    intCounter = .MaxRows
                    .set_RowHeight(.MaxRows, 300)
                    .Row = intCounter
                    .Col = Col_Item_Code
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                    .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
                    .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
                    strItemCode = oRS.Fields("item_code").Value
                    .Text = strItemCode
                    .Row = intCounter
                    .Col = Col_Item_Description
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                    .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
                    .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
                    .Text = Replace(oRS.Fields("DESCRIPTION").Value & " [" & oRS.Fields("MATERIALSPEC").Value & "]", "[]", String.Empty)
                    .Row = intCounter
                    .Col = Col_Item_UOM
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                    .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
                    .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
                    strMeasureCode = oRS.Fields("UOM").Value
                    mstrMeasureCode = strMeasureCode
                    .Text = strMeasureCode
                    intUOMdecimalPlaces = GetUOMDecimalPlacesAllowed(strMeasureCode)
                    .Row = intCounter
                    .Col = Col_Item_Rate
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                    .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                    .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
                    .Text = Format(oRS.Fields("ITEM_RATE").Value, "#0.0000")
                    .Row = intCounter
                    .Col = Col_Challan_Qty
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                    .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                    .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
                    If intUOMdecimalPlaces > 0 Then
                        .Text = Format(oRS.Fields("challan_quantity").Value, "#0.0000")
                    Else
                        .Text = Format(oRS.Fields("challan_quantity").Value, "#0")
                    End If
                    .Row = intCounter
                    .Col = Col_Actual_Qty
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                    .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                    .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
                    If intUOMdecimalPlaces > 0 Then
                        .Text = Format(oRS.Fields("actual_quantity").Value, "#0.0000")
                    Else
                        .Text = Format(oRS.Fields("actual_quantity").Value, "#0")
                    End If
                    .Row = intCounter
                    .Col = Col_Excess_PO_Qty
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                    .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                    .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
                    If intUOMdecimalPlaces > 0 Then
                        .Text = Format(oRS.Fields("Excess_PO_Quantity").Value, "#0.0000")
                    Else
                        .Text = Format(oRS.Fields("Excess_PO_Quantity").Value, "#0")
                    End If
                    .Row = intCounter
                    .Col = Col_Rejected_Qty
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                    .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                    .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
                    If intUOMdecimalPlaces > 0 Then
                        .Text = Format(oRS.Fields("actual_quantity").Value, "#0.0000")
                    Else
                        .Text = Format(oRS.Fields("actual_quantity").Value, "#0")
                    End If
                    .Row = intCounter
                    .Col = Col_Receipt_Qty
                    .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                    .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
                    If intUOMdecimalPlaces > 0 Then
                        .Text = Format(oRS.Fields("receipt_quantity").Value, "#0.0000")
                    Else
                        .Text = Format(oRS.Fields("receipt_quantity").Value, "#0")
                    End If
                    If intUOMdecimalPlaces > 0 Then
                        .Row = intCounter
                        .Col = Col_Accepted_Qty
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                        .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                        .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
                        .TypeFloatMin = 0.0#
                        .TypeFloatMax = 99999999.9999
                        .TypeFloatDecimalPlaces = 4
                        .Text = Format(0, "#0.0000")
                    Else
                        .Row = intCounter
                        .Col = Col_Accepted_Qty
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeInteger
                        .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                        .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
                        .TypeIntegerMin = 0
                        .TypeIntegerMax = 99999999
                        .Text = Format(0, "#0")
                    End If
                    .Row = intCounter
                    .Col = Col_Rejection_Reason
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                    .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
                    .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
                    .TypeMaxEditLen = 50
                    .Text = oRS.Fields("remarks").Value
                    .Row = intCounter
                    .Col = Col_Asseccible_Rate
                    .Text = oRS.Fields("Accessible_rate").Value
                    .Row = intCounter
                    .Col = Col_GL_Group
                    .Text = oRS.Fields("GL_Group_code").Value
                    .Row = intCounter
                    .Col = Col_Project_Code
                    .Text = oRS.Fields("Project_code").Value
                    .Row = intCounter
                    .Col = Col_Discount_Per
                    .Text = oRS.Fields("Discount_per").Value
                    .Row = intCounter
                    .Col = Col_Rejection_Details
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                    .TypeButtonText = "Rejection Details"
                    .Row = intCounter
                    .Col = Col_Inspection_Entry
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
                    .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
                    .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter

                    If mblnBatchTracking = True Then
                        .Row = intCounter
                        .Col = Col_Batch_Details
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                        .TypeButtonText = "Batch Details"
                    End If

                    If mbln_Inspection_Control_Details And optGRIN.Checked = True Then
                        .Row = 0
                        .Col = col_Inspection_Control_Details
                        .Text = "Inspection Details"
                        .ColHidden = False
                        .Row = intCounter
                        .Col = col_Inspection_Control_Details
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                        .TypeButtonText = "Inspection Details"
                        .Row = 0
                        .Row2 = .MaxRows
                        .Col = Col_Inspection_Entry
                        .Col2 = Col_Inspection_Entry
                        .ColHidden = False
                    Else
                        .Row = 0
                        .Row2 = .MaxRows
                        .Col = Col_Inspection_Entry
                        .Col2 = Col_Inspection_Entry
                        .ColHidden = True
                    End If

                    ''Issue id-10384524

                    .Row = intCounter
                    .Col = Col_ViewDoc
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                    .TypeButtonText = "View Doc"

                    If mblnIQARequired Then
                        oRSDocExist.GetResult("select top 1 1 from Item_MST where unit_code = '" & gstrUNITID & "' and item_code = '" & strItemCode.Trim & "'  and (ITEM_MAIN_GRP IN ('R','C')  OR (ITEM_MAIN_GRP IN ('S','F') AND  Source in ('B','J') AND TYPE_CODE='B'))")
                        If oRSDocExist.GetNoRows > 0 Then   'Praveen On 2 FEB 2021
                            .Row = intCounter
                            .Col = Col_IQA
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                            .TypeButtonText = "IQA Required"
                        Else
                            .Row = intCounter
                            .Col = Col_IQA
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                        End If
                    End If

                    oRSDocExist.GetResult("select top 1 1 from ItemDocs where unit_code = '" & gstrUNITID & "' and item_code = '" & strItemCode.Trim & "'")

                    If oRSDocExist.GetNoRows > 0 Then
                        .BlockMode = True
                        .Row = intCounter
                        .Row2 = intCounter
                        .Col = Col_ViewDoc
                        .Col2 = Col_ViewDoc
                        .TypeButtonColor = Color.LightBlue

                        .Col = Col_Rejection_Details
                        .Col2 = Col_Batch_Details
                        .TypeButtonColor = Me.BackColor
                        .BlockMode = False

                        ' oRSDocExist.ResultSetClose()
                    End If

                    '*10857384
                    .Row = intCounter
                    .Col = Col_ESI
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                    .TypeButtonText = "ESI Detail"
                    '*
                    oRS.MoveNext()


                End While

                If mbln_Inspection_Control_Details = False Then
                    .Col = col_Inspection_Control_Details
                    .ColHidden = True
                End If

                .Row = 1
                .Col = Col_Item_UOM
                .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                .EditModePermanent = True
                .EditModeReplace = True
            End With
            Call SetInspectionEvironment()
            cmdGrpAuthorise1.Enabled(0) = True
        Else
            MsgBox("No Items to be displayed in the GRID.", MsgBoxStyle.Information, ResolveResString(100))
            ctlitem_c1.Focus()
        End If
        oRS.Close()
        oRS = Nothing

        If Not IsNothing(oRSDocExist) Then
            oRSDocExist = Nothing
        End If

        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        oRS = Nothing
    End Sub
    Private Sub frmQATRN0001_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error GoTo ErrHandler
        Me.Dispose()
        mdifrmMain.RemoveFormNameFromWindowList = mintIndex
        frmModules.NodeFontBold(Me.Tag) = False
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        gblnCancelUnload = True
    End Sub
    Private Sub fsPOWiseRejectionDtls_EditChange(ByVal eventSender As System.Object, ByVal eventArgs As AxFPSpreadADO._DSpreadEvents_EditChangeEvent) Handles fsPOWiseRejectionDtls.EditChange
        '***************************************************************************
        'Creation Date  - 07-Oct-2006
        'Created By     - Davinder Singh
        'Issue No       - 17893
        'Description    - To refresh the quantities in the labels and frame when user change the quantities
        '                 in grid
        '***************************************************************************
        On Error GoTo ErrHandler
        Dim curActualQty As Double
        Dim curRejectedQty As Double
        Dim curTotalRejectedQty As Double
        Dim intCtr As Short
        With fsPOWiseRejectionDtls
            Select Case eventArgs.col
                Case Enm_fsPOWiseRejectionDtls.Col_Rejected_Quantity
                    .Row = eventArgs.row
                    .Col = Enm_fsPOWiseRejectionDtls.Col_Rejected_Quantity
                    curRejectedQty = Val(.Text)
                    .Col = Enm_fsPOWiseRejectionDtls.Col_Actual_Quantity
                    curActualQty = Val(.Text)
                    .Col = Enm_fsPOWiseRejectionDtls.Col_Accepted_Quantity
                    .Text = CStr(curActualQty - curRejectedQty)
                    curTotalRejectedQty = 0
                    For intCtr = 1 To .MaxRows Step 1
                        .Row = intCtr
                        .Col = Enm_fsPOWiseRejectionDtls.Col_Rejected_Quantity
                        curTotalRejectedQty = curTotalRejectedQty + Val(.Text)
                    Next intCtr
                    If InStr(1, lblRejectedAgainstPO.Text, ".", CompareMethod.Text) <> 0 Then
                        lblRejectedAgainstPO.Text = VB6.Format(curTotalRejectedQty, "0.0000")
                        lblRemainingRejectedQty.Text = VB6.Format(CDbl(lblTotalRejectedQty.Text) - CDbl(lblRejectedAgainstPO.Text), "0.0000")
                    Else
                        lblRejectedAgainstPO.Text = VB6.Format(curTotalRejectedQty, "0")
                        lblRemainingRejectedQty.Text = VB6.Format(CDbl(lblTotalRejectedQty.Text) - CDbl(lblRejectedAgainstPO.Text), "0")
                    End If
            End Select
        End With
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub optGRIN_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optGRIN.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo ErrHandler
            Dim strsql As String
            Call RefreshFrm()
            With sprdata
                .Row = 0
                .Col = Col_Challan_Qty
                .ColHidden = False
                .Row = 0
                .Col = Col_Actual_Qty
                .Text = "Actual Qty."
            End With
            If mbln_Inspection_Control_Details = True Then
                With sprdata
                    .Row = 0
                    .Col = col_Inspection_Control_Details : .Text = "Inspection Details"
                    .ColHidden = False
                    .Row = 1
                    .Row2 = .MaxRows
                    .Col = col_Inspection_Control_Details
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                    .TypeButtonText = "Inspection Details"
                    .Row = 0
                    .Row2 = .MaxRows
                    .Col = Col_Inspection_Entry
                    .Col2 = Col_Inspection_Entry
                    .ColHidden = False
                End With
            Else
                With sprdata
                    .Row = 0
                    .Row2 = .MaxRows
                    .Col = Col_Inspection_Entry
                    .Col2 = Col_Inspection_Entry
                    .ColHidden = True
                End With
            End If
            txtrecloc.Enabled = True
            txtrecloc.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            ctlitem_c1.Enabled = True
            strsql = "select LOCATION_CODE from location_mst where UNIT_CODE = '" & gstrUNITID & "' AND" & _
                " upper(loc_type)='R'"
            txtrecloc.Text = GetQryOutput(strsql)
            strsql = "select description from location_mst where UNIT_CODE = '" & gstrUNITID & "' AND" & _
                " LOCATION_CODE = '" & txtrecloc.Text.Trim & "'"
            lblRecLoc.Text = GetQryOutput(strsql)
            ctlitem_c1.Focus()
            Exit Sub
ErrHandler:
            gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub
    Private Sub optGRIN_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles optGRIN.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If KeyAscii = System.Windows.Forms.Keys.Return Then
            If Len(Trim(Me.txtrecloc.Text)) > 0 Then
                ctlitem_c1.Focus() : GoTo EventExitSub
            Else
                txtrecloc.Enabled = True : txtrecloc.Focus()
            End If
        End If
        GoTo EventExitSub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub optSuppGRIN_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optSuppGRIN.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo ErrHandler
            Dim strsql As String
            Call RefreshFrm()
            If mbln_Inspection_Control_Details = True Then
                With Me.sprdata
                    .Row = 0
                    .Col = col_Inspection_Control_Details : .Text = "Inspection Details"
                    .ColHidden = True
                    .Row = 1
                    .Row2 = .MaxRows
                    .Col = col_Inspection_Control_Details
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                    .TypeButtonText = "Inspection Details"
                    .Row = 0
                    .Row2 = .MaxRows
                    .Col = Col_Inspection_Entry
                    .Col2 = Col_Inspection_Entry
                    .ColHidden = False
                End With
            Else
                With Me.sprdata
                    .Row = 0
                    .Row2 = .MaxRows
                    .Col = Col_Inspection_Entry
                    .Col2 = Col_Inspection_Entry
                    .ColHidden = True
                End With
            End If
            With Me.sprdata
                .Row = 0
                .Col = Col_Challan_Qty
                .ColHidden = True
                .Row = 0
                .Col = Col_Actual_Qty : .Text = "ReInspection Qty."
            End With
            strsql = "select location_code from location_mst where UNIT_CODE = '" & gstrUNITID & "' AND" & _
                " upper(loc_type)='J'"
            txtrecloc.Text = GetQryOutput(strsql)
            strsql = "select description from location_mst where UNIT_CODE = '" & gstrUNITID & "' AND" & _
                " LOCATION_CODE = '" & txtrecloc.Text.Trim & "'"
            lblRecLoc.Text = GetQryOutput(strsql)
            ctlitem_c1.Focus()
            Exit Sub
ErrHandler:
            gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub
    Private Sub optSuppGRIN_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles optSuppGRIN.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If KeyAscii = System.Windows.Forms.Keys.Return Then
            Me.txtrecloc.Enabled = True : Me.txtrecloc.Focus()
        End If
        GoTo EventExitSub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtDocNo_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtDocNo.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '********************************************************************************************
        'Created By     -  Davinder Singh
        'Date           -  09-Oct-2006
        'Arguments      -  Keycode of the key pressed
        'Return Value   -  NIL
        'Function       -  To display the Doc No. help
        '***********************************************************************************
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then
            cmdDocHlp.PerformClick()
        End If
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtDocNo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtDocNo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '********************************************************************************************
        'Created By     -  Davinder Singh
        'Date           -  09-Oct-2006
        'Arguments      -  Keycode of the key pressed
        'Return Value   -  NIL
        'Function       -  To display the Doc No. help
        '***********************************************************************************
        On Error GoTo ErrHandler
        If KeyAscii = System.Windows.Forms.Keys.Return Then
            cmdPrint.PerformClick()
        End If
        GoTo EventExitSub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtEmpQA_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtEmpQA.TextChanged
        On Error GoTo ErrHandler
        lblEmpCodeDesc.Text = String.Empty
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtEmpQA_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtEmpQA.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '----------------------------------------------------------
        'Call Validate Event of Emp Qa  - Nitin
        '----------------------------------------------------------
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                txtEmpQA_Validating(txtEmpQA, New System.ComponentModel.CancelEventArgs(False)) 'Validate Employee Code
            Case 34, 39, 96
                KeyAscii = 0 'Disable ',",`
        End Select
        GoTo EventExitSub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtEmpQA_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtEmpQA.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If (KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0) Then Call CmdHelpEmpQA_Click(CmdHelpEmpQA, New System.EventArgs()) : Exit Sub
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtEmpQA_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtEmpQA.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        Dim strSql As String
        If txtEmpQA.Text.Trim.Length > 0 Then
            strSql = "SELECT fullname from user_mst where UNIT_CODE = '" & gstrUNITID & "' AND" & _
                " user_id='" & txtEmpQA.Text.Trim & "'and InActive='N'"
            '" user_id='" & txtEmpQA.Text.Trim & "' "  '102278094
            lblEmpCodeDesc.Text = GetQryOutput(strSql)
            If lblEmpCodeDesc.Text.Trim.Length > 0 Then
                dtpQADate.Focus()
            Else
                Call ConfirmWindow(10333)
                txtEmpQA.Text = String.Empty
                txtEmpQA.Focus()
                Cancel = True
            End If
        Else
            dtpQADate.Focus()
        End If
        GoTo EventExitSub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtgrtype_Validate(ByRef Cancel As Boolean)
        On Error GoTo ErrHandler
        Dim strsql As String
        mstrGrinType = txtgrtype.Text.Trim
        strsql = "SELECT Description FROM type_mst" & _
                " WHERE UNIT_CODE = '" & gstrUNITID & "' AND type_code = '" & txtgrtype.Text & "'"
        txtgrtype.Text = GetQryOutput(strsql)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub grnhdr(ByVal flag As Boolean)
        On Error GoTo ErrHandler
        txtrecloc.Enabled = flag
        txtdesloc.Enabled = flag
        txtgrdt.Enabled = flag
        txtpo.Enabled = flag
        txtvencd.Enabled = flag
        txtinv.Enabled = flag
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Function CheckForvalidAcceptance(ByVal strDocNo As String) As Integer
        'Added By   : Shabbir Hussain
        'On Date    : 24-Oct-2005
        'Purpose    : To check whether accepted and rejected Quantity of Grin is zero
        'Input      : Grin No
        'Output     : returns -1 If  Accepted and rejected quantity is 0 and actual quantity is non zero
        '             Otherwise returns 0.
        On Error GoTo ErrHandler
        Dim rstDB As ClsResultSetDB
        Dim strsql As String
        CheckForvalidAcceptance = 0
        strsql = "SELECT * FROM Grn_Dtl"
        strsql = strsql & " WHERE UNIT_CODE = '" & gstrUNITID & "' AND Doc_No='" & strDocNo & "'"
        strsql = strsql & " AND Accepted_Quantity=0"
        strsql = strsql & " AND Rejected_Quantity=0"
        strsql = strsql & " AND Actual_Quantity<>0"
        rstDB = New ClsResultSetDB
        If rstDB.GetResult(strsql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly) = False Then GoTo ErrHandler
        If rstDB.GetNoRows > 0 Then
            CheckForvalidAcceptance = -1
        End If
        rstDB.ResultSetClose()
        rstDB = Nothing
        Exit Function
ErrHandler:
        CheckForvalidAcceptance = -1
        rstDB = Nothing
    End Function
    Private Function validate_data(ByRef pstrItemCode As String, ByRef plngActualQty As Double, ByRef pdblExcessPOQty As Double, ByRef plngacceptedqty As Double, ByRef plngRejected_Quantity As Double, ByRef plngreaforrej As String) As Boolean
        '**************************************************************************************************************************************
        'Revision Date  - 13-Aug-2007
        'Revised By     - Davinder Singh
        'Issue No       - 20852
        'History        - A new Stored Procedure PRC_UPDATE_CUSTOMER_SUPPLIED_MATERIAL
        '                 is called for solving Primary key error occured in case of
        '                 Customer Supplied type material(Z) type of Grin.
        '*********************************************************************************************************************
        ' Revised By                 -   Davinder Singh
        ' Revision Date              -   02 FEB 2009
        ' Issue Id                   -   eMpro-20090202-26873 
        ' Revision History           -   Object not closed error Occuring. To resolved this error  
        '                                Values fetched from recordset oRS and closed it before using
        '                                command object CMD    
        '*********************************************************************************************************************
        On Error GoTo ErrHandler
        Dim strsql As String
        Dim oRS As ADODB.Recordset
        Dim strRejectReason As String
        Dim intCtr As Short
        Dim lngPONO As Integer
        Dim curAcceptedQty As Double
        Dim curRejectedQty As Double
        Dim curActualQty As Double
        Dim curExcessPOQty As Double
        Dim CMD As ADODB.Command
        Dim strGrnType As String
        Dim strCustCode As String
        Dim strInvNo As String
        Dim strInvDt As String
        Dim strPOStatus As String = ""
        Dim strType As String = ""
        validate_data = True
        strCustCode = String.Empty
        strInvNo = String.Empty
        strInvDt = String.Empty
        strsql = "SELECT doc_category,vendor_code,invoice_no,invoice_date,PO_STATUS" & _
                " FROM DBO.UFN_GET_GRIN_HDR_DATA(" & _
                " '" & gstrUNITID & "','" & txtrecloc.Text.Trim & "'," & _
                 ctlitem_c1.Text & ")"
        oRS = mP_Connection.Execute(strsql)
        strGrnType = UCase(oRS.Fields("Doc_category").Value.ToString)
        If strGrnType = "Z" Then
            strCustCode = oRS.Fields("vendor_code").Value
            strInvNo = oRS.Fields("invoice_no").Value
            strInvDt = Format(oRS.Fields("invoice_date").Value, "dd MMM yyyy")
            strPOStatus = oRS.Fields("PO_STATUS").Value
        End If
        oRS.Close()
        oRS = Nothing
        ''---- In case of Customer supplied material update stock at Customer Location
        If strGrnType = "Z" Then
            ' IF GRIN IS MADE AGAINST PO THEN @TYPE CODE IS 'V' ELSE 'C'
            If strPOStatus = "A" Then
                strType = "V"
            Else
                strType = "C"
            End If
            CMD = New ADODB.Command
            With CMD
                .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                .CommandText = "PRC_UPDATE_CUSTOMER_SUPPLIED_MATERIAL"
                .ActiveConnection = mP_Connection
                .Parameters.Append(.CreateParameter("@UNITCODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                .Parameters.Append(.CreateParameter("@POSTATUS", ADODB.DataTypeEnum.adChar, ADODB.ParameterDirectionEnum.adParamInput, 1, strType))
                .Parameters.Append(.CreateParameter("@CUSTCODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 8, strCustCode))
                .Parameters.Append(.CreateParameter("@INVNO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 50, strInvNo))
                .Parameters.Append(.CreateParameter("@INVDATE", ADODB.DataTypeEnum.adDBTimeStamp, ADODB.ParameterDirectionEnum.adParamInput, ADODB.DataTypeEnum.adDBTimeStamp, strInvDt))
                .Parameters.Append(.CreateParameter("@ITEMCODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 16, pstrItemCode))
                .Parameters.Append(.CreateParameter("@QTY", ADODB.DataTypeEnum.adCurrency, ADODB.ParameterDirectionEnum.adParamInput, ADODB.DataTypeEnum.adCurrency, plngacceptedqty))
                .Parameters.Append(.CreateParameter("@MP_USER", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 16, mP_User))
                .Parameters.Append(.CreateParameter("@MSG", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamOutput, 100))
                .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End With
            If Not CMD.Parameters(CMD.Parameters.Count - 1).Value Is System.DBNull.Value Then
                MsgBox(CMD.Parameters(CMD.Parameters.Count - 1).Value, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                validate_data = False
                CMD = Nothing
                Exit Function
            End If
            CMD = Nothing
            'FOR DIFF IN RATE VALIDATION FOR GRN MADE AGAINST PO
            If strPOStatus = "A" Then
                If DiffInRates(pstrItemCode) = True Then
                    MsgBox("As their is Difference in Invoice Rate and PO Rate,Please amend the Purchase Order To Authorize the GRIN!." & vbCrLf & "Grin can not be authorized.", MsgBoxStyle.Information, ResolveResString(100))
                    validate_data = False
                    Exit Function
                Else
                    CMD = New ADODB.Command
                    With CMD
                        .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                        .CommandText = "USP_UPDATE_CSIRATE"
                        .ActiveConnection = mP_Connection
                        .Parameters.Append(.CreateParameter("@UNITCODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                        .Parameters.Append(.CreateParameter("@GRN_NO", ADODB.DataTypeEnum.adBigInt, ADODB.ParameterDirectionEnum.adParamInput, , Val(Me.ctlitem_c1.Text.Trim)))
                        .Parameters.Append(.CreateParameter("@ITEM_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 16, pstrItemCode.Trim))
                        .Parameters.Append(.CreateParameter("@USER_ID", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 16, mP_User))
                        .Parameters.Append(.CreateParameter("@ERR_CODE", ADODB.DataTypeEnum.adBigInt, ADODB.ParameterDirectionEnum.adParamOutput))
                        .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End With
                    If CMD.Parameters(CMD.Parameters.Count - 1).Value <> 0 Then
                        MsgBox("Error in updating Item Rate in CustomerSuppliedItem_Mst", MsgBoxStyle.Critical, ResolveResString(100))
                        validate_data = False
                        CMD = Nothing
                        Exit Function
                    End If
                    CMD = Nothing
                End If
            End If
            'End of addition by Vinod
        End If
        ''---- If Grin is against PO then update the associated tables
        If mstrPOstatus <> "W" Then
            For intCtr = 1 To UBound(mArrPOdtl) Step 1
                If StrComp(pstrItemCode, mArrPOdtl(intCtr).ItemCode, CompareMethod.Text) = 0 Then
                    lngPONO = mArrPOdtl(intCtr).PONO
                    curAcceptedQty = mArrPOdtl(intCtr).AcceptedQty
                    curActualQty = mArrPOdtl(intCtr).ActualQty
                    curExcessPOQty = mArrPOdtl(intCtr).ExcessPOQty
                    curRejectedQty = mArrPOdtl(intCtr).RejectedQty
                    ''---- Update Accepted and Rejected Qty in Grn_PO_Dtl table
                    strsql = "Update Grn_PO_Dtl" & _
                            " SET Accepted_Qty=" & curAcceptedQty & "," & _
                            " Rejected_Qty=" & curRejectedQty & _
                            " ,Upd_dt=getdate(),Upd_Userid='" & mP_User & "'" & _
                            " WHERE UNIT_CODE = '" & gstrUNITID & "' AND Doc_NO ='" & ctlitem_c1.Text & "'" & _
                            " AND Item_Code='" & pstrItemCode & "'" & _
                            " AND PO_No=" & lngPONO
                    mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    ''---- Update Grn_PO_Dtl for Supplementary grin
                    If optSuppGRIN.Checked = True Then
                        strsql = "UPDATE GRN_PO_Dtl" & _
                                " SET Inspected_Qty = ISNULL(Inspected_Qty,0) - " & curActualQty & "+" & curAcceptedQty & _
                                " ,Upd_dt=getdate(),Upd_Userid='" & mP_User.Trim & "'" & _
                                " WHERE UNIT_CODE = '" & gstrUNITID & "' AND Doc_no=" & txtRefGRINNo.Text.Trim & _
                                " AND Item_code='" & pstrItemCode & "'" & _
                                " AND PO_No=" & lngPONO
                        mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End If
                End If
            Next intCtr
        End If
        ''---- Update Grn_Hdr,Grn_Dtl
        strRejectReason = plngreaforrej
        Call QuoteString(strRejectReason)
        strsql = "UPDATE Grn_dtl" & _
                " SET Accepted_Quantity = " & plngacceptedqty & _
                ",Rejected_Quantity = " & plngRejected_Quantity & _
                ",Remarks = '" & strRejectReason & "'" & _
                ",Upd_dt=getdate(),Upd_Userid='" & mP_User & "'" & _
                " WHERE UNIT_CODE = '" & gstrUNITID & "' AND Doc_no = '" & ctlitem_c1.Text & "'" & _
                " AND item_code ='" & pstrItemCode & "'" & _
                " AND from_location='" & txtrecloc.Text.Trim & "'"
        mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

        ''---- Update Grn_Dtl for Supplementary grin
        If optSuppGRIN.Checked = True Then
            strsql = "UPDATE GRN_Dtl" & _
                    " SET Inspected_quantity = ISNULL(inspected_quantity,0) - " & plngActualQty & "+" & plngacceptedqty & _
                    " ,Upd_dt=getdate(),Upd_Userid='" & mP_User & "'" & _
                    " WHERE UNIT_CODE = '" & gstrUNITID & "' AND Doc_type=10" & _
                    " AND From_location='01R1'" & _
                    " AND Doc_no=" & txtRefGRINNo.Text.Trim & _
                    " AND Item_code='" & pstrItemCode & "'"
            mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        End If
        If strGrnType = "Z" Then
            CMD = New ADODB.Command
            With CMD
                .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                .CommandText = "USP_INSERT_CSIGRIN_FOR_INVOICE"
                .ActiveConnection = mP_Connection
                .Parameters.Append(.CreateParameter("@UNITCODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                .Parameters.Append(.CreateParameter("@GRN_NO", ADODB.DataTypeEnum.adBigInt, ADODB.ParameterDirectionEnum.adParamInput, , Val(Me.ctlitem_c1.Text.Trim)))
                .Parameters.Append(.CreateParameter("@ITEM_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 16, pstrItemCode.Trim))
                .Parameters.Append(.CreateParameter("@USER_ID", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 16, mP_User))
                .Parameters.Append(.CreateParameter("@ERR_CODE", ADODB.DataTypeEnum.adBigInt, ADODB.ParameterDirectionEnum.adParamOutput))
                .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End With
            If CMD.Parameters(CMD.Parameters.Count - 1).Value <> 0 Then
                MsgBox("Error in Updating CSM Data For Invoice", MsgBoxStyle.Critical, ResolveResString(100))
                validate_data = False
                CMD = Nothing
                Exit Function
            End If
            CMD = Nothing
        End If
        strsql = "Exec st_proc '" & gstrUNITID & "', '" & txtrecloc.Text.Trim & "','" & txtdesloc.Text.Trim & "','" & pstrItemCode & "'," & plngacceptedqty & ",0,'N', 'Y','" & mP_User & "'"
        mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        ''---- If Qty Received in Excess or Rejected during inspection
        strsql = "Exec st_proc '" & gstrUNITID & "', '" & txtrecloc.Text.Trim & "', '01J1', '" & pstrItemCode & "', " & plngRejected_Quantity + pdblExcessPOQty & ",0,'N', 'Y','" & Trim(mP_User) & "'"
        If (optGRIN.Checked = True And (plngRejected_Quantity + pdblExcessPOQty) > 0) Then mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        ''---- For supplementary Grin
        strsql = "EXEC st_proc '" & gstrUNITID & "', '" & txtrecloc.Text.Trim & "','01M1','" & pstrItemCode & "'," & plngacceptedqty & ",0,'Y','N','" & mP_User & "'"
        If optSuppGRIN.Checked = True Then mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        Exit Function
ErrHandler:
        validate_data = False
        oRS = Nothing
        MsgBox(Err.Description & " - VD [ " & Err.Number & " ] ", MsgBoxStyle.Information, ResolveResString(100))
    End Function
    Private Sub txtRecLoc_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtrecloc.TextChanged
        On Error GoTo ErrHandler
        txtrecloc.Text = Replace(txtrecloc.Text, "'", String.Empty)
        lblRecLoc.Text = String.Empty
        If txtrecloc.Text.Trim.Length = 0 Then
            Call RefreshFrm()
            If txtrecloc.Enabled = True Then txtrecloc.Focus()
        End If
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtRecLoc_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtrecloc.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If KeyAscii >= 97 And KeyAscii <= 122 Then KeyAscii = KeyAscii - 32
        Select Case KeyAscii
            Case 34, 39
                KeyAscii = 0 : GoTo EventExitSub
            Case System.Windows.Forms.Keys.Return
                If txtrecloc.Text.Trim.Length = 0 Then
                    MsgBox("Enter Location Code", MsgBoxStyle.Information, ResolveResString(100))
                    txtrecloc.Focus()
                    GoTo EventExitSub
                ElseIf txtrecloc.Text.Trim.Length > 0 Then
                    Call txtRecLoc_Validating(txtrecloc, New System.ComponentModel.CancelEventArgs(False))
                    GoTo EventExitSub
                End If
        End Select
        GoTo EventExitSub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtRecLoc_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtrecloc.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If (KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0) Then Call cmdRecLocList_Click(cmdRecLocList, New System.EventArgs()) : Exit Sub
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtRecLoc_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtrecloc.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        Dim strsql As String
        If txtrecloc.Text.Length > 0 Then
            If optGRIN.Checked = True Then
                strsql = "SELECT DESCRIPTION FROM LOCATION_MST" & _
                        " WHERE UNIT_CODE = '" & gstrUNITID & "' AND upper(loc_type) = 'R' AND Location_Code='" & txtrecloc.Text.Trim & "'"
            ElseIf optSuppGRIN.Checked = True Then
                strsql = "SELECT DESCRIPTION FROM LOCATION_MST" & _
                        " WHERE UNIT_CODE = '" & gstrUNITID & "' AND upper(loc_type) = 'J' AND Location_Code='" & txtrecloc.Text.Trim & "'"
            End If
            lblRecLoc.Text = GetQryOutput(strsql)
            If lblRecLoc.Text.Trim.Length Then
                If ctlitem_c1.Enabled = True Then ctlitem_c1.Focus()
            Else
                Cancel = True
                Call ConfirmWindow(10402)
                lblRecLoc.Text = String.Empty
                txtrecloc.Clear()
                GoTo EventExitSub
            End If
        End If
        GoTo EventExitSub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub MakeRowHeight()
        On Error GoTo ErrHandler
        Dim intRow As Short
        With sprdata
            For intRow = 0 To .MaxRows
                .set_RowHeight(intRow, 300)
            Next intRow
        End With
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Function ValidateBeforeAuthorize() As Boolean
        'On Error GoTo ErrHandler
        Dim strsql As String
        Dim blnStatus As Boolean
        Dim intCounter As Short
        Dim strItemCode As String
        Dim strRejectionReason As String
        Dim dblChallanQty As Double
        Dim dblActualQty As Double
        Dim dblAcceptedQty As Double
        Dim dblRejectedQty As Double
        Dim blnValidFlag As Boolean
        Dim blnRejectionDtls As Boolean
        Dim dblbinnedbatchqty As Double
        Dim dblTotalBatchQty As Double
        Dim dblAcceptedBatchQty As Double
        Dim intCounter1 As Short
        Dim strMSG As String
        Dim objbatchbinnedqty As ClsResultSetDB
        Dim Rs As ADODB.Recordset
        Dim strMissingPkt As String = ""
        Dim strItemDesc As String

        blnValidFlag = True
        mstrErrorDesc = String.Empty
        mlngErrorNo = 1
        mctlError = Nothing
        Try
            ''----- Validate Location Code
            If txtrecloc.Text.Trim.Length = 0 Then
                mstrErrorDesc = mstrErrorDesc & vbCrLf & mlngErrorNo & ".  Location From."
                mlngErrorNo = mlngErrorNo + 1
                blnValidFlag = False
                If mctlError Is Nothing Then mctlError = Me.txtrecloc
            Else
                If optGRIN.Checked = True Then
                    strsql = "SELECT TOP 1 1 FROM location_mst" & _
                    " WHERE UNIT_CODE = '" & gstrUNITID & "' AND upper(loc_type) in ('R') and location_code='" & txtrecloc.Text.Trim & "'"
                ElseIf optSuppGRIN.Checked = True Then
                    strsql = "SELECT TOP 1 1  FROM location_mst" & _
                    " WHERE UNIT_CODE = '" & gstrUNITID & "' AND upper(loc_type) in ('J') and location_code='" & txtrecloc.Text.Trim & "'"
                End If
                blnStatus = DataExist(strsql)
                If blnStatus = False Then
                    mstrErrorDesc = mstrErrorDesc & vbCrLf & mlngErrorNo & ".  Location From."
                    mlngErrorNo = mlngErrorNo + 1
                    blnValidFlag = False
                    If mctlError Is Nothing Then mctlError = Me.txtrecloc
                End If
            End If
            ''---- Validate Grin No.
            If ctlitem_c1.Text.Trim.Length = 0 Then
                mstrErrorDesc = mstrErrorDesc & vbCrLf & mlngErrorNo & ". GRIN Number."
                mlngErrorNo = mlngErrorNo + 1
                blnValidFlag = False
                If mctlError Is Nothing Then mctlError = Me.ctlitem_c1
            Else
                strsql = "SELECT TOP 1 1  FROM grn_hdr" & _
                        " WHERE UNIT_CODE = '" & gstrUNITID & "' AND doc_type=10" & _
                        " AND ISNULL(LEN(qa_authorized_code),0)=0" & _
                        " AND from_location='" & txtrecloc.Text.Trim & "'" & _
                        " AND doc_no=" & ctlitem_c1.Text
                blnStatus = DataExist(strsql)
                If blnStatus = False Then
                    mstrErrorDesc = mstrErrorDesc & vbCrLf & mlngErrorNo & ".  GRIN Number."
                    mlngErrorNo = mlngErrorNo + 1
                    blnValidFlag = False
                    If mctlError Is Nothing Then mctlError = Me.ctlitem_c1
                End If
            End If
            ''---- validate the Inspection Authority
            If txtEmpQA.Text.Trim.Length = 0 Then
                mstrErrorDesc = mstrErrorDesc & vbCrLf & mlngErrorNo & ".  Inspected By."
                mlngErrorNo = mlngErrorNo + 1
                blnValidFlag = False
                If mctlError Is Nothing Then mctlError = Me.txtEmpQA
            Else
                strsql = "SELECT TOP 1 1 FROM user_mst" & _
                        " WHERE UNIT_CODE = '" & gstrUNITID & "' AND user_id='" & txtEmpQA.Text.Trim & "'"
                blnStatus = DataExist(strsql)
                If blnStatus = False Then
                    mstrErrorDesc = mstrErrorDesc & vbCrLf & mlngErrorNo & ".  Inspected By."
                    mlngErrorNo = mlngErrorNo + 1
                    blnValidFlag = False
                    If mctlError Is Nothing Then mctlError = Me.txtEmpQA
                End If
            End If
            ''---- Validate Quantities in Grid
            With sprdata
                For intCount = 1 To .MaxRows
                    .Row = intCount
                    .Col = Col_Item_Code : strItemCode = .Text.Trim
                    .Col = Col_Item_Description : strItemDesc = .Text.Trim
                    .Col = Col_Challan_Qty : dblChallanQty = Val(.Text)
                    .Col = Col_Actual_Qty : dblActualQty = Val(.Text)
                    .Col = Col_Rejected_Qty : dblRejectedQty = Val(.Text)
                    .Col = Col_Accepted_Qty : dblAcceptedQty = Val(.Text)
                    .Col = Col_Rejection_Reason : strRejectionReason = .Text.Trim

                    ''VID CHANGES
                    If mblnVendInvDateOlderThanPODate AndAlso txtReqApprovalStatus.Text.Trim.ToUpper = "REJECTED" Then
                        If dblAcceptedQty > 0 Then
                            mstrErrorDesc = mstrErrorDesc & vbCrLf & mlngErrorNo & ".GRIN should be fully rejected as Vendor Invoice Date is less than PO Date and same has been rejected for acceptance from approval workflow ! Kindly reject full Qty."
                            mlngErrorNo = mlngErrorNo + 1
                            blnValidFlag = False
                            If mctlError Is Nothing Then
                                mctlError = Me.sprdata
                                .Row = intCount
                                .Col = Col_Rejected_Qty
                                .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                Exit For
                            End If
                        End If
                    End If

                    If dblActualQty <> 0 And dblAcceptedQty = 0 And dblRejectedQty = 0 Then
                        mstrErrorDesc = mstrErrorDesc & vbCrLf & mlngErrorNo & ". Both Accepted and Rejected Quantities can't be Zero for Item " & strItemCode
                        mlngErrorNo = mlngErrorNo + 1
                        blnValidFlag = False
                        If mctlError Is Nothing Then
                            mctlError = Me.sprdata
                            .Row = intCount
                            .Col = Col_Rejected_Qty
                            .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                            Exit For
                        End If
                    Else
                        If dblAcceptedQty > dblActualQty Then
                            mstrErrorDesc = mstrErrorDesc & vbCrLf & mlngErrorNo & ". Accepted Quantity Can't Exceed Actual Qty for Item " & strItemCode
                            mlngErrorNo = mlngErrorNo + 1
                            blnValidFlag = False
                            If mctlError Is Nothing Then
                                mctlError = Me.sprdata
                                .Row = intCount
                                .Col = Col_Rejected_Qty
                                .Action = FPSpreadADO.ActionConstants.ActionActiveCell : Exit For
                            End If
                        End If
                    End If

                    If mblnBarcodeGRN = True And Barcode_Location(txtrecloc.Text, txtdesloc.Text) Then
                        'If (GetBarcodeTracking(strItemCode) = True) And (mstrGrinType <> "E" And mstrGrinType <> "U") Then
                        If (GetBarcodeTracking(strItemCode) = True) Then
                            Rs = mP_Connection.Execute("SELECT DBO.UFN_GETBINNEDQTY('" & gstrUNITID & "'," & ctlitem_c1.Text & ",'" & strItemCode & "','" & Me.txtdesloc.Text & "') AS BINNEDQTY")
                            If Not (Rs.BOF And Rs.EOF) Then
                                If Math.Round(Rs.Fields("BINNEDQTY").Value, 2) > Math.Round(dblAcceptedQty, 2) Then
                                    mstrErrorDesc = mstrErrorDesc & vbCrLf & mlngErrorNo & ". Bags Available Qty.: [" & Rs.Fields("BINNEDQTY").Value & "] In Store Is Greater Than Accepted Qty.: [" & dblAcceptedQty & "] GRIN Can't be Authorised"
                                    mlngErrorNo = mlngErrorNo + 1
                                    blnValidFlag = False
                                    If mctlError Is Nothing Then
                                        mctlError = Me.sprdata
                                        .Row = intCount
                                        .Col = Col_Accepted_Qty
                                        .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                        Rs.Close()
                                        Rs = Nothing
                                        Exit For
                                    End If
                                End If
                                If Math.Round(Rs.Fields("BINNEDQTY").Value, 2) < Math.Round(dblAcceptedQty, 2) Then
                                    strMissingPkt = GetMissingPakcets(ctlitem_c1.Text, strItemCode)
                                    mstrErrorDesc = mstrErrorDesc & vbCrLf & mlngErrorNo & ". Bags Available Qty.[" & Rs.Fields("BINNEDQTY").Value & "] In Store Is Less Than Accepted Qty.[" & dblAcceptedQty & "] GRIN Can't be Authorised" & vbCrLf & strMissingPkt & vbCrLf
                                    mlngErrorNo = mlngErrorNo + 1
                                    blnValidFlag = False
                                    If mctlError Is Nothing Then
                                        mctlError = Me.sprdata
                                        .Row = intCount
                                        .Col = Col_Accepted_Qty
                                        .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                        Rs.Close()
                                        Rs = Nothing
                                        Exit For
                                    End If
                                End If
                            End If
                            Rs.Close()
                            Rs = Nothing
                        End If
                    End If
                Next
            End With
            ''----- Batch Details Validation Checking
            If mblnBatchTracking = True Then
                With sprdata
                    For intCounter = 1 To .MaxRows
                        .Row = intCounter
                        .Col = Col_Item_Code : strItemCode = .Text.Trim
                        .Row = intCounter
                        .Col = Col_Accepted_Qty : dblAcceptedQty = Val(.Text)
                        If BatchDetailsExist(strItemCode) = False Then
                            mstrErrorDesc = mstrErrorDesc & vbCrLf & mlngErrorNo & ". Batch Authorization Details have not been entered for Item: " & strItemCode
                            mlngErrorNo = mlngErrorNo + 1
                            blnValidFlag = False
                            If mctlError Is Nothing Then
                                mctlError = Me.sprdata
                                .Row = intCounter
                                .Col = Col_Batch_Details
                                .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                            End If
                            Exit For
                        Else ' If Batch Details Exist
                            If Math.Round(BatchAcceptedOrRejectedQty(4, strItemCode), 4) <> Math.Round(dblAcceptedQty, 4) Then
                                mstrErrorDesc = mstrErrorDesc & vbCrLf & mlngErrorNo & ". Complete Authorization Batch Details for Accepted Quantity " & dblAcceptedQty & " have not been entered for Item " & strItemCode
                                mlngErrorNo = mlngErrorNo + 1
                                blnValidFlag = False
                                If mctlError Is Nothing Then
                                    mctlError = Me.sprdata
                                    .Row = intCount
                                    .Col = Col_Batch_Details
                                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                                End If
                                Exit For
                            End If
                        End If
                    Next
                End With
            End If
            If mblnBarcodeGRN = True And Barcode_Location(txtrecloc.Text, txtdesloc.Text) And mstrGrinType <> "E" And mstrGrinType <> "U" Then
                If mblnBatchTracking = True Then
                    With sprdata
                        For intCount = 1 To .MaxRows
                            .Row = intCount
                            .Col = Col_Item_Code
                            strItemCode = .Text.Trim
                            If GetBarcodeTracking(strItemCode) = True Then
                                .Row = intCount
                                .Col = Col_Accepted_Qty
                                dblAcceptedBatchQty = CDbl(.Text)
                                If BatchDetailsExist(strItemCode) = True Then ' If Batch Details Exist
                                    dblTotalBatchQty = 0
                                    For intCounter1 = 1 To UBound(garrBatchDetails, 2)
                                        If StrComp(strItemCode, garrBatchDetails(1, intCounter1), CompareMethod.Text) = 0 Then
                                            objbatchbinnedqty = New ClsResultSetDB
                                            objbatchbinnedqty.GetResult("select dbo.fn_getbinnedqty_batchwise" & _
                                                " ('" & gstrUNITID & "', '" & garrBatchDetails(2, intCounter1) & "','" & strItemCode & "','" & Trim(ctlitem_c1.Text) & "') as batchwisebinned_qty")
                                            If Not objbatchbinnedqty.EOFRecord Then
                                                dblbinnedbatchqty = objbatchbinnedqty.GetValue("batchwisebinned_qty")
                                            End If
                                            objbatchbinnedqty.ResultSetClose()
                                            objbatchbinnedqty = Nothing
                                            dblTotalBatchQty = CDbl(garrBatchDetails(4, intCounter1))
                                            If Math.Round(dblTotalBatchQty, 2) <> Math.Round(dblbinnedbatchqty, 2) Then
                                                strMSG = "Binned Quantity Of Batch No (" & garrBatchDetails(2, intCounter1) & ") Is Not Equal To " & vbCrLf & " Batch Accepted Quantity Of Batch No (" & garrBatchDetails(2, intCounter1) & ")" & vbCrLf & " So Grin Cannot Be Authorized" & vbCrLf & "Binned Batch Quantity Of Batch No (" & garrBatchDetails(2, intCounter1) & ") Is " & dblbinnedbatchqty & vbCrLf & "Batch Accepted Quantity Of Batch No (" & garrBatchDetails(2, intCounter1) & ") Is " & garrBatchDetails(4, intCounter1)
                                                mstrErrorDesc = mstrErrorDesc & vbCrLf & mlngErrorNo & strMSG
                                                mlngErrorNo = mlngErrorNo + 1
                                                blnValidFlag = False
                                                If mctlError Is Nothing Then
                                                    mctlError = Me.sprdata
                                                    .Row = intCount
                                                    .Col = Col_Batch_Details
                                                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                                                End If
                                                Exit Function
                                            End If
                                        End If
                                    Next
                                End If
                            End If
                        Next
                    End With
                End If
            End If
            ''---- Validate the inspection details
            Dim strIEValue As String
            If mbln_Inspection_Control_Details = True And optGRIN.Checked = True Then
                With Me.sprdata
                    For intCounter = 1 To .MaxRows
                        .Row = intCounter
                        .Col = Col_Item_Code : strItemCode = .Text.Trim
                        .Row = intCounter
                        .Col = Col_Inspection_Entry : strIEValue = .Text.Trim
                        If InspectionDetailsExist(strItemCode) = False And strIEValue = "1" Then
                            mstrErrorDesc = mstrErrorDesc & vbCrLf & mlngErrorNo & ". Inspection Details have not been entered for Item " & strItemCode
                            mlngErrorNo = mlngErrorNo + 1
                            blnValidFlag = False
                            If mctlError Is Nothing Then
                                mctlError = Me.sprdata
                                .Row = intCounter
                                .Col = col_Inspection_Control_Details
                                .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                            End If
                            Exit For
                        End If
                    Next
                End With
            End If
            ''----- Rejection Details Validation Checking
            Dim strAccptancecriteria As String
            '101161852
            If Not mblnIQARequired Then
                With sprdata
                    For intCount = 1 To .MaxRows
                        .Row = intCount
                        .Col = Col_Item_Code : strItemCode = .Text.Trim
                        .Row = intCount
                        .Col = Col_Accepted_Qty : dblAcceptedQty = Val(.Text)
                        .Row = intCount
                        .Col = Col_Rejected_Qty : dblRejectedQty = Val(.Text)
                        .Row = intCount
                        .Col = Col_Inspection_Entry : strAccptancecriteria = .Text.Trim
                        blnRejectionDtls = False
                        If mstrGRType.Value = "U" Or mstrGRType.Value = "E" Then
                            For intCounter = 1 To UBound(marrDefectDetails, 2)
                                If StrComp(strItemCode, marrDefectDetails(0, intCounter)) = 0 Then
                                    blnRejectionDtls = True
                                    Exit For
                                Else
                                    blnRejectionDtls = False
                                End If
                            Next
                            If (strAccptancecriteria = "0" Or strAccptancecriteria = "") And dblRejectedQty > 0 Then
                                If blnRejectionDtls = False Then
                                    mstrErrorDesc = mstrErrorDesc & vbCrLf & mlngErrorNo & ". Defect Details have not been entered for Item " & strItemCode
                                    mlngErrorNo = mlngErrorNo + 1
                                    blnValidFlag = False
                                    If mctlError Is Nothing Then
                                        mctlError = Me.sprdata
                                        .Row = intCount
                                        .Col = Col_Rejection_Details
                                        .Action = FPSpreadADO.ActionConstants.ActionActiveCell : Exit For
                                    End If
                                End If
                            End If
                        Else
                            If dblRejectedQty > 0 Then
                                For intCounter = 1 To UBound(marrDefectDetails, 2)
                                    If StrComp(strItemCode, marrDefectDetails(0, intCounter)) = 0 Then
                                        blnRejectionDtls = True
                                        Exit For
                                    Else
                                        blnRejectionDtls = False
                                    End If
                                Next
                                If (strAccptancecriteria = "0" Or strAccptancecriteria = "") And dblRejectedQty > 0 Then
                                    If blnRejectionDtls = False Then
                                        mstrErrorDesc = mstrErrorDesc & vbCrLf & mlngErrorNo & ". Defect Details have not been entered for Item " & strItemCode
                                        mlngErrorNo = mlngErrorNo + 1
                                        blnValidFlag = False
                                        If mctlError Is Nothing Then
                                            mctlError = Me.sprdata
                                            .Row = intCount
                                            .Col = Col_Rejection_Details
                                            .Action = FPSpreadADO.ActionConstants.ActionActiveCell : Exit For
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                End With
            End If
            '*10857384
            If mstrGRType.ToString.ToUpper = "W" Then
                Dim Rows As New ArrayList
                Dim strRows As String = String.Empty
                With sprdata
                    For intRow As Integer = 1 To .MaxRows
                        .Row = intRow
                        .Col = Col_Item_Code
                        strItemCode = .Text.Trim.ToUpper
                        Dim intCount As Integer = ListESIDetail.Where(Function(x) x.ItemCode.ToUpper = strItemCode.ToUpper).Count
                        If intCount = 0 Then
                            Rows.Add(intRow)
                        End If
                    Next
                End With
                If Rows.Count > 0 Then
                    strRows = Rows.ToArray.Aggregate(Function(x, y) x.ToString + "," + y.ToString)
                    If strRows.Length > 0 Then
                        mstrErrorDesc = mstrErrorDesc & vbCrLf & mlngErrorNo & ". Please Add/Update ESI Details at row(s) : " & strRows
                        mlngErrorNo = mlngErrorNo + 1
                        blnValidFlag = False
                        If mctlError Is Nothing Then
                            mctlError = Me.sprdata
                        End If
                    End If
                End If
            End If
            '*
            ValidateBeforeAuthorize = blnValidFlag
            Exit Function
        Catch ex As Exception
            gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            objbatchbinnedqty = Nothing
            Rs = Nothing
        End Try
    End Function
    Private Sub CreateDataString()
        '***************************************************************************
        'Revision Date  - 07-Oct-2006
        'Revised By     - Davinder Singh
        'Issue No       - 17893
        'Description    - To prepare the Strings to pass data to the 'ClsGrinInspection'
        '                 If grin is without PO then Prepare string from grid otherwise from array
        '***************************************************************************
        On Error GoTo ErrHandler
        Dim strSql As String
        Dim intCtr As Short
        Dim ldblExchangeRate As String
        Dim strRejectionRem As String
        Dim strItemCode As String
        Dim curItemRate As Double
        Dim curQty As Double
        Dim curLandingRate As Double
        mstrMasterString = String.Empty
        mstrDetailString = String.Empty

        '' Add By Rajeev Gupta on 17 Oct 2012
        '' Against Issue ID 10268417 - Exchange rate issue in PV
        'strSql = "SELECT CExch_MultiFactor from Gen_CurExchMaster" & _
        '        " WHERE UNIT_CODE = '" & gstrUNITID & "' AND CExch_InOut=0" & _
        '        " AND CExch_CurrencyTo='" & mstrCurrencyID & "'"
        strSql = "SET DATEFORMAT 'DMY' SELECT CExch_MultiFactor from Gen_CurExchMaster" & _
                " WHERE UNIT_CODE = '" & gstrUNITID & "' AND CExch_InOut=0" & _
                " AND CExch_CurrencyTo='" & mstrCurrencyID.Trim & "' AND CEXCH_DATEFROM <= '" & txtgrdt.Text & "' " & _
                " ORDER BY CEXCH_DATEFROM DESC"
        '' Changes 10268417 Ends Here        

        If GetQryOutput(strSql).ToString.Length > 0 Then
            ldblExchangeRate = GetQryOutput(strSql)
        Else
            ldblExchangeRate = 0
        End If
        If ldblExchangeRate = 0 Then
            MsgBox("Exchange Rate not Defined for Currency Code [" & mstrCurrencyID & "]", MsgBoxStyle.Critical, ResolveResString(100))
            Exit Sub
        End If
        ''---- Prepare the master string from form header
        If mstrPOstatus = "W" Then
            mstrMasterString = mstrMasterString & Trim(ctlitem_c1.Text) & "»" & Trim(txtgrdt.Text) & "»" & dtpQADate.Text & "»" & gstrUNITID & "»C»" & Trim(txtpo.Text) & "»" & Trim(txtinv.Text) & "»" & Trim(txtinvdt.Text) & "»" & mstrChallanNo & "»" & Trim(txtvencd.Text) & "»" & Trim(txtvend.Text) & "»" & mstrCurrencyID & "»" & ldblExchangeRate & "»" & mP_User & "»" & VB.Left(Trim(txtRemarks.Text), 35) & "»" & Val(txtInvValue.Text) & "»" & mstrGateEntryID
        Else
            mstrMasterString = mstrMasterString & Trim(ctlitem_c1.Text) & "»" & Trim(txtgrdt.Text) & "»" & dtpQADate.Text & "»" & gstrUNITID & "»P»" & Trim(txtpo.Text) & "»" & Trim(txtinv.Text) & "»" & Trim(txtinvdt.Text) & "»" & mstrChallanNo & "»" & Trim(txtvencd.Text) & "»" & Trim(txtvend.Text) & "»" & mstrCurrencyID & "»" & ldblExchangeRate & "»" & mP_User & "»" & VB.Left(Trim(txtRemarks.Text), 35) & "»" & Val(txtInvValue.Text) & "»" & mstrGateEntryID
        End If
        ''---- Prepare the detailstring from Grid
        If mstrPOstatus = "W" Then
            With sprdata
                For intCtr = 1 To .MaxRows Step 1
                    .Row = intCtr
                    mstrDetailString = mstrDetailString & intCtr & "»"
                    .Col = Col_Project_Code
                    mstrDetailString = mstrDetailString & .Text.Trim & "»"
                    .Col = Col_Item_Description
                    mstrDetailString = mstrDetailString & VB.Left(.Text.Trim, 35) & "»"
                    .Col = Col_Item_Code
                    strItemCode = .Text.Trim
                    mstrDetailString = mstrDetailString & strItemCode & "»"
                    .Col = Col_GL_Group
                    mstrDetailString = mstrDetailString & .Text.Trim & "»"
                    .Col = Col_Challan_Qty
                    mstrDetailString = mstrDetailString & .Text.Trim & "»"
                    .Col = Col_Actual_Qty
                    mstrDetailString = mstrDetailString & .Text.Trim & "»"
                    curQty = Val(.Text.Trim)
                    .Col = Col_Item_Rate
                    mstrDetailString = mstrDetailString & .Text.Trim & "»"
                    curItemRate = Val(.Text.Trim)
                    .Col = Col_Item_UOM
                    mstrDetailString = mstrDetailString & .Text.Trim & "»"
                    .Col = Col_Discount_Per
                    mstrDetailString = mstrDetailString & .Text.Trim & "»"
                    .Col = Col_Asseccible_Rate
                    mstrDetailString = mstrDetailString & .Text.Trim & "»"
                    .Col = Col_Accepted_Qty
                    mstrDetailString = mstrDetailString & .Text.Trim & "»"
                    .Col = Col_Rejection_Reason
                    strRejectionRem = .Text.Trim
                    Call QuoteString(strRejectionRem)
                    strRejectionRem = VB.Left(strRejectionRem, 35)
                    mstrDetailString = mstrDetailString & strRejectionRem & "»"
                    .Col = Col_Rejected_Qty
                    mstrDetailString = mstrDetailString & .Text.Trim & "»"
                    curLandingRate = CalcLandingRate(strItemCode, curQty, curItemRate)
                    mstrDetailString = mstrDetailString & curLandingRate & "»¦"
                Next intCtr
            End With
        Else
            For intCtr = 1 To UBound(mArrPOdtl) Step 1
                mstrDetailString = mstrDetailString & intCtr & "»"
                mstrDetailString = mstrDetailString & mArrPOdtl(intCtr).ProjectCode & "»"
                mstrDetailString = mstrDetailString & VB.Left(mArrPOdtl(intCtr).Item_Desc, 35) & "»"
                strItemCode = mArrPOdtl(intCtr).ItemCode
                mstrDetailString = mstrDetailString & strItemCode & "»"
                mstrDetailString = mstrDetailString & mArrPOdtl(intCtr).GLGrpCode & "»"
                mstrDetailString = mstrDetailString & mArrPOdtl(intCtr).ChallanQty & "»"
                mstrDetailString = mstrDetailString & mArrPOdtl(intCtr).ActualQty & "»"
                curQty = mArrPOdtl(intCtr).ActualQty
                mstrDetailString = mstrDetailString & mArrPOdtl(intCtr).Rate_Renamed & "»"
                curItemRate = mArrPOdtl(intCtr).Rate_Renamed
                mstrDetailString = mstrDetailString & mArrPOdtl(intCtr).UOM & "»"
                mstrDetailString = mstrDetailString & mArrPOdtl(intCtr).DiscountPer & "»"
                mstrDetailString = mstrDetailString & mArrPOdtl(intCtr).AssesableRate & "»"
                mstrDetailString = mstrDetailString & mArrPOdtl(intCtr).AcceptedQty & "»"
                strRejectionRem = mArrPOdtl(intCtr).Rej_reason
                Call QuoteString(strRejectionRem)
                strRejectionRem = VB.Left(strRejectionRem, 35)
                mstrDetailString = mstrDetailString & strRejectionRem & "»"
                mstrDetailString = mstrDetailString & mArrPOdtl(intCtr).RejectedQty & "»"
                curLandingRate = CalcLandingRate(strItemCode, curQty, curItemRate)
                mstrDetailString = mstrDetailString & curLandingRate & "»"
                mstrDetailString = mstrDetailString & mArrPOdtl(intCtr).PONO & "¦"
            Next intCtr
        End If
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub FillDefects(ByVal pstrItemCode As String, ByVal strMeasureCode As String)
        On Error GoTo ErrHandler
        Dim strsql As String
        Dim rstDB As ClsResultSetDB
        Dim intTotalDefects As Short
        Dim intCounter As Short
        Dim intCounter1 As Short

        '10847531

        'strsql = "Select Defect_C,defect_nm from defect_mst WHERE UNIT_CODE = '" & gstrUNITID & "'"
        Dim obj As Object = SqlConnectionclass.ExecuteScalar("SELECT COUNT(*) FROM DEFECT_MST (NOLOCK) WHERE UNIT_CODE='" + gstrUNITID + "' AND LEN(ISNULL(DEFECT_CATEGORY ,''))>0")
        If Val(obj.ToString()) > 0 Then
            strsql = " SELECT Defect_C,defect_nm,DEFECT_CATEGORY " & _
                " FROM defect_mst " & _
                " WHERE UNIT_CODE = '" & gstrUNITID & "' " & _
                "       AND DEFECT_CATEGORY='Incoming QA' " & _
                " ORDER BY defect_nm"
        Else
            strsql = "Select Defect_C,defect_nm,DEFECT_CATEGORY from defect_mst WHERE UNIT_CODE = '" & gstrUNITID & "' ORDER BY defect_nm"
        End If
       
        obj = Nothing

        rstDB = New ClsResultSetDB
        Call rstDB.GetResult(strsql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        intTotalDefects = rstDB.GetNoRows
        If intTotalDefects > 0 Then
            With Me.vaDefects
                .Enabled = True
                .MaxRows = 0
                .MaxRows = intTotalDefects
                lblDefRejection.Text = "0"
                'lblBalRejQty.Text = "0"
                rstDB.MoveFirst()
                For intCounter = 1 To intTotalDefects
                    .set_RowHeight(intCounter, 250)
                    .Row = intCounter

                    .Col = Col_Defect_Code
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                    .Text = rstDB.GetValue("Defect_C")
                    .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
                    .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
                    .Row = intCounter

                    .Col = Col_Defect_Help
                    .ColHidden = True
                    .Row = intCounter
                    .Col = Col_Defect_Description
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                    .Text = rstDB.GetValue("Defect_nm")
                    .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
                    .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
                    .Row = intCounter

                    .Col = Col_Defect_qty
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                    .TypeFloatDecimalPlaces = 4
                    .TypeFloatMax = 99999999.9999
                    .TypeFloatMin = 0.0#
                    .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                    .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter

                    '10847531
                    .Col = Col_Defect_Category
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                    .Text = rstDB.GetValue("DEFECT_CATEGORY")
                    .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
                    .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
                    'End Here

                    For intCounter1 = 1 To UBound(marrDefectDetails, 2)
                        If StrComp(marrDefectDetails(0, intCounter1), pstrItemCode) = 0 Then ' Item is already in the Array
                            If StrComp(marrDefectDetails(1, intCounter1), rstDB.GetValue("defect_C")) = 0 Then 'Defect is Already in the Array
                                If CDbl(marrDefectDetails(2, intCounter1)) > 0 Then
                                    .Row = intCounter
                                    .Col = Col_Defect_qty
                                    .Text = marrDefectDetails(2, intCounter1) 'Defect Qty is written.
                                    lblDefRejection.Text = CStr(Val(lblDefRejection.Text) + CDbl(marrDefectDetails(2, intCounter1)))
                                    lblBalRejQty.Text = CStr(Val(lblBalRejQty.Text) - CDbl(marrDefectDetails(2, intCounter1)))
                                    If GetUOMDecimalPlacesAllowed(strMeasureCode) > 0 Then
                                        lblDefRejection.Text = Format(Val(lblDefRejection.Text), "#0.0000")
                                        lblBalRejQty.Text = Format(Val(lblBalRejQty.Text), "#0.0000")
                                    Else
                                        lblDefRejection.Text = Format(Val(lblDefRejection.Text), "#0")
                                        lblBalRejQty.Text = Format(Val(lblBalRejQty.Text), "#0")
                                    End If
                                End If
                            Else
                            End If
                        Else
                        End If
                    Next
                    rstDB.MoveNext()
                Next
            End With
        Else
            MsgBox("No Defect Code(s) are defined in the System. First Define Defect Code(s).", MsgBoxStyle.Information, ResolveResString(100))
        End If
        rstDB.ResultSetClose()
        rstDB = Nothing
        Exit Sub
ErrHandler:
        rstDB = Nothing
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Function Defects_Details_Insertion(ByVal pstrItemCode As String) As Boolean
        On Error GoTo ErrHandler
        Dim strsql As String
        Dim intCounter As Short
        Dim strDefectItemCode As String
        Dim strDefectCode As String
        Dim dblDefectQty As Double
        Defects_Details_Insertion = True
        If UBound(marrDefectDetails, 2) > 0 Then
            For intCounter = 1 To UBound(marrDefectDetails, 2)
                strDefectItemCode = marrDefectDetails(0, intCounter)
                If StrComp(pstrItemCode, strDefectItemCode) = 0 Then
                    strDefectCode = marrDefectDetails(1, intCounter)
                    dblDefectQty = Val(marrDefectDetails(2, intCounter))
                    strsql = "INSERT INTO Inspection_Defect_Details"
                    strsql = strsql & " (GRIN_No,GRIN_From_Location,Item_Code,"
                    strsql = strsql & " Defect_Code ,Defect_Quantity,"
                    strsql = strsql & " Ent_Dt,Ent_Userid,Upd_dt,Upd_Userid,UNIT_CODE)"
                    strsql = strsql & " VALUES(" & ctlitem_c1.Text & ",'" & Trim(txtrecloc.Text) & "','" & pstrItemCode & "','"
                    strsql = strsql & strDefectCode & "'," & dblDefectQty
                    strsql = strsql & ",getdate(),'" & Trim(mP_User) & "',getdate(),'" & Trim(mP_User) & "','" & gstrUNITID & "')"
                    If dblDefectQty > 0 Then mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                Else
                End If
            Next
        End If
        Exit Function
ErrHandler:
        Defects_Details_Insertion = False
        MsgBox(Err.Description & " - DDI [ " & Err.Number & " ] ", MsgBoxStyle.Information, ResolveResString(100))
    End Function
    Private Function CalcLandingRate(ByVal ItemCode As String, ByVal Qty As Double, ByVal Rate_Renamed As Double) As Double
        On Error GoTo ErrHandler
        Dim rstemp As ADODB.Recordset
        Dim dblBasic As Double
        Dim dblExcise As Double
        Dim dblsalestax As Double
        CalcLandingRate = 0
        If Qty = 0 Then Exit Function
        rstemp = New ADODB.Recordset
        rstemp.Open("SELECT isnull(excise_amount,0) as excise_amount ," & _
            " isnull(salestax_amount,0) as salestax_amount FROM grn_dtl a,vend_item b" & _
            " WHERE A.UNIT_CODE = B.UNIT_CODE AND A.UNIT_CODE = '" & gstrUNITID & "' AND" & _
            " a.item_Code ='" & ItemCode & "' AND a.item_code=b.item_code" & _
            " AND a.po_no=b.Pur_Order_No", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
        If Not rstemp.EOF Then
            dblBasic = Qty * Rate_Renamed
            dblExcise = dblBasic * rstemp.Fields("excise_amount").Value / 100
            dblsalestax = (dblBasic + dblExcise) * rstemp.Fields("salestax_amount").Value / 100
            dblBasic = dblBasic + dblExcise + dblsalestax
            CalcLandingRate = dblBasic / Qty
        End If
        rstemp.Close()
        rstemp = Nothing
        Exit Function
ErrHandler:
        rstemp = Nothing
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub FillGRINBatchDetails(ByVal pstrItemCode As String, ByVal dblAccQty As Double)
        '***************************************************************************
        'Revised By       - Davinder Singh
        'Revision Date    - 14 Sep 2007
        'Issue No         - 21102
        'Revision History - 1) Concept of Shelf life is added and
        '                       MFG and EXP dates are calculated
        '                   2) To check for Expired batches
        '***************************************************************************
        On Error GoTo ErrHandler
        Dim strsql As String
        Dim strbatchno As String
        Dim rstDB As ClsResultSetDB
        Dim intCounter As Short
        Dim intTotalrecs As Short
        Dim dblBatchQty As Double
        Dim lngShelfLife As Integer
        Dim blnVendorBarcode As Boolean

        ''ADDED AGAINST ISSUE ID : 10504051
        blnVendorBarcode = VendorBarcode(txtvencd.Text.Trim)
        ''

        strsql = "Select Batch_No,Batch_Date,Batch_Qty from ItemBatch_Dtl" & _
                " where UNIT_CODE = '" & gstrUNITID & "' AND" & _
                " doc_type=10 and Doc_No=" & ctlitem_c1.Text & _
                " and from_location='" & txtrecloc.Text.Trim & "'" & _
                " and item_Code='" & pstrItemCode & "'" & _
                " Order by Serial_No"
        rstDB = New ClsResultSetDB
        Call rstDB.GetResult(strsql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rstDB.GetNoRows > 0 Then
            intTotalrecs = rstDB.GetNoRows : rstDB.MoveFirst()
            frmAuthBatchDetails.vaBatchDetails.MaxRows = 0
            With frmAuthBatchDetails.vaBatchDetails
                For intCounter = 1 To intTotalrecs
                    .MaxRows = .MaxRows + 1
                    .set_RowHeight(.MaxRows, 300)
                    .Row = .MaxRows
                    .Col = Col_Batch_No
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                    .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
                    .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
                    .Text = rstDB.GetValue("Batch_No") : strbatchno = .Text.Trim
                    .Row = .MaxRows
                    .Col = Col_Batch_Date
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                    .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
                    .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
                    .Text = VB6.Format(rstDB.GetValue("Batch_Date"), gstrDateFormat)

                    ''Changed and Written against Issue Id : 10504051
                    If blnVendorBarcode Then
                        lngShelfLife = Convert.ToDouble(SqlConnectionclass.ExecuteScalar("SELECT SHELF_LIFE FROM UFN_GET_VENDOR_BATCH_DTL('" & gstrUNITID & "'," & Val(ctlitem_c1.Text) & ",'" & txtvencd.Text & "','" & rstDB.GetValue("Batch_No") & "','" & pstrItemCode & "') A"))
                    End If
                    If lngShelfLife = 0 Then
                        lngShelfLife = GetShelfLife(pstrItemCode)
                    End If
                    ''

                    If lngShelfLife = 0 Then
                        .Row = intCounter
                        .Col = Col_Batch_MFG_Date
                        .ColHidden = True
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                        .Row = intCounter
                        .Col = Col_Batch_EXP_Date
                        .ColHidden = True
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                    Else
                        .Row = intCounter
                        .Col = Col_Batch_MFG_Date
                        .ColHidden = False
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                        .Row = intCounter
                        .Col = Col_Batch_EXP_Date
                        .ColHidden = False
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                    End If

                    If GetUOMDecimalPlacesAllowed(frmAuthBatchDetails.mstrMeasureCode) > 0 Then
                        .Row = .MaxRows
                        .Col = Col_Batch_Qty
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                        .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                        .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
                        .Text = VB6.Format(rstDB.GetValue("Batch_Qty"), "0.0000") : dblBatchQty = CDbl(.Text)
                        If BatchDetailsExist(pstrItemCode) Then
                            .Row = .MaxRows
                            .Col = Col_Batch_Accepted_Qty
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                            .TypeFloatDecimalPlaces = 4
                            .TypeFloatMax = VB6.Format(dblBatchQty, "0.0000")
                            .TypeFloatMin = 0.0#
                            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .Text = VB6.Format(FillAcceptedOrRejectedQty(pstrItemCode, strbatchno, 4), "0.0000")
                            .Row = .MaxRows
                            .Col = Col_Batch_Rejected_Qty
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .Text = VB6.Format(FillAcceptedOrRejectedQty(pstrItemCode, strbatchno, 5), "0.0000")
                            .Row = .MaxRows
                            .Col = Col_Batch_MFG_Date
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
                            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .Text = VB6.Format(FillMFGOrEXP(pstrItemCode, strbatchno, 6), gstrDateFormat)
                            .Row = .MaxRows
                            .Col = Col_Batch_EXP_Date
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
                            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .Text = VB6.Format(FillMFGOrEXP(pstrItemCode, strbatchno, 7), gstrDateFormat)
                        Else
                            .Row = .MaxRows
                            .Col = Col_Batch_Accepted_Qty
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                            .TypeFloatDecimalPlaces = 4
                            .TypeFloatMax = VB6.Format(dblBatchQty, "0.0000")
                            .TypeFloatMin = 0.0#
                            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .Text = "0.0000"
                            If Val(CStr(dblAccQty)) = 0 Then
                                .Row = .MaxRows
                                .Col = Col_Batch_Rejected_Qty
                                .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                                .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                                .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
                                .Text = VB6.Format(rstDB.GetValue("Batch_Qty"), "0.0000")
                            Else
                                .Row = .MaxRows
                                .Col = Col_Batch_Rejected_Qty
                                .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                                .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                                .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .Text = "0.0000"
                            End If
                            If lngShelfLife > 0 Then
                                .Row = .MaxRows
                                .Col = Col_Batch_MFG_Date
                                .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                                .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
                                .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
                                .Text = VB6.Format(rstDB.GetValue("Batch_Date"), gstrDateFormat)
                                .Row = .MaxRows
                                .Col = Col_Batch_EXP_Date
                                .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                                .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
                                .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
                                .Text = VB6.Format(CStr(GetBatchExpDate(CDate(rstDB.GetValue("Batch_Date")), lngShelfLife)), gstrDateFormat)
                            End If
                        End If
                    Else
                        .Row = .MaxRows
                        .Col = Col_Batch_Qty
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                        .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                        .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
                        .Text = VB6.Format(rstDB.GetValue("Batch_Qty"), "0") : dblBatchQty = CDbl(.Text)
                        If BatchDetailsExist(pstrItemCode) Then
                            .Row = .MaxRows
                            .Col = Col_Batch_Accepted_Qty
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeInteger
                            .TypeIntegerMin = 0
                            .TypeIntegerMax = VB6.Format(dblBatchQty, "0")
                            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .Text = VB6.Format(FillAcceptedOrRejectedQty(pstrItemCode, strbatchno, 4), "0")
                            .Row = .MaxRows
                            .Col = Col_Batch_Rejected_Qty
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .Text = VB6.Format(FillAcceptedOrRejectedQty(pstrItemCode, strbatchno, 5), "0")
                            .Row = .MaxRows
                            .Col = Col_Batch_MFG_Date
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
                            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .Text = VB6.Format(FillMFGOrEXP(pstrItemCode, strbatchno, 6), gstrDateFormat)
                            .Row = .MaxRows
                            .Col = Col_Batch_EXP_Date
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
                            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .Text = VB6.Format(FillMFGOrEXP(pstrItemCode, strbatchno, 7), gstrDateFormat)
                        Else
                            .Row = .MaxRows
                            .Col = Col_Batch_Accepted_Qty
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeInteger
                            .TypeIntegerMax = VB6.Format(dblBatchQty, "0")
                            .TypeIntegerMin = 0
                            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .Text = "0"
                            .Row = .MaxRows
                            .Col = Col_Batch_Rejected_Qty
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .Text = "0"
                            If Val(CStr(dblAccQty)) = 0 Then
                                .Row = .MaxRows
                                .Col = Col_Batch_Rejected_Qty
                                .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                                .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                                .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
                                .Text = VB6.Format(rstDB.GetValue("Batch_Qty"), "0")
                            Else
                                .Row = .MaxRows
                                .Col = Col_Batch_Rejected_Qty
                                .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                                .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                                .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter : .Text = "0"
                            End If
                            If lngShelfLife > 0 Then
                                .Row = .MaxRows
                                .Col = Col_Batch_MFG_Date
                                .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                                .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
                                .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
                                .Text = VB6.Format(CStr(rstDB.GetValue("Batch_Date")), gstrDateFormat)
                                .Row = .MaxRows
                                .Col = Col_Batch_EXP_Date
                                .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                                .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
                                .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
                                .Text = VB6.Format(CStr(GetBatchExpDate(CDate(rstDB.GetValue("Batch_Date")), lngShelfLife)), gstrDateFormat)
                            End If
                        End If
                    End If
                    If lngShelfLife > 0 Then
                        If optGRIN.Checked = True Then
                            .Row = .MaxRows
                            .Col = Col_Batch_EXP_Date
                            If ConvertToDate(.Text) < GetServerDate() Then
                                Call .SetText(Col_Batch_Rejected_Qty, .MaxRows, dblBatchQty)
                                .Row = .MaxRows
                                .Row2 = .MaxRows
                                .Col = 1
                                .Col2 = .MaxCols
                                .BlockMode = True
                                .Lock = True
                                .ForeColor = System.Drawing.Color.Red
                                .BlockMode = False
                                frmAuthBatchDetails.lblRedBatches.Visible = True
                            End If
                        Else
                            .Row = .MaxRows
                            .Col = Col_Batch_EXP_Date
                            If ConvertToDate(.Text) < GetServerDate() Then
                                .Row = .MaxRows
                                .Row2 = .MaxRows
                                .Col = 1
                                .Col2 = .MaxCols
                                .BlockMode = True
                                .ForeColor = System.Drawing.Color.Red
                                .BlockMode = False
                                frmAuthBatchDetails.lblRedBatches.Visible = True
                            End If
                        End If
                    End If
                    rstDB.MoveNext()
                Next
                frmAuthBatchDetails.cmdGrpBatchAuth.Enabled(0) = True
                .Enabled = True
                .Row = 1
                .Col = Col_Batch_Accepted_Qty
                .EditModePermanent = True
                .EditModeReplace = True
                .Action = FPSpreadADO.ActionConstants.ActionActiveCell : Exit Sub
            End With
        Else
            MsgBox("No Batch Details Exist for Item : " & pstrItemCode, MsgBoxStyle.Information, ResolveResString(100))
        End If
        rstDB.ResultSetClose()
        rstDB = Nothing
        Exit Sub
ErrHandler:
        rstDB = Nothing
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Function FillAcceptedOrRejectedQty(ByVal pstrItemCode As String, ByVal pstrBatchNo As String, ByVal pintIndex As Short) As Double
        On Error GoTo ErrHandler
        Dim intCounter As Short
        Dim strItemCode As String
        Dim strbatchno As String
        With frmBatch.vaBatchDetails
            If UBound(garrBatchDetails, 2) > 0 Then
                For intCounter = 1 To UBound(garrBatchDetails, 2)
                    strItemCode = garrBatchDetails(1, intCounter)
                    strbatchno = garrBatchDetails(2, intCounter)
                    If (StrComp(strItemCode, pstrItemCode, CompareMethod.Text) = 0 And StrComp(strbatchno, pstrBatchNo, CompareMethod.Text) = 0) Then
                        FillAcceptedOrRejectedQty = CDbl(garrBatchDetails(pintIndex, intCounter))
                    End If
                Next
            End If
        End With
        Exit Function
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function BatchAcceptedOrRejectedQty(ByVal pintIndex As Short, ByRef pstrItemCode As String) As Object '   ##
        On Error GoTo ErrHandler
        Dim intCounter As Short
        Dim dblqty As Double
        dblqty = 0
        For intCounter = 1 To UBound(garrBatchDetails, 2)
            If StrComp(garrBatchDetails(1, intCounter), pstrItemCode, CompareMethod.Text) = 0 Then
                dblqty = dblqty + CDbl(garrBatchDetails(pintIndex, intCounter))
            End If
        Next
        BatchAcceptedOrRejectedQty = dblqty
        Exit Function
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function UpdateBatchInfo(ByVal pstrItemCode As String) As Boolean
        '***************************************************************************
        'Revised By     - Davinder Singh
        'Revision Date  - 14 Sep 2007
        'Issue No       - 21102
        'Description    - New Stored Procedure is called to update ItemBatch_Mst
        '                 which also updates MFG and EXP dates
        '***************************************************************************
        On Error GoTo ErrHandler
        Dim strsql As String
        Dim strItemCode As String
        Dim intCounter As Short
        Dim strbatchno As String
        Dim strBatchDate As String
        Dim dblBatchAcceptedQty As Double
        Dim dblBatchRejectedQty As Double
        Dim oCmd As ADODB.Command
        Dim MfgDate As Object
        Dim ExpDate As Object
        strsql = String.Empty
        UpdateBatchInfo = True
        For intCounter = 1 To UBound(garrBatchDetails, 2)
            strItemCode = Trim(garrBatchDetails(1, intCounter))
            strbatchno = Trim(garrBatchDetails(2, intCounter))
            Call QuoteString(strbatchno)
            strBatchDate = garrBatchDetails(3, intCounter)
            dblBatchAcceptedQty = Val(garrBatchDetails(4, intCounter))
            dblBatchRejectedQty = Val(garrBatchDetails(5, intCounter))
            If Len(Trim(garrBatchDetails(6, intCounter))) <> 0 Then
                MfgDate = VB6.Format(garrBatchDetails(6, intCounter), "dd MMM yyyy")
            Else
                MfgDate = System.DBNull.Value
            End If
            If Len(Trim(garrBatchDetails(6, intCounter))) <> 0 Then
                ExpDate = VB6.Format(garrBatchDetails(7, intCounter), "dd MMM yyyy")
            Else
                ExpDate = System.DBNull.Value
            End If
            If (StrComp(strItemCode, pstrItemCode, CompareMethod.Text) = 0) Then
                ''--- Accepted and Rejected Quantities are Updated Batch Wise in ItemBatch_Dtl Table.
                strsql = strsql & " UPDATE ItemBatch_Dtl" & _
                    " SET Batch_Accepted_Qty=" & dblBatchAcceptedQty & "," & _
                    " Batch_Rejected_Qty=" & dblBatchRejectedQty & "" & _
                    " WHERE UNIT_CODE = '" & gstrUNITID & "' AND Doc_Type=10" & _
                    " AND Doc_No=" & Trim(ctlitem_c1.Text) & " AND From_Location='" & Trim(txtrecloc.Text) & "'" & _
                    " AND Item_Code='" & pstrItemCode & "'" & _
                    " AND Batch_No='" & strbatchno & "'" & vbCrLf
                ''--- Accepted Quantities are Updated Batch Wise at To Location. in ItemBatch_Mst
                If dblBatchAcceptedQty > 0 Then
                    oCmd = New ADODB.Command
                    With oCmd
                        .let_ActiveConnection(mP_Connection)
                        .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                        .CommandText = "UPDATEBATCHMSTINFO"
                        .Parameters.Append(.CreateParameter("@UNITCODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                        .Parameters.Append(.CreateParameter("@FROM_LOC", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 8, Trim(txtrecloc.Text)))
                        .Parameters.Append(.CreateParameter("@TO_LOC", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 8, Trim(txtdesloc.Text)))
                        .Parameters.Append(.CreateParameter("@ITEM_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(pstrItemCode)))
                        .Parameters.Append(.CreateParameter("@BATCH_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 25, strbatchno))
                        .Parameters.Append(.CreateParameter("@BATCH_DATE", ADODB.DataTypeEnum.adDBTimeStamp, ADODB.ParameterDirectionEnum.adParamInput, , strBatchDate))
                        .Parameters.Append(.CreateParameter("@BATCH_QTY", ADODB.DataTypeEnum.adDouble, ADODB.ParameterDirectionEnum.adParamInput, , dblBatchAcceptedQty))
                        .Parameters.Append(.CreateParameter("@UPDATE_FROM_LOC", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, , IIf(optGRIN.Checked = True, 0, 1)))
                        .Parameters.Append(.CreateParameter("@UPDATE_TO_LOC", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, , 1))
                        .Parameters.Append(.CreateParameter("@USER_ID", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 16, Trim(mP_User)))
                        .Parameters.Append(.CreateParameter("@MFG_DATE", ADODB.DataTypeEnum.adDBTimeStamp, ADODB.ParameterDirectionEnum.adParamInput, , MfgDate))
                        .Parameters.Append(.CreateParameter("@EXP_DATE", ADODB.DataTypeEnum.adDBTimeStamp, ADODB.ParameterDirectionEnum.adParamInput, , ExpDate))
                        .Parameters.Append(.CreateParameter("@ERRCODE", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamOutput))
                        .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End With
                    If oCmd.Parameters(oCmd.Parameters.Count - 1).Value <> 0 Then
                        MsgBox("Error while updating Batch Mst ", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100))
                        oCmd = Nothing
                        UpdateBatchInfo = False
                        GoTo ErrHandler
                    End If
                    oCmd = Nothing
                End If
                ''--- Rejected Quantities are Updated Batch Wise at 01J1 Location.
                If (dblBatchRejectedQty > 0 And optGRIN.Checked = True) Then
                    oCmd = New ADODB.Command
                    With oCmd
                        .let_ActiveConnection(mP_Connection)
                        .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                        .CommandText = "UPDATEBATCHMSTINFO"
                        .Parameters.Append(.CreateParameter("@UNITCODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                        .Parameters.Append(.CreateParameter("@FROM_LOC", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 8, Trim(Me.txtrecloc.Text)))
                        .Parameters.Append(.CreateParameter("@TO_LOC", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 8, "01J1"))
                        .Parameters.Append(.CreateParameter("@ITEM_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(pstrItemCode)))
                        .Parameters.Append(.CreateParameter("@BATCH_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 25, strbatchno))
                        .Parameters.Append(.CreateParameter("@BATCH_DATE", ADODB.DataTypeEnum.adDBTimeStamp, ADODB.ParameterDirectionEnum.adParamInput, , strBatchDate))
                        .Parameters.Append(.CreateParameter("@BATCH_QTY", ADODB.DataTypeEnum.adDouble, ADODB.ParameterDirectionEnum.adParamInput, , dblBatchRejectedQty))
                        .Parameters.Append(.CreateParameter("@UPDATE_FROM_LOC", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, , 0))
                        .Parameters.Append(.CreateParameter("@UPDATE_TO_LOC", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, , 1))
                        .Parameters.Append(.CreateParameter("@USER_ID", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 16, Trim(mP_User)))
                        .Parameters.Append(.CreateParameter("@MFG_DATE", ADODB.DataTypeEnum.adDBTimeStamp, ADODB.ParameterDirectionEnum.adParamInput, , MfgDate))
                        .Parameters.Append(.CreateParameter("@EXP_DATE", ADODB.DataTypeEnum.adDBTimeStamp, ADODB.ParameterDirectionEnum.adParamInput, , ExpDate))
                        .Parameters.Append(.CreateParameter("@ERRCODE", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamOutput))
                        .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End With
                    If oCmd.Parameters(oCmd.Parameters.Count - 1).Value <> 0 Then
                        MsgBox("Error while updating Batch Mst ", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100))
                        oCmd = Nothing
                        UpdateBatchInfo = False
                        GoTo ErrHandler
                    End If
                    oCmd = Nothing
                End If
            End If
        Next
        If strsql.Trim.Length > 0 Then mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        Exit Function
ErrHandler:
        UpdateBatchInfo = False
        oCmd = Nothing
        MsgBox(Err.Description & " - UBI [ " & Err.Number & " ] ", MsgBoxStyle.Information, ResolveResString(100))
    End Function
    Private Function GateEntryGrn_Req() As Boolean
        '****************************************************
        'Created By     -  Preety Jain
        'Description    -  Shall Check Whether Gate Entry for Grin is required
        'Arguments      -  None
        'Return Value   -  True If Gate Entry for Grin is required
        '****************************************************
        On Error GoTo ErrHandler
        Dim strsql As String
        Dim rstDB As ClsResultSetDB
        strsql = "Select GateEntryGrn_Req From Stores_Configmst WHERE UNIT_CODE = '" & gstrUNITID & "'"
        rstDB = New ClsResultSetDB
        Call rstDB.GetResult(strsql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rstDB.GetNoRows > 0 Then
            If rstDB.GetValue("GateEntryGrn_Req") = True Then
                GateEntryGrn_Req = True
            Else
                GateEntryGrn_Req = False
            End If
        End If
        rstDB.ResultSetClose()
        rstDB = Nothing
        Exit Function
ErrHandler:
        rstDB = Nothing
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function Inspection_Control_Detail() As Boolean
        '****************************************************
        'Created By     -  Preety Jain
        'Description    -  Shall Check Whether Inspection Control Details  is required
        'Arguments      -  None
        'Return Value   -  True If Inspection Control Details Flag in Store Config Master is true
        '****************************************************
        On Error GoTo ErrHandler
        Dim strsql As String
        Dim rstDB As ClsResultSetDB
        strsql = "Select Inspection_Control_Req From Stores_Configmst WHERE UNIT_CODE = '" & gstrUNITID & "'"
        rstDB = New ClsResultSetDB
        Call rstDB.GetResult(strsql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rstDB.GetNoRows > 0 Then
            If rstDB.GetValue("Inspection_Control_Req") = True Then
                Inspection_Control_Detail = True
            Else
                Inspection_Control_Detail = False
            End If
        End If
        rstDB.ResultSetClose()
        rstDB = Nothing
        Exit Function
ErrHandler:
        rstDB = Nothing
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function UpdateInspectionDetails(ByVal pstrItemCode As String) As Boolean '   ##,   **
        On Error GoTo ErrHandler
        Dim strsql As String
        UpdateInspectionDetails = True
        ''---- Insert data in Hdr Table
        strsql = "INSERT INTO inspection_entry_hdr" & vbCrLf
        strsql = strsql & " SELECT * FROM #inspection_entry_hdr"
        strsql = strsql & " WHERE UNIT_CODE = '" & gstrUNITID & "' AND item_code='" & Trim(pstrItemCode) & "'"
        strsql = strsql & " AND Grin_no = '" & ctlitem_c1.Text & "'"
        mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        ''---- Insert data in Dtl Table
        strsql = "INSERT INTO inspection_entry_dtl" & vbCrLf
        strsql = strsql & " SELECT * FROM #inspection_entry_dtl"
        strsql = strsql & " WHERE UNIT_CODE = '" & gstrUNITID & "' AND item_code='" & Trim(pstrItemCode) & "'"
        strsql = strsql & " AND Grin_no = '" & ctlitem_c1.Text & "'"
        mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        Exit Function
ErrHandler:
        UpdateInspectionDetails = False
        MsgBox(Err.Description & " - UID [ " & Err.Number & " ] ", MsgBoxStyle.Information, ResolveResString(100))
    End Function
    Private Function InspectionDetailsExist(ByVal pstrItemCode As String) As Boolean
        On Error GoTo ErrHandler
        Dim strsql As String
        Dim rstDB As ClsResultSetDB
        Dim rstDB1 As ClsResultSetDB
        Dim checkHdr, CheckDtl As Boolean
        rstDB = New ClsResultSetDB
        rstDB1 = New ClsResultSetDB
        strsql = " select item_code from #inspection_entry_hdr where UNIT_CODE = '" & gstrUNITID & "' AND" & _
            " item_code= '" & Trim(pstrItemCode) & "' and grin_no='" & ctlitem_c1.Text & "'"
        Call rstDB.GetResult(strsql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rstDB.GetNoRows > 0 Then
            'InspectionDetailsExist = True
            checkHdr = True
        Else
            'InspectionDetailsExist = False
            checkHdr = False
        End If
        strsql = " select item_code from #inspection_entry_dtl where UNIT_CODE = '" & gstrUNITID & "' AND" & _
            " item_code= '" & Trim(pstrItemCode) & "' and grin_no='" & ctlitem_c1.Text & "'"
        Call rstDB1.GetResult(strsql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rstDB1.GetNoRows > 0 Then
            CheckDtl = True
        Else
            CheckDtl = False
        End If
        rstDB.ResultSetClose()
        rstDB = Nothing
        rstDB1.ResultSetClose()
        rstDB1 = Nothing
        If checkHdr = True And CheckDtl = True Then
            InspectionDetailsExist = True
        Else
            InspectionDetailsExist = False
        End If
        Exit Function
ErrHandler:
        rstDB = Nothing
        rstDB1 = Nothing
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub CmdCertificates_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdCertificates.Click
        With frmSTRTRN0001a
            .ParentForm_Renamed = "FRMQATRN0001"
            If optGRIN.Checked = True Then
                .mlngDocNo = Val(ctlitem_c1.Text)
            Else
                .mlngDocNo = Val(txtRefGRINNo.Text)
            End If
            .Left = VB6.PixelsToTwipsX(Me.Left) + 100
            .Top = VB6.PixelsToTwipsY(Me.Top) + 500
            .ShowDialog()
        End With
    End Sub
    Private Function UpdateRGPJobOrderQty() As String
        '---------------------------------------------------------------------------------------
        'Name       :   UpdateRGPJobOrderQty
        'Type       :   Function
        'Author     :   Sunita Gupta
        'Return     :   0 if fail, 1 if Success
        'Purpose    :   Update the data in Accepted OK Qty against the JOB Order in GRIN_JOBORDER_DTL table
        '---------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim strsql As String
        Dim strGrinItem As String
        Dim curAcceptedQty As Double
        Dim objRSTemp As New ADODB.Recordset
        Dim objRsTemp1 As New ADODB.Recordset
        With sprdata
            For intCount = 1 To .MaxRows
                .Row = intCount
                .Col = Col_Item_Code
                strGrinItem = .Text.Trim
                .Row = intCount
                .Col = Col_Accepted_Qty
                curAcceptedQty = Val(.Text)
                strsql = "SELECT RGP_No FROM RGP_Grin_Details"
                strsql = strsql & " WHERE UNIT_CODE = '" & gstrUNITID & "' AND Grin_No = '" & Trim(ctlitem_c1.Text) & "'"
                If objRSTemp.State = ADODB.ObjectStateEnum.adStateOpen Then objRSTemp.Close()
                objRSTemp.Open(strsql, mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                If Not objRSTemp.EOF Then
                    Do While Not objRSTemp.EOF
                        strsql = "SELECT JobOrder_No, Accepted_Quantity,"
                        strsql = strsql & " ActualGRINOK_Quantity"
                        strsql = strsql & " FROM Grin_JobOrder_Dtl"
                        strsql = strsql & " WHERE UNIT_CODE = '" & gstrUNITID & "' AND" & _
                            " Doc_No = '" & Trim(ctlitem_c1.Text) & "'"
                        strsql = strsql & " AND Item_Code = '" & strGrinItem & "'"
                        strsql = strsql & " AND Location_Code  = '" & Trim(txtdesloc.Text) & "'"
                        strsql = strsql & " AND Accepted_Quantity > ActualGRINOK_Quantity"
                        strsql = strsql & " AND JobOrder_No In"
                        strsql = strsql & " (SELECT Pur_Order_No FROM RGP_Hdr WHERE UNIT_CODE = '" & gstrUNITID & "'" & _
                            " AND Doc_No = '" & objRSTemp.Fields("RGP_No").Value & "')"
                        strsql = strsql & " ORDER BY JobOrder_No"
                        If objRsTemp1.State = ADODB.ObjectStateEnum.adStateOpen Then objRsTemp1.Close()
                        objRsTemp1.Open(strsql, mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        Do While Not objRsTemp1.EOF
                            If curAcceptedQty > 0 Then
                                If objRsTemp1.Fields("Accepted_Quantity").Value <= curAcceptedQty Then
                                    strsql = "UPDATE Grin_JobOrder_Dtl"
                                    strsql = strsql & " SET ActualGRINOK_Quantity = " & objRsTemp1.Fields("Accepted_Quantity").Value
                                    strsql = strsql & " WHERE UNIT_CODE = '" & gstrUNITID & "' AND Doc_No = '" & Trim(ctlitem_c1.Text) & "'"
                                    strsql = strsql & " AND JobOrder_No = '" & objRsTemp1.Fields("JobOrder_No").Value & "'"
                                    strsql = strsql & " AND Item_Code = '" & strGrinItem & "'"
                                    strsql = strsql & " AND Location_Code IN (SELECT To_Location FROM Grn_Hdr WHERE UNIT_CODE = '" & gstrUNITID & "' AND Doc_No = '" & Trim(ctlitem_c1.Text) & "')"
                                    curAcceptedQty = curAcceptedQty - objRsTemp1.Fields("Accepted_Quantity").Value
                                Else
                                    strsql = "UPDATE Grin_JobOrder_Dtl"
                                    strsql = strsql & " SET ActualGRINOK_Quantity = " & curAcceptedQty
                                    strsql = strsql & " WHERE UNIT_CODE = '" & gstrUNITID & "' AND Doc_No = '" & Trim(ctlitem_c1.Text) & "'"
                                    strsql = strsql & " AND JobOrder_No = '" & objRsTemp1.Fields("JobOrder_No").Value & "'"
                                    strsql = strsql & " AND Item_Code = '" & strGrinItem & "'"
                                    strsql = strsql & " AND Location_Code IN (SELECT To_Location FROM Grn_Hdr WHERE UNIT_CODE = '" & gstrUNITID & "' AND" & _
                                        " Doc_No = '" & Trim(ctlitem_c1.Text) & "')"
                                    curAcceptedQty = 0
                                End If
                                mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            Else
                                Exit Do
                            End If
                            objRsTemp1.MoveNext()
                        Loop
                        objRSTemp.MoveNext()
                    Loop
                End If
            Next
        End With
        If objRSTemp.State = ADODB.ObjectStateEnum.adStateOpen Then objRSTemp.Close()
        If objRsTemp1.State = ADODB.ObjectStateEnum.adStateOpen Then objRsTemp1.Close()
        objRSTemp = Nothing
        objRsTemp1 = Nothing
        Exit Function
ErrHandler:
        objRSTemp = Nothing
        objRsTemp1 = Nothing
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function InspectionDtlEntryIconsLoad() As Object
        On Error GoTo Errorhandler
        '********************************************************************************************
        'Created By     -  Davinder Singh
        'Date           -  09-Oct-2006
        'Arguments      -  NIL
        'Return Value   -  NIL
        'Function       -  To Load the icons on the buttons in the frame for Inspection detail printing
        '***********************************************************************************
        cmdDocHlp.Image = My.Resources.ico111.ToBitmap
        cmdPrint.Image = My.Resources.ico231.ToBitmap
        cmdcansel.Image = My.Resources.ico230.ToBitmap
        fraPrint.Visible = False
        Exit Function
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub ShowGrinReport(Optional ByVal toPrinter As Integer = 0)
        On Error GoTo Errorhandler
        '********************************************************************************************
        'Created By     -  Davinder Singh
        'Date           -  14 Nov 2006
        'Function       -  To Display the report
        '***********************************************************************************
        Dim straddress As String
        Dim strQSNo As String
        Dim dtDocDate As Date
        Dim strSql As String
        Dim blnStatus As Boolean
        REPVIEWER = New eMProCrystalReportViewer
        REPDOC = New ReportDocument
        REPDOC = REPVIEWER.GetReportDocument()
        STRREPPATH = My.Application.Info.DirectoryPath & "\Reports\rptGrinPrinting.rpt"
        REPDOC.Load(STRREPPATH)
        With REPDOC
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.WaitCursor)
            straddress = gstr_WRK_ADDRESS1 & gstr_WRK_ADDRESS2
            .SetParameterValue("@unitcode", gstrUNITID)
            .SetParameterValue("@GRIN_NO", Trim(txtDocNo.Text))
            .DataDefinition.FormulaFields("CompanyName").Text = "'" & gstrCOMPANY & "'"
            .DataDefinition.FormulaFields("companyaddress").Text = "'" & straddress & "'"
            strSql = "SELECT ISNULL(EOU_FLAG,0) AS EOU_FLAG FROM COMPANY_MST WHERE UNIT_CODE = '" & gstrUNITID & "'"
            blnStatus = DataExist(strSql)
            If blnStatus = False Then
                .DataDefinition.FormulaFields("Display_RateAndValue").Text = "True"
            Else
                .DataDefinition.FormulaFields("Display_RateAndValue").Text = "False"
            End If
            If mblnBatchTracking = True Then
                .DataDefinition.FormulaFields("SuppressBatches").Text = "False"
            Else
                .DataDefinition.FormulaFields("SuppressBatches").Text = "True"
            End If
            strSql = "SELECT GRN_DATE FROM GRN_HDR WHERE UNIT_CODE = '" & gstrUNITID & "' AND DOC_NO = " & txtDocNo.Text
            dtDocDate = GetQryOutput(strSql)
            If QSRequired() = True Then
                'dtDocDate = CDate(VB6.Format(dtDocDate, gstrDateFormat))
                strQSNo = QSFormatNumber("rptGRINPrinting", dtDocDate)
                .DataDefinition.FormulaFields("QSFormatNo").Text = "'" & Trim(strQSNo) & "'"
            Else
                .DataDefinition.FormulaFields("QSFormatNo").Text = "'" & "" & "'"
            End If
            REPVIEWER.Zoom = 150
            If toPrinter = 1 Then
                REPVIEWER.SetReportDocument()
                REPDOC.PrintToPrinter(1, False, 0, 0)
            Else
                REPVIEWER.Show()
            End If

            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        End With
        Exit Sub
Errorhandler:
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Function InspectionDtlEntryEnable(ByRef Bln As Boolean) As Object
        On Error GoTo Errorhandler
        '********************************************************************************************
        'Created By     -  Davinder Singh
        'Date           -  09-Oct-2006
        'Arguments      -  A boolean variable according to which controls are Enabled/Disabled
        'Return Value   -  NIL
        'Function       -  To enable/Disable the controls when user want to print inspection entry report
        '***********************************************************************************
        fraPrint.Enabled = Bln
        cmdDocHlp.Enabled = Bln
        cmdPrint.Enabled = Bln
        cmdcansel.Enabled = Bln
        txtDocNo.Text = ""
        txtDocNo.Enabled = Bln
        cmdDocHlp.Enabled = Bln
        fraPrint.Visible = Bln
        If Bln = True Then
            txtDocNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            fraGRINType.Enabled = False
            txtrecloc.Enabled = False
            txtrecloc.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            cmdRecLocList.Enabled = False
            ctlitem_c1.Enabled = False
            ctlitem_c1.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            cmdHelp.Enabled = False
            dtpQADate.Enabled = False
            txtEmpQA.Enabled = False
            txtEmpQA.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            CmdHelpEmpQA.Enabled = False
            CmdCertificates.Enabled = False
            txtDocNo.Focus()
        Else
            txtDocNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            fraGRINType.Enabled = True
            txtrecloc.Enabled = True
            txtrecloc.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            cmdRecLocList.Enabled = True
            ctlitem_c1.Enabled = True
            ctlitem_c1.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            cmdHelp.Enabled = True
            dtpQADate.Enabled = True
            CmdCertificates.Enabled = True
            txtEmpQA.Enabled = True
            txtEmpQA.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        End If
        Exit Function
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function ShowPOWiseRejectionDtls(ByVal pItemCode As String, ByVal pItemDescription As String, ByVal pItemUOM As String, ByVal pRejectedQty As Double) As Object
        '********************************************************************************************
        'Created By     -  Davinder Singh
        'Date           -  09-Oct-2006
        'Arguments      -  A boolean variable according to which controls are Enabled/Disabled
        'Return Value   -  NIL
        'Function       -  To show the grid for taking the PO Wise rejection inputs
        '***********************************************************************************
        On Error GoTo Errorhandler
        Dim strsql As String
        Dim blnFloatFormat As Boolean
        Dim rsdb As ClsResultSetDB
        blnFloatFormat = False
        strsql = "SELECT GPD.PO_No,GPD.Challan_Qty,"
        strsql = strsql & " GPD.Actual_Qty,GPD.Excess_PO_Qty"
        strsql = strsql & " FROM Grn_PO_Dtl GPD"
        strsql = strsql & " WHERE GPD.UNIT_CODE = '" & gstrUNITID & "' AND GPD.Doc_No='" & ctlitem_c1.Text & "'"
        strsql = strsql & " AND GPD.Item_Code='" & pItemCode & "'"
        rsdb = New ClsResultSetDB
        If rsdb.GetResult(strsql) = False Then GoTo Errorhandler
        If rsdb.GetNoRows > 0 Then
            rsdb.MoveFirst()
            With fsPOWiseRejectionDtls
                .MaxRows = 0
                .MaxCols = 6
                .UnitType = FPSpreadADO.UnitTypeConstants.UnitTypeTwips
                .Font = Me.Font
                .set_RowHeight(.MaxRows, 300)
                Call .SetText(Enm_fsPOWiseRejectionDtls.Col_PO_No, 0, "PO No")
                Call .SetText(Enm_fsPOWiseRejectionDtls.Col_Challan_Quantity, 0, "Challan Qty")
                Call .SetText(Enm_fsPOWiseRejectionDtls.Col_Actual_Quantity, 0, "Actual Qty")
                Call .SetText(Enm_fsPOWiseRejectionDtls.Col_Excess_PO_Quantity, 0, "Excess PO Qty")
                Call .SetText(Enm_fsPOWiseRejectionDtls.Col_Accepted_Quantity, 0, "Accepted Qty")
                Call .SetText(Enm_fsPOWiseRejectionDtls.Col_Rejected_Quantity, 0, "Rejected Qty")
                If GetUOMDecimalPlacesAllowed(pItemUOM) > 0 Then
                    blnFloatFormat = True
                    While Not rsdb.EOFRecord
                        .MaxRows = .MaxRows + 1
                        .Row = .MaxRows
                        .set_RowHeight(.MaxRows, 300)
                        Call .SetText(Enm_fsPOWiseRejectionDtls.Col_PO_No, .MaxRows, rsdb.GetValue("PO_No"))
                        Call .SetText(Enm_fsPOWiseRejectionDtls.Col_Challan_Quantity, .MaxRows, VB6.Format(rsdb.GetValue("Challan_Qty"), "0.0000"))
                        Call .SetText(Enm_fsPOWiseRejectionDtls.Col_Actual_Quantity, .MaxRows, VB6.Format(rsdb.GetValue("Actual_Qty"), "0.0000"))
                        Call .SetText(Enm_fsPOWiseRejectionDtls.Col_Excess_PO_Quantity, .MaxRows, VB6.Format(rsdb.GetValue("Excess_PO_Qty"), "0.0000"))
                        .Col = Enm_fsPOWiseRejectionDtls.Col_Rejected_Quantity : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatDecimalPlaces = 4 : .TypeFloatMin = 0.0#
                        .TypeFloatMax = CDbl(VB6.Format(rsdb.GetValue("Actual_Qty"), "0.0000")) : .EditModePermanent = True : .EditModeReplace = True : .Text = "0.0000"
                        Call .SetText(Enm_fsPOWiseRejectionDtls.Col_Accepted_Quantity, .MaxRows, VB6.Format(rsdb.GetValue("Actual_Qty"), "0.0000"))
                        rsdb.MoveNext()
                    End While
                Else
                    While Not rsdb.EOFRecord
                        .MaxRows = .MaxRows + 1
                        .Row = .MaxRows
                        .set_RowHeight(.MaxRows, 300)
                        Call .SetText(Enm_fsPOWiseRejectionDtls.Col_PO_No, .MaxRows, rsdb.GetValue("PO_No"))
                        Call .SetText(Enm_fsPOWiseRejectionDtls.Col_Challan_Quantity, .MaxRows, VB6.Format(rsdb.GetValue("Challan_Qty"), "0"))
                        Call .SetText(Enm_fsPOWiseRejectionDtls.Col_Actual_Quantity, .MaxRows, VB6.Format(rsdb.GetValue("Actual_Qty"), "0"))
                        Call .SetText(Enm_fsPOWiseRejectionDtls.Col_Excess_PO_Quantity, .MaxRows, VB6.Format(rsdb.GetValue("Excess_PO_Qty"), "0"))
                        .Col = Enm_fsPOWiseRejectionDtls.Col_Rejected_Quantity : .CellType = FPSpreadADO.CellTypeConstants.CellTypeInteger : .TypeIntegerMin = 0
                        .TypeIntegerMax = CInt(VB6.Format(rsdb.GetValue("Actual_Qty"), "0")) : .EditModePermanent = True : .EditModeReplace = True : .Text = "0"
                        Call .SetText(Enm_fsPOWiseRejectionDtls.Col_Accepted_Quantity, .MaxRows, VB6.Format(rsdb.GetValue("Actual_Qty"), "0"))
                        rsdb.MoveNext()
                    End While
                End If
                .Row = 1
                .Col = Enm_fsPOWiseRejectionDtls.Col_Challan_Quantity
                .Row2 = .MaxRows
                .Col2 = .MaxCols
                .BlockMode = True
                .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
                .BlockMode = False
                .Row = 1
                .Col = Enm_fsPOWiseRejectionDtls.Col_PO_No
                .Row2 = .MaxRows
                .Col2 = .MaxCols
                .BlockMode = True
                .Lock = True
                .BlockMode = False
                .Row = 1
                .Col = Enm_fsPOWiseRejectionDtls.Col_Rejected_Quantity
                .Row2 = .MaxRows
                .Col2 = Enm_fsPOWiseRejectionDtls.Col_Rejected_Quantity
                .BlockMode = True
                .Lock = False
                .BlockMode = False
                .set_ColWidth(Enm_fsPOWiseRejectionDtls.Col_Challan_Quantity, 1200)
                .set_ColWidth(Enm_fsPOWiseRejectionDtls.Col_Actual_Quantity, 1200)
                .set_ColWidth(Enm_fsPOWiseRejectionDtls.Col_Excess_PO_Quantity, 1200)
                .set_ColWidth(Enm_fsPOWiseRejectionDtls.Col_Accepted_Quantity, 1200)
                .set_ColWidth(Enm_fsPOWiseRejectionDtls.Col_Rejected_Quantity, 1200)
                fraPoWiseRejectionDtl.Visible = True
                InspectionDtlEntryEnable(True)
                cmdPORejOk.Enabled = True
                cmdPORejCancel.Enabled = True
                lblItemCode.Text = pItemCode & " [" & pItemDescription & "]"
                If blnFloatFormat = True Then
                    lblTotalRejectedQty.Text = VB6.Format(pRejectedQty, "0.0000")
                    lblRemainingRejectedQty.Text = VB6.Format(pRejectedQty, "0.0000")
                    lblRejectedAgainstPO.Text = "0.0000"
                Else
                    lblTotalRejectedQty.Text = VB6.Format(pRejectedQty, "0")
                    lblRemainingRejectedQty.Text = VB6.Format(pRejectedQty, "0")
                    lblRejectedAgainstPO.Text = "0"
                End If
                .Row = 1
                .Col = Enm_fsPOWiseRejectionDtls.Col_Rejected_Quantity
                .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                .Focus()
            End With
        End If
        rsdb.ResultSetClose()
        rsdb = Nothing
        Exit Function
Errorhandler:
        rsdb = Nothing
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function GetPOList() As String
        '********************************************************************************************
        'Created By     -  Davinder Singh
        'Date           -  09-Oct-2006
        'Arguments      -  A boolean variable according to which controls are Enabled/Disabled
        'Return Value   -  NIL
        'Function       -  To get list of PO's in the Grin
        '***********************************************************************************
        On Error GoTo Errorhandler
        Dim strsql As String
        Dim strPOList As String
        Dim rsdb As ClsResultSetDB
        strPOList = String.Empty
        rsdb = New ClsResultSetDB
        strsql = "SELECT DISTINCT PO_No FROM Grn_PO_Dtl WHERE UNIT_CODE = '" & gstrUNITID & "' AND" & _
            " Doc_No='" & Trim(ctlitem_c1.Text) & "'"
        If rsdb.GetResult(strsql) = False Then GoTo Errorhandler
        If rsdb.GetNoRows > 0 Then
            rsdb.MoveFirst()
            While Not rsdb.EOFRecord
                strPOList = strPOList & "," & rsdb.GetValue("PO_No")
                rsdb.MoveNext()
            End While
            strPOList = Mid(strPOList, 2)
        End If
        rsdb.ResultSetClose()
        rsdb = Nothing
        GetPOList = strPOList
        Exit Function
Errorhandler:
        rsdb = Nothing
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function MultiplePO_Allowed() As Boolean
        '********************************************************************************************
        'Created By     -  Davinder Singh
        'Date           -  09-Oct-2006
        'Arguments      -  NIL
        'Return Value   -  NIL
        'Function       -  To check if Multiple PO Flag is on or not
        '***********************************************************************************
        On Error GoTo Errorhandler
        Dim strsql As String
        Dim blnStatus As Boolean
        strsql = "SELECT TOP 1 1 FROM Stores_ConfigMst WHERE UNIT_CODE = '" & gstrUNITID & "' AND MultiplePOAllowed =1"
        blnStatus = DataExist(strsql)
        MultiplePO_Allowed = blnStatus
        Exit Function
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub FillArrayWithPOdtls()
        '********************************************************************************************
        'Created By     -  Davinder Singh
        'Date           -  09-Oct-2006
        'Arguments      -  NIL
        'Return Value   -  NIL
        'Function       -  Fill PO wise details in array
        '***********************************************************************************
        On Error GoTo Errorhandler
        Dim rsdb As ClsResultSetDB
        Dim strsql As String
        Dim intArrIndex As Short
        strsql = "SELECT Item_Code,PO_No,Challan_Qty,Receipt_Qty,Actual_Qty,"
        strsql = strsql & "Excess_PO_Qty,item_Rate,Assesable_Rate,Discount_Amount,"
        strsql = strsql & "GL_Group_Code,project_code,Discount_Per"
        strsql = strsql & " FROM Grn_PO_Dtl WHERE UNIT_CODE = '" & gstrUNITID & "' AND Doc_No=" & ctlitem_c1.Text
        ReDim mArrPOdtl(0)
        rsdb = New ClsResultSetDB
        If rsdb.GetResult(strsql) = False Then GoTo Errorhandler
        If rsdb.GetNoRows > 0 Then
            rsdb.MoveFirst()
            While Not rsdb.EOFRecord
                intArrIndex = UBound(mArrPOdtl)
                intArrIndex = intArrIndex + 1
                ReDim Preserve mArrPOdtl(intArrIndex)
                mArrPOdtl(intArrIndex).ItemCode = Trim(rsdb.GetValue("Item_Code"))
                mArrPOdtl(intArrIndex).PONO = CInt(Trim(rsdb.GetValue("PO_No")))
                mArrPOdtl(intArrIndex).ChallanQty = CDbl(Trim(rsdb.GetValue("Challan_Qty")))
                mArrPOdtl(intArrIndex).ReceiptQty = CDbl(Trim(rsdb.GetValue("Receipt_Qty")))
                mArrPOdtl(intArrIndex).ActualQty = CDbl(Trim(rsdb.GetValue("Actual_Qty")))
                mArrPOdtl(intArrIndex).ExcessPOQty = CDbl(Trim(rsdb.GetValue("Excess_PO_Qty")))
                mArrPOdtl(intArrIndex).Rate_Renamed = Trim(rsdb.GetValue("item_Rate"))
                mArrPOdtl(intArrIndex).AssesableRate = CDbl(Trim(rsdb.GetValue("Assesable_Rate")))
                mArrPOdtl(intArrIndex).DiscountAmt = CDbl(Trim(rsdb.GetValue("Discount_Amount")))
                mArrPOdtl(intArrIndex).GLGrpCode = Trim(rsdb.GetValue("GL_Group_Code"))
                mArrPOdtl(intArrIndex).ProjectCode = Trim(rsdb.GetValue("project_code"))
                mArrPOdtl(intArrIndex).DiscountPer = CDbl(Trim(rsdb.GetValue("Discount_Per")))
                mArrPOdtl(intArrIndex).AcceptedQty = 0
                mArrPOdtl(intArrIndex).RejectedQty = 0
                rsdb.MoveNext()
            End While
        Else
            MsgBox("PO details not found for selected Grin No.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
        End If
        rsdb.ResultSetClose()
        rsdb = Nothing
        Exit Sub
Errorhandler:
        rsdb = Nothing
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub AutoFillPOwiseRejectionDtls(ByVal pItemCode As String, ByVal pDescription As String, ByVal pUOM As String, ByVal pRejectedQty As Double)
        '***************************************************************************
        'Creation Date  - 09-Oct-2006
        'Created By     - Davinder Singh
        'Issue No       - 17893
        'Description    - To fill the Rejection details in the array without showing input frame
        '***************************************************************************
        On Error GoTo Errorhandler
        Dim intCtrArr As Short
        For intCtrArr = 1 To UBound(mArrPOdtl) Step 1
            If StrComp(pItemCode, mArrPOdtl(intCtrArr).ItemCode, CompareMethod.Text) = 0 Then
                mArrPOdtl(intCtrArr).AcceptedQty = mArrPOdtl(intCtrArr).ActualQty - pRejectedQty
                mArrPOdtl(intCtrArr).RejectedQty = pRejectedQty
                mArrPOdtl(intCtrArr).UOM = pUOM
                mArrPOdtl(intCtrArr).Item_Desc = pDescription
            End If
        Next intCtrArr
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub FillRejReasonInArray()
        '***************************************************************************
        'Creation Date  - 09-Oct-2006
        'Created By     - Davinder Singh
        'Issue No       - 17893
        'Description    - To fill the Reason of Rejection in the array
        '***************************************************************************
        On Error GoTo Errorhandler
        Dim intCtrArr As Short
        Dim intCtrGrid As Short
        Dim strItemCode As String
        With sprdata
            For intCtrGrid = 1 To .MaxRows Step 1
                .Row = intCtrGrid
                .Col = Col_Item_Code
                strItemCode = .Text.Trim
                For intCtrArr = 1 To UBound(mArrPOdtl) Step 1
                    If StrComp(strItemCode, mArrPOdtl(intCtrArr).ItemCode, CompareMethod.Text) = 0 Then
                        .Col = Col_Rejection_Reason
                        mArrPOdtl(intCtrArr).Rej_reason = .Text.Trim
                    End If
                Next intCtrArr
            Next intCtrGrid
        End With
        Exit Sub
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Function ValidatePODtlArr() As Boolean
        '*******************************************************************************
        'Creation Date  - 03-Nov-2006
        'Created By     - Davinder Singh
        'Issue No       - 17893
        'Description    - To Check if PO wise dtls are equal to item wise dtls in grid
        '*******************************************************************************
        On Error GoTo Errorhandler
        Dim strItemCode As String
        Dim stritemdesc As String
        Dim strUOM As String
        Dim curGridRejectedQty As Double
        Dim curGridAcceptedQty As Double
        Dim curArrRejectedQty As Double
        Dim curArrAcceptedQty As Double
        Dim intArrCtr As Short
        Dim intGridCtr As Short
        ValidatePODtlArr = True
        With sprdata
            For intGridCtr = 1 To .MaxRows Step 1
                .Row = intGridCtr
                .Col = Col_Item_Code
                strItemCode = .Text.Trim
                .Col = Col_Item_Description
                stritemdesc = .Text.Trim
                .Col = Col_Item_UOM
                strUOM = .Text.Trim
                .Col = Col_Rejected_Qty
                curGridRejectedQty = CDbl(.Text.Trim)
                .Col = Col_Accepted_Qty
                curGridAcceptedQty = CDbl(.Text.Trim)
                curArrAcceptedQty = 0
                curArrRejectedQty = 0
                For intArrCtr = 1 To UBound(mArrPOdtl) Step 1
                    If StrComp(mArrPOdtl(intArrCtr).ItemCode, strItemCode, CompareMethod.Text) = 0 Then
                        curArrRejectedQty = curArrRejectedQty + mArrPOdtl(intArrCtr).RejectedQty
                        curArrAcceptedQty = curArrAcceptedQty + mArrPOdtl(intArrCtr).AcceptedQty
                    End If
                Next intArrCtr
                If ((curGridRejectedQty <> curArrRejectedQty) Or (curGridAcceptedQty <> curArrAcceptedQty)) Then
                    If mstrPOstatus = "A" Then
                        Call AutoFillPOwiseRejectionDtls(strItemCode, stritemdesc, strUOM, curGridRejectedQty)
                    ElseIf mstrPOstatus = "M" Then
                        mstrItemCode = strItemCode
                        mstrUOM = strUOM
                        mstrDescription = stritemdesc
                        MsgBox("PO wise dtls have not been entered for Item Code: " & strItemCode, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                        Call ShowPOWiseRejectionDtls(strItemCode, stritemdesc, strUOM, curGridRejectedQty)
                        ValidatePODtlArr = False
                        Exit Function
                    End If
                End If
            Next intGridCtr
        End With
        Exit Function
Errorhandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function GetExpBatchQty(ByVal pItem_code As String) As Double
        '-------------------------------------------------------------------------
        'Revised By     - Davinder Singh
        'Revision Date  - 14 Sep 2007
        'Description    - To get the qty associated with expired batches
        'ISSUE ID       - 21102
        '-------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim strQry As String
        strQry = "SELECT DBO.UFN_GETEXPIREDQTY('" & gstrUNITID & "','" & ctlitem_c1.Text.Trim & "','" & pItem_code & "','" & txtrecloc.Text.Trim & "','" & VB6.Format(GetServerDate(), "dd mmm yyyy") & "') AS EXPQTY"
        GetExpBatchQty = GetQryOutput(strQry)
        Exit Function
ErrHandler:
        GetExpBatchQty = -1
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function FillMFGOrEXP(ByVal pstrItemCode As String, ByVal pstrBatchNo As String, ByVal pintIndex As Short) As String
        '-------------------------------------------------------------------------
        'Revised By     - Davinder Singh
        'Revision Date  - 14 Sep 2007
        'Description    - To read the MFG. or EXP. date from the array according to the passed index
        'ISSUE ID       - 21102
        '-------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim intCounter As Short
        Dim strItemCode As String
        Dim strbatchno As String
        If UBound(garrBatchDetails, 2) > 0 Then
            For intCounter = 1 To UBound(garrBatchDetails, 2)
                strItemCode = garrBatchDetails(1, intCounter)
                strbatchno = garrBatchDetails(2, intCounter)
                If (StrComp(strItemCode, pstrItemCode, CompareMethod.Text) = 0 And StrComp(strbatchno, pstrBatchNo, CompareMethod.Text) = 0) Then
                    FillMFGOrEXP = garrBatchDetails(pintIndex, intCounter)
                End If
            Next
        End If
        Exit Function
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function PendingForSAN(ByVal plngGrnNo As Integer) As Boolean
        '-------------------------------------------------------------------------
        'Revised By     - Davinder Singh
        'Revision Date  - 14 Sep 2007
        'Description    - To check for Pending WIP entry for SAN
        'ISSUE ID       - 21102
        '-------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim strQry As String
        Dim Rs As ADODB.Recordset
        PendingForSAN = False
        strQry = "SELECT DBO.UFN_PENDINGFORSAN('" & gstrUNITID & "'," & plngGrnNo & ") AS Flag"
        Rs = mP_Connection.Execute(strQry)
        If Not (Rs.BOF And Rs.EOF) Then
            If Rs.Fields("flag").Value = True Then
                MsgBox("There is pending WIP Entry(s) for stock adjustment for this grin." & vbCrLf & "Cannot Authorize this Grin before stock adjustment completed.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                PendingForSAN = True
            End If
        End If
        Rs.Close()
        Rs = Nothing
        Exit Function
ErrHandler:
        Rs = Nothing
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function GetBarcodeProcessGRN() As Boolean
        '-------------------------------------------------------------------------
        'Revised By     - Davinder Singh
        'Revision Date  - 14 Sep 2007
        'Description    - To Check if Barcode Functionality is required for GRIN
        'ISSUE ID       - 21102
        '-------------------------------------------------------------------------
        On Error GoTo ErrHandler
        GetBarcodeProcessGRN = False
        If GetBarcode_Process() = True Then
            If GetBarcode_GRN() = True Then
                GetBarcodeProcessGRN = True
            End If
        End If
        Exit Function
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function BatchDetailsExist(ByVal pstrItemCode As String) As Boolean '   ##
        On Error GoTo ErrHandler
        Dim intCounter As Short
        Dim strItemCode As String
        With frmAuthBatchDetails.vaBatchDetails
            If UBound(garrBatchDetails, 2) > 0 Then
                For intCounter = 1 To UBound(garrBatchDetails, 2)
                    strItemCode = garrBatchDetails(1, intCounter)
                    If StrComp(strItemCode, pstrItemCode, CompareMethod.Text) = 0 Then
                        BatchDetailsExist = True
                        Exit For
                    End If
                Next
            End If
        End With
        Exit Function
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function IsItemGrnPendingForAuth() As Boolean
        '-------------------------------------------------------------------------
        'Created By     - Davinder Singh
        'Creation Date  - 23 Oct 2007
        'Description    - To Check if previous Grins with same items
        '                 pending for Authorization
        'ISSUE ID       - 21345
        '-------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim str_ItCode As String
        Dim oRs As New ADODB.Recordset
        Dim intCount As Short
        Dim strMSG As String
        Dim strDocNo As String
        Dim strsql As String
        intCount = 0
        IsItemGrnPendingForAuth = False
        strMSG = "Can't Authorize the GRIN." & vbCrLf & "As previous Grin(s) for the following Item(s) are pending for Authorization : " & vbCrLf
        With sprdata
            For intCount = 1 To .MaxRows
                .Row = intCount
                .Col = Col_Item_Code : str_ItCode = .Text.Trim
                If Len(Trim(str_ItCode)) <> 0 Then
                    strsql = "SELECT DBO.UFN_GRNPENDINGFORAUTH('" & gstrUNITID & "', " & Trim(ctlitem_c1.Text) & ",'" & str_ItCode & "') AS PENDGRINS"
                    oRs = mP_Connection.Execute(strsql)
                    If Len(oRs.Fields("PENDGRINS").Value) > 0 Then
                        strDocNo = oRs.Fields("PENDGRINS").Value
                        strMSG = strMSG & " [ " & str_ItCode & "]  : " & strDocNo & vbCrLf
                        IsItemGrnPendingForAuth = True
                    End If
                End If
            Next
        End With
        If IsItemGrnPendingForAuth = True Then
            MsgBox(strMSG, MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100))
        End If
        oRs.Close()
        oRs = Nothing
        Exit Function
ErrHandler:
        oRs = Nothing
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        IsItemGrnPendingForAuth = False
    End Function
    Private Function RaiseSCAR() As Boolean
        '--------------------------------------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Return Value  : Boolean
        ' Function      : To Check whether to raise SCAR or not
        ' Issue ID      : 21345
        ' Datetime      : 22 Oct 2007
        '--------------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim strItemCode As String
        Dim StrItemDesription As String
        Dim dblRejectedQty As Double
        Dim intmaxrows As Short
        Dim blnRaiseScar As Boolean
        Dim intLoop As Short
        Dim blnIsRejectionExists As Boolean

        
        RaiseSCAR = True

        '10949734 
        Dim blnSCARMandatory As Boolean = False
        blnSCARMandatory = IsRecordExists("SELECT TOP 1 1 FROM STORES_CONFIGMST (nolock) WHERE UNIT_CODE='" & gstrUNITID & "' AND SCAR_MANDATORY=1")

        mblnSaveScar = False
        blnIsRejectionExists = False
        If txtvencd.Text.Trim.Length = 0 Then Exit Function
        blnRaiseScar = RaiseScarForGrinType(mstrGRType.Value)
        If blnRaiseScar = False Then Exit Function
        If IsContinueIfOpenSCAR(CInt(Trim(ctlitem_c1.Text)), 10, "QC") = False Then
            RaiseSCAR = False
            Exit Function
        End If
        With frmScarDetails
            .TxtVendorCode.Text = txtvencd.Text.Trim
            .TxtVendorCode.Enabled = False
            .lblVendorName.Text = txtvend.Text.Trim
            .lblSource.Text = "GRIN (QC)"
            .DtpSCARDate.Value = GetServerDate()
            .sspr.MaxRows = 0
        End With
        intmaxrows = sprdata.MaxRows
        With sprdata
            For intLoop = 1 To intmaxrows
                .Row = intLoop
                .Col = Col_Item_Code
                strItemCode = .Text.Trim
                .Row = intLoop
                .Col = Col_Item_Description
                StrItemDesription = .Text.Trim
                .Row = intLoop
                .Col = Col_Rejected_Qty
                dblRejectedQty = Val(.Text)
                If dblRejectedQty > 0 Then
                    blnIsRejectionExists = True
                    If frmScarDetails.AddSCARDetails(strItemCode, StrItemDesription) = False Then
                        MsgBox("Error while showing SCAR details.Please try again !", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100))
                        frmScarDetails.Dispose()
                        RaiseSCAR = False
                        Exit Function
                    End If
                End If
            Next
        End With

        If blnIsRejectionExists = True Then
            'If MsgBox("Do you want to raise SCAR for the rejection entered", MsgBoxStyle.Question + MsgBoxStyle.YesNo, ResolveResString(100)) = MsgBoxResult.Yes Then
            FitToParent(frmScarDetails, Me)
            frmScarDetails.ShowDialog()
            If frmScarDetails.dicSCAR.Count = 0 Then
                If blnSCARMandatory = True Then
                    MsgBox("Please enter SCAR Details !", MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, ResolveResString(100))
                Else
                    MsgBox("No SCAR has been raised for the items having rejection", MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, ResolveResString(100))
                End If
                frmScarDetails.Dispose()

                '10949734 
                If blnSCARMandatory = True Then
                    RaiseSCAR = False
                End If
                '

                mblnSaveScar = False
            Else
                mblnSaveScar = True
            End If
            'Else
            '    frmScarDetails.Dispose()
            '    mblnSaveScar = False
            'End If
        End If
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        RaiseSCAR = False
    End Function
    Private Function SaveSCAR() As Boolean
        '--------------------------------------------------------------------------------------------------
        ' Author        : Davinder singh
        ' Return Value  : Boolean
        ' Function      : To Check whether to raise SCAR or not
        ' Issue ID      : 21345
        ' Datetime      : 22 Oct 2007
        '--------------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim strItem As Object
        Dim oCmdSCAR As ADODB.Command
        SaveSCAR = True
        If frmScarDetails.dicSCAR.Count > 0 Then
            For Each strItem In frmScarDetails.dicSCAR.Keys
                oCmdSCAR = New ADODB.Command
                With oCmdSCAR
                    .let_ActiveConnection(mP_Connection)
                    .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                    .CommandText = "INSERTSCAR_DTL"
                    .Parameters.Append(.CreateParameter("@UNITCODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                    .Parameters.Append(.CreateParameter("@VENDOR_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 8, Trim(frmScarDetails.TxtVendorCode.Text)))
                    .Parameters.Append(.CreateParameter("@ITEM_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, Trim(strItem)))
                    .Parameters.Append(.CreateParameter("@SCAROPENDATE", ADODB.DataTypeEnum.adDBTimeStamp, ADODB.ParameterDirectionEnum.adParamInput, , getDateForDB(frmScarDetails.DtpSCARDate.Value)))
                    .Parameters.Append(.CreateParameter("@REMARKS", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 100, frmScarDetails.dicSCAR.Item(strItem)))
                    .Parameters.Append(.CreateParameter("@SOURCE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 5, "QC"))
                    .Parameters.Append(.CreateParameter("@SOURCE_DOC_NO", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamInput, , Trim(ctlitem_c1.Text)))
                    .Parameters.Append(.CreateParameter("@ERRCODE", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamOutput))
                    .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                End With
                If oCmdSCAR.Parameters(oCmdSCAR.Parameters.Count - 1).Value <> 0 Then
                    MsgBox("Error while raising SCAR details", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100))
                    oCmdSCAR = Nothing
                    SaveSCAR = False
                End If
                oCmdSCAR = Nothing
            Next strItem
        End If
        frmScarDetails.Dispose()
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        SaveSCAR = False
    End Function
    Private Sub cmdGrpAuthorise1_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.cmdGrpAuthorise.ButtonClickEventArgs) Handles cmdGrpAuthorise1.ButtonClick
        On Error GoTo ErrHandler
        '---------------------------------------------------------------------
        ' Purpose   : This Method is used to Add/Edit/Delete/Save/Print
        '             the date of the Location Master.
        ' Parameter : A read only parameter, it gives you the index of the
        '             button clicked by the user.
        '---------------------------------------------------------------------
        Dim strItemCode As String
        Dim dblActualQty As Double
        Dim dblExcessPOQty As Double
        Dim dblAcceptedQty As Double
        Dim dblRejectedQty As Double
        Dim StrreaForrej As String
        Dim intLoop, intmaxrows As Short
        Dim strsql As String
        Dim lstrReturnStr As String
        Dim strInspectValue As String
        Dim rstDB As ClsResultSetDB
        Dim lobjSaveData As clsGrinInspection
        Dim IsRejected As Boolean = False
        Call InspectionDtlEntryEnable(False)
        Dim objCmd As ADODB.Command
        Select Case e.Button
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_AUTHORIZE
                '101161852
                If mblnGRINDeviationForm Then
                    If Not mblnIQARequired Then
                        MsgBox("Please first ON the flag [IQA_REQUIRED] in [STORES_CONFIGMST].", MsgBoxStyle.Information, ResolveResString(100))
                        Exit Sub
                    End If
                End If
                Dim strtime As String = GetServerDateTime()
                If mstrPOstatus <> "W" Then
                    If ValidatePODtlArr() = False Then
                        Exit Sub
                    End If
                End If
                '101161852
                If Not mblnGRINDeviationForm Then
                    If mblnIQARequired Then
                        With sprdata
                            For intCount = 1 To .MaxRows
                                .Row = intCount
                                .Col = Col_IQA
                                If (.CellType = FPSpreadADO.CellTypeConstants.CellTypeButton) Then ''Condition Added by praveen on 02 FEB 2021
                                    .Col = Col_Item_Code
                                    ''Start---Added by praveen on 08-04-2019 for not allowing Grin to clear if Inspection master not defined.
                                    strsql = "select top 1 1 From inspection_parameter_linkage_master where UNIT_CODE = '" & gstrUNITID & "' and item_code= '" & Convert.ToString(.Text).Trim() & "' AND Process = 'IQA'"
                                    If Not IsRecordExists(strsql) = True Then
                                        MessageBox.Show("Please define [INSPECTION PARAMETER LINKAGE MASTER] in IQA section for item code: " & Convert.ToString(.Text), ResolveResString(100), MessageBoxButtons.OK)
                                        Exit Sub
                                    End If
                                    ''END
                                    strsql = "Select Top 1 1 from IQA_Parameter_Dtl (NOLOCK) WHERE UNIT_CODE='" + gstrUNITID + "' AND  DOC_NO='" & strDoc & "' AND ITEM_CODE='" & Convert.ToString(.Text).Trim() & "'"
                                    If Not IsRecordExists(strsql) = True Then
                                        MessageBox.Show("Please enter IQA Parameter Detail for item code: " & Convert.ToString(.Text), ResolveResString(100), MessageBoxButtons.OK)
                                        Exit Sub
                                    End If

                                    strsql = "Select Top 1 1 from IQA_Sample_Dtl (NOLOCK) WHERE UNIT_CODE='" + gstrUNITID + "' AND  DOC_NO='" & strDoc & "' AND ITEM_CODE='" & Convert.ToString(.Text).Trim() & "'"
                                    If Not IsRecordExists(strsql) = True Then
                                        MessageBox.Show("Please enter IQA Sample Detail for item code: " & Convert.ToString(.Text), ResolveResString(100), MessageBoxButtons.OK)
                                        Exit Sub
                                    End If
                                End If
                                
                            Next
                        End With

                        strsql = "SELECT TOP 1 1 FROM IQA_Parameter_Dtl (NOLOCK)  WHERE UNIT_CODE='" + gstrUnitId + "' AND  DOC_NO='" & strDoc & "'  AND ISNULL(IQA_Status,'') = '2' OR ISNULL(IQA_Status,'') = '0' "
                        If IsRecordExists(strsql) = True Then
                            MessageBox.Show("IQA Parameters either not entered or deviated.", ResolveResString(100), MessageBoxButtons.OK)
                            Exit Sub
                        End If
                    End If
                End If

                ''Start---Added by praveen on 31 OCT 2019 for Checking the rejection value Entered if User Marked at least one parameter as NOT OK.
                Dim strtempitemCode As String
                Dim strtempitemCodeFinal As String
                strtempitemCodeFinal = ""
                If mblnIQARequired Then
                    If mblnGRINDeviationForm Then
                        With sprdata
                            For intCount = 1 To .MaxRows
                                .Row = intCount
                                .Col = Col_IQA
                                If (.CellType = FPSpreadADO.CellTypeConstants.CellTypeButton) Then ''Condition Added by praveen on 02 FEB 2021
                                    .Col = Col_Item_Code
                                    strtempitemCode = Convert.ToString(.Text).Trim()
                                    strsql = "Select Top 1 1 from IQA_Sample_Dtl (NOLOCK) WHERE UNIT_CODE='" + gstrUNITID + "' AND  DOC_NO='" & strDoc & "' AND ITEM_CODE='" & strtempitemCode & "' AND Status='Not Ok'"
                                    If IsRecordExists(strsql) = True Then
                                        .Col = Col_Rejected_Qty
                                        If Convert.ToDecimal(.Text) <= 0 Then
                                            strtempitemCodeFinal = strtempitemCodeFinal + strtempitemCode + ","
                                            .Col = Col_Accepted_Qty
                                            .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                                        End If
                                    End If
                                End If
                                
                            Next
                        End With
                        If strtempitemCodeFinal <> "" Then
                            strtempitemCodeFinal = strtempitemCodeFinal.Trim(",")
                            Dim result As Integer = MessageBox.Show("One of the Sample is marked as [NOT OK].Do you want to continue without entering Rejection Quantity for items: " + strtempitemCodeFinal, "Empro", MessageBoxButtons.YesNoCancel)
                            If result = DialogResult.Cancel Then
                                Exit Sub
                            ElseIf result = DialogResult.No Then
                                Exit Sub
                            ElseIf result = DialogResult.Yes Then
                                'Exit Sub
                            End If
                        End If
                        
                    Else
                        With sprdata
                            For intCount = 1 To .MaxRows
                                .Row = intCount
                                .Col = Col_IQA
                                If (.CellType = FPSpreadADO.CellTypeConstants.CellTypeButton) Then ''Condition Added by praveen on 02 FEB 2021
                                    .Col = Col_Item_Code
                                    strtempitemCode = Convert.ToString(.Text).Trim()
                                    strsql = "Select Top 1 1 from IQA_Sample_Dtl (NOLOCK) WHERE UNIT_CODE='" + gstrUNITID + "' AND  DOC_NO='" & strDoc & "' AND ITEM_CODE='" & strtempitemCode & "' AND Status='Not Ok'"
                                    If IsRecordExists(strsql) = True Then
                                        .Col = Col_Rejected_Qty
                                        If Convert.ToDecimal(.Text) <= 0 Then
                                            MessageBox.Show("One of the Sample is marked as [NOT OK].Please enter Rejection Quantity for item code: " & strtempitemCode, ResolveResString(100), MessageBoxButtons.OK)
                                            Exit Sub
                                        End If
                                    End If
                                End If
                              

                            Next
                        End With
                    End If
                    
                End If
                ''END

                If ValidateBeforeAuthorize() Then
                    If RestrictInspectionForPendingItemGrin() = True Then
                        If IsItemGrnPendingForAuth() = True Then
                            Exit Sub
                        End If
                    End If
                    If mblnBarcodeGRN Then
                        If PendingForSAN(CInt(Trim(ctlitem_c1.Text))) = True Then Exit Sub
                    End If
                    If RaiseSCAR() = False Then Exit Sub
                    Call ResetDatabaseConnection()
                    mP_Connection.BeginTrans()
                    mP_Connection.Execute("set dateformat 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    If mblnSaveScar = True Then
                        If SaveSCAR() = False Then GoTo ErrHandler
                    End If
                    '10517328
                    '    strsql = "UPDATE Grn_hdr" & _
                    '" SET QA_Authorized_Code = '" & Trim(mP_User) & "'," & _
                    '" QA_Date ='" & Format(dtpQADate.Value, "dd MMM yyyy") & "'," & _
                    '" Upd_dt=getdate(),Upd_Userid='" & mP_User & "'" & _
                    '" WHERE UNIT_CODE = '" & gstrUNITID & "' AND Doc_no = '" & ctlitem_c1.Text & "'" & _
                    '" AND From_Location='" & txtrecloc.Text.Trim & "'"
                    '    mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

                    strsql = " UPDATE Grn_hdr" & _
                           " SET QA_Authorized_Code = '" & Trim(mP_User) & "'," & _
                           " QA_Date ='" & getDateForDB(GetServerDate()) & "'," & _
                           " Upd_dt=getdate(),Upd_Userid='" & mP_User & "',GRN_IQA_DEVIATION=" & IIf(mblnGRINDeviationForm, 1, 0) & "" & _
                           " WHERE UNIT_CODE = '" & gstrUnitId & "' AND Doc_no = '" & ctlitem_c1.Text & "'" & _
                           " AND From_Location='" & txtrecloc.Text.Trim & "'"
                    mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

                    With sprdata
                        intmaxrows = .MaxRows
                        For intLoop = 1 To intmaxrows
                            .Row = intLoop
                            .Col = Col_Item_Code : strItemCode = .Text.Trim
                            .Col = Col_Actual_Qty : dblActualQty = Val(.Text)
                            .Col = Col_Excess_PO_Qty : dblExcessPOQty = Val(.Text)
                            .Col = Col_Accepted_Qty : dblAcceptedQty = Val(.Text)
                            .Col = Col_Rejected_Qty : dblRejectedQty = Val(.Text)
                            .Col = Col_Rejection_Reason : StrreaForrej = .Text.Trim
                            .Col = Col_Inspection_Entry : strInspectValue = .Text.Trim
                            Call QuoteString(StrreaForrej)
                            If Val(dblRejectedQty) > 0 And IsRejected = False Then
                                IsRejected = True
                            End If
                            If validate_data(strItemCode, dblActualQty, dblExcessPOQty, dblAcceptedQty, dblRejectedQty, StrreaForrej) = False Then
                                GoTo ErrHandler
                            End If
                            If Defects_Details_Insertion(strItemCode) = False Then
                                GoTo ErrHandler
                            End If
                            If mblnBatchTracking = True Then
                                If UpdateBatchInfo(strItemCode) = False Then
                                    GoTo ErrHandler
                                End If
                            End If
                            If mbln_Inspection_Control_Details = True And optGRIN.Checked = True And strInspectValue = "1" Then
                                If UpdateInspectionDetails(strItemCode) = False Then
                                    GoTo ErrHandler
                                End If
                            End If
                            'CHANGES AGAINST ISSUE ID : 10378778
                            objCmd = New ADODB.Command
                            With objCmd
                                .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                                .ActiveConnection = mP_Connection
                                .CommandText = "USP_ENABLE_GLOBAL_TOOL_MAPPING"
                                .CommandTimeout = 0
                                .Parameters.Append(.CreateParameter("@UNIT_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUnitId))
                                .Parameters.Append(.CreateParameter("@TOOL_ITEM_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 16, strItemCode))
                                .Parameters.Append(.CreateParameter("@USER_ID", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 16, mP_User))
                                .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            End With
                            'END OF CHANGES
                        Next
                    End With
                    If mstrPOstatus <> "W" Then
                        ''--- Revert Schedules updated by Grin
                        If RevertPurchaseSchedule(txtrecloc.Text.Trim, CLng(ctlitem_c1.Text), False) = False Then
                            Exit Sub
                        End If
                        ''---- Purchase Schedule KnockOff
                        If KnockOffPurchaseSchedule(txtrecloc.Text.Trim, CLng(ctlitem_c1.Text)) = False Then
                            Exit Sub
                        End If
                    End If
                    If CheckForvalidAcceptance(ctlitem_c1.Text) = -1 Then
                        MsgBox("Authorize the Grin again !", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100))
                        GoTo ErrHandler
                    End If
                    ''---- To fill the reason of rejection in the array by picking it from grid
                    If mstrPOstatus <> "W" Then
                        Call FillRejReasonInArray()
                    End If
                    ''---- Prepare the strings for sending data to Grin inspection class to make PV in accounts
                    Call CreateDataString()
                    lobjSaveData = New clsGrinInspection
                    ''---- Sending the data to the Grin Inspection class and catching the returned status
                    lstrReturnStr = lobjSaveData.SaveGrin(mstrMasterString, mstrDetailString, True, mP_User)
                    If VB.Left(lstrReturnStr, 1) <> "Y" Then
                        mP_Connection.RollbackTrans()
                        MsgBox(VB.Right(lstrReturnStr, Len(lstrReturnStr) - 4), MsgBoxStyle.Critical, ResolveResString(100))
                        Exit Sub
                    Else
                        Call UpdateRGPJobOrderQty()
                    End If
                    'Added by Vinod for DTF Transaction
                    If gblnBarcodeProcess And DTF_ISSUE_ALLOWED() Then
                        If KNOCKOFF_DTF_TRANSACTION() = False Then
                            'mP_Connection.RollbackTrans()
                            Exit Sub
                        End If
                    End If
                    'End of Addition
                    'CHANGES AGAINST ISSUE ID : 10378778
                    objCmd = New ADODB.Command
                    With objCmd
                        .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                        .ActiveConnection = mP_Connection
                        .CommandText = "USP_SAVE_GLOBAL_TOOL_ITEM_GRIN"
                        .CommandTimeout = 0
                        .Parameters.Append(.CreateParameter("@UNIT_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUnitId))
                        .Parameters.Append(.CreateParameter("@GRIN_NO", ADODB.DataTypeEnum.adBigInt, ADODB.ParameterDirectionEnum.adParamInput, , ctlitem_c1.Text))
                        .Parameters.Append(.CreateParameter("@FROM_LOCATION", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, Me.txtrecloc.Text))
                        .Parameters.Append(.CreateParameter("@USER_ID", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 16, mP_User))
                        .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End With
                    'END OF CHANGES

                    '*SAVE ESI DETAILS, ISSUE ID : 10857384
                    SaveESIDetails()
                    '*

                    ''*AUTO CONSUME SERVICE GRIN STOCK, ISSUE ID : 10857384
                    'objCmd = New ADODB.Command
                    'With objCmd
                    '    .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                    '    .ActiveConnection = mP_Connection
                    '    .CommandText = "USP_KNOCKOFF_SERVICE_GRIN"
                    '    .CommandTimeout = 0
                    '    .Parameters.Append(.CreateParameter("@UNIT_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                    '    .Parameters.Append(.CreateParameter("@FROM_LOCATION", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, Me.txtrecloc.Text))
                    '    .Parameters.Append(.CreateParameter("@GRIN_NO", ADODB.DataTypeEnum.adBigInt, ADODB.ParameterDirectionEnum.adParamInput, , ctlitem_c1.Text))
                    '    .Parameters.Append(.CreateParameter("@USER_ID", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 16, mP_User))
                    '    .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    'End With
                    ''*
                    'auto mtn transfer


                    mP_Connection.CommitTrans()

                    Call Logging_Starting_End_Time("GRIN Inspection", strtime, "Saved", ctlitem_c1.Text)
                    If optGRIN.Checked = True Then
                        If MsgBox("GRIN Successfully Authorized.", MsgBoxStyle.Information, ResolveResString(100)) = MsgBoxResult.Ok Then
                            Dim intGrinNo As Integer
                            If Val(ctlitem_c1.Text) > 0 Then
                                intGrinNo = ctlitem_c1.Text
                            End If
                            ''auto mtn transfer
                            AutoMTNforJobworkGRN()

                            gblnCancelUnload = False
                            gblnFormAddEdit = False
                            With sprdata
                                .Row = 1
                                .Col = Col_Item_UOM
                                .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                            End With
                            cmdGrpAuthorise1.Enabled(0) = False
                            Call RefreshFrm()
                            optGRIN.Checked = True
                            Call optGRIN_CheckedChanged(optGRIN, New System.EventArgs())
                            ViewGrinAuthReport(intGrinNo)
                            If IsRejected = True Then
                                Using SqlCmd As New SqlCommand()
                                    With SqlCmd
                                        .CommandText = "USP_MATERIALREJECTION_AUTOMAILER"
                                        .CommandType = CommandType.StoredProcedure
                                        .CommandTimeout = 0
                                        .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUnitId
                                        .Parameters.Add("@DOC_NO", SqlDbType.VarChar, 10).Value = intGrinNo
                                        .Parameters.Add("@ERRMSG", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
                                        SqlConnectionclass.ExecuteNonQuery(SqlCmd)
                                    End With

                                    If SqlCmd.Parameters(SqlCmd.Parameters.Count - 1).Value.ToString.Trim.Length <> 0 Then
                                        MessageBox.Show("Error while sending the Mail !", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                                        Exit Sub
                                    Else
                                        MessageBox.Show("Rejection alert mail has been sent to the concerned person.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                                    End If
                                End Using
                            End If
                            Exit Sub
                        End If
                    ElseIf optSuppGRIN.Checked = True Then
                        If MsgBox("Supplementry GRIN Successfully Authorized.", MsgBoxStyle.Information, ResolveResString(100)) = MsgBoxResult.Ok Then
                            Dim intGrinNo As Integer
                            If Val(ctlitem_c1.Text) > 0 Then
                                intGrinNo = ctlitem_c1.Text
                            End If
                            gblnCancelUnload = False
                            gblnFormAddEdit = False
                            With sprdata
                                .Row = 1
                                .Col = Col_Item_UOM
                                .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                            End With
                            cmdGrpAuthorise1.Enabled(0) = False
                            Call RefreshFrm()
                            optGRIN.Checked = True
                            Call optGRIN_CheckedChanged(optGRIN, New System.EventArgs())
                            If IsRejected = True Then
                                Using SqlCmd As New SqlCommand()
                                    With SqlCmd
                                        .CommandText = "USP_MATERIALREJECTION_AUTOMAILER"
                                        .CommandType = CommandType.StoredProcedure
                                        .CommandTimeout = 0
                                        .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUnitId
                                        .Parameters.Add("@DOC_NO", SqlDbType.VarChar, 10).Value = intGrinNo
                                        .Parameters.Add("@ERRMSG", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
                                        SqlConnectionclass.ExecuteNonQuery(SqlCmd)
                                    End With

                                    If SqlCmd.Parameters(SqlCmd.Parameters.Count - 1).Value.ToString.Trim.Length <> 0 Then
                                        MessageBox.Show("Error while sending the Mail !", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                                        Exit Sub
                                    Else
                                        MessageBox.Show("Rejection alert mail has been sent to the concerned person.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                                    End If
                                End Using
                            End If
                            Exit Sub
                        End If
                    End If
                    Call Logging_Starting_End_Time("QA GRIN Inspection", strtime, "Authorize", ctlitem_c1.Text)
                Else
                    mstrErrorDesc = mstrErrorCaption & vbCrLf & mstrErrorDesc
                    If MsgBox(mstrErrorDesc, MsgBoxStyle.OkOnly, ResolveResString(100)) = MsgBoxResult.Ok Then
                        mctlError.Focus()
                    End If
                    gblnCancelUnload = True
                    gblnFormAddEdit = True
                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT
                InspectionDtlEntryEnable(True)
                If ctlitem_c1.Text.Trim.Length > 0 Then
                    If mblnShowInpectionGRINFRame = False Then
                        fraPrint.Visible = False
                        txtDocNo.Text = ctlitem_c1.Text
                        Call cmdPrint_Click(cmdPrint, New System.EventArgs())
                        InspectionDtlEntryEnable(False)
                    End If
                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_REFRESH
                If MsgBox("Warning! Refresh Will Undo All Changes Done. Proceed?", MsgBoxStyle.YesNo, ResolveResString(100)) = MsgBoxResult.Yes Then
                    Call RefreshFrm()
                    Call optGRIN_CheckedChanged(optGRIN, New System.EventArgs())
                    Exit Sub
                Else
                    ctlitem_c1.Focus()
                    Exit Sub
                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                ReDim marrDefectDetails(0, 0)
                If mbln_Inspection_Control_Details = True Then
                    mP_Connection.BeginTrans()
                    strsql = "Drop table #Inspection_Entry_Hdr"
                    mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    strsql = "Drop table #Inspection_Entry_Dtl"
                    mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    mP_Connection.CommitTrans()
                End If
                Me.Close()
        End Select
        Exit Sub
ErrHandler:
        rstDB = Nothing
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub SetGridHdrs()
        On Error GoTo ErrHandler
        With Me.sprdata
            .set_ColWidth(Col_Item_Code, 1560)
            .SetText(Col_Item_Code, 0, "Item Code")

            .set_ColWidth(Col_Item_Description, 2300)
            .SetText(Col_Item_Description, 0, "Deascription")

            .set_ColWidth(Col_Item_UOM, 555)
            .SetText(Col_Item_UOM, 0, "UOM")

            .set_ColWidth(Col_Item_Rate, 810)
            .SetText(Col_Item_Rate, 0, "Rate")

            .set_ColWidth(Col_Challan_Qty, 1125)
            .SetText(Col_Challan_Qty, 0, "Challan Qty.")

            .set_ColWidth(Col_Actual_Qty, 1125)
            .SetText(Col_Actual_Qty, 0, "Actual Qty.")

            .set_ColWidth(Col_Excess_PO_Qty, 1830)
            .SetText(Col_Excess_PO_Qty, 0, "Excess To PO/Schedule")

            .set_ColWidth(Col_Accepted_Qty, 1230)
            .SetText(Col_Accepted_Qty, 0, "Accepted Qty.")

            .set_ColWidth(Col_Rejected_Qty, 1125)
            .SetText(Col_Rejected_Qty, 0, "Rejected Qty.")

            .set_ColWidth(Col_Rejection_Reason, 1725)
            .SetText(Col_Rejection_Reason, 0, "Reason Of Rejection")

            .set_ColWidth(Col_Inspection_Entry, 790)
            .SetText(Col_Inspection_Entry, 0, "Insp. Ent.")

            .set_ColWidth(Col_Asseccible_Rate, 300)
            .SetText(Col_Asseccible_Rate, 0, "Assessible Rate")

            .set_ColWidth(Col_GL_Group, 300)
            .SetText(Col_GL_Group, 0, "GL Group")

            .set_ColWidth(Col_Project_Code, 300)
            .SetText(Col_Project_Code, 0, "Project Code")

            .set_ColWidth(Col_Discount_Per, 1000)
            .SetText(Col_Discount_Per, 0, "Discount Per")

            .set_ColWidth(Col_Rejection_Details, 1300)
            .SetText(Col_Rejection_Details, 0, "Defect Details")

            .set_ColWidth(Col_Batch_Details, 1200)
            .SetText(Col_Batch_Details, 0, "Batch Details")

            '*10857384
            .set_ColWidth(Col_ESI, 1000)
            .SetText(Col_ESI, 0, "ESI Detail")

            .set_ColWidth(Col_IQA, 1200)
            .SetText(Col_IQA, 0, "IQA Required")
            '.SetText(21, 0, "Item Doc")
            '.SetText(22, 0, "Item Doc")
            ''.SetText(Col_Receipt_Qty, 0, "")
            ''.SetText(col_Inspection_Control_Details, 0, "")
        End With
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub sprdata_ButtonClicked(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ButtonClickedEvent) Handles sprdata.ButtonClicked
        '***************************************************************************
        'Revised By     - Davinder Singh
        'Revision Date  - 14 Sep 2007
        'Issue No       - 21102
        'Description    - Validation added to Reject the qty associated with
        '                 Expired Batches
        '***************************************************************************
        Dim strActualQty As Integer = 0
        Dim strItemCode As String
        Dim strSql As String
        Dim blnStatus As Boolean
        Dim stritemdesc As String
        Dim strMeasureCode As String
        Dim strvendorCode As String
        Dim dblAcceptedQty As Double
        Dim dblRejectedQty As Double
        Dim dblActualQty As Double
        Dim dbladjstQty As Double
        Dim dblTotRejectedQty As Double
        Dim Intcounter As Integer
        Dim lngLotQty As Long
        Dim checkItemConditions As Boolean
        Dim curExpQty As Double
        Try
            Call SetInspectionEvironment()
            Select Case e.col
                Case Col_Batch_Details
                    With sprdata
                        .Col = Col_Item_Code : .Row = .ActiveRow : strItemCode = .Text.Trim
                        .Col = Col_Item_Description : .Row = .ActiveRow : stritemdesc = " [" & .Text.Trim & "]"
                        .Col = Col_Item_UOM : .Row = .ActiveRow : strMeasureCode = .Text.Trim : mstrMeasureCode = .Text
                        .Col = Col_Accepted_Qty : .Row = .ActiveRow : dblAcceptedQty = Val(.Text)
                        .Col = Col_Rejected_Qty : .Row = .ActiveRow : dblRejectedQty = Val(.Text)
                        If (dblAcceptedQty = 0 And dblRejectedQty = 0) Then
                            MsgBox("First Enter Accepted and/or Rejected Quantity for Item :" & strItemCode, vbInformation, ResolveResString(100))
                            .Row = .ActiveRow : .Col = Col_Accepted_Qty : .EditModePermanent = True : .EditModeReplace = True : .Focus() : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : Exit Sub
                        Else
                            If optGRIN.Checked = True Then
                                curExpQty = GetExpBatchQty(strItemCode)
                                If curExpQty > 0 Then
                                    If dblRejectedQty < curExpQty Then
                                        MsgBox("You have to Reject [" & curExpQty & "] Qty. for Item Code [" & strItemCode & "] as this Qty. is associated with Expired Batches", vbInformation + vbOKOnly, ResolveResString(100))
                                        .Focus()
                                        .Row = .ActiveRow
                                        .Col = Col_Accepted_Qty
                                        .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                        Exit Sub
                                    End If
                                End If
                            End If
                            With frmAuthBatchDetails
                                .Left = Me.Left + 100
                                .Top = Me.Top + 100
                                .mstrItemCode = strItemCode
                                .mstrItemDesc = stritemdesc
                                .mstrMeasureCode = strMeasureCode
                                .mintDocType = 10
                                .mlngDocNo = Trim(Me.ctlitem_c1.Text)
                                .mstrVendorcode = Trim$(txtvencd.Text)
                                .mstrGrnType = IIf(optGRIN.Checked = True, "G", "S")
                                Call FillGRINBatchDetails(strItemCode, dblAcceptedQty)
                                .mdblTotalAcceptedtQty = dblAcceptedQty
                                .mdblTotalRejectedQty = dblRejectedQty
                                If BatchDetailsExist(strItemCode) Then
                                    .mdblAdjustedAcceptedQty = BatchAcceptedOrRejectedQty(4, strItemCode)
                                    .mdblAdjustedRejectedQty = BatchAcceptedOrRejectedQty(5, strItemCode)
                                Else
                                    .mdblAdjustedAcceptedQty = 0
                                    With .vaBatchDetails
                                        For Intcounter = 1 To .MaxRows
                                            .Row = Intcounter : .Col = 5 : dbladjstQty = dbladjstQty + Val(.Text)
                                        Next
                                    End With
                                    .mdblAdjustedRejectedQty = dbladjstQty
                                End If
                                .ctlFormHeader1.HeaderFontSize = 3
                                .ctlFormHeader1.HeaderString = "For Item : " & .mstrItemCode & " " & .mstrItemDesc
                                If GetUOMDecimalPlacesAllowed(.mstrMeasureCode) > 0 Then
                                    .lblTotalAccepted.Text = Format(.mdblTotalAcceptedtQty, "0.0000")
                                    .lblTotalRejected.Text = Format(.mdblTotalRejectedQty, "0.0000")
                                    .lblAdjustedAccepted.Text = Format(.mdblAdjustedAcceptedQty, "0.0000")
                                    .lblAdjustedRejected.Text = Format(.mdblAdjustedRejectedQty, "0.0000")
                                    .lblBalanceAccepted.Text = Format(.mdblTotalAcceptedtQty - .mdblAdjustedAcceptedQty, "0.0000")
                                    .lblBalanceRejected.Text = Format(.mdblTotalRejectedQty - .mdblAdjustedRejectedQty, "0.0000")
                                Else
                                    .lblTotalAccepted.Text = Format(.mdblTotalAcceptedtQty, "0")
                                    .lblTotalRejected.Text = Format(.mdblTotalRejectedQty, "0")
                                    .lblAdjustedAccepted.Text = Format(.mdblAdjustedAcceptedQty, "0")
                                    .lblAdjustedRejected.Text = Format(.mdblAdjustedRejectedQty, "0")
                                    .lblBalanceAccepted.Text = Format(.mdblTotalAcceptedtQty - .mdblAdjustedAcceptedQty, "0")
                                    .lblBalanceRejected.Text = Format(.mdblTotalRejectedQty - .mdblAdjustedRejectedQty, "0")
                                End If
                                .ShowDialog()
                                ''DoEvents()
                                Exit Sub
                            End With
                        End If
                    End With
                Case Col_Rejection_Details
                    With Me
                        With .sprdata
                            .Row = e.row : .Col = Col_Item_Code : strItemCode = .Text.Trim : mstrItemCode = strItemCode
                            .Row = e.row : .Col = Col_Item_Description : stritemdesc = .Text.Trim
                            .Row = e.row : .Col = Col_Item_UOM : strMeasureCode = .Text.Trim : mstrMeasureCode = .Text.Trim
                            .Row = e.row : .Col = Col_Actual_Qty : dblActualQty = Val(.Text)
                            .Row = e.row : .Col = Col_Rejected_Qty : dblRejectedQty = Val(.Text)
                        End With
                        fraDefectDetails.Visible = True
                        .lblDefectDetails.Text = "Defect Details for the Item: " & strItemCode _
                        & " [" & stritemdesc & "]"
                        If mstrGrinType = "U" Or mstrGRType.Value = "E" Then
                            .lblTotalRej.Text = dblActualQty
                            .lblDefRejection.Text = "0"
                            .lblBalRejQty.Text = dblActualQty
                        Else
                            .lblTotalRej.Text = dblRejectedQty
                            .lblDefRejection.Text = "0"
                            .lblBalRejQty.Text = dblRejectedQty
                        End If
                        cmdOK.Enabled = True : cmdCancel.Enabled = True
                        strSql = "SELECT TOP 1 1 FROM DEFECT_MST WHERE UNIT_CODE = '" & gstrUNITID & "'"
                        blnStatus = DataExist(strSql)
                        If blnStatus = True Then
                            Call FillDefects(strItemCode, strMeasureCode)
                            If GetUOMDecimalPlacesAllowed(strMeasureCode) > 0 Then
                                .lblTotalRej.Text = Format(Val(.lblTotalRej.Text), "#0.0000")
                                .lblDefRejection.Text = Format(Val(.lblDefRejection.Text), "#0.0000")
                                .lblBalRejQty.Text = Format(Val(.lblBalRejQty.Text), "#0.0000")
                            Else
                                .lblTotalRej.Text = Format(Val(.lblTotalRej.Text), "#0")
                                .lblDefRejection.Text = Format(Val(.lblDefRejection.Text), "#0")
                                .lblBalRejQty.Text = Format(Val(.lblBalRejQty.Text), "#0")
                            End If
                            dblTotRejectedQty = Val(Me.lblTotalRej.Text)
                            With vaDefects
                                If GetUOMDecimalPlacesAllowed(strMeasureCode) > 0 Then
                                    .Row = 1 : .Col = Col_Defect_qty : .TypeFloatMax = dblTotRejectedQty : .TypeFloatMin = 0.0# : .TypeFloatDecimalPlaces = 4 : .EditModePermanent = True : .EditModeReplace = True : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus() : Exit Sub
                                Else
                                    .Row = 1 : .Col = Col_Defect_qty : .TypeFloatMax = dblTotRejectedQty : .TypeFloatMin = 0 : .TypeFloatDecimalPlaces = 0 : .EditModePermanent = True : .EditModeReplace = True : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus() : Exit Sub
                                End If
                            End With
                        Else
                            MsgBox("No Defect Code(s) are defined in the System." & vbCrLf _
                            & "First Define Defect Code(s).", vbInformation, ResolveResString(100))
                            cmdCancel.Focus() : Exit Sub
                        End If
                    End With
                Case col_Inspection_Control_Details
                    Dim strUOM As String
                    strvendorCode = txtvend.Text
                    With Me.sprdata
                        .Col = Col_Item_Code : .Row = .ActiveRow : strItemCode = .Text.Trim
                        .Col = Col_Item_Description : .Row = .ActiveRow : stritemdesc = " [" & .Text.Trim & "]"
                        .Col = Col_Actual_Qty : .Row = .ActiveRow : lngLotQty = Val(.Text)
                        .Col = 3 : .Row = .ActiveRow : strUOM = .Text
                        With frmQATRN0006
                            .Left = Me.Left + 100
                            .Top = Me.Top + 100
                            .mstrItemCode = strItemCode
                            .mstrItemDesc = stritemdesc
                            .mstrMeasureCode = strMeasureCode
                            .mintDocType = 10
                            .mlngDocNo = Me.ctlitem_c1.Text
                            .mlnglotqty = lngLotQty
                            .mstrEmpCode = txtEmpQA.Text
                            .mstrvendCode = txtvencd.Text
                            .txtLotNo.Text = ctlitem_c1.Text
                            checkItemConditions = .CmdItemCodeHelp_Click
                            If checkItemConditions = True Then
                                .cmdEmpCodeHelp_Click()
                                .ShowDialog()
                            End If
                        End With
                    End With
                    '101161852
                Case Col_IQA
                    With frmIQAParameterDetails
                        .Left = Me.Left + 100
                        .Top = Me.Top + 100

                        With Me.sprdata
                            .Col = Col_Item_Code : .Row = .ActiveRow : strItemCode = .Text.Trim
                            .Col = Col_Actual_Qty : .Row = .ActiveRow : strActualQty = .Text.Trim
                        End With

                        .mstrItemCodeParameter = strItemCode
                        .mstrActualQtyParameter = strActualQty
                        .mstrGrinNoParameter = strDoc
                        .mstrGrinLocParameter = txtrecloc.Text.Trim

                        .ShowDialog()

                        With sprdata
                            Dim strMainSql As String
                            Dim strSqlSub As String
                            Dim varActualQty As Object
                            varActualQty = Nothing
                            Call .GetText(Col_Actual_Qty, .ActiveRow, varActualQty)
                            strMainSql = "SELECT TOP 1 1 FROM IQA_Parameter_Dtl  WHERE UNIT_CODE='" + gstrUnitId + "' AND  DOC_NO='" & strDoc & "' and item_code='" & strItemCode & "' and IQA_Status= '1' "
                            If IsRecordExists(strMainSql) Then
                                strSqlSub = "SELECT TOP 1 1 FROM IQA_Sample_Dtl  WHERE UNIT_CODE='" + gstrUnitId + "' AND  DOC_NO='" & strDoc & "' and item_code='" & strItemCode & "' and upper(Status)='NOT OK' "
                                If IsRecordExists(strSqlSub) = True Then
                                    .Row = .ActiveRow
                                    .Col = Col_Accepted_Qty
                                    .Text = 0
                                    .Col = Col_Rejected_Qty
                                    .Text = varActualQty
                                Else
                                    .Row = .ActiveRow
                                    .Col = Col_Accepted_Qty
                                    .Text = varActualQty
                                    .Col = Col_Rejected_Qty
                                    .Text = 0
                                End If

                            End If
                        End With
                    End With
                Case Col_ViewDoc
                    ViewDoc(e.row)
                Case Col_ESI
                    '*10857384
                    Dim strItem As String
                    Dim ESIDetail As cls_GRIN_ESI_Detail

                    With sprdata
                        .Row = e.row
                        .Col = Col_Item_Code
                        strItem = .Text.Trim.ToUpper
                    End With

                    Dim ESI = ListESIDetail.Where(Function(X) X.ItemCode.ToUpper = strItem.ToUpper)
                    For Each objESI As cls_GRIN_ESI_Detail In ESI
                        ESIDetail = objESI
                        Exit For
                    Next

                    If ESIDetail Is Nothing Then
                        ESIDetail = New cls_GRIN_ESI_Detail
                        ESIDetail.PO_No = txtpo.Text
                        ESIDetail.PO_Type = txtgrtype.Text
                        With sprdata
                            .Row = e.row

                            .Col = Col_Item_Code
                            ESIDetail.ItemCode = .Text.Trim

                            .Col = Col_Item_Description
                            ESIDetail.ItemDescription = .Text.Trim

                            .Col = Col_Accepted_Qty
                            ESIDetail.AcceptedQty = Val(.Text)

                            .Col = Col_Actual_Qty
                            ESIDetail.ReceiptQty = Val(.Text)

                            .Col = Col_Rejected_Qty
                            ESIDetail.RejectedQty = Val(.Text)

                            .Col = Col_Item_Rate
                            ESIDetail.BasicValue = Math.Round(ESIDetail.AcceptedQty * Val(.Text), 4)

                        End With
                    End If

                    Dim frmESIDetail As New FrmQATRN0001A
                    With frmESIDetail
                        .ShowFields(ESIDetail)
                        .ShowDialog()
                        If .DialogResult = System.Windows.Forms.DialogResult.OK Then
                            ListESIDetail.RemoveAll(Function(x) x.ItemCode.ToUpper = strItem.ToUpper)
                            ListESIDetail.Add(.GetFieldsValue)
                        End If
                    End With
                    frmESIDetail = Nothing
                    '*
            End Select
            Exit Sub
        Catch ex As Exception
            gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End Try
    End Sub

    Public Sub MakeFileFromBytes(ByVal strfilePath As String, ByVal FileBytes As Byte())

        Try
            If strfilePath.Trim.Length = 0 Then Exit Sub
            Dim oFs As New FileStream(strfilePath, FileMode.Create)
            Dim oBinaryWriter As New BinaryWriter(oFs)

            oBinaryWriter.Write(FileBytes)
            oBinaryWriter.Flush()
            oBinaryWriter.Close()
            oBinaryWriter = Nothing
            oFs.Close()
            oFs = Nothing
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Sub

    Private Sub ViewDoc(ByVal intActiveRow As Integer)

        Dim objDR As SqlDataReader = Nothing
        Dim strSql As String = ""
        Dim strItemCode As String = String.Empty

        With sprdata
            .Row = intActiveRow
            .Col = Col_Item_Code
            strItemCode = .Text.Trim
            strSql = "select isnull(Doc_Filename,'')Doc_Filename  ,isnull(Doc_Image,'')Doc_Image  from ItemDocs where unit_code='" & gstrUNITID & "' and item_code='" & strItemCode & "'"
        End With
        Try
            objDR = SqlConnectionclass.ExecuteReader(strSql)
            If objDR.HasRows Then
                While objDR.Read
                    Dim strfilename As String = Path.GetTempPath() & objDR("Doc_Filename").ToString().Trim
                    Dim FileBytes As Byte() = objDR("Doc_Image")
                    MakeFileFromBytes(strfilename, FileBytes)
                    If objDR("DOC_FILENAME").ToString().Trim.Length > 0 Then
                        Try
                            With FrmPDFViewer
                                If InStr(UCase(objDR("DOC_FILENAME").ToString().Trim), ".PDF") > 0 Then
                                    .AxAcroPDF1.LoadFile(strfilename)
                                    .AxAcroPDF1.setZoom(60)
                                    .GrpPDF.Visible = True
                                    .grpImage.Visible = False
                                    .GrpPDF.Dock = DockStyle.Fill
                                Else
                                    .PictureBox1.ImageLocation = Path.GetTempPath() & objDR("Doc_Filename").ToString().Trim
                                    .GrpPDF.Visible = False
                                    .grpImage.Visible = True
                                    .grpImage.Show()
                                    .grpImage.Dock = DockStyle.Fill
                                    .PictureBox1.Dock = DockStyle.Fill
                                    .PictureBox1.Show()
                                End If
                                .ShowDialog()
                            End With
                        Catch ex As Exception
                            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
                        End Try
                    End If
                End While

            Else
                MessageBox.Show("No Doc/Image found against this item.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If


            If objDR.IsClosed = False Then objDR.Close()
            objDR = Nothing
        Catch ex As Exception
            If objDR.IsClosed = False Then objDR.Close()
            objDR = Nothing
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Sub

    Private Sub sprdata_EditChange(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_EditChangeEvent) Handles sprdata.EditChange
        'On Error GoTo ErrHandler
        Dim dblActualQty As Double
        Dim strMeasureCode As String
        Dim dblAcceptedQty As Double
        Dim dblReceiptQty As Double
        '*10857384
        Dim ESIDetail As cls_GRIN_ESI_Detail
        Dim strItemCode As String = String.Empty
        '*
        Try
            If e.col = Col_Accepted_Qty Then
                With Me.sprdata
                    .Row = e.row
                    .Col = Col_Item_UOM : strMeasureCode = .Text.Trim
                    .Col = Col_Actual_Qty : dblActualQty = Val(.Text)
                    .Col = Col_Accepted_Qty : dblAcceptedQty = Val(.Text)
                    .Col = Col_Receipt_Qty
                    If .ColHidden = True Then
                        dblReceiptQty = dblActualQty
                    Else
                        dblReceiptQty = Val(.Text)
                    End If
                    If GetUOMDecimalPlacesAllowed(strMeasureCode) > 0 Then
                        .Col = Col_Rejected_Qty : .Text = Format(dblReceiptQty - dblAcceptedQty, "0.0000")
                    Else
                        .Col = Col_Rejected_Qty : .Text = Format(dblReceiptQty - dblAcceptedQty, "0")
                    End If

                    '*10857384
                    .Col = Col_Item_Code
                    strItemCode = .Text.Trim.ToUpper
                    '*
                End With

                '*10857384
                Dim ESI = ListESIDetail.Where(Function(X) X.ItemCode.ToUpper = strItemCode.ToUpper)
                For Each objESI As cls_GRIN_ESI_Detail In ESI
                    If Math.Round(objESI.AcceptedQty, 4) <> Math.Round(dblAcceptedQty, 4) Then
                        ListESIDetail.RemoveAll(Function(x) x.ItemCode.ToUpper = strItemCode.ToUpper)
                    End If
                    Exit For
                Next
                '*
            End If
            Exit Sub
        Catch ex As Exception
            gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End Try
    End Sub
    Private Sub sprdata_KeyDownEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyDownEvent) Handles sprdata.KeyDownEvent
        On Error GoTo ErrHandler
        Dim dblRejectedQty As Double
        Dim intcheck As Integer
        If e.keyCode = Keys.Return Then
            With sprdata
                .Row = .ActiveRow : .Col = Col_Inspection_Entry : intcheck = Val(.Text)
                .Row = .ActiveRow : .Col = Col_Rejected_Qty : dblRejectedQty = Val(.Text)
                If mbln_Inspection_Control_Details = True And mblnBatchTracking = False And .ActiveRow = .MaxRows And .ActiveCol = Col_Inspection_Entry And dblRejectedQty = 0 And intcheck = 0 Then
                    cmdGrpAuthorise1.Focus()
                ElseIf mbln_Inspection_Control_Details = True And mblnBatchTracking = False And .ActiveRow = .MaxRows And .ActiveCol = col_Inspection_Control_Details Then
                    cmdGrpAuthorise1.Focus()
                End If
                If mbln_Inspection_Control_Details = False And mblnBatchTracking = True And dblRejectedQty = 0 And .ActiveCol = Col_Batch_Details And .ActiveRow = .MaxRows Then
                    cmdGrpAuthorise1.Focus()
                ElseIf mbln_Inspection_Control_Details = False And mblnBatchTracking = True And dblRejectedQty <> 0 And .ActiveCol = Col_Batch_Details And .ActiveRow = .MaxRows Then
                    cmdGrpAuthorise1.Focus()
                End If
                If mbln_Inspection_Control_Details = True And mblnBatchTracking = True And .ActiveRow = .MaxRows And .ActiveCol = Col_Batch_Details And intcheck = 0 Then
                    cmdGrpAuthorise1.Focus()
                ElseIf mbln_Inspection_Control_Details = True And mblnBatchTracking = True And .ActiveRow = .MaxRows And .ActiveCol = col_Inspection_Control_Details Then
                    cmdGrpAuthorise1.Focus()
                End If
            End With
        End If
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub sprdata_KeyPressEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyPressEvent) Handles sprdata.KeyPressEvent
        On Error GoTo ErrHandler
        With sprdata
            If e.keyAscii = 46 Then
                .Col = .ActiveCol : .Row = .ActiveRow
                If .CellType = FPSpreadADO.CellTypeConstants.CellTypeInteger Then
                    Call ConfirmWindow(10455)
                    .Focus() : .Row = .ActiveRow : .Col = .ActiveCol : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus() : Exit Sub
                End If
            End If
        End With
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub sprdata_LeaveCell(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_LeaveCellEvent) Handles sprdata.LeaveCell
        '***************************************************************************
        'Revision Date  - 09-Oct-2006
        'Revised By     - Davinder Singh
        'Issue No       - 17893
        'Description    - To fill the rejection details in the array directly in case if grin is against single PO
        '                 and take the input of rejection details by showing PO wise rejection frame in case of Grin against Multiple PO's
        '                 on leaving the Accepted Qty cell of the grid
        '**************************************************************************************************************************************
        'Revised By        -   Davinder Singh
        'Revision Date     -   17 Sep 2007
        'Issue Id          -   21102
        'Revision History  -   If Barcode tracking is on for the Item then check if Binning of this
        '                      Item has been done or not. If Binning has been done then Qty.
        '                      associated with active packets should be equal to the accepted Qty. of GRIN
        '                      Otherwise user may have to do Binning or Cancel Packets
        '**************************************************************************************************************************************
        On Error GoTo ErrHandler
        Dim dblAcceptedQuantity As Double
        Dim dblActualQty As Double
        Dim dblRejectedQty As Double
        Dim dblReceiptQty As Double
        Dim strMeasureCode As String
        Dim Rs As ADODB.Recordset
        Dim strMissingPkt As String = ""
        Dim strItemDesc As String
        If e.newCol = -1 Then Exit Sub
        If e.row <> 0 Then
            With sprdata
                If e.col = Col_Accepted_Qty Then
                    .Row = e.row
                    .Col = Col_Actual_Qty
                    dblActualQty = Val(.Text)
                    .Col = Col_Receipt_Qty
                    If .ColHidden = True Then
                        dblReceiptQty = dblActualQty
                    Else
                        dblReceiptQty = Val(.Text)
                    End If
                    .Col = Col_Item_Description : strItemDesc = .Text.Trim
                    .Col = Col_Rejected_Qty : dblRejectedQty = Val(.Text)
                    .Col = Col_Accepted_Qty : dblAcceptedQuantity = Val(.Text)

                    If dblAcceptedQuantity > dblActualQty Then
                        sprdata.Enabled = True
                        MsgBox("Accepted Qty Can't Exceed Actual Quantity  " & dblActualQty, vbInformation, ResolveResString(100))
                        .Col = Col_Rejected_Qty : .Text = dblReceiptQty
                        .Col = Col_Accepted_Qty : .Text = Format(0, "0.0000") : .EditModePermanent = True : .EditModeReplace = True : .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                        Exit Sub
                    Else
                        .Col = Col_Accepted_Qty : dblAcceptedQuantity = Val(.Text)
                        .Col = Col_Rejected_Qty : .Text = dblReceiptQty - dblAcceptedQuantity
                        With sprdata
                            .Col = Col_Item_UOM : strMeasureCode = .Text.Trim
                            .Col = Col_Actual_Qty : dblActualQty = Val(.Text)
                            .Col = Col_Accepted_Qty : dblAcceptedQuantity = Val(.Text)
                            If GetUOMDecimalPlacesAllowed(strMeasureCode) > 0 Then
                                .Col = Col_Rejected_Qty : .Text = Format(dblReceiptQty - dblAcceptedQuantity, "0.0000")
                            Else
                                .Col = Col_Rejected_Qty : .Text = Format(dblReceiptQty - dblAcceptedQuantity, "0")
                            End If
                            .Row = e.row
                            .Col = Col_Item_Code
                            mstrItemCode = .Text.Trim
                            .Col = Col_Item_UOM
                            mstrUOM = .Text.Trim
                            .Col = Col_Rejected_Qty
                            dblRejectedQty = Val(.Text)
                            .Col = Col_Item_Description
                            mstrDescription = .Text.Trim
                            .Col = Col_Rejection_Reason
                            mstrRejReason = .Text.Trim
                            If mstrGRType.Value = "U" Or mstrGRType.Value = "E" Then  ''---- Customer Rejection type GRIN.
                                .Col = Col_Rejection_Details : .Lock = False
                            Else
                                If dblAcceptedQuantity = dblReceiptQty Then
                                    .Col = Col_Rejection_Details : .Lock = True
                                    If mstrPOstatus <> "W" Then
                                        Call AutoFillPOwiseRejectionDtls(mstrItemCode, mstrDescription, mstrUOM, 0)
                                    End If
                                Else
                                    .Col = Col_Rejection_Details : .Lock = False
                                    If mstrPOstatus = "A" Then
                                        Call AutoFillPOwiseRejectionDtls(mstrItemCode, mstrDescription, mstrUOM, dblRejectedQty)
                                    ElseIf mstrPOstatus = "M" Then
                                        Call ShowPOWiseRejectionDtls(mstrItemCode, mstrDescription, mstrUOM, dblRejectedQty)
                                    End If
                                End If
                            End If

                            If mblnVendInvDateOlderThanPODate AndAlso txtReqApprovalStatus.Text.Trim.ToUpper = "REJECTED" Then
                                If dblAcceptedQuantity > 0 Then
                                    MsgBox("GRIN should be fully rejected as Vendor Invoice Date is less than PO Date and same has been rejected for acceptance from approval workflow ! Kindly rejected full Qty.", MsgBoxStyle.Exclamation, ResolveResString(100))
                                    .Row = e.row
                                    .Col = Col_Accepted_Qty
                                    .Text = Format(0, "0.0000")
                                    .EditModePermanent = True
                                    .EditModeReplace = True
                                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                    e.cancel = True
                                    Exit Sub
                                End If
                            End If

                            If mblnBarcodeGRN = True And Barcode_Location(txtrecloc.Text, txtdesloc.Text) = True Then ''--- If Barcode Process is Enabled in GRIN
                                If (mstrGRType.Value <> "E" And mstrGRType.Value <> "U") Then
                                    If GetBarcodeTracking(mstrItemCode) = True Then
                                        Rs = mP_Connection.Execute("SELECT DBO.UFN_GETBINNEDQTY('" & gstrUNITID & "'," & ctlitem_c1.Text & ",'" & mstrItemCode & "','" & txtdesloc.Text & "') AS BINNEDQTY")
                                        If Not (Rs.BOF And Rs.EOF) Then
                                            If Math.Round(Rs.Fields("BINNEDQTY").Value, 2) > Math.Round(dblAcceptedQuantity, 2) Then
                                                MsgBox("Bags Available Qty.[" & Rs.Fields("BINNEDQTY").Value & "] In Store Is Greater Than Accepted Qty.[" & dblAcceptedQuantity & "] GRIN Can't be Authorised " & vbCrLf & _
                                                       "Please Cancel Remaining Bags. for Item Code:" & mstrItemCode, vbInformation, ResolveResString(100))
                                                ShowBarcodeLabelScreen(CInt(ctlitem_c1.Text), mstrItemCode, strItemDesc)
                                                .Row = e.row
                                                .Col = Col_Accepted_Qty
                                                .Text = Format(0, "0.0000")
                                                .EditModePermanent = True
                                                .EditModeReplace = True
                                                .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                                e.cancel = True
                                                Rs.Close()
                                                Rs = Nothing
                                                Exit Sub
                                            End If
                                            If Math.Round(Rs.Fields("BINNEDQTY").Value, 2) < Math.Round(dblAcceptedQuantity, 2) Then
                                                strMissingPkt = GetMissingPakcets(ctlitem_c1.Text, mstrItemCode)
                                                strMissingPkt = "Bags Available Qty.[" & Rs.Fields("BINNEDQTY").Value & "] In Store Is Less Than Accepted Qty.[" & dblAcceptedQuantity & "] GRIN Can't be Authorised " & vbCrLf & _
                                                       "Please Do the Binning For Remaining Qty. of Item Code: " & mstrItemCode & vbCrLf & strMissingPkt
                                                'MsgBox("Bags Available Qty.[" & Rs.Fields("BINNEDQTY").Value & "] In Store Is Less Than Accepted Qty.[" & dblAcceptedQuantity & "] GRIN Can't be Authorised " & vbCrLf & _
                                                '"Please Do the Binning For Remaining Qty. of Item Code: " & mstrItemCode, vbInformation, ResolveResString(100))
                                                MsgBox(strMissingPkt, MsgBoxStyle.Information, ResolveResString(100))
                                                ShowBarcodeLabelScreen(CInt(ctlitem_c1.Text), mstrItemCode, strItemDesc)
                                                .Row = e.row
                                                .Col = Col_Accepted_Qty
                                                .Text = Format(0, "0.0000")
                                                .EditModePermanent = True
                                                .EditModeReplace = True
                                                .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                                e.cancel = True
                                                Rs.Close()
                                                Rs = Nothing
                                                Exit Sub
                                            End If
                                        End If
                                        Rs.Close()
                                        Rs = Nothing
                                    End If
                                End If
                            End If
                        End With
                    End If
                ElseIf e.col = Col_Rejection_Reason Then
                    .Row = e.row
                    .Col = Col_Rejected_Qty : dblRejectedQty = Val(.Text)
                    If mbln_Inspection_Control_Details = False And mblnBatchTracking = False And Me.sprdata.Row = Me.sprdata.MaxRows And dblRejectedQty = 0 Then
                        cmdGrpAuthorise1.Focus() : Exit Sub
                    End If
                ElseIf e.col = Col_Rejection_Details Then
                    If fraDefectDetails.Visible = True Then
                        vaDefects.Focus()
                    Else
                        .Row = e.row + 1 : .Col = Col_Accepted_Qty : .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                    End If
                End If
            End With
        End If
        Exit Sub
ErrHandler:
        Rs = Nothing
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub vaDefects_EditChange(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_EditChangeEvent) Handles vaDefects.EditChange
        On Error GoTo ErrHandler
        Dim dblTotalRejectedQty As Double
        Dim dblDefectRejection As Double
        Dim intCounter As Integer
        Dim dblTillRejectedQty As Double
        Select Case e.col
            Case Col_Defect_qty
                With vaDefects
                    .Row = e.row : .Col = Col_Defect_qty : dblDefectRejection = Val(.Text)
                    For intCounter = 1 To .MaxRows
                        .Row = intCounter : .Col = Col_Defect_qty : dblDefectRejection = Val(.Text)
                        dblTillRejectedQty = dblTillRejectedQty + dblDefectRejection
                    Next
                    lblDefRejection.Text = dblTillRejectedQty
                    dblTotalRejectedQty = Me.lblTotalRej.Text
                    Me.lblBalRejQty.Text = dblTotalRejectedQty - dblTillRejectedQty
                    If GetUOMDecimalPlacesAllowed(mstrMeasureCode) > 0 Then
                        Me.lblTotalRej.Text = Format(CDbl(Me.lblTotalRej.Text), "0.0000")
                        Me.lblDefRejection.Text = Format(CDbl(Me.lblDefRejection.Text), "0.0000")
                        Me.lblBalRejQty.Text = Format(CDbl(Me.lblBalRejQty.Text), "0.0000")
                    Else
                        Me.lblTotalRej.Text = Format(CDbl(Me.lblTotalRej.Text), "0")
                        Me.lblDefRejection.Text = Format(CDbl(Me.lblDefRejection.Text), "0")
                        Me.lblBalRejQty.Text = Format(CDbl(Me.lblBalRejQty.Text), "0")
                    End If
                End With
        End Select
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub dtpQADate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles dtpQADate.KeyPress
        With sprdata
            .Row = 1
            .Col = Col_Accepted_Qty
            .Action = FPSpreadADO.ActionConstants.ActionActiveCell
            .Focus()
        End With
    End Sub
    Private Function DiffInRates(ByVal pItemCode As String) As Boolean
        Dim rs As ClsResultSetDB
        Dim StrQry As String
        Try
            DiffInRates = False
            StrQry = " SELECT TOP 1 1 FROM GRN_HDR A INNER JOIN GRN_PO_DTL B " & _
                    " ON A.UNIT_CODE = B.UNIT_CODE AND A.DOC_NO=B.DOC_NO AND A.DOC_TYPE=10" & _
                    " WHERE A.UNIT_CODE = '" & gstrUNITID & "' AND A.DOC_CATEGORY='Z' AND B.Rate_Flag=1 AND B.ITEM_CODE='" & pItemCode.Trim & "' AND  A.DOC_NO=" & Me.ctlitem_c1.Text.Trim & ""
            rs = New ClsResultSetDB
            rs.GetResult(StrQry)
            If Not rs.EOFRecord Then
                DiffInRates = True
            End If
            rs.ResultSetClose()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
        End Try
    End Function
    Private Function GetMissingPakcets(ByVal GrinNo As Integer, ByVal ItemCode As String) As String
        Dim strQry As String
        Dim strMsg As String = ""
        Dim rs As New ClsResultSetDB
        On Error GoTo errHandler
        strQry = "SELECT F.FIFO_LOTNO,F.FIFO_QTY,I.ITEM_CODE"
        strQry += " FROM ITEM_MST I INNER JOIN BAR_FIFOLABEL F"
        strQry += " ON (I.UNIT_CODE = F.UNIT_CODE AND I.ITM_ITEMALIAS = F.FIFO_PARTNO	)"
        strQry += " LEFT OUTER JOIN BAR_CROSSREFERENCE C"
        strQry += " ON (C.UNIT_CODE = F.UNIT_CODE AND C.CREF_PACKETNO = F.FIFO_LOTNO AND C.CANCEL_FLAG=0 )"
        strQry += " WHERE I.UNIT_CODE = '" & gstrUNITID & "' AND I.ITEM_CODE ='" & ItemCode & "' AND F.FIFO_PDOCID = '" & GrinNo & "' "
        strQry += " AND F.FIFO_DOCTYPE = 10 AND C.CREF_PACKETNO IS NULL "
        rs.GetResult(strQry)
        If Not rs.EOFRecord Then
            strMsg = "Missing Packets : " & vbCrLf & "------------------" & vbCrLf
            rs.MoveFirst()
            While Not rs.EOFRecord
                strMsg += rs.GetValue("FIFO_LOTNO") & ", "
                rs.MoveNext()
            End While
            strMsg = strMsg.Trim.Substring(0, strMsg.Length - 2)
        End If
        rs.ResultSetClose()
        rs = Nothing
        Return strMsg
        Exit Function
errHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function KNOCKOFF_DTF_TRANSACTION() As Boolean
        Dim oCmd As New ADODB.Command
        Dim strQry As String = ""
        On Error GoTo errHandler
        KNOCKOFF_DTF_TRANSACTION = False
        strQry = "DELETE FROM TMPPENDINGDTF WHERE UNIT_CODE = '" & gstrUNITID & "' AND IP_ADDR='" & gstrIpaddressWinSck & "'"
        mP_Connection.Execute(strQry, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        strQry = "INSERT INTO TMPPENDINGDTF(UNIT_CODE, ITEM_CODE,IP_ADDR) "
        strQry += " SELECT A.UNIT_CODE, A.ITEM_CODE,'" & gstrIpaddressWinSck & "' FROM GRN_DTL A INNER JOIN ITEM_MST B" & _
            " ON A.UNIT_CODE = B.UNIT_CODE AND A.ITEM_CODE=B.ITEM_CODE"
        strQry += " WHERE A.UNIT_CODE = '" & gstrUNITID & "' AND" & _
            " A.DOC_NO='" & Val(Me.ctlitem_c1.Text) & "' AND A.DOC_TYPE=10 AND BARCODE_TRACKING=1"
        strQry += " GROUP BY A.UNIT_CODE, A.ITEM_CODE "
        mP_Connection.Execute(strQry, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        With oCmd
            .let_ActiveConnection(mP_Connection)
            .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
            .CommandText = "USP_BAR_KNOCKOFF_DTF"
            .CommandTimeout = 0
            .Parameters.Append(.CreateParameter("@UNITCODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
            .Parameters.Append(.CreateParameter("@IP_ADDR", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, gstrIpaddressWinSck))
            .Parameters.Append(.CreateParameter("@USER_ID", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 16, Trim(mP_User)))
            .Parameters.Append(.CreateParameter("@ERR_CODE", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamOutput))
            .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        End With
        If oCmd.Parameters(oCmd.Parameters.Count - 1).Value <> 0 Then
            oCmd = Nothing
            KNOCKOFF_DTF_TRANSACTION = False
            Err.Raise(123, , "Error while Knocking off DTF Transaction ")
            Exit Function
        End If
        KNOCKOFF_DTF_TRANSACTION = True
        oCmd = Nothing
        Exit Function
errHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function DTF_ISSUE_ALLOWED() As Boolean
        Dim rs As New ClsResultSetDB
        Dim strQry As String
        Try
            DTF_ISSUE_ALLOWED = False
            strQry = "Select top 1 1 from barcode_config_mst where DTF_ISSUE=1 AND UNIT_CODE = '" & gstrUNITID & "'"
            rs.GetResult(strQry)
            If Not rs.EOFRecord Then
                DTF_ISSUE_ALLOWED = True
            End If
            rs.ResultSetClose()
            rs = Nothing
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Private Sub ViewGrinAuthReport(ByVal intGRNNo As Integer)
        'CREATED BY     :   VINOD
        'CREATED DATE   :   08/03/2011
        On Error GoTo errHandler
        Dim strQry As String
        Dim rs As New ClsResultSetDB
        strQry = "SELECT TOP 1 1 FROM STORES_CONFIGMST WHERE UNIT_CODE = '" & gstrUNITID & "' AND VIEWGRINAUTHREPORT=1"
        rs.GetResult(strQry)
        If Not rs.EOFRecord Then
            If intGRNNo > 0 Then
                txtDocNo.Text = intGRNNo
                ShowGrinReport(1)
            End If
        End If
        rs.ResultSetClose()
        rs = Nothing
        Exit Sub
errHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub ShowBarcodeLabelScreen(ByVal GRIN_No As Integer, ByVal Item_code As String, ByVal ItemDesc As String)
        'ISSUE ID       :   10504051
        'DESCRIOPTION   :   SHOWS BARCODE LABEL SCREEN TO EDIT/UPDATE LABEL QTY
        Try
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.WaitCursor)
            Dim frmLabel As New frmBARTRN0020A
            With frmLabel
                .DocumentNo = GRIN_No
                .ParentScreen = frmBARTRN0020A._EnmParent.QAAuth
                .ItemCode = Item_code
                .ItemDescription = ItemDesc
                .Mode = frmBARTRN0020A.EnmMode.ADD
                .ShowDialog()
                frmLabel = Nothing
            End With
        Catch ex As Exception
            RaiseException(ex)
        Finally
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        End Try

    End Sub
    Private Function VendorBarcode(ByVal strVendorCode As String) As Boolean
        'ISSUE ID   :   10504051
        'Checks is Vendor barcode functionality is enabled or not.
        Dim strQry As String = String.Empty
        Dim dt As DataTable

        Try
            'ISSUE ID 10639761  CONDITION ADDED FOR GRIN TYPE : M
            If Me.txtgrtype.Text <> "S" And Me.txtgrtype.Text <> "M" Then
                strQry = "SELECT TOP 1 1 FROM BARCODE_CONFIG_MST WHERE VENDOR_BARCODE=1 AND VendorBarcodeAtGRIN=1 AND UNIT_CODE='" & gstrUNITID & "'"
                dt = SqlConnectionclass.GetDataTable(strQry)
                If dt.Rows.Count = 0 Then Return False

                strQry = "SELECT TOP 1 1 FROM  VENDOR_MST WHERE VENDOR_CODE='" & Trim(strVendorCode.Trim) & "' and VENDOR_BARCODE=1 AND UNIT_CODE='" & gstrUNITID & "'"
                dt = SqlConnectionclass.GetDataTable(strQry)
                If dt.Rows.Count = 0 Then
                    Return False
                Else
                    Return True
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Private Sub SaveESIDetails()
        Dim strItemCode As String
        Dim strSQL As String
        Try
            If mstrGRType.ToString.ToUpper = "W" Then
                With sprdata
                    For intRow As Integer = 1 To .MaxRows
                        .Row = intRow
                        .Col = Col_Item_Code
                        strItemCode = .Text.Trim.ToUpper
                        Dim objESIDtl = ListESIDetail.Where(Function(x) x.ItemCode.ToUpper = strItemCode.ToUpper)
                        For Each esi As cls_GRIN_ESI_Detail In objESIDtl
                            strsql = "INSERT INTO GRN_ESI_DTL(UNIT_CODE,GRN_NO,PO_NO,PO_TYPE,ITEM_CODE,RECEIPT_QTY,ACCEPTED_QTY,REJECTED_QTY, "
                            strSQL += " SERVICE_PERC,SERVICE_VALUE,ESI_PERC,ESI_VALUE,ENT_DT,ENT_USERID,REMARKS) "
                            strSQL += " VALUES ('" & gstrUNITID & "'," & ctlitem_c1.Text & "," & txtpo.Text & ",'" & mstrGRType.ToString.ToUpper & "','" & esi.ItemCode & "'," & esi.ReceiptQty & "," & esi.AcceptedQty & "," & esi.RejectedQty & ", "
                            strSQL += " " & esi.ServiceRateInPercentage & "," & esi.ServiceValue & "," & esi.ESIRateInPercentage & "," & esi.ESIValue & ",GETDATE(),'" & mP_User & "','" & esi.Remarks & "')"
                            mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        Next
                    Next
                End With
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    '101161852
    Private Sub CheckIQARequired()
        Try
            mblnIQARequired = Convert.ToBoolean(SqlConnectionclass.ExecuteScalar("SELECT ISNULL(IQA_REQUIRED,0) IQA_REQUIRED FROM STORES_CONFIGMST (NOLOCK) WHERE UNIT_CODE='" & gstrUnitId & "'"))
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub btnIQCReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnIQCReport.Click
        Try
            Dim FRM As New frmQATRN0001B()
            FRM.ShowDialog()
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString())
        End Try
    End Sub

    Private Sub CheckFlags()
        'VID CHANGES
        Dim strSQL As String
        Try
            strSQL = "SELECT ISNULL(VendInvDtLessThanPODt_Easy_Approval,0) VendInvDtLessThanPODt_Easy_Approval FROM Stores_ConfigMst WHERE UNIT_CODE ='" & gstrUNITID & "' "
            Using dt As DataTable = SqlConnectionclass.GetDataTable(strSQL)
                If dt.Rows.Count > 0 Then
                    mblnVendInvDateOlderThanPODate = Convert.ToBoolean(dt.Rows(0)("VendInvDtLessThanPODt_Easy_Approval"))
                End If
            End Using
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub GetVendorInvoiceDateOlderThanPODateApprovalDetail()
        'VALIDATE INVOICE DATE WITH PO DATE
        Dim strSQL As String
        Try
            txtEasyReqNo.Text = ""
            txtReqApprovalStatus.Text = ""
            If mblnVendInvDateOlderThanPODate = False Then Return
            strSQL = "SELECT EASY_NO,APPROVAL_STATUS FROM VEND_INV_DT_EXCEPTION_EASY_APPROVAL WHERE UNIT_CODE ='" & gstrUNITID & "' AND GE_NO='" & txtInwRegNo.Text.Trim & "'"
            Using dt As DataTable = SqlConnectionclass.GetDataTable(strSQL)
                If dt.Rows.Count > 0 Then
                    txtEasyReqNo.Text = Convert.ToString(dt.Rows(0)("EASY_NO"))
                    txtReqApprovalStatus.Text = Convert.ToString(dt.Rows(0)("APPROVAL_STATUS"))
                Else
                    txtEasyReqNo.Text = ""
                    txtReqApprovalStatus.Text = ""
                End If
            End Using
        Catch ex As Exception
            Throw ex
        End Try


    End Sub
    Private Sub AutoMTNforJobworkGRN()
        Dim strSQL As String = String.Empty
        Dim blnBarcodeItemSkipped As Boolean
        Dim intNewMTNNo As Integer
        Try
            strSQL = "SELECT TOP 1 1 FROM STORES_CONFIGMST WHERE UNIT_CODE ='" & gstrUNITID & "' AND Auto_MTN_For_GRIN=1"
            If IsRecordExists(strSQL) = False Then Exit Sub
            Using sqlCmd As SqlCommand = New SqlCommand
                With sqlCmd
                    .CommandText = "USP_AUTO_MTN_FOR_GRIN"
                    .CommandType = CommandType.StoredProcedure
                    .CommandTimeout = 0
                    .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                    .Parameters.Add("@GRIN_NO", SqlDbType.Int).Value = ctlitem_c1.Text
                    .Parameters.Add("@USER_ID", SqlDbType.VarChar, 20).Value = mP_User
                    .Parameters.Add("@BARCODE_ITEMS_SKIPPED", SqlDbType.Bit).Direction = ParameterDirection.Output
                    .Parameters.Add("@NEW_MTN_NO", SqlDbType.Int).Direction = ParameterDirection.Output
                    SqlConnectionclass.ExecuteNonQuery(sqlCmd)
                    blnBarcodeItemSkipped = Convert.ToBoolean(.Parameters("@BARCODE_ITEMS_SKIPPED").Value)
                    intNewMTNNo = Convert.ToInt32(.Parameters("@NEW_MTN_NO").Value)
                    If intNewMTNNo > 0 And blnBarcodeItemSkipped = True Then
                        MsgBox("Auto MTN[" & intNewMTNNo & "] successfully done except for barcode items, Manual MTN transaction is required for barcode items !", MsgBoxStyle.Exclamation, ResolveResString(100))
                    ElseIf intNewMTNNo > 0 And blnBarcodeItemSkipped = False Then
                        MsgBox("Auto MTN[" & intNewMTNNo & "] successfully done.", MsgBoxStyle.Information, ResolveResString(100))
                    End If
                End With
            End Using

        Catch ex As Exception
            RaiseException(ex)
        End Try

    End Sub
    
End Class