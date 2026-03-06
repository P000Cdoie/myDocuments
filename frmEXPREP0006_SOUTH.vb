Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports System.IO
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Drawing.Imaging
Imports System.Runtime.InteropServices
Imports System.Drawing.Drawing2D
Imports System.Drawing

Friend Class frmEXPREP0006_SOUTH
    Inherits System.Windows.Forms.Form
    '===================================================================================
    ' (c) MothersonSumi INfotech & Design Ltd. All rights reserverd.
    ' File Name         :   frmEXPREP0006.frm
    ' Function          :   Used to Print & View Export Invoice deatails
    ' Created By        :   Nimesh Verma
    ' Created On        :   09 May, 2001
    ' Revision History  :   05 Dec, 2001
    '19 feb, 2002, nimesh, changed - currency code display
    '15/01/2002 CHANGED FOR DOCUMENT NO. ON FORM NO. 4069
    '03/22/2002 IN CREASED THE SIZE AS REQUIRED BY MSSED
    '29/05/02 To remove Stock Check in Case of Reprint Of invoice.
    '29/06/2002  added msgbox to show what items have current stock less than the sales qty
    '29/06/2002  To remove stock check during preview
    '20/03/2003 for Financial Year RollOver By nisha       28/03/2003
    'changes done by nisha on 26/07/2003 for Catia invoice format.
    'chnaged by nisha 20/01/2004 To Remove Greater ">" Sign From Select Query
    '--------------------------------------------------------
    'Changed by Arshad on 22/09/2004
    'option for Print Excise Format is included
    'Fresh credit is received against a selected invoice
    '--------------------------------------------------------
    '--------------------------------------------------------
    'Changed by Nisha on 30/09/2004
    'To update on the basis of Single Series
    '--------------------------------------------------------
    '--------------------------------------------------------
    'Changed by Nisha on 04/10/2004
    'To update companyMst deatils in Report
    '--------------------------------------------------------
    'To stop the posting of invoice in case of sample for chennei only by nisha on 21 Nov 2005
    '===================================================================================
    'Revised By     : Arul Mozhi
    'Revised On     : 25-01-2005
    '                 1) Getting Type Mismatch Error in generate/invoiceNo Function
    '                 2) Geenrate Excise Format & Grid Make as Invisible Mode at Design time
    '                    This is only used in 100% EOU units only as Per Mr Nirat sir confirmation
    '===================================================================================
    'Revised By         : Davinder Singh
    'Revised on         : 12-Feb-2007
    'Revision History   : To Update Invoice No agst Despatch Advise
    '===================================================================================
    'Revised By         :   Manoj Kr. Vaish
    'Revised on         :   10-Sep-2007
    'Revision History   :   To Update Invoice No in bar_palette_mst agst Despatch Advise
    '===================================================================================
    'Revised By         :   Manoj Vaish
    'Revision Date      :   04 Mar 2009
    'Issue ID           :   eMpro-20090227-27987
    'Revision History   :   Changes for commercial invoice at Mate Units
    '-----------------------------------------------------------------------------
    'Revised By        -    Vinod Singh
    'Revision Date     -    07/06/2011
    'Revision History  -    Changes for Multi Unit
    '-----------------------------------------------------------------------------------------------------
    'Revised By         :   prashant Rajpal
    'Revision Date      :   14-dec-2011
    'Issue ID           :   10170550 
    'Revision History   :   Changes for ASN MULTIPLE PATH
    '-----------------------------------------------------------------------------
    'Revised By         :   Prashant Rajpal
    'Revision Date      :   23-dec-2011
    'Issue ID           :   10174688 
    'Revision History   :   Changes for error - Master insertion Failure error, rectified.
    '-----------------------------------------------------------------------------
    'Modified By Deepak kumar on 31 Jan 2012 to support multiunit change management
    '-----------------------------------------------------------------------------
    '****************************************************************************************
    'Revised By         :   Prashant Rajpal
    'Revision Date      :   27-Aug 2012
    'Issue ID           :   10266201  
    'Revision History   :   Incorporate the validation of Sales Order (Despatch quantity should not be greater than Schedule Qty )
    '**********************************************************************************************************************
    '**********************************************************************************************************************
    'Revised By         :   Prashant Rajpal
    'Revision Date      :   27-Aug 2012
    'Issue ID           :   10274457   
    'Revision History   :   For mtl sharjah , no need for validation of sales order schedule vs disaptch
    '**********************************************************************************************************************
    'Revised By         :   Prashant Rajpal
    'Revision Date      :   19 Oct 2012
    'issue id           :   10274457
    'Revision History   :   For Groupo ASN 
    '**********************************************************************************************************************
    '**********************************************************************************************************************
    'Revised By         :   Prashant Rajpal
    'Revision Date      :   21 Nov 2012
    'issue id           :   10310582
    'Revision History   :   For Groupo Customer ASN -change 
    '**********************************************************************************************************************
    'Revised By         :   Vinod Singh
    'Revision Date      :   28 March 2013
    'Revision History   :   Changes for Stock In Transit functionality
    '*********************************************************************************************************************
    'Revised By         :   Prashant Rajpal
    'Revision Date      :   18-Mar-2013-08 apr 2013
    'Issue ID           :   10354980   
    'Revision History   :   Woco migration changes
    '**********************************************************************************************************************
    'Revised By         :   PRASHANT RAJPAL
    'Issue ID           :   10597202    
    'Revision Date      :   15 May 2014
    'HISTORY            :   SINGLE  INVOICE SERIES  FOR UNIT 3 : TELECOM , DOMESTIC AND EXPORT 
    '***************************************************************************************************
    'Revised By         :   Shalini Singh
    'Issue ID           :   10646968    
    'Revision Date      :   05 Aug 2014
    'HISTORY            :   Item was coming in report 7 times and its value as well.
    '***************************************************************************************************
    'Revised By         :   PRASHANT RAJPAL
    'Issue ID           :   10853890  
    'Revision Date      :   01 Sep 2015
    'HISTORY            :   internmediate consignee code for unit 3 export
    '***************************************************************************************************
    'REVISED BY        -    PRASHANT RAJPAL
    'REVISION DATE     -    25/11/2015
    'REVISION HISTORY  -    Credit term picked wrong 
    'ISSUE ID          -    10895403    
    '**********************************************************************************************************************
    'REVISED BY     -  ASHISH SHARMA
    'REVISED ON     -  24 AUG 2020
    'PURPOSE        -  102027599 - IRN CHANGES
    '***************************************************************************************************

    Dim mStrCustMst As String
    Dim mresult As New ClsResultSetDB
    'Dim intnumber As Integer

    Dim mintIndex As Short
    Dim WindowHnd As Integer
    Dim salesconf As String
    Dim mExDuty As Double
    Dim mInvNo As Double
    Dim mBasicAmt As Double
    Dim msubTotal As Double
    Dim mOtherAmt As Double
    Dim mGrTotal As Double
    Dim mStAmt As Double
    Dim mFrAmt As Double
    'Dim mDoc_No As Integer
    Dim mCustmtrl As Double
    Dim mInvType As String
    Dim mSubCat As String
    Dim mAccount_Code As String
    Dim mupdatestock As Boolean
    Dim mupdatepo As Boolean
    Dim strsalesheadersql As String
    Dim strsaledetails As String
    Dim strSQLDuePayment As String
    Dim strupdate As String
    Dim StrDeletesaledtl As String
    Dim strDeletesalesconf As String
    Dim strupdateitbalmst As String
    Dim strupdatecustodtdtl As String
    Dim mCust_Ref, mAmendment_No As String
    Dim saleschallan As String
    Dim mSalesChallanLorryNo As String
    Dim ValidRecord As Boolean
    Dim mexchange_rate As Double
    'MTL Sharjah Exchange Rate Start
    Dim exchange_rate_aed As Double = 0
    'MTL Sharjah Exchange Rate End
    Dim mstrMasterString As String
    Dim mstrDetailString As String
    Dim mstrReportFilename As String
    Dim mSaleConfNo As Double
    Dim salesDtl As String
    '---------------------------------
    Dim sngAdditionalDuty, sngExcisePer, sngCessOnCVD As Double
    Dim dtCurrentDate As Date
    '------------------------------
    'Code Added By Arul on 27-08-2005
    Dim STREXPDET As String
    'Addition ends here
    Dim mstrDespAdvice As String
    'Added for Issue ID 21054 Starts
    Dim mstrupdatebarpalmst As String
    'Added for Issue ID 21054 Ends

    Dim mblnASNExist As Boolean
    Dim mstrupdateASNdtl As String
    Dim mstrupdateASNCumFig As String
    Dim mblnInvocieforMTL As Boolean
    Dim mblnInvoicelike_MTLsharjah As Boolean
    Dim mblnEwaybill_Print As Boolean
    Dim CR As ReportDocument

    Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
        Dim intCount As Short
        With fpInvoice
            For intCount = 1 To .MaxRows
                .Col = 1
                .Row = intCount
                .Value = CStr(System.Windows.Forms.CheckState.Unchecked)
            Next intCount
        End With
    End Sub
    Private Sub cmdClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClose.Click
        Me.Close()
    End Sub
    Private Sub cmdSave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSave.Click
        Dim intCount As Short
        Dim blnFound As Boolean

        If Trim(txtInvoice.Text) = "" Then
            MsgBox("Please select invoice for which Excise Format is to be printed.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, "empower")
            txtInvoice.Focus()
            Exit Sub
        End If
        With fpInvoice
            blnFound = False
            For intCount = 1 To .MaxRows
                .Row = intCount
                .Col = 1
                If CBool(.Value) = True Then
                    blnFound = True
                End If
            Next intCount

            If Not blnFound Then
                If (MsgBox("No Invoice selected for Fresh Credit, Still Do you want to continue?", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2 + MsgBoxStyle.Information, "empower")) = MsgBoxResult.No Then
                    .Row = 1
                    .Col = 1
                    .Focus()
                    Exit Sub
                End If
            End If

            If MsgBox("Once saved, It cann't be reverted back, Are you sure to proceed?", MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2 + MsgBoxStyle.Information, "empower") = MsgBoxResult.No Then
                Exit Sub
            End If

        End With
        Call SaveData()
    End Sub
    Private Sub cmdUnitCodeList_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdUnitCodeList.Click
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Show name of Unit on click of this
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Call ShowCode_Desc("SELECT Unt_CodeID,Unt_UnitName FROM Gen_UnitMaster WHERE Unt_Status=1 and Unt_CodeID='" & gstrUnitId & "'", txtUnitCode)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub ctlFormHeader1_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        On Error GoTo ErrHandler
        Call ShowHelp("HLP" & Mid(Me.Name, 4, Len(Me.Name)) & ".htm")
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub Opt_no_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles Opt_no.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then System.Windows.Forms.SendKeys.Send("{TAB}")
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub Opt_yes_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles Opt_yes.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then System.Windows.Forms.SendKeys.Send("{TAB}")
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub optExciseFormat_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optExciseFormat.CheckedChanged
        If eventSender.Checked Then
            If optExciseFormat.Checked = True Then
                fraExciseFormat.Visible = True
                cmdSave.Enabled = True
                cmdCancel.Enabled = True
            Else
                fraExciseFormat.Visible = False
            End If
        End If
    End Sub
    Private Sub optExciseFormat_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles optExciseFormat.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then
            System.Windows.Forms.SendKeys.Send("{TAB}")
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtInvoice_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtInvoice.TextChanged
        If Len(Trim(txtInvoice.Text)) = 0 Then
            fpInvoice.MaxRows = 0
            txtOTLNo.Text = ""
            txtLorryNo.Text = ""

            txtFreight.Text = ""
        End If
    End Sub
    Private Sub txtUnitCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUnitCode.TextChanged
        If Trim(txtUnitCode.Text) = "" Then
            txtInvoice.Enabled = False
            txtInvoice.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            cmdHelp.Enabled = False
        End If
    End Sub
    Private Sub txtUnitCode_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUnitCode.Enter
        On Error GoTo ErrHandler
        With txtUnitCode
            .SelectionStart = 0 : .SelectionLength = Len(Trim(.Text))
        End With
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub txtUnitCode_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtUnitCode.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If Shift <> 0 Then Exit Sub
        If KeyCode = System.Windows.Forms.Keys.F1 Then Call cmdUnitCodeList_Click(cmdUnitCodeList, New System.EventArgs())
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub txtUnitCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtUnitCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then
            System.Windows.Forms.SendKeys.Send("{TAB}")
        ElseIf KeyAscii = 187 Or KeyAscii = 166 Or KeyAscii = 164 Or KeyAscii = 172 Then
            KeyAscii = 0
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub TxtUnitCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtUnitCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   to validate data through accounts COM
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim strUnitDesc As String
        Dim mobjGLTrans As New prj_GLTransactions.cls_GLTransactions(gstrUnitId, GetServerDate)
        strUnitDesc = mobjGLTrans.GetUnit(gstrUnitId, ConnectionString:=gstrCONNECTIONSTRING)
        If Trim(txtUnitCode.Text) = "" Then GoTo EventExitSub
        If CheckString(strUnitDesc) <> "Y" Then
            MsgBox(CheckString(strUnitDesc), MsgBoxStyle.Critical, "empower")
            txtUnitCode.Text = ""
            txtInvoice.Enabled = False
            txtInvoice.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            cmdHelp.Enabled = False
            Cancel = True
        Else
            txtInvoice.Enabled = True
            txtInvoice.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            cmdHelp.Enabled = True
        End If
        GoTo EventExitSub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Show help of Help buttons
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        Dim strHelp As Object
        Dim strQry As String = String.Empty
        Dim listHelpFlag As Boolean = False
        On Error GoTo Err_Handler
        With Me.txtInvoice
            If Len(Trim(.Text)) = 0 Then
                If Opt_yes.Checked = True Then
                    strHelp = ShowList(1, .MaxLength, "", "Doc_No", "location_code", "SalesChallan_dtl", " and Invoice_Type = 'EXP' and Sub_category in('E','S','A') and print_flag=0 and location_code='" & Trim(txtUnitCode.Text) & "'")
                ElseIf Opt_no.Checked = True Then
                    '102027599
                    If mblnEwaybill_Print Then
                        If chkPrintReprint.Checked Then
                            strQry = "SELECT S.Doc_No,S.location_code from SalesChallan_dtl S where S.unit_code='" & gstrUnitId & "' and S.location_code='" & Trim(txtUnitCode.Text) & "' and S.Invoice_Type = 'EXP' and S.Sub_category in('E','S','A') and S.print_flag=1 and S.EWAY_IRN_REQUIRED='N' "
                            strQry += " UNION Select S.Doc_No,S.location_code from SalesChallan_Dtl S LEFT JOIN SALESCHALLAN_DTL_IRN I ON I.UNIT_CODE=S.UNIT_CODE AND I.DOC_NO=S.DOC_NO where  S.UNIT_CODE = '" & gstrUnitId & "' and S.Invoice_Type = 'EXP' and S.Sub_category in('E','S','A') "
                            strQry += " AND ((S.EWAY_IRN_REQUIRED='E' AND ISNULL(S.EWAY_BILL_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='I' AND ISNULL(I.IRN_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='B' AND ISNULL(S.EWAY_BILL_NO,'')<>'' AND ISNULL(I.IRN_NO,'')<>'')) "
                            strQry += " AND S.Location_Code='" & Trim(txtUnitCode.Text) & "'" & " and S.print_flag =1 "
                            strQry += " AND EXISTS (SELECT TOP 1 1 FROM FIRSTTIME_INVOICEPRINTING F (NOLOCK) WHERE F.UNIT_CODE =S.UNIT_CODE AND F.DOC_NO=S.DOC_NO) "
                        Else
                            strQry = "Select S.Doc_No,S.location_code from SalesChallan_Dtl S LEFT JOIN SALESCHALLAN_DTL_IRN I ON I.UNIT_CODE=S.UNIT_CODE AND I.DOC_NO=S.DOC_NO where  S.UNIT_CODE = '" & gstrUnitId & "' and S.Invoice_Type = 'EXP' and S.Sub_category in('E','S','A') "
                            strQry += " AND ((S.EWAY_IRN_REQUIRED='E' AND ISNULL(S.EWAY_BILL_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='I' AND ISNULL(I.IRN_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='B' AND ISNULL(S.EWAY_BILL_NO,'')<>'' AND ISNULL(I.IRN_NO,'')<>'')) "
                            strQry += " AND S.Location_Code='" & Trim(txtUnitCode.Text) & "'" & " and S.print_flag =1 "
                            strQry += " AND NOT EXISTS (SELECT TOP 1 1 FROM FIRSTTIME_INVOICEPRINTING F (NOLOCK) WHERE F.UNIT_CODE =S.UNIT_CODE AND F.DOC_NO=S.DOC_NO) "
                        End If
                        strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQry)
                        If UBound(strHelp) <> -1 Then
                            If strHelp(0) <> "0" Then
                                txtInvoice.Text = Trim(strHelp(0))
                                Call ShowInvoiceDetail(Me.txtInvoice.Text)
                            Else
                                Call ConfirmWindow(10512, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION)
                                txtInvoice.Focus()
                            End If
                        End If
                        listHelpFlag = True
                    Else
                        strHelp = ShowList(1, .MaxLength, "", "Doc_No", "location_code", "SalesChallan_dtl", " and Invoice_Type = 'EXP' and Sub_category in('E','S','A') and print_flag=1 and location_code='" & Trim(txtUnitCode.Text) & "'")
                    End If
                ElseIf optExciseFormat.Checked = True Then
                    strHelp = ShowList(1, .MaxLength, "", "Doc_No", "location_code", "SalesChallan_dtl", " and Invoice_Type = 'EXP' and Sub_category = 'E' and bill_flag=1  and location_code='" & Trim(txtUnitCode.Text) & "'")
                End If

                If Not listHelpFlag Then
                    If Val(strHelp) = -1 Then ' No record
                        Call ConfirmWindow(10512, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION)
                        .Focus()
                    Else
                        Me.txtInvoice.Text = strHelp
                        Call ShowInvoiceDetail(Me.txtInvoice.Text)
                    End If
                End If
            Else
                If Opt_yes.Checked = True Then
                    strHelp = ShowList(1, .MaxLength, .Text, "Doc_No", "location_code", "SalesChallan_dtl", " and Invoice_Type = 'EXP' and Sub_category in('E','S','A') and print_flag=0 and location_code='" & Trim(txtUnitCode.Text) & "'")
                ElseIf Opt_no.Checked = True Then
                    '102027599
                    If mblnEwaybill_Print Then
                        If chkPrintReprint.Checked Then
                            strQry = "SELECT S.Doc_No,S.location_code from SalesChallan_dtl S where S.unit_code='" & gstrUnitId & "' and S.location_code='" & Trim(txtUnitCode.Text) & "' and S.Invoice_Type = 'EXP' and S.Sub_category in('E','S','A') and S.print_flag=1 and S.EWAY_IRN_REQUIRED='N' and S.doc_no='" & Trim(txtInvoice.Text) & "' "
                            strQry += " UNION Select S.Doc_No,S.location_code from SalesChallan_Dtl S LEFT JOIN SALESCHALLAN_DTL_IRN I ON I.UNIT_CODE=S.UNIT_CODE AND I.DOC_NO=S.DOC_NO where  S.UNIT_CODE = '" & gstrUnitId & "' and S.Invoice_Type = 'EXP' and S.Sub_category in('E','S','A') "
                            strQry += " AND ((S.EWAY_IRN_REQUIRED='E' AND ISNULL(S.EWAY_BILL_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='I' AND ISNULL(I.IRN_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='B' AND ISNULL(S.EWAY_BILL_NO,'')<>'' AND ISNULL(I.IRN_NO,'')<>'')) "
                            strQry += " AND S.Location_Code='" & Trim(txtUnitCode.Text) & "'" & " and S.print_flag =1 and S.doc_no='" & Trim(txtInvoice.Text) & "' "
                            strQry += " AND EXISTS (SELECT TOP 1 1 FROM FIRSTTIME_INVOICEPRINTING F (NOLOCK) WHERE F.UNIT_CODE =S.UNIT_CODE AND F.DOC_NO=S.DOC_NO) "
                        Else
                            strQry += "Select S.Doc_No,S.location_code from SalesChallan_Dtl S LEFT JOIN SALESCHALLAN_DTL_IRN I ON I.UNIT_CODE=S.UNIT_CODE AND I.DOC_NO=S.DOC_NO where  S.UNIT_CODE = '" & gstrUnitId & "' and S.Invoice_Type = 'EXP' and S.Sub_category in('E','S','A') "
                            strQry += " AND ((S.EWAY_IRN_REQUIRED='E' AND ISNULL(S.EWAY_BILL_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='I' AND ISNULL(I.IRN_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='B' AND ISNULL(S.EWAY_BILL_NO,'')<>'' AND ISNULL(I.IRN_NO,'')<>'')) "
                            strQry += " AND S.Location_Code='" & Trim(txtUnitCode.Text) & "'" & " and S.print_flag =1 and S.doc_no='" & Trim(txtInvoice.Text) & "' "
                            strQry += " AND NOT EXISTS (SELECT TOP 1 1 FROM FIRSTTIME_INVOICEPRINTING F (NOLOCK) WHERE F.UNIT_CODE =S.UNIT_CODE AND F.DOC_NO=S.DOC_NO) "
                        End If
                        strHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strQry)
                        If UBound(strHelp) <> -1 Then
                            If strHelp(0) <> "0" Then
                                txtInvoice.Text = Trim(strHelp(0))
                                Call ShowInvoiceDetail(Me.txtInvoice.Text)
                            Else
                                Call ConfirmWindow(10512, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION)
                                txtInvoice.Focus()
                            End If
                        End If
                        listHelpFlag = True
                    Else
                        strHelp = ShowList(1, .MaxLength, .Text, "Doc_No", "location_code", "SalesChallan_dtl", " and Invoice_Type = 'EXP' and Sub_category in('E','S','A') and print_flag=1 and location_code='" & Trim(txtUnitCode.Text) & "'")
                    End If
                ElseIf optExciseFormat.Checked = True Then
                    strHelp = ShowList(1, .MaxLength, .Text, "Doc_No", "location_code", "SalesChallan_dtl", " and Invoice_Type = 'EXP' and Sub_category = 'E' and bill_flag=1 and location_code='" & Trim(txtUnitCode.Text) & "'")
                End If

                If Not listHelpFlag Then
                    If Val(strHelp) = -1 Then ' No record
                        Call ConfirmWindow(10512, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION)
                        .Focus()
                    Else
                        Me.txtInvoice.Text = strHelp
                        Call ShowInvoiceDetail(Me.txtInvoice.Text)
                        .Focus()
                    End If
                End If
            End If
        End With
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub ShowInvoiceDetail(ByVal pstrInvoiceno As String)
        '*******************************************************************************
        'Author             :   Manoj Vaish
        'Argument(s)if any  :   Invoice No.
        'Return Value       :   NA
        'Function           :   To Show Lorry No,OTL No and Freight Amount
        'Comments           :   NA
        'Creation Date      :   06 Mar 2009 Issue ID eMpro-20090227-27987
        '*******************************************************************************
        Dim rsshowdata As ClsResultSetDB
        Dim strquery As String

        On Error GoTo Err_Handler

        rsshowdata = New ClsResultSetDB
        strquery = "select lorry_no,otl_no,frieght_amount from saleschallan_dtl where unit_code='" & gstrUnitId & "' and  doc_no='" & pstrInvoiceno & "'"
        rsshowdata.GetResult(strquery, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        If rsshowdata.GetNoRows > 0 Then
            txtLorryNo.Text = Trim(rsshowdata.GetValue("lorry_no"))
            txtOTLNo.Text = Trim(rsshowdata.GetValue("otl_no"))
            txtFreight.Text = Val(rsshowdata.GetValue("frieght_amount"))
        Else
            txtLorryNo.Text = ""
            txtOTLNo.Text = ""
            txtFreight.Text = "0.00"
        End If
        rsshowdata.ResultSetClose()
        rsshowdata = Nothing
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdInvoice_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.UCfraRepCmd.ButtonClickEventArgs) Handles Cmdinvoice.ButtonClick
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Add Funtionality on PREVIEW/PRINT/CLOSE Button.
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        Dim rsCompMst As ClsResultSetDB
        Dim rsSalesConf As ClsResultSetDB
        Dim rssaledtl As ClsResultSetDB
        Dim rsItembal As ClsResultSetDB
        Dim Address, EccNo, RegNo, Range, Phone As String
        Dim CST, PLA, Fax, EMail, UPST, Division As String
        Dim ExpCode As String
        Dim Commissionerate As String
        Dim strsql As String
        Dim strCompMst, SALEDTL As String
        Dim ItemCode As New VB6.FixedLengthString(20)
        Dim strStockLocation As String
        Dim intLoopcount, intRow, int_counter As Short
        Dim salesQuantity As Double
        Dim bol_cur_bal As Boolean 'For checking if for any item cur_bal is less than sales qty
        Dim str_cur_bal As String 'To store the items for which the current stock is less than the sales qty
        Dim str_current_bal As New VB6.FixedLengthString(20)
        Dim strRetval As String
        Dim strInvoiceDate As String
        Dim strAccountCode As String
        Dim strSuffix As String
        Dim dblExistingInvNo As Double
        Dim rsSalesChallandtl As ClsResultSetDB
        Dim objDrCr As New prj_DrCrNote.cls_DrCrNote(GetServerDate)
        Dim blnExpCatiaInvFormat As Boolean
        Dim strInvSubType As String
        Dim frmRpt As New eMProCrystalReportViewer
        Dim CR As New ReportDocument
        Dim rsCustorddtl As ClsResultSetDB
        Dim strCustdrgno As String
        Dim strCustref As String
        Dim strAmendmentNo As String
        Dim LUT_NO As String
        Dim LUT_DATEFROMDATETO As String
        Dim blnIsPDFExported As Boolean = False
        Dim DeliveredAdd As String
        Dim shipname As String


        On Error GoTo Err_Handler
        rssaledtl = New ClsResultSetDB
        If e.Button = UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE Then
            Me.Close()
            Exit Sub
        Else
            Call ValidSelection()
            If ValidRecord = False Then Exit Sub
        End If
        rsSalesConf = New ClsResultSetDB
        rsSalesChallandtl = New ClsResultSetDB
        rsSalesChallandtl.GetResult("SELECT * FROM  saleschallan_dtl where unit_code='" & gstrUnitId & "' and doc_no = " & txtInvoice.Text & " and Location_code = '" & Trim(txtUnitCode.Text) & "'")
        strInvSubType = rsSalesChallandtl.GetValue("Sub_category")
        strInvoiceDate = VB6.Format(rsSalesChallandtl.GetValue("Invoice_Date"), gstrDateFormat)
        strAccountCode = rsSalesChallandtl.GetValue("account_code")
        strCustref = rsSalesChallandtl.GetValue("cust_ref")
        strAmendmentNo = rsSalesChallandtl.GetValue("Amendment_no")

        If Len(Trim(rsSalesChallandtl.GetValue("ServiceInvoiceformatExport"))) = 0 Then
            blnExpCatiaInvFormat = False
        Else
            blnExpCatiaInvFormat = rsSalesChallandtl.GetValue("ServiceInvoiceformatExport")
        End If
        rsSalesConf.GetResult("Select Stock_Location, SecReport_filename,Report_filename,Suffix, ExciseFormatReport from SaleConf Where unit_code='" & gstrUnitId & "' and Invoice_Type ='EXP' AND Location_Code ='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0")
        If blnExpCatiaInvFormat = True Then
            mstrReportFilename = rsSalesConf.GetValue("SecReport_filename")
        Else
            mstrReportFilename = rsSalesConf.GetValue("Report_filename")
        End If
        If optExciseFormat.Checked = True Then
            mstrReportFilename = rsSalesConf.GetValue("ExciseFormatReport")
        End If
        strStockLocation = rsSalesConf.GetValue("Stock_Location")
        strSuffix = rsSalesConf.GetValue("Suffix")
        If Len(Trim(strStockLocation)) = False Then
            MsgBox("Please Define Stock Location in SalesConf for Export Invoice", MsgBoxStyle.Information, "empower")
            Exit Sub
        End If
        rsSalesConf.ResultSetClose()
        rsSalesConf = Nothing
        If Opt_yes.Checked = True Then
            rsSalesConf = New ClsResultSetDB
            rsSalesConf.GetResult("Select Stock_Location, SecReport_filename,Report_filename from SaleConf Where unit_code='" & gstrUNITID & "' and Invoice_Type ='EXP' AND Location_Code ='" & Trim(txtUnitCode.Text) & "'and datediff(dd,'" & getDateForDB(strInvoiceDate) & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0")
            If blnExpCatiaInvFormat = True Then
                mstrReportFilename = rsSalesConf.GetValue("SecReport_filename")
            Else
                mstrReportFilename = rsSalesConf.GetValue("Report_filename")
            End If
            strStockLocation = rsSalesConf.GetValue("Stock_Location")
            If Len(Trim(strStockLocation)) = False Then
                MsgBox("Please Define Stock Location in SalesConf for Export Invoice", MsgBoxStyle.Information, "empower")
                Exit Sub
            End If
            SALEDTL = "Select Sales_Quantity,Item_code,cust_item_code from sales_Dtl where unit_code='" & gstrUNITID & "' and Doc_No = " & Me.txtInvoice.Text & " AND Location_Code='" & Trim(txtUnitCode.Text) & "'"
            rssaledtl.GetResult(SALEDTL, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            intRow = rssaledtl.GetNoRows
            rssaledtl.MoveFirst()
            rsItembal = New ClsResultSetDB
            rsCustorddtl = New ClsResultSetDB
            str_cur_bal = "Current Balance of the following items is less than the Sales Qty" & vbCrLf & " " & vbCrLf
            str_cur_bal = str_cur_bal & "S.No   Location Code      Item Code               Current Balance     Sales Quantity" & vbCrLf
            str_cur_bal = str_cur_bal & "----------------------------------------------------------------------------------------------------" & vbCrLf
            If e.Button = UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT_TO_PRINTER Then
                int_counter = 0
                For intLoopcount = 1 To intRow
                    ItemCode.Value = rssaledtl.GetValue("Item_code")
                    salesQuantity = rssaledtl.GetValue("Sales_quantity")
                    strCustdrgno = rssaledtl.GetValue("Cust_item_code")
                    rsItembal.GetResult("Select Cur_bal from Itembal_Mst where unit_code='" & gstrUNITID & "' and Item_code = '" & ItemCode.Value & "'and Location_code ='" & strStockLocation & "'", ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)

                    If salesQuantity > rsItembal.GetValue("Cur_Bal") Then
                        bol_cur_bal = True
                        int_counter = int_counter + 1
                        str_current_bal.Value = CStr(rsItembal.GetValue("Cur_Bal"))
                        str_cur_bal = str_cur_bal & int_counter & ".        " & strStockLocation & "                      " & ItemCode.Value & "    " & str_current_bal.Value & "        " & salesQuantity & vbCrLf
                    End If
                    'If mblnInvocieforMTL = False Or mblnInvoicelike_MTLsharjah = False Then
                    If Not (mblnInvocieforMTL = True Or mblnInvoicelike_MTLsharjah = True) Then
                        'issue id 10266201
                        rsCustorddtl.GetResult("Select openso,balance_Qty = order_qty - Despatch_qty from Cust_ord_dtl where " & _
                                                "Account_code ='" & strAccountCode & "'" & " and Item_code ='" & _
                                                ItemCode.Value.Trim & "' and cust_drgNo ='" & strCustdrgno & _
                                                "' and Authorized_flag = 1 and cust_ref = '" & strCustref & "' and amendment_no='" & strAmendmentNo & "'")

                        If rsCustorddtl.GetValue("openso") = False Then 'should not be open PO
                            If salesQuantity > rsCustorddtl.GetValue("balance_Qty") Then
                                MsgBox("Balance Quantity available in SO for Customer Part code [ " & strCustdrgno & "] is " & Val(rsCustorddtl.GetValue("Balance_Qty")) & ".", MsgBoxStyle.Information, ResolveResString(100))
                                Exit Sub
                            End If
                        End If
                        'issue id 10266201
                    End If
                    rssaledtl.MoveNext()
                Next
                If bol_cur_bal = True Then
                    MsgBox(str_cur_bal, MsgBoxStyle.Information + MsgBoxStyle.SystemModal, "empower")
                    Exit Sub
                End If
            End If
        End If

        Call InitializeVariable()
        If ValuetoVariables() = True Then
            Exit Sub
        End If
        If mblnInvocieforMTL = False Then

            If DataExist("SELECT SOUPLD_FOREXCURRENCY FROM CUSTOMER_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND SOUPLD_FOREXCURRENCY=1 AND CUSTOMER_CODE='" & strAccountCode & "'") = True Then
                If DataExist("SELECT TOP 1 1 FROM SALES_dTL WHERE UNIT_CODE = '" & gstrUNITID & "' AND DOC_NO= '" & txtInvoice.Text & "' AND len(isnull(EXTERNAL_SALESORDER_NO,''))>0") Then
                    mstrReportFilename = mstrReportFilename + "_External_SO"
                End If
            End If
        End If

        Call UpdateinSale_Dtl()
        Call updatesalesconfandsaleschallan()
        Call updateLorryNo()
        mP_Connection.BeginTrans()
        mP_Connection.Execute("set Dateformat 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        mP_Connection.Execute(mSalesChallanLorryNo, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        mP_Connection.CommitTrans()

        If Not CreateStringForAccounts() Then Exit Sub
        rsCompMst = New ClsResultSetDB
        strsql = "{SalesChallan_Dtl.Unit_Code}='" & gstrUNITID & "' and {SalesChallan_Dtl.Location_Code}='" & Trim(txtUnitCode.Text) & "' and {SalesChallan_Dtl.Doc_No} =" & Trim(txtInvoice.Text) & " and {SalesChallan_Dtl.Invoice_Type} = 'EXP'  and {SalesChallan_Dtl.Sub_Category} = 'E'"
        strCompMst = "Select * from Company_Mst where unit_code='" & gstrUNITID & "'"
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
            ExpCode = rsCompMst.GetValue("Exporter_Code")
        End If
        Address = gstr_WRK_ADDRESS1 & gstr_WRK_ADDRESS2
        rsCompMst.ResultSetClose()
        If Trim(mstrReportFilename) = "" Or mstrReportFilename.Trim.ToLower = "unknown" Then
            MsgBox("No Report filename selected for the invoice. Invoice cannot be printed", MsgBoxStyle.Information, "empower")
            Exit Sub
        End If
        CR = frmRpt.GetReportDocument
        CR.Load(My.Application.Info.DirectoryPath & "\Reports\" & mstrReportFilename & ".rpt")
        If optExciseFormat.Checked = True Then
            CR.DataDefinition.FormulaFields("CompanyName").Text = "'" & "For " & gstrCOMPANY & "'"
            CR.DataDefinition.FormulaFields("companyAddress").Text = "'" & Address & "'"
            CR.DataDefinition.FormulaFields("exchangerate").Text = "'" & mexchange_rate & "'"
            CR.DataDefinition.FormulaFields("RegNo").Text = "'" & RegNo & "'"
            CR.DataDefinition.FormulaFields("ECC").Text = "'" & EccNo & "'"
            CR.DataDefinition.FormulaFields("Range").Text = "'" & Range & "'"
            CR.DataDefinition.FormulaFields("Phone").Text = "'" & Phone & "'"
            CR.DataDefinition.FormulaFields("Fax").Text = "'" & Fax & "'"
            CR.DataDefinition.FormulaFields("EMail").Text = "'" & EMail & "'"
            CR.DataDefinition.FormulaFields("PLA").Text = "'" & PLA & "'"
            CR.DataDefinition.FormulaFields("UPST").Text = "'" & UPST & "'"
            CR.DataDefinition.FormulaFields("CST").Text = "'" & CST & "'"
            CR.DataDefinition.FormulaFields("Division").Text = "'" & Division & "'"
            CR.DataDefinition.FormulaFields("Commissionerate").Text = "'" & Commissionerate & "'"
        Else
            CR.DataDefinition.FormulaFields("Comp_name").Text = "'" & gstrCOMPANY & "'"
            CR.DataDefinition.FormulaFields("comp_add").Text = "'" & Address & "'"
            CR.DataDefinition.FormulaFields("exchangerate").Text = "'" & mexchange_rate & "'"
            CR.DataDefinition.FormulaFields("ExpCode").Text = "'" & ExpCode & "'"
            'MTL Sharjah Exchange Rate Start
            If mblnInvocieforMTL Then
                CR.DataDefinition.FormulaFields("Exchange_Rate_AED").Text = "'" & exchange_rate_aed & "'"
            End If
            'MTL Sharjah Exchange Rate End
        End If
        If Opt_yes.Checked = False Then
            If Not optExciseFormat.Checked = True Then
                If Len(Trim(strSuffix)) > 0 Then
                    If Val(strSuffix) > 0 Then
                        dblExistingInvNo = Val(Mid(CStr(mInvNo), Len(Trim(strSuffix)) + 1))
                    Else
                        dblExistingInvNo = CDbl(txtInvoice.Text)
                    End If
                Else
                    dblExistingInvNo = CDbl(txtInvoice.Text)
                End If
                If GetPlantName() = "HILEX" Then
                    CR.DataDefinition.FormulaFields("CurrentNo").Text = "'" & txtInvoice.Text & "'"
                Else
                    CR.DataDefinition.FormulaFields("CurrentNo").Text = "'" & dblExistingInvNo & "'"
                End If

            End If
        Else
            If GetPlantName() = "HILEX" Then
                CR.DataDefinition.FormulaFields("CurrentNo").Text = "'" & mInvNo & "'"
            Else
                CR.DataDefinition.FormulaFields("CurrentNo").Text = "'" & mSaleConfNo & "'"
            End If

        End If
        If GetPlantName() = "HILEX" Then
            CR.DataDefinition.FormulaFields("InvoiceDate").Text = "'" & VB6.Format(GetServerDateTime(), "dd/MM/yyyy") & "'"
        End If

        If DataExist("Select top 1 1 from SaleConf Where unit_code = '" & gstrUNITID & "' and  Invoice_type = 'EXP' and sub_type='E' and Location_code ='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0 And REQD_LUTNO = 1") = True Then
            LUT_NO = CStr(Find_Value("select TOP 1 LUT_NO from COMPANY_LUT_DETAILS where unit_code='" & gstrUNITID & "' and status=1 and datediff(dd,'" & getDateForDB(strInvoiceDate) & "',valid_from)<=0  and datediff(dd,valid_to,'" & getDateForDB(strInvoiceDate) & "')<=0"))
            LUT_DATEFROMDATETO = CStr(Find_Value("SELECT TOP 1 + 'DATE FROM ' + CONVERT(VARCHAR(12),CONVERT(DATETIME ,VALID_FROM,106)) + ' DATE TO '  + CONVERT(VARCHAR(12),CONVERT(DATETIME ,VALID_TO,106))  FROM  COMPANY_LUT_DETAILS where unit_code='" & gstrUNITID & "' and status=1 and LUT_NO='" & LUT_NO & "'"))
            If LUT_NO <> "" Then
                CR.DataDefinition.FormulaFields("LUTNO").Text = "'LUT NO:" & LUT_NO & "'"
                CR.DataDefinition.FormulaFields("LUT_DATEFROMDATETO").Text = "'" & LUT_DATEFROMDATETO & "'"
            End If
        End If
        If mblnInvocieforMTL = True Then
            rsCompMst = New ClsResultSetDB
            rsCompMst.GetResult("Select * from VW_SHIPPINGCODE_DESC_EXPORTINVOICE_MTL   where UNIT_CODE = '" & gstrUNITID & "'and INVOICE_NO = '" & txtInvoice.Text & "'")
            If rsCompMst.GetNoRows > 0 Then

                DeliveredAdd = Trim(rsCompMst.GetValue("Ship_address1"))
                If Len(Trim(DeliveredAdd)) Then
                    DeliveredAdd = Trim(DeliveredAdd) & "," & Trim(rsCompMst.GetValue("Ship_address2"))
                Else
                    DeliveredAdd = Trim(rsCompMst.GetValue("Ship_address2"))
                End If
            End If
            shipname = Trim(rsCompMst.GetValue("Shipping_Desc"))
            rsCompMst.ResultSetClose()

            CR.DataDefinition.FormulaFields("shipname").Text = "'" & shipname & "'"
            CR.DataDefinition.FormulaFields("Address2").Text = "'" & DeliveredAdd & "'"
        End If


        CR.RecordSelectionFormula = "{SalesChallan_Dtl.Unit_Code}='" & gstrUNITID & "' and {SalesChallan_Dtl.Location_Code}='" & Trim(txtUnitCode.Text) & "' and {SalesChallan_Dtl.Doc_No}=" & Trim(txtInvoice.Text)
        '10646968        
        CR.RecordSelectionFormula = CR.RecordSelectionFormula & " And {SalesChallan_Dtl.Invoice_Date} >= {SaleConf.Fin_start_date} and {SalesChallan_Dtl.Invoice_Date} <= {SaleConf.Fin_end_date}"
        Select Case e.Button
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                Me.Close()
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT_TO_WINDOW
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                frmRpt.ShowPrintButton = True
                frmRpt.ShowExportButton = True

                '102027599
                If Opt_yes.Checked = False Then
                    If mblnEwaybill_Print Then
                        Call IRN_QRBarcode()
                    End If
                End If

                If Opt_yes.Checked = True And AllowASNPrinting(strAccountCode) = True Then
                    If mblnASNExist = True Then
                        mP_Connection.Execute("Update CreatedASN Set ASN_NO='" & Trim$(txtASNNumber.Text) & "',Updatedon=getdate() where unit_code='" & gstrUNITID & "' and doc_no='" & Trim$(Me.txtInvoice.Text) & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    Else
                        mP_Connection.Execute("Insert into CreatedASN(Doc_no,ASN_NO,Createdon,CreatedBy,Updatedon,UNIT_CODE) values ('" & Trim$(Me.txtInvoice.Text) & "','" & Trim$(txtASNNumber.Text) & "',getdate(),'" & mP_User & "',getdate(),'" & gstrUNITID & "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End If
                End If
                frmRpt.Show()
                'CHANGED FOR MTL SHARJAH -19 SEP 2022 BY PRASHANT RAJPAL
                If GetPlantName() = "MTL" Then 'ÓNLY REPRINT
                    If Opt_yes.Checked = False Then
                        EXPORTINVOICETOPDF_ONPRINTREPRINT(strAccountCode, "EXP", "E", CR)
                    End If
                End If

                'CHANGED ENDED FOR MTL SHARJAH -19 SEP 2022 BY PRASHANT RAJPAL
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                WindowHnd = GetActiveWindow
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT_TO_PRINTER
                '102027599

                If Opt_yes.Checked = False Then
                    If mblnEwaybill_Print Then
                        Call IRN_QRBarcode()
                    End If
                End If
                frmRpt.SetReportDocument()
                '102027599
                'CHANGED FOR MTL SHARJAH -19 SEP 2022 BY PRASHANT RAJPAL
                If GetPlantName() = "MTL" Then 'ÓNLY REPRINT
                    If Opt_yes.Checked = False Then
                        EXPORTINVOICETOPDF_ONPRINTREPRINT(strAccountCode, "EXP", "E", CR)
                    End If
                End If

                'CHANGED ENDED FOR MTL SHARJAH -19 SEP 2022 BY PRASHANT RAJPAL

                If mblnEwaybill_Print = False Then
                    CR.PrintToPrinter(1, False, 0, 0)

                    If GetPlantName() = "MTL" Then
                        EXPORTINVOICETOPDF_ONPRINTREPRINT(strAccountCode, "EXP", "E", CR)

                    End If
                Else
                    If chkPrintReprint.Checked And Opt_no.Checked Then
                        CR.PrintToPrinter(1, False, 0, 0)
                    Else
                        If Opt_no.Checked Then
                            If Not DataExist("SELECT TOP 1 1 FROM FIRSTTIME_INVOICEPRINTING  WHERE UNIT_CODE = '" & gstrUNITID & "' and doc_no= " & Trim(txtInvoice.Text) & "") = True Then
                                mP_Connection.Execute("Insert into FIRSTTIME_INVOICEPRINTING(unit_code,doc_no,ent_dt,ent_userid) values('" & gstrUNITID & "','" & Trim(txtInvoice.Text) & "',getdate(),'" & mP_User & "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                CR.PrintToPrinter(1, False, 0, 0)
                            Else
                                CR.PrintToPrinter(1, False, 0, 0)
                            End If
                        End If
                    End If
                End If
                If Opt_yes.Checked = True Then
                    If ConfirmWindow(10344, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                        mP_Connection.BeginTrans()
                        mP_Connection.Execute("set Dateformat 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        mP_Connection.Execute(salesconf, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        mP_Connection.Execute(saleschallan, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        mP_Connection.Execute(salesDtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        mP_Connection.Execute(mstrDespAdvice, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        mP_Connection.Execute(mstrupdatebarpalmst, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        mP_Connection.Execute(STREXPDET, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        If mupdatepo = True Then
                            mP_Connection.Execute(strupdatecustodtdtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        End If
                        If mupdatestock = True Then
                            mP_Connection.Execute(strupdateitbalmst, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        End If
                        'If UCase(strInvSubType) <> "S" Then
                        '    strRetval = objDrCr.SetARInvoiceDocument(gstrUNITID, mstrMasterString, mstrDetailString, prj_GLTransactions.cls_GLTransactions.udtOperationType.optInsert, gstrCONNECTIONSTRING)
                        '    strRetval = CheckString(strRetval)
                        '    If Not strRetval = "Y" Then
                        '        MsgBox(strRetval, MsgBoxStyle.Information, "empower")
                        '        mP_Connection.RollbackTrans()
                        '        Exit Sub
                        '    End If
                        'End If
                        '10174688
                        If AllowASNTextFileGeneration(strAccountCode) = True Then
                            mP_Connection.Execute("UPDATE MKT_ASN_INVDTL SET DOC_NO=" & Trim(mInvNo) & " where Unit_code='" & gstrUNITID & "' and dOC_NO=" & Trim(txtInvoice.Text) & "", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            If CheckASNTYPE(strAccountCode) = "GROUPO" Then
                                If GROUPOASNFileGeneration(mInvNo, strAccountCode) = False Then
                                    mP_Connection.RollbackTrans()
                                    Exit Sub
                                Else
                                    If Len(mstrupdateASNdtl) > 0 Then
                                        mP_Connection.Execute(mstrupdateASNdtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                        mP_Connection.Execute(mstrupdateASNCumFig, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                    End If
                                End If
                            Else
                                If FordASNFileGeneration(mInvNo, strAccountCode) = False Then
                                    mP_Connection.RollbackTrans()
                                    Exit Sub
                                Else
                                    If Len(mstrupdateASNdtl) > 0 Then
                                        mP_Connection.Execute(mstrupdateASNdtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                        mP_Connection.Execute(mstrupdateASNCumFig, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                    End If
                                End If

                            End If
                        End If
                        If Opt_yes.Checked = True Then
                            If DataExist("SELECT TOP 1 1 FROM SALES_PARAMETER WHERE INVOICE_LOCKING_ENTRY_SAMEDATE=1  and UNIT_CODE = '" & gstrUNITID & "'") Or ((GetPlantName() = "HILEX" Or GetPlantName() = "MTL")) Then
                                mP_Connection.Execute("update Saleschallan_dtl set invoice_date= Convert(varchar(12), getdate(), 106) WHERE UNIT_CODE='" + gstrUNITID + "' AND  Doc_no = " & mInvNo, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            End If
                        End If
                        If UCase(strInvSubType) <> "S" Then
                            strRetval = objDrCr.SetARInvoiceDocument(gstrUNITID, mstrMasterString, mstrDetailString, prj_GLTransactions.cls_GLTransactions.udtOperationType.optInsert, gstrCONNECTIONSTRING)
                            strRetval = CheckString(strRetval)
                            If Not strRetval = "Y" Then
                                MsgBox(strRetval, MsgBoxStyle.Information, "empower")
                                mP_Connection.RollbackTrans()
                                Exit Sub
                            End If
                        End If

                        'If AllowASNTextFileGeneration(strAccountCode) = True Then
                        '    mP_Connection.Execute("UPDATE MKT_ASN_INVDTL SET DOC_NO=" & Trim(mInvNo) & " where dOC_NO=" & Trim(txtInvoice.Text) & "", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        '    If FordASNFileGeneration(mInvNo, strAccountCode) = False Then
                        '        mP_Connection.RollbackTrans()
                        '        Exit Sub
                        '    Else
                        '        If Len(mstrupdateASNdtl) > 0 Then
                        '            'Changed for Issue Id eMpro-20090709-33409 Starts
                        '            mP_Connection.Execute(mstrupdateASNdtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        '            mP_Connection.Execute(mstrupdateASNCumFig, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        '            'Changed for Issue Id eMpro-20090709-33409 Ends
                        '        End If
                        '    End If
                        'End If

                        If Find_Value("select isnull(ALLOWSTOCKINTRANSIT,0) as ALLOWSTOCKINTRANSIT from sales_parameter where unit_code='" & gstrUNITID & "'") = True Then
                            mP_Connection.Execute("Exec PROC_STOCK_IN_TRANSIT '" & mInvNo & "','EXP','" & gstrUNITID & "','" & mP_User & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        End If
                        mP_Connection.CommitTrans()
                        MsgBox("Invoice has been locked successfully with number " & mInvNo, MsgBoxStyle.Information, "empower")
                        txtInvoice.Text = ""
                        rsCompMst = Nothing
                        rsSalesConf = Nothing
                        rssaledtl.ResultSetClose()
                        rsSalesConf = Nothing
                        rsItembal.ResultSetClose()
                        rsItembal = Nothing
                        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                    End If
                End If
                Call RefreshForm()
        End Select
        frmRpt = Nothing
        Exit Sub
Err_Handler:
        frmRpt = Nothing
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub EXPORTINVOICETOPDF_ONPRINTREPRINT(ByVal strAccountCode As String, ByVal strInvoiceType As String, ByVal strInvoiceSubType As String, ByRef RPTDoc As ReportDocument)
        ''AMIT RANA 20 Jun 2019
        Try

            Dim OBJPdfConfig As Object = SqlConnectionclass.ExecuteScalar("SELECT COUNT(*) FROM INVOICE_PDF_CONFIG (NOLOCK) WHERE UNIT_CODE='" + gstrUnitId + "' AND CUSTOMER_CODE='" + strAccountCode.ToString() + "' AND INVOICE_TYPE='" & strInvoiceType & "' AND INVOICE_SUB_TYPE='" & strInvoiceSubType & "' And IS_ACTIVE=1")
            If Val(OBJPdfConfig.ToString()) > 0 Then

                Dim strCreatedPDFPath As String = String.Empty
                Dim strRESULT As String = String.Empty
                Dim STRINVOICENO_PDF As String = String.Empty
                If Opt_no.Checked = True Then 'GENERATE THE INVOICE
                    STRINVOICENO_PDF = mInvNo
                Else
                    STRINVOICENO_PDF = txtInvoice.Text
                End If
                'RPTDoc.DataDefinition.FormulaFields("CopyName").Text = "'ORIGINAL FOR BUYER'"

                'Start code to merge original invoice and annexture
                
                Dim Frm As Object = Nothing
                Dim RPTDocTemp As ReportDocument

                    RPTDoc.Export(GetExportOptions(STRINVOICENO_PDF, strCreatedPDFPath))

                'End code to merge original invoice and annexture

                
            End If

            ''AMIT RANA 20 Jun 2019
        Catch ex As Exception
            MsgBox(ex.Message.ToString())
        End Try

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

    Private Function AllowASNPrinting(ByVal pstraccountcode As String) As Boolean
        On Error GoTo ErrHandler
        Dim strQry As String
        Dim Rs As ClsResultSetDB
        AllowASNPrinting = False
        strQry = "Select isnull(AllowASNPrinting,0) as AllowASNPrinting from customer_mst where Unit_code='" & gstrUNITID & "' and Customer_Code='" & Trim(pstraccountcode) & "'"
        Rs = New ClsResultSetDB
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
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub Opt_no_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Opt_no.CheckedChanged
        If eventSender.Checked Then
            txtInvoice.Text = ""
            If Opt_no.Checked = True Then
                fraExciseFormat.Visible = False
            End If
            Cmdinvoice.Visible = True
            txtFreight.Enabled = True
        End If

    End Sub
    Private Sub Opt_yes_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Opt_yes.CheckedChanged
        If eventSender.Checked Then
            txtInvoice.Text = ""
            If Opt_yes.Checked = True Then
                fraExciseFormat.Visible = False
            End If
            Cmdinvoice.Visible = True
        End If
    End Sub
    Private Sub txtInvoice_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtInvoice.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   At Enter Key Press Set Focus To Next Control
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo Err_Handler
        If KeyAscii = 13 Then
            If Len(Trim(Me.txtInvoice.Text)) = 0 Then
                Me.txtInvoice.Focus()
                GoTo EventExitSub
            Else
                txtinvoice_Validating(txtInvoice, New System.ComponentModel.CancelEventArgs(False))
            End If
        End If
        If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
            KeyAscii = KeyAscii
        Else
            KeyAscii = 0
        End If
        GoTo EventExitSub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtInvoice_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtInvoice.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo Err_Handler
        If KeyCode = 112 Then
            Call cmdHelp_Click(cmdHelp, New System.EventArgs())
        End If
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtinvoice_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtInvoice.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim strsql As String
        Dim strAccountCode As String
        On Error GoTo Err_Handler
        If Len(txtInvoice.Text) = 0 Then GoTo EventExitSub
        If Opt_yes.Checked = True Then
            strsql = mStrCustMst & "'" & txtInvoice.Text & "'" & " AND unit_code = '" & gstrUnitId & "' and  Bill_flag=0 AND Location_code ='" & Trim(txtUnitCode.Text) & "'  AND Invoice_Type = 'EXP'  AND Sub_category in('E','S','A')"
        ElseIf Opt_no.Checked = True Then
            '102027599
            If mblnEwaybill_Print Then
                If chkPrintReprint.Checked Then
                    strsql = "SELECT S.Doc_No,S.Invoice_type from SalesChallan_dtl S where S.unit_code='" & gstrUnitId & "' and S.location_code='" & Trim(txtUnitCode.Text) & "' and S.Invoice_Type = 'EXP' and S.Sub_category in('E','S','A') and S.Bill_flag=1 and S.EWAY_IRN_REQUIRED='N' and S.Doc_No='" & Trim(txtInvoice.Text) & "' "
                    strsql += " UNION Select S.Doc_No,S.Invoice_type from SalesChallan_Dtl S LEFT JOIN SALESCHALLAN_DTL_IRN I ON I.UNIT_CODE=S.UNIT_CODE AND I.DOC_NO=S.DOC_NO where  S.UNIT_CODE = '" & gstrUnitId & "' and S.Invoice_Type = 'EXP' and S.Sub_category in('E','S','A') "
                    strsql += " AND ((S.EWAY_IRN_REQUIRED='E' AND ISNULL(S.EWAY_BILL_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='I' AND ISNULL(I.IRN_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='B' AND ISNULL(S.EWAY_BILL_NO,'')<>'' AND ISNULL(I.IRN_NO,'')<>'')) "
                    strsql += " AND S.Location_Code='" & Trim(txtUnitCode.Text) & "'" & " and S.Bill_flag =1 and S.Doc_No='" & Trim(txtInvoice.Text) & "' "
                    strsql += " AND EXISTS (SELECT TOP 1 1 FROM FIRSTTIME_INVOICEPRINTING F (NOLOCK) WHERE F.UNIT_CODE =S.UNIT_CODE AND F.DOC_NO=S.DOC_NO) "
                Else
                    strsql = "Select S.Doc_No,S.Invoice_type from SalesChallan_Dtl S LEFT JOIN SALESCHALLAN_DTL_IRN I ON I.UNIT_CODE=S.UNIT_CODE AND I.DOC_NO=S.DOC_NO where  S.UNIT_CODE = '" & gstrUnitId & "' and S.Invoice_Type = 'EXP' and S.Sub_category in('E','S','A') "
                    strsql += " AND ((S.EWAY_IRN_REQUIRED='E' AND ISNULL(S.EWAY_BILL_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='I' AND ISNULL(I.IRN_NO,'')<>'') OR (S.EWAY_IRN_REQUIRED='B' AND ISNULL(S.EWAY_BILL_NO,'')<>'' AND ISNULL(I.IRN_NO,'')<>'')) "
                    strsql += " AND S.Location_Code='" & Trim(txtUnitCode.Text) & "'" & " and S.Bill_flag =1 and S.Doc_No='" & Trim(txtInvoice.Text) & "' "
                    strsql += " AND NOT EXISTS (SELECT TOP 1 1 FROM FIRSTTIME_INVOICEPRINTING F (NOLOCK) WHERE F.UNIT_CODE =S.UNIT_CODE AND F.DOC_NO=S.DOC_NO) "
                End If
            Else
                strsql = mStrCustMst & "'" & txtInvoice.Text & "'" & " AND unit_code = '" & gstrUnitId & "' and Bill_flag=1 AND Location_code ='" & Trim(txtUnitCode.Text) & "'  AND Invoice_Type = 'EXP'  AND Sub_category in('E','S','A')"
            End If

        End If
        If Not mresult Is Nothing Then mresult = Nothing
        mresult = New ClsResultSetDB
        mresult.GetResult(strsql)
        If Len(Trim(txtInvoice.Text)) > 0 Then 'Checking if Item Field is not Blank
            If mresult.RowCount <= 0 Then 'Checking if the Record Exists
                MsgBox("Invoice No. does not exist.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "empower")
                fpInvoice.MaxRows = 0
                txtOTLNo.Text = ""
                txtLorryNo.Text = ""
                txtFreight.Text = ""
                txtInvoice.Text = ""
                If Opt_yes.Checked = True Then
                    Me.txtASNNumber.Visible = False
                    Me.txtASNNumber.Enabled = False
                    Me.txtASNNumber.Text = ""
                    Me.txtASNNumber.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    Me.lblASNNumber.Visible = False
                End If
                GoTo EventExitSub
            Else
                If optExciseFormat.Checked = True Then
                    Call fillAllinvoiceInGrid()
                    fpInvoice.Focus()
                Else
                    Cmdinvoice.Focus()
                End If
                ''Added By Ashutosh on 26 Mar 2007 ,Issue Id:19661

                strsql = ""
                strsql = "Select account_code,Lorry_no,OTL_No,Frieght_Amount from saleschallan_dtl where Unit_code='" & gstrUnitId & "' and doc_no='" & Trim(txtInvoice.Text) & "' "
                If Not mresult Is Nothing Then mresult = Nothing
                mresult = New ClsResultSetDB
                mresult.GetResult(strsql)
                If mresult.GetNoRows > 0 Then
                    txtLorryNo.Text = IIf(IsDBNull(mresult.GetValue("Lorry_no")), "", mresult.GetValue("Lorry_no"))
                    txtOTLNo.Text = IIf(IsDBNull(mresult.GetValue("OTL_No")), "", mresult.GetValue("OTL_No"))
                    txtFreight.Text = IIf(IsDBNull(mresult.GetValue("Frieght_Amount")), 0, mresult.GetValue("Frieght_Amount"))
                    strAccountCode = mresult.GetValue("Frieght_Amount")
                End If

                If Opt_yes.Checked = True And AllowASNPrinting(strAccountCode) = True Then
                    Me.txtASNNumber.Visible = True
                    Me.txtASNNumber.Enabled = True
                    Me.txtASNNumber.Text = ""
                    Me.txtASNNumber.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    Me.lblASNNumber.Visible = True
                    Me.txtASNNumber.Text = CheckASNExist(Me.txtInvoice.Text)        'Get Saved ASN Number
                    Me.txtASNNumber.Focus()
                Else
                    Me.txtASNNumber.Visible = False
                    Me.txtASNNumber.Enabled = False
                    Me.txtASNNumber.Text = ""
                    Me.txtASNNumber.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    Me.lblASNNumber.Visible = False
                    Me.txtASNNumber.Text = CheckASNExist(Me.txtInvoice.Text)        'Get Saved ASN Number
                    Me.txtASNNumber.Focus()
                End If
                ''Changes for Issue Id:19661 end here.
            End If
        End If
        mresult.ResultSetClose()
        mresult = Nothing
        GoTo EventExitSub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Function CheckASNExist(ByVal pstrInvoiceNo As String) As String
        On Error GoTo ErrHandler
        Dim rsgetASNNumber As ClsResultSetDB
        Dim strsql As String
        rsgetASNNumber = New ClsResultSetDB
        strsql = "select ASN_NO from CreatedASN where Unit_code='" & gstrUnitId & "' and doc_no='" & Trim(pstrInvoiceNo) & "'"
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
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

    End Function
    Private Sub frmEXPREP0006_SOUTH_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Initialise the values on Form Activation
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo Err_Handler
        mdifrmMain.CheckFormName = mintIndex
        frmModules.NodeFontBold(Tag) = True
        Call RefreshForm()
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmEXPREP0006_SOUTH_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To release the memory on Deactivate
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo Err_Handler
        frmModules.NodeFontBold(Tag) = False
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub Form_Initialize_Renamed()
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To initialise the required data
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo Err_Handler
        mStrCustMst = "Select Doc_No,Invoice_type from SalesChallan_Dtl where unit_code='" & gstrUnitId & "' and Doc_No = "
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmEXPREP0006_SOUTH_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To write the code on form load
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo Err_Handler
        mintIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader1.Tag)
        Call FillLabelFromResFile(Me) 'To Fill label description from Resource file
        Call FitToClient(Me, fraInvoice, ctlFormHeader1, Cmdinvoice) 'To fit the form in the MDI
        Call EnableControls(False, Me) 'To Disable controls
        txtUnitCode.Enabled = True
        txtUnitCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        cmdUnitCodeList.Enabled = True
        gblnCancelUnload = False
        cmdHelp.Enabled = True
        optExciseFormat.Enabled = True
        cmdSave.Enabled = True
        cmdCancel.Enabled = True
        dtCurrentDate = GetServerDate()
        Call AddColumnsInSpread()
        txtOTLNo.Enabled = True : txtOTLNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        txtLorryNo.Enabled = True : txtLorryNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        txtFreight.Enabled = True
        txtFreight.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        Opt_yes.Checked = True
        txtFreight.Enabled = True
        mblnInvocieforMTL = Find_Value("select isnull(InvoiceForMTLSharjah,0)as InvoiceForMTLSharjah from sales_parameter where unit_code='" & gstrUnitId & "'")
        mblnInvoicelike_MTLsharjah = Find_Value("select isnull(Invoicelike_MTLsharjah,0)as Invoicelike_MTLsharjah from sales_parameter WHERE UNIT_CODE='" & gstrUnitId & "'")
        '102027599
        mblnEwaybill_Print = Find_Value("SELECT ISNULL(MAX(CAST(EWAY_BILL_FUNCTIONALITY AS INT)),0) FROM SALECONF (NOLOCK) WHERE  UNIT_CODE = '" & gstrUnitId & "' AND INVOICE_TYPE='EXP' AND DATEDIFF(DD,GETDATE(),FIN_START_DATE)<=0  AND DATEDIFF(DD,FIN_END_DATE,GETDATE())<=0 ")
        If mblnEwaybill_Print Then
            chkPrintReprint.Enabled = True
            chkPrintReprint.Checked = True
            btnExceptionInvoices.Enabled = True
        Else
            chkPrintReprint.Enabled = False
            chkPrintReprint.Checked = False
            btnExceptionInvoices.Enabled = False
        End If
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmEXPREP0006_SOUTH_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To write code on form unload
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo Err_Handler
        mdifrmMain.RemoveFormNameFromWindowList = mintIndex
        frmModules.NodeFontBold(Tag) = False
        Me.Dispose()
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Function ValuetoVariables() As Boolean
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   Store the calculated values in Variables before invoice Posting
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        Dim strsql As String
        Dim strInvoiceDate As Object
        On Error GoTo Err_Handler
        strsql = "Select Account_code,Exchange_Rate,Invoice_date,Account_code,Exchange_Rate,Invoice_date,Invoice_Type,Sub_Category from SalesChallan_Dtl where Unit_code='" & gstrUnitId & "' and Doc_No=" & Me.txtInvoice.Text & " AND Location_Code='" & Trim(txtUnitCode.Text) & "'"
        mresult = New ClsResultSetDB
        mresult.GetResult(strsql)
        mAccount_Code = mresult.GetValue("Account_Code")
        mexchange_rate = mresult.GetValue("Exchange_rate")
        mInvType = mresult.GetValue("Invoice_Type")
        mSubCat = mresult.GetValue("Sub_Category")
        strInvoiceDate = VB6.Format(mresult.GetValue("Invoice_Date"), gstrDateFormat)
        mresult.ResultSetClose()
        mresult = Nothing
        'MTL Sharjah Exchange Rate Start
        If mblnInvocieforMTL Then
            strsql = "SELECT ISNULL(CEXCH_MULTIFACTOR,0) CEXCH_MULTIFACTOR FROM GEN_CUREXCHMASTER WHERE UNIT_CODE='" & gstrUnitId & "' AND CEXCH_CURRENCYTO='AED'  AND CEXCH_INOUT=0 AND datediff(dd,'" & getDateForDB(strInvoiceDate) & "',CEXCH_DATEFROM)<=0  and datediff(dd,CEXCH_DATETO,'" & getDateForDB(strInvoiceDate) & "')<=0"
            mresult = New ClsResultSetDB
            mresult.GetResult(strsql)
            exchange_rate_aed = IIf(Convert.ToString(mresult.GetValue("CEXCH_MULTIFACTOR")).ToUpper() = "UNKNOWN", 0, mresult.GetValue("CEXCH_MULTIFACTOR"))
            mresult.ResultSetClose()
            mresult = Nothing
        End If
        'MTL Sharjah Exchange Rate End
        If Opt_yes.Checked = True Then
            strsql = "Select Current_No,updateStock_Flag,updatePO_Flag from Saleconf where Unit_code='" & gstrUnitId & "' and Invoice_Type='" & mInvType & "' and sub_type='"
            strsql = strsql & mSubCat & "' AND Location_Code='" & Trim(txtUnitCode.Text) & "'and datediff(dd,'" & getDateForDB(strInvoiceDate) & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
            mresult = New ClsResultSetDB
            mresult.GetResult(strsql)
            mInvNo = mresult.GetValue("Current_no")
            mupdatepo = mresult.GetValue("updatePO_Flag")
            mupdatestock = mresult.GetValue("updateStock_Flag")
            mInvNo = CDbl(GenerateInvoiceNo(mInvType, strInvoiceDate))
            mresult.ResultSetClose()
        Else
            mInvNo = CDbl(Trim(txtInvoice.Text))
        End If
        mresult = New ClsResultSetDB
        strsql = "select Basic = sum(sales_Quantity*(Rate * " & mexchange_rate & "))"
        strsql = strsql & "from sales_Dtl where Unit_code='" & gstrUnitId & "' and Doc_No = " & Me.txtInvoice.Text
        mresult.GetResult(strsql)
        mBasicAmt = IIf(mresult.GetValue("Basic") = "", 0, mresult.GetValue("Basic"))
        mresult.ResultSetClose()
        mresult = Nothing
        msubTotal = mExDuty + mBasicAmt
        mOtherAmt = 0
        mGrTotal = (msubTotal - mCustmtrl) + mStAmt + mFrAmt
        ValuetoVariables = False
        Exit Function
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Public Sub updatesalesconfandsaleschallan()
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   to generate the database updation string.
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo Err_Handler
        Dim Doc_No As Short
        Dim Suffix As String
        Dim Excise_Type As String
        Dim Excise_SerialNo As Short
        Dim From_Box As Short
        Dim To_Box As Short
        Dim Invoice_Date As Date
        Dim Account_Code As String
        Dim Form3 As String
        Dim Form3Date As String
        Dim Carriage_Name As String

        Dim Year_Renamed As Short
        Dim Excise_Duty_Per As Double
        Dim Extra_Excise_Duty_Per As Double
        Dim Sales_Tax As Double
        Dim Surcharge_Sales_Tax As Double
        Dim Frieght_Tax As Double
        Dim Invoice_Type As String
        Dim Ref_Doc_No As String
        Dim Cust_Name As String
        Dim SalesTax_Type As String
        Dim SalesTax_FormNo As String
        Dim SalesTax_FormValue As Double
        Dim Annex_no As Short
        Dim Ent_dt As Date
        Dim Ent_UserId As String
        Dim Transport_Type As String
        Dim Vehicle_No As String
        Dim From_Station As String
        Dim To_Station As String
        Dim mStrSQL As String
        Dim strbill_flag As String
        Dim strbuyer_description_of_goods As String
        Dim strbuyer_id As String
        Dim strctry_destination_goods As String
        Dim strcurrency_code As String
        Dim StrCust_Ref As String
        Dim dblcvd_amount As Double
        Dim dblcvd_per As Double
        Dim strdelivery_terms As String
        Dim strdispatch_mode As String
        Dim strexchange_rate As String
        Dim dtexchange_date As Date
        Dim dblExcise_Amount As Double
        Dim strfinal_destination As String
        Dim dblfreight_amount As Double
        Dim dblinsurance As Double
        Dim strinvoice_description_of_EPC As String
        Dim strInvoice_Type As String
        Dim strmode_of_shipment As String
        Dim strnature_of_contract As String
        Dim strorigin_status As String
        Dim dblpacking_amount As Double
        Dim strpayment_terms As String
        Dim strport_of_discharge As String
        Dim strport_of_loading As String
        Dim strprecarriage_by As String
        Dim strprint_date_time As String
        Dim strprint_flag As New VB6.FixedLengthString(1)
        Dim strreceipt_pre_carriage_by As String
        Dim dblsale_amount As Double
        Dim strsub_category As New VB6.FixedLengthString(1)
        Dim dblSurcharge_Sales_Amount As Double
        Dim dblSVD_amount As Double
        Dim dblSVD_per As Double
        Dim dbltotal_amount As Double
        Dim dbltotal_quantity As Double
        Dim strvessel_flight_number As String
        Dim strInvoiceDate As String
        Dim rsSalesChallan As ClsResultSetDB
        rsSalesChallan = New ClsResultSetDB
        rsSalesChallan.GetResult("SELECT * FROM  saleschallan_dtl where Unit_code='" & gstrUnitId & "' and doc_no = " & txtInvoice.Text & " and Location_code = '" & Trim(txtUnitCode.Text) & "'")
        strInvoiceDate = VB6.Format(rsSalesChallan.GetValue("Invoice_Date"), gstrDateFormat)
        '10597202
        If DataExist("SELECT TOP 1 1 FROM SALES_PARAMETER  WHERE SINGLE_INVOICE_SERIES= 1 and UNIT_CODE='" + gstrUnitId + "'") Then
            rsSalesChallan.GetResult("select single_series from saleconf where Unit_code='" & gstrUnitId & "' and Invoice_type = 'EXP' and Location_code ='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0")
            If rsSalesChallan.GetValue("Single_series") = True Then
                salesconf = "update saleconf set current_No = " & mSaleConfNo & " where UNIT_CODE IN( SELECT UNIT_CODE FROM SALES_PARAMETER WHERE SINGLE_INVOICE_SERIES=1) and datediff(dd,'" & getDateForDB(strInvoiceDate) & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0 and Single_series = 1"
            Else
                salesconf = "update saleconf set current_No = " & mSaleConfNo & " where UNIT_CODE IN( SELECT UNIT_CODE FROM SALES_PARAMETER WHERE SINGLE_INVOICE_SERIES=1) and Invoice_type = 'EXP' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
            End If
        Else
            rsSalesChallan.GetResult("select single_series from saleconf where Unit_code='" & gstrUnitId & "' and Invoice_type = 'EXP' and Location_code ='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0")
            If rsSalesChallan.GetValue("Single_series") = True Then
                salesconf = "update saleconf set current_No = " & mSaleConfNo & " where Unit_code='" & gstrUnitId & "' and Location_code ='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0 and Single_series = 1"
            Else
                salesconf = "update saleconf set current_No = " & mSaleConfNo & " where Unit_code='" & gstrUnitId & "' and Invoice_type = 'EXP' and Location_code ='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
            End If

        End If
        'rsSalesChallan.GetResult("select single_series from saleconf where Unit_code='" & gstrUNITID & "' and Invoice_type = 'EXP' and sub_type = 'E' and Location_code ='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0")
        'If rsSalesChallan.GetValue("Single_series") = True Then
        'salesconf = "update saleconf set current_No = " & mSaleConfNo & " where Unit_code='" & gstrUNITID & "' and Location_code ='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0 and Single_series = 1"
        'Else
        'salesconf = "update saleconf set current_No = " & mSaleConfNo & " where Unit_code='" & gstrUNITID & "' and Invoice_type = 'EXP' and sub_type = 'E' and Location_code ='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
        'End If
        saleschallan = "UPDATE Saleschallan_dtl SET Doc_No=" & mInvNo & ",bill_flag = 1,print_flag = 1,Lorry_No='" & Trim(txtLorryNo.Text) & "',OTL_No='" & Trim(txtOTLNo.Text) & "',Frieght_Amount=" & Val(txtFreight.Text) & " WHERE Unit_code='" & gstrUnitId & "' and Doc_No = " & Me.txtInvoice.Text & " and Invoice_type = '" & mInvType & "' and  sub_category =  '" & mSubCat & "' and Location_Code='" & Trim(txtUnitCode.Text) & "'"
        salesDtl = "UPDATE Sales_dtl SET Doc_No=" & mInvNo & " WHERE Unit_code='" & gstrUnitId & "' and Doc_No = " & Me.txtInvoice.Text & "  and Location_Code='" & Trim(txtUnitCode.Text) & "'"
        STREXPDET = "UPDATE EXPORT_SALES_EXTRA_DETAIL SET Doc_No = " & mInvNo & " WHERE Unit_code='" & gstrUnitId & "' and Doc_No = " & Me.txtInvoice.Text & "  and Unt_CodeID = '" & Trim(txtUnitCode.Text) & "'"
        mstrDespAdvice = "UPDATE BAR_DISPATCHADVICE_hdr SET INVOICENO=" & mInvNo & "  WHERE UNIT_CODE='" & gstrUnitId & "' AND INVOICENO=" & Trim(txtInvoice.Text) & " "
        mstrupdatebarpalmst = "UPDATE bar_palette_mst SET INVOICE_NO=" & mInvNo & "  WHERE UNIT_CODE='" & gstrUnitId & "' AND INVOICE_NO=" & Trim(txtInvoice.Text) & " "
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub updateLorryNo()
        On Error GoTo Err_Handler
        mSalesChallanLorryNo = "UPDATE Saleschallan_dtl SET Lorry_No='" & Trim(txtLorryNo.Text) & "',OTL_No='" & Trim(txtOTLNo.Text) & "',Frieght_Amount=" & Val(txtFreight.Text) & " WHERE Unit_code='" & gstrUnitId & "' and Doc_No = " & Me.txtInvoice.Text & " and Invoice_type = '" & mInvType & "' and  sub_category =  '" & mSubCat & "' and Location_Code='" & Trim(txtUnitCode.Text) & "'"
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Sub UpdateinSale_Dtl()
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To generate database update string in Database
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        Dim rssaledtl As ClsResultSetDB
        Dim rsSaleConf As ClsResultSetDB
        Dim strsql As String
        Dim strStockLocCode As String
        Dim intRow, intLoopcount As Short
        Dim mItem_Code, mSuffix, mCust_Item_Code As String
        Dim mCust_Item_Desc As String
        Dim mSales_Quantity As Double
        Dim strSQLSales As String
        Dim rsSalesChallan As ClsResultSetDB
        Dim strInvoiceDate As String
        Dim strcustomercode As String

        On Error GoTo Err_Handler
        rsSaleConf = New ClsResultSetDB
        rsSalesChallan = New ClsResultSetDB
        rsSalesChallan.GetResult("SELECT * FROM  saleschallan_dtl where Unit_code='" & gstrUnitId & "' and doc_no = " & txtInvoice.Text & " and Location_code = '" & Trim(txtUnitCode.Text) & "'")
        strInvoiceDate = VB6.Format(rsSalesChallan.GetValue("Invoice_Date"), gstrDateFormat)
        strSQLSales = "Select Stock_Location from saleconf where UNIT_code='" & gstrUnitId & "' and Invoice_Type = 'EXP' and Sub_type ='E' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0"
        Call rsSaleConf.GetResult(strSQLSales)

        strStockLocCode = rsSaleConf.GetValue("Stock_Location")
        strcustomercode = Find_Value("SELECT account_code FROM  saleschallan_dtl where Unit_code='" & gstrUNITID & "' and doc_no = " & txtInvoice.Text & " and Location_code = '" & Trim(txtUnitCode.Text) & "'")
        If DataExist("SELECT SOUPLD_LINE_LEVEL_SALESORDER FROM CUSTOMER_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & strcustomercode & "'") = True Then
            strsql = "Select a.suffix,a.item_code,a.Cust_Item_Code,a.Cust_Item_Desc,a.Sales_Quantity,a.external_salesorder_no as cust_ref,'' as amendment_no from sales_Dtl a, saleschallan_dtl b  where a.Unit_Code=b.Unit_Code and a.Unit_code='" & gstrUNITID & "' and a.Doc_No = " & Me.txtInvoice.Text & " and a.Location_Code='" & Trim(txtUnitCode.Text) & "' and a.doc_no=b.doc_no and a.location_code = b.location_code "
        Else
            strsql = "Select a.suffix,a.item_code,a.Cust_Item_Code,a.Cust_Item_Desc,a.Sales_Quantity,b.cust_ref,b.amendment_no from sales_Dtl a, saleschallan_dtl b  where a.Unit_Code=b.Unit_Code and a.Unit_code='" & gstrUNITID & "' and a.Doc_No = " & Me.txtInvoice.Text & " and a.Location_Code='" & Trim(txtUnitCode.Text) & "' and a.doc_no=b.doc_no and a.location_code = b.location_code "
        End If

        rssaledtl = New ClsResultSetDB
        Call rssaledtl.GetResult(strsql)
        If rssaledtl.GetNoRows > 0 Then
            intRow = rssaledtl.GetNoRows
            rssaledtl.MoveFirst()
            For intLoopcount = 1 To intRow
                mSuffix = rssaledtl.GetValue("Suffix")
                mItem_Code = rssaledtl.GetValue("Item_Code")
                mCust_Item_Code = rssaledtl.GetValue("Cust_Item_Code")
                mCust_Item_Desc = rssaledtl.GetValue("Cust_Item_Desc")
                mSales_Quantity = rssaledtl.GetValue("Sales_Quantity")
                mCust_Ref = rssaledtl.GetValue("cust_ref")
                mAmendment_No = rssaledtl.GetValue("amendment_no")
                strupdateitbalmst = Trim(strupdateitbalmst) & "Update Itembal_mst set cur_bal= cur_bal-"
                strupdateitbalmst = strupdateitbalmst & mSales_Quantity & " where Unit_code='" & gstrUnitId & "' and Location_code = '" & strStockLocCode
                strupdateitbalmst = strupdateitbalmst & "' and item_code = '" & mItem_Code & "' "

                strupdatecustodtdtl = Trim(strupdatecustodtdtl) & "Update Cust_ord_dtl set Despatch_Qty = Despatch_Qty + "
                strupdatecustodtdtl = strupdatecustodtdtl & mSales_Quantity & " where Unit_code='" & gstrUnitId & "' and Account_code ='"
                strupdatecustodtdtl = strupdatecustodtdtl & mAccount_Code & "' and Cust_DrgNo = '"

                If Not DataExist("SELECT SOUPLD_LINE_LEVEL_SALESORDER FROM CUSTOMER_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & strcustomercode & "'") = True Then
                    strupdatecustodtdtl = strupdatecustodtdtl & mCust_Item_Code & "' and Cust_ref = '" & mCust_Ref & "'"
                    strupdatecustodtdtl = strupdatecustodtdtl & " and amendment_no = '" & mAmendment_No & "' "
                Else
                    If mCust_Ref = "" Or mCust_Ref = "0" Then
                        Dim strCustRef = SqlConnectionclass.ExecuteScalar("select Cust_ref from saleschallan_dtl where unit_code='" & gstrUNITID & "'  and doc_no=" & txtInvoice.Text & " ")
                        Dim strAmend = SqlConnectionclass.ExecuteScalar("select amendment_no from saleschallan_dtl where unit_code='" & gstrUNITID & "'  and doc_no=" & txtInvoice.Text & " ")
                        strupdatecustodtdtl = strupdatecustodtdtl & mCust_Item_Code & "' and Cust_ref = '" & strCustRef & "' and amendment_no = '" & strAmend & "'"
                    Else
                        strupdatecustodtdtl = strupdatecustodtdtl & mCust_Item_Code & "' and EXTERNAL_SALESORDER_NO = '" & mCust_Ref & "'"
                    End If

                End If
                rssaledtl.MoveNext()
            Next intLoopcount
        End If
        rssaledtl.ResultSetClose()

        rssaledtl = Nothing
        rsSaleConf.ResultSetClose()

        rsSaleConf = Nothing
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Sub ValidSelection()
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   to check required data before PREVIEW OR Print
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        Dim blnInvalidData As Boolean
        Dim strErrMsg As String
        Dim ctlBlank As System.Windows.Forms.Control
        Dim lNo As Integer
        On Error GoTo Err_Handler
        ValidRecord = False
        lNo = 1
        strErrMsg = ResolveResString(10059) & vbCrLf & vbCrLf
        If Len(Trim(txtInvoice.Text)) = 0 Then
            blnInvalidData = True
            strErrMsg = strErrMsg & vbCrLf & lNo & "." & ResolveResString(60373)
            lNo = lNo + 1
            If ctlBlank Is Nothing Then ctlBlank = txtInvoice
        End If
        strErrMsg = VB.Left(strErrMsg, Len(strErrMsg) - 1)
        strErrMsg = strErrMsg & "."
        lNo = lNo + 1
        If blnInvalidData = True Then
            gblnCancelUnload = True
            Call MsgBox(strErrMsg, MsgBoxStyle.Information + MsgBoxStyle.SystemModal, "empower")
            ctlBlank.Focus()
            Exit Sub
        End If
        ValidRecord = True
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Sub RefreshForm()
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To refresh all values
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        cmdHelp.Enabled = True
        Me.Fra_yes_no.Enabled = True
        Me.Opt_no.Enabled = True
        Me.Opt_yes.Enabled = True
        Me.txtInvoice.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        Me.txtInvoice.Enabled = True
        Me.txtInvoice.Text = ""
        Me.Opt_yes.Focus()
    End Sub
    Public Sub InitializeVariable()
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To initialise the values of variables
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        mInvType = ""
        mSubCat = ""
        mInvNo = 0
        mBasicAmt = 0
        mExDuty = 0
        msubTotal = 0
        mAccount_Code = ""
        mGrTotal = 0
        strsalesheadersql = ""
        strsaledetails = ""
        strSQLDuePayment = ""
        strupdate = ""
        StrDeletesaledtl = ""
        strupdateitbalmst = ""
        strupdatecustodtdtl = ""
        saleschallan = ""
        strDeletesalesconf = ""
        salesconf = ""
        salesDtl = ""
        STREXPDET = ""
    End Sub
    Private Sub frmEXPREP0006_SOUTH_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   At Enter Key Press Set Focus To Next Control
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo Err_Handler
        If KeyAscii = 13 Then
            System.Windows.Forms.SendKeys.SendWait("{Tab}") : GoTo EventExitSub
        End If
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Function CreateStringForAccounts() As Boolean
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To generate COM string for Accounts
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        Dim objRecordSet As New ADODB.Recordset
        Dim objTmpRecordset As New ADODB.Recordset
        Dim strRetval As String
        Dim strInvoiceNo As String
        Dim strInvoiceDate As String
        Dim strCurrencyCode As String
        Dim dblInvoiceAmt As Double
        Dim dblExchangeRate As Double
        Dim dblBasicAmount As Double
        Dim dblBaseCurrencyAmount As Double
        Dim strCreditTermsID As String
        Dim strBasicDueDate As String
        Dim strPaymentDueDate As String
        Dim strExpectedDueDate As String
        Dim strCustomerGL As String
        Dim strCustomerSL As String
        Dim strItemGL As String
        Dim strItemSL As String
        Dim strGlGroupId As String
        Dim varTmp As Object
        Dim iCtr As Short
        Dim strCustCode As String
        Dim dblCCShare As Double
        Dim dblFreightAmt As Double
        Dim strTaxGL As String
        Dim strTaxSL As String
        Dim strTaxType As String
        Dim dblTaxAmt As Double
        Dim dblTaxRate As Double
        Dim dblsalestaxamount As Double
        Dim strsalestaxtype As String

        On Error GoTo ErrHandler
        If mblnInvocieforMTL = True Or mblnInvoicelike_MTLsharjah = True Then
            objRecordSet.Open("SELECT * FROM  saleschallan_dtl WHERE Doc_No='" & Trim(txtInvoice.Text) & "' and Location_Code='" & Trim(txtUnitCode.Text) & "'and unit_code='" & gstrUnitId & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
        Else
            objRecordSet.Open("SELECT * FROM  saleschallan_dtl a,cust_ord_hdr b WHERE a.Unit_code = b.Unit_Code and a.Unit_code='" & gstrUnitId & "' and a.cust_ref = b.cust_ref and a.amendment_no = b.amendment_no and a.Doc_No='" & Trim(txtInvoice.Text) & "' AND a.Location_Code='" & Trim(txtUnitCode.Text) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
        End If
        If objRecordSet.EOF Then
            MsgBox("Invoice details not found", MsgBoxStyle.Information, "empower")
            CreateStringForAccounts = False
            If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                objRecordSet.Close()
                objRecordSet = Nothing
            End If
            Exit Function
        End If
        strInvoiceNo = CStr(mInvNo)
        'strInvoiceDate = VB6.Format(objRecordSet.Fields("Invoice_Date").Value, "dd-MMM-yyyy")
        '11 Jun 18 --Changes for same date locking ádded by priti on 08 Jan 2020
        If DataExist("SELECT TOP 1 1 FROM SALES_PARAMETER WHERE INVOICE_LOCKING_ENTRY_SAMEDATE=1  and UNIT_CODE = '" & gstrUNITID & "'") And (GetPlantName() = "HILEX" Or GetPlantName() = "MTL") Then
            strInvoiceDate = VB6.Format(GetServerDateTime(), "dd-MMM-yyyy")
        Else
            strInvoiceDate = VB6.Format(objRecordSet.Fields("Invoice_Date").Value, "dd-MMM-yyyy")
        End If
        strCurrencyCode = Trim(IIf(IsDBNull(objRecordSet.Fields("Currency_Code").Value), "", objRecordSet.Fields("Currency_Code").Value))
        dblInvoiceAmt = IIf(IsDBNull(objRecordSet.Fields("total_amount").Value), 0, objRecordSet.Fields("total_amount").Value)
        dblExchangeRate = IIf(IsDBNull(objRecordSet.Fields("Exchange_Rate").Value), 1, objRecordSet.Fields("Exchange_Rate").Value)
        dblFreightAmt = objRecordSet.Fields("frieght_amount").Value
        dblsalestaxamount = objRecordSet.Fields("sales_tax_amount").Value
        strsalestaxtype = objRecordSet.Fields("salestax_type").Value.ToString

        strCustCode = " "
        strCustCode = mAccount_Code
        '10895403
        strCreditTermsID = Trim(IIf(IsDBNull(objRecordSet.Fields("payment_terms").Value), "", objRecordSet.Fields("payment_terms").Value))
        '10895403
        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
        objTmpRecordset.Open("SELECT Cst_ArCode, Cst_slCode, Cst_CreditTerm FROM Sal_CustomerMaster where Unit_code='" & gstrUnitId & "' and Prty_PartyID='" & strCustCode & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
        If objTmpRecordset.EOF Then
            MsgBox("Customer details not found", MsgBoxStyle.Information, "empower")
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
        strCustomerGL = Trim(IIf(IsDBNull(objTmpRecordset.Fields("Cst_ArCode").Value), "", objTmpRecordset.Fields("Cst_ArCode").Value))
        strCustomerSL = Trim(IIf(IsDBNull(objTmpRecordset.Fields("Cst_slCode").Value), "", objTmpRecordset.Fields("Cst_slCode").Value))
        '10895403
        If strCreditTermsID = "" Then
            strCreditTermsID = Trim(IIf(IsDBNull(objTmpRecordset.Fields("Cst_CreditTerm").Value), "", objTmpRecordset.Fields("Cst_CreditTerm").Value))
        End If
        '10895403
        If strCreditTermsID = "" Then
            MsgBox("Customer Credit Terms not found", MsgBoxStyle.Information, "empower")
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

        Dim objCreditTerms As New prj_CreditTerm.clsCR_Term_Resolver
        strRetval = objCreditTerms.RetCR_Term_Dates("", "INV", strCreditTermsID, strInvoiceDate, gstrUnitId, "", "", gstrCONNECTIONSTRING)
        If CheckString(strRetval) = "Y" Then
            strRetval = Mid(strRetval, 3)
            varTmp = Split(strRetval, "»")
            strBasicDueDate = VB6.Format(varTmp(0), "dd-MMM-yyyy")
            strPaymentDueDate = VB6.Format(varTmp(1), "dd-MMM-yyyy")
            strExpectedDueDate = VB6.Format(varTmp(1), "dd-MMM-yyyy")
        Else
            MsgBox(CheckString(strRetval), MsgBoxStyle.Information, "empower")
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


        mstrMasterString = ""
        mstrDetailString = ""

        mstrMasterString = "I»" & strInvoiceNo & "»Dr»»" & strInvoiceDate & "»»»»»EXP»I»" & strInvoiceNo & "»" & strInvoiceDate & "»" & Trim(strCustCode) & "»" & Trim(txtUnitCode.Text) & "»" & strCurrencyCode & "»" & dblInvoiceAmt & "»" & dblInvoiceAmt * dblExchangeRate & "»" & dblExchangeRate & "»" & strCreditTermsID & "»" & strBasicDueDate & "»" & strPaymentDueDate & "»" & strExpectedDueDate & "»" & strCustomerGL & "»" & strCustomerSL & "»" & mP_User & "»getdate()»»"
        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close()
        If mblnInvocieforMTL = True Or mblnInvoicelike_MTLsharjah = True Then
            objRecordSet.Open("SELECT isnull(sum(a.basic_amount),0) as Basic_Amount,a.item_code, b.GlGrp_code" & _
             " FROM sales_dtl a, item_mst b WHERE a.unit_code=b.unit_code and a.Doc_No='" & Trim(txtInvoice.Text) & "' and a.Item_Code=b.Item_Code and a.Location_Code='" & Trim(txtUnitCode.Text) & "'" & _
             " and a.unit_code='" & gstrUnitId & "' group by a.item_code,b.GlGrp_code")
        Else

            objRecordSet.Open("SELECT sales_dtl.*, item_mst.GlGrp_code FROM sales_dtl, item_mst WHERE sales_dtl.Unit_code = item_mst.Unit_code and sales_dtl.Unit_code='" & gstrUnitId & "' and sales_dtl.Doc_No='" & Trim(txtInvoice.Text) & "' and sales_dtl.Item_Code=item_mst.Item_Code AND sales_dtl.Location_Code='" & Trim(txtUnitCode.Text) & "'")
        End If
        If objRecordSet.EOF Then
            MsgBox("Item details not found.", MsgBoxStyle.Information, "empower")
            CreateStringForAccounts = False
            If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then
                objRecordSet.Close()
                objRecordSet = Nothing
            End If
            Exit Function
        End If

        While Not objRecordSet.EOF
            strGlGroupId = Trim(IIf(IsDBNull(objRecordSet.Fields("GlGrp_code").Value), "", objRecordSet.Fields("GlGrp_code").Value))
            dblBasicAmount = IIf(IsDBNull(objRecordSet.Fields("Basic_Amount").Value), 0, objRecordSet.Fields("Basic_Amount").Value)
            dblBaseCurrencyAmount = dblBasicAmount
            If dblBaseCurrencyAmount > 0 Then
                strRetval = GetItemGLSL(strGlGroupId, "Export_Sales")
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
                objTmpRecordset.Open("SELECT Location_Code,Invoice_Type,Sub_Type,ccM_ccCode,ccM_cc_percentage FROM invcc_dtl WHERE Unit_code='" & gstrUnitId & "' and Invoice_Type='EXP' AND Sub_Type = 'E' AND Location_Code ='" & Trim(txtUnitCode.Text) & "' AND ccM_cc_Percentage > 0", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
                If Not objTmpRecordset.EOF Then
                    While Not objTmpRecordset.EOF
                        dblCCShare = (dblBaseCurrencyAmount / 100) * objTmpRecordset.Fields("ccM_cc_Percentage").Value
                        mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»ITM»EXP»" & iCtr & "»" & Trim(objRecordSet.Fields("item_code").Value) & "»" & strGlGroupId & "»0»" & strItemGL & "»" & strItemSL & "»" & dblCCShare & "»Cr»»" & Trim(objTmpRecordset.Fields("ccM_ccCode").Value) & "»»»»0»0»0»0»0¦"
                        objTmpRecordset.MoveNext()
                        iCtr = iCtr + 1
                    End While
                Else
                    mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»ITM»EXP»" & iCtr & "»" & Trim(objRecordSet.Fields("item_code").Value) & "»" & strGlGroupId & "»0»" & strItemGL & "»" & strItemSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                End If
                iCtr = iCtr + 1
            End If
            objRecordSet.MoveNext()
        End While
        If dblFreightAmt > 0 And GetPlantName() = "WCS" Then
            'initializing the tax gl and sl here
            strRetval = GetTaxGlSl("FRTIM")
            If strRetval = "N" Then
                MsgBox("GL for ARTAX is not defined for FRTIM  ", MsgBoxStyle.Information, "eMPro")
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

            mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»FRT»0»" & "»»0»" & strTaxGL & "»" & strTaxSL & "»" & dblFreightAmt & "»Cr»»»»»»0»0»0»0»0" & "¦"
        End If

        'Sale Tax Changes'
        If dblsalestaxamount > 0 And mblnInvocieforMTL = True Then
            'initializing the tax gl and sl here
            strRetval = GetTaxGlSl("VAT")
            If strRetval = "N" Then
                MsgBox("GL for ARTAX is not defined for VAT", MsgBoxStyle.Information, "eMPro")
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

            mstrDetailString = mstrDetailString & "I»" & strInvoiceNo & "»" & iCtr & "»TAX»VAT»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblsalestaxamount & "»Cr»»»»»»0»0»0»0»0" & "¦"
        End If

        'Sale Tax Changes Ended'
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
    Private Function GetItemGLSL(ByVal InventoryGlGroup As String, ByVal PurposeCode As String) As String
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :   1.InventoryGlGroup in string format
        '                       2.PurposeCode in String Format
        'Return Value       :   retuns GL/SL as String
        'Function           :   To Get GL & SL of item code in Invoice
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        Dim objRecordSet As New ADODB.Recordset
        Dim strGL As String
        Dim strSL As String
        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close()
        objRecordSet.Open("SELECT invGld_glcode, invGld_slcode FROM fin_InvGLGrpDtl WHERE Unit_code='" & gstrUnitId & "' and invGld_prpsCode = '" & PurposeCode & "' AND invGld_invGLGrpId = '" & InventoryGlGroup & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
        If objRecordSet.EOF Then
            objRecordSet.Close()
            objRecordSet.Open("SELECT gbl_glCode, gbl_slCode FROM fin_globalGL WHERE Unit_code='" & gstrUnitId & "' and gbl_prpsCode = '" & PurposeCode & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly)
            If objRecordSet.EOF Then
                GetItemGLSL = "N"
                MsgBox("GL and SL not defined for Purpose Code: " & PurposeCode, MsgBoxStyle.Information, "empower")
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
            MsgBox("GL and SL not defined for Purpose Code:" & PurposeCode, MsgBoxStyle.Information, "empower")
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
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        varHelp = ctlHelp.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, pstrQuery)
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        If UBound(varHelp) <> -1 Then

            If varHelp(0) <> "0" Then
                pctlCode.Text = Trim(varHelp(0))
                If Not (pctlDesc Is Nothing) Then
                    pctlDesc.Text = Trim(varHelp(1))
                End If
                pctlCode.Focus()
            Else
                MsgBox("No Record Available", MsgBoxStyle.Information, "empower")
            End If
        End If
        Exit Sub 'This is to avoid the execution of the error handler

ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Function GenerateInvoiceNo(ByVal pstrInvoiceType As String, ByVal pstrRequiredDate As String) As String
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To generat invoice no from databse
        'Comments           :   NA
        'Creation Date      :   15/05/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim clsInstEMPDBDbase As New EMPDataBase.EMPDB(gstrUnitId)
        Dim strCheckDOcNo As String 'Gets the Doc Number from Back End
        Dim strTempSeries As String 'Find the Numeric series in Doc No
        Dim strSuffix As String 'Generate a NEW Series
        Dim strZeroSuffix As String
        Dim strFin_Start_Date As String
        Dim strFin_End_Date As String
        Dim strsql As String 'String SQL Query
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        If Len(Trim(pstrInvoiceType)) > 0 Then 'For Dated Docs
            strsql = "Set dateformat 'dmy' Select Current_No,Suffix,Fin_start_date,Fin_end_Date From saleConf Where  Unit_code='" & gstrUnitId & "' and"
            strsql = strsql & " Invoice_Type ='" & pstrInvoiceType & "' AND Location_Code ='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(pstrRequiredDate) & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(pstrRequiredDate) & "')<=0"
            With clsInstEMPDBDbase.CConnection
                .OpenConnection(gstrDSNName, gstrDatabaseName)
                .ExecuteSQL("Set Dateformat 'dmy'")
            End With
            clsInstEMPDBDbase.CRecordset.OpenRecordset(strsql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
            If Not clsInstEMPDBDbase.CRecordset.EOF_Renamed Then
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
        Dim clsErrorInst As New EMPDataBase.EMPDB(gstrUnitId)
        clsErrorInst.CError.RaiseError(20008, "[frmexptrn0006]", "[GenerateInvoiceNo]", "", "No. Not Generated For DocType = " & pstrInvoiceType & " due to [ " & Err.Description & " ].", My.Application.Info.DirectoryPath, gstrDSNName, gstrDatabaseName)
    End Function

    Sub AddColumnsInSpread()
        On Error GoTo ErrHandler
        With fpInvoice
            .MaxRows = 0
            .MaxCols = 8
            .Row = 0
            .Col = 1 : .Text = "Mark" : .set_ColWidth(1, 4)
            .Col = 2 : .Text = "Invoice No" : .set_ColWidth(2, 10)
            .Col = 3 : .Text = "FreshCr Date" : .set_ColWidth(3, 10)
            .Col = 4 : .Text = "Excise Per" : .set_ColWidth(4, 0)
            .Col = 5 : .Text = "Excise Amount" : .set_ColWidth(5, 10)
            .Col = 6 : .Text = "Cess Per" : .set_ColWidth(6, 0)
            .Col = 7 : .Text = "Cess Amount" : .set_ColWidth(7, 10)
            .Col = 8 : .Text = "Remarks" : .set_ColWidth(8, 20)
        End With
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Sub fillAllinvoiceInGrid()
        On Error GoTo ErrHandler
        Dim rsInvoice As New ADODB.Recordset
        Dim rsRate As New ClsResultSetDB
        Dim strsql As String
        Dim intCount As Short
        Dim lngTotalRecord As Integer
        Dim dblExciseAmt, dblCessAmt As Double
        Dim blnCheckPrintExciseFormat As Boolean

        strsql = "select PrintExciseFormat_excisePer, PrintExciseFormat_addDuty, PrintExciseFormat_cessOnCVD from sales_parameter where unit_code='" & gstrUnitId & "'"
        rsRate.GetResult(strsql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If Not rsRate.EOFRecord Then
            sngExcisePer = rsRate.GetValue("PrintExciseFormat_excisePer")
            sngAdditionalDuty = rsRate.GetValue("PrintExciseFormat_addDuty")
            sngCessOnCVD = rsRate.GetValue("PrintExciseFormat_cessOnCVD")
        End If
        rsRate.ResultSetClose()
        blnCheckPrintExciseFormat = CBool(Find_Value("select printExciseFormat from SalesChallan_dtl where Unit_code='" & gstrUnitId & "' and doc_no = '" & Trim(txtInvoice.Text) & "'"))

        If blnCheckPrintExciseFormat Then
            strsql = "select FreshCrInvoice_no,FreshCrRecdDate, Excise_per, Excise_Amount, Cess_per, Cess_Amount, Remarks from printExciseFormat_dtl "
            strsql = strsql & " where unit_code='" & gstrUnitId & "' and invoice_no ='" & Trim(txtInvoice.Text) & "'"
            rsInvoice.Open(strsql, mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            lngTotalRecord = rsInvoice.RecordCount
            If lngTotalRecord = 0 Then
                rsInvoice.Close()
                rsInvoice = Nothing
                cmdSave.Enabled = False : cmdCancel.Enabled = False : cmdClose.Enabled = False
                Cmdinvoice.Visible = True
                fpInvoice.MaxRows = 0
                Exit Sub
            End If
            rsInvoice.MoveFirst()
            fpInvoice.MaxRows = 0
            For intCount = 1 To lngTotalRecord
                With fpInvoice
                    .UnitType = FPSpreadADO.UnitTypeConstants.UnitTypeTwips
                    .set_RowHeight(intCount, 315)
                    .MaxRows = .MaxRows + 1
                    .Row = intCount
                    .Col = 1 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox : .TypeCheckCenter = True : .set_ColWidth(1, 0)
                    .Col = 2 : .Text = rsInvoice.Fields("FreshCrInvoice_no").Value : .Lock = True
                    .Col = 3 : .Text = VB6.Format(rsInvoice.Fields("FreshCrRecdDate").Value, gstrDateFormat) : .CellType = FPSpreadADO.CellTypeConstants.CellTypeDate : .Lock = True
                    .Col = 4 : .Text = rsInvoice.Fields("Excise_per").Value : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .Lock = True
                    dblExciseAmt = rsInvoice.Fields("Excise_Amount").Value
                    .Col = 5 : .Text = CStr(dblExciseAmt) : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .Lock = True
                    .Col = 6 : .Text = rsInvoice.Fields("cess_per").Value : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .Lock = True
                    dblCessAmt = rsInvoice.Fields("cess_Amount").Value
                    .Col = 7 : .Text = CStr(dblCessAmt) : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .Lock = True
                    .Col = 8 : .Text = rsInvoice.Fields("Remarks").Value : .Lock = True
                End With
                rsInvoice.MoveNext() 'move to next record
            Next intCount
        Else
            strsql = "select m.doc_no, sum(accessible_amount) as accessible_amount, m.exchange_rate from salesChallan_dtl m"
            strsql = strsql & " inner join sales_dtl d on M.unit_code=d.Unit_code and m.location_code = d.location_code and m.doc_no = d.doc_no"
            strsql = strsql & " Where M.unit_code='" & gstrUnitId & "' and invoice_type='EXP' and sub_category='E' and  bill_flag = 1 And FreshCrRecd = 0 and m.doc_no <> '" & Trim(txtInvoice.Text) & "'"
            strsql = strsql & " group by m.doc_no, m.exchange_rate"
            rsInvoice.Open(strsql, mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            lngTotalRecord = rsInvoice.RecordCount
            rsInvoice.MoveFirst()
            If lngTotalRecord = 0 Then
                rsInvoice.Close()
                rsInvoice = Nothing : Exit Sub
            End If
            fpInvoice.MaxRows = 0
            For intCount = 1 To lngTotalRecord
                With fpInvoice
                    .UnitType = FPSpreadADO.UnitTypeConstants.UnitTypeTwips
                    .set_RowHeight(intCount, 315)
                    .MaxRows = .MaxRows + 1
                    .Row = intCount
                    .Col = 1 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox : .TypeCheckCenter = True : .set_ColWidth(1, 500)
                    .Col = 2 : .Text = rsInvoice.Fields("doc_no").Value : .Lock = True
                    .Col = 3 : .Text = VB6.Format(dtCurrentDate, gstrDateFormat) : .CellType = FPSpreadADO.CellTypeConstants.CellTypeDate : .Lock = True
                    .Col = 4 : .Text = CStr(sngExcisePer) : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .Lock = True
                    dblExciseAmt = CalculateExciseAmount(rsInvoice.Fields("accessible_amount").Value, rsInvoice.Fields("exchange_rate").Value)
                    .Col = 5 : .Text = CStr(dblExciseAmt) : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .Lock = True
                    .Col = 6 : .Text = CStr(sngCessOnCVD) : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .Lock = True
                    dblCessAmt = CalculateCessAmount(dblExciseAmt)
                    .Col = 7 : .Text = CStr(dblCessAmt) : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .Lock = True
                    .Col = 8 : .Text = "" : .Lock = False
                End With
                rsInvoice.MoveNext() 'move to next record
            Next intCount
            cmdSave.Enabled = True : cmdCancel.Enabled = True : cmdClose.Enabled = True
            Cmdinvoice.Visible = False
        End If
        If rsInvoice.State = ADODB.ObjectStateEnum.adStateOpen Then
            rsInvoice.Close()
            rsInvoice = Nothing
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Function CalculateExciseAmount(ByRef pAccessibleAmount As Double, ByRef pExchangeRate As Double) As Object
        '*******************************************************************************
        'Author             :   Arshad Ali
        'Return Value       :   Calculated total excise amount or total duty
        'Function           :   CalculateExciseAmount
        'Creation Date      :   22/09/2004
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim strsql As String
        Dim dblAdditionalDuty, dblExciseTotal, dblCessOnCVD As Double
        If sngExcisePer > 0 Then
            dblExciseTotal = System.Math.Round((pAccessibleAmount * pExchangeRate) / 100 * sngExcisePer, 4)
        Else
            dblExciseTotal = 0
        End If

        If sngAdditionalDuty > 0 Then
            dblAdditionalDuty = (pAccessibleAmount * pExchangeRate) + dblExciseTotal
            dblAdditionalDuty = System.Math.Round(dblAdditionalDuty / 100 * sngAdditionalDuty, 4)
        Else
            dblAdditionalDuty = 0
        End If
        If sngCessOnCVD > 0 Then
            dblCessOnCVD = System.Math.Round(dblAdditionalDuty / 100 * sngCessOnCVD, 4)
        Else
            dblCessOnCVD = 0
        End If
        CalculateExciseAmount = dblExciseTotal + dblAdditionalDuty + dblCessOnCVD
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

    Function CalculateCessAmount(ByRef pExciseAmount As Double) As Object
        '*******************************************************************************
        'Author             :   Arshad Ali
        'Return Value       :   Calculated Cess Amount on Total Duty
        'Function           :   CalculateCessAmount
        'Creation Date      :   22/09/2004
        '*******************************************************************************
        On Error GoTo ErrHandler

        Dim rsRate As New ClsResultSetDB
        Dim strsql As String
        Dim sngCessPer As Double
        strsql = "select PrintExciseFormat_cessPer from sales_parameter where unit_code='" & gstrUnitId & "'"
        rsRate.GetResult(strsql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

        If Not rsRate.EOFRecord Then
            sngCessPer = rsRate.GetValue("PrintExciseFormat_cessPer")
        End If

        If sngCessPer > 0 Then
            CalculateCessAmount = System.Math.Round((pExciseAmount / 100) * sngCessPer, 4)
        Else
            CalculateCessAmount = 0
        End If

        rsRate.ResultSetClose()
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

    Sub SaveData()
        '*******************************************************************************
        'Author             :   Arshad Ali
        'Return Value       :   N/A
        'Function           :   Saves all Fresh Credit Received, update bond17openingbalance
        'Creation Date      :   22/09/2004
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim intCount As Short
        Dim blnSuccess As Boolean
        Dim strsql As String
        Dim rsRate As New ClsResultSetDB
        Dim strInvoice As String
        Dim strRemarks As String
        Dim dblFreshCrRecdExcise As Double
        Dim dblPrintExcise As Double
        Dim dtFreshCrRecdDate As Date
        Dim dblAdditionalDuty, dblExciseTotal, dblCessOnCVD As Double
        Dim dblNewBond17OpeningBalance As Double

        With fpInvoice
            blnSuccess = False
            mP_Connection.BeginTrans()
            For intCount = 1 To .MaxRows
                .Col = 1
                .Row = intCount
                If .Value = CStr(System.Windows.Forms.CheckState.Checked) Then
                    .Col = 2
                    strInvoice = .Text
                    .Col = 3
                    dtFreshCrRecdDate = ConvertToDate(.Text)
                    .Col = 5
                    dblExciseTotal = CDbl(.Text)
                    dblFreshCrRecdExcise = dblFreshCrRecdExcise + dblExciseTotal
                    .Col = 7
                    dblCessOnCVD = CDbl(.Text)
                    dblFreshCrRecdExcise = dblFreshCrRecdExcise + dblCessOnCVD
                    .Col = 8
                    strRemarks = .Text

                    strsql = "Insert into PrintExciseFormat_dtl(Unit_code,location_code, invoice_no, FreshCrRecdDate, FreshCrInvoice_no, Excise_per, Excise_Amount, Cess_per, Cess_Amount, Remarks)"
                    strsql = strsql & " values('" & gstrUnitId & "','" & Trim(txtUnitCode.Text) & "','" & Trim(txtInvoice.Text) & "','" & getDateForDB(dtFreshCrRecdDate) & "','" & strInvoice & "'," & sngExcisePer & "," & dblExciseTotal & "," & sngCessOnCVD & "," & dblCessOnCVD & ",'" & strRemarks & "')"

                    strsql = strsql & " Update SalesChallan_dtl set FreshCrRecd=1"
                    strsql = strsql & " Where unit_code='" & gstrUnitId & "' and location_code='" & Trim(txtUnitCode.Text) & "' and doc_no='" & strInvoice & "'"

                    mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

                End If
            Next intCount
            strsql = "Update SalesChallan_dtl set printExciseFormat=1"
            strsql = strsql & " Where unit_code='" & gstrUnitId & "' and location_code='" & Trim(txtUnitCode.Text) & "' and doc_no='" & Trim(txtInvoice.Text) & "'"
            mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

            dblNewBond17OpeningBalance = System.Math.Round(CheckBond17OpeningBalance((txtInvoice.Text), dblFreshCrRecdExcise), 2)
            If dblNewBond17OpeningBalance = 0 Then
                MsgBox("No suffcient Bond 17 Opening Balance.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "empower")
                mP_Connection.RollbackTrans()
                Exit Sub
            Else
                strsql = "Update SaleConf set bond17OpeningBal= " & dblNewBond17OpeningBalance
                strsql = strsql & "where unit_code='" & gstrUnitId & "' and Invoice_Type ='EXP' AND Location_Code ='" & Trim(txtUnitCode.Text) & "' and datediff(dd,getdate(),fin_start_date)<=0  and datediff(dd,fin_end_date,getdate())<=0"
                mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End If

            mP_Connection.CommitTrans()
            blnSuccess = True
            MsgBox("Record Saved Successfully.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, "empower")
            Call fillAllinvoiceInGrid()
        End With
        Exit Sub
ErrHandler:
        If Not blnSuccess Then mP_Connection.RollbackTrans()
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Function CheckBond17OpeningBalance(ByRef pstrInvoice As String, ByRef dblTotalExcise As Double) As Double
        '*******************************************************************************
        'Author             :   Arshad Ali
        'Return Value       :   Returns Bond17OpeningBalance
        'Function           :   CheckBond17OpeningBalance
        'Creation Date      :   22/09/2004
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim strsql As String
        Dim rsInvoice As New ADODB.Recordset
        Dim dblPrintExciseFormatExcise, dblOpeningBal, dblFreshCrRecd As Double

        strsql = "select sum(accessible_amount) as accessible_amount,  m.exchange_rate from salesChallan_dtl m"
        strsql = strsql & " inner join sales_dtl d on m.Unit_code = d.Unit_code and m.location_code = d.location_code and m.doc_no = d.doc_no"
        strsql = strsql & " Where m.unit_code='" & gstrUnitId & "' and m.doc_no='" & Trim(pstrInvoice) & "'"
        strsql = strsql & " group by m.doc_no, m.exchange_rate"
        rsInvoice.Open(strsql, mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

        If Not rsInvoice.EOF Then
            strsql = "select bond17OpeningBal from saleConf"
            strsql = strsql & " where unit_code='" & gstrUnitId & "' and Invoice_Type ='EXP' AND Location_Code ='" & Trim(txtUnitCode.Text) & "' and datediff(dd,getdate(),fin_start_date)<=0  and datediff(dd,fin_end_date,getdate())<=0"

            dblOpeningBal = Val(Find_Value(strsql))
            mP_Connection.Execute("Update salesChallan_dtl set Bond17OpeningBal=" & dblOpeningBal & " where unit_code='" & gstrUnitId & "' and location_code='" & Trim(txtUnitCode.Text) & "' and doc_no='" & pstrInvoice & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            dblPrintExciseFormatExcise = CalculateExciseAmount(rsInvoice.Fields("accessible_amount").Value, rsInvoice.Fields("exchange_rate").Value)
            dblPrintExciseFormatExcise = dblPrintExciseFormatExcise + CalculateCessAmount(dblPrintExciseFormatExcise)
            If (dblOpeningBal + dblTotalExcise - dblPrintExciseFormatExcise) >= 0 Then
                CheckBond17OpeningBalance = dblOpeningBal + dblTotalExcise - dblPrintExciseFormatExcise
            Else
                CheckBond17OpeningBalance = 0
            End If
        Else
            CheckBond17OpeningBalance = 0
        End If
        If rsInvoice.State = ADODB.ObjectStateEnum.adStateOpen Then
            rsInvoice.Close()

            rsInvoice = Nothing
        End If
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

    Private Sub txtLorryNo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtLorryNo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*******************************************************************************
        'Author             :   Ashutosh Verma
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :
        'Comments           :   NA
        'Creation Date      :   26 Mar 2007
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                If txtOTLNo.Enabled Then txtOTLNo.Focus()
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

    Private Sub txtOTLNo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtOTLNo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*******************************************************************************
        'Author             :   Ashutosh Verma
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :
        'Comments           :   NA
        'Creation Date      :   26 Mar 2007
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
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

    Private Sub fpInvoice_Change(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ChangeEvent) Handles fpInvoice.Change
        With fpInvoice
            If .ActiveCol = 3 Then
                e.col = .ActiveCol
                e.row = .ActiveRow
                If ConvertToDate(.Text) > dtCurrentDate Then
                    .Text = VB6.Format(dtCurrentDate, gstrDateFormat)
                End If
            End If
        End With
    End Sub

    Private Function FordASNFileGeneration(ByVal pintdocno As Integer, ByVal pstraccountcode As String) As Boolean
        'Revised By     : Manoj Kr. Vaish
        'Revised On     : 14 May 2009
        'Arguments      : INvoice No
        'Issue ID       : eMpro-20090513-31282
        'Reason         : Generate ASN File for FORD
        '--------------------------------------------------------------------------------------
        On Error GoTo ErrHandler

        Dim rsgetData As New ClsResultSetDB
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
        strquery = "select * from dbo.FN_GETASNDETAIL(" & pintdocno & ",'" & gstrUnitId & "')"
        rsgetData.GetResult(strquery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsgetData.GetNoRows > 0 Then
            If rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Length = 0 Then
                MessageBox.Show("Customer vendor code is not defined for Customer : " & pstraccountcode, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                FordASNFileGeneration = False
                Exit Function

            Else
                If rsgetData.GetValue("CUST_PLANTCODE").ToString.Trim().Length = 0 Then
                    MessageBox.Show("Unable To Get Plant Code For The Customer: " & pstraccountcode & " While Generating ASN File." & vbCrLf & _
                                    "Invoice Can't Be Locked", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                    FordASNFileGeneration = False
                    Exit Function
                End If
                strASNdata = "856HD20200000000000000000" & Space(5 - rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim.Length) + rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim & rsgetData.GetValue("CUST_PLANTCODE").ToString.Trim & Space(5 - rsgetData.GetValue("CUST_PLANTCODE").ToString.Trim.Length) & ",     ,     *" & vbCrLf
                '10853890
                If rsgetData.GetValue("INTERMEDIATE_CONSIGNEE_CODE").ToString.Trim().Length = 0 Then
                    strASNdata = strASNdata & "856A M" & mInvNo.ToString.Trim() & Space(10 - mInvNo.ToString.Trim().Length) & Space(5 - rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim.Length) & rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim & "09" & VB6.Format(rsgetData.GetValue("INVOICE_DATE"), "yymmdd") & VB6.Format(rsgetData.GetValue("INVOICE_DATE"), "hhmm") & VB6.Format(rsgetData.GetValue("INVOICE_DATE"), "yymmdd") & VB6.Format(rsgetData.GetValue("INVOICE_DATE"), "hhmm") & Space(10) & "+00000000100KG+00000000080KG" & "AE0N" & Space(8) & rsgetData.GetValue("TRANSPORT_TYPE").ToString() & Space("12") & "M" & VB.Right(mInvNo, 5) & Space(4) & Space(35) & "M" & VB.Right(mInvNo, 5) & Space(5) & rsgetData.GetValue("CUST_PLANTCODE").ToString.Trim() & Space(5 - rsgetData.GetValue("CUST_PLANTCODE").ToString.Trim.Length) & rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim & Space(6) & rsgetData.GetValue("ARL_CODE").ToString.Trim() & Space(5 - rsgetData.GetValue("ARL_CODE").ToString.Trim().Length) & VB6.Format(GetServerDateTime, "mmddhhmm") & Space(3) & "0000000.00" & vbCrLf
                Else
                    strASNdata = strASNdata & "856A M" & mInvNo.ToString.Trim() & Space(10 - mInvNo.ToString.Trim().Length) & Space(5 - rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim.Length) & rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim & "09" & VB6.Format(rsgetData.GetValue("INVOICE_DATE"), "yymmdd") & VB6.Format(rsgetData.GetValue("INVOICE_DATE"), "hhmm") & VB6.Format(rsgetData.GetValue("INVOICE_DATE"), "yymmdd") & VB6.Format(rsgetData.GetValue("INVOICE_DATE"), "hhmm") & Space(10) & "+00000000100KG+00000000080KG" & "AE0N" & Space(8) & rsgetData.GetValue("TRANSPORT_TYPE").ToString() & Space("12") & "M" & VB.Right(mInvNo, 5) & Space(4) & Space(35) & "M" & VB.Right(mInvNo, 5) & Space(5) & rsgetData.GetValue("CUST_PLANTCODE").ToString.Trim() & Space(5 - rsgetData.GetValue("CUST_PLANTCODE").ToString.Trim.Length) & rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim & Space(6) & rsgetData.GetValue("INTERMEDIATE_CONSIGNEE_CODE").ToString.Trim() & Space(5 - rsgetData.GetValue("INTERMEDIATE_CONSIGNEE_CODE").ToString.Trim().Length) & VB6.Format(GetServerDateTime, "mmddhhmm") & Space(3) & "0000000.00" & vbCrLf
                End If
                '10853890
                Dcount = 2
                strcontainerdespQty = Find_Value("select sum(isnull(to_box,0)-isnull(from_box,0)+1) as Desp_Qty from sales_dtl where Unit_code='" & gstrUnitId & "' and doc_no=" & pintdocno)
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

                strtotalquantity = CInt(Find_Value("select sum(isnull(sales_quantity,0)) from sales_dtl where  UNIT_CODE = '" & gstrUnitId & "' and  doc_no=" & pintdocno))
                While Len(strtotalquantity) < 8
                    strtotalquantity = "0" + strtotalquantity
                End While
                strASNdata = strASNdata & strtotalquantity
                strnoofItems = CInt(Find_Value("select COUNT(*) NOOFITEMS  from sales_dtl where  UNIT_CODE = '" & gstrUnitId & "' and  doc_no=" & pintdocno))
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
                    dblcummulativeQty = Find_Value("SELECT DBO.UDF_GET_CUMMULATIVEQTY('" & gstrUnitId & "','" & rsgetData.GetValue("CUST_PLANTCODE").ToString() & "','" & rsgetData.GetValue("CUST_PART_CODE").ToString() & "'," & pintdocno & ")")
                    dblSalesQty = rsgetData.GetValue("SALES_QUANTITY")
                    dblcummulativeQty = dblcummulativeQty + dblSalesQty
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
                    mstrupdateASNdtl = Trim(mstrupdateASNdtl) & "UPDATE MKT_ASN_INVDTL SET ASN_STATUS=1,CUMMULATIVE_QTY=" & dblcummulativeQty & " WHERE UNIT_CODE='" & gstrUnitId & "' AND DOC_NO=" & pintdocno & " AND CUST_PART_CODE='" & rsgetData.GetValue("CUST_PART_CODE").ToString().Trim() & "' AND CUST_PLANTCODE='" & rsgetData.GetValue("CUST_PLANTCODE").ToString().Trim & "'" & vbCrLf
                    mstrupdateASNCumFig = Trim(mstrupdateASNCumFig) & "UPDATE MKT_ASN_CUMFIG SET CUMMULATIVE_QTY=" & dblcummulativeQty & " WHERE UNIT_CODE='" & gstrUnitId & "' AND CUST_PART_CODE='" & rsgetData.GetValue("CUST_PART_CODE").ToString().Trim() & "' AND CUST_PLANTCODE='" & rsgetData.GetValue("CUST_PLANTCODE").ToString().Trim & "'" & vbCrLf
                    rsgetData.MoveNext()
                Loop

                Dcount = Dcount + 1
                strASNdata = strASNdata & "856T " & Mid("0000", Dcount.ToString.Length, 5) & Dcount & Mid("00000000", strTotalQty.ToString.Length(), 9) & strTotalQty
                gstrASNPath = ReadValueFromINI(Application.StartupPath & "\mind.cfg", "ASNPATH-" & gstrUnitId, "Filepath")
                gstrASNPathForEDI = ReadValueFromINI(Application.StartupPath & "\mind.cfg", "ASNPATH-" & gstrUnitId, "FilepathforEDI")

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
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

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
        Dim Rs As ClsResultSetDB
        AllowASNTextFileGeneration = False
        strQry = "Select isnull(AllowASNTextGeneration,0) as AllowASNTextGeneration from customer_mst where UNIT_CODE='" & gstrUnitId & "' AND Customer_Code='" & Trim(pstraccountcode) & "'"
        Rs = New ClsResultSetDB
        If Rs.GetResult(strQry) = False Then GoTo ErrHandler
        If Rs.GetValue("AllowASNTextGeneration") = "True" Then
            AllowASNTextFileGeneration = True
        Else
            AllowASNTextFileGeneration = False
        End If

        Rs.ResultSetClose()
        Rs = Nothing
        Exit Function
ErrHandler:
        Rs = Nothing
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub txtASNNumber_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtASNNumber.KeyPress
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
    Private Function GROUPOASNFileGeneration(ByVal pintdocno As Integer, ByVal pstraccountcode As String) As Boolean
        'Revised By     : PRASHANT RAJPAL
        'Revised On     : 29 AUG 2012
        'Arguments      : INvoice No
        'Issue ID       : 10265065 
        'Reason         : Generate ASN File for GROUPO
        '--------------------------------------------------------------------------------------
        On Error GoTo ErrHandler

        Dim rsgetData As New ClsResultSetDB
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

        strASNdata = "856HDR"
        strquery = "select * from dbo.FN_GETASNDETAIL_GROUPO(" & pintdocno & ",'" & gstrUnitId & "')"
        rsgetData.GetResult(strquery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsgetData.GetNoRows > 0 Then
            If rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Length = 0 Then
                MessageBox.Show("Customer vendor code is not defined for Customer : " & pstraccountcode, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                GROUPOASNFileGeneration = False
                Exit Function

            Else
                If rsgetData.GetValue("CUST_PLANTCODE").ToString.Trim().Length = 0 Then
                    MessageBox.Show("Unable To Get Plant Code For The Customer: " & pstraccountcode & " While Generating ASN File." & vbCrLf & _
                                    "Invoice Can't Be Locked", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                    GROUPOASNFileGeneration = False
                    Exit Function
                End If
                strASNdata = strASNdata & Space(25 - rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim.Length) + rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim
                strASNdata = strASNdata & Space(25 - rsgetData.GetValue("ACCOUNT_CODE").ToString.Trim.Length) + rsgetData.GetValue("ACCOUNT_CODE").ToString.Trim
                strASNdata = strASNdata & Space(15 - mInvNo.ToString.Trim().Length) & +rsgetData.GetValue("doc_no").ToString.Trim
                strASNdata = strASNdata & Space(14 - VB6.Format(rsgetData.GetValue("INVOICE_DATE").ToString.Trim, "YYYYMMDD").Length) + VB6.Format(rsgetData.GetValue("INVOICE_DATE").ToString.Trim, "YYYYMMDD")
                strASNdata = strASNdata & VB6.Format(rsgetData.GetValue("INVOICE_TIME").ToString.Trim, "HHMMSS")
                strASNdata = strASNdata & Space(14 - VB6.Format(rsgetData.GetValue("INVOICE_DATE").ToString.Trim, "YYYYMMDD").Length) + VB6.Format(rsgetData.GetValue("INVOICE_DATE").ToString.Trim, "YYYYMMDD")
                strASNdata = strASNdata & VB6.Format(rsgetData.GetValue("INVOICE_TIME").ToString.Trim, "HHMMSS")
                strASNdata = strASNdata & Space(14 - VB6.Format(rsgetData.GetValue("INVOICE_DATE").ToString.Trim, "YYYYMMDD").Length) + VB6.Format(rsgetData.GetValue("INVOICE_DATE").ToString.Trim, "YYYYMMDD")
                strASNdata = strASNdata & VB6.Format(rsgetData.GetValue("INVOICE_TIME").ToString.Trim, "HHMMSS")
                strASNdata = strASNdata & Space(20 - rsgetData.GetValue("CONSIGNEE_CODE").ToString.Trim.Length) + rsgetData.GetValue("CONSIGNEE_CODE").ToString.Trim
                strASNdata = strASNdata & Space(20 - rsgetData.GetValue("DOCK_CODE").ToString.Trim.Length) + rsgetData.GetValue("DOCK_CODE").ToString.Trim
                strASNdata = strASNdata & Space(25 - rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim.Length) + rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim
                strASNdata = strASNdata & Space(25 - rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim.Length) + rsgetData.GetValue("CUST_VENDOR_CODE").ToString.Trim
                strASNdata = strASNdata & Space(25 - rsgetData.GetValue("ACCOUNT_CODE").ToString.Trim.Length) + rsgetData.GetValue("ACCOUNT_CODE").ToString.Trim
                strASNdata = strASNdata & Space(10 - rsgetData.GetValue("CUSTOMER_EDICODE").ToString.Trim.Length) + rsgetData.GetValue("CUSTOMER_EDICODE").ToString.Trim
                strASNdata = strASNdata & Space(4) + "1"
                strASNdata = strASNdata & Space(15 - rsgetData.GetValue("CONTAINER_DESP_QTY").ToString.Trim.Length) + rsgetData.GetValue("CONTAINER_DESP_QTY").ToString.Trim
                strASNdata = strASNdata & Space(15 - rsgetData.GetValue("TYPE_OF_PKGS").ToString.Trim.Length) + rsgetData.GetValue("TYPE_OF_PKGS").ToString.Trim
                strASNdata = strASNdata & Space(15 - rsgetData.GetValue("CONTAINER_QTY").ToString.Trim.Length) + rsgetData.GetValue("CONTAINER_QTY").ToString.Trim & vbCrLf


                rsgetData.MoveFirst()
                Do While Not rsgetData.EOFRecord
                    dblcummulativeQty = 0
                    dblSalesQty = 0
                    dblContainerQty = 0
                    dblcummulativeQty = Find_Value("SELECT DBO.UDF_GET_CUMMULATIVEQTY('" & gstrUnitId & "','" & rsgetData.GetValue("CUST_PLANTCODE").ToString() & "','" & rsgetData.GetValue("CUST_PART_CODE").ToString() & "'," & pintdocno & ")")
                    dblSalesQty = rsgetData.GetValue("SALES_QUANTITY")
                    dblcummulativeQty = dblcummulativeQty + dblSalesQty
                    dblContainerQty = rsgetData.GetValue("CONTAINER_QTY")

                    strASNdata = strASNdata & "856DTL"
                    strASNdata = strASNdata & Space(30 - rsgetData.GetValue("CUST_PART_CODE").ToString.Trim.Length) + rsgetData.GetValue("CUST_PART_CODE").ToString.Trim
                    'strASNdata = strASNdata & rsgetData.GetValue("CUST_PART_CODE").ToString().Trim & Space(30 - rsgetData.GetValue("CUST_PART_CODE").ToString().Length())
                    dblSalesQty = rsgetData.GetValue("Sales_Quantity").ToString().Trim & Space(10 - rsgetData.GetValue("Sales_Quantity").ToString().Length())
                    strASNdata = strASNdata & Mid("000000000", dblSalesQty.ToString.Length(), 10) & dblSalesQty

                    strTotalQty = Val(strTotalQty) + Val(rsgetData.GetValue("SALES_QUANTITY"))
                    Dcount = Dcount + 1

                    strASNdata = strASNdata & Mid("000000000", dblcummulativeQty.ToString().Length(), 10) & dblcummulativeQty '& vbCrLf
                    strASNdata = strASNdata & Space(35 - rsgetData.GetValue("CUST_REF").ToString.Trim.Length) + rsgetData.GetValue("CUST_REF").ToString.Trim & vbCrLf
                    Dcount = Dcount + 2
                    Dcount = Dcount + 1
                    mstrupdateASNdtl = Trim(mstrupdateASNdtl) & "UPDATE MKT_ASN_INVDTL SET ASN_STATUS=1,CUMMULATIVE_QTY=" & dblcummulativeQty & " WHERE UNIT_CODE='" & gstrUnitId & "' AND DOC_NO=" & pintdocno & " AND CUST_PART_CODE='" & rsgetData.GetValue("CUST_PART_CODE").ToString().Trim() & "' AND CUST_PLANTCODE='" & rsgetData.GetValue("CUST_PLANTCODE").ToString().Trim & "'" & vbCrLf
                    mstrupdateASNCumFig = Trim(mstrupdateASNCumFig) & "UPDATE MKT_ASN_CUMFIG SET CUMMULATIVE_QTY=" & dblcummulativeQty & " WHERE UNIT_CODE='" & gstrUnitId & "' AND CUST_PART_CODE='" & rsgetData.GetValue("CUST_PART_CODE").ToString().Trim() & "' AND CUST_PLANTCODE='" & rsgetData.GetValue("CUST_PLANTCODE").ToString().Trim & "'" & vbCrLf
                    rsgetData.MoveNext()
                Loop

                Dcount = Dcount + 1
                strASNFilepath = Find_Value("select isnull(ASN_GROUP_localPath,0)as ASN_GROUP_localPath from sales_parameter where unit_code='" & gstrUnitId & "'")
                strASNFilepathforEDI = Find_Value("select isnull(ASN_GROUP_EDIPath,0)as ASN_GROUP_EDIPath from sales_parameter where unit_code='" & gstrUnitId & "'")
                If Directory.Exists(strASNFilepath) = False Then
                    Directory.CreateDirectory(strASNFilepath)
                End If
                If Directory.Exists(strASNFilepathforEDI) = False Then
                    Directory.CreateDirectory(strASNFilepathforEDI)
                End If
                strASNFilepath = strASNFilepath & "\" & mInvNo.ToString() & ".dat"
                strASNFilepathforEDI = strASNFilepathforEDI & "\" & mInvNo.ToString() & ".dat"

                fs = File.Create(strASNFilepath)
                sw = New StreamWriter(fs)
                strASNdata = strASNdata.Substring(0, strASNdata.ToString.Length - 1)
                sw.Write(strASNdata)
                sw.Close()
                fs.Close()
                If File.Exists(strASNFilepathforEDI) = False Then
                    File.Copy(strASNFilepath, strASNFilepathforEDI)
                End If
                rsgetData.ResultSetClose()
                rsgetData = Nothing

                GROUPOASNFileGeneration = True
            End If
        Else
            MessageBox.Show("Unable To Generate ASN File . Invoice Can't Be Locked", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
            GROUPOASNFileGeneration = False
        End If

        Exit Function
ErrHandler:
        GROUPOASNFileGeneration = False
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function CheckASNTYPE(ByVal pstraccountcode As String) As String
        On Error GoTo ErrHandler
        Dim rsgetASNNumber As ClsResultSetDB
        Dim strsql As String
        rsgetASNNumber = New ClsResultSetDB
        strsql = "select ISNULL(ASN_type,'') ASN_type from customer_mst where Unit_code='" & gstrUnitId & "' and CUSTOMER_CODE='" & pstraccountcode & "'"
        rsgetASNNumber.GetResult(strsql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)

        If rsgetASNNumber.GetNoRows > 0 Then
            CheckASNTYPE = IIf(IsDBNull(rsgetASNNumber.GetValue("ASN_type")), "", rsgetASNNumber.GetValue("ASN_type"))
        End If

        rsgetASNNumber.ResultSetClose()
        rsgetASNNumber = Nothing
        Exit Function
ErrHandler:
        rsgetASNNumber = Nothing
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

    End Function
    Private Function GetTaxGlSl(ByVal TaxType As String) As String
        Dim objRecordSet As New ADODB.Recordset
        On Error GoTo ErrHandler
        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close()
        objRecordSet.Open("SELECT tx_glCode, tx_slCode FROM fin_TaxGlRel where tx_rowType = 'ARTAX' AND tx_taxId ='" & TaxType & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
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

    Private Sub chkPrintReprint_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkPrintReprint.CheckedChanged
        Try
            If chkPrintReprint.Checked Then
                txtInvoice.Text = String.Empty
                Opt_no.Text = "Reprint Invoice"
            Else
                txtInvoice.Text = String.Empty
                Opt_no.Text = "Print Invoice"
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub btnExceptionInvoices_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExceptionInvoices.Click
        Try
            Dim objExceptionInvoices As New frmExceptionInvoices
            objExceptionInvoices.SetInvoiceType = "EXP"
            objExceptionInvoices.ShowDialog()
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub IRN_QRBarcode()
        Try
            Dim rsGENERATEBARCODE As ClsResultSetDB
            Dim straccountcode As String
            Dim strPrintMethod As String = ""
            Dim strSQL As String = ""
            Dim intTotalNoofSlabs As Integer = 0
            Dim intRow As Short
            Dim strBarcodeMsg As String
            Dim strBarcodeMsg_paratemeter As String
            Dim ObjBarcodeHMI As New Prj_BCHMI.cls_BCHMI(gstrUnitId)
            Dim stimage As ADODB.Stream
            Dim strQuery As String
            Dim Rs As ADODB.Recordset
            Dim pstrPath As String = ""
            Dim blnCROP_QRIMAGE As Boolean = False


            pstrPath = gstrUserMyDocPath
            strSQL = "SELECT TOP 1 1 FROM SALESCHALLAN_DTL_IRN I INNER JOIN SALESCHALLAN_DTL_IRN_BARCODE B ON I.UNIT_CODE=B.UNIT_CODE AND I.DOC_NO=B.DOC_NO WHERE I.UNIT_CODE = '" & gstrUnitId & "' AND I.DOC_NO='" & Trim(Me.txtInvoice.Text) & "'" & " AND ISNULL(I.IRN_NO,'')<>'' AND ISNULL(B.BARCODE_DATA,'')<>'' "

            If DataExist(strSQL) = True Then
                strBarcodeMsg = ObjBarcodeHMI.GenerateQRBarCodeForIRN(gstrUserMyDocPath, Trim(txtInvoice.Text), gstrCONNECTIONSTRING)

                If VB.Left(strBarcodeMsg, 1) <> "Y" Then
                    MsgBox("Problem While Generating Barcode Image.", vbInformation, ResolveResString(100))
                    Exit Sub
                Else
                    strBarcodeMsg_paratemeter = Mid(strBarcodeMsg, 3)
                    stimage = New ADODB.Stream
                    stimage.Type = ADODB.StreamTypeEnum.adTypeBinary
                    stimage.Open()
                    pstrPath = pstrPath & "QRBarcodeImgIRN.wmf"

                    blnCROP_QRIMAGE = CBool(Find_Value("SELECT CROP_IRN_QRBARCODE  FROM SALES_PARAMETER (NOLOCK) WHERE UNIT_CODE='" + gstrUnitId + "'"))
                    If blnCROP_QRIMAGE = True Then
                        Dim bmp As New Bitmap(pstrPath)
                        Dim picturebox1 As New PictureBox
                        picturebox1.Image = ImageTrim(bmp)
                        picturebox1.Image.Save(pstrPath)
                        picturebox1 = Nothing
                    End If

                    stimage.LoadFromFile(pstrPath)

                    strQuery = "select  BARCODE_DATA,Doc_No ,barcodeimage from SALESCHALLAN_DTL_IRN_BARCODE where UNIT_CODE = '" & gstrUnitId & "' AND Doc_No=" & Trim(Me.txtInvoice.Text)

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
End Class