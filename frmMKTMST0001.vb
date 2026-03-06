Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports FPSpreadADO
Imports ADODB
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Data.SqlClient

Friend Class frmMKTMST0001
	Inherits System.Windows.Forms.Form
	'---------------------------------------------------------------------------
    'Copyright          :   MIND Ltd.
	'Form Name          :   frmMKTMST0001.frm
	'Created By         :   Ananya Nath
	'Created on         :   04/02/2002
	'Revised By         :
	'Revision Date      :
	'Description        :   Customer Master.
	'Revision History   :   on 11-10-2002 dll ref. is changed, desc. fields are added
	'Revision History   :   on 12-10-2002 Fax and Email, bank not mendatory.
	'Revision History   :   on 11-10-2002 dll ref. is changed, desc. fields are added
	'Revision History   :   on 12-10-2002 Fax and Email, bank not mendatory.
	'Revision History   :   By Ajay Vashistha on 17-07-2003
	'                       1) To check if GL and SL for Customer Exists in Ar_Docmaster if Yes then
	'                          the details should not be editable
	'                       2) Field Added - Customer Supplied Material Inc. In CST
    '                   :   By Jyolsna VN on 15-Dec-2004 (as part of RFQ - development)
	'                       1) Added option group for Confirmed and Non Confirmed Customers.
	'                       2) Added HO Address tab in General Details to capture HO Address details
	'                          for Non Confirmed Customers.
	'---------------------------------------------------------------------------
	'Revision Date      :   16/10/2007
	'Revised By         :   NEHA CHADHA
	'Description        :   21258 -Additional fields for capturing (non mandatory)Pan number,service tax number,Bank details like Bank name, bank address, account number (alphanumeric),swift number,IBAN/SOT number
	'---------------------------------------------------------------------------------
	'Revision Date      :   28 Aug 2008
	'Revised By         :   Manoj Kr Vaish
	'Issue ID           :   eMpro-20080828-21178
	'History            :   Merging of Customer Master Form of North and South.
	'                   :   Adding a new functionality of multiple shipping address
	'---------------------------------------------------------------------------------
    'Revised By         : Manoj Kr. Vaish
    'Issue ID           : eMpro-20090223-27780
    'Revision Date      : 24 Feb 2009
    'History            : Shipping address shouldn't be validated ,if the address is inactive.
    '***********************************************************************************
    'Revised By         : Manoj Kr. Vaish
    'Issue ID           : eMpro-20090611-32362
    'Revision Date      : 12 Jun 2009
    'History            : Add New field of Customer EDI Code-Hilex Nissan CSV File Genaration
    '****************************************************************************************
    'Revised By         : Manoj Kr. Vaish
    'Issue ID           : eMpro-20090624-32847
    'Revision Date      : 24 Jun 2009
    'History            : Add New field of Customer Plant Code-Ford ASN File Genaration
    'Modified By Amit Rana on 20/April/2011 for multiunit change
    '****************************************************************************************
    'Revised By         : Prashant Rajpal
    'Issue ID           : 10153377
    '****************************************************************************************
    'Revised By     -  Shubhra Verma
    'Revised On     -  20 Mar 2013
    'Issue ID       -  10368468
    'Revised History-  Checkbox added for T6 Box Label
    '-------------------------------------------------------------------------------------------------------------------------------------------------------
    'Revised By     -  Parveen Kumar
    'Revised On     -  31 Oct 2014
    'Issue ID       -  10690771.
    'Revised History-  ASN Arrival Time-Ship Duration Added.
    '-------------------------------------------------------------------------------------------------------------------------------------------------------
    'Revised By     -  Geetanjali Aggarwal
    'Revised On     -  31 Oct 2014
    'Purpose        -  10688280 - KAM code addition for Sales Provisioning
    '-------------------------------------------------------------------------------------------------------------------------------------------------------
    'Revised By     -  Vinod Singh
    'Revised On     -  08 jan 2015
    'Purpose        -  10736222 — eMPro - CT2 - ARE3 functionality
    '-------------------------------------------------------------------------------------------------------------------------------------------------------
    'Revised By     -  Prashant Rajpal
    'Revised On     -  10 May 2017 -11-May 2017
    'Purpose        -  GST CHANGES
    '-------------------------------------------------------------------------------------------------------------------------------------------------------
    'Modified By         :   Ashish Sharma
    'Modified On         :   22 MAR 2018
    'Issue ID            :   101482956 VAT Report Format for MTL Sharjah
    '-------------------------------------------------------------------------------------------------------

    'Declare Variables
	Dim mintFormIndex As Short
    Dim mobjEmpDll As New EMPDataBase.EMPDB(gstrUNITID)
    Dim mrsEmpDll As New EMPDataBase.CRecordset
    Dim plantName As String = String.Empty

    'Samiksha customer master authorization
    Dim authorize_active As Boolean = False
    'Samiksha Credit limit changes
    Dim isCreditLimitMandatory As Boolean = False


    Private Enum enmshipdetail
		VAL_DEFAULT = 1
		VAL_INACTIVE = 2
		VAL_SHIPCODE = 3
		VAL_SHIPDESC = 4
		VAL_SHIPADD1 = 5
		VAL_SHIPADD2 = 6
		VAL_CITY = 7
        VAL_DISCT = 8
        VAL_STATE = 9
        VAL_SHIP_GSTIN_ID = 10 'By Abhijit 17-Aug-2017
        VAL_GSTSTATECODE = 11
        VAL_GSTSTATEDESC = 12
        VAL_COUNTRY = 13
        VAL_Pin = 14
        VAL_PHONE = 15
        VAL_FAX = 16
        VAL_EMAILID = 17
        VAL_CONTACTPERSON = 18
        VAL_DESIGNATION = 19
        VAL_DISTANCE_FROM_UNIT = 20
        VAL_INACTIVEDATE = 21
        VAL_DELETE = 22
	End Enum
	
    'Const MaxGridHdrCols As Short = 18 
    Const MaxGridHdrCols As Short = 22
	Dim mstrShippingCode As String
	Dim mstrCustomerCode As String
	Dim mstrInsertUpdate As String
    Dim mintTotalRecord As Integer
    Dim mstrA4insert As String


    Private Sub chkConsigneeWiseLoc_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkConsigneeWiseLoc.CheckStateChanged
        On Error GoTo ErrHandler
        Dim rsChk As ClsResultSetDB
        If chkConsigneeWiseLoc.CheckState = 0 Then
            rsChk = New ClsResultSetDB
            If rsChk.GetResult("Select distinct Warehousefile_location,releaseFile_location,BackUpLocation from scheduleparameter_mst Where Unit_Code='" & gstrUNITID & "' And customer_code='" & Trim(txtCustCode.Text) & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly) = False Then GoTo ErrHandler
            If rsChk.GetNoRows > 1 Then
                MsgBox("You can't uncheck as different consignee locations are present for this customer." & vbCrLf & "Please make all locations same in Release File Parameter Master " & vbCrLf & "and then try again!", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100))
                chkConsigneeWiseLoc.CheckState = System.Windows.Forms.CheckState.Checked
                rsChk = Nothing
                Exit Sub
            End If
            rsChk = Nothing
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        rsChk = Nothing
    End Sub
    Private Sub ChkCSI_GT_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ChkCSI_GT.CheckStateChanged
        On Error GoTo ErrHandler
        If ChkCSI_GT.CheckState = 1 Then
            fraLedger.Visible = True
        Else
            fraLedger.Visible = False
        End If
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub chkJobWork_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles chkJobWork.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If KeyAscii = 13 Then
            chkmilkvan.Focus()
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

    Private Sub cmdhelpCreditTerms_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdhelpCreditTerms.Click
        On Error GoTo errHandler
        Dim strCreditTerms() As String
        Dim strString As String
        Dim strcode As String
        strString = txtCreditTermId.Text & "%"
        If txtCreditTermId.Text <> "" Then strcode = "Where Unit_Code='" & gstrUNITID & "' And CrTrm_TermId like '" & strString & "'" Else strcode = " Where Unit_Code='" & gstrUNITID & "'"
        strCreditTerms = ctlCurrCode.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select CrTrm_TermId, CrTrm_Desc from Gen_CreditTrmMaster " & strcode & " ", "Credit Term Listing", 1)
        If UBound(strCreditTerms) = -1 Then Me.txtCreditTermId.Focus() : Exit Sub
        If (strCreditTerms(0)) = "0" Then
            ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
            txtCreditTermId.Text = ""
            txtCreditTermId.Focus()
        Else
            txtCreditTermId.Text = Trim(strCreditTerms(0))
            lblCreditDesc.Text = Trim(strCreditTerms(1))
            txtCurrencyCode.Focus()
        End If
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdhelpCurrency_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdhelpCurrency.Click
        On Error GoTo ErrHandler
        Dim strCurrency() As String
        Dim strString As String
        Dim strcode As String
        strString = txtCurrencyCode.Text & "%"
        If txtCurrencyCode.Text <> "" Then strcode = "Where Unit_Code='" & gstrUNITID & "' And currency_code like '" & strString & "'" Else strcode = " Where Unit_Code='" & gstrUNITID & "'"
        strCurrency = ctlCurrCode.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select currency_code, description from currency_mst " & strcode & " ", "Currency Listing")
        If UBound(strCurrency) = -1 Then Me.txtCurrencyCode.Focus() : Exit Sub
        If (strCurrency(0)) = "0" Then
            ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
            txtCurrencyCode.Text = ""
            txtCurrencyCode.Focus()
        Else
            txtCurrencyCode.Text = Trim(strCurrency(0))
            lblCurrDesc.Text = Trim(strCurrency(1))
            txtCustvendCode.Focus()
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    'Samiksha : Changes For Customer Master Authorization
    Private Sub cmdhelpCustCode_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdhelpCustCode.Click
        Dim Index As Short = cmdhelpCustCode.GetIndex(eventSender) 'If user clicks on Customer Help Button then the listing of Customer Code will be populated.
        On Error GoTo ErrHandler
        Dim strCustCode() As String = Nothing
        Dim strCust As String
        Dim strString As String
        strString = txtCustCode.Text & "%"
        With ctlCustCode
            .CreateDSN(gstrCONNECTIONSERVER, gstrCONNECTIONDESCRIPTION, gstrCONNECTIONDSN, gstrCONNECTIONDATABASE)
            .ConnectAsUser = gstrCONNECTIONUSER
            .ConnectThroughDSN = gstrCONNECTIONDSN
            .ConnectWithPWD = gstrCONNECTIONPASSWORD
        End With
        If authorize_active = False Then
            If txtCustCode.Text <> "" Then
                strCust = " Where Unit_Code='" & gstrUNITID & "' And Customer_Code like '" & strString & "'"
            Else
                strCust = " Where Unit_Code='" & gstrUNITID & "'"
            End If
        ElseIf authorize_active = True Then
            If txtCustCode.Text <> "" Then
                strCust = "  Where Unit_Code='" & gstrUNITID & "' And Customer_Code like '" & strString & "' union Select Customer_code,Cust_name from customer_mst_authorization Where Unit_Code='" & gstrUNITID & "' And Customer_Code like '" & strString & "' AND ( ISNULL(Authorization_Status,'') = 'Return' or ISNULL(Authorization_Status,'')='')"
            Else
                strCust = " Where Unit_Code='" & gstrUNITID & "' union Select Customer_code,Cust_name from customer_mst_authorization Where Unit_Code='" & gstrUNITID & "' AND ( ISNULL(Authorization_Status,'') = 'Return' or ISNULL(Authorization_Status,'')='')"
            End If
        Else
            strCust = " Where Unit_Code='" & gstrUNITID & "'"

        End If
        If Me.OptNonConfCust.Checked = True Then
            If txtCustCode.Text <> "" Then
                strCust = " Where Unit_Code='" & gstrUNITID & "' And Customer_Code like '" & strString & "'"
            Else
                strCust = " Where Unit_Code='" & gstrUNITID & "'"
            End If
        End If
        If Me.OptConfirmedCust.Checked = True Then
            strCustCode = ctlCustCode.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select Customer_code,Cust_name from customer_mst " & strCust & " ", "Customer Listing")
        ElseIf Me.OptNonConfCust.Checked = True Then
            strCustCode = ctlCustCode.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select Customer_code,Cust_name from emp_non_conf_cust_mst " & strCust & " ", "Customer Listing")
        End If


        If UBound(strCustCode) = -1 Then
            txtCustCode.Text = ""
            Exit Sub
        End If

        txtCustCode.Tag = ""
        If strCustCode(0) = "0" Then
            ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO) : txtCustCode.Text = "" : txtCustCode.Focus() : Exit Sub
        Else
            Me.txtCustCode.Text = strCustCode(0)
            txtCustCode.Tag = txtCustCode.Text
        End If
        FillASN_Status()
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    'Samiksha : Changes for Customer Master Authorization
    Private Sub cmdgrpCustMst_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.UCbtndgrp.ButtonClickEventArgs) Handles cmdgrpCustMst.ButtonClick
        On Error GoTo ErrHandler
        Dim BolCheck As Boolean
        Dim strshippingCode As String
        Dim intMaxLoop As Short
        Dim strTblname As String = String.Empty

        Select Case e.Button
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD
                Me.GroupBox1.Visible = False
                If OptConfirmedCust.Checked = True Then
                    tabDetails.SelectedIndex = 0
                    tabCustomer.SelectedIndex = 0
                ElseIf OptNonConfCust.Checked = True Then
                    tabDetails.SelectedIndex = 0
                    tabCustomer.SelectedIndex = 2
                End If
                Me._cmdHelpGlobalCustomer_0.Enabled = True
                txtCustCode.Tag = lblGlobalCustCodeDesc.Text
                ' txtCustCode.Tag = txtCustCode.Text
                Me.cmdgrpCustMst.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE) = True
                Call RefreshForm()
                If OptConfirmedCust.Checked = True Then Call DisableHOAddress()
                If OptNonConfCust.Checked = True Then Call DisableForNonConfCust()
                txtCustCode.Enabled = False
                txtCustCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                Me.chkStandard.CheckState = System.Windows.Forms.CheckState.Checked
                Me.txtAcctLedger.Enabled = False
                cmdhelpCustCode(1).Enabled = False
                cmdhelpCurrency(1).Enabled = True
                cmdhelpCreditTerms.Enabled = True
                BtnHelpKAM.Enabled = True   '10688280 -Add KAM Code in Customer Master.
                'GST CHANGES
                cmdGSTBillStateHelp.Enabled = True
                txtGSTBillState.Enabled = False
                txtGSTBillState.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                'GST CHANGES
                '101482956
                VisibleTaxRegistrationNumber()
                If OptNonConfCust.Checked = False Then
                    Me.cmdhelpLedger(0).Enabled = True
                    Me.cmdhelpLedger(1).Enabled = True
                    Me.txtAccSubLedger.Enabled = False
                    Me.cmdhelpSubLedger(0).Enabled = False
                    Me.cmdhelpSubLedger(2).Enabled = False
                    Me.txtAcctLedger.Enabled = True
                    Me.txtAcctLedgerExc.Enabled = True
                    chkShpmntThruWh.CheckState = System.Windows.Forms.CheckState.Unchecked
                    chkConsigneeWiseLoc.CheckState = System.Windows.Forms.CheckState.Checked
                End If
                chktT6Label.CheckState = CheckState.Unchecked
                Call AddBlankRowinGrid()
                '  mstrCustomerCode = GenerateCustomerNo()
                mstrCustomerCode = lblGlobalCustCodeDesc.Text
                mstrShippingCode = GenerateShippingCode(mstrCustomerCode)
                Call spShippingAddess.SetText(enmshipdetail.VAL_SHIPCODE, spShippingAddess.MaxRows, mstrShippingCode)
                mintTotalRecord = 0

                '  Me.txtCustName.Text = lblGlobalCustNameDesc.Text
                Me.txtWebSite.Focus()
                ' Me.txtCustName.Focus()
                Me.txtCustCode.ReadOnly = True
                Me.txtCustCode.Text = lblGlobalCustCodeDesc.Text
                Me.optDomestic.Checked = True
                Me.txtCustName.Enabled = False
                Me.txtWebSite.Enabled = False
                Me.txtCustLoc.Enabled = False

            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE 'Saves a Rec
                Dim custauthStatus As String = String.Empty
                If (Me.cmdgrpCustMst.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD) Then 'Add Mode
                    'Samiksha  : Customer Master Authorisation
                    If authorize_active = True Then
                        strTblname = "customer_mst_authorization"
                    ElseIf authorize_active = False Then
                        strTblname = "customer_mst"
                    End If
                    'Samiksha Credit limit changes
                    'Samiksha branchcode changes
                    customer_mst_insert(strTblname)
                End If
                If Me.cmdgrpCustMst.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then 'Edit Mode
                    'Samiksha: Customer Master Authorization
                    If authorize_active = True Then
                        custauthStatus = getAuthorisationstatus()
                        If custauthStatus = "Auth" Or custauthStatus = "Return" Or custauthStatus = "" Then
                            strTblname = "customer_mst_authorization"

                        End If
                    ElseIf authorize_active = False Then
                        strTblname = "customer_mst"
                    End If
                    'Samiksha Credit limit changes
                    customer_mst_update(strTblname)

                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT 'Edit Record

                txtCustCode.Tag = txtCustCode.Text
                Call FillASN_Status()
                Call Enable()
                If OptConfirmedCust.Checked = True Then Call DisableHOAddress()
                If OptNonConfCust.Checked = True Then Call DisableForNonConfCust()
                txtCustCode.Enabled = False
                txtCustCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                cmdhelpCustCode(1).Enabled = False
                cmdhelpCurrency(1).Enabled = True
                cmdhelpCreditTerms.Enabled = True
                BtnHelpKAM.Enabled = True   '10688280 -Add KAM Code in Customer Master.
                'GST CHANGES
                cmdGSTBillStateHelp.Enabled = True
                txtGSTBillState.Enabled = False
                txtGSTBillState.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                'GST CHANGES
                '101482956
                VisibleTaxRegistrationNumber()
                Me.cmdgrpCustMst.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD) = False
                Me.cmdgrpCustMst.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                If OptConfirmedCust.Checked = True Then
                    Me.cmdhelpLedger(0).Enabled = True
                    Me.cmdhelpSubLedger(2).Enabled = True
                    Me.cmdhelpLedger(1).Enabled = True
                    Me.cmdhelpSubLedger(0).Enabled = True
                    Me.txtAccSubLedgerExc.Enabled = True
                    Me.txtAcctLedgerExc.Enabled = True
                End If
                Me.txtAccSubLedger.Enabled = True
                Me.txtAcctLedger.Enabled = True
                With spShippingAddess
                    For intMaxLoop = 1 To .MaxRows Step 1
                        .Row = intMaxLoop
                        .Col = enmshipdetail.VAL_INACTIVE
                        If .Value = False Then
                            .Row = intMaxLoop
                            .Col = enmshipdetail.VAL_DEFAULT
                            .Lock = False
                            .Row = intMaxLoop
                            .Col = enmshipdetail.VAL_INACTIVE
                            .Lock = False
                        End If
                        .Row = intMaxLoop
                        .Col = enmshipdetail.VAL_DEFAULT
                        If .Value = True Then
                            .Row = intMaxLoop
                            .Col = enmshipdetail.VAL_INACTIVE
                            .Lock = True
                        End If
                    Next
                End With
                'Anupam Kumar, Date:21052024
                'strshippingCode = GenerateShippingCode(Trim(txtCustCode.Text))
                'Call AddBlankRowinGrid()
                'Call AddShippingDetailinNewRow(Trim(txtCustCode.Text))
                'Call spShippingAddess.SetText(enmshipdetail.VAL_SHIPCODE, spShippingAddess.MaxRows, strshippingCode)
                Me.txtCustName.Focus()
                mobjEmpDll.CConnection.OpenConnection(gstrDSNName, gstrDatabaseName)
                Call mrsEmpDll.OpenRecordset("Select top 1 docM_voType,docM_voNo,docM_PartyID,docM_TermId,docM_transNo,docM_drcrNoteType,docM_voDate,docM_custDocNo,docM_custDocDate,docM_srcMod,docM_srcDocType,docM_srcDocNo,docM_srcDocDate,docM_currency,docM_unit,docM_Amt,docM_xchgRate,docM_baseCurAmt,docM_dueDate,docM_payDueDate,docM_expDueDate,docM_ctrlGLAc,docM_ctrlSLAc,docM_contestAmt,docM_contestReason,docM_contestRem,docM_amtPaid,docM_amtPaidBaseCur,docM_open,docM_cancel,docM_cancelRem,docM_cancelBy,docM_cancelOn,docM_glTransNo,docM_remarks,docM_shipTo,docM_soldTo,docM_drCrNoteReason,docM_glRevTransNo,docM_revRemarks,docM_ourRefNo,docM_ourRefDate,docM_partyBalance,docM_refdetails,docM_InterUnit,docM_Contra,docm_refdetail,docM_subType from ar_docmaster Where docM_unit='" & gstrUNITID & "' And docm_partyid='" & txtCustCode.Text & "' and docm_ctrlglac='" & txtAcctLedger.Text & "' and docm_ctrlslac='" & txtAccSubLedger.Text & "'", ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                If Not mrsEmpDll.EOF_Renamed Then
                    txtAccSubLedger.Enabled = False
                    txtAcctLedger.Enabled = False
                    cmdhelpLedger(0).Enabled = False
                    cmdhelpSubLedger(2).Enabled = False
                End If
                Me.txtCustName.Enabled = False
                If lblGlobalCustCodeDesc.Text.Trim() = "".Trim() Then
                    Me._cmdHelpGlobalCustomer_0.Enabled = True
                Else
                    Me._cmdHelpGlobalCustomer_0.Enabled = False
                End If


                mrsEmpDll.CloseRecordset()
                mobjEmpDll.CConnection.CloseConnection()
                Me.txtCustName.Enabled = False
                Me.txtWebSite.Enabled = False
                Me.txtCustLoc.Enabled = False

            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE 'Deletion of a rec
                txtCustCode.Tag = txtCustCode.Text
                If ConfirmWindow(10054, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_CRITICAL, 60096) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then 'Confirmation before deletion
                    If OptConfirmedCust.Checked = True Then
                        BolCheck = ExistInTransaction() ' Checks if the particular rec. has taken part in any transaction
                        If BolCheck = True Then
                            ConfirmWindow(10231, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO) 'If so, than Record can not be deleted.
                            Me.txtCustCode.Focus()
                            Exit Sub
                        End If
                        Call delete() ' Delete Function will be called.
                    ElseIf OptNonConfCust.Checked = True Then
                        'Check whether customer code has been referenced in other transactions
                        mobjEmpDll.CConnection.OpenConnection(gstrDSNName, gstrDatabaseName)
                        Call mrsEmpDll.OpenRecordset("select customer_code from emp_cust_enquiry_hdr Where Unit_Code='" & gstrUNITID & "' And customer_code='" & Trim(txtCustCode.Text) & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic)
                        If Not mrsEmpDll.EOF_Renamed Then
                            ConfirmWindow(10231, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO) 'If so, than Record can not be deleted.
                            mrsEmpDll.CloseRecordset()
                            mobjEmpDll.CConnection.CloseConnection()
                            Me.txtCustCode.Focus()
                            Exit Sub
                        Else
                            mrsEmpDll.CloseRecordset()
                            mobjEmpDll.CConnection.CloseConnection()
                            Call DeleteNonConfCust() ' Delete Function will be called.
                        End If
                    End If
                    Call RefreshForm()
                    Call Disabled()
                    cmdhelpCurrency(1).Enabled = False
                    cmdhelpCreditTerms.Enabled = False
                    cmdhelpCustCode(1).Enabled = True
                    BtnHelpKAM.Enabled = False   '10688280 -Add KAM Code in Customer Master.
                    'GST CHANGES
                    cmdGSTBillStateHelp.Enabled = False
                    txtGSTBillState.Enabled = False
                    txtGSTBillState.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)

                    'GST CHANGES
                    txtCustCode.Enabled = True
                    txtCustCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    Me.cmdgrpCustMst.Focus()
                    txtCustCode.Focus()
                Else
                    Me.cmdgrpCustMst.Focus()
                    txtCustCode.Focus()
                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL ' Cancellation
                If ConfirmWindow(10053, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION, 60095) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_NO Then
                    Me.cmdgrpCustMst.Focus()
                    Me._cmdHelpGlobalCustomer_0.Enabled = False
                Else
                    RefreshForm()
                    Me.cmdgrpCustMst.Revert()
                    If txtCustCode.Tag <> "" And Me.cmdgrpCustMst.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                        Call display()
                    Else
                        Me.cmdgrpCustMst.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                        Me.cmdgrpCustMst.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                        Me.cmdgrpCustMst.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL) = False
                        Me.cmdgrpCustMst.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                        Me.lblGlobalCustCodeDesc.Text = ""
                        Me.lblGlobalCustCodeDesc.Tag = ""
                    End If
                    Call Disabled()
                    Me.txtCustCode.Enabled = False
                    txtCustCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    Me.txtCustCode.Text = ""
                    Me.cmdhelpCustCode(1).Enabled = True
                    Me._cmdHelpGlobalCustomer_0.Enabled = False
                    Me._cmdHelpGlobalCustomer_0.Focus()
                    Me.cmdgrpCustMst.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                    Me.cmdgrpCustMst.Focus()

                    ' Me.txtCustCode.Focus()

                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE ' If user clicks on close button
                Me.Close()
        End Select
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        mobjEmpDll.CConnection.CloseConnection()
        mrsEmpDll.CloseRecordset()
    End Sub
    'Samiksha  : Customer Master Authorisation

    Public Sub customer_mst_insert(ByVal tblName As String)
        If OptConfirmedCust.Checked = True Then
            If ValidateRowData(spShippingAddess.ActiveRow, enmshipdetail.VAL_DISTANCE_FROM_UNIT) = False Then Exit Sub
            If ValBeforesave() = False Then Exit Sub 'Checks for mendtory fields
        ElseIf OptNonConfCust.Checked = True Then
            If ValNonConfBeforesave() = False Then Exit Sub
        End If
        txtCustCode.Tag = txtCustCode.Text
        If OptConfirmedCust.Checked = True Then
            '   txtCustCode.Text = GenerateCustomerNo()
            'Swati Code...
            txtCustCode.Enabled = False
            txtCustName.Enabled = False
            txtCustCode.Text = lblGlobalCustCodeDesc.Text
            '  txtCustName.Text = lblGlobalCustNameDesc.Text
            txtCustName.Text = lblGlobalCustCodeDesc.Tag.ToString
            'end
            txtCustLoc.Text = GenerateCustLoc()
            'Samiksha Credit limit changes

            'Samiksha branchcode changes
            Call Insert(tblName)
        ElseIf OptNonConfCust.Checked = True Then
            ' txtCustCode.Text = GenerateNonConfCustomerNo()
            txtCustCode.Text = lblGlobalCustCodeDesc.Text
            txtCustName.Text = lblGlobalCustCodeDesc.Tag.ToString
            txtCustCode.Enabled = False
            txtCustName.Enabled = False
            'Samiksha Credit limit changes
            'Samiksha branchcode changes

            Call InsertNonConfDetails("customer_mst")
        End If
        Call Disabled()
        'If OptConfirmedCust.Checked = True Then
        '    MsgBox("Customer Code : " & txtCustCode.Text & " is Successfully Generated in the system" & vbCrLf & "Customer Location is : " & txtCustLoc.Text, MsgBoxStyle.Information, "eMPro")
        'ElseIf OptNonConfCust.Checked = True Then
        '    MsgBox("Customer Code : " & txtCustCode.Text & " is Successfully Generated in the system" & vbCrLf, MsgBoxStyle.Information, "eMPro")
        'End If
        cmdhelpCustCode(1).Enabled = True
        Me.cmdgrpCustMst.Revert()
        cmdgrpCustMst.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
        txtCustCode.Enabled = True
        txtCustCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        txtCustCode.Focus()
        txtCustCode.Text = ""

        Exit Sub
        Exit Sub
    End Sub
    'Samiskha: Changes for Customer Master Authorisation
    Public Sub customer_mst_update(ByVal tblname As String)
        If OptConfirmedCust.Checked = True Then
            If ValidateRowData(spShippingAddess.ActiveRow, enmshipdetail.VAL_DISTANCE_FROM_UNIT) = False Then Exit Sub
            If ValBeforesave() = False Then Exit Sub
        ElseIf OptNonConfCust.Checked = True Then
            If ValNonConfBeforesave() = False Then Exit Sub
        End If
        txtCustCode.Tag = txtCustCode.Text
        Call InsertA4customer(txtCustCode.Text)
        If Not IsNothing(mstrA4insert) Then
            If Len(mstrA4insert.ToString) > 0 Then
                mP_Connection.Execute(mstrA4insert, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End If
        End If

        If OptConfirmedCust.Checked = True Then
            Call Update_Renamed_DotNet(tblname)
        ElseIf OptNonConfCust.Checked = True Then
            Call UpdateNonConfDetails("customer_mst")
        End If
        Call Disabled()
        cmdhelpCustCode(1).Enabled = True
        Me.cmdgrpCustMst.Revert()
        cmdgrpCustMst.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
        txtCustCode.Enabled = True
        txtCustCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        Call ShowShippingDetail(Trim(txtCustCode.Text))
        txtCustCode.Text = ""
        txtCustCode.Focus()
        Me.txtCustName.Enabled = False
        Me._cmdHelpGlobalCustomer_0.Enabled = False

        Exit Sub
    End Sub
    'Samiksha : Changes for Customer MAster Authorization
    Public Function getAuthorisationstatus() As String
        On Error GoTo ErrHandler

        Dim StrSql As String = ""
        Dim authStatus As String = ""

        StrSql = "Select ISNULL(Authorization_Status,'') from customer_mst_authorization(nolock) where UNIT_CODE='" & gstrUNITID & "' and Customer_Code='" & txtCustCode.Text & "'"
        authStatus = SqlConnectionclass.ExecuteScalar(StrSql)

        Return authStatus

ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

    End Function


    Private Sub cmdhelpLedger_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdhelpLedger.Click
        Dim Index As Short = cmdhelpLedger.GetIndex(eventSender)
        On Error GoTo ErrHandler
        Dim strLedger() As String
        Dim strLedg As String
        Dim strString As String
        Select Case Index
            Case 0
                strString = txtAcctLedger.Text & "%"
                If txtAcctLedger.Text <> "" Then strLedg = "Where Unit_Code='" & gstrUNITID & "' And glM_glCode like '" & strString & "' and glm_transtag ='1' " Else strLedg = "Where Unit_Code='" & gstrUNITID & "' And glm_transtag ='1' "
                strLedger = Me.ctlhelpLedger.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select glM_glCode,glm_desc from fin_glmaster " & strLedg & " ", "General Ledger Listing")
                If UBound(strLedger) = -1 Then txtAcctLedger.Focus() : Exit Sub
                If strLedger(0) = "0" Then
                    ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO) : txtAcctLedger.Focus() : Exit Sub
                Else
                    Me.txtAccSubLedger.Enabled = True
                    Me.cmdhelpSubLedger(2).Enabled = True
                    Me.txtAcctLedger.Text = Trim(strLedger(0))
                    lblleddesc.Text = Trim(strLedger(1))
                End If
                If Me.txtAccSubLedger.Enabled = True Then Me.txtAccSubLedger.Focus() Else Me.txtBankAcct1.Focus()
            Case 1
                strString = txtAcctLedgerExc.Text & "%"
                If txtAcctLedgerExc.Text <> "" Then strLedg = "Where Unit_Code='" & gstrUNITID & "' And  glM_glCode like '" & strString & "' and glm_transtag ='1' " Else strLedg = "Where Unit_Code='" & gstrUNITID & "' And glm_transtag ='1' "
                strLedger = Me.ctlhelpLedger.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select glM_glCode,glm_desc from fin_glmaster " & strLedg & " ", "General Ledger Listing")
                If UBound(strLedger) = -1 Then txtAcctLedger.Focus() : Exit Sub
                If strLedger(0) = "0" Then
                    ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO) : txtAcctLedger.Focus() : Exit Sub
                Else
                    Me.txtAccSubLedgerExc.Enabled = True
                    Me.cmdhelpSubLedger(0).Enabled = True
                    Me.txtAcctLedgerExc.Text = Trim(strLedger(0))
                    lblleddescExc.Text = Trim(strLedger(1))
                End If
                If Me.txtAccSubLedgerExc.Enabled = True Then Me.txtAccSubLedgerExc.Focus() Else Me.txtExciseRange.Focus()
        End Select
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub GSTBillstateDesc()

        On Error GoTo ErrHandler

        Dim strStatecode() As String
        Dim strSubLedg As String
        Dim strString As String
        With spShippingAddess
            strStatecode = ctlhelpsubledger.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select GST_STATE_CODE,STATE_NAME from state_mst ", "State Master ")
            If UBound(strStatecode) = -1 Then Exit Sub
            If strStatecode(0) = "0" Then
                ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO) : Exit Sub
            Else
                Call .SetText(enmshipdetail.VAL_GSTSTATECODE, .ActiveRow, strStatecode(0))
                Call .SetText(enmshipdetail.VAL_GSTSTATEDESC, .ActiveRow, strStatecode(1))
            End If

        End With
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdhelpSubLedger_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdhelpSubLedger.Click
        Dim Index As Short = cmdhelpSubLedger.GetIndex(eventSender)
        On Error GoTo ErrHandler
        Dim strSubLedger() As String
        Dim strSubLedg As String
        Dim strString As String
        Select Case Index
            Case 2
                If Me.txtAcctLedger.Text <> "" Then
                    strString = txtAccSubLedger.Text & "%"
                    If txtAccSubLedger.Text <> "" Then strSubLedg = "Where Unit_Code='" & gstrUNITID & "' And slM_slCode like '" & strString & "' and slm_transtag = '1' and slm_glcode = '" & txtAcctLedger.Text & "' " Else strSubLedg = "Where Unit_Code='" & gstrUNITID & "' And slm_transtag = '1' and slm_glcode = '" & txtAcctLedger.Text & "'"
                    strSubLedger = Me.ctlhelpsubledger.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select slM_slCode,slm_desc from fin_slmaster " & strSubLedg & " ", "Sub Ledger Listing")
                    If UBound(strSubLedger) = -1 Then txtAccSubLedger.Focus() : Exit Sub
                    If strSubLedger(0) = "0" Then
                        ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO) : txtAccSubLedger.Text = "" : txtAccSubLedger.Focus() : Exit Sub
                    Else
                        Me.txtAccSubLedger.Text = Trim(strSubLedger(0))
                        Me.lblsubdesc.Text = Trim(strSubLedger(1))
                        Me.txtBankAcct1.Focus()
                    End If
                End If
            Case 0
                If Me.txtAcctLedgerExc.Text <> "" Then
                    strString = txtAccSubLedgerExc.Text & "%"
                    If txtAccSubLedgerExc.Text <> "" Then strSubLedg = "Where Unit_Code='" & gstrUNITID & "' And slM_slCode like '" & strString & "' and slm_transtag = '1' and slm_glcode = '" & txtAcctLedgerExc.Text & "' " Else strSubLedg = "Where Unit_Code='" & gstrUNITID & "' And slm_transtag = '1' and slm_glcode = '" & txtAcctLedgerExc.Text & "'"
                    strSubLedger = Me.ctlhelpsubledger.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select slM_slCode,slm_desc from fin_slmaster " & strSubLedg & " ", "Sub Ledger Listing")
                    If UBound(strSubLedger) = -1 Then txtAccSubLedgerExc.Focus() : Exit Sub
                    If strSubLedger(0) = "0" Then
                        ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO) : txtAccSubLedgerExc.Text = "" : txtAccSubLedgerExc.Focus() : Exit Sub
                    Else
                        Me.txtAccSubLedgerExc.Text = Trim(strSubLedger(0))
                        Me.lblsubdescExc.Text = Trim(strSubLedger(1))
                        Me.txtExciseRange.Focus()
                    End If
                End If
        End Select
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub ctlFormHeader1_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlFormHeader1.Click
        On Error GoTo errHandler
        Call ShowHelp("HLPMKTMST0001.htm")
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub Form_Initialize_Renamed()
        '' This Procedure is Added by Rajeev Gupta on 25-May-2005
        '' There a New Column added in sales_parameter ie ShowTaxExciseDetails
        '' if ShowTaxExciseDetails = 1 then Tax and Excise Detail is Visible = True
        '' if ShowTaxExciseDetails = 0 then Tax and Excise Detail is Visible = False
        '' By default ShowTaxExciseDetails = 1
        On Error GoTo Err_Handler
        Dim gobjDB1 As New ClsResultSetDB
        gobjDB1.GetResult("SELECT ShowTaxExciseDetails FROM sales_parameter Where Unit_Code='" & gstrUNITID & "'")
        If gobjDB1.GetValue("ShowTaxExciseDetails") = True Then
        Else
            tabDetails.TabPages.RemoveAt(3)
        End If

        gobjDB1.ResultSetClose()
        gobjDB1 = Nothing

        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub OptConfirmedCust_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OptConfirmedCust.CheckedChanged
        If eventSender.Checked Then
            Call ConfCustOpt()
        End If
    End Sub
    Private Sub OptNonConfCust_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OptNonConfCust.CheckedChanged
        If eventSender.Checked Then
            Call NonConfCustOpt()
        End If
    End Sub
    Private Sub TxtAccNo_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles TxtAccNo.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Return Then cmdgrpCustMst.Focus()
    End Sub
    Private Sub txtAccSubLedger_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtAccSubLedger.KeyPress
        On Error GoTo ErrHandler
        Dim keyascii As Short
        keyascii = Asc(e.KeyChar)
        Select Case keyascii
            Case 39, 34, 96, 45
                e.Handled = True
        End Select
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtAccSubLedger_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAccSubLedger.TextChanged
        On Error GoTo ErrHandler
        If txtAcctLedger.Text = "" Then txtAccSubLedger.Text = ""
        Me.lblsubdesc.Text = ""
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtAccSubLedger_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtAccSubLedger.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then Call cmdhelpSubLedger_Click(cmdhelpSubLedger.Item(2), New System.EventArgs()) 'Listing of AccountLedgers will be displayed
        Exit Sub ' if user presses F1 Key , while cursor is in acctledger Field.
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtAccSubLedger_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtAccSubLedger.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        If Me.txtAccSubLedger.Text <> "" Then
            mobjEmpDll.CConnection.OpenConnection(gstrDSNName, gstrDatabaseName)
            mrsEmpDll.OpenRecordset("select slm_desc, slm_slcode from fin_slmaster Where Unit_Code='" & gstrUNITID & "' And slm_slcode = '" & txtAccSubLedger.Text & "' and slm_transtag = '1' and slm_glcode = '" & txtAcctLedger.Text & "' ", ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockReadOnly)
            mrsEmpDll.Filter_Renamed = " slm_slcode ='" & txtAccSubLedger.Text & "'"
            If mrsEmpDll.Recordcount = 0 Then
                ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                Me.txtAccSubLedger.Text = ""
                Me.txtAccSubLedger.Focus()
            Else
                Me.lblsubdesc.Text = mrsEmpDll.GetFieldValue("slm_desc", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
            End If
            mrsEmpDll.CloseRecordset()
            mobjEmpDll.CConnection.CloseConnection()
        End If
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        mrsEmpDll.CloseRecordset()
        mobjEmpDll.CConnection.CloseConnection()
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtAccSubLedger_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAccSubLedger.Enter
        On Error GoTo ErrHandler
        txtAccSubLedger.SelectionStart = 0
        txtAccSubLedger.SelectionLength = Len(txtAccSubLedger.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtAccSubLedgerExc_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtAccSubLedgerExc.KeyPress
        On Error GoTo ErrHandler
        Dim keyascii As Short
        keyascii = Asc(e.KeyChar)
        Select Case keyascii
            Case 39, 34, 96, 45
                e.Handled = True
        End Select
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtAccSubLedgerExc_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAccSubLedgerExc.TextChanged
        On Error GoTo ErrHandler
        If txtAcctLedgerExc.Text = "" Then txtAccSubLedgerExc.Text = ""
        Me.lblsubdescExc.Text = ""
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtAccSubLedgerExc_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAccSubLedgerExc.Enter
        On Error GoTo ErrHandler
        txtAccSubLedgerExc.SelectionStart = 0
        txtAccSubLedgerExc.SelectionLength = Len(txtAccSubLedgerExc.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtAccSubLedgerExc_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtAccSubLedgerExc.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then Call cmdhelpSubLedger_Click(cmdhelpSubLedger.Item(0), New System.EventArgs()) 'Listing of AccountLedgers will be displayed
        Exit Sub ' if user presses F1 Key , while cursor is in acctledger Field.
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtAccSubLedgerExc_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtAccSubLedgerExc.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        If Me.txtAccSubLedgerExc.Text <> "" Then
            mobjEmpDll.CConnection.OpenConnection(gstrDSNName, gstrDatabaseName)
            mrsEmpDll.OpenRecordset("select slm_desc, slm_slcode from fin_slmaster Where Unit_Code='" & gstrUNITID & "' And slm_slcode = '" & txtAccSubLedgerExc.Text & "' and slm_transtag = '1' and slm_glcode = '" & txtAcctLedgerExc.Text & "' ", ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockReadOnly)
            mrsEmpDll.Filter_Renamed = " slm_slcode ='" & txtAccSubLedgerExc.Text & "'"
            If mrsEmpDll.Recordcount = 0 Then
                ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                Me.txtAccSubLedgerExc.Text = ""
                Me.txtAccSubLedgerExc.Focus()
            Else
                Me.lblsubdescExc.Text = mrsEmpDll.GetFieldValue("slm_desc", EMPDataBase.EMPDB.ADODataType.ADOChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
            End If
            mrsEmpDll.CloseRecordset()
            mobjEmpDll.CConnection.CloseConnection()
        End If
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        mrsEmpDll.CloseRecordset()
        mobjEmpDll.CConnection.CloseConnection()
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtAcctLedger_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtAcctLedger.KeyPress
        On Error GoTo ErrHandler
        Dim keyascii As Short
        keyascii = Asc(e.KeyChar)
        Select Case keyascii
            Case 39, 34, 96, 45
                e.Handled = True
        End Select
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtAcctLedger_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAcctLedger.TextChanged
        On Error GoTo ErrHandler
        Me.lblleddesc.Text = "" : Me.lblsubdesc.Text = "" : Me.txtAccSubLedger.Text = ""
        If txtAcctLedger.Text = "" Then
            txtAccSubLedger.Text = ""
            Me.txtAccSubLedger.Enabled = False
            Me.cmdhelpSubLedger(2).Enabled = False
            Me.lblleddesc.Text = "" : Me.lblsubdesc.Text = ""
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtAcctLedger_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAcctLedger.Enter
        On Error GoTo ErrHandler
        txtAcctLedger.SelectionStart = 0
        txtAcctLedger.SelectionLength = Len(txtAcctLedger.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtAcctLedger_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtAcctLedger.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then Call cmdhelpLedger_Click(cmdhelpLedger.Item(0), New System.EventArgs()) 'Listing of AccountLedgers will be displayed
        Exit Sub ' if user presses F1 Key , while cursor is in acctledger Field.
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtAcctLedger_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtAcctLedger.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        If Len(txtAcctLedger.Text) > 0 Then
            mobjEmpDll.CConnection.OpenConnection(gstrDSNName, gstrDatabaseName)
            mrsEmpDll.OpenRecordset("select glm_glcode, glm_desc from fin_glmaster Where Unit_Code='" & gstrUNITID & "' and  glm_glcode = '" & txtAcctLedger.Text & "' and glm_transtag = '1' ", ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockReadOnly)
            mrsEmpDll.Filter_Renamed = " glm_glcode ='" & txtAcctLedger.Text & "'"
            If mrsEmpDll.Recordcount <= 0 Then
                ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                Me.txtAcctLedger.Text = ""
            Else
                lblleddesc.Text = mrsEmpDll.GetFieldValue("glm_desc", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                Me.txtAccSubLedger.Enabled = True
                Me.cmdhelpSubLedger(2).Enabled = True
            End If
            mrsEmpDll.CloseRecordset()
            mobjEmpDll.CConnection.CloseConnection()
        End If
        Me.cmdgrpCustMst.Focus()
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        mobjEmpDll.CConnection.CloseConnection()
        mrsEmpDll.CloseRecordset()
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtAcctLedgerExc_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtAcctLedgerExc.KeyPress
        On Error GoTo ErrHandler
        Dim keyascii As Short
        keyascii = Asc(e.KeyChar)
        Select Case keyascii
            Case 39, 34, 96, 45
                e.Handled = True
        End Select
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtAcctLedgerExc_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAcctLedgerExc.TextChanged
        On Error GoTo ErrHandler
        Me.lblleddescExc.Text = "" : Me.lblsubdescExc.Text = "" : Me.txtAccSubLedgerExc.Text = ""
        If txtAcctLedgerExc.Text = "" Then
            txtAccSubLedgerExc.Text = ""
            Me.txtAccSubLedgerExc.Enabled = False
            Me.cmdhelpSubLedger(0).Enabled = False
            Me.lblleddescExc.Text = "" : Me.lblsubdescExc.Text = ""
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtAcctLedgerExc_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAcctLedgerExc.Enter
        On Error GoTo ErrHandler
        txtAcctLedgerExc.SelectionStart = 0
        txtAcctLedgerExc.SelectionLength = Len(txtAcctLedgerExc.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtAcctLedgerExc_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtAcctLedgerExc.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then Call cmdhelpLedger_Click(cmdhelpLedger.Item(1), New System.EventArgs()) 'Listing of AccountLedgers will be displayed
        Exit Sub ' if user presses F1 Key , while cursor is in acctledger Field.
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtAcctLedgerExc_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtAcctLedgerExc.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        If Len(txtAcctLedgerExc.Text) > 0 Then
            mobjEmpDll.CConnection.OpenConnection(gstrDSNName, gstrDatabaseName)
            mrsEmpDll.OpenRecordset("select glm_glcode, glm_desc from fin_glmaster Where Unit_Code='" & gstrUNITID & "' And glm_glcode = '" & txtAcctLedgerExc.Text & "' and glm_transtag = '1' ", ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockReadOnly)
            mrsEmpDll.Filter_Renamed = " glm_glcode ='" & txtAcctLedgerExc.Text & "'"
            If mrsEmpDll.Recordcount <= 0 Then
                ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                Me.txtAcctLedgerExc.Text = ""
            Else
                lblleddescExc.Text = mrsEmpDll.GetFieldValue("glm_desc", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                Me.txtAccSubLedgerExc.Enabled = True
                Me.cmdhelpSubLedger(0).Enabled = True
            End If
            mrsEmpDll.CloseRecordset()
            mobjEmpDll.CConnection.CloseConnection()
        End If
        Me.cmdgrpCustMst.Focus()
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        mobjEmpDll.CConnection.CloseConnection()
        mrsEmpDll.CloseRecordset()
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtBankAcct3_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtBankAcct3.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If KeyAscii = 13 Then
            tabDetails.SelectedIndex = 3
            txtExciseRange.Focus()
        End If
        Select Case KeyAscii
            Case 39, 34, 96, 45
                eventArgs.Handled = True
        End Select
        Exit Sub
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtBillContPer_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtBillContPer.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If (KeyAscii < 97 Or KeyAscii > 122) And (KeyAscii < 65 Or KeyAscii > 90) And (KeyAscii <> 32) And (KeyAscii <> 45) And (KeyAscii <> 8) And (KeyAscii <> 127) Then
            KeyAscii = 0
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
    Private Sub txtBillDesig_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtBillDesig.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If KeyAscii = 13 Then
            tabCustomer.SelectedIndex = 2
            spShippingAddess.Focus()
        End If
        If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii < 97 Or KeyAscii > 122) And (KeyAscii < 65 Or KeyAscii > 90) And (KeyAscii <> 32) And (KeyAscii <> 45) And (KeyAscii <> 8) And (KeyAscii <> 127) Then
            KeyAscii = 0
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
    Private Sub txtBillEmail_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtBillEmail.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If (KeyAscii = 39) Then
            KeyAscii = 0
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
    Private Sub txtComRate_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtComRate.TextChanged
        On Error GoTo ErrHandler
        txtComRate.SelectionLength = Len(txtComRate.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtconsigneecode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtconsigneecode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo Err_Renamed
        If (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 8) Or KeyAscii = 45 Then
            KeyAscii = KeyAscii
        ElseIf KeyAscii = 13 Then
            chkStandard.Focus()
        Else
            KeyAscii = 0
        End If
        GoTo EventExitSub
Err_Renamed:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        GoTo EventExitSub
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtContPer1_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtContPer1.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If (KeyAscii < 97 Or KeyAscii > 122) And (KeyAscii < 65 Or KeyAscii > 90) And (KeyAscii <> 32) And (KeyAscii <> 45) And (KeyAscii <> 8) And (KeyAscii <> 127) Then
            KeyAscii = 0
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
    Private Sub txtCreditTermId_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCreditTermId.TextChanged
        On Error GoTo ErrHandler
        lblCreditDesc.Text = ""
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtCreditTermId_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCreditTermId.Enter
        On Error GoTo ErrHandler
        txtCreditTermId.SelectionStart = 0
        txtCreditTermId.SelectionLength = Len(txtCreditTermId.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtCreditTermId_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCreditTermId.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then Call cmdhelpCreditTerms_Click(cmdhelpCreditTerms, New System.EventArgs())
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtCreditTermId_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCreditTermId.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar) ' only digit can be entered.
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case 32, 39, 34, 96, 45
                KeyAscii = 0
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
    Private Sub txtCurrencyCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCurrencyCode.TextChanged
        On Error GoTo ErrHandler
        lblCurrDesc.Text = ""
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtCurrencyCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCurrencyCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If KeyAscii = 13 And txtCurrencyCode.Text <> "" Then
            If ValidateCur() = False Then
                ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                Me.txtCurrencyCode.Text = ""
                Me.txtCurrencyCode.Focus()
            End If
        ElseIf KeyAscii = 39 Or KeyAscii = 34 Or KeyAscii = 96 Or KeyAscii = 45 Then
            KeyAscii = 0
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
    Private Sub txtDesig1_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtDesig1.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If KeyAscii = 13 Then
            tabCustomer.SelectedIndex = 1
            txtBilladd1.Focus()
        End If
        If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii < 97 Or KeyAscii > 122) And (KeyAscii < 65 Or KeyAscii > 90) And (KeyAscii <> 32) And (KeyAscii <> 45) And (KeyAscii <> 8) And (KeyAscii <> 127) Then
            KeyAscii = 0
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
    Private Sub txtEmail1_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtEmail1.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If (KeyAscii = 39) Then
            KeyAscii = 0
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
    Private Sub txtHOfax_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtHOfax.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 44) And (KeyAscii <> 45) And (KeyAscii <> 8) And (KeyAscii <> 127) Then
            KeyAscii = 0
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
    Private Sub txtHOPhone_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtHOPhone.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 44) And (KeyAscii <> 45) And (KeyAscii <> 8) And (KeyAscii <> 127) Then
            KeyAscii = 0
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
    Private Sub txtShipContPer_KeyPress(ByRef KeyAscii As Short)
        On Error GoTo ErrHandler
        If (KeyAscii < 97 Or KeyAscii > 122) And (KeyAscii < 65 Or KeyAscii > 90) And (KeyAscii <> 32) And (KeyAscii <> 45) And (KeyAscii <> 8) And (KeyAscii <> 127) Then
            KeyAscii = 0
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtShipDesig_KeyPress(ByRef KeyAscii As Short)
        On Error GoTo ErrHandler
        If KeyAscii = 13 Then
            tabDetails.SelectedIndex = 1
            txtCreditTermId.Focus()
        End If
        If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii < 97 Or KeyAscii > 122) And (KeyAscii < 65 Or KeyAscii > 90) And (KeyAscii <> 32) And (KeyAscii <> 45) And (KeyAscii <> 8) And (KeyAscii <> 127) Then
            KeyAscii = 0
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtShipEmail_KeyPress(ByRef KeyAscii As Short)
        On Error GoTo ErrHandler
        If (KeyAscii = 39) Then
            KeyAscii = 0
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtContPer1_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtContPer1.Enter
        On Error GoTo ErrHandler
        txtContPer1.SelectionStart = 0
        txtContPer1.SelectionLength = Len(txtContPer1.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtContPer1_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtContPer1.Leave
        On Error GoTo ErrHandler
        txtContPer1.Text = StrConv(txtContPer1.Text, VbStrConv.ProperCase)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtBillContPer_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBillContPer.Enter
        On Error GoTo ErrHandler
        txtBillContPer.SelectionStart = 0
        txtBillContPer.SelectionLength = Len(txtBillContPer.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtBillContPer_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBillContPer.Leave
        On Error GoTo ErrHandler
        txtBillContPer.Text = StrConv(txtBillContPer.Text, VbStrConv.ProperCase)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtCurrencyCode_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCurrencyCode.Leave
        On Error GoTo ErrHandler
        txtCurrencyCode.Text = UCase(txtCurrencyCode.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtDesig1_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDesig1.Enter
        On Error GoTo ErrHandler
        txtDesig1.SelectionStart = 0
        txtDesig1.SelectionLength = Len(txtDesig1.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtDesig1_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDesig1.Leave
        On Error GoTo ErrHandler
        txtDesig1.Text = StrConv(txtDesig1.Text, VbStrConv.ProperCase)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtBillDesig_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBillDesig.Enter
        On Error GoTo ErrHandler
        txtBillDesig.SelectionStart = 0
        txtBillDesig.SelectionLength = Len(txtBillDesig.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtBillDesig_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBillDesig.Leave
        On Error GoTo ErrHandler
        txtBillDesig.Text = StrConv(txtBillDesig.Text, VbStrConv.ProperCase)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtDivision_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDivision.Enter
        On Error GoTo ErrHandler
        txtDivision.SelectionStart = 0
        txtDivision.SelectionLength = Len(txtDivision.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtEmail1_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtEmail1.Enter
        On Error GoTo ErrHandler
        txtEmail1.SelectionStart = 0
        txtEmail1.SelectionLength = Len(txtEmail1.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtBillEmail_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBillEmail.Enter
        On Error GoTo ErrHandler
        txtBillEmail.SelectionStart = 0
        txtBillEmail.SelectionLength = Len(txtBillEmail.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtExciseRange_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtExciseRange.Enter
        On Error GoTo ErrHandler
        txtExciseRange.SelectionStart = 0
        txtExciseRange.SelectionLength = Len(txtExciseRange.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtLst_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLst.Enter
        On Error GoTo ErrHandler
        txtLst.SelectionStart = 0
        txtLst.SelectionLength = Len(txtLst.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtoffadd11_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtoffadd11.KeyPress
        On Error GoTo ErrHandler
        Dim keyascii As Short
        keyascii = Asc(e.KeyChar)
        Select Case keyascii
            Case 39, 34, 96, 45
                e.Handled = True
        End Select
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtoffadd11_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtoffadd11.Leave
        On Error GoTo ErrHandler
        txtoffadd11.Text = StrConv(txtoffadd11.Text, VbStrConv.ProperCase)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtbilladd1_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBilladd1.Enter
        On Error GoTo ErrHandler
        txtBilladd1.SelectionStart = 0
        txtBilladd1.SelectionLength = Len(txtBilladd1.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtBilladd1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtBilladd1.KeyPress
        On Error GoTo ErrHandler
        Dim keyascii As Short
        keyascii = Asc(e.KeyChar)
        Select Case keyascii
            Case 39, 34, 96, 45
                e.Handled = True
        End Select
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtbilladd1_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBilladd1.Leave
        On Error GoTo ErrHandler
        txtBilladd1.Text = StrConv(txtBilladd1.Text, VbStrConv.ProperCase)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtOffAdd21_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtOffAdd21.KeyPress
        On Error GoTo ErrHandler
        Dim keyascii As Short
        keyascii = Asc(e.KeyChar)
        Select Case keyascii
            Case 39, 34, 96, 45
                e.Handled = True
        End Select
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtOffAdd21_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtOffAdd21.Leave
        On Error GoTo ErrHandler
        txtOffAdd21.Text = StrConv(txtOffAdd21.Text, VbStrConv.ProperCase)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtbilladd2_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBillAdd2.Enter
        On Error GoTo ErrHandler
        txtBillAdd2.SelectionStart = 0
        txtBillAdd2.SelectionLength = Len(txtBillAdd2.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtBillAdd2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtBillAdd2.KeyPress
        On Error GoTo ErrHandler
        Dim keyascii As Short
        keyascii = Asc(e.KeyChar)
        Select Case keyascii
            Case 39, 34, 96, 45
                e.Handled = True
        End Select
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtbilladd2_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBillAdd2.Leave
        On Error GoTo ErrHandler
        txtBillAdd2.Text = StrConv(txtBillAdd2.Text, VbStrConv.ProperCase)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtBillCity_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtbillCity.Enter
        Dim Index As Short = txtbillCity.GetIndex(eventSender)
        On Error GoTo ErrHandler
        txtbillCity(0).SelectionStart = 0
        txtbillCity(0).SelectionLength = Len(txtbillCity(0).Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtBillCountry_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBillCountry.Enter
        On Error GoTo ErrHandler
        txtBillCountry.SelectionStart = 0
        txtBillCountry.SelectionLength = Len(txtBillCountry.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtOffDist1_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtOffDist1.Enter
        On Error GoTo ErrHandler
        txtOffDist1.SelectionStart = 0
        txtOffDist1.SelectionLength = Len(txtOffDist1.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtOffDist1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtOffDist1.KeyPress
        On Error GoTo ErrHandler
        Dim keyascii As Short
        keyascii = Asc(e.KeyChar)
        Select Case keyascii
            Case 39, 34, 96, 45
                e.Handled = True
        End Select
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtOffDist1_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtOffDist1.Leave
        On Error GoTo ErrHandler
        txtOffDist1.Text = StrConv(txtOffDist1.Text, VbStrConv.ProperCase)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtBillDist_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBillDist.Enter
        On Error GoTo ErrHandler
        txtBillDist.SelectionStart = 0
        txtBillDist.SelectionLength = Len(txtBillDist.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtBillDist_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtBillDist.KeyPress
        On Error GoTo ErrHandler
        Dim keyascii As Short
        keyascii = Asc(e.KeyChar)
        Select Case keyascii
            Case 39, 34, 96, 45
                e.Handled = True
        End Select
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtBillDist_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBillDist.Leave
        On Error GoTo ErrHandler
        txtBillDist.Text = StrConv(txtBillDist.Text, VbStrConv.ProperCase)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtOffFax1_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtOffFax1.Enter
        On Error GoTo ErrHandler
        txtOffFax1.SelectionStart = 0
        txtOffFax1.SelectionLength = Len(txtOffFax1.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtOffFax1_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtOffFax1.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 44) And (KeyAscii <> 45) And (KeyAscii <> 8) And (KeyAscii <> 127) Then
            KeyAscii = 0
            GoTo EventExitSub
ErrHandler:
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtBillfax_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBillFax.Enter
        On Error GoTo ErrHandler
        txtBillFax.SelectionStart = 0
        txtBillFax.SelectionLength = Len(txtBillFax.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtBillfax_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtBillFax.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 44) And (KeyAscii <> 45) And (KeyAscii <> 8) And (KeyAscii <> 127) Then
            KeyAscii = 0
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
    Private Sub txtOffPhone1_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtOffPhone1.Enter
        On Error GoTo ErrHandler
        txtOffPhone1.SelectionStart = 0
        txtOffPhone1.SelectionLength = Len(txtOffPhone1.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtOffPhone1_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtOffPhone1.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 44) And (KeyAscii <> 45) And (KeyAscii <> 8) And (KeyAscii <> 127) Then
            KeyAscii = 0
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
    Private Sub txtBillPhone_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBillPhone.Enter
        On Error GoTo ErrHandler
        txtBillPhone.SelectionStart = 0
        txtBillPhone.SelectionLength = Len(txtBillPhone.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtBillPhone_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtBillPhone.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 44) And (KeyAscii <> 45) And (KeyAscii <> 8) And (KeyAscii <> 127) Then
            KeyAscii = 0
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
    Private Sub txtOffPin1_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtOffPin1.Enter
        On Error GoTo ErrHandler
        txtOffPin1.SelectionStart = 0
        txtOffPin1.SelectionLength = Len(txtOffPin1.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtBillPin_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBillPin.Enter
        On Error GoTo ErrHandler
        txtBillPin.SelectionStart = 0
        txtBillPin.SelectionLength = Len(txtBillPin.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtBillPin_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtBillPin.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8) And (KeyAscii <> 127) Then
            KeyAscii = 0
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
    Private Sub txtOffState1_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtOffState1.Enter
        On Error GoTo ErrHandler
        txtOffState1.SelectionStart = 0
        txtOffState1.SelectionLength = Len(txtOffState1.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtBillState_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBillState.Enter
        On Error GoTo ErrHandler
        txtBillState.SelectionStart = 0
        txtBillState.SelectionLength = Len(txtBillState.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtCustCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCustCode.TextChanged
        On Error GoTo ErrHandler
        If Me.cmdgrpCustMst.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW And txtCustCode.Text = "" Then ' If CustomerCode textbox does not contain any data than
            Call RefreshForm() ' all the fields will be blank.
            Call Disabled()
            Me.cmdgrpCustMst.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
            Me.cmdgrpCustMst.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
            Me.cmdgrpCustMst.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE) = False
            txtCustCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            Me.cmdhelpCustCode(1).Enabled = True
            Me.txtCustCode.Enabled = True
        End If
        'Samiksha branchcode changes
        'Samiksha Creditlimit changes
        If Me.cmdgrpCustMst.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW And txtCustCode.Text <> "" Then
            If Me.OptConfirmedCust.Checked = True Then
                If ValCustomerCode() = True Then Call display()
            ElseIf Me.OptNonConfCust.Checked = True Then
                If ValNonConfCustomerCode() = True Then Call NonConfCustDisplay()
            End If
        End If
        If Me.txtCustCode.Enabled = False Then Me.txtCustCode.Enabled = True : Me.txtCustCode.Focus()
        If Me.OptNonConfCust.Checked = True Then
            tabDetails.SelectedIndex = 0
            tabCustomer.SelectedIndex = 2
        Else
            tabDetails.SelectedIndex = 0
            tabCustomer.SelectedIndex = 0
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtCustCode_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCustCode.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then Call cmdhelpCustCode_Click(cmdhelpCustCode.Item(1), New System.EventArgs()) 'Listing of Customer Codes will be displayed
        Exit Sub ' if user presses F1 Key , while cursor is in CustCode Field.
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtCustCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCustCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Select Case KeyAscii
            Case 39, 34, 96, 45
                KeyAscii = 0
            Case 8
                Return
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
    'Samiksha : Customer Master Changes
    'Samiksha Credit limit changes
    'Samiksha branchcode changes

    Private Sub txtcustcode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtCustCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        Dim bolExist As Boolean
        If cmdgrpCustMst.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW And Len(txtCustCode.Text) > 0 Then
            If OptConfirmedCust.Checked = True Then
                bolExist = ValCustomerCode()
            ElseIf OptNonConfCust.Checked = True Then
                bolExist = ValNonConfCustomerCode()
            End If
            If bolExist = True Then
                If OptConfirmedCust.Checked = True Then Call display()
                If OptNonConfCust.Checked = True Then Call NonConfCustDisplay()
            Else
                Call RefreshForm()
                Call Disabled()
                Me.cmdhelpCustCode(1).Enabled = True
                Me.cmdgrpCustMst.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                Me.cmdgrpCustMst.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                Me.txtCustCode.Enabled = True
                Me.txtCustCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO) 'Else message 'corresponding details are not available.
                Cancel = True
                Me.cmdgrpCustMst.Focus()
                Me.txtCustCode.Focus()
            End If
        End If
        Call DisplayA4customer(txtCustCode.Text)
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub frmMKTMST0001_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo ErrHandler
        'Checking the form name in the Windows list
        mdifrmMain.CheckFormName = mintFormIndex
        frmModules.NodeFontBold(Tag) = True
        'This is to avoid the execution of the error handler
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTMST0001_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Tag) = False
        gblnCancelUnload = False
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTMST0001_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        Dim varShipcode As Object = Nothing
        On Error GoTo ErrHandler

        If (Shift = VB6.ShiftConstants.CtrlMask And KeyCode = System.Windows.Forms.Keys.N) Then   ' For example, Ctrl + N to Add New Row  '21052024
            If txtCustCode.Text.Trim.Length > 0 AndAlso tabCustomer.SelectedIndex = 2 AndAlso Me.cmdgrpCustMst.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                Dim strshippingCode As String
                strshippingCode = GenerateShippingCode(Trim(txtCustCode.Text))


                Call spShippingAddess.GetText(enmshipdetail.VAL_SHIPCODE, spShippingAddess.MaxRows, varShipcode)
                If spShippingAddess.MaxRows > 0 Then
                    If CInt((Microsoft.VisualBasic.Right(varShipcode, 2))) + 1 = CInt((Microsoft.VisualBasic.Right(strshippingCode, 2))) Then
                        Call AddBlankRowinGrid()
                        Call AddShippingDetailinNewRow(Trim(txtCustCode.Text))
                        Call spShippingAddess.SetText(enmshipdetail.VAL_SHIPCODE, spShippingAddess.MaxRows, strshippingCode)
                    End If
                Else
                    Call AddBlankRowinGrid()
                    Call AddShippingDetailinNewRow(Trim(txtCustCode.Text))
                    Call spShippingAddess.SetText(enmshipdetail.VAL_SHIPCODE, spShippingAddess.MaxRows, strshippingCode)
                End If

                





            End If
        End If

        If (KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0) Then Call ctlFormHeader1_Click(ctlFormHeader1, New System.EventArgs()) 'F4 key is used to call empHelp
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTMST0001_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                System.Windows.Forms.SendKeys.SendWait("{TAB}") 'If user press the Enter Key ,the focus will be advanced        Case vbKeyEscape  'If user press Escape than valCancel will be callked.
            Case System.Windows.Forms.Keys.Escape
                If Me.cmdgrpCustMst.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                    Call cmdgrpCustMst_ButtonClick(cmdgrpCustMst, New UCActXCtl.UCbtndgrp.ButtonClickEventArgs(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL))
                End If
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

    Private Sub frmMKTMST0001_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo ErrHandler
        Dim oRs As New ADODB.Recordset
        Dim intCount As Short
        Dim strSql As String = ""

        Me.GroupBox1.Visible = False
        mintFormIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader1.Tag)
        Call FitToClient(Me, fraMain, ctlFormHeader1, cmdgrpCustMst, 500)
        Call RefreshForm()
        SetGridHeading()

        oRs = mP_Connection.Execute("Select * From Lists Where Key1='INVBARPRINT' And Unit_Code = '" & gstrUNITID & "'")
        If Not (oRs.EOF And oRs.BOF) Then
            While Not oRs.EOF
                CmbPrintMethod.Items.Add(oRs("Key2").Value.ToString.Trim)
                oRs.MoveNext()
            End While
        End If
        oRs = Nothing
		With Me
            For intCount = 1 To 5
                cmdgrpCustMst.Enabled(intCount) = False
            Next
            '101482956
            plantName = UCase(Trim(GetPlantName()))
            VisibleTaxRegistrationNumber()
            Call Disabled()
            Me.chkStandard.CheckState = System.Windows.Forms.CheckState.Checked
            Me.cmdhelpCustCode(1).Enabled = True
            Me.txtCustCode.Enabled = True
            '  Me._cmdHelpGlobalCustomer_0.Enabled = True
            Me.txtCustCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            Me.OptConfirmedCust.Enabled = True
            Me.OptNonConfCust.Enabled = True
            Me.OptConfirmedCust.Checked = True
            Me.tabDetails.SelectedIndex = 0
            Me.tabCustomer.SelectedIndex = 0
            
        End With
        Me.optDomestic.Checked = True

        'Samiksha customer master Authorization

        SqlConnectionclass.OpenGlobalConnection()
        strSql = "select ISNULL(authorization_active,0) from Sales_Parameter where unit_code='" & gstrUNITID & "'"
        authorize_active = SqlConnectionclass.ExecuteScalar(strSql)

        'Samiksha Credit limit changes
        strSql = ""
        strSql = "select ISNULL(isCreditLimitMandatory,0) from Sales_Parameter where unit_code='" & gstrUNITID & "'"
        isCreditLimitMandatory = SqlConnectionclass.ExecuteScalar(strSql)

        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTMST0001_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        On Error GoTo errHandler
        Dim enmValue As eMPowerFunctions.ConfirmWindowReturnEnum
        If UnloadMode >= 0 And UnloadMode <= 5 Then
            If Me.cmdgrpCustMst.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Or Me.cmdgrpCustMst.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then 'If not View Mode
                enmValue = ConfirmWindow(10055, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNOCANCEL, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) 'Confirm before unloading the FORM
                If enmValue <> eMPowerFunctions.ConfirmWindowReturnEnum.VAL_CANCEL Then 'If  'YES' or 'NO'
                    If enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then 'If YES
                        Call cmdgrpCustMst_ButtonClick(cmdgrpCustMst, New UCActXCtl.UCbtndgrp.ButtonClickEventArgs(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE))
                        eventArgs.Cancel = True
                    ElseIf enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_NO Then  'If NO than Unload the Form
                        gblnCancelUnload = False 'Variable used in MDI Form before unloading
                        gblnFormAddEdit = False
                        Me.cmdgrpCustMst.Focus()
                    Else
                        gblnCancelUnload = True : gblnFormAddEdit = True ' If Cancel than Focus will be set on in the first field.
                        Me.cmdgrpCustMst.Focus()
                    End If
                Else
                    If cmdgrpCustMst.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Or cmdgrpCustMst.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then ' if Mode is add than focus will be fixed on Plant Code
                        Me.txtCustName.Focus()
                    Else
                        Me.txtCustCode.Focus() ' Else Focus will be set on Plant Name.
                    End If
                    gblnCancelUnload = True
                    gblnFormAddEdit = True
                End If
            Else
                gblnCancelUnload = False
            End If
        End If
        If gblnCancelUnload = True Then eventArgs.Cancel = True 'Do not unload FORM, if the value of gblncancelUnload is False
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub frmMKTMST0001_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error GoTo ErrHandler
        'Removing the form name from list
        mdifrmMain.RemoveFormNameFromWindowList = mintFormIndex
        frmModules.NodeFontBold(Tag) = False
        Me.Dispose()
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    '*********************************************'
    'Author         :       Ananya Nath
    'Arguments      :       None
    'Return Value   :       None
    'Description    :       This Function is used to clear the contents of all form level controls.
    '*********************************************'
    Private Sub RefreshForm()
        On Error GoTo ErrHandler
        Dim txtControl As System.Windows.Forms.Control
        LoopAllControlsInsideForm_MKTMST0001(Me.Controls)
        Me.lblGlobalCustCodeDesc.Text = ""
        Me.txtCreditTermId.Text = ""
        Me.txtCustLoc.Enabled = False
        txtCustLoc.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        fraLedger.Visible = False
        Call InitializeShippingSpread()
        optDomestic.Checked = True
        Me.lblGlobalCustCodeDesc.Text = ""
        txtA4orginal.Text = ""
        spA4grid.MaxRows = 0
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    '*********************************************'
    'Author         :       Ananya Nath
    'Arguments      :       None
    'Return Value   :       None
    'Description    :       Used to Desable all form level controls.
    '*********************************************'
    Private Sub Disabled() ' All the form level controls will be Disabled
        On Error GoTo ErrHandler
        DisableAllControlsInsideForm_MKTMST0001(Me.Controls)
        Me.txtAcctLedger.Enabled = False
        Me.txtAcctLedger.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        Me.txtCustLoc.Enabled = False
        Me.txtCustLoc.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        Me.OptConfirmedCust.Enabled = True
        Me.OptNonConfCust.Enabled = True
        Me._cmdHelpGlobalCustomer_0.Enabled = False
        With spShippingAddess
            .Row = 0 : .Row2 = .MaxRows
            .Col = enmshipdetail.VAL_DEFAULT
            .Col2 = enmshipdetail.VAL_INACTIVEDATE
            .BlockMode = True : .Lock = True
            .BlockMode = False
        End With
        Me.GroupBox1.Visible = False
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    '*********************************************'
    'Author         :       Ananya Nath
    'Arguments      :       None
    'Return Value   :       None
    'Description    :       Used to Enable all form level controls.
    '*********************************************'
    Private Sub Enable() ' All the form level controls will be enabled
        On Error GoTo ErrHandler
        EnableAllControlsInsideForm_MKTMST0001(Me.Controls)
        Me.txtAcctLedger.Enabled = False
        Me.txtCustLoc.Enabled = False
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtBankAcct1_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBankAcct1.Enter
        On Error GoTo ErrHandler
        Me.txtBankAcct1.SelectionStart = 0
        Me.txtBankAcct1.SelectionLength = Len(txtBankAcct1.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtBankAcct2_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBankAcct2.Enter
        On Error GoTo ErrHandler
        Me.txtBankAcct2.SelectionStart = 0
        Me.txtBankAcct2.SelectionLength = Len(txtBankAcct2.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtBankAcct3_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtBankAcct3.Enter
        On Error GoTo ErrHandler
        Me.txtBankAcct3.SelectionStart = 0
        Me.txtBankAcct3.SelectionLength = Len(txtBankAcct3.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtoffPin1_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtOffPin1.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar) ' Only 0-9 can be entered.
        On Error GoTo ErrHandler
        If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8) And (KeyAscii <> 127) Then
            KeyAscii = 0
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
    Private Sub txtCst_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCst.Enter
        On Error GoTo ErrHandler
        Me.txtCst.SelectionStart = 0
        Me.txtCst.SelectionLength = Len(txtCst.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtCurrencyCode_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCurrencyCode.Enter
        On Error GoTo ErrHandler
        Me.txtCurrencyCode.SelectionStart = 0
        Me.txtCurrencyCode.SelectionLength = Len(txtCurrencyCode.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtCurrencyCode_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCurrencyCode.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo errHandler
        If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then Call cmdhelpCurrency_Click(cmdhelpCurrency, New System.EventArgs())
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtEcc_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtEcc.Enter
        On Error GoTo ErrHandler
        Me.txtEcc.SelectionStart = 0
        Me.txtEcc.SelectionLength = Len(txtEcc.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtOffAdd11_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtoffadd11.Enter
        On Error GoTo ErrHandler
        txtoffadd11.SelectionStart = 0
        txtoffadd11.SelectionLength = Len(txtoffadd11.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtOffAdd21_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtOffAdd21.Enter
        On Error GoTo ErrHandler
        txtOffAdd21.SelectionStart = 0
        txtOffAdd21.SelectionLength = Len(txtOffAdd21.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtOffCity1_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtOffCity1.Enter
        On Error GoTo ErrHandler
        txtOffCity1.SelectionStart = 0
        txtOffCity1.SelectionLength = Len(txtOffCity1.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtOffCountry1_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtOffCountry1.Enter
        On Error GoTo ErrHandler
        txtOffCountry1.SelectionStart = 0
        txtOffCountry1.SelectionLength = Len(txtOffCountry1.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtcustname_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCustName.Enter
        On Error GoTo ErrHandler
        txtCustName.SelectionStart = 0
        txtCustName.SelectionLength = Len(txtCustName.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtWebSite_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtWebSite.Enter
        On Error GoTo ErrHandler
        txtWebSite.SelectionStart = 0
        txtWebSite.SelectionLength = Len(txtWebSite.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    '*********************************************'
    'Author         :       Ananya Nath
    'Arguments      :       None
    'Return Value   :       True or False
    'Description    :       Used to validate all the mandatory fields have been properly entered.
    '*********************************************'
    Public Function ValBeforesave() As Boolean ' checks for validity
        On Error GoTo ErrHandler
        Dim strControls As String
        Dim strTag As String
        Dim getYYint As Short
        Dim strFocus1 As System.Windows.Forms.Control = Nothing
        Dim lNo As Integer
        Dim intFirstPos As Short
        Dim intSecondPos As Short
        Dim intSecondPlace As Short
        Dim intDefault As Short
        Dim intctr As Short
        Dim varDefault As Object = Nothing
        Dim varShipcode As Object = Nothing
        Dim varShipDesc As Object = Nothing
        Dim varShipAdd1 As Object = Nothing
        Dim varShipAdd2 As Object = Nothing
        Dim varCity As Object = Nothing
        Dim varCountry As Object = Nothing
        Dim varDelete As Object = Nothing
        Dim varPinCode As Object = Nothing
        Dim intShipDetail As Short

        Dim VAL_SHIP_GSTIN_ID As Object = Nothing

        ValBeforesave = True
        lNo = 1
        strControls = ResolveResString(10059) & vbCrLf
        If Len(Trim(Me.txtCustName.Text)) < 1 Then
            strControls = strControls & vbCrLf & lNo & ". Customer Name." 'Add message to String
            lNo = lNo + 1
            strFocus1 = Me.txtCustName
            ValBeforesave = False
        End If
        If Len(Trim(Me.txtWebSite.Text)) > 0 Then 'validation for WebSite
            If Mid(txtWebSite.Text, 1, 4) <> "www." Then
                strControls = strControls & vbCrLf & lNo & ". WebSite." 'Add message to String
                lNo = lNo + 1
                If strFocus1 Is Nothing Then strFocus1 = Me.txtWebSite
                ValBeforesave = False
            Else
                intSecondPlace = InStr(5, txtWebSite.Text, ".")
                If (Not intSecondPlace > 5) Or (Not Len(txtWebSite.Text) > intSecondPlace) Then
                    strControls = strControls & vbCrLf & lNo & ". WebSite." 'Add message to String
                    lNo = lNo + 1
                    If strFocus1 Is Nothing Then strFocus1 = Me.txtWebSite
                    ValBeforesave = False
                End If
            End If
        End If
        If Len(Trim(txtoffadd11.Text)) = 0 And Len(Trim(txtOffAdd21.Text)) = 0 Then
            strControls = strControls & vbCrLf & lNo & ". Office Address (General Details - Office Address)."
            lNo = lNo + 1
            tabDetails.SelectedIndex = 0
            tabCustomer.SelectedIndex = 0
            If strFocus1 Is Nothing Then strFocus1 = txtoffadd11
            ValBeforesave = False
        End If
        If Len(Trim(txtOffCity1.Text)) = 0 Then
            strControls = strControls & vbCrLf & lNo & ". City (General Details - Office Address)."
            lNo = lNo + 1
            tabDetails.SelectedIndex = 0
            tabCustomer.SelectedIndex = 0
            If strFocus1 Is Nothing Then strFocus1 = txtOffCity1
            ValBeforesave = False
        End If
        If Len(Trim(txtOffCountry1.Text)) = 0 Then
            strControls = strControls & vbCrLf & lNo & ". Country (General Details - Office Address)."
            lNo = lNo + 1
            tabDetails.SelectedIndex = 0
            tabCustomer.SelectedIndex = 0
            If strFocus1 Is Nothing Then strFocus1 = txtOffCountry1
            ValBeforesave = False
        End If
        If Len(Trim(txtOffPin1.Text)) = 0 Then
            strControls = strControls & vbCrLf & lNo & ". Pin (General Details - Office Address)."
            lNo = lNo + 1
            tabDetails.SelectedIndex = 0
            tabCustomer.SelectedIndex = 0
            If strFocus1 Is Nothing Then strFocus1 = txtOffPin1
            ValBeforesave = False
        End If
        If Len(Trim(txtEmail1.Text)) > 0 Then
            intFirstPos = InStr(1, txtEmail1.Text, "@")
            intSecondPos = InStr(1, txtEmail1.Text, ".")
            If (intFirstPos = 0 Or intSecondPos = 0) Or (intFirstPos = 1) Or (intSecondPos - intFirstPos = 1) Or (Not Len(txtEmail1.Text) > intSecondPos) Then
                strControls = strControls & vbCrLf & lNo & ". Email (General Details - Office Address)."
                lNo = lNo + 1
                If strFocus1 Is Nothing Then strFocus1 = txtEmail1
                ValBeforesave = False
            End If
        End If
        If Len(Trim(txtCreditTermId.Text)) = 0 Then
            strControls = strControls & vbCrLf & lNo & ". Credit Term ID (Marketing Details)."
            lNo = lNo + 1
            If strFocus1 Is Nothing Then strFocus1 = txtCreditTermId
            ValBeforesave = False
        End If
        If Len(Trim(txtCurrencyCode.Text)) = 0 Or (ValidateCur()) = False Then
            strControls = strControls & vbCrLf & lNo & ". Currency Code (Purchase Details)."
            lNo = lNo + 1
            If strFocus1 Is Nothing Then strFocus1 = txtCurrencyCode
            ValBeforesave = False
        End If
        If chkJobWork.CheckState = 0 And chkStandard.CheckState = 0 Then
            strControls = strControls & vbCrLf & lNo & ". Customer Type (Marketing Details)."
            lNo = lNo + 1
            If strFocus1 Is Nothing Then strFocus1 = chkJobWork
            ValBeforesave = False
        End If
        If Len(Trim(txtAcctLedger.Text)) > 0 Then
            mobjEmpDll.CConnection.OpenConnection(gstrDSNName, gstrDatabaseName)
            mrsEmpDll.OpenRecordset("select glM_glCode,glM_desc,glM_grpCode,glM_transTag,glM_stDate,glM_endDate,glM_slTag,glM_jvAllowed,glM_controlaccount,glM_costcentreflag,glM_projectflag,glM_empflag,glM_GroupID,glM_remarks,glm_GLType from fin_glmaster Where Unit_Code='" & gstrUNITID & "' And glm_glcode = '" & txtAcctLedger.Text & "' and glm_transtag = '1' ", ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockReadOnly)
            mrsEmpDll.Filter_Renamed = " glm_glcode ='" & txtAcctLedger.Text & "'"
            If mrsEmpDll.Recordcount <= 0 Then
                strControls = strControls & vbCrLf & lNo & ". Ledger (Account Details)."
                lNo = lNo + 1
                If strFocus1 Is Nothing Then strFocus1 = txtAcctLedger
                ValBeforesave = False
            End If
            mrsEmpDll.CloseRecordset()
            mobjEmpDll.CConnection.CloseConnection()
        ElseIf Len(Trim(txtAcctLedger.Text)) > 0 Then
            mobjEmpDll.CConnection.OpenConnection(gstrDSNName, gstrDatabaseName)
            mrsEmpDll.OpenRecordset("select glm_glcode from fin_glmaster Where Unit_Code='" & gstrUNITID & "' And glm_glcode = '" & txtAcctLedger.Text & "'", ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            If Not mrsEmpDll.Recordcount > 0 Then
                strControls = strControls & vbCrLf & lNo & ". SubLedger (Account Details)."
                lNo = lNo + 1
                If strFocus1 Is Nothing Then strFocus1 = txtAccSubLedger
                ValBeforesave = False
            End If
            mrsEmpDll.CloseRecordset()
            mobjEmpDll.CConnection.CloseConnection()
        End If
        If ChkCSI_GT.CheckState = 1 And Len(txtAcctLedgerExc.Text) = 0 Then
            strControls = strControls & vbCrLf & lNo & ". Excise on Cust Supp Matl GL."
            lNo = lNo + 1
            If strFocus1 Is Nothing Then strFocus1 = txtAcctLedgerExc
            ValBeforesave = False
        End If
        If ChkCSI_GT.CheckState = 1 And Len(txtAccSubLedgerExc.Text) = 0 Then
            strControls = strControls & vbCrLf & lNo & ". Excise on Cust Supp Matl SL."
            lNo = lNo + 1
            If strFocus1 Is Nothing Then strFocus1 = txtAccSubLedgerExc
            ValBeforesave = False
        End If
        intShipDetail = 0
        For intctr = 1 To spShippingAddess.MaxRows
            varDelete = Nothing
            Call spShippingAddess.GetText(enmshipdetail.VAL_DELETE, intctr, varDelete)
            If UCase(varDelete) <> "D" Then
                varShipcode = Nothing
                Call spShippingAddess.GetText(enmshipdetail.VAL_SHIPCODE, intctr, varShipcode)
                varShipDesc = Nothing
                Call spShippingAddess.GetText(enmshipdetail.VAL_SHIPDESC, intctr, varShipDesc)
                varShipAdd1 = Nothing
                Call spShippingAddess.GetText(enmshipdetail.VAL_SHIPADD1, intctr, varShipAdd1)
                varShipAdd2 = Nothing
                Call spShippingAddess.GetText(enmshipdetail.VAL_SHIPADD2, intctr, varShipAdd2)
                varCity = Nothing
                Call spShippingAddess.GetText(enmshipdetail.VAL_CITY, intctr, varCity)
                varCountry = Nothing
                Call spShippingAddess.GetText(enmshipdetail.VAL_COUNTRY, intctr, varCountry)
                varPinCode = Nothing
                Call spShippingAddess.GetText(enmshipdetail.VAL_Pin, intctr, varPinCode)

                VAL_SHIP_GSTIN_ID = Nothing 'by abhijit on 17-Aug-2017
                Call spShippingAddess.GetText(enmshipdetail.VAL_SHIP_GSTIN_ID, intctr, VAL_SHIP_GSTIN_ID)

               

                If gblnGSTUnit = True And chkGSTINnotrequired.Checked = False Then
                    If Len(Trim(varShipcode)) > 0 And Len(Trim(varShipDesc)) > 0 And Len(Trim(varShipAdd1)) > 0 And Len(Trim(varShipAdd2)) > 0 And Len(Trim(varCity)) > 0 And Len(Trim(varCountry)) > 0 And Len(Trim(varPinCode)) > 0 And Len(Trim(VAL_SHIP_GSTIN_ID)) > 0 Then
                        intShipDetail = intShipDetail + 1
                    End If
                Else
                    If Len(Trim(varShipcode)) > 0 And Len(Trim(varShipDesc)) > 0 And Len(Trim(varShipAdd1)) > 0 And Len(Trim(varShipAdd2)) > 0 And Len(Trim(varCity)) > 0 And Len(Trim(varCountry)) > 0 And Len(Trim(varPinCode)) > 0 Then
                        intShipDetail = intShipDetail + 1
                    End If
                End If

            End If
        Next intctr
        If intShipDetail = 0 Then
            strControls = strControls & vbCrLf & lNo & ". Shipping Address Details."
            lNo = lNo + 1
            If strFocus1 Is Nothing Then strFocus1 = spShippingAddess
            ValBeforesave = False
        End If
        If intShipDetail > 0 Then
            intDefault = 0
            For intctr = 1 To spShippingAddess.MaxRows
                varDelete = Nothing
                Call spShippingAddess.GetText(enmshipdetail.VAL_DELETE, intctr, varDelete)
                If UCase(varDelete) <> "D" Then
                    varDefault = Nothing
                    Call spShippingAddess.GetText(enmshipdetail.VAL_DEFAULT, intctr, varDefault)
                    varShipcode = Nothing
                    Call spShippingAddess.GetText(enmshipdetail.VAL_SHIPCODE, intctr, varShipcode)
                    varShipDesc = Nothing
                    Call spShippingAddess.GetText(enmshipdetail.VAL_SHIPDESC, intctr, varShipDesc)
                    varShipAdd1 = Nothing
                    Call spShippingAddess.GetText(enmshipdetail.VAL_SHIPADD1, intctr, varShipAdd1)
                    varShipAdd2 = Nothing
                    Call spShippingAddess.GetText(enmshipdetail.VAL_SHIPADD2, intctr, varShipAdd2)
                    varCity = Nothing
                    Call spShippingAddess.GetText(enmshipdetail.VAL_CITY, intctr, varCity)
                    varCountry = Nothing
                    Call spShippingAddess.GetText(enmshipdetail.VAL_COUNTRY, intctr, varCountry)
                    varPinCode = Nothing
                    Call spShippingAddess.GetText(enmshipdetail.VAL_Pin, intctr, varPinCode)

                    VAL_SHIP_GSTIN_ID = Nothing 'by abhijit 17-AUG-2017
                    Call spShippingAddess.GetText(enmshipdetail.VAL_SHIP_GSTIN_ID, intctr, VAL_SHIP_GSTIN_ID)

                    If gblnGSTUnit = True And chkGSTINnotrequired.Checked = False Then
                        If Len(Trim(varShipcode)) > 0 And Len(Trim(varShipDesc)) > 0 And Len(Trim(varShipAdd1)) > 0 And Len(Trim(varShipAdd2)) > 0 And Len(Trim(varCity)) > 0 And Len(Trim(varCountry)) > 0 And Len(Trim(varPinCode)) > 0 And Len(Trim(VAL_SHIP_GSTIN_ID)) > 0 Then
                            If Val(varDefault) = 1 Then intDefault = intDefault + 1
                        End If
                    Else
                        If Len(Trim(varShipcode)) > 0 And Len(Trim(varShipDesc)) > 0 And Len(Trim(varShipAdd1)) > 0 And Len(Trim(varShipAdd2)) > 0 And Len(Trim(varCity)) > 0 And Len(Trim(varCountry)) > 0 And Len(Trim(varPinCode)) > 0 Then
                            If Val(varDefault) = 1 Then intDefault = intDefault + 1
                        End If
                    End If


                End If
            Next intctr
            If Chkbarcodeprinting.CheckState = CheckState.Checked And CmbPrintMethod.Text.Trim.Length = 0 Then
                strControls = strControls & vbCrLf & lNo & ". Barcode Printing Method."
                lNo = lNo + 1
                If strFocus1 Is Nothing Then strFocus1 = CmbPrintMethod
                ValBeforesave = False
            End If
            If intDefault = 0 Then
                strControls = strControls & vbCrLf & lNo & ". Default Shipping Address"
                lNo = lNo + 1
                If strFocus1 Is Nothing Then strFocus1 = spShippingAddess
                ValBeforesave = False
            End If
        End If



        Dim strSQL As String = String.Empty
        strSQL = "select dbo.Check_SubLedger_For_CustMst( '" & gstrUNITID & "','" & txtCustCode.Text.Trim & "','" & txtAccSubLedger.Text.Trim & "' )"
        If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSQL)) = True Then
            strControls = strControls & vbCrLf & lNo & ". Account SubLedger"
            lNo = lNo + 1
            If strFocus1 Is Nothing Then strFocus1 = txtAccSubLedger
            ValBeforesave = False
        End If

        'GST CHANGES
        If Len(Trim(txtKeyContactEmail_1.Text)) > 0 Then
            intFirstPos = InStr(1, Trim(txtKeyContactEmail_1.Text), "@") ' find out the position of @ in the string
            intSecondPos = InStr(1, Trim(txtKeyContactEmail_1.Text), ".") ' Find out the position of . in the string
            If (intFirstPos = 0 Or intSecondPos = 0) Or (intFirstPos = 1) Or (intSecondPos - intFirstPos = 1) Or (Not Len(txtKeyContactEmail_1.Text) > intSecondPos) Then
                strControls = strControls & vbCrLf & lNo & ". Email Id. (GST Details - Key Contact for Mismatch)."
                lNo = lNo + 1
                If strFocus1 Is Nothing Then strFocus1 = txtKeyContactEmail_1
                ValBeforesave = False
            End If
        End If

        If Len(Trim(txtKeyContactEmail_2.Text)) > 0 Then
            intFirstPos = InStr(1, Trim(txtKeyContactEmail_2.Text), "@") ' find out the position of @ in the string
            intSecondPos = InStr(1, Trim(txtKeyContactEmail_2.Text), ".") ' Find out the position of . in the string
            If (intFirstPos = 0 Or intSecondPos = 0) Or (intFirstPos = 1) Or (intSecondPos - intFirstPos = 1) Or (Not Len(txtKeyContactEmail_2.Text) > intSecondPos) Then
                strControls = strControls & vbCrLf & lNo & ". Email Id. (GST Details - Key Contact for Mismatch-Escalation Level)."
                lNo = lNo + 1
                If strFocus1 Is Nothing Then strFocus1 = txtKeyContactEmail_2
                ValBeforesave = False
            End If
        End If
        'GST CHANGES
        If gblnGSTUnit Then
            If chkGSTINnotrequired.Checked = False Then
                strSQL = "SELECT TOP 1 1 FROM CUSTOMER_MST WHERE UNIT_CODE='" & gstrUnitId & "' AND CUSTOMER_CODE='" & txtCustCode.Text & "' AND CLASSIFICATIONTYPE ='B' AND CUST_TYPE<>'O'"
                If IsRecordExists(strSQL) Then
                    If Len(Trim(txtGSTINId.Text)) > 0 And Len(Trim(txtGSTINId.Text)) <> 15 Then
                        strControls = strControls & vbCrLf & lNo & ".GSTIN ID must be of 15 characters."
                        lNo = lNo + 1
                        If strFocus1 Is Nothing Then strFocus1 = txtGSTINId
                        ValBeforesave = False
                    ElseIf Len(Trim(txtGSTINId.Text)) = 0 Then
                        strControls = strControls & vbCrLf & lNo & ".GSTIN ID is manadatory."
                        lNo = lNo + 1
                        If strFocus1 Is Nothing Then strFocus1 = txtGSTINId
                        ValBeforesave = False
                    End If
                End If
            End If
        End If

        'Samiksha Credit limit changes
        If isCreditLimitMandatory = True Then
            If txtboxCreditLimit.Text.Length <= 0 Then
                strControls = strControls & vbCrLf & lNo & ".Credit Limit is mandatory"
                If strFocus1 Is Nothing Then strFocus1 = txtboxCreditLimit
                lNo = lNo + 1
                ValBeforesave = False
            ElseIf (IsNumeric(txtboxCreditLimit.Text) = False) Then
                strControls = strControls & vbCrLf & lNo & ".Credit Limit must be Numeric"
                If strFocus1 Is Nothing Then strFocus1 = txtboxCreditLimit
                lNo = lNo + 1
                ValBeforesave = False
            ElseIf (Convert.ToDouble(txtboxCreditLimit.Text) <= 0) Then
                strControls = strControls & vbCrLf & lNo & ".Credit Limit must be greater than zero"
                If strFocus1 Is Nothing Then strFocus1 = txtboxCreditLimit
                lNo = lNo + 1
                ValBeforesave = False
            End If
        End If


        If ValBeforesave = False Then 'If any invalid field is there than set the focus on that field(after displaying message).                  MsgBox strControls, vbInformation, "eMPro"
            MsgBox(strControls, MsgBoxStyle.Information, "eMPro")
            strFocus1.Focus()
        End If
        strFocus1 = Nothing
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        mrsEmpDll.CloseRecordset()
        mobjEmpDll.CConnection.CloseConnection()
    End Function
    '*********************************************'
    'Author         :       Ananya Nath
    'Arguments      :       None
    'Return Value   :       None
    'Description    :       This Function is Used to display corresponding values, if user enters/selects an existing CustomerCode.
    '*********************************************'
    'Changes by samiksha : Customer Master Authorization
    'Samiksha Credit limit changes
    'Samiksha branchcode changes
    Public Function display() As Object ' Used to display the record.
        On Error GoTo ErrHandler
        Dim strdetails As String
        Dim strsql As String
        Dim strcheck As String
        Dim strmilk As Boolean
        Dim rsLocation As New ClsResultSetDB
        Dim strCust As String = ""
        Dim strcustClassification As String = ""
        Dim isNewCustomer As Boolean = False
        If txtCustCode.Tag <> "" Then txtCustCode.Text = txtCustCode.Tag
        '10690771
        If Me.txtCustCode.Text <> "" Then

            isNewCustomer = checkifnewcustomer()
            mobjEmpDll.CConnection.OpenConnection(gstrDSNName, gstrDatabaseName)
            'GST CHANGES - NEW COLUMN ADDED
            If isNewCustomer = False Then
                mobjEmpDll.CRecordset.OpenRecordset("select Customer_Code,Account_ledger,Account_subLedger,Cust_Name,Cust_Vendor_Code,Shipping_Duration,Excise_range,Commisionrate,Division,ECC_Code,LST_No,CST_No,Cust_Location,Ship_Contact_person,Ship_Person_desig,Ship_email_id,Ship_Address1,Ship_Address2,Ship_City,Ship_State,Ship_Pin,Ship_dist,Ship_Country,Ship_Phone,Ship_Fax,bill_Contact_person,Bill_Person_desig,bill_email_id,Bill_Address1,Bill_Address2,Bill_City,Bill_State,Bill_Pin,Bill_dist,Bill_Country,Bill_Phone,Bill_Fax,Office_Contact_person,Office_person_design,office_email_id,Office_Address1,Office_Address2,Office_City,Office_State,Office_dist,Office_Pin,Office_Country,Office_Phone,Office_fax,Website,customer_type,currency_code,Bank_ac1,Bank_ac2,Bank_ac3,Credit_days,Has_Trans_flag,ScheduleCode,CST_Eval,tin_no,Milk_Van,AllowExcessSchedule,CUST_ALIAS,DOCK_CODE,ShipmentThruWh,PANNo,ServiceTaxNo,SwiftCode,IBANNo_SOT,BAccNo,BankName,BankAdd1,Tool_GL,Tool_SL,ASN_REQD,CST_amt,CSIEX_Inc,CSIEX_GL,CSIEX_SL,Group_Customer,AllowBarCodePrinting,CONSIGNEEWISELOC,AllowASNPrinting,AllowASNTextGeneration,IS_AEROBIN_CUST,Customer_EDICode,Plant_Code,ASNFunctionCode,INVOICEAGAINSTAGREEMENTMST,CUST_TYPE,GLOBAL_CUATOMER_CODE,FordLabelReqd, KAMCODE,PRINT_METHOD,CT2_REQD,GSTIN_Id,GST_BILLSTATE_CODE,GST_BILLSTATE_DESC,ContactForMismatch_1,ContactForMismatch_PhoneNo_1,ContactForMismatch_MobileNo_1,ContactForMismatch_Email_1,ContactForMismatch_2,ContactForMismatch_PhoneNo_2,ContactForMismatch_MobileNo_2,ContactForMismatch_Email_2,classificationtype,GSTIN_notrequired,ISNULL(PARTY_TRN_NO,'') PARTY_TRN_NO,SEZ,CreditLimit,branchcode  from customer_mst Where Unit_Code='" & gstrUNITID & "' And Customer_code = '" & txtCustCode.Text & "'", ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                rsLocation.GetResult("select Customer_Code,Account_ledger,Account_subLedger,Cust_Name," &
                " Cust_Vendor_Code,Excise_range,Commisionrate,Division,ECC_Code,LST_No,CST_No,Cust_Location," &
                " Ship_Contact_person,Ship_Person_desig,Ship_email_id,Ship_Address1,Ship_Address2,Ship_City," &
                " Ship_State,Ship_Pin,Ship_dist,Ship_Country,Ship_Phone,Ship_Fax,bill_Contact_person," &
                " Bill_Person_desig,bill_email_id,Bill_Address1,Bill_Address2,Bill_City,Bill_State,Bill_Pin," &
                " Bill_dist,Bill_Country,Bill_Phone,Bill_Fax,Office_Contact_person,Office_person_design," &
                " office_email_id,Office_Address1,Office_Address2,Office_City,Office_State,Office_dist," &
                " Office_Pin,Office_Country,Office_Phone,Office_fax,Website,customer_type,currency_code," &
                " Bank_ac1,Bank_ac2,Bank_ac3,Credit_days,Has_Trans_flag,ScheduleCode,CST_Eval,tin_no,Milk_Van," &
                " AllowExcessSchedule,CUST_ALIAS,DOCK_CODE,ShipmentThruWh,PANNo,ServiceTaxNo,SwiftCode," &
                " IBANNo_SOT,BAccNo,BankName,BankAdd1,Tool_GL,Tool_SL,ASN_REQD,CST_amt,CSIEX_Inc,CSIEX_GL," &
                " CSIEX_SL,Group_Customer,AllowBarCodePrinting,CONSIGNEEWISELOC,AllowASNPrinting," &
                " AllowASNTextGeneration,IS_AEROBIN_CUST,Customer_EDICode,Plant_Code,ASNFunctionCode," &
                " INVOICEAGAINSTAGREEMENTMST,CUST_TYPE,GLOBAL_CUATOMER_CODE,FordLabelReqd, KAMCODE,CT2_REQD,GSTIN_Id,GST_BILLSTATE_CODE,GST_BILLSTATE_DESC,ContactForMismatch_1,ContactForMismatch_PhoneNo_1,ContactForMismatch_MobileNo_1,ContactForMismatch_Email_1,ContactForMismatch_2,ContactForMismatch_PhoneNo_2,ContactForMismatch_MobileNo_2,ContactForMismatch_Email_2,classificationtype,GSTIN_notrequired,SEZ,ISNULL(PARTY_TRN_NO,'') PARTY_TRN_NO,CreditLimit,branchcode from customer_mst" &
                " Where Unit_Code='" & gstrUNITID & "' And Customer_code = '" & txtCustCode.Text & "'")
                '" & union select Customer_Code,Account_ledger,Account_subLedger,Cust_Name," &
                '" Cust_Vendor_Code,Excise_range,Commisionrate,Division,ECC_Code,LST_No,CST_No,Cust_Location," &
                '" Ship_Contact_person,Ship_Person_desig,Ship_email_id,Ship_Address1,Ship_Address2,Ship_City," &
                '" Ship_State,Ship_Pin,Ship_dist,Ship_Country,Ship_Phone,Ship_Fax,bill_Contact_person," &
                '" Bill_Person_desig,bill_email_id,Bill_Address1,Bill_Address2,Bill_City,Bill_State,Bill_Pin," &
                '" Bill_dist,Bill_Country,Bill_Phone,Bill_Fax,Office_Contact_person,Office_person_design," &
                '" office_email_id,Office_Address1,Office_Address2,Office_City,Office_State,Office_dist," &
                '" Office_Pin,Office_Country,Office_Phone,Office_fax,Website,customer_type,currency_code," &
                '" Bank_ac1,Bank_ac2,Bank_ac3,Credit_days,Has_Trans_flag,ScheduleCode,CST_Eval,tin_no,Milk_Van," &
                '" AllowExcessSchedule,CUST_ALIAS,DOCK_CODE,ShipmentThruWh,PANNo,ServiceTaxNo,SwiftCode," &
                '" IBANNo_SOT,BAccNo,BankName,BankAdd1,Tool_GL,Tool_SL,ASN_REQD,CST_amt,CSIEX_Inc,CSIEX_GL," &
                '" CSIEX_SL,Group_Customer,AllowBarCodePrinting,CONSIGNEEWISELOC,AllowASNPrinting," &
                '" AllowASNTextGeneration,IS_AEROBIN_CUST,Customer_EDICode,Plant_Code,ASNFunctionCode," &
                '" INVOICEAGAINSTAGREEMENTMST,CUST_TYPE,GLOBAL_CUATOMER_CODE,FordLabelReqd, KAMCODE,CT2_REQD,GSTIN_Id,GST_BILLSTATE_CODE,GST_BILLSTATE_DESC,ContactForMismatch_1,ContactForMismatch_PhoneNo_1,ContactForMismatch_MobileNo_1,ContactForMismatch_Email_1,ContactForMismatch_2,ContactForMismatch_PhoneNo_2,ContactForMismatch_MobileNo_2,ContactForMismatch_Email_2,classificationtype,GSTIN_notrequired,SEZ,ISNULL(PARTY_TRN_NO,'') PARTY_TRN_NO,CreditLimit,branchcode from customer_mst_authorization" &
                '" Where Unit_Code='" & gstrUNITID & "' And Customer_code = '" & txtCustCode.Text & "' and (ISNULL(authorization_status,'')='Return'  Or ISNULL(authorization_status,'')='')"
                ') '10688280 -Add KAM Code in Customer Master.
            Else
                mobjEmpDll.CRecordset.OpenRecordset("select Customer_Code,Account_ledger,Account_subLedger,Cust_Name,Cust_Vendor_Code,Shipping_Duration,Excise_range,Commisionrate,Division,ECC_Code,LST_No,CST_No,Cust_Location,Ship_Contact_person,Ship_Person_desig,Ship_email_id,Ship_Address1,Ship_Address2,Ship_City,Ship_State,Ship_Pin,Ship_dist,Ship_Country,Ship_Phone,Ship_Fax,bill_Contact_person,Bill_Person_desig,bill_email_id,Bill_Address1,Bill_Address2,Bill_City,Bill_State,Bill_Pin,Bill_dist,Bill_Country,Bill_Phone,Bill_Fax,Office_Contact_person,Office_person_design,office_email_id,Office_Address1,Office_Address2,Office_City,Office_State,Office_dist,Office_Pin,Office_Country,Office_Phone,Office_fax,Website,customer_type,currency_code,Bank_ac1,Bank_ac2,Bank_ac3,Credit_days,Has_Trans_flag,ScheduleCode,CST_Eval,tin_no,Milk_Van,AllowExcessSchedule,CUST_ALIAS,DOCK_CODE,ShipmentThruWh,PANNo,ServiceTaxNo,SwiftCode,IBANNo_SOT,BAccNo,BankName,BankAdd1,Tool_GL,Tool_SL,ASN_REQD,CST_amt,CSIEX_Inc,CSIEX_GL,CSIEX_SL,Group_Customer,AllowBarCodePrinting,CONSIGNEEWISELOC,AllowASNPrinting,AllowASNTextGeneration,IS_AEROBIN_CUST,Customer_EDICode,Plant_Code,ASNFunctionCode,INVOICEAGAINSTAGREEMENTMST,CUST_TYPE,GLOBAL_CUATOMER_CODE,FordLabelReqd, KAMCODE,PRINT_METHOD,CT2_REQD,GSTIN_Id,GST_BILLSTATE_CODE,GST_BILLSTATE_DESC,ContactForMismatch_1,ContactForMismatch_PhoneNo_1,ContactForMismatch_MobileNo_1,ContactForMismatch_Email_1,ContactForMismatch_2,ContactForMismatch_PhoneNo_2,ContactForMismatch_MobileNo_2,ContactForMismatch_Email_2,classificationtype,GSTIN_notrequired,ISNULL(PARTY_TRN_NO,'') PARTY_TRN_NO,SEZ,CreditLimit,branchcode  from customer_mst_authorization Where Unit_Code='" & gstrUNITID & "' And Customer_code = '" & txtCustCode.Text & "'", ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
                rsLocation.GetResult("select Customer_Code,Account_ledger,Account_subLedger,Cust_Name," &
                " Cust_Vendor_Code,Excise_range,Commisionrate,Division,ECC_Code,LST_No,CST_No,Cust_Location," &
                " Ship_Contact_person,Ship_Person_desig,Ship_email_id,Ship_Address1,Ship_Address2,Ship_City," &
                " Ship_State,Ship_Pin,Ship_dist,Ship_Country,Ship_Phone,Ship_Fax,bill_Contact_person," &
                " Bill_Person_desig,bill_email_id,Bill_Address1,Bill_Address2,Bill_City,Bill_State,Bill_Pin," &
                " Bill_dist,Bill_Country,Bill_Phone,Bill_Fax,Office_Contact_person,Office_person_design," &
                " office_email_id,Office_Address1,Office_Address2,Office_City,Office_State,Office_dist," &
                " Office_Pin,Office_Country,Office_Phone,Office_fax,Website,customer_type,currency_code," &
                " Bank_ac1,Bank_ac2,Bank_ac3,Credit_days,Has_Trans_flag,ScheduleCode,CST_Eval,tin_no,Milk_Van," &
                " AllowExcessSchedule,CUST_ALIAS,DOCK_CODE,ShipmentThruWh,PANNo,ServiceTaxNo,SwiftCode," &
                " IBANNo_SOT,BAccNo,BankName,BankAdd1,Tool_GL,Tool_SL,ASN_REQD,CST_amt,CSIEX_Inc,CSIEX_GL," &
                " CSIEX_SL,Group_Customer,AllowBarCodePrinting,CONSIGNEEWISELOC,AllowASNPrinting," &
                " AllowASNTextGeneration,IS_AEROBIN_CUST,Customer_EDICode,Plant_Code,ASNFunctionCode," &
                " INVOICEAGAINSTAGREEMENTMST,CUST_TYPE,GLOBAL_CUATOMER_CODE,FordLabelReqd, KAMCODE,CT2_REQD,GSTIN_Id,GST_BILLSTATE_CODE,GST_BILLSTATE_DESC,ContactForMismatch_1,ContactForMismatch_PhoneNo_1,ContactForMismatch_MobileNo_1,ContactForMismatch_Email_1,ContactForMismatch_2,ContactForMismatch_PhoneNo_2,ContactForMismatch_MobileNo_2,ContactForMismatch_Email_2,classificationtype,GSTIN_notrequired,SEZ,ISNULL(PARTY_TRN_NO,'') PARTY_TRN_NO,CreditLimit,branchcode from customer_mst_authorization" &
                " Where Unit_Code='" & gstrUNITID & "' And Customer_code = '" & txtCustCode.Text & "' and ( ISNULL(authorization_status,'')='' or ISNULL(authorization_status,'')='Return')")
                '" & union select Customer_Code,Account_ledger,Account_subLedger,Cust_Name," &
                '" Cust_Vendor_Code,Excise_range,Commisionrate,Division,ECC_Code,LST_No,CST_No,Cust_Location," &
                '" Ship_Contact_person,Ship_Person_desig,Ship_email_id,Ship_Address1,Ship_Address2,Ship_City," &
                '" Ship_State,Ship_Pin,Ship_dist,Ship_Country,Ship_Phone,Ship_Fax,bill_Contact_person," &
                '" Bill_Person_desig,bill_email_id,Bill_Address1,Bill_Address2,Bill_City,Bill_State,Bill_Pin," &
                '" Bill_dist,Bill_Country,Bill_Phone,Bill_Fax,Office_Contact_person,Office_person_design," &
                '" office_email_id,Office_Address1,Office_Address2,Office_City,Office_State,Office_dist," &
                '" Office_Pin,Office_Country,Office_Phone,Office_fax,Website,customer_type,currency_code," &
                '" Bank_ac1,Bank_ac2,Bank_ac3,Credit_days,Has_Trans_flag,ScheduleCode,CST_Eval,tin_no,Milk_Van," &
                '" AllowExcessSchedule,CUST_ALIAS,DOCK_CODE,ShipmentThruWh,PANNo,ServiceTaxNo,SwiftCode," &
                '" IBANNo_SOT,BAccNo,BankName,BankAdd1,Tool_GL,Tool_SL,ASN_REQD,CST_amt,CSIEX_Inc,CSIEX_GL," &
                '" CSIEX_SL,Group_Customer,AllowBarCodePrinting,CONSIGNEEWISELOC,AllowASNPrinting," &
                '" AllowASNTextGeneration,IS_AEROBIN_CUST,Customer_EDICode,Plant_Code,ASNFunctionCode," &
                '" INVOICEAGAINSTAGREEMENTMST,CUST_TYPE,GLOBAL_CUATOMER_CODE,FordLabelReqd, KAMCODE,CT2_REQD,GSTIN_Id,GST_BILLSTATE_CODE,GST_BILLSTATE_DESC,ContactForMismatch_1,ContactForMismatch_PhoneNo_1,ContactForMismatch_MobileNo_1,ContactForMismatch_Email_1,ContactForMismatch_2,ContactForMismatch_PhoneNo_2,ContactForMismatch_MobileNo_2,ContactForMismatch_Email_2,classificationtype,GSTIN_notrequired,SEZ,ISNULL(PARTY_TRN_NO,'') PARTY_TRN_NO,CreditLimit,branchcode from customer_mst_authorization" &
                '" Where Unit_Code='" & gstrUNITID & "' And Customer_code = '" & txtCustCode.Text & "' and (ISNULL(authorization_status,'')='Return'  Or ISNULL(authorization_status,'')='')"
                ') '10688280 -Add KAM Code in Customer Master.
            End If


            If Not mobjEmpDll.CRecordset.EOF_Renamed Then
                mobjEmpDll.CRecordset.MoveFirst()
                lblGlobalCustCodeDesc.Text = mobjEmpDll.CRecordset.GetFieldValue("GLOBAL_CUATOMER_CODE", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtCustCode.Text = mobjEmpDll.CRecordset.GetFieldValue("customer_code", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtCustName.Text = mobjEmpDll.CRecordset.GetFieldValue("cust_name", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtCustLoc.Text = rsLocation.GetValue("Cust_location")
                txtWebSite.Text = mobjEmpDll.CRecordset.GetFieldValue("website", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtAcctLedger.Text = mobjEmpDll.CRecordset.GetFieldValue("account_ledger", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtAccSubLedger.Text = mobjEmpDll.CRecordset.GetFieldValue("account_subledger", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtoffadd11.Text = mobjEmpDll.CRecordset.GetFieldValue("office_address1", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtOffAdd21.Text = mobjEmpDll.CRecordset.GetFieldValue("office_address2", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtOffCity1.Text = mobjEmpDll.CRecordset.GetFieldValue("office_city", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtOffDist1.Text = mobjEmpDll.CRecordset.GetFieldValue("office_dist", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtOffPin1.Text = mobjEmpDll.CRecordset.GetFieldValue("office_pin", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtOffState1.Text = mobjEmpDll.CRecordset.GetFieldValue("office_state", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtOffCountry1.Text = mobjEmpDll.CRecordset.GetFieldValue("office_country", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtOffPhone1.Text = mobjEmpDll.CRecordset.GetFieldValue("office_phone", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtOffFax1.Text = mobjEmpDll.CRecordset.GetFieldValue("office_fax", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtContPer1.Text = mobjEmpDll.CRecordset.GetFieldValue("office_contact_person", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtDesig1.Text = mobjEmpDll.CRecordset.GetFieldValue("office_person_design", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtEmail1.Text = mobjEmpDll.CRecordset.GetFieldValue("office_email_id", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtCustvendCode.Text = mobjEmpDll.CRecordset.GetFieldValue("cust_vendor_code", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtShipDur.Text = mobjEmpDll.CRecordset.GetFieldValue("Shipping_Duration", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                '**********
                txtBilladd1.Text = mobjEmpDll.CRecordset.GetFieldValue("Bill_address1", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtBillAdd2.Text = mobjEmpDll.CRecordset.GetFieldValue("Bill_address2", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtbillCity(0).Text = mobjEmpDll.CRecordset.GetFieldValue("Bill_city", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtBillDist.Text = mobjEmpDll.CRecordset.GetFieldValue("Bill_dist", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtBillPin.Text = mobjEmpDll.CRecordset.GetFieldValue("Bill_pin", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtBillState.Text = mobjEmpDll.CRecordset.GetFieldValue("Bill_state", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtBillCountry.Text = mobjEmpDll.CRecordset.GetFieldValue("Bill_country", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtBillPhone.Text = mobjEmpDll.CRecordset.GetFieldValue("Bill_phone", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtBillFax.Text = mobjEmpDll.CRecordset.GetFieldValue("Bill_fax", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtBillContPer.Text = mobjEmpDll.CRecordset.GetFieldValue("Bill_contact_person", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtBillDesig.Text = mobjEmpDll.CRecordset.GetFieldValue("Bill_person_desig", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtBillEmail.Text = mobjEmpDll.CRecordset.GetFieldValue("Bill_email_id", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                '*************
                txtBankAcct1.Text = mobjEmpDll.CRecordset.GetFieldValue("BANK_AC1", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtBankAcct2.Text = mobjEmpDll.CRecordset.GetFieldValue("BANK_AC2", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtBankAcct3.Text = mobjEmpDll.CRecordset.GetFieldValue("BANK_AC3", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                '********************
                'Purchase Details'
                txtCurrencyCode.Text = mobjEmpDll.CRecordset.GetFieldValue("currency_code", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtCreditTermId.Text = mobjEmpDll.CRecordset.GetFieldValue("credit_days", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                strcheck = mobjEmpDll.CRecordset.GetFieldValue("Customer_type", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                strmilk = mobjEmpDll.CRecordset.GetFieldValue("milk_van", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)

                If strmilk = True Then
                    Me.chkmilkvan.CheckState = System.Windows.Forms.CheckState.Checked
                Else
                    Me.chkmilkvan.CheckState = System.Windows.Forms.CheckState.Unchecked
                End If

                TxtPanNo.Text = mobjEmpDll.CRecordset.GetFieldValue("PANNo", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                TxtServiceTaxNo.Text = mobjEmpDll.CRecordset.GetFieldValue("ServiceTaxNo", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                TxtSwiftNo.Text = mobjEmpDll.CRecordset.GetFieldValue("SwiftCode", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                TxtIBANNo.Text = mobjEmpDll.CRecordset.GetFieldValue("IBANNo_SOT", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                TxtBankname.Text = mobjEmpDll.CRecordset.GetFieldValue("BankName", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                TxtBankAddress.Text = mobjEmpDll.CRecordset.GetFieldValue("BankAdd1", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                TxtAccNo.Text = mobjEmpDll.CRecordset.GetFieldValue("BAccNo", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)


                If strcheck = "J" Then
                    Me.chkJobWork.CheckState = System.Windows.Forms.CheckState.Checked
                    Me.chkStandard.CheckState = System.Windows.Forms.CheckState.Unchecked
                ElseIf strcheck = "S" Then
                    Me.chkJobWork.CheckState = System.Windows.Forms.CheckState.Unchecked
                    Me.chkStandard.CheckState = System.Windows.Forms.CheckState.Checked
                ElseIf strcheck = "B" Then
                    Me.chkJobWork.CheckState = System.Windows.Forms.CheckState.Checked
                    Me.chkStandard.CheckState = System.Windows.Forms.CheckState.Checked
                End If

                '********************
                txtExciseRange.Text = mobjEmpDll.CRecordset.GetFieldValue("excise_range", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtComRate.Text = mobjEmpDll.CRecordset.GetFieldValue("commisionrate", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtDivision.Text = mobjEmpDll.CRecordset.GetFieldValue("division", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtEcc.Text = mobjEmpDll.CRecordset.GetFieldValue("ecc_code", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtLst.Text = mobjEmpDll.CRecordset.GetFieldValue("lst_no", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtCst.Text = mobjEmpDll.CRecordset.GetFieldValue("cst_no", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtTinNo.Text = mobjEmpDll.CRecordset.GetFieldValue("Tin_no", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                chkShpmntThruWh.CheckState = IIf(rsLocation.GetValue("SHIPMENTTHRUWH") = True, 1, 0)
                chktT6Label.CheckState = IIf(rsLocation.GetValue("FordLabelReqd") = True, 1, 0)
                '*************10688280 -Add KAM Code in Customer Master********************
                txtKAMName.Text = mobjEmpDll.CRecordset.GetFieldValue("KAMCode", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                lblKAMDesc.Text = Convert.ToString(SqlConnectionclass.ExecuteScalar("select isnull(Name,'') from employee_mst where Unit_Code='" + gstrUNITID + "' and employee_code='" + txtKAMName.Text.Trim() + "'"))
                '************************************
                If rsLocation.GetValue("ConsigneeWiseLoc") = True Then
                    chkConsigneeWiseLoc.CheckState = System.Windows.Forms.CheckState.Checked
                Else
                    chkConsigneeWiseLoc.CheckState = System.Windows.Forms.CheckState.Unchecked
                End If
                If mobjEmpDll.CRecordset.GetFieldValue("CST_Eval", EMPDataBase.EMPDB.ADODataType.ADOBit) = True Then
                    ChkCSTEval.CheckState = System.Windows.Forms.CheckState.Checked
                Else
                    ChkCSTEval.CheckState = System.Windows.Forms.CheckState.Unchecked
                End If
                If mobjEmpDll.CRecordset.GetFieldValue("CST_amt", EMPDataBase.EMPDB.ADODataType.ADOBit) = True Then
                    ChKAMtEval.CheckState = System.Windows.Forms.CheckState.Checked
                Else
                    ChKAMtEval.CheckState = System.Windows.Forms.CheckState.Unchecked
                End If
                If mobjEmpDll.CRecordset.GetFieldValue("CSIEx_Inc", EMPDataBase.EMPDB.ADODataType.ADOBit) = True Then
                    ChkCSI_GT.CheckState = System.Windows.Forms.CheckState.Checked
                Else
                    ChkCSI_GT.CheckState = System.Windows.Forms.CheckState.Unchecked
                End If
                If mobjEmpDll.CRecordset.GetFieldValue("Has_Trans_flag", EMPDataBase.EMPDB.ADODataType.ADOBit) = True Then
                    ChkActive.CheckState = System.Windows.Forms.CheckState.Checked
                Else
                    ChkActive.CheckState = System.Windows.Forms.CheckState.Unchecked
                End If
                If mobjEmpDll.CRecordset.GetFieldValue("Group_Customer", EMPDataBase.EMPDB.ADODataType.ADOBit) = True Then
                    chkGroup.CheckState = System.Windows.Forms.CheckState.Checked
                Else
                    chkGroup.CheckState = System.Windows.Forms.CheckState.Unchecked
                End If
                If mobjEmpDll.CRecordset.GetFieldValue("CSIex_INC", EMPDataBase.EMPDB.ADODataType.ADOBit) = True Then
                    ChkCSI_GT.CheckState = System.Windows.Forms.CheckState.Checked
                    fraLedger.Visible = True
                    txtAcctLedgerExc.Text = mobjEmpDll.CRecordset.GetFieldValue("CSIEX_GL", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                    txtAccSubLedgerExc.Text = mobjEmpDll.CRecordset.GetFieldValue("CSIEX_SL", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                Else
                    ChkCSI_GT.CheckState = System.Windows.Forms.CheckState.Unchecked
                    fraLedger.Visible = False
                End If
                txtconsigneecode.Text = UCase(mobjEmpDll.CRecordset.GetFieldValue("Dock_code", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString))
                txtCustEDICode.Text = UCase(mobjEmpDll.CRecordset.GetFieldValue("Customer_EDIcode", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString))
                txtPlantCode.Text = UCase(mobjEmpDll.CRecordset.GetFieldValue("Plant_Code", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString))
                Me.cmdgrpCustMst.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = True
                Me.cmdgrpCustMst.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = True
                Me.cmdgrpCustMst.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE) = True
                If Me.txtCustCode.Enabled = False Then Me.txtCustCode.Enabled = True
                '10153377
                If mobjEmpDll.CRecordset.GetFieldValue("AllowBarcodePrinting", EMPDataBase.EMPDB.ADODataType.ADOBit) = True Then
                    Chkbarcodeprinting.CheckState = System.Windows.Forms.CheckState.Checked
                    CmbPrintMethod.SelectedIndex = CmbPrintMethod.FindStringExact(mobjEmpDll.CRecordset.GetFieldValue("PRINT_METHOD", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString))
                Else
                    Chkbarcodeprinting.CheckState = System.Windows.Forms.CheckState.Unchecked
                    CmbPrintMethod.SelectedIndex = -1
                End If

                strCust = mobjEmpDll.CRecordset.GetFieldValue("Cust_type", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                If strCust = "L" Then
                    Me.optDomestic.Checked = System.Windows.Forms.CheckState.Checked
                    Me.optOverseas.Checked = System.Windows.Forms.CheckState.Unchecked
                    Me.optInterState.Checked = System.Windows.Forms.CheckState.Unchecked
                ElseIf strCust = "O" Then
                    Me.optDomestic.Checked = System.Windows.Forms.CheckState.Unchecked
                    Me.optOverseas.Checked = System.Windows.Forms.CheckState.Checked
                    Me.optInterState.Checked = System.Windows.Forms.CheckState.Unchecked
                ElseIf strCust = "I" Then
                    Me.optDomestic.Checked = System.Windows.Forms.CheckState.Unchecked
                    Me.optOverseas.Checked = System.Windows.Forms.CheckState.Unchecked
                    Me.optInterState.Checked = System.Windows.Forms.CheckState.Checked
                End If
                'GST CHANGES

                strcustClassification = mobjEmpDll.CRecordset.GetFieldValue("Classificationtype", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)

                If strcustClassification = "B" Then
                    OptB2B.Checked = System.Windows.Forms.CheckState.Checked
                    optB2C.Checked = System.Windows.Forms.CheckState.Unchecked
                    OptGovtAgency.Checked = System.Windows.Forms.CheckState.Unchecked
                    OptInternationexemption.Checked = System.Windows.Forms.CheckState.Unchecked
                ElseIf strcustClassification = "C" Then
                    OptB2B.Checked = System.Windows.Forms.CheckState.Unchecked
                    optB2C.Checked = System.Windows.Forms.CheckState.Checked
                    OptGovtAgency.Checked = System.Windows.Forms.CheckState.Unchecked
                    OptInternationexemption.Checked = System.Windows.Forms.CheckState.Unchecked
                ElseIf strcustClassification = "G" Then
                    OptB2B.Checked = System.Windows.Forms.CheckState.Unchecked
                    optB2C.Checked = System.Windows.Forms.CheckState.Unchecked
                    OptGovtAgency.Checked = System.Windows.Forms.CheckState.Checked
                    OptInternationexemption.Checked = System.Windows.Forms.CheckState.Unchecked
                ElseIf strcustClassification = "S" Then
                    OptB2B.Checked = System.Windows.Forms.CheckState.Unchecked
                    optB2C.Checked = System.Windows.Forms.CheckState.Unchecked
                    OptGovtAgency.Checked = System.Windows.Forms.CheckState.Unchecked
                    OptInternationexemption.Checked = System.Windows.Forms.CheckState.Unchecked
                Else
                    OptB2B.Checked = System.Windows.Forms.CheckState.Unchecked
                    optB2C.Checked = System.Windows.Forms.CheckState.Unchecked
                    OptGovtAgency.Checked = System.Windows.Forms.CheckState.Unchecked
                    OptInternationexemption.Checked = System.Windows.Forms.CheckState.Checked
                End If

                txtGSTINId.Text = mobjEmpDll.CRecordset.GetFieldValue("GSTIN_Id", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtGSTBillState.Text = mobjEmpDll.CRecordset.GetFieldValue("GST_BILLSTATE_CODE", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                lblGSTStateDesc.Text = mobjEmpDll.CRecordset.GetFieldValue("GST_BILLSTATE_DESC", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtKeyContactName_1.Text = mobjEmpDll.CRecordset.GetFieldValue("ContactForMismatch_1", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtKeyContactPhNo_1.Text = mobjEmpDll.CRecordset.GetFieldValue("ContactForMismatch_PhoneNo_1", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtKeyContactMBNo_1.Text = mobjEmpDll.CRecordset.GetFieldValue("ContactForMismatch_MobileNo_1", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtKeyContactEmail_1.Text = mobjEmpDll.CRecordset.GetFieldValue("ContactForMismatch_Email_1", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtKeyContactName_2.Text = mobjEmpDll.CRecordset.GetFieldValue("ContactForMismatch_2", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtKeyContactPhNo_2.Text = mobjEmpDll.CRecordset.GetFieldValue("ContactForMismatch_PhoneNo_2", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtKeyContactMBNo_2.Text = mobjEmpDll.CRecordset.GetFieldValue("ContactForMismatch_MobileNo_2", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtKeyContactEmail_2.Text = mobjEmpDll.CRecordset.GetFieldValue("ContactForMismatch_Email_2", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)


                'GST CHANGES ENDED
                If mobjEmpDll.CRecordset.GetFieldValue("GSTIN_notrequired", EMPDataBase.EMPDB.ADODataType.ADOBit) = True Then
                    chkGSTINnotrequired.CheckState = System.Windows.Forms.CheckState.Checked
                Else
                    chkGSTINnotrequired.CheckState = System.Windows.Forms.CheckState.Unchecked
                End If
                '101482956
                If mobjEmpDll.CRecordset.GetFieldValue("SEZ", EMPDataBase.EMPDB.ADODataType.ADOBit) = True Then
                    chkSEZ.CheckState = System.Windows.Forms.CheckState.Checked
                Else
                    chkSEZ.CheckState = System.Windows.Forms.CheckState.Unchecked
                End If
                txtTRNNo.Text = mobjEmpDll.CRecordset.GetFieldValue("PARTY_TRN_NO", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                '10736222

                chkCT2.Checked = Convert.ToBoolean(mobjEmpDll.CRecordset.GetFieldValue("CT2_REQD", EMPDataBase.EMPDB.ADODataType.ADOBit, EMPDataBase.EMPDB.ADOCustomFormat.CustomString))
                'Samiksha Credit limit changes
                txtboxCreditLimit.Text = Convert.ToString(mobjEmpDll.CRecordset.GetFieldValue("CreditLimit", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString))
                'Samiksha branchcode changes
                txtboxBranchCode.Text = Convert.ToString(mobjEmpDll.CRecordset.GetFieldValue("branchcode", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString))

                Me.txtCustCode.Focus()
            End If
        End If

        mrsEmpDll.CloseRecordset()
        rsLocation.ResultSetClose()
        mrsEmpDll.OpenRecordset("select slm_desc from fin_slmaster Where Unit_Code='" & gstrUNITID & "' And slm_slcode = '" & Trim(txtAccSubLedgerExc.Text) & "' and slm_transtag = '1' and slm_glcode = '" & txtAcctLedgerExc.Text & "' ", ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockReadOnly)
        If mrsEmpDll.Recordcount > 0 Then
            lblsubdescExc.Text = mrsEmpDll.GetFieldValue("slm_desc", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
        End If
        mrsEmpDll.CloseRecordset()
        '----------------------------------------
        mrsEmpDll.OpenRecordset("select glm_desc from fin_glmaster Where Unit_Code='" & gstrUNITID & "' And glm_glcode = '" & Trim(txtAcctLedgerExc.Text) & "' and glm_transtag = '1'", ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockReadOnly)
        If mrsEmpDll.Recordcount > 0 Then
            Me.lblleddescExc.Text = mrsEmpDll.GetFieldValue("glm_desc", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
        End If
        mrsEmpDll.CloseRecordset()
        Call ShowShippingDetail(txtCustCode.Text) 'Show the Customer shipping Detail
        '----------------------------------------
        mrsEmpDll.OpenRecordset("select slm_desc from fin_slmaster Where Unit_Code='" & gstrUNITID & "' And slm_slcode = '" & Trim(txtAccSubLedger.Text) & "' and slm_transtag = '1' and slm_glcode = '" & txtAcctLedger.Text & "' ", ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockReadOnly)
        If mrsEmpDll.Recordcount > 0 Then
            lblsubdesc.Text = mrsEmpDll.GetFieldValue("slm_desc", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
        End If
        mrsEmpDll.CloseRecordset()
        '----------------------------------------
        mrsEmpDll.OpenRecordset("select glm_desc from fin_glmaster Where Unit_Code='" & gstrUNITID & "' And glm_glcode = '" & Trim(txtAcctLedger.Text) & "' and glm_transtag = '1'", ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockReadOnly)
        If mrsEmpDll.Recordcount > 0 Then
            Me.lblleddesc.Text = mrsEmpDll.GetFieldValue("glm_desc", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
        End If
        mrsEmpDll.CloseRecordset()
        '------------------------------------------
        mrsEmpDll.OpenRecordset("select currency_code, description from currency_mst where currency_code = '" & txtCurrencyCode.Text & "' and  Unit_Code='" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockReadOnly)
        If mrsEmpDll.Recordcount > 0 Then lblCurrDesc.Text = mrsEmpDll.GetFieldValue("description", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
        mrsEmpDll.CloseRecordset()
        '-------------------------------------------
        mrsEmpDll.OpenRecordset("select CrTrm_TermId, crtrm_desc from Gen_CreditTrmMaster Where Unit_Code='" & gstrUNITID & "' And CrTrm_TermId = '" & txtCreditTermId.Text & "' and CrTrm_Status = '1'", ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockReadOnly)
        If mrsEmpDll.Recordcount > 0 Then lblCreditDesc.Text = mrsEmpDll.GetFieldValue("crtrm_desc", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
        mrsEmpDll.CloseRecordset()
        '--------------------------------------------
        mobjEmpDll.CConnection.CloseConnection()
        '-----------------------------------------
        Me.tabDetails.SelectedIndex = 0
        Me.tabCustomer.SelectedIndex = 0
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        mobjEmpDll.CConnection.RollbackTransaction()
        mobjEmpDll.CConnection.CloseConnection()
        mrsEmpDll.CloseRecordset()
    End Function
    Public Function checkifnewcustomer() As Boolean
        Dim result As Boolean = False
        Dim strSQL As String = ""
        strSQL = "select dbo.ufn_CheckCustomerExistsInMst  ( '" + gstrUNITID + "','" + txtCustCode.Text + "')"
        SqlConnectionclass.OpenGlobalConnection()
        result = SqlConnectionclass.ExecuteScalar(strSQL)
        Return result

    End Function
    '*********************************************'
    'Author         :       Ananya Nath
    'Arguments      :       None
    'Return Value   :       Customer Code
    'Description    :       Generates and returns CustomerCode.
    '*********************************************'
    Public Function GenerateCustomerNo() As String ' used to generate Customer No.
        On Error GoTo ErrHandler
        Dim lngVar As Integer
        Dim merror As Excel.Error
        Dim strCustCode As String
        Dim strValue As String
        Dim strVar As String

        mobjEmpDll.CConnection.OpenConnection(gstrDSNName, gstrDatabaseName)

        mobjEmpDll.CConnection.BeginTransaction()
        mrsEmpDll.OpenRecordset("select customer_code from customer_mst Where Unit_Code='" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
        If mrsEmpDll.Recordcount > 0 Then
            mrsEmpDll.OpenRecordset("select max(SUBSTRING(customer_code,2,7)) from customer_mst Where Unit_Code='" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
            mrsEmpDll.MoveFirst()
            strCustCode = mrsEmpDll.GetString(ADODB.StringFormatEnum.adClipString, 1, "-")
            lngVar = CInt(Mid(strCustCode, 1, 7)) + 1
            strVar = VB6.Format(lngVar, "000000#")
            GenerateCustomerNo = "C" & strVar
        Else
            GenerateCustomerNo = "C0000001"
        End If
        mobjEmpDll.CConnection.CommitTransaction()
        mobjEmpDll.CConnection.CloseConnection()
        mrsEmpDll.CloseRecordset()
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        mobjEmpDll.CConnection.RollbackTransaction()
        mobjEmpDll.CConnection.CloseConnection()
        mrsEmpDll.CloseRecordset()
    End Function
    '*********************************************'
    'Author         :       Ananya Nath
    'Arguments      :       None
    'Return Value   :       None
    'Description    :       Inserts a new row in the customer_mst table.
    '*********************************************'
    'Function to insert row in the customer_mst
    Public Function Insert(ByVal strTblname As String) As Object
        On Error GoTo ErrHandler
        Dim strCustType As String
        Dim strVendCat As String = ""
        Dim strCustClassificationType As String
        Dim GSTINnotrequired As Boolean
        Dim isSEZ As Boolean
        mobjEmpDll.CConnection.OpenConnection(gstrDSNName, gstrDatabaseName)
        '10690771
        '10688280 -Add KAM Code in Customer Master.
        'Samiksha : Customer Master Authorization
        'Samiksha Credit limit changes
        'Samiksha branchcode changes
        If strTblname = "customer_mst" Then
            Call mrsEmpDll.OpenRecordset("SELECT Customer_Code,FordLabelReqd,Account_ledger,Account_subLedger,Cust_Name,Cust_Vendor_Code,Excise_range,Commisionrate,Division,ECC_Code,LST_No,CST_No,Cust_Location,Ship_Contact_person,Ship_Person_desig,Ship_email_id,Ship_Address1,Ship_Address2,Ship_City,Ship_State,Ship_Pin,Ship_dist,Ship_Country,Ship_Phone,Ship_Fax,bill_Contact_person,Bill_Person_desig,bill_email_id,Bill_Address1,Bill_Address2,Bill_City,Bill_State,Bill_Pin,Bill_dist,Bill_Country,Bill_Phone,Bill_Fax,Office_Contact_person,Office_person_design,office_email_id,Office_Address1,Office_Address2,Office_City,Office_State,Office_dist,Office_Pin,Office_Country,Office_Phone,Office_fax,Website,customer_type,currency_code,Bank_ac1,Bank_ac2,Bank_ac3,Credit_days,Has_Trans_flag,ScheduleCode,CST_Eval,tin_no,Milk_Van,AllowExcessSchedule,CUST_ALIAS,DOCK_CODE,ShipmentThruWh,PANNo,ServiceTaxNo,SwiftCode,IBANNo_SOT,BAccNo,BankName,BankAdd1,Tool_GL,Tool_SL,ASN_REQD,CST_amt,CSIEX_Inc,CSIEX_GL,CSIEX_SL,Group_Customer,AllowBarCodePrinting,CONSIGNEEWISELOC,AllowASNPrinting,AllowASNTextGeneration,IS_AEROBIN_CUST,Customer_EDICode,Plant_Code,ASNFunctionCode,INVOICEAGAINSTAGREEMENTMST,CUST_TYPE,Ent_dt,Ent_Userid,Upd_dt,Upd_Userid,Unit_Code,GLOBAL_CUATOMER_CODE,PRINT_METHOD,Shipping_Duration,KAMCODE,CT2_REQD,GSTIN_Id,GST_BILLSTATE_CODE,GST_BILLSTATE_DESC,ContactForMismatch_1,ContactForMismatch_PhoneNo_1,ContactForMismatch_MobileNo_1,ContactForMismatch_Email_1,ContactForMismatch_2,ContactForMismatch_PhoneNo_2,ContactForMismatch_MobileNo_2,ContactForMismatch_Email_2,ClassificationType,GSTIN_notrequired,PARTY_TRN_NO,SEZ,CreditLimit,branchcode   FROM customer_mst Where Unit_Code='" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
        ElseIf strTblname = "customer_mst_authorization" Then
            Call mrsEmpDll.OpenRecordset("SELECT Customer_Code,FordLabelReqd,Account_ledger,Account_subLedger,Cust_Name,Cust_Vendor_Code,Excise_range,Commisionrate,Division,ECC_Code,LST_No,CST_No,Cust_Location,Ship_Contact_person,Ship_Person_desig,Ship_email_id,Ship_Address1,Ship_Address2,Ship_City,Ship_State,Ship_Pin,Ship_dist,Ship_Country,Ship_Phone,Ship_Fax,bill_Contact_person,Bill_Person_desig,bill_email_id,Bill_Address1,Bill_Address2,Bill_City,Bill_State,Bill_Pin,Bill_dist,Bill_Country,Bill_Phone,Bill_Fax,Office_Contact_person,Office_person_design,office_email_id,Office_Address1,Office_Address2,Office_City,Office_State,Office_dist,Office_Pin,Office_Country,Office_Phone,Office_fax,Website,customer_type,currency_code,Bank_ac1,Bank_ac2,Bank_ac3,Credit_days,Has_Trans_flag,ScheduleCode,CST_Eval,tin_no,Milk_Van,AllowExcessSchedule,CUST_ALIAS,DOCK_CODE,ShipmentThruWh,PANNo,ServiceTaxNo,SwiftCode,IBANNo_SOT,BAccNo,BankName,BankAdd1,Tool_GL,Tool_SL,ASN_REQD,CST_amt,CSIEX_Inc,CSIEX_GL,CSIEX_SL,Group_Customer,AllowBarCodePrinting,CONSIGNEEWISELOC,AllowASNPrinting,AllowASNTextGeneration,IS_AEROBIN_CUST,Customer_EDICode,Plant_Code,ASNFunctionCode,INVOICEAGAINSTAGREEMENTMST,CUST_TYPE,Ent_dt,Ent_Userid,Upd_dt,Upd_Userid,Unit_Code,GLOBAL_CUATOMER_CODE,PRINT_METHOD,Shipping_Duration,KAMCODE,CT2_REQD,GSTIN_Id,GST_BILLSTATE_CODE,GST_BILLSTATE_DESC,ContactForMismatch_1,ContactForMismatch_PhoneNo_1,ContactForMismatch_MobileNo_1,ContactForMismatch_Email_1,ContactForMismatch_2,ContactForMismatch_PhoneNo_2,ContactForMismatch_MobileNo_2,ContactForMismatch_Email_2,ClassificationType,GSTIN_notrequired,PARTY_TRN_NO,SEZ,CreditLimit,branchcode   FROM customer_mst_authorization Where Unit_Code='" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

        End If
        mobjEmpDll.CConnection.BeginTransaction()
        mrsEmpDll.AddNew()
        With Me
            If chkJobWork.CheckState = 1 And chkStandard.CheckState = 0 Then
                strCustType = "J"
            Else
                If .chkStandard.CheckState = 1 And .chkJobWork.CheckState = 0 Then
                    strCustType = "S"
                ElseIf .chkJobWork.CheckState = 1 And .chkStandard.CheckState = 1 Then
                    strCustType = "B"
                End If
            End If
        End With
        'GST CHANGES
        If OptB2B.Checked = True Then
            strCustClassificationType = "B"
        ElseIf optB2C.Checked = True Then
            strCustClassificationType = "C"
        ElseIf OptGovtAgency.Checked = True Then
            strCustClassificationType = "G"
        Else
            strCustClassificationType = "I"
        End If


        Call mrsEmpDll.SetValue("GSTIN_Id", txtGSTINId.Text)
        Call mrsEmpDll.SetValue("GST_BILLSTATE_CODE", txtGSTBillState.Text)
        Call mrsEmpDll.SetValue("GST_BILLSTATE_DESC", lblGSTStateDesc.Text)
        'GST CHANGES
        Call mrsEmpDll.SetValue("customer_code", txtCustCode.Text)
        Call mrsEmpDll.SetValue("Account_ledger", txtAcctLedger.Text)
        Call mrsEmpDll.SetValue("Account_subledger", txtAccSubLedger.Text)
        Call mrsEmpDll.SetValue("cust_name", txtCustName.Text)
        Call mrsEmpDll.SetValue("excise_range", txtExciseRange.Text)
        Call mrsEmpDll.SetValue("commisionrate", txtComRate.Text)
        Call mrsEmpDll.SetValue("division", txtDivision.Text)
        Call mrsEmpDll.SetValue("ecc_code", txtEcc.Text)
        Call mrsEmpDll.SetValue("lst_no", txtLst.Text)
        Call mrsEmpDll.SetValue("cst_no", txtCst.Text)
        Call mrsEmpDll.SetValue("Cust_location", txtCustLoc.Text)
        Call mrsEmpDll.SetValue("Cust_Vendor_code", txtCustvendCode.Text)
        '***************
        Call mrsEmpDll.SetValue("office_Address1", txtoffadd11.Text)
        Call mrsEmpDll.SetValue("office_address2", txtOffAdd21.Text)
        Call mrsEmpDll.SetValue("Office_city", txtOffCity1.Text)
        Call mrsEmpDll.SetValue("Office_State", txtOffState1.Text)
        Call mrsEmpDll.SetValue("Office_Pin", txtOffPin1.Text)
        Call mrsEmpDll.SetValue("Office_Country", txtOffCountry1.Text)
        Call mrsEmpDll.SetValue("Office_Phone", txtOffPhone1.Text)
        Call mrsEmpDll.SetValue("Office_dist", txtOffDist1.Text)
        Call mrsEmpDll.SetValue("Office_email_id", txtEmail1.Text)
        Call mrsEmpDll.SetValue("Office_fax", txtOffFax1.Text)
        Call mrsEmpDll.SetValue("Office_contact_person", txtContPer1.Text)
        Call mrsEmpDll.SetValue("Office_person_design", txtDesig1.Text)
        '***************
        Call mrsEmpDll.SetValue("Bill_Address1", txtBilladd1.Text)
        Call mrsEmpDll.SetValue("Bill_address2", txtBillAdd2.Text)
        Call mrsEmpDll.SetValue("Bill_city", txtbillCity(0).Text)
        Call mrsEmpDll.SetValue("Bill_State", txtBillState.Text)
        Call mrsEmpDll.SetValue("Bill_Pin", txtBillPin.Text)
        Call mrsEmpDll.SetValue("Bill_Country", txtBillCountry.Text)
        Call mrsEmpDll.SetValue("Bill_Phone", txtBillPhone.Text)
        Call mrsEmpDll.SetValue("Bill_dist", txtBillDist.Text)
        Call mrsEmpDll.SetValue("Bill_email_id", txtBillEmail.Text)
        Call mrsEmpDll.SetValue("Bill_fax", txtBillFax.Text)
        Call mrsEmpDll.SetValue("Bill_contact_person", txtBillContPer.Text)
        Call mrsEmpDll.SetValue("Bill_person_desig", txtBillDesig.Text)
        '**************
        Call mrsEmpDll.SetValue("Customer_type", strCustType)
        Call mrsEmpDll.SetValue("currency_code", txtCurrencyCode.Text)
        Call mrsEmpDll.SetValue("website", txtWebSite.Text)
        Call mrsEmpDll.SetValue("Bank_ac1", txtBankAcct1.Text)
        Call mrsEmpDll.SetValue("Bank_ac2", txtBankAcct2.Text)
        Call mrsEmpDll.SetValue("Bank_ac3", txtBankAcct3.Text)
        Call mrsEmpDll.SetValue("Credit_days", Trim(txtCreditTermId.Text))
        Call mrsEmpDll.SetValue("has_trans_flag", False)
        Call mrsEmpDll.SetValue("Ent_dt", getDateForDB(GetServerDate()))
        Call mrsEmpDll.SetValue("Ent_Userid", mP_User)
        Call mrsEmpDll.SetValue("Upd_dt", getDateForDB(GetServerDate()))
        Call mrsEmpDll.SetValue("Upd_Userid", mP_User)
        Call mrsEmpDll.SetValue("milk_van", chkmilkvan.CheckState)
        Call mrsEmpDll.SetValue("CST_Eval", ChkCSTEval.CheckState)
        Call mrsEmpDll.SetValue("Tin_no", txtTinNo.Text)
        Call mrsEmpDll.SetValue("PANNo", Trim(TxtPanNo.Text))
        Call mrsEmpDll.SetValue("ServiceTaxNo", Trim(TxtServiceTaxNo.Text))
        Call mrsEmpDll.SetValue("SwiftCode", Trim(TxtSwiftNo.Text))
        Call mrsEmpDll.SetValue("IBANNo_SOT", Trim(TxtIBANNo.Text))
        Call mrsEmpDll.SetValue("BankName", Trim(TxtBankname.Text))
        Call mrsEmpDll.SetValue("BankAdd1", Trim(TxtBankAddress.Text))
        Call mrsEmpDll.SetValue("BAccNo", Trim(TxtAccNo.Text))
        Call mrsEmpDll.SetValue("ShipmentThruWh", chkShpmntThruWh.CheckState)
        Call mrsEmpDll.SetValue("FordLabelReqd", chktT6Label.CheckState)
        Call mrsEmpDll.SetValue("ConsigneeWiseLoc", chkConsigneeWiseLoc.CheckState)
        Call mrsEmpDll.SetValue("CST_Amt", ChKAMtEval.CheckState)
        Call mrsEmpDll.SetValue("CSIex_INC", ChkCSI_GT.CheckState)
        Call mrsEmpDll.SetValue("has_trans_flag", ChkActive.CheckState)
        Call mrsEmpDll.SetValue("Group_Customer", chkGroup.CheckState)
        If ChkCSI_GT.CheckState = 1 Then
            Call mrsEmpDll.SetValue("CSIEX_GL", txtAcctLedgerExc.Text)
            Call mrsEmpDll.SetValue("CSIEX_SL", txtAccSubLedgerExc.Text)
        End If
        Call mrsEmpDll.SetValue("Dock_code", UCase(Trim(txtconsigneecode.Text)))
        Call mrsEmpDll.SetValue("Customer_EDICode", txtCustEDICode.Text.Trim().ToUpper())
        Call mrsEmpDll.SetValue("Plant_Code", txtPlantCode.Text.Trim().ToUpper())
        If Me.optDomestic.Checked = True Then strVendCat = "L" Else If Me.optOverseas.Checked = True Then strVendCat = "O" Else If Me.optInterState.Checked = True Then strVendCat = "I"
        Call mrsEmpDll.SetValue("Cust_Type", strVendCat.Trim().ToUpper())
        Call mrsEmpDll.SetValue("Unit_Code", gstrUNITID)
        Call mrsEmpDll.SetValue("GLOBAL_CUATOMER_CODE", lblGlobalCustCodeDesc.Text.ToString.Trim)
        Call mrsEmpDll.SetValue("Shipping_Duration", Val(txtShipDur.Text))
        Call mrsEmpDll.SetValue("AllowBarcodePrinting", Chkbarcodeprinting.CheckState)
        Call mrsEmpDll.SetValue("PRINT_METHOD", CmbPrintMethod.Text.ToString().Trim)
        '10688280 -Add KAM Code in Customer Master.
        Call mrsEmpDll.SetValue("KAMCode", txtKAMName.Text.Trim)
        '10736222 — eMPro-- CT2 - ARE3 functionality
        Call mrsEmpDll.SetValue("CT2_REQD", chkCT2.CheckState)
        'GST CHNAGES
        Call mrsEmpDll.SetValue("ClassificationType", strCustClassificationType)
        'GST CHNAGES

        '101459347 — Changes in Customer & Vendor Master
        If chkGSTINnotrequired.Checked = True Then
            GSTINnotrequired = 1
        Else
            GSTINnotrequired = 0
        End If
        Call mrsEmpDll.SetValue("GSTIN_notrequired", GSTINnotrequired)
        'end here 
        If chkSEZ.Checked = True Then
            isSEZ = 1
        Else
            isSEZ = 0
        End If
        Call mrsEmpDll.SetValue("SEZ", isSEZ)
        'Samiksha Credit limit changes
        Call mrsEmpDll.SetValue("CreditLimit", If(isCreditLimitMandatory = True, Convert.ToDouble(txtboxCreditLimit.Text), Convert.ToDouble(0)))
        'Samiksha branchcode changes
        Call mrsEmpDll.SetValue("branchcode", If(Len(txtboxBranchCode.Text) = 0, Convert.ToString(""), Convert.ToString(txtboxBranchCode.Text)))

        Call mrsEmpDll.SetValue("PARTY_TRN_NO", Trim(txtTRNNo.Text))
        Call InsertUpdateShippingDetails(txtCustCode.Text)
        Call InsertA4customer(txtCustCode.Text)
        mrsEmpDll.Update()
        mobjEmpDll.CConnection.CommitTransaction()
        mrsEmpDll.CloseRecordset()
        mobjEmpDll.CConnection.CloseConnection()
        ResetDatabaseConnection()
        mP_Connection.BeginTrans()
        mP_Connection.Execute(mstrInsertUpdate, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        If Not IsNothing(mstrA4insert) Then
            If Len(mstrA4insert.ToString) > 0 Then
                mP_Connection.Execute(mstrA4insert, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End If
        End If
        mP_Connection.CommitTrans()
        ConfirmWindow(10049, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
        Exit Function
        ' End If
ErrHandler:
        If (Err.Number.ToString().Trim = "-2147217900".Trim Or Err.Number.ToString.Trim = "5") Then
            MsgBox("This Customer Code is already saved.", MsgBoxStyle.Information, "eMPro")
            RefreshForm()
            'mobjEmpDll.CConnection.RollbackTransaction()
            'mobjEmpDll.CConnection.CloseConnection()
            'mrsEmpDll.CloseRecordset()
            mobjEmpDll.CConnection.RollbackTransaction()
            '  mrsEmpDll.CloseRecordset()
            mobjEmpDll.CConnection.CloseConnection()
        Else
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
            mP_Connection.RollbackTrans()
            mobjEmpDll.CConnection.RollbackTransaction()
            mobjEmpDll.CConnection.CloseConnection()
            mrsEmpDll.CloseRecordset()
        End If
    End Function
    '*********************************************'
    'Author         :       Ananya Nath
    'Arguments      :       None
    'Return Value   :       None
    'Description    :       Updates record.
    '*********************************************'

    ' check for Duplicate value in  customer_mst table''''''
    'Public Function getduplicatevalue() As Boolean
    '    Dim strSql As String = String.Empty
    '    Dim rsgetdetail As ClsResultSetDB
    '    Dim flag As Boolean
    '    rsgetdetail = New ClsResultSetDB
    '    strSql = "select GLOBAL_CUATOMER_CODE from customer_mst Where Unit_Code='" & gstrUNITID & "' And GLOBAL_CUATOMER_CODE='" & Trim(lblGlobalCustCodeDesc.Text) & "'"
    '    rsgetdetail.GetResult(strSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
    '    If rsgetdetail.GetNoRows > 0 Then
    '        flag = True
    '        rsgetdetail.ResultSetClose()
    '        rsgetdetail = Nothing

    '    Else
    '        flag = False

    '    End If
    '    Return flag
    'End Function


    'end
    'Samiksha: Customer Master Authorisation
    Public Function Update_Renamed_DotNet(ByVal strTblName As String) As Object 'This Function Updates the customer_mst in Dot Net

        Dim SqlCmd As New SqlCommand
        Dim strCustomerType As String
        Dim strCustType As String
        Dim strCustClassificationType As String
        Dim StrQuery As String = ""
        Dim StrDeleteCustMst As String = ""

        Try
            'Initialization Section for objects
            If Me.chkJobWork.CheckState = 1 And Me.chkStandard.CheckState = 1 Then
                strCustomerType = "B"
            ElseIf Me.chkStandard.CheckState = 1 And Me.chkJobWork.CheckState = 0 Then
                strCustomerType = "S"
            ElseIf Me.chkJobWork.CheckState = 1 And Me.chkStandard.CheckState = 0 Then
                strCustomerType = "J"
            Else
                strCustomerType = ""
            End If

            If optDomestic.Checked = True Then
                strCustType = "L"
            ElseIf optOverseas.Checked = True Then
                strCustType = "O"
            ElseIf optInterState.Checked = True Then
                strCustType = "I"
            Else
                strCustType = ""
            End If

            If OptB2B.Checked = True Then
                strCustClassificationType = "B"
            ElseIf optB2C.Checked = True Then
                strCustClassificationType = "C"
            ElseIf OptGovtAgency.Checked = True Then
                strCustClassificationType = "G"
            Else
                strCustClassificationType = "I"
            End If

            'Ends Initialization Section for objects



            If strTblName = "customer_mst_authorization" Then

                If checkifexistsincustmstauth() = False Then

                    Dim strErrormsg As String = String.Empty
                    Using cmd As SqlCommand = New SqlCommand
                        With cmd
                            .CommandType = CommandType.StoredProcedure
                            .CommandText = "USP_INSERT_CUSTOMER_MST_AUTHORIZATION"
                            .CommandTimeout = 0
                            .Parameters.Clear()
                            .Parameters.AddWithValue("@UNIT_CODE", gstrUNITID)
                            .Parameters.AddWithValue("@CUSTOMER_CODE", Trim(txtCustCode.Text))
                            .Parameters.Add("@Error_Msg", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
                            SqlConnectionclass.ExecuteNonQuery(cmd)

                            If cmd.Parameters("@Error_Msg").Value.ToString() <> "" Then
                                strErrormsg = Convert.ToString(cmd.Parameters("@Error_Msg").SqlValue)
                                MsgBox(strErrormsg, MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100))
                                Exit Function
                            End If
                        End With
                    End Using
                End If

                StrQuery = "Update customer_mst_authorization Set Account_ledger ='" & txtAcctLedger.Text & "'," & vbCrLf &
                "Account_subledger = '" & txtAccSubLedger.Text & "', Cust_name= '" & txtCustName.Text & "', excise_range = '" & txtExciseRange.Text & "'," & vbCrLf &
                "commisionrate = '" & txtComRate.Text & "', division ='" & txtDivision.Text & "', ecc_code = '" & txtEcc.Text & "'," & vbCrLf &
                "lst_no ='" & txtLst.Text & "', cst_no='" & txtCst.Text & "', Cust_location='" & txtCustLoc.Text & "', cust_vendor_code='" & txtCustvendCode.Text & "'," & vbCrLf &
                "office_Address1='" & txtoffadd11.Text & "', office_address2='" & txtOffAdd21.Text & "', Office_city='" & txtOffCity1.Text & "'," & vbCrLf &
                "Office_State='" & txtOffState1.Text & "', Office_Pin='" & txtOffPin1.Text & "', Office_Country='" & txtOffCountry1.Text & "'," & vbCrLf &
                "Office_Phone='" & txtOffPhone1.Text & "', Office_dist='" & txtOffDist1.Text & "', Office_email_id='" & txtEmail1.Text & "'," & vbCrLf &
                "Office_fax='" & txtOffFax1.Text & "', Office_contact_person='" & txtContPer1.Text & "', Office_person_design='" & txtDesig1.Text & "'," & vbCrLf &
                "Bill_Address1='" & txtBilladd1.Text & "', Bill_address2='" & txtBillAdd2.Text & "', Bill_city='" & txtbillCity(0).Text & "'," & vbCrLf &
                "Bill_State='" & txtBillState.Text & "', Bill_Pin='" & txtBillPin.Text & "', Bill_Country='" & txtBillCountry.Text & "'," & vbCrLf &
                "Bill_Phone='" & txtBillPhone.Text & "', Bill_dist='" & txtBillDist.Text & "', Bill_email_id='" & txtBillEmail.Text & "'," & vbCrLf &
                "Bill_fax='" & txtBillFax.Text & "', Bill_contact_person='" & txtBillContPer.Text & "', Bill_person_desig='" & txtBillDesig.Text & "'," & vbCrLf &
                "Customer_type= '" & strCustomerType & "', currency_code= '" & txtCurrencyCode.Text & "', website= '" & txtWebSite.Text & "'," & vbCrLf &
                "Bank_ac1= '" & txtBankAcct1.Text & "', Bank_ac2= '" & txtBankAcct2.Text & "', Bank_ac3= '" & txtBankAcct3.Text & "'," & vbCrLf &
                "Credit_days= '" & Trim(txtCreditTermId.Text) & "', Upd_dt= '" & getDateForDB(GetServerDate()) & "', Upd_Userid= '" & mP_User & "'," & vbCrLf &
                "Milk_Van=" & chkmilkvan.CheckState & ", CST_Eval= " & ChkCSTEval.CheckState & ", Tin_no='" & txtTinNo.Text & "'," & vbCrLf &
                "PANNo='" & Trim(TxtPanNo.Text) & "', ServiceTaxNo='" & Trim(TxtServiceTaxNo.Text) & "', SwiftCode='" & Trim(TxtSwiftNo.Text) & "'," & vbCrLf &
                "IBANNo_SOT='" & Trim(TxtIBANNo.Text) & "', BankName= '" & Trim(TxtBankname.Text) & "', BankAdd1='" & Trim(TxtBankAddress.Text) & "'," & vbCrLf &
                "BAccNo='" & Trim(TxtAccNo.Text) & "', ShipmentThruWh='" & chkShpmntThruWh.CheckState & "', FordLabelReqd=" & chktT6Label.CheckState & ", " & vbCrLf &
                "ConsigneeWiseLoc=" & chkConsigneeWiseLoc.CheckState & ", CST_amt=" & ChKAMtEval.CheckState & ", CSIex_INC=" & ChkCSI_GT.CheckState & ", " & vbCrLf &
                "Has_Trans_flag=" & ChkActive.CheckState & ", Group_Customer=" & chkGroup.CheckState & ", Shipping_Duration='" & Val(txtShipDur.Text) & "'," & vbCrLf &
                "CSIEX_GL = '" & IIf(ChkCSI_GT.CheckState = 1, txtAcctLedgerExc.Text, "") & "'," & vbCrLf &
                "CSIEX_SL = '" & IIf(ChkCSI_GT.CheckState = 1, txtAccSubLedgerExc.Text, "") & "'," & vbCrLf &
                "Dock_code='" & UCase(Trim(txtconsigneecode.Text)) & "', Customer_EDICode='" & txtCustEDICode.Text.Trim().ToUpper() & "'," & vbCrLf &
                "Plant_Code='" & txtPlantCode.Text.Trim().ToUpper() & "',KAMCode='" & txtKAMName.Text.Trim() & "',CT2_REQD=" & chkCT2.CheckState & "," & vbCrLf &
                "Cust_Type='" & strCustType.Trim().ToUpper() & "'," & vbCrLf &
                "GLOBAL_CUATOMER_CODE = " & IIf(lblGlobalCustCodeDesc.Text.Trim = String.Empty, "Null", "'" & lblGlobalCustCodeDesc.Text.ToString.Trim & "'") & "," & vbCrLf &
                "classificationtype= '" & strCustClassificationType & "', GSTIN_Id= '" & txtGSTINId.Text & "'," & vbCrLf &
                "GST_BILLSTATE_CODE= '" & txtGSTBillState.Text.Trim() & "', GST_BILLSTATE_DESC= '" & lblGSTStateDesc.Text & "'," & vbCrLf &
                "ContactForMismatch_1= '" & txtKeyContactName_1.Text & "', ContactForMismatch_PhoneNo_1= '" & txtKeyContactPhNo_1.Text & "'," & vbCrLf &
                "ContactForMismatch_MobileNo_1= '" & txtKeyContactMBNo_1.Text & "', ContactForMismatch_Email_1= '" & txtKeyContactEmail_1.Text & "'," & vbCrLf &
                "ContactForMismatch_2= '" & txtKeyContactName_2.Text & "', ContactForMismatch_PhoneNo_2= '" & txtKeyContactPhNo_2.Text & "'," & vbCrLf &
                "ContactForMismatch_MobileNo_2= '" & txtKeyContactMBNo_2.Text & "', ContactForMismatch_Email_2= '" & txtKeyContactEmail_2.Text & "'," & vbCrLf &
                "GSTIN_notrequired= '" & IIf(chkGSTINnotrequired.Checked, True, False) & "',PARTY_TRN_NO ='" & Trim(txtTRNNo.Text) & "', " & vbCrLf &
                "SEZ= '" & IIf(chkSEZ.Checked, True, False) & "', AllowBarcodePrinting= " & Chkbarcodeprinting.CheckState & ", PRINT_METHOD='" & CmbPrintMethod.Text.ToString().Trim & "',authorization_status=NULL,Authorization_Remark = NULL," & vbCrLf &
                "CreditLimit= " & If(isCreditLimitMandatory = True, Convert.ToDouble(txtboxCreditLimit.Text), Convert.ToDouble(0)) & "," & vbCrLf &
                "branchcode='" & If(Len(txtboxBranchCode.Text) = 0, "", Convert.ToString(txtboxBranchCode.Text)) & "' Where unit_code = '" & gstrUNITID & "' and customer_code = '" & txtCustCode.Text & "'"
                'Samiksha Credit limit changes
                'Samiksha branchcode changes

            ElseIf strTblName = "customer_mst" Then



                StrQuery = "Update customer_mst Set Account_ledger ='" & txtAcctLedger.Text & "'," & vbCrLf &
              "Account_subledger = '" & txtAccSubLedger.Text & "', Cust_name= '" & txtCustName.Text & "', excise_range = '" & txtExciseRange.Text & "'," & vbCrLf &
              "commisionrate = '" & txtComRate.Text & "', division ='" & txtDivision.Text & "', ecc_code = '" & txtEcc.Text & "'," & vbCrLf &
              "lst_no ='" & txtLst.Text & "', cst_no='" & txtCst.Text & "', Cust_location='" & txtCustLoc.Text & "', cust_vendor_code='" & txtCustvendCode.Text & "'," & vbCrLf &
              "office_Address1='" & txtoffadd11.Text & "', office_address2='" & txtOffAdd21.Text & "', Office_city='" & txtOffCity1.Text & "'," & vbCrLf &
              "Office_State='" & txtOffState1.Text & "', Office_Pin='" & txtOffPin1.Text & "', Office_Country='" & txtOffCountry1.Text & "'," & vbCrLf &
              "Office_Phone='" & txtOffPhone1.Text & "', Office_dist='" & txtOffDist1.Text & "', Office_email_id='" & txtEmail1.Text & "'," & vbCrLf &
              "Office_fax='" & txtOffFax1.Text & "', Office_contact_person='" & txtContPer1.Text & "', Office_person_design='" & txtDesig1.Text & "'," & vbCrLf &
              "Bill_Address1='" & txtBilladd1.Text & "', Bill_address2='" & txtBillAdd2.Text & "', Bill_city='" & txtbillCity(0).Text & "'," & vbCrLf &
              "Bill_State='" & txtBillState.Text & "', Bill_Pin='" & txtBillPin.Text & "', Bill_Country='" & txtBillCountry.Text & "'," & vbCrLf &
              "Bill_Phone='" & txtBillPhone.Text & "', Bill_dist='" & txtBillDist.Text & "', Bill_email_id='" & txtBillEmail.Text & "'," & vbCrLf &
              "Bill_fax='" & txtBillFax.Text & "', Bill_contact_person='" & txtBillContPer.Text & "', Bill_person_desig='" & txtBillDesig.Text & "'," & vbCrLf &
              "Customer_type= '" & strCustomerType & "', currency_code= '" & txtCurrencyCode.Text & "', website= '" & txtWebSite.Text & "'," & vbCrLf &
              "Bank_ac1= '" & txtBankAcct1.Text & "', Bank_ac2= '" & txtBankAcct2.Text & "', Bank_ac3= '" & txtBankAcct3.Text & "'," & vbCrLf &
              "Credit_days= '" & Trim(txtCreditTermId.Text) & "', Upd_dt= '" & getDateForDB(GetServerDate()) & "', Upd_Userid= '" & mP_User & "'," & vbCrLf &
              "Milk_Van=" & chkmilkvan.CheckState & ", CST_Eval= " & ChkCSTEval.CheckState & ", Tin_no='" & txtTinNo.Text & "'," & vbCrLf &
              "PANNo='" & Trim(TxtPanNo.Text) & "', ServiceTaxNo='" & Trim(TxtServiceTaxNo.Text) & "', SwiftCode='" & Trim(TxtSwiftNo.Text) & "'," & vbCrLf &
              "IBANNo_SOT='" & Trim(TxtIBANNo.Text) & "', BankName= '" & Trim(TxtBankname.Text) & "', BankAdd1='" & Trim(TxtBankAddress.Text) & "'," & vbCrLf &
              "BAccNo='" & Trim(TxtAccNo.Text) & "', ShipmentThruWh='" & chkShpmntThruWh.CheckState & "', FordLabelReqd=" & chktT6Label.CheckState & ", " & vbCrLf &
              "ConsigneeWiseLoc=" & chkConsigneeWiseLoc.CheckState & ", CST_amt=" & ChKAMtEval.CheckState & ", CSIex_INC=" & ChkCSI_GT.CheckState & ", " & vbCrLf &
              "Has_Trans_flag=" & ChkActive.CheckState & ", Group_Customer=" & chkGroup.CheckState & ", Shipping_Duration='" & Val(txtShipDur.Text) & "'," & vbCrLf &
              "CSIEX_GL = '" & IIf(ChkCSI_GT.CheckState = 1, txtAcctLedgerExc.Text, "") & "'," & vbCrLf &
              "CSIEX_SL = '" & IIf(ChkCSI_GT.CheckState = 1, txtAccSubLedgerExc.Text, "") & "'," & vbCrLf &
              "Dock_code='" & UCase(Trim(txtconsigneecode.Text)) & "', Customer_EDICode='" & txtCustEDICode.Text.Trim().ToUpper() & "'," & vbCrLf &
              "Plant_Code='" & txtPlantCode.Text.Trim().ToUpper() & "',KAMCode='" & txtKAMName.Text.Trim() & "',CT2_REQD=" & chkCT2.CheckState & "," & vbCrLf &
              "Cust_Type='" & strCustType.Trim().ToUpper() & "'," & vbCrLf &
              "GLOBAL_CUATOMER_CODE = " & IIf(lblGlobalCustCodeDesc.Text.Trim = String.Empty, "Null", "'" & lblGlobalCustCodeDesc.Text.ToString.Trim & "'") & "," & vbCrLf &
              "classificationtype= '" & strCustClassificationType & "', GSTIN_Id= '" & txtGSTINId.Text & "'," & vbCrLf &
              "GST_BILLSTATE_CODE= '" & txtGSTBillState.Text.Trim() & "', GST_BILLSTATE_DESC= '" & lblGSTStateDesc.Text & "'," & vbCrLf &
              "ContactForMismatch_1= '" & txtKeyContactName_1.Text & "', ContactForMismatch_PhoneNo_1= '" & txtKeyContactPhNo_1.Text & "'," & vbCrLf &
              "ContactForMismatch_MobileNo_1= '" & txtKeyContactMBNo_1.Text & "', ContactForMismatch_Email_1= '" & txtKeyContactEmail_1.Text & "'," & vbCrLf &
              "ContactForMismatch_2= '" & txtKeyContactName_2.Text & "', ContactForMismatch_PhoneNo_2= '" & txtKeyContactPhNo_2.Text & "'," & vbCrLf &
              "ContactForMismatch_MobileNo_2= '" & txtKeyContactMBNo_2.Text & "', ContactForMismatch_Email_2= '" & txtKeyContactEmail_2.Text & "'," & vbCrLf &
              "GSTIN_notrequired= '" & IIf(chkGSTINnotrequired.Checked, True, False) & "',PARTY_TRN_NO ='" & Trim(txtTRNNo.Text) & "', " & vbCrLf &
              "SEZ= '" & IIf(chkSEZ.Checked, True, False) & "', AllowBarcodePrinting= " & Chkbarcodeprinting.CheckState & ", PRINT_METHOD='" & CmbPrintMethod.Text.ToString().Trim & "'," & vbCrLf &
              "CreditLimit= " & IIf(isCreditLimitMandatory = True, Convert.ToDouble(txtboxCreditLimit.Text), Convert.ToDouble(0)) & "," & vbCrLf &
              "branchcode='" & IIf(Len(txtboxBranchCode.Text) = 0, "", Convert.ToString(txtboxBranchCode.Text)) & "' Where unit_code = '" & gstrUNITID & "' and customer_code = '" & txtCustCode.Text & "'"
                'Samiksha Credit limit changes
                'Samiksha branchcode changes


            End If

            'If strTblName = "customer_mst_authorization" Then
            '    Dim authStatus As String = getAuthorisationstatus()
            '    If authStatus = "Auth" Then
            '        StrDeleteCustMst = "Delete from customer_mst where UNIT_CODE='" & gstrUNITID & "' and Customer_Code='" & txtCustCode.Text & "'"
            '    End If
            'End If

            SqlCmd = New SqlCommand()
            SqlCmd.Connection = SqlConnectionclass.GetConnection()
            SqlCmd.Transaction = SqlCmd.Connection.BeginTransaction

            SqlCmd.CommandType = CommandType.Text
            SqlCmd.CommandText = StrQuery
            SqlCmd.ExecuteNonQuery()

            InsertUpdateShippingDetails_DotNet(Trim(txtCustCode.Text))
            If mstrInsertUpdate.Length > 0 Then
                SqlCmd.CommandText = mstrInsertUpdate
                SqlCmd.ExecuteNonQuery()
            End If

            If StrDeleteCustMst.Length > 0 Then
                SqlCmd.CommandText = StrDeleteCustMst
                SqlCmd.ExecuteNonQuery()
            End If

            SqlCmd.Transaction.Commit()
            MessageBox.Show("Transaction Completed Successfully.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            SqlCmd.Transaction.Rollback()
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
            SqlCmd.Dispose()
        End Try
    End Function
    Public Function checkifexistsincustmstauth() As Boolean
        Dim result As Boolean = False
        Dim strSQL As String = ""
        strSQL = "select dbo.ufn_CheckCustomerExistsInAuthMst  ( '" + gstrUNITID + "','" + txtCustCode.Text + "')"
        SqlConnectionclass.OpenGlobalConnection()
        result = SqlConnectionclass.ExecuteScalar(strSQL)
        Return result
    End Function
    Public Function Update_Renamed() As Object 'This Function Updates the customer_mst
        On Error GoTo ErrHandler
        Dim strCustType As String
        Dim strCustClassificationType As String
        Dim GSTINnotrequired As Boolean
        Dim isSEZ As Boolean
        mobjEmpDll.CConnection.OpenConnection(gstrDSNName, gstrDatabaseName)
        mobjEmpDll.CConnection.BeginTransaction()
        '10690771
        '10688280 -Add KAM Code in Customer Master.
        Call mrsEmpDll.OpenRecordset("select Customer_Code,FordLabelReqd,Account_ledger,Account_subLedger,Cust_Name,Cust_Vendor_Code,Excise_range,Commisionrate,Division,ECC_Code,LST_No,CST_No,Cust_Location,Ship_Contact_person,Ship_Person_desig,Ship_email_id,Ship_Address1,Ship_Address2,Ship_City,Ship_State,Ship_Pin,Ship_dist,Ship_Country,Ship_Phone,Ship_Fax,bill_Contact_person,Bill_Person_desig,bill_email_id,Bill_Address1,Bill_Address2,Bill_City,Bill_State,Bill_Pin,Bill_dist,Bill_Country,Bill_Phone,Bill_Fax,Office_Contact_person,Office_person_design,office_email_id,Office_Address1,Office_Address2,Office_City,Office_State,Office_dist,Office_Pin,Office_Country,Office_Phone,Office_fax,Website,customer_type,currency_code,Bank_ac1,Bank_ac2,Bank_ac3,Credit_days,Has_Trans_flag,ScheduleCode,CST_Eval,tin_no,Milk_Van,AllowExcessSchedule,CUST_ALIAS,DOCK_CODE,ShipmentThruWh,PANNo,ServiceTaxNo,SwiftCode,IBANNo_SOT,BAccNo,BankName,BankAdd1,Tool_GL,Tool_SL,ASN_REQD,CST_amt,CSIEX_Inc,CSIEX_GL,CSIEX_SL,Group_Customer,AllowBarCodePrinting,CONSIGNEEWISELOC,AllowASNPrinting,AllowASNTextGeneration,IS_AEROBIN_CUST,Customer_EDICode,Plant_Code,ASNFunctionCode,INVOICEAGAINSTAGREEMENTMST,CUST_TYPE,Ent_dt,Ent_Userid,Upd_dt,Upd_Userid,UNIT_CODE,GLOBAL_CUATOMER_CODE,PRINT_METHOD,Shipping_Duration, KAMCode,CT2_REQD,GSTIN_Id,GST_BILLSTATE_CODE,GST_BILLSTATE_DESC,ContactForMismatch_1,ContactForMismatch_PhoneNo_1,ContactForMismatch_MobileNo_1,ContactForMismatch_Email_1,ContactForMismatch_2,ContactForMismatch_PhoneNo_2,ContactForMismatch_MobileNo_2,ContactForMismatch_Email_2,classificationtype,GSTIN_notrequired,PARTY_TRN_NO,SEZ  from customer_mst Where Unit_Code='" & gstrUNITID & "' And customer_code = '" & txtCustCode.Text & "' ", ADODB.CursorTypeEnum.adOpenDynamic)
        With Me
            If .chkJobWork.CheckState = 1 And .chkStandard.CheckState = 0 Then
                strCustType = "J"
            Else
                If .chkStandard.CheckState = 1 And .chkJobWork.CheckState = 0 Then
                    strCustType = "S"
                ElseIf .chkJobWork.CheckState = 1 And .chkStandard.CheckState = 1 Then
                    strCustType = "B"
                End If
            End If
        End With
        Call mrsEmpDll.SetValue("customer_code", txtCustCode.Text)
        Call mrsEmpDll.SetValue("Account_ledger", txtAcctLedger.Text)
        Call mrsEmpDll.SetValue("Account_subledger", txtAccSubLedger.Text)
        Call mrsEmpDll.SetValue("Cust_name", txtCustName.Text)
        Call mrsEmpDll.SetValue("excise_range", txtExciseRange.Text)
        Call mrsEmpDll.SetValue("commisionrate", txtComRate.Text)
        Call mrsEmpDll.SetValue("division", txtDivision.Text)
        Call mrsEmpDll.SetValue("ecc_code", txtEcc.Text)
        Call mrsEmpDll.SetValue("lst_no", txtLst.Text)
        Call mrsEmpDll.SetValue("cst_no", txtCst.Text)
        Call mrsEmpDll.SetValue("Cust_location", txtCustLoc.Text)
        Call mrsEmpDll.SetValue("cust_vendor_code", txtCustvendCode.Text)
        '***************
        Call mrsEmpDll.SetValue("office_Address1", txtoffadd11.Text)
        Call mrsEmpDll.SetValue("office_address2", txtOffAdd21.Text)
        Call mrsEmpDll.SetValue("Office_city", txtOffCity1.Text)
        Call mrsEmpDll.SetValue("Office_State", txtOffState1.Text)
        Call mrsEmpDll.SetValue("Office_Pin", txtOffPin1.Text)
        Call mrsEmpDll.SetValue("Office_Country", txtOffCountry1.Text)
        Call mrsEmpDll.SetValue("Office_Phone", txtOffPhone1.Text)
        Call mrsEmpDll.SetValue("Office_dist", txtOffDist1.Text)
        Call mrsEmpDll.SetValue("Office_email_id", txtEmail1.Text)
        Call mrsEmpDll.SetValue("Office_fax", txtOffFax1.Text)
        Call mrsEmpDll.SetValue("Office_contact_person", txtContPer1.Text)
        Call mrsEmpDll.SetValue("Office_person_design", txtDesig1.Text)
        '***************
        Call mrsEmpDll.SetValue("Bill_Address1", txtBilladd1.Text)
        Call mrsEmpDll.SetValue("Bill_address2", txtBillAdd2.Text)
        Call mrsEmpDll.SetValue("Bill_city", txtbillCity(0).Text)
        Call mrsEmpDll.SetValue("Bill_State", txtBillState.Text)
        Call mrsEmpDll.SetValue("Bill_Pin", txtBillPin.Text)
        Call mrsEmpDll.SetValue("Bill_Country", txtBillCountry.Text)
        Call mrsEmpDll.SetValue("Bill_Phone", txtBillPhone.Text)
        Call mrsEmpDll.SetValue("Bill_dist", txtBillDist.Text)
        Call mrsEmpDll.SetValue("Bill_email_id", txtBillEmail.Text)
        Call mrsEmpDll.SetValue("Bill_fax", txtBillFax.Text)
        Call mrsEmpDll.SetValue("Bill_contact_person", txtBillContPer.Text)
        Call mrsEmpDll.SetValue("Bill_person_desig", txtBillDesig.Text)
        '**************
        Call mrsEmpDll.SetValue("Customer_type", strCustType)
        Call mrsEmpDll.SetValue("currency_code", txtCurrencyCode.Text)
        Call mrsEmpDll.SetValue("website", txtWebSite.Text)
        Call mrsEmpDll.SetValue("Bank_ac1", txtBankAcct1.Text)
        Call mrsEmpDll.SetValue("Bank_ac2", txtBankAcct2.Text)
        Call mrsEmpDll.SetValue("Bank_ac3", txtBankAcct3.Text)
        Call mrsEmpDll.SetValue("Credit_days", Trim(txtCreditTermId.Text))
        Call mrsEmpDll.SetValue("Upd_dt", getDateForDB(GetServerDate()))
        Call mrsEmpDll.SetValue("Upd_Userid", mP_User)
        Call mrsEmpDll.SetValue("Milk_Van", chkmilkvan.CheckState)
        Call mrsEmpDll.SetValue("CST_Eval", ChkCSTEval.CheckState)
        Call mrsEmpDll.SetValue("Tin_no", txtTinNo.Text)
        Call mrsEmpDll.SetValue("PANNo", Trim(TxtPanNo.Text))
        Call mrsEmpDll.SetValue("ServiceTaxNo", Trim(TxtServiceTaxNo.Text))
        Call mrsEmpDll.SetValue("SwiftCode", Trim(TxtSwiftNo.Text))
        Call mrsEmpDll.SetValue("IBANNo_SOT", Trim(TxtIBANNo.Text))
        Call mrsEmpDll.SetValue("BankName", Trim(TxtBankname.Text))
        Call mrsEmpDll.SetValue("BankAdd1", Trim(TxtBankAddress.Text))
        Call mrsEmpDll.SetValue("BAccNo", Trim(TxtAccNo.Text))
        Call mrsEmpDll.SetValue("ShipmentThruWh", chkShpmntThruWh.CheckState)
        Call mrsEmpDll.SetValue("FordLabelReqd", chktT6Label.CheckState)
        Call mrsEmpDll.SetValue("ConsigneeWiseLoc", chkConsigneeWiseLoc.CheckState)
        Call mrsEmpDll.SetValue("CST_amt", ChKAMtEval.CheckState)
        Call mrsEmpDll.SetValue("CSIex_INC", ChkCSI_GT.CheckState)
        Call mrsEmpDll.SetValue("Has_Trans_flag", ChkActive.CheckState)
        Call mrsEmpDll.SetValue("Group_Customer", chkGroup.CheckState)
        Call mrsEmpDll.SetValue("Shipping_Duration", Val(txtShipDur.Text))
        If ChkCSI_GT.CheckState = 1 Then
            Call mrsEmpDll.SetValue("CSIEX_GL", txtAcctLedgerExc.Text)
            Call mrsEmpDll.SetValue("CSIEX_SL", txtAccSubLedgerExc.Text)
        Else
            Call mrsEmpDll.SetValue("CSIEX_GL", "")
            Call mrsEmpDll.SetValue("CSIEX_SL", "")
        End If
        Call mrsEmpDll.SetValue("Dock_code", UCase(Trim(txtconsigneecode.Text)))
        Call mrsEmpDll.SetValue("Customer_EDICode", txtCustEDICode.Text.Trim().ToUpper())
        Call mrsEmpDll.SetValue("Plant_Code", txtPlantCode.Text.Trim().ToUpper())
        '10688280 -Add KAM Code in Customer Master.
        Call mrsEmpDll.SetValue("KAMCode", txtKAMName.Text.Trim())
        '10736222 — eMPro-- CT2 - ARE3 functionality
        Call mrsEmpDll.SetValue("CT2_REQD", chkCT2.CheckState)
        strCustType = ""
        If optDomestic.Checked = True Then
            strCustType = "L"
        End If
        If optOverseas.Checked = True Then
            strCustType = "O"
        End If
        If optInterState.Checked = True Then
            strCustType = "I"
        End If
        Call mrsEmpDll.SetValue("Cust_Type", strCustType.Trim().ToUpper())
        If lblGlobalCustCodeDesc.Text.Trim = String.Empty Then
            Call mrsEmpDll.SetValue("GLOBAL_CUATOMER_CODE", DBNull.Value)
        Else
            Call mrsEmpDll.SetValue("GLOBAL_CUATOMER_CODE", lblGlobalCustCodeDesc.Text.Trim())
        End If
        'GST CHANGES
        If OptB2B.Checked = True Then
            strCustClassificationType = "B"
        ElseIf optB2C.Checked = True Then
            strCustClassificationType = "C"
        ElseIf OptGovtAgency.Checked = True Then
            strCustClassificationType = "G"
        Else
            strCustClassificationType = "I"
        End If

        Call mrsEmpDll.SetValue("classificationtype", strCustClassificationType)
        Call mrsEmpDll.SetValue("GSTIN_Id", txtGSTINId.Text)
        Call mrsEmpDll.SetValue("GST_BILLSTATE_CODE", txtGSTBillState.Text.Trim())
        Call mrsEmpDll.SetValue("GST_BILLSTATE_DESC", lblGSTStateDesc.Text)

        Call mrsEmpDll.SetValue("ContactForMismatch_1", txtKeyContactName_1.Text)
        Call mrsEmpDll.SetValue("ContactForMismatch_PhoneNo_1", txtKeyContactPhNo_1.Text)
        Call mrsEmpDll.SetValue("ContactForMismatch_MobileNo_1", txtKeyContactMBNo_1.Text)
        Call mrsEmpDll.SetValue("ContactForMismatch_Email_1", txtKeyContactEmail_1.Text)
        Call mrsEmpDll.SetValue("ContactForMismatch_2", txtKeyContactName_2.Text)
        Call mrsEmpDll.SetValue("ContactForMismatch_PhoneNo_2", txtKeyContactPhNo_2.Text)
        Call mrsEmpDll.SetValue("ContactForMismatch_MobileNo_2", txtKeyContactMBNo_2.Text)
        Call mrsEmpDll.SetValue("ContactForMismatch_Email_2", txtKeyContactEmail_2.Text)
        '101459347 — Changes in Customer & Vendor Master
        If chkGSTINnotrequired.Checked = True Then
            GSTINnotrequired = 1
        Else
            GSTINnotrequired = 0
        End If
        Call mrsEmpDll.SetValue("GSTIN_notrequired", GSTINnotrequired)
        '101482956
        Call mrsEmpDll.SetValue("PARTY_TRN_NO", Trim(txtTRNNo.Text))

        'GST CHANGES
        If chkSEZ.Checked = True Then
            isSEZ = 1
        Else
            isSEZ = 0
        End If
        Call mrsEmpDll.SetValue("SEZ", isSEZ)
        Call InsertUpdateShippingDetails(Trim(txtCustCode.Text))

        ResetDatabaseConnection()
        mP_Connection.BeginTrans()
        mP_Connection.Execute(mstrInsertUpdate, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        mP_Connection.CommitTrans()
        Call mrsEmpDll.SetValue("AllowBarcodePrinting", Chkbarcodeprinting.CheckState)
        Call mrsEmpDll.SetValue("PRINT_METHOD", CmbPrintMethod.Text.ToString().Trim)
        mrsEmpDll.Update()
        mobjEmpDll.CConnection.CommitTransaction()
        mobjEmpDll.CConnection.CloseConnection()
        mrsEmpDll.CloseRecordset()
        ConfirmWindow(10049, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        mP_Connection.RollbackTrans()
        mobjEmpDll.CConnection.RollbackTransaction()
        mobjEmpDll.CConnection.CloseConnection()
        mrsEmpDll.CloseRecordset()
    End Function
    '*********************************************'
    'Author         :       Ananya Nath
    'Arguments      :       None
    'Return Value   :       None
    'Description    :       Deletes from customer_mst.
    '*********************************************'
    Public Function delete() As Object ' this func is used to Delete a rec.
        Dim strSql As String = String.Empty
        Dim rsgetdetail As ClsResultSetDB
        On Error GoTo ErrHandler
        rsgetdetail = New ClsResultSetDB
        strSql = "select Customer_Code,Account_ledger,Account_subLedger,Cust_Name,Cust_Vendor_Code,Excise_range,Commisionrate,Division,ECC_Code,LST_No,CST_No,Cust_Location,Ship_Contact_person,Ship_Person_desig,Ship_email_id,Ship_Address1,Ship_Address2,Ship_City,Ship_State,Ship_Pin,Ship_dist,Ship_Country,Ship_Phone,Ship_Fax,bill_Contact_person,Bill_Person_desig,bill_email_id,Bill_Address1,Bill_Address2,Bill_City,Bill_State,Bill_Pin,Bill_dist,Bill_Country,Bill_Phone,Bill_Fax,Office_Contact_person,Office_person_design,office_email_id,Office_Address1,Office_Address2,Office_City,Office_State,Office_dist,Office_Pin,Office_Country,Office_Phone,Office_fax,Website,customer_type,currency_code,Bank_ac1,Bank_ac2,Bank_ac3,Credit_days,Has_Trans_flag,ScheduleCode,CST_Eval,tin_no,Milk_Van,AllowExcessSchedule,CUST_ALIAS,DOCK_CODE,ShipmentThruWh,PANNo,ServiceTaxNo,SwiftCode,IBANNo_SOT,BAccNo,BankName,BankAdd1,Tool_GL,Tool_SL,ASN_REQD,CST_amt,CSIEX_Inc,CSIEX_GL,CSIEX_SL,Group_Customer,AllowBarCodePrinting,CONSIGNEEWISELOC,AllowASNPrinting,AllowASNTextGeneration,IS_AEROBIN_CUST,Customer_EDICode,Plant_Code,ASNFunctionCode,INVOICEAGAINSTAGREEMENTMST,CUST_TYPE from customer_mst Where Unit_Code='" & gstrUNITID & "' And customer_code='" & Trim(txtCustCode.Text) & "'"
        rsgetdetail.GetResult(strSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
        If rsgetdetail.GetNoRows > 0 Then
            strSql = "delete from Customer_Shipping_Dtl where customer_code='" & Trim(txtCustCode.Text) & "' and unit_code='" & gstrUNITID & "' " & vbCrLf
            strSql = strSql & "delete from Customer_mst where customer_code='" & Trim(txtCustCode.Text) & "' and unit_code='" & gstrUNITID & "'"
            ResetDatabaseConnection()
            mP_Connection.BeginTrans()
            mP_Connection.Execute(strSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            mP_Connection.CommitTrans()
            ConfirmWindow(10051, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
            rsgetdetail.ResultSetClose()
            rsgetdetail = Nothing
            Exit Function
        End If
ErrHandler:
        If Err.Number = -2147217873 Then
            mP_Connection.RollbackTrans()
            MessageBox.Show("This Customer Code is used in other Transaction,so unable to delete.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Function
        End If
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        mP_Connection.RollbackTrans()
    End Function
    '*********************************************'
    'Author         :       Ananya Nath
    'Arguments      :       None
    'Return Value   :       Customer Location
    'Description    :       Generates and Returns Customer Location
    '*********************************************'
    Public Function GenerateCustLoc() As String ' Generates Customer Location
        On Error GoTo ErrHandler
        Dim intVar As Short
        Dim merror As Excel.Error
        Dim strCustLoc As String
        mobjEmpDll.CConnection.OpenConnection(gstrDSNName, gstrDatabaseName)
        mrsEmpDll.OpenRecordset("select cust_location from customer_mst Where Unit_Code='" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenDynamic)
        If mrsEmpDll.Recordcount > 0 Then
            mrsEmpDll.OpenRecordset("select max(cast (substring(cust_location,2,len(cust_location)) as int)) from customer_mst Where Unit_Code='" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenDynamic)
            strCustLoc = mrsEmpDll.GetString(ADODB.StringFormatEnum.adClipString, 1)
            intVar = CDbl(strCustLoc) + 1
            If intVar <10 Then
                GenerateCustLoc="C00" & intVar
            ElseIf intVar < 100 Then
                GenerateCustLoc = "C0" & intVar
            Else
                GenerateCustLoc = "C" & intVar : End If
        Else
            GenerateCustLoc = "C001"
        End If
        mobjEmpDll.CConnection.CloseConnection()
        mrsEmpDll.CloseRecordset()
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        mobjEmpDll.CConnection.CloseConnection()
        mrsEmpDll.CloseRecordset()
    End Function
    '*********************************************'
    'Author         :       Ananya Nath
    'Arguments      :       None
    'Return Value   :       True/False
    'Description    :       Retursns True if the user enters an exiting Currency Code, else returns False
    '*********************************************'
    Public Function ValidateCur() As Boolean ' Checks for existion Currency Code
        On Error GoTo ErrHandler
        mobjEmpDll.CConnection.OpenConnection(gstrDSNName, gstrDatabaseName)
        mrsEmpDll.OpenRecordset("select currency_code, description from currency_mst where currency_code = '" & txtCurrencyCode.Text & "' and unit_code='" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockReadOnly)
        mrsEmpDll.Filter_Renamed = "currency_Code='" & txtCurrencyCode.Text & "'"
        If mrsEmpDll.Recordcount > 0 Then
            ValidateCur = True
            lblCurrDesc.Text = mrsEmpDll.GetFieldValue("description", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
        Else
            ValidateCur = False
        End If
        mrsEmpDll.CloseRecordset()
        mobjEmpDll.CConnection.CloseConnection()
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        mobjEmpDll.CConnection.CloseConnection()
        mrsEmpDll.CloseRecordset()
    End Function
    '*********************************************'
    'Author         :       Ananya Nath
    'Arguments      :       None
    'Return Value   :       True/False
    'Description    :       Returns True, if entered/selected Currency code has taken part in transaction
    '*********************************************'
    Public Function ExistInTransaction() As Boolean ' Checks for existance of rec. in transaction
        On Error GoTo ErrHandler
        Dim strTrans As Boolean
        mobjEmpDll.CConnection.OpenConnection(gstrDSNName, gstrDatabaseName)
        mrsEmpDll.OpenRecordset("select Has_Trans_flag from customer_mst Where Unit_Code='" & gstrUNITID & "' And customer_code = '" & txtCustCode.Text & "'", ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
        strTrans = mrsEmpDll.GetString(ADODB.StringFormatEnum.adClipString, 1)
        ExistInTransaction = strTrans
        mrsEmpDll.CloseRecordset()
        mobjEmpDll.CConnection.CloseConnection()
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        mobjEmpDll.CConnection.CloseConnection()
        mrsEmpDll.CloseRecordset()
    End Function
    '*********************************************'
    'Author         :       Ananya Nath
    'Arguments      :       None
    'Return Value   :       True/False
    'Description    :       Returns True, if Customer Code is existing, else False.
    '*********************************************'
    'Changes by : Samiksha Customer Master Authorization
    'Samiksha Credit limit changes
    'Samiksha branchcode changes
    Public Function ValCustomerCode() As Boolean ' Checks for Customer Code
        On Error GoTo ErrHandler
        Dim ms As String
        mobjEmpDll.CConnection.OpenConnection(gstrDSNName, gstrDatabaseName)
        '10690771
        Call mobjEmpDll.CRecordset.OpenRecordset("select Customer_Code,Account_ledger,Account_subLedger,Cust_Name,Cust_Vendor_Code,Shipping_Duration,Excise_range,Commisionrate,Division,ECC_Code,LST_No,CST_No,Cust_Location,Ship_Contact_person,Ship_Person_desig,Ship_email_id,Ship_Address1,Ship_Address2,Ship_City,Ship_State,Ship_Pin,Ship_dist,Ship_Country,Ship_Phone,Ship_Fax,bill_Contact_person,Bill_Person_desig,bill_email_id,Bill_Address1,Bill_Address2,Bill_City,Bill_State,Bill_Pin,Bill_dist,Bill_Country,Bill_Phone,Bill_Fax,Office_Contact_person,Office_person_design,office_email_id,Office_Address1,Office_Address2,Office_City,Office_State,Office_dist,Office_Pin,Office_Country,Office_Phone,Office_fax,Website,customer_type,currency_code,Bank_ac1,Bank_ac2,Bank_ac3,Credit_days,Has_Trans_flag,ScheduleCode,CST_Eval,tin_no,Milk_Van,AllowExcessSchedule,CUST_ALIAS,DOCK_CODE,ShipmentThruWh,PANNo,ServiceTaxNo,SwiftCode,IBANNo_SOT,BAccNo,BankName,BankAdd1,Tool_GL,Tool_SL,ASN_REQD,CST_amt,CSIEX_Inc,CSIEX_GL,CSIEX_SL,Group_Customer,AllowBarCodePrinting,CONSIGNEEWISELOC,AllowASNPrinting,AllowASNTextGeneration,IS_AEROBIN_CUST,Customer_EDICode,Plant_Code,ASNFunctionCode,INVOICEAGAINSTAGREEMENTMST,CUST_TYPE,CreditLimit,branchcode from customer_mst Where Unit_Code='" & gstrUNITID & "' And customer_code = '" & txtCustCode.Text & "' union select Customer_Code,Account_ledger,Account_subLedger,Cust_Name,Cust_Vendor_Code,Shipping_Duration,Excise_range,Commisionrate,Division,ECC_Code,LST_No,CST_No,Cust_Location,Ship_Contact_person,Ship_Person_desig,Ship_email_id,Ship_Address1,Ship_Address2,Ship_City,Ship_State,Ship_Pin,Ship_dist,Ship_Country,Ship_Phone,Ship_Fax,bill_Contact_person,Bill_Person_desig,bill_email_id,Bill_Address1,Bill_Address2,Bill_City,Bill_State,Bill_Pin,Bill_dist,Bill_Country,Bill_Phone,Bill_Fax,Office_Contact_person,Office_person_design,office_email_id,Office_Address1,Office_Address2,Office_City,Office_State,Office_dist,Office_Pin,Office_Country,Office_Phone,Office_fax,Website,customer_type,currency_code,Bank_ac1,Bank_ac2,Bank_ac3,Credit_days,Has_Trans_flag,ScheduleCode,CST_Eval,tin_no,Milk_Van,AllowExcessSchedule,CUST_ALIAS,DOCK_CODE,ShipmentThruWh,PANNo,ServiceTaxNo,SwiftCode,IBANNo_SOT,BAccNo,BankName,BankAdd1,Tool_GL,Tool_SL,ASN_REQD,CST_amt,CSIEX_Inc,CSIEX_GL,CSIEX_SL,Group_Customer,AllowBarCodePrinting,CONSIGNEEWISELOC,AllowASNPrinting,AllowASNTextGeneration,IS_AEROBIN_CUST,Customer_EDICode,Plant_Code,ASNFunctionCode,INVOICEAGAINSTAGREEMENTMST,CUST_TYPE,CreditLimit,branchcode from customer_mst_authorization Where Unit_Code='" & gstrUNITID & "' And customer_code = '" & txtCustCode.Text & "' and (ISNULL(authorization_status,'')='Return' or ISNULL(authorization_status,'')='')", ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockReadOnly)
        mobjEmpDll.CRecordset.Filter_Renamed = "customer_code='" & txtCustCode.Text & "'"
        If mobjEmpDll.CRecordset.Recordcount > 0 Then ValCustomerCode = True Else ValCustomerCode = False
        mobjEmpDll.CConnection.CloseConnection()
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        mobjEmpDll.CConnection.CloseConnection()
        mrsEmpDll.CloseRecordset()
    End Function
    Private Sub txtCreditTermId_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtCreditTermId.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        If (txtCreditTermId.Text) <> "" Then
            mobjEmpDll.CConnection.OpenConnection(gstrDSNName, gstrDatabaseName)
            mrsEmpDll.OpenRecordset("select CrTrm_TermId, crtrm_desc from Gen_CreditTrmMaster Where Unit_Code='" & gstrUNITID & "' And CrTrm_TermId = '" & txtCreditTermId.Text & "' and CrTrm_Status = '1'", ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockReadOnly)
            mrsEmpDll.Filter_Renamed = " CrTrm_TermId ='" & txtCreditTermId.Text & "'"
            If mrsEmpDll.Recordcount <= 0 Then
                ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                Me.txtCreditTermId.Text = ""
                Cancel = True
            Else
                lblCreditDesc.Text = mrsEmpDll.GetFieldValue("crtrm_desc", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                mrsEmpDll.CloseRecordset()
                mobjEmpDll.CConnection.CloseConnection()
                GoTo EventExitSub
            End If
            mrsEmpDll.CloseRecordset()
            mobjEmpDll.CConnection.CloseConnection()
        End If
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        mrsEmpDll.CloseRecordset()
        mobjEmpDll.CConnection.CloseConnection()
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtWebsite_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtWebSite.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If KeyAscii = 13 Then
            tabDetails.SelectedIndex = 0
            tabCustomer.SelectedIndex = 0
            Me.txtoffadd11.Focus()
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
    '*********************************************'
    'Author         :       Jyolsna VN
    'Arguments      :       None
    'Return Value   :       True/False
    'Description    :       Returns True, if Non Confirmed Customer Code is existing, else False.
    '*********************************************'
    'Changes : Samiksha Customer master authorization
    'Samiksha branchcode changes
    Public Function ValNonConfCustomerCode() As Boolean ' Checks for Non Confirmed Customer Code
        On Error GoTo ErrHandler
        Dim ms As String
        mobjEmpDll.CConnection.OpenConnection(gstrDSNName, gstrDatabaseName)
        Call mobjEmpDll.CRecordset.OpenRecordset("select Customer_Code,Cust_Name,Cust_Vendor_Code,Ship_Contact_Person,Ship_Person_Desig,Ship_Email_Id,Ship_Address1,Ship_Address2,Ship_City,Ship_State,Ship_Pin,Ship_Dist,Ship_Country,Ship_Phone,Ship_Fax,HO_Contact_Person,HO_Person_Desig,HO_Email_Id,HO_Address1,HO_Address2,HO_City,HO_State,HO_Pin,HO_Dist,HO_Country,HO_Phone,HO_Fax,Customer_Type,Currency_Code,Credit_Days,milk_van,CreditLimit,branchcode from emp_non_conf_cust_mst Where Unit_Code='" & gstrUNITID & "' And customer_code = '" & txtCustCode.Text & "'", ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockReadOnly)
        mobjEmpDll.CRecordset.Filter_Renamed = "customer_code='" & txtCustCode.Text & "'"
        If mobjEmpDll.CRecordset.Recordcount > 0 Then ValNonConfCustomerCode = True Else ValNonConfCustomerCode = False
        mobjEmpDll.CConnection.CloseConnection()
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        mobjEmpDll.CConnection.CloseConnection()
        mrsEmpDll.CloseRecordset()
    End Function
    '*********************************************'
    'Author         :       Jyolsna VN
    'Arguments      :       None
    'Return Value   :       None
    'Description    :       This Function is Used to display Non Confirmed Customer details, if user enters/selects an existing CustomerCode.
    '*********************************************'
    'Samiksha : Changes For Customer Master Authorization
    'Samiskha creditlimit changes
    'Samiksha branchcode changes
    Public Function NonConfCustDisplay() As Object ' Used to display the record.
        On Error GoTo ErrHandler
        Dim strdetails As String
        Dim strsql As String
        Dim strmilk As Boolean
        Dim strcheck As String
        If txtCustCode.Tag <> "" Then txtCustCode.Text = txtCustCode.Tag
        If Me.txtCustCode.Text <> "" Then
            mobjEmpDll.CConnection.OpenConnection(gstrDSNName, gstrDatabaseName)
            'mobjEmpDll.CRecordset.OpenRecordset("select Customer_Code,Cust_Name,Cust_Vendor_Code,Ship_Contact_Person,Ship_Person_Desig,Ship_Email_Id,Ship_Address1,Ship_Address2,Ship_City,Ship_State,Ship_Pin,Ship_Dist,Ship_Country,Ship_Phone,Ship_Fax,HO_Contact_Person,HO_Person_Desig,HO_Email_Id,HO_Address1,HO_Address2,HO_City,HO_State,HO_Pin,HO_Dist,HO_Country,HO_Phone,HO_Fax,Customer_Type,Currency_Code,Credit_Days,milk_van,CreditLimit from emp_non_conf_cust_mst Where Unit_Code='" & gstrUNITID & "' And Customer_code = '" & txtCustCode.Text & "' union select Customer_Code,Cust_Name,Cust_Vendor_Code,Ship_Contact_Person,Ship_Person_Desig,Ship_Email_Id,Ship_Address1,Ship_Address2,Ship_City,Ship_State,Ship_Pin,Ship_Dist,Ship_Country,Ship_Phone,Ship_Fax,HO_Contact_Person,HO_Person_Desig,HO_Email_Id,HO_Address1,HO_Address2,HO_City,HO_State,HO_Pin,HO_Dist,HO_Country,HO_Phone,HO_Fax,Customer_Type,Currency_Code,Credit_Days,milk_van,CreditLimit from customer_mst_authorization  Where Unit_Code='" & gstrUNITID & "' And Customer_code = '" & txtCustCode.Text & "'", ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            mobjEmpDll.CRecordset.OpenRecordset("select Customer_Code,Cust_Name,Cust_Vendor_Code,Ship_Contact_Person,Ship_Person_Desig,Ship_Email_Id,Ship_Address1,Ship_Address2,Ship_City,Ship_State,Ship_Pin,Ship_Dist,Ship_Country,Ship_Phone,Ship_Fax,HO_Contact_Person,HO_Person_Desig,HO_Email_Id,HO_Address1,HO_Address2,HO_City,HO_State,HO_Pin,HO_Dist,HO_Country,HO_Phone,HO_Fax,Customer_Type,Currency_Code,Credit_Days,milk_van,CreditLimit,branchcode from emp_non_conf_cust_mst Where Unit_Code='" & gstrUNITID & "' And Customer_code = '" & txtCustCode.Text & "'", ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

            If Not mobjEmpDll.CRecordset.EOF_Renamed Then
                mobjEmpDll.CRecordset.MoveFirst()
                txtCustCode.Text = mobjEmpDll.CRecordset.GetFieldValue("customer_code", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtCustName.Text = mobjEmpDll.CRecordset.GetFieldValue("cust_name", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtCustvendCode.Text = mobjEmpDll.CRecordset.GetFieldValue("cust_vendor_code", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                '********** HO Address
                txtHOadd1.Text = mobjEmpDll.CRecordset.GetFieldValue("ho_address1", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtHOadd2.Text = mobjEmpDll.CRecordset.GetFieldValue("ho_address2", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtHOcity.Text = mobjEmpDll.CRecordset.GetFieldValue("ho_city", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtHODist.Text = mobjEmpDll.CRecordset.GetFieldValue("ho_dist", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtHOPin.Text = mobjEmpDll.CRecordset.GetFieldValue("ho_pin", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtHOState.Text = mobjEmpDll.CRecordset.GetFieldValue("ho_state", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtHOCountry.Text = mobjEmpDll.CRecordset.GetFieldValue("ho_country", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtHOPhone.Text = mobjEmpDll.CRecordset.GetFieldValue("ho_phone", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtHOfax.Text = mobjEmpDll.CRecordset.GetFieldValue("ho_fax", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtHOContPer.Text = mobjEmpDll.CRecordset.GetFieldValue("ho_contact_person", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtHODesig.Text = mobjEmpDll.CRecordset.GetFieldValue("ho_person_desig", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtHOemail.Text = mobjEmpDll.CRecordset.GetFieldValue("ho_email_id", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                '********** Purchase Details
                txtCurrencyCode.Text = mobjEmpDll.CRecordset.GetFieldValue("currency_code", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                txtCreditTermId.Text = mobjEmpDll.CRecordset.GetFieldValue("credit_days", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                strmilk = mobjEmpDll.CRecordset.GetFieldValue("milk_van", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                If strmilk = True Then
                    Me.chkmilkvan.CheckState = System.Windows.Forms.CheckState.Checked
                Else
                    Me.chkmilkvan.CheckState = System.Windows.Forms.CheckState.Unchecked
                End If
                strcheck = mobjEmpDll.CRecordset.GetFieldValue("Customer_type", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
                If strcheck = "J" Then
                    Me.chkJobWork.CheckState = System.Windows.Forms.CheckState.Checked
                    Me.chkStandard.CheckState = System.Windows.Forms.CheckState.Unchecked
                ElseIf strcheck = "S" Then
                    Me.chkJobWork.CheckState = System.Windows.Forms.CheckState.Unchecked
                    Me.chkStandard.CheckState = System.Windows.Forms.CheckState.Checked
                ElseIf strcheck = "B" Then
                    Me.chkJobWork.CheckState = System.Windows.Forms.CheckState.Checked
                    Me.chkStandard.CheckState = System.Windows.Forms.CheckState.Checked
                End If
                Me.cmdgrpCustMst.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = True
                Me.cmdgrpCustMst.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = True
                Me.cmdgrpCustMst.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE) = True
                If Me.txtCustCode.Enabled = False Then Me.txtCustCode.Enabled = True

                'Samiksha Credit limit changes
                txtboxCreditLimit.Text = Convert.ToString(mobjEmpDll.CRecordset.GetFieldValue("CreditLimit", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString))
                'Samiksha branchcode changes
                txtboxBranchCode.Text = Convert.ToString(mobjEmpDll.CRecordset.GetFieldValue("branchcode", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString))
                Me.txtCustCode.Focus()

            End If
        End If
        mrsEmpDll.CloseRecordset()
        '------------------------------------------
        mrsEmpDll.OpenRecordset("select currency_code, description from currency_mst where currency_code = '" & txtCurrencyCode.Text & "' and unit_code='" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockReadOnly)
        If mrsEmpDll.Recordcount > 0 Then lblCurrDesc.Text = mrsEmpDll.GetFieldValue("description", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
        mrsEmpDll.CloseRecordset()
        '-------------------------------------------
        mrsEmpDll.OpenRecordset("select CrTrm_TermId, crtrm_desc from Gen_CreditTrmMaster Where Unit_Code='" & gstrUNITID & "' And  CrTrm_TermId = '" & txtCreditTermId.Text & "' and CrTrm_Status = '1'", ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockReadOnly)
        If mrsEmpDll.Recordcount > 0 Then lblCreditDesc.Text = mrsEmpDll.GetFieldValue("crtrm_desc", EMPDataBase.EMPDB.ADODataType.ADOVarChar, EMPDataBase.EMPDB.ADOCustomFormat.CustomString)
        mrsEmpDll.CloseRecordset()
        '--------------------------------------------
        mobjEmpDll.CConnection.CloseConnection()
        '-----------------------------------------
        Me.tabDetails.SelectedIndex = 0
        Me.tabCustomer.SelectedIndex = 2
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        mobjEmpDll.CConnection.RollbackTransaction()
        mobjEmpDll.CConnection.CloseConnection()
        mrsEmpDll.CloseRecordset()
    End Function
    '*********************************************'
    'Author         :       Jyolsna VN
    'Arguments      :       None
    'Return Value   :       None
    'Description    :       Used to Enable all form level controls.
    '*********************************************'
    Private Sub DisableForNonConfCust() ' All the form level controls will be enabled
        On Error GoTo ErrHandler
        txtCustLoc.Enabled = False
        txtCustLoc.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        txtWebSite.Enabled = False
        txtWebSite.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        txtAcctLedger.Enabled = False
        txtAcctLedger.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        txtAccSubLedger.Enabled = False
        txtAccSubLedger.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        cmdhelpLedger(0).Enabled = False
        cmdhelpSubLedger(2).Enabled = False
        txtBankAcct1.Enabled = False
        txtBankAcct1.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        txtBankAcct2.Enabled = False
        txtBankAcct2.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        txtBankAcct3.Enabled = False
        txtBankAcct3.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        txtoffadd11.Enabled = False
        txtoffadd11.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        txtOffAdd21.Enabled = False
        txtOffAdd21.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        txtOffCity1.Enabled = False
        txtOffCity1.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        txtOffDist1.Enabled = False
        txtOffDist1.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        txtOffPin1.Enabled = False
        txtOffPin1.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        txtOffState1.Enabled = False
        txtOffState1.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        txtOffCountry1.Enabled = False
        txtOffCountry1.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        txtOffPhone1.Enabled = False
        txtOffPhone1.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        txtOffFax1.Enabled = False
        txtOffFax1.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        txtContPer1.Enabled = False
        txtContPer1.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        txtDesig1.Enabled = False
        txtDesig1.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        txtEmail1.Enabled = False
        txtEmail1.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        txtBilladd1.Enabled = False
        txtBilladd1.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        txtBillAdd2.Enabled = False
        txtBillAdd2.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        txtbillCity(0).Enabled = False
        txtbillCity(0).BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        txtBillDist.Enabled = False
        txtBillDist.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        txtBillPin.Enabled = False
        txtBillPin.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        txtBillState.Enabled = False
        txtBillState.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        txtBillCountry.Enabled = False
        txtBillCountry.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        txtBillPhone.Enabled = False
        txtBillPhone.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        txtBillFax.Enabled = False
        txtBillFax.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        txtBillContPer.Enabled = False
        txtBillContPer.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        txtBillDesig.Enabled = False
        txtBillDesig.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        txtBillEmail.Enabled = False
        txtBillEmail.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        txtExciseRange.Enabled = False
        txtExciseRange.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        txtComRate.Enabled = False
        txtComRate.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        txtDivision.Enabled = False
        txtDivision.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        txtEcc.Enabled = False
        txtEcc.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        txtLst.Enabled = False
        txtLst.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        txtCst.Enabled = False
        txtCst.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        ChkCSTEval.CheckState = False
        ChkCSTEval.Enabled = False
        txtTinNo.Enabled = False
        txtTinNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        TxtPanNo.Enabled = False
        TxtPanNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        TxtServiceTaxNo.Enabled = False
        TxtServiceTaxNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        TxtSwiftNo.Enabled = False
        TxtSwiftNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        TxtIBANNo.Enabled = False
        TxtIBANNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        TxtBankname.Enabled = False
        TxtBankname.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        TxtBankAddress.Enabled = False
        TxtBankAddress.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        TxtAccNo.Enabled = False
        TxtAccNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    '*********************************************'
    'Author         :       Jyolsna VN
    'Arguments      :       None
    'Return Value   :       None
    'Description    :       Used to refresh and disable the form controls for Non Confirmed Customer Option.
    '*********************************************'
    Private Sub NonConfCustOpt()
        On Error GoTo ErrHandler
        Call RefreshForm()
        'Call Disabled()
        Call DisableForNonConfCust()
        cmdhelpCurrency(1).Enabled = False
        cmdhelpCreditTerms.Enabled = False
        BtnHelpKAM.Enabled = False   '10688280 -Add KAM Code in Customer Master.
        cmdhelpCustCode(1).Enabled = True
        txtCustCode.Enabled = True
        txtCustCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        cmdhelpCustCode(1).Enabled = True
        Me.cmdgrpCustMst.Revert()
        Me.cmdgrpCustMst.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
        Me.cmdgrpCustMst.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
        Me.cmdgrpCustMst.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL) = False
        Me.cmdgrpCustMst.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
        tabDetails.SelectedIndex = 0
        tabCustomer.SelectedIndex = 2
        txtCustCode.Focus()
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    '*********************************************'
    'Author         :       Jyolsna VN
    'Arguments      :       None
    'Return Value   :       None
    'Description    :       Used to refresh and disable the form controls for Confirmed Customer Option.
    '*********************************************'
    Private Sub ConfCustOpt()
        On Error GoTo ErrHandler
        Call DisableHOAddress()
        cmdhelpCurrency(1).Enabled = False
        cmdhelpCreditTerms.Enabled = False
        cmdhelpCustCode(1).Enabled = True
        BtnHelpKAM.Enabled = False   '10688280 -Add KAM Code in Customer Master.
        txtCustCode.Enabled = True
        txtCustCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        Me.cmdgrpCustMst.Revert()
        Me.cmdgrpCustMst.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
        Me.cmdgrpCustMst.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
        Me.cmdgrpCustMst.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE) = False
        Me.cmdgrpCustMst.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
        tabDetails.SelectedIndex = 0
        tabCustomer.SelectedIndex = 0
        txtCustCode.Focus()
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    '*********************************************'
    'Author         :       Jyolsna VN
    'Arguments      :       None
    'Return Value   :       None
    'Description    :       Used disable the HO Address details.
    '*********************************************'
    Private Sub DisableHOAddress()
        On Error GoTo ErrHandler
        txtHOadd1.Enabled = False
        txtHOadd1.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        txtHOadd2.Enabled = False
        txtHOadd2.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        txtHOcity.Enabled = False
        txtHOcity.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        txtHODist.Enabled = False
        txtHODist.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        txtHOPin.Enabled = False
        txtHOPin.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        txtHOState.Enabled = False
        txtHOState.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        txtHOCountry.Enabled = False
        txtHOCountry.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        txtHOPhone.Enabled = False
        txtHOPhone.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        txtHOfax.Enabled = False
        txtHOfax.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        txtHOContPer.Enabled = False
        txtHOContPer.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        txtHODesig.Enabled = False
        txtHODesig.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        txtHOemail.Enabled = False
        txtHOemail.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    '*********************************************'
    'Author         :       Jyolsna VN
    'Arguments      :       None
    'Return Value   :       Non Confirmed Customer Code
    'Description    :       Generates and returns Non Confirmed CustomerCode.
    '*********************************************'
    Public Function GenerateNonConfCustomerNo() As String ' used to generate Non Confirmed Customer No.
        On Error GoTo ErrHandler
        Dim lngVar As Integer
        Dim merror As Excel.Error
        Dim strCustCode As String
        Dim strValue As String
        Dim strVar As String
        mobjEmpDll.CConnection.OpenConnection(gstrDSNName, gstrDatabaseName)
        mobjEmpDll.CConnection.BeginTransaction()
        mrsEmpDll.OpenRecordset("select customer_code from emp_non_conf_cust_mst Where Unit_Code='" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
        If mrsEmpDll.Recordcount > 0 Then
            mrsEmpDll.OpenRecordset("select max(SUBSTRING(customer_code,3,6)) from emp_non_conf_cust_mst Where Unit_Code='" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
            mrsEmpDll.MoveFirst()
            strCustCode = mrsEmpDll.GetString(ADODB.StringFormatEnum.adClipString, 1, "-") '("customer_code", ADOVarChar, CustomString)
            lngVar = CInt(Mid(strCustCode, 1, 7)) + 1
            strVar = VB6.Format(lngVar, "00000#")
            GenerateNonConfCustomerNo = "NC" & strVar
        Else
            GenerateNonConfCustomerNo = "NC000001"
        End If
        mobjEmpDll.CConnection.CommitTransaction()
        mobjEmpDll.CConnection.CloseConnection()
        mrsEmpDll.CloseRecordset()
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        mobjEmpDll.CConnection.RollbackTransaction()
        mobjEmpDll.CConnection.CloseConnection()
        mrsEmpDll.CloseRecordset()
    End Function
    '*********************************************'
    'Author         :       Jyolsna VN
    'Arguments      :       None
    'Return Value   :       True or False
    'Description    :       Used to validate all the mandatory fields for Non Confirmed Customers have been properly entered.
    '*********************************************'
    Public Function ValNonConfBeforesave() As Boolean ' checks for validity
        On Error GoTo ErrHandler
        Dim strControls As String
        Dim strTag As String
        Dim getYYint As Short
        Dim strFocus1 As System.Windows.Forms.Control
        Dim lNo As Integer
        Dim intFirstPos As Short
        Dim intSecondPos As Short
        Dim intSecondPlace As Short
        ValNonConfBeforesave = True
        lNo = 1
        strFocus1 = Nothing
        strControls = ResolveResString(10059) & vbCrLf
        If Len(Trim(Me.txtCustName.Text)) < 1 Then
            strControls = strControls & vbCrLf & lNo & ". Customer Name." 'Add message to String
            lNo = lNo + 1
            strFocus1 = Me.txtCustName
            ValNonConfBeforesave = False
        End If
        If Len(Trim(txtCreditTermId.Text)) = 0 Then
            strControls = strControls & vbCrLf & lNo & ". Credit Term ID (Marketing Details)."
            lNo = lNo + 1
            tabDetails.SelectedIndex = 1
            If strFocus1 Is Nothing Then strFocus1 = txtCreditTermId
            ValNonConfBeforesave = False
        End If
        If Len(Trim(txtCurrencyCode.Text)) = 0 Or (ValidateCur()) = False Then
            strControls = strControls & vbCrLf & lNo & ". Currency Code (Marketing Details)."
            lNo = lNo + 1
            tabDetails.SelectedIndex = 1
            If strFocus1 Is Nothing Then strFocus1 = txtCurrencyCode
            ValNonConfBeforesave = False
        End If
        If chkJobWork.CheckState = 0 And chkStandard.CheckState = 0 Then
            strControls = strControls & vbCrLf & lNo & ". Customer Type (Marketing Details)."
            lNo = lNo + 1
            tabDetails.SelectedIndex = 1
            If strFocus1 Is Nothing Then strFocus1 = chkJobWork
            ValNonConfBeforesave = False
        End If
        'Samiksha Credit limit changes
        If isCreditLimitMandatory = True Then
            If txtboxCreditLimit.Text.Length <= 0 Then
                strControls = strControls & vbCrLf & lNo & ".Credit Limit is mandatory"
                If strFocus1 Is Nothing Then strFocus1 = txtboxCreditLimit
                lNo = lNo + 1
                ValNonConfBeforesave = False
            ElseIf (IsNumeric(txtboxCreditLimit.Text) = False) Then
                strControls = strControls & vbCrLf & lNo & ".Credit Limit must be Numeric"
                If strFocus1 Is Nothing Then strFocus1 = txtboxCreditLimit
                lNo = lNo + 1
                ValNonConfBeforesave = False
            ElseIf (Convert.ToDouble(txtboxCreditLimit.Text) <= 0) Then
                strControls = strControls & vbCrLf & lNo & ".Credit Limit must be greater than zero"
                If strFocus1 Is Nothing Then strFocus1 = txtboxCreditLimit
                lNo = lNo + 1
                ValNonConfBeforesave = False
            End If
        End If

        If ValNonConfBeforesave = False Then 'If any invalid field is there than set the focus on that field(after displaying message).                  MsgBox strControls, vbInformation, "eMPro"
            MsgBox(strControls, MsgBoxStyle.Information, "eMPro")
            strFocus1.Focus()
        End If
        strFocus1 = Nothing
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        mrsEmpDll.CloseRecordset()
        mobjEmpDll.CConnection.CloseConnection()
    End Function
    '*********************************************'
    'Author         :       Jyolsna VN
    'Arguments      :       None
    'Return Value   :       None
    'Description    :       Inserts a new row in the emp_non_conf_cust_mst table.
    '*********************************************'
    'Function to insert row in the emp_non_conf_cust_mst table.
    Public Function InsertNonConfDetails(ByVal strTblname As String) As Object
        On Error GoTo ErrHandler
        Dim strCustType As String = ""
        mobjEmpDll.CConnection.OpenConnection(gstrDSNName, gstrDatabaseName)
        'Samiksha : Customer Master Authorization
        'Samiksha Credit limit changes
        'Samiksha branchcode changes

        If strTblname = "customer_mst" Then
            Call mrsEmpDll.OpenRecordset("SELECT Customer_Code,Cust_Name,Cust_Vendor_Code,Ship_Contact_Person,Ship_Person_Desig,Ship_Email_Id,Ship_Address1,Ship_Address2,Ship_City,Ship_State,Ship_Pin,Ship_Dist,Ship_Country,Ship_Phone,Ship_Fax,HO_Contact_Person,HO_Person_Desig,HO_Email_Id,HO_Address1,HO_Address2,HO_City,HO_State,HO_Pin,HO_Dist,HO_Country,HO_Phone,HO_Fax,Customer_Type,Currency_Code,Credit_Days,milk_van,Ent_dt,Ent_Userid,Upd_dt,Upd_Userid,UNIT_CODE,CreditLimit,branchcode FROM emp_non_conf_cust_mst Where Unit_Code='" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)
            'ElseIf strTblname = "customer_mst_authorization" Then
            '    Call mrsEmpDll.OpenRecordset("SELECT Customer_Code,Cust_Name,Cust_Vendor_Code,Ship_Contact_Person,Ship_Person_Desig,Ship_Email_Id,Ship_Address1,Ship_Address2,Ship_City,Ship_State,Ship_Pin,Ship_Dist,Ship_Country,Ship_Phone,Ship_Fax,HO_Contact_Person,HO_Person_Desig,HO_Email_Id,HO_Address1,HO_Address2,HO_City,HO_State,HO_Pin,HO_Dist,HO_Country,HO_Phone,HO_Fax,Customer_Type,Currency_Code,Credit_Days,milk_van,CreditLimit FROM customer_mst_authorization Where Unit_Code='" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockOptimistic)

        End If

        mobjEmpDll.CConnection.BeginTransaction()
        mrsEmpDll.AddNew()
        With Me
            If chkJobWork.CheckState = 1 And chkStandard.CheckState = 0 Then
                strCustType = "J"
            Else
                If .chkStandard.CheckState = 1 And .chkJobWork.CheckState = 0 Then
                    strCustType = "S"
                ElseIf .chkJobWork.CheckState = 1 And .chkStandard.CheckState = 1 Then
                    strCustType = "B"
                End If
            End If
        End With
        Call mrsEmpDll.SetValue("customer_code", txtCustCode.Text)
        Call mrsEmpDll.SetValue("cust_name", txtCustName.Text)
        Call mrsEmpDll.SetValue("Cust_Vendor_code", txtCustvendCode.Text)
        '**** HO Address
        Call mrsEmpDll.SetValue("HO_Address1", txtHOadd1.Text)
        Call mrsEmpDll.SetValue("HO_address2", txtHOadd2.Text)
        Call mrsEmpDll.SetValue("HO_city", txtHOcity.Text)
        Call mrsEmpDll.SetValue("HO_State", txtHOState.Text)
        Call mrsEmpDll.SetValue("HO_Pin", txtHOPin.Text)
        Call mrsEmpDll.SetValue("HO_Country", txtHOCountry.Text)
        Call mrsEmpDll.SetValue("HO_Phone", txtHOPhone.Text)
        Call mrsEmpDll.SetValue("HO_dist", txtHODist.Text)
        Call mrsEmpDll.SetValue("HO_email_id", txtHOemail.Text)
        Call mrsEmpDll.SetValue("HO_fax", txtHOfax.Text)
        Call mrsEmpDll.SetValue("HO_contact_person", txtHOContPer.Text)
        Call mrsEmpDll.SetValue("HO_person_desig", txtHODesig.Text)
        '*******
        Call mrsEmpDll.SetValue("Customer_type", strCustType)

        'Field added by prashant rajpal on 29-12-2005

        Call mrsEmpDll.SetValue("milk_van", chkmilkvan.CheckState)
        Call mrsEmpDll.SetValue("currency_code", txtCurrencyCode.Text)
        Call mrsEmpDll.SetValue("Credit_days", Trim(txtCreditTermId.Text))
        Call mrsEmpDll.SetValue("Ent_dt", getDateForDB(GetServerDate))
        Call mrsEmpDll.SetValue("Ent_Userid", mP_User)
        Call mrsEmpDll.SetValue("Upd_dt", getDateForDB(GetServerDate))
        Call mrsEmpDll.SetValue("Upd_Userid", mP_User)
        Call mrsEmpDll.SetValue("UNIT_CODE", gstrUNITID)


        'Samiksha Credit limit changes
        Call mrsEmpDll.SetValue("CreditLimit", If(isCreditLimitMandatory = True, Convert.ToDouble(txtboxCreditLimit.Text), Convert.ToDouble(0)))
        'Samiksha branchcode changes
        Call mrsEmpDll.SetValue("branchcode", If(Len(txtboxCreditLimit.Text) = 0, Convert.ToString(""), Convert.ToString(txtboxCreditLimit.Text)))

        mrsEmpDll.Update()
        mobjEmpDll.CConnection.CommitTransaction()
        mrsEmpDll.CloseRecordset()
        mobjEmpDll.CConnection.CloseConnection()
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        mobjEmpDll.CConnection.RollbackTransaction()
        mobjEmpDll.CConnection.CloseConnection()
        mrsEmpDll.CloseRecordset()
    End Function
    '*********************************************'
    'Author         :       Jyolsna vn
    'Arguments      :       None
    'Return Value   :       None
    'Description    :       Updates Non Confirmed Customer details.
    '*********************************************'
    'Samiksha  Creditlimit Changes
    'Samiksha branchcode changes
    Public Function UpdateNonConfDetails(ByVal strTblName As String) As Object 'This Function Updates the emp_non_conf_cust_mst
        On Error GoTo ErrHandler
        Dim strCustType As String = ""
        mobjEmpDll.CConnection.OpenConnection(gstrDSNName, gstrDatabaseName)
        mobjEmpDll.CConnection.BeginTransaction()
        If strTblName = "customer_mst" Then
            Call mrsEmpDll.OpenRecordset("select Customer_Code,Cust_Name,Cust_Vendor_Code,Ship_Contact_Person,Ship_Person_Desig,Ship_Email_Id,Ship_Address1,Ship_Address2,Ship_City,Ship_State,Ship_Pin,Ship_Dist,Ship_Country,Ship_Phone,Ship_Fax,HO_Contact_Person,HO_Person_Desig,HO_Email_Id,HO_Address1,HO_Address2,HO_City,HO_State,HO_Pin,HO_Dist,HO_Country,HO_Phone,HO_Fax,Customer_Type,Currency_Code,Credit_Days,milk_van,CreditLimit,branchcode from emp_non_conf_cust_mst Where Unit_Code='" & gstrUNITID & "' And customer_code = '" & txtCustCode.Text & "' ", ADODB.CursorTypeEnum.adOpenDynamic)
            With Me
                If .chkJobWork.CheckState = 1 And .chkStandard.CheckState = 0 Then
                    strCustType = "J"
                Else
                    If .chkStandard.CheckState = 1 And .chkJobWork.CheckState = 0 Then
                        strCustType = "S"
                    ElseIf .chkJobWork.CheckState = 1 And .chkStandard.CheckState = 1 Then
                        strCustType = "B"
                    End If
                End If
            End With
            Call mrsEmpDll.SetValue("customer_code", txtCustCode.Text)
            Call mrsEmpDll.SetValue("cust_name", txtCustName.Text)
            Call mrsEmpDll.SetValue("Cust_Vendor_code", txtCustvendCode.Text)
            '**** HO Address
            Call mrsEmpDll.SetValue("HO_Address1", txtHOadd1.Text)
            Call mrsEmpDll.SetValue("HO_address2", txtHOadd2.Text)
            Call mrsEmpDll.SetValue("HO_city", txtHOcity.Text)
            Call mrsEmpDll.SetValue("HO_State", txtHOState.Text)
            Call mrsEmpDll.SetValue("HO_Pin", txtHOPin.Text)
            Call mrsEmpDll.SetValue("HO_Country", txtHOCountry.Text)
            Call mrsEmpDll.SetValue("HO_Phone", txtHOPhone.Text)
            Call mrsEmpDll.SetValue("HO_dist", txtHODist.Text)
            Call mrsEmpDll.SetValue("HO_email_id", txtHOemail.Text)
            Call mrsEmpDll.SetValue("HO_fax", txtHOfax.Text)
            Call mrsEmpDll.SetValue("HO_contact_person", txtHOContPer.Text)
            Call mrsEmpDll.SetValue("HO_person_desig", txtHODesig.Text)
            '*******
            Call mrsEmpDll.SetValue("Customer_type", strCustType)
            Call mrsEmpDll.SetValue("currency_code", txtCurrencyCode.Text)
            Call mrsEmpDll.SetValue("Credit_days", Trim(txtCreditTermId.Text))
            Call mrsEmpDll.SetValue("Upd_dt", getDateForDB(GetServerDate))
            Call mrsEmpDll.SetValue("Upd_Userid", mP_User)
            Call mrsEmpDll.SetValue("milk_van", chkmilkvan.CheckState)
            'Samiksha Credit limit changes
            Call mrsEmpDll.SetValue("CreditLimit", If(isCreditLimitMandatory = True, Convert.ToDouble(txtboxCreditLimit.Text), Convert.ToDouble(0)))
            'Samiksh branchcode changes
            Call mrsEmpDll.SetValue("branchcode", If(Len(txtboxBranchCode.Text) = 0, "", Convert.ToString(txtboxBranchCode.Text)))

            'ElseIf strTblName = "customer_mst_authorization" Then
            '    Call mrsEmpDll.OpenRecordset("select Customer_Code,Cust_Name,Cust_Vendor_Code,Ship_Contact_Person,Ship_Person_Desig,Ship_Email_Id,Ship_Address1,Ship_Address2,Ship_City,Ship_State,Ship_Pin,Ship_Dist,Ship_Country,Ship_Phone,Ship_Fax,HO_Contact_Person,HO_Person_Desig,HO_Email_Id,HO_Address1,HO_Address2,HO_City,HO_State,HO_Pin,HO_Dist,HO_Country,HO_Phone,HO_Fax,Customer_Type,Currency_Code,Credit_Days,milk_van,authorization_status,authorization_remarks,CreditLimit from customer_mst_authorization Where Unit_Code='" & gstrUNITID & "' And customer_code = '" & txtCustCode.Text & "' ", ADODB.CursorTypeEnum.adOpenDynamic)
            '    With Me
            '        If .chkJobWork.CheckState = 1 And .chkStandard.CheckState = 0 Then
            '            strCustType = "J"
            '        Else
            '            If .chkStandard.CheckState = 1 And .chkJobWork.CheckState = 0 Then
            '                strCustType = "S"
            '            ElseIf .chkJobWork.CheckState = 1 And .chkStandard.CheckState = 1 Then
            '                strCustType = "B"
            '            End If
            '        End If
            '    End With
            '    Call mrsEmpDll.SetValue("customer_code", txtCustCode.Text)
            '    Call mrsEmpDll.SetValue("cust_name", txtCustName.Text)
            '    Call mrsEmpDll.SetValue("Cust_Vendor_code", txtCustvendCode.Text)
            '    '**** HO Address
            '    Call mrsEmpDll.SetValue("HO_Address1", txtHOadd1.Text)
            '    Call mrsEmpDll.SetValue("HO_address2", txtHOadd2.Text)
            '    Call mrsEmpDll.SetValue("HO_city", txtHOcity.Text)
            '    Call mrsEmpDll.SetValue("HO_State", txtHOState.Text)
            '    Call mrsEmpDll.SetValue("HO_Pin", txtHOPin.Text)
            '    Call mrsEmpDll.SetValue("HO_Country", txtHOCountry.Text)
            '    Call mrsEmpDll.SetValue("HO_Phone", txtHOPhone.Text)
            '    Call mrsEmpDll.SetValue("HO_dist", txtHODist.Text)
            '    Call mrsEmpDll.SetValue("HO_email_id", txtHOemail.Text)
            '    Call mrsEmpDll.SetValue("HO_fax", txtHOfax.Text)
            '    Call mrsEmpDll.SetValue("HO_contact_person", txtHOContPer.Text)
            '    Call mrsEmpDll.SetValue("HO_person_desig", txtHODesig.Text)
            '    '*******
            '    Call mrsEmpDll.SetValue("Customer_type", strCustType)
            '    Call mrsEmpDll.SetValue("currency_code", txtCurrencyCode.Text)
            '    Call mrsEmpDll.SetValue("Credit_days", Trim(txtCreditTermId.Text))
            '    Call mrsEmpDll.SetValue("Upd_dt", getDateForDB(GetServerDate))
            '    Call mrsEmpDll.SetValue("Upd_Userid", mP_User)
            '    Call mrsEmpDll.SetValue("milk_van", chkmilkvan.CheckState)
            '    Call mrsEmpDll.SetValue("authorization_status", vbNull)
            '    Call mrsEmpDll.SetValue("Authorization_Remark", vbNull)
            '    'Samiksha Credit limit changes
            '    Call mrsEmpDll.SetValue("CreditLimit", If(isCreditLimitMandatory = True, Convert.ToDouble(txtboxCreditLimit.Text), Convert.ToDouble(0)))
        End If
        'Dim authStatus As String = getAuthorisationstatus()
        'Dim strSql As String = ""
        'If authStatus = "Auth" Then
        '    strSql = "Delete from customer_mst where customer_code='" & txtCustCode.Text & " and UNIT_CODE='" & gstrUNITID & "'"
        '    SqlConnectionclass.ExecuteNonQuery(strSql)
        'End If

        mrsEmpDll.Update()
        mobjEmpDll.CConnection.CommitTransaction()
        mobjEmpDll.CConnection.CloseConnection()
        mrsEmpDll.CloseRecordset()
        ConfirmWindow(10049, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        mobjEmpDll.CConnection.RollbackTransaction()
        mobjEmpDll.CConnection.CloseConnection()
        mrsEmpDll.CloseRecordset()
    End Function
    '*********************************************'
    'Author         :       Jyolsna VN
    'Arguments      :       None
    'Return Value   :       None
    'Description    :       Deletes from emp_non_conf_cust_mst.
    '*********************************************'
    Public Function DeleteNonConfCust() As Object ' this func is used to Delete a record.
        On Error GoTo ErrHandler
        mobjEmpDll.CConnection.OpenConnection(gstrDSNName, gstrDatabaseName)
        mobjEmpDll.CConnection.BeginTransaction()
        Call mrsEmpDll.OpenRecordset("select Customer_Code,Cust_Name,Cust_Vendor_Code,Ship_Contact_Person,Ship_Person_Desig,Ship_Email_Id,Ship_Address1,Ship_Address2,Ship_City,Ship_State,Ship_Pin,Ship_Dist,Ship_Country,Ship_Phone,Ship_Fax,HO_Contact_Person,HO_Person_Desig,HO_Email_Id,HO_Address1,HO_Address2,HO_City,HO_State,HO_Pin,HO_Dist,HO_Country,HO_Phone,HO_Fax,Customer_Type,Currency_Code,Credit_Days,milk_van from emp_non_conf_cust_mst Where Unit_Code='" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic)
        mrsEmpDll.Filter_Renamed = "customer_code='" & txtCustCode.Text & "'"
        If Not mrsEmpDll.EOF_Renamed Then
            mrsEmpDll.Delete()
            mrsEmpDll.UpdateBatch()
            mobjEmpDll.CConnection.CommitTransaction()
            mobjEmpDll.CConnection.CloseConnection()
            mrsEmpDll.CloseRecordset()
        End If
        ConfirmWindow(10051, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        mobjEmpDll.CConnection.RollbackTransaction()
        mobjEmpDll.CConnection.CloseConnection()
        mrsEmpDll.CloseRecordset()
    End Function
    Private Sub txtPANNo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TxtPanNo.Enter
        On Error GoTo ErrHandler
        TxtPanNo.SelectionStart = 0
        TxtPanNo.SelectionLength = Len(TxtPanNo.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub TxtServiceTaxNo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TxtServiceTaxNo.Enter
        On Error GoTo ErrHandler
        TxtServiceTaxNo.SelectionStart = 0
        TxtServiceTaxNo.SelectionLength = Len(TxtServiceTaxNo.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub TxtSwiftNo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TxtSwiftNo.Enter
        On Error GoTo ErrHandler
        TxtSwiftNo.SelectionStart = 0
        TxtSwiftNo.SelectionLength = Len(TxtSwiftNo.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub TxtIBANNo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TxtIBANNo.Enter
        On Error GoTo ErrHandler
        TxtIBANNo.SelectionStart = 0
        TxtIBANNo.SelectionLength = Len(TxtIBANNo.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtBankName_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TxtBankname.Enter
        On Error GoTo ErrHandler
        TxtBankname.SelectionStart = 0
        TxtBankname.SelectionLength = Len(TxtBankname.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub TxtBankAddress_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TxtBankAddress.Enter
        On Error GoTo ErrHandler
        TxtBankAddress.SelectionStart = 0
        TxtBankAddress.SelectionLength = Len(TxtBankAddress.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub TxtAccNo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TxtAccNo.Enter
        On Error GoTo ErrHandler
        TxtAccNo.SelectionStart = 0
        TxtAccNo.SelectionLength = Len(TxtAccNo.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub TxtAccNo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles TxtAccNo.KeyPress, TxtSwiftNo.KeyPress, TxtIBANNo.KeyPress, TxtServiceTaxNo.KeyPress, TxtPanNo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If (KeyAscii < 97 Or KeyAscii > 122) And (KeyAscii < 65 Or KeyAscii > 90) And (KeyAscii <> 32) And (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8) Then
            KeyAscii = 0
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
    Private Sub TxtServiceTaxNo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles TxtServiceTaxNo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If (KeyAscii < 97 Or KeyAscii > 122) And (KeyAscii < 65 Or KeyAscii > 90) And (KeyAscii <> 32) And (KeyAscii <> 8) And (KeyAscii < 48 Or KeyAscii > 57) Then
            KeyAscii = 0
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



    Private Sub TxtSwiftNo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles TxtSwiftNo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If (KeyAscii < 97 Or KeyAscii > 122) And (KeyAscii < 65 Or KeyAscii > 90) And (KeyAscii <> 32) And (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8) Then
            KeyAscii = 0
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

    Private Sub TxtIBANNo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles TxtIBANNo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If (KeyAscii < 97 Or KeyAscii > 122) And (KeyAscii < 65 Or KeyAscii > 90) And (KeyAscii <> 32) And (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8) Then
            KeyAscii = 0
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
    Private Sub InitializeShippingSpread()
        '-------------------------------------------------------------------------------------------
        ' Function      : Intialize Spread with all the required columns
        '--------------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        With spShippingAddess
            .MaxRows = 0
            .Row = 0
            .MaxCols = MaxGridHdrCols
            .UnitType = FPSpreadADO.UnitTypeConstants.UnitTypeTwips
            .Font = VB6.FontChangeName(.Font, Me.Font.Name)
            .CursorStyle = FPSpreadADO.CursorStyleConstants.CursorStyleArrow
            .Col = enmshipdetail.VAL_DEFAULT : .Text = "Default"
            .set_ColWidth(enmshipdetail.VAL_DEFAULT, 600)
            .ColHidden = False
            .Col = enmshipdetail.VAL_INACTIVE : .Text = "Inactive"
            .set_ColWidth(enmshipdetail.VAL_INACTIVE, 650)
            .ColHidden = False
            .Col = enmshipdetail.VAL_SHIPCODE : .Text = "Shipping Code"
            .set_ColWidth(enmshipdetail.VAL_SHIPCODE, 1200)
            .ColHidden = False
            .Col = enmshipdetail.VAL_SHIPDESC : .Text = "Description"
            .set_ColWidth(enmshipdetail.VAL_SHIPDESC, 1300)
            .ColHidden = False
            .Col = enmshipdetail.VAL_SHIPADD1 : .Text = "Ship Address1"
            .set_ColWidth(enmshipdetail.VAL_SHIPADD1, 1800)
            .ColHidden = False
            .Col = enmshipdetail.VAL_SHIPADD2 : .Text = "Ship Address2"
            .set_ColWidth(enmshipdetail.VAL_SHIPADD2, 1800)
            .ColHidden = False
            .Col = enmshipdetail.VAL_CITY : .Text = "City"
            .set_ColWidth(enmshipdetail.VAL_CITY, 1000)
            .ColHidden = False
            .Col = enmshipdetail.VAL_DISCT : .Text = "District"
            .set_ColWidth(enmshipdetail.VAL_DISCT, 1100)
            .ColHidden = False
            .Col = enmshipdetail.VAL_STATE : .Text = "State"
            .set_ColWidth(enmshipdetail.VAL_STATE, 1000)
            .ColHidden = False
            'GST CHANGES

            'Abhijit
            .Col = enmshipdetail.VAL_SHIP_GSTIN_ID : .Text = "GSTIN ID"
            .set_ColWidth(enmshipdetail.VAL_SHIP_GSTIN_ID, 1500)
            .ColHidden = False
            'Abhijit

            .Col = enmshipdetail.VAL_GSTSTATECODE : .Text = "GST State Code "
            .set_ColWidth(enmshipdetail.VAL_GSTSTATECODE, 1500)
            .ColHidden = False
            .Col = enmshipdetail.VAL_GSTSTATEDESC : .Text = "GST State Name"
            .set_ColWidth(enmshipdetail.VAL_GSTSTATEDESC, 1500)
            .ColHidden = False
            'GST CHANGES
            .Col = enmshipdetail.VAL_COUNTRY : .Text = "Country"
            .set_ColWidth(enmshipdetail.VAL_COUNTRY, 1100)
            .ColHidden = False
            .Col = enmshipdetail.VAL_Pin : .Text = "PIN Code"
            .set_ColWidth(enmshipdetail.VAL_Pin, 1100)
            .ColHidden = False
            .Col = enmshipdetail.VAL_PHONE : .Text = "Phone"
            .set_ColWidth(enmshipdetail.VAL_PHONE, 1100)
            .ColHidden = False
            .Col = enmshipdetail.VAL_FAX : .Text = "Fax No."
            .set_ColWidth(enmshipdetail.VAL_FAX, 1100)
            .ColHidden = False
            .Col = enmshipdetail.VAL_EMAILID : .Text = "Email Id."
            .set_ColWidth(enmshipdetail.VAL_EMAILID, 1500)
            .ColHidden = False
            .Col = enmshipdetail.VAL_CONTACTPERSON : .Text = "Contact Person"
            .set_ColWidth(enmshipdetail.VAL_CONTACTPERSON, 1500)
            .ColHidden = False
            .Col = enmshipdetail.VAL_DESIGNATION : .Text = "Designation"
            .set_ColWidth(enmshipdetail.VAL_DESIGNATION, 1500)
            .ColHidden = False
            .Col = enmshipdetail.VAL_DISTANCE_FROM_UNIT : .Text = "Distance From Unit(Km.)"
            .set_ColWidth(enmshipdetail.VAL_DISTANCE_FROM_UNIT, 1900)
            .ColHidden = False
            .Col = enmshipdetail.VAL_INACTIVEDATE : .Text = "Inactive Date"
            .set_ColWidth(enmshipdetail.VAL_INACTIVEDATE, 1500)
            .ColHidden = False
            .Col = enmshipdetail.VAL_DELETE : .Text = "Delete"
            .set_ColWidth(enmshipdetail.VAL_DELETE, 600)
            .ColHidden = True
            .ColsFrozen = 3
        End With
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub AddBlankRowinGrid()
        '-------------------------------------------------------------------------------------------
        ' Function      : Add blank row while Enter Press at the last column of the Grid
        '--------------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        With spShippingAddess
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .set_RowHeight(.MaxRows, 300)
            .Col = enmshipdetail.VAL_DEFAULT
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
            .TypeCheckCenter = True
            .Col = enmshipdetail.VAL_INACTIVE
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
            .TypeCheckCenter = True
            .Col = enmshipdetail.VAL_SHIPCODE
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Col = enmshipdetail.VAL_SHIPDESC
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .TypeMaxEditLen = 50
            .Col = enmshipdetail.VAL_SHIPADD1
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .TypeMaxEditLen = 60
            .Col = enmshipdetail.VAL_SHIPADD2
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .TypeMaxEditLen = 60
            .Col = enmshipdetail.VAL_CITY
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .TypeMaxEditLen = 50
            .Col = enmshipdetail.VAL_DISCT
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .TypeMaxEditLen = 50
            .Col = enmshipdetail.VAL_STATE
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .TypeMaxEditLen = 50
            'GST CHANGES

            'Abhijit
            .Col = enmshipdetail.VAL_SHIP_GSTIN_ID
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .TypeMaxEditLen = 50
            'Abhijit

            .Col = enmshipdetail.VAL_GSTSTATECODE
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .TypeMaxEditLen = 50
            .Col = enmshipdetail.VAL_GSTSTATEDESC
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .TypeMaxEditLen = 50
            'GST CHANGES
            .Col = enmshipdetail.VAL_COUNTRY
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .TypeMaxEditLen = 50
            .Col = enmshipdetail.VAL_Pin
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .TypeMaxEditLen = 20
            .Col = enmshipdetail.VAL_PHONE
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .TypeMaxEditLen = 20
            .Col = enmshipdetail.VAL_FAX
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .TypeMaxEditLen = 20
            .Col = enmshipdetail.VAL_EMAILID
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .TypeMaxEditLen = 100
            .Col = enmshipdetail.VAL_CONTACTPERSON
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .TypeMaxEditLen = 50
            .Col = enmshipdetail.VAL_DESIGNATION
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .TypeMaxEditLen = 50
            .Col = enmshipdetail.VAL_DISTANCE_FROM_UNIT
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeNumber
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .TypeNumberMin = 0.0
            .TypeNumberMax = CDbl("99999.99")
            .Col = enmshipdetail.VAL_INACTIVEDATE
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeDate
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Lock = True
            .Col = enmshipdetail.VAL_DELETE
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
            .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
            .Lock = True
        End With
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Function GenerateShippingCode(ByVal pstrCustomerCode As String) As String
        '-------------------------------------------------------------------------------------------
        ' Author        : Manoj Kr. Vaish
        ' Arguments     : Customer Code
        ' Return Value  : String
        ' Function      : Generate Shipping Code from Customer_Shipping_Dtl
        ' Datetime      : 29 Aug 2008
        'Issue ID       : eMpro-20080828-21178
        '----------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim intVar As Short
        Dim merror As Excel.Error
        Dim strShipCode As String
        mobjEmpDll.CConnection.OpenConnection(gstrDSNName, gstrDatabaseName)
        mobjEmpDll.CConnection.BeginTransaction()
        mrsEmpDll.OpenRecordset("select Shipping_Code from Customer_Shipping_Dtl Where Unit_Code='" & gstrUNITID & "' And customer_code='" & Trim(pstrCustomerCode) & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
        If mrsEmpDll.Recordcount > 0 Then
            mrsEmpDll.OpenRecordset("select max(cast (substring(Shipping_Code,2,len(Shipping_Code)) as int)) from Customer_Shipping_Dtl Where Unit_Code='" & gstrUNITID & "' And customer_code='" & Trim(pstrCustomerCode) & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
            strShipCode = mrsEmpDll.GetString(ADODB.StringFormatEnum.adClipString, 1)
            intVar = CDbl(strShipCode) + 1
            If intVar >= 1 And intVar <= 9 Then
                GenerateShippingCode = "S000000" & intVar
            ElseIf intVar > 9 And intVar <= 99 Then
                GenerateShippingCode = "S00000" & intVar
            ElseIf intVar > 99 And intVar <= 999 Then
                GenerateShippingCode = "S0000" & intVar
            ElseIf intVar > 999 And intVar <= 9999 Then
                GenerateShippingCode = "S000" & intVar
            ElseIf intVar > 9999 And intVar <= 99999 Then
                GenerateShippingCode = "S00" & intVar
            ElseIf intVar > 99999 And intVar <= 9999999 Then
                GenerateShippingCode = "S0" & intVar
            ElseIf intVar > 999999 And intVar <= 99999999 Then
                GenerateShippingCode = "S" & intVar
            End If
        Else
            GenerateShippingCode = "S0000001"
        End If
        mobjEmpDll.CConnection.CommitTransaction()
        mobjEmpDll.CConnection.CloseConnection()
        mrsEmpDll.CloseRecordset()
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        mobjEmpDll.CConnection.CloseConnection()
        mrsEmpDll.CloseRecordset()
    End Function
    Private Function GetNextShippingCode() As String
        '-------------------------------------------------------------------------------------------
        ' Function      : Generate Shipping Code from Customer_Shipping_Dtl
        '----------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim intVar As Short
        Dim merror As Excel.Error
        Dim lngShipCode As Integer
        Dim varShipcode As Object
        With spShippingAddess
            If .MaxRows >= 1 Then
                .Col = enmshipdetail.VAL_SHIPCODE
                .Row = .MaxRows
                varShipcode = Trim(.Text)
                lngShipCode = CInt(Mid(varShipcode, 2, Len(varShipcode)))
                intVar = lngShipCode + 1
                If intVar >= 1 And intVar <= 9 Then
                    GetNextShippingCode = "S000000" & intVar
                ElseIf intVar > 9 And intVar <= 99 Then
                    GetNextShippingCode = "S00000" & intVar
                ElseIf intVar > 99 And intVar <= 999 Then
                    GetNextShippingCode = "S0000" & intVar
                ElseIf intVar > 999 And intVar <= 9999 Then
                    GetNextShippingCode = "S000" & intVar
                ElseIf intVar > 9999 And intVar <= 99999 Then
                    GetNextShippingCode = "S00" & intVar
                ElseIf intVar > 99999 And intVar <= 9999999 Then
                    GetNextShippingCode = "S0" & intVar
                ElseIf intVar > 999999 And intVar <= 99999999 Then
                    GetNextShippingCode = "S" & intVar
                End If
            End If
        End With
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub InsertUpdateShippingDetails_DotNet(ByVal pstrCustomerCode As String)
        '-------------------------------------------------------------------------------------------
        ' Function      : Insert and Update the Customer Shipping Detail
        '----------------------------------------------------------------------------------------------
        Dim intLoopCounter As Short
        Dim varDelete As Object = Nothing
        Dim strQuery As String
        Dim varShipcode As Object = Nothing
        Dim varShipDesc As Object = Nothing
        Dim varShipAdd1 As Object = Nothing
        Dim varShipAdd2 As Object = Nothing
        Dim varCity As Object = Nothing
        Dim varDist As Object = Nothing
        Dim varState As Object = Nothing
        Dim varCountry As Object = Nothing
        Dim varPinCode As Object = Nothing
        Dim varPhone As Object = Nothing
        Dim varFax As Object = Nothing
        Dim varEmailId As Object = Nothing
        Dim varContact As Object = Nothing
        Dim varDesignation As Object = Nothing
        Dim varInactiveDate As Object = Nothing
        Dim intDefalut As Short
        Dim intInactive As Short
        Dim varGSTStatecode As Object = Nothing
        Dim varGSTStatename As Object = Nothing
        Dim VAL_SHIP_GSTIN_ID As Object = Nothing
        Dim VAR_DISTANCE_FROM_UNIT As Object = Nothing
        Dim CntShippingDetails As Integer = Nothing

        Try
            If cmdgrpCustMst.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Or cmdgrpCustMst.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                With spShippingAddess
                    mstrInsertUpdate = ""
                    For intLoopCounter = 1 To .MaxRows
                        varDelete = Nothing
                        Call .GetText(enmshipdetail.VAL_DELETE, intLoopCounter, varDelete)
                        If UCase(varDelete) <> "D" Then
                            .Row = intLoopCounter
                            .Col = enmshipdetail.VAL_DEFAULT
                            If .Value = System.Windows.Forms.CheckState.Checked Then
                                intDefalut = 1
                            Else
                                intDefalut = 0
                            End If
                            .Row = intLoopCounter
                            .Col = enmshipdetail.VAL_INACTIVE
                            If .Value = System.Windows.Forms.CheckState.Checked Then
                                intInactive = 1
                            Else
                                intInactive = 0
                            End If
                            varShipcode = Nothing
                            Call .GetText(enmshipdetail.VAL_SHIPCODE, intLoopCounter, varShipcode)
                            varShipDesc = Nothing
                            Call .GetText(enmshipdetail.VAL_SHIPDESC, intLoopCounter, varShipDesc)
                            varShipAdd1 = Nothing
                            Call .GetText(enmshipdetail.VAL_SHIPADD1, intLoopCounter, varShipAdd1)
                            varShipAdd2 = Nothing
                            Call .GetText(enmshipdetail.VAL_SHIPADD2, intLoopCounter, varShipAdd2)
                            varCity = Nothing
                            Call .GetText(enmshipdetail.VAL_CITY, intLoopCounter, varCity)
                            varDist = Nothing
                            Call .GetText(enmshipdetail.VAL_DISCT, intLoopCounter, varDist)
                            varState = Nothing
                            Call .GetText(enmshipdetail.VAL_STATE, intLoopCounter, varState)
                            varCountry = Nothing
                            Call .GetText(enmshipdetail.VAL_COUNTRY, intLoopCounter, varCountry)
                            varPinCode = Nothing
                            Call .GetText(enmshipdetail.VAL_Pin, intLoopCounter, varPinCode)
                            varPhone = Nothing
                            Call .GetText(enmshipdetail.VAL_PHONE, intLoopCounter, varPhone)
                            varFax = Nothing
                            Call .GetText(enmshipdetail.VAL_FAX, intLoopCounter, varFax)
                            varEmailId = Nothing
                            Call .GetText(enmshipdetail.VAL_EMAILID, intLoopCounter, varEmailId)
                            varContact = Nothing
                            Call .GetText(enmshipdetail.VAL_CONTACTPERSON, intLoopCounter, varContact)
                            varDesignation = Nothing
                            Call .GetText(enmshipdetail.VAL_DESIGNATION, intLoopCounter, varDesignation)
                            varInactiveDate = Nothing
                            Call .GetText(enmshipdetail.VAL_INACTIVEDATE, intLoopCounter, varInactiveDate)
                            VAL_SHIP_GSTIN_ID = Nothing
                            Call .GetText(enmshipdetail.VAL_SHIP_GSTIN_ID, intLoopCounter, VAL_SHIP_GSTIN_ID)
                            
                            varGSTStatecode = Nothing
                            Call .GetText(enmshipdetail.VAL_GSTSTATECODE, intLoopCounter, varGSTStatecode)
                            varGSTStatename = Nothing
                            Call .GetText(enmshipdetail.VAL_GSTSTATEDESC, intLoopCounter, varGSTStatename)

                            VAR_DISTANCE_FROM_UNIT = Nothing
                            Call .GetText(enmshipdetail.VAL_DISTANCE_FROM_UNIT, intLoopCounter, VAR_DISTANCE_FROM_UNIT)
                            If Len(Trim(varShipcode)) > 0 And Len(Trim(varShipDesc)) > 0 Then

                                strQuery = "select Count(shipping_code)CntShippingDetails from Customer_Shipping_Dtl Where Unit_Code='" & gstrUNITID & "' And customer_code='" & pstrCustomerCode & "' and shipping_code='" & Trim(varShipcode) & "'"
                                CntShippingDetails = SqlConnectionclass.ExecuteScalar(strQuery)

                                If CntShippingDetails > 0 Then
                                    mstrInsertUpdate = Trim(mstrInsertUpdate) & "Update Customer_Shipping_Dtl set Inactive_Flag=" & intInactive & ",Default_Address=" & intDefalut
                                    mstrInsertUpdate = mstrInsertUpdate & ",Inactive_Date='" & getDateForDB(varInactiveDate) & "',Upd_Dt=getdate(),Upd_Userid='" & mP_User & "'"
                                    mstrInsertUpdate = mstrInsertUpdate & " Where Unit_Code='" & gstrUNITID & "' And  shipping_code='" & varShipcode & "' and customer_code='" & pstrCustomerCode & "' " & vbCrLf
                                Else

                                    mstrInsertUpdate = Trim(mstrInsertUpdate) & "Insert into Customer_Shipping_Dtl(Customer_Code,Shipping_Code,Shipping_Desc,Default_Address,InActive_Flag,InActive_date,Ship_Address1,Ship_Address2,"
                                    mstrInsertUpdate = mstrInsertUpdate & "Ship_City,Ship_dist,Ship_State,Ship_Country,Ship_Pin,Ship_Phone,Ship_Fax,Ship_email_id,Ship_Contact_person,Ship_Person_desig,Ent_Dt,Ent_Userid,Upd_Dt,Upd_Userid,Unit_Code,GST_STATE_CODE,GST_STATE_DESC,GSTIN_ID,DISTANCE_FROM_UNIT_KM)"
                                    mstrInsertUpdate = mstrInsertUpdate & "values('" & Trim(pstrCustomerCode) & "','" & Trim(varShipcode) & "','" & Trim(varShipDesc) & "'," & intDefalut & "," & intInactive & ",'" & getDateForDB(varInactiveDate) & "','" & Trim(varShipAdd1) & "'"
                                    mstrInsertUpdate = mstrInsertUpdate & ",'" & Trim(varShipAdd2) & "','" & Trim(varCity) & "','" & Trim(varDist) & "','" & Trim(varState) & "','" & Trim(varCountry) & "','" & Trim(varPinCode) & "','" & Trim(varPhone) & "'"
                                    mstrInsertUpdate = mstrInsertUpdate & ",'" & Trim(varFax) & "','" & Trim(varEmailId) & "','" & Trim(varContact) & "','" & Trim(varDesignation) & "',getdate(),'" & mP_User & "',getdate(),'" & mP_User & "','" & gstrUNITID & "','" & varGSTStatecode & "','" & varGSTStatename & "','" & VAL_SHIP_GSTIN_ID & "'," & VAR_DISTANCE_FROM_UNIT & ")"
                                End If

                                If intDefalut = 1 Then
                                    mstrInsertUpdate = Trim(mstrInsertUpdate) & " update Customer_mst set GST_SHIPSTATE_CODE='" & varGSTStatecode & "' where Customer_Code='" & pstrCustomerCode & "' and UNIT_CODE='" & gstrUNITID & "'"
                                End If
                            End If
                        End If
                    Next
                End With
            End If

        Catch ex As Exception
            RaiseException(ex)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Sub

    Private Sub InsertUpdateShippingDetails(ByVal pstrCustomerCode As String)
        '-------------------------------------------------------------------------------------------
        ' Function      : Insert and Update the Customer Shipping Detail
        '----------------------------------------------------------------------------------------------
        Dim intLoopCounter As Short
        Dim varDelete As Object = Nothing
        Dim rsGetShippingdtl As ClsResultSetDB
        Dim strQuery As String
        Dim varShipcode As Object = Nothing
        Dim varShipDesc As Object = Nothing
        Dim varShipAdd1 As Object = Nothing
        Dim varShipAdd2 As Object = Nothing
        Dim varCity As Object = Nothing
        Dim varDist As Object = Nothing
        Dim varState As Object = Nothing
        Dim varCountry As Object = Nothing
        Dim varPinCode As Object = Nothing
        Dim varPhone As Object = Nothing
        Dim varFax As Object = Nothing
        Dim varEmailId As Object = Nothing
        Dim varContact As Object = Nothing
        Dim varDesignation As Object = Nothing
        Dim varInactiveDate As Object = Nothing
        Dim intDefalut As Short
        Dim intInactive As Short
        'GST CHANGES
        Dim varGSTStatecode As Object = Nothing
        Dim varGSTStatename As Object = Nothing

        Dim VAL_SHIP_GSTIN_ID As Object = Nothing

        'GST CHANGES
        Dim VAR_DISTANCE_FROM_UNIT As Object = Nothing
        On Error GoTo ErrHandler
        If cmdgrpCustMst.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Or cmdgrpCustMst.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
            rsGetShippingdtl = New ClsResultSetDB
            With spShippingAddess
                mstrInsertUpdate = ""
                For intLoopCounter = 1 To .MaxRows
                    varDelete = Nothing
                    Call .GetText(enmshipdetail.VAL_DELETE, intLoopCounter, varDelete)
                    If UCase(varDelete) <> "D" Then
                        .Row = intLoopCounter
                        .Col = enmshipdetail.VAL_DEFAULT
                        If .Value = System.Windows.Forms.CheckState.Checked Then
                            intDefalut = 1
                        Else
                            intDefalut = 0
                        End If
                        .Row = intLoopCounter
                        .Col = enmshipdetail.VAL_INACTIVE
                        If .Value = System.Windows.Forms.CheckState.Checked Then
                            intInactive = 1
                        Else
                            intInactive = 0
                        End If
                        varShipcode = Nothing
                        Call .GetText(enmshipdetail.VAL_SHIPCODE, intLoopCounter, varShipcode)
                        varShipDesc = Nothing
                        Call .GetText(enmshipdetail.VAL_SHIPDESC, intLoopCounter, varShipDesc)
                        varShipAdd1 = Nothing
                        Call .GetText(enmshipdetail.VAL_SHIPADD1, intLoopCounter, varShipAdd1)
                        varShipAdd2 = Nothing
                        Call .GetText(enmshipdetail.VAL_SHIPADD2, intLoopCounter, varShipAdd2)
                        varCity = Nothing
                        Call .GetText(enmshipdetail.VAL_CITY, intLoopCounter, varCity)
                        varDist = Nothing
                        Call .GetText(enmshipdetail.VAL_DISCT, intLoopCounter, varDist)
                        varState = Nothing
                        Call .GetText(enmshipdetail.VAL_STATE, intLoopCounter, varState)
                        varCountry = Nothing
                        Call .GetText(enmshipdetail.VAL_COUNTRY, intLoopCounter, varCountry)
                        varPinCode = Nothing
                        Call .GetText(enmshipdetail.VAL_Pin, intLoopCounter, varPinCode)
                        varPhone = Nothing
                        Call .GetText(enmshipdetail.VAL_PHONE, intLoopCounter, varPhone)
                        varFax = Nothing
                        Call .GetText(enmshipdetail.VAL_FAX, intLoopCounter, varFax)
                        varEmailId = Nothing
                        Call .GetText(enmshipdetail.VAL_EMAILID, intLoopCounter, varEmailId)
                        varContact = Nothing
                        Call .GetText(enmshipdetail.VAL_CONTACTPERSON, intLoopCounter, varContact)
                        varDesignation = Nothing
                        Call .GetText(enmshipdetail.VAL_DESIGNATION, intLoopCounter, varDesignation)
                        varInactiveDate = Nothing
                        Call .GetText(enmshipdetail.VAL_INACTIVEDATE, intLoopCounter, varInactiveDate)
                        'GST CHANGES
                        'Abhijit
                        VAL_SHIP_GSTIN_ID = Nothing
                        Call .GetText(enmshipdetail.VAL_SHIP_GSTIN_ID, intLoopCounter, VAL_SHIP_GSTIN_ID)
                        'Abhijit


                        varGSTStatecode = Nothing
                        Call .GetText(enmshipdetail.VAL_GSTSTATECODE, intLoopCounter, varGSTStatecode)
                        varGSTStatename = Nothing
                        Call .GetText(enmshipdetail.VAL_GSTSTATEDESC, intLoopCounter, varGSTStatename)


                        'GST CHANGES
                        VAR_DISTANCE_FROM_UNIT = Nothing
                        Call .GetText(enmshipdetail.VAL_DISTANCE_FROM_UNIT, intLoopCounter, VAR_DISTANCE_FROM_UNIT)
                        If Len(Trim(varShipcode)) > 0 And Len(Trim(varShipDesc)) > 0 Then
                          
                            strQuery = "select customer_code,shipping_code from Customer_Shipping_Dtl Where Unit_Code='" & gstrUNITID & "' And customer_code='" & pstrCustomerCode & "' and shipping_code='" & Trim(varShipcode) & "'"
                            rsGetShippingdtl.GetResult(strQuery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
                            If rsGetShippingdtl.GetNoRows > 0 Then
                                mstrInsertUpdate = Trim(mstrInsertUpdate) & "Update Customer_Shipping_Dtl set Inactive_Flag=" & intInactive & ",Default_Address=" & intDefalut
                                mstrInsertUpdate = mstrInsertUpdate & ",Inactive_Date='" & getDateForDB(varInactiveDate) & "',Upd_Dt=getdate(),Upd_Userid='" & mP_User & "'"
                                mstrInsertUpdate = mstrInsertUpdate & " Where Unit_Code='" & gstrUNITID & "' And  shipping_code='" & varShipcode & "' and customer_code='" & pstrCustomerCode & "'" & vbCrLf
                            Else

                                mstrInsertUpdate = Trim(mstrInsertUpdate) & "Insert into Customer_Shipping_Dtl(Customer_Code,Shipping_Code,Shipping_Desc,Default_Address,InActive_Flag,InActive_date,Ship_Address1,Ship_Address2,"
                                'GST CHANGES VALUE ADDED
                                mstrInsertUpdate = mstrInsertUpdate & "Ship_City,Ship_dist,Ship_State,Ship_Country,Ship_Pin,Ship_Phone,Ship_Fax,Ship_email_id,Ship_Contact_person,Ship_Person_desig,Ent_Dt,Ent_Userid,Upd_Dt,Upd_Userid,Unit_Code,GST_STATE_CODE,GST_STATE_DESC,GSTIN_ID,DISTANCE_FROM_UNIT_KM)"
                                'GST CHANGES VALUE ADDED
                                mstrInsertUpdate = mstrInsertUpdate & "values('" & Trim(pstrCustomerCode) & "','" & Trim(varShipcode) & "','" & Trim(varShipDesc) & "'," & intDefalut & "," & intInactive & ",'" & getDateForDB(varInactiveDate) & "','" & Trim(varShipAdd1) & "'"
                                mstrInsertUpdate = mstrInsertUpdate & ",'" & Trim(varShipAdd2) & "','" & Trim(varCity) & "','" & Trim(varDist) & "','" & Trim(varState) & "','" & Trim(varCountry) & "','" & Trim(varPinCode) & "','" & Trim(varPhone) & "'"
                                'GST CHANGES VALUE ADDED
                                'mstrInsertUpdate = mstrInsertUpdate & ",'" & Trim(varFax) & "','" & Trim(varEmailId) & "','" & Trim(varContact) & "','" & Trim(varDesignation) & "',getdate(),'" & mP_User & "',getdate(),'" & mP_User & "','" & gstrUNITID & "')"
                                mstrInsertUpdate = mstrInsertUpdate & ",'" & Trim(varFax) & "','" & Trim(varEmailId) & "','" & Trim(varContact) & "','" & Trim(varDesignation) & "',getdate(),'" & mP_User & "',getdate(),'" & mP_User & "','" & gstrUNITID & "','" & varGSTStatecode & "','" & varGSTStatename & "','" & VAL_SHIP_GSTIN_ID & "'," & VAR_DISTANCE_FROM_UNIT & ")"
                                'GST CHANGES
                                
                            End If
                            If intDefalut = 1 Then
                                mstrInsertUpdate = Trim(mstrInsertUpdate) & "update Customer_mst set GST_SHIPSTATE_CODE='" & varGSTStatecode & "' where Customer_Code='" & pstrCustomerCode & "' and UNIT_CODE='" & gstrUNITID & "'"
                            End If
                        End If
                    End If
                Next
                rsGetShippingdtl.ResultSetClose()
                rsGetShippingdtl = Nothing
            End With
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub ShowShippingDetail(ByVal pstrCustomerCode As String)
        '-------------------------------------------------------------------------------------------
        ' Function      : Show the Customer Shipping Detail in View and Edit Mode
        '---------------------------------------------------------------------------------------------
        Dim rsgetDetail As ClsResultSetDB
        Dim strQuery As String
        Dim intMaxRecord As Short
        Dim intLoopCounter As Short
        On Error GoTo ErrHandler
        If cmdgrpCustMst.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            With spShippingAddess
                rsgetDetail = New ClsResultSetDB
                strQuery = "select Customer_Code,Shipping_Code,Shipping_Desc,Default_Address,InActive_Flag,InActive_date,Ship_Address1,Ship_Address2,Ship_City,Ship_dist,Ship_State,Ship_Country,Ship_Pin,Ship_Phone,Ship_Fax,Ship_email_id,Ship_Contact_person,Ship_Person_desig,GST_STATE_CODE,GST_STATE_DESC,GSTIN_ID,ISNULL(DISTANCE_FROM_UNIT_KM,0) DISTANCE_FROM_UNIT_KM from Customer_Shipping_Dtl Where Unit_Code='" & gstrUnitId & "' And customer_code='" & Trim(pstrCustomerCode) & "' order by Shipping_code,default_Address,inactive_flag "
                rsgetDetail.GetResult(strQuery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
                If rsgetDetail.GetNoRows > 0 Then
                    .MaxRows = 0
                    intMaxRecord = rsgetDetail.GetNoRows
                    mintTotalRecord = intMaxRecord
                    rsgetDetail.MoveFirst()
                    For intLoopCounter = 1 To intMaxRecord
                        Call AddBlankRowinGrid()
                        Call .SetText(enmshipdetail.VAL_DEFAULT, intLoopCounter, IIf(rsgetDetail.GetValue("Default_Address") = "True", 1, 0))
                        Call .SetText(enmshipdetail.VAL_INACTIVE, intLoopCounter, IIf(rsgetDetail.GetValue("Inactive_Flag") = "True", 1, 0))
                        If rsgetDetail.GetValue("Inactive_Flag") = True Then
                            .Row = intLoopCounter
                            .Col = enmshipdetail.VAL_INACTIVE
                            .Lock = True
                        End If
                        Call .SetText(enmshipdetail.VAL_SHIPCODE, intLoopCounter, rsgetDetail.GetValue("Shipping_Code"))
                        Call .SetText(enmshipdetail.VAL_SHIPDESC, intLoopCounter, rsgetDetail.GetValue("Shipping_Desc"))
                        Call .SetText(enmshipdetail.VAL_SHIPADD1, intLoopCounter, rsgetDetail.GetValue("Ship_Address1"))
                        Call .SetText(enmshipdetail.VAL_SHIPADD2, intLoopCounter, rsgetDetail.GetValue("Ship_Address2"))
                        Call .SetText(enmshipdetail.VAL_CITY, intLoopCounter, rsgetDetail.GetValue("Ship_City"))
                        Call .SetText(enmshipdetail.VAL_DISCT, intLoopCounter, rsgetDetail.GetValue("Ship_Dist"))
                        Call .SetText(enmshipdetail.VAL_STATE, intLoopCounter, rsgetDetail.GetValue("Ship_State"))
                        Call .SetText(enmshipdetail.VAL_COUNTRY, intLoopCounter, rsgetDetail.GetValue("Ship_country"))
                        Call .SetText(enmshipdetail.VAL_Pin, intLoopCounter, rsgetDetail.GetValue("Ship_pin"))
                        Call .SetText(enmshipdetail.VAL_PHONE, intLoopCounter, rsgetDetail.GetValue("Ship_Phone"))
                        Call .SetText(enmshipdetail.VAL_FAX, intLoopCounter, rsgetDetail.GetValue("Ship_fax"))
                        Call .SetText(enmshipdetail.VAL_EMAILID, intLoopCounter, rsgetDetail.GetValue("Ship_email_id"))
                        Call .SetText(enmshipdetail.VAL_CONTACTPERSON, intLoopCounter, rsgetDetail.GetValue("Ship_contact_person"))
                        Call .SetText(enmshipdetail.VAL_DESIGNATION, intLoopCounter, rsgetDetail.GetValue("Ship_person_desig"))

                        If (rsgetDetail.GetValue("Inactive_Flag")) = True Then
                            Call .SetText(enmshipdetail.VAL_INACTIVEDATE, intLoopCounter, VB6.Format(rsgetDetail.GetValue("Inactive_Date"), gstrDateFormat))
                        End If

                        ''abhijit
                        Call .SetText(enmshipdetail.VAL_GSTSTATECODE, intLoopCounter, rsgetDetail.GetValue("GST_STATE_CODE"))
                        Call .SetText(enmshipdetail.VAL_GSTSTATEDESC, intLoopCounter, rsgetDetail.GetValue("GST_STATE_DESC"))
                        Call .SetText(enmshipdetail.VAL_SHIP_GSTIN_ID, intLoopCounter, rsgetDetail.GetValue("GSTIN_ID"))
                        ''abhijit
                        Call .SetText(enmshipdetail.VAL_DISTANCE_FROM_UNIT, intLoopCounter, rsgetDetail.GetValue("DISTANCE_FROM_UNIT_KM"))
                        rsgetDetail.MoveNext()
                    Next
                    .Row = 0
                    .Row2 = .MaxRows
                    .Col = enmshipdetail.VAL_DEFAULT
                    .Col2 = enmshipdetail.VAL_INACTIVEDATE
                    .BlockMode = True
                    .Lock = True
                    .BlockMode = False
                End If
                rsgetDetail.ResultSetClose()
                rsgetDetail = Nothing
            End With
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub LoopAllControlsInsideForm_MKTMST0001(ByVal Ctrls As Control.ControlCollection)
        For Each txtControl As Control In Ctrls
            If TypeOf txtControl Is System.Windows.Forms.TextBox Then
                DirectCast(txtControl, TextBox).Text = String.Empty
                DirectCast(txtControl, TextBox).Enabled = True
                DirectCast(txtControl, TextBox).BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            ElseIf TypeOf txtControl Is CtlGeneral Then
                DirectCast(txtControl, CtlGeneral).Text = String.Empty
                DirectCast(txtControl, CtlGeneral).Enabled = True
                DirectCast(txtControl, CtlGeneral).BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            ElseIf TypeOf txtControl Is UCActXCtl.UCctlNumberBox Then
                DirectCast(txtControl, UCActXCtl.UCctlNumberBox).Text = String.Empty
                DirectCast(txtControl, UCActXCtl.UCctlNumberBox).Enabled = True
                DirectCast(txtControl, UCActXCtl.UCctlNumberBox).BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            Else
                If TypeOf txtControl Is System.Windows.Forms.RadioButton Then
                    DirectCast(txtControl, RadioButton).Enabled = True
                Else
                    If TypeOf txtControl Is System.Windows.Forms.CheckBox Then
                        DirectCast(txtControl, CheckBox).Checked = False
                        DirectCast(txtControl, CheckBox).Enabled = True
                    End If
                End If
            End If
            LoopAllControlsInsideForm_MKTMST0001(txtControl.Controls)
        Next
    End Sub
    Private Sub DisableAllControlsInsideForm_MKTMST0001(ByVal Ctrls As Control.ControlCollection)
        For Each txtControl As Control In Ctrls
            If TypeOf txtControl Is System.Windows.Forms.TextBox Then
                DirectCast(txtControl, TextBox).Enabled = False
                DirectCast(txtControl, TextBox).BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            ElseIf TypeOf txtControl Is System.Windows.Forms.ComboBox Then
                DirectCast(txtControl, ComboBox).Enabled = False
                DirectCast(txtControl, ComboBox).BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            ElseIf TypeOf txtControl Is CtlGeneral Then
                DirectCast(txtControl, CtlGeneral).Enabled = False
                DirectCast(txtControl, CtlGeneral).BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            ElseIf TypeOf txtControl Is System.Windows.Forms.Button And txtControl.Parent.ProductName <> "UCActXCtl" Then
                DirectCast(txtControl, Button).Enabled = False
                DirectCast(txtControl, Button).BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            ElseIf TypeOf txtControl Is UCActXCtl.UCctlNumberBox Then
                DirectCast(txtControl, UCActXCtl.UCctlNumberBox).Enabled = False
                DirectCast(txtControl, UCActXCtl.UCctlNumberBox).BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            Else
                If TypeOf txtControl Is System.Windows.Forms.CheckBox Then
                    DirectCast(txtControl, CheckBox).Enabled = False
                End If
                If TypeOf txtControl Is System.Windows.Forms.RadioButton Then
                    DirectCast(txtControl, RadioButton).Enabled = False
                End If
            End If
            DisableAllControlsInsideForm_MKTMST0001(txtControl.Controls)
        Next
    End Sub
    Private Sub EnableAllControlsInsideForm_MKTMST0001(ByVal Ctrls As Control.ControlCollection)
        For Each txtControl As Control In Ctrls
            If TypeOf txtControl Is System.Windows.Forms.TextBox Then
                DirectCast(txtControl, TextBox).Enabled = True
                DirectCast(txtControl, TextBox).BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            ElseIf TypeOf txtControl Is CtlGeneral Then
                DirectCast(txtControl, CtlGeneral).Enabled = True
                DirectCast(txtControl, CtlGeneral).BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            ElseIf TypeOf txtControl Is UCActXCtl.UCctlNumberBox Then
                DirectCast(txtControl, UCActXCtl.UCctlNumberBox).Enabled = True
                DirectCast(txtControl, UCActXCtl.UCctlNumberBox).BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            ElseIf TypeOf txtControl Is System.Windows.Forms.RadioButton Then
                DirectCast(txtControl, RadioButton).Enabled = True
            Else
                If TypeOf txtControl Is System.Windows.Forms.CheckBox Then
                    DirectCast(txtControl, CheckBox).Enabled = True
                End If
            End If
            EnableAllControlsInsideForm_MKTMST0001(txtControl.Controls)
        Next
    End Sub

    Private Sub spShippingAddess_KeyDownEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyDownEvent) Handles spShippingAddess.KeyDownEvent
        Dim KeyCode As Short = e.keyCode
        Dim strshippingCode As String = Nothing
        Dim strQuery As String
        Dim CntShippingNOS As Integer = Nothing
        Dim varShipcode As Object = Nothing
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F1 And spShippingAddess.ActiveCol = enmshipdetail.VAL_GSTSTATECODE Then Call GSTBillstateDesc()
        If KeyCode = System.Windows.Forms.Keys.N AndAlso tabCustomer.SelectedIndex = 2 AndAlso Me.cmdgrpCustMst.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then

            strshippingCode = GenerateShippingCode(Trim(txtCustCode.Text))

            Call spShippingAddess.GetText(enmshipdetail.VAL_SHIPCODE, spShippingAddess.MaxRows, varShipcode)
            If spShippingAddess.MaxRows > 0 Then
                If CInt((Microsoft.VisualBasic.Right(varShipcode, 2))) + 1 = CInt((Microsoft.VisualBasic.Right(strshippingCode, 2))) Then
                    Call AddBlankRowinGrid()
                    Call AddShippingDetailinNewRow(Trim(txtCustCode.Text))
                    Call spShippingAddess.SetText(enmshipdetail.VAL_SHIPCODE, spShippingAddess.MaxRows, strshippingCode)
                End If

            End If

            
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub spShippingAddess_KeyPressEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyPressEvent) Handles spShippingAddess.KeyPressEvent
        Dim strshippingCode As String = Nothing
        Dim varDelete As Object = Nothing
        Dim Row As Integer
        On Error GoTo ErrHandler
        If cmdgrpCustMst.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            With spShippingAddess
                Select Case e.keyAscii
                    Case 39, 34, 96, 45
                        e.keyAscii = 0
                    Case 1 To 7, 14 To 43, 58 To 200
                        If e.keyAscii = 4 And ChkActive.Checked = False Then
                            .Row = .MaxRows - 1
                        End If
                        If .ActiveCol = enmshipdetail.VAL_Pin Or .ActiveCol = enmshipdetail.VAL_PHONE Or .ActiveCol = enmshipdetail.VAL_FAX Then
                            e.keyAscii = 0
                        End If

                    Case System.Windows.Forms.Keys.Return
                        Call .GetText(enmshipdetail.VAL_DELETE, .MaxRows, varDelete)
                        Row = IIf(UCase(varDelete) = "D", .MaxRows - 1, .MaxRows)
                        If .ActiveCol = enmshipdetail.VAL_DISTANCE_FROM_UNIT Then
                            If ValidateRowData(.ActiveRow, .ActiveCol) = True Then
                                strshippingCode = GetNextShippingCode()           'Get the Next Shipping Code
                                Call AddBlankRowinGrid()                          'Add new Row
                                .SetText(enmshipdetail.VAL_SHIPCODE, spShippingAddess.MaxRows, strshippingCode)
                            End If
                        End If
                End Select
            End With
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub spShippingAddess_KeyUpEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyUpEvent) Handles spShippingAddess.KeyUpEvent
        '-------------------------------------------------------------------------------------------
        ' Function      : Delete Row on CTRL+D
        '--------------------------------------------------------------------------------------------
        Dim intRow As Integer
        Dim intDelete As Integer
        Dim intLoopCount As Integer
        Dim intMaxLoop As Integer
        Dim varDelete As Object = Nothing
        Dim VarDefault As Object = Nothing
        Dim varInactive As Object = Nothing
        On Error GoTo ErrHandler
        If ((e.shift = 2) And (e.keyCode = System.Windows.Forms.Keys.D)) Then
            If cmdgrpCustMst.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                With spShippingAddess
                    If .ActiveRow > mintTotalRecord Then
                        Call .GetText(enmshipdetail.VAL_DEFAULT, .ActiveRow, VarDefault)
                        Call .GetText(enmshipdetail.VAL_INACTIVE, .ActiveRow, varInactive)
                        If Val(VarDefault) = 1 Then
                            MsgBox("Default Address can not be deleted.", vbInformation, ResolveResString(100))
                            Exit Sub
                        ElseIf Val(varInactive) = 1 Then
                            MsgBox("Inactive Address can not be deleted.", vbInformation, ResolveResString(100))
                            Exit Sub
                        Else
                            If MsgBox("Sure to delete the Row -" & .ActiveRow, vbInformation + vbYesNo, ResolveResString(100)) = vbYes Then
                                intRow = .ActiveRow : intMaxLoop = spShippingAddess.MaxRows
                                For intLoopCount = 1 To intMaxLoop
                                    If intLoopCount <> intRow Then
                                        varDelete = Nothing
                                        Call .GetText(enmshipdetail.VAL_DELETE, intLoopCount, varDelete)
                                        If UCase(varDelete) = "D" Then
                                            intDelete = intDelete + 1
                                        End If
                                    End If
                                Next
                                If (intMaxLoop - intDelete) > 1 Then
                                    Call .SetText(enmshipdetail.VAL_DELETE, intRow, "D")
                                    .Row = .ActiveRow : .Row2 = .ActiveRow : .BlockMode = True : .RowHidden = True : .BlockMode = False
                                End If
                            End If
                        End If
                    Else
                        MessageBox.Show("Existing record cannot be deleted.", ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Exit Sub
                    End If
                End With
            End If
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub spShippingAddess_LeaveCell(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_LeaveCellEvent) Handles spShippingAddess.LeaveCell
        '-------------------------------------------------------------------------------------------
        ' Function      : Validate the cell of the Grid
        '--------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        If cmdgrpCustMst.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            With spShippingAddess
                If e.newCol = -1 Or e.newRow = -1 Then
                    Exit Sub
                End If
                Call ValidateRowData(.ActiveRow, .ActiveCol)
            End With
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub spShippingAddess_ButtonClicked(ByVal sender As System.Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ButtonClickedEvent) Handles spShippingAddess.ButtonClicked
        '-------------------------------------------------------------------------------------------
        ' Function      : Select Deafult Address
        '--------------------------------------------------------------------------------------------
        Dim intSubItem As Integer
        Dim intDefault As Integer
        Dim intFound As Integer
        Dim intctr As Integer
        Dim varDefault As Object = Nothing
        On Error GoTo ErrHandler
        If cmdgrpCustMst.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            With spShippingAddess
                If .ActiveCol = enmshipdetail.VAL_DEFAULT Then
                    intFound = 0
                    For intctr = 1 To spShippingAddess.MaxRows
                        varDefault = Nothing
                        Call spShippingAddess.GetText(enmshipdetail.VAL_DEFAULT, intctr, varDefault)
                        If Val(varDefault) = 1 Then intFound = intFound + 1
                        If intFound > 1 Then Call spShippingAddess.SetText(enmshipdetail.VAL_DEFAULT, e.row, False)
                    Next intctr
                    .Row = e.row : .Col = enmshipdetail.VAL_DEFAULT
                    If .Value = CStr(System.Windows.Forms.CheckState.Checked) Then
                        .Row = e.row
                        .Col = enmshipdetail.VAL_INACTIVE
                        .Lock = True
                        Exit Sub
                    Else
                        .Row = e.row
                        .Col = enmshipdetail.VAL_INACTIVE
                        .Lock = False
                        Exit Sub
                    End If
                ElseIf .ActiveCol = enmshipdetail.VAL_INACTIVE Then
                    .Row = e.row : .Col = enmshipdetail.VAL_INACTIVE
                    If .Value = CStr(System.Windows.Forms.CheckState.Checked) Then
                        Call .SetText(enmshipdetail.VAL_INACTIVEDATE, e.row, VB6.Format(GetServerDate(), gstrDateFormat))
                        .Row = e.row
                        .Col = enmshipdetail.VAL_DEFAULT
                        .Lock = True
                        Exit Sub
                    Else
                        Call .SetText(enmshipdetail.VAL_INACTIVEDATE, e.row, "")
                        .Row = e.row
                        .Col = enmshipdetail.VAL_DEFAULT
                        .Lock = False
                        Exit Sub
                    End If
                End If
            End With
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    
    Private Function ValidateRowData(ByVal pintRow As Long, ByVal pintCol As Long) As Boolean
        Dim varShippingDesc As Object = Nothing
        Dim varShipAdd1 As Object = Nothing
        Dim varShipAdd2 As Object = Nothing
        Dim varShipcity As Object = Nothing
        Dim varShipState As Object = Nothing
        Dim varShipCountry As Object = Nothing
        Dim varShipPinCode As Object = Nothing
        Dim varEmailId As Object = Nothing
        Dim VarDelete As Object = Nothing
        Dim strValidateField As String
        Dim intFirstPos As Integer
        Dim intSecPos As Integer
        Dim strEmail As String
        Dim Row As Long
        Dim Col As Long
        Dim varInactive As Object = Nothing
        'GST CHANGES
        Dim varGSTCode As Object = Nothing
        Dim varGSTNAME As Object = Nothing
        Dim varGSTINID As Object = Nothing
        'GST CHANGES
        Dim varDistinctFromUnit As Object = Nothing
        strValidateField = "Followings are invalid or have not been entered:" & vbCrLf
        Dim blnsetColumn As Boolean
        blnsetColumn = True
        ValidateRowData = True
        On Error GoTo ErrHandler
        With spShippingAddess
            Row = pintRow
            Col = pintCol
            Call .GetText(enmshipdetail.VAL_DELETE, Row, VarDelete)
            If UCase(VarDelete) = "D" Then Exit Function
            Call .GetText(enmshipdetail.VAL_INACTIVE, Row, varInactive)
            If Val(varInactive) = 1 Then Exit Function
            Call .GetText(enmshipdetail.VAL_SHIPDESC, Row, varShippingDesc)
            Call .GetText(enmshipdetail.VAL_SHIPADD1, Row, varShipAdd1)
            Call .GetText(enmshipdetail.VAL_SHIPADD2, Row, varShipAdd2)
            Call .GetText(enmshipdetail.VAL_CITY, Row, varShipcity)
            Call .GetText(enmshipdetail.VAL_STATE, Row, varShipState)
            Call .GetText(enmshipdetail.VAL_COUNTRY, Row, varShipCountry)
            Call .GetText(enmshipdetail.VAL_Pin, Row, varShipPinCode)
            Call .GetText(enmshipdetail.VAL_EMAILID, Row, varEmailId)
            'GST CHANGES
            Call .GetText(enmshipdetail.VAL_GSTSTATECODE, Row, varGSTCode)
            Call .GetText(enmshipdetail.VAL_GSTSTATEDESC, Row, varGSTNAME)
            Call .GetText(enmshipdetail.VAL_SHIP_GSTIN_ID, Row, varGSTINID)
            If Trim(varGSTCode) <> "" Then
                If Not DataExist("SELECT TOP 1 1 FROM VW_GST_STATE_MST WHERE STATE_CODE ='" & varGSTCode & "'") Then
                    strValidateField = strValidateField & vbCrLf & ". Invalid GST State Code "
                    If blnsetColumn = True Then .Col = enmshipdetail.VAL_GSTSTATECODE : blnsetColumn = False
                End If

                If Not DataExist("SELECT TOP 1 1 FROM VW_GST_STATE_MST WHERE STATE_NAME ='" & varGSTNAME & "'") Then
                    strValidateField = strValidateField & vbCrLf & ". Invalid GST State Name"
                    If blnsetColumn = True Then .Col = enmshipdetail.VAL_GSTSTATEDESC : blnsetColumn = False
                End If
            End If

            'GST CHANGES
            Call .GetText(enmshipdetail.VAL_DISTANCE_FROM_UNIT, Row, varDistinctFromUnit)
            If Len(Trim(varShippingDesc)) = 0 And enmshipdetail.VAL_SHIPDESC <= Col Then
                strValidateField = strValidateField & vbCrLf & ". Shipping Description"
                If blnsetColumn = True Then .Col = enmshipdetail.VAL_SHIPDESC : blnsetColumn = False
            End If
            If Len(Trim(varShipAdd1)) = 0 And enmshipdetail.VAL_SHIPADD1 <= Col Then
                strValidateField = strValidateField & vbCrLf & ". Shipping Address 1"
                If blnsetColumn = True Then .Col = enmshipdetail.VAL_SHIPADD1 : blnsetColumn = False
            End If
            If Len(Trim(varShipAdd2)) = 0 And enmshipdetail.VAL_SHIPADD2 <= Col Then
                strValidateField = strValidateField & vbCrLf & ". Shipping Address 2"
                If blnsetColumn = True Then .Col = enmshipdetail.VAL_SHIPADD2 : blnsetColumn = False
            End If
            If Len(Trim(varShipcity)) = 0 And enmshipdetail.VAL_CITY <= Col Then
                strValidateField = strValidateField & vbCrLf & ". Shipping City"
                If blnsetColumn = True Then .Col = enmshipdetail.VAL_CITY : blnsetColumn = False
            End If
            If Len(Trim(varShipState)) = 0 And enmshipdetail.VAL_STATE <= Col Then
                strValidateField = strValidateField & vbCrLf & ". Shipping State"
                If blnsetColumn = True Then .Col = enmshipdetail.VAL_STATE : blnsetColumn = False
            End If
            If Len(Trim(varShipCountry)) = 0 And enmshipdetail.VAL_COUNTRY <= Col Then
                strValidateField = strValidateField & vbCrLf & ". Shipping Country"
                If blnsetColumn = True Then .Col = enmshipdetail.VAL_COUNTRY : blnsetColumn = False
            End If
            If Len(Trim(varShipPinCode)) = 0 And enmshipdetail.VAL_Pin <= Col Then
                strValidateField = strValidateField & vbCrLf & ". Shipping PinCode"
                If blnsetColumn = True Then .Col = enmshipdetail.VAL_Pin : blnsetColumn = False
            End If
            If Len(Trim(varGSTCode)) <> 99 And Len(Trim(varShipPinCode)) > 0 Then '25112019  validate PIN code when user copy paste  
                'If UCase$(varShipPinCode) Like "[0-9][0-9][0-9][0-9][0-9][0-9]" And Val(varShipPinCode) <> 0 Then    
                'Dim st1 As String = InStr(1, "1234567890.", "1234567890.")
                'Dim st2 As String = InStr(1, "1234567890.", "1234")
                'Dim st3 As String = InStr(1, "1234567890.", "1234ssdsd")
                'Dim st4 As String = InStr(1, "1234567890.", "fdfdf")
                'Dim st5 As Integer = InStr(1, "1234567890.", "0200")
                'If InStr(1, "1234567890.", varShipPinCode) <> 0 Then
                If IsNumeric(varShipPinCode) = False Then
                    strValidateField = strValidateField & vbCrLf & ". Shipping PinCode"
                    If blnsetColumn = True Then .Col = enmshipdetail.VAL_Pin : blnsetColumn = False
                End If
            End If

            If Len(Trim(varEmailId)) > 0 And enmshipdetail.VAL_EMAILID <= Col Then
                strEmail = varEmailId
                intFirstPos = InStr(1, strEmail, "@")
                intSecPos = InStr(1, strEmail, ".")
                If (intFirstPos = 0 Or intSecPos = 0) Or (intFirstPos = 1) Or (intSecPos - intFirstPos = 1) Or (Not Len(strEmail) > intSecPos) Then
                    strValidateField = strValidateField & vbCrLf & ". E-Mail ID"
                    If blnsetColumn = True Then .Col = enmshipdetail.VAL_EMAILID : blnsetColumn = False
                End If
            End If
            If gblnGSTUnit Then
                If Val(varDistinctFromUnit) = 0 And enmshipdetail.VAL_DISTANCE_FROM_UNIT <= Col Then
                    strValidateField = strValidateField & vbCrLf & ". Distance From Unit(Km)"
                    If blnsetColumn = True Then .Col = enmshipdetail.VAL_DISTANCE_FROM_UNIT : blnsetColumn = False
                End If
            End If
            If blnsetColumn = False Then
                MsgBox(strValidateField, vbInformation, ResolveResString(100))
                .Row = pintRow
                .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                .Focus()
                ValidateRowData = False
                Exit Function
            End If
        End With
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub txtBankAcct1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtBankAcct1.KeyPress, txtBankAcct2.KeyPress, TxtBankAddress.KeyPress, TxtBankname.KeyPress, txtbillCity.KeyPress, txtBillCountry.KeyPress, txtCustLoc.KeyPress, txtCustName.KeyPress, txtCustvendCode.KeyPress, txtTinNo.KeyPress, txtOffCity1.KeyPress, txtOffCountry1.KeyPress, txtOffState1.KeyPress, txtBillState.KeyPress, txtExciseRange.KeyPress, txtComRate.KeyPress, txtDivision.KeyPress, txtLst.KeyPress, txtCst.KeyPress, txtEcc.KeyPress, txtWebSite.KeyPress
        On Error GoTo ErrHandler
        Dim keyascii As Short
        keyascii = Asc(e.KeyChar)
        Select Case keyascii
            Case 39, 34, 96, 45
                e.Handled = True
        End Select
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub txtBankAcct2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtBankAcct2.KeyPress
        On Error GoTo ErrHandler
        Dim keyascii As Short
        keyascii = Asc(e.KeyChar)
        Select Case keyascii
            Case 39, 34, 96, 45
                e.Handled = True
        End Select
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub TxtBankAddress_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtBankAddress.KeyPress
        On Error GoTo ErrHandler
        Dim keyascii As Short
        keyascii = Asc(e.KeyChar)
        Select Case keyascii
            Case 39, 34, 96, 45
                e.Handled = True
        End Select
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub TxtBankname_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TxtBankname.KeyPress
        On Error GoTo ErrHandler
        Dim keyascii As Short
        keyascii = Asc(e.KeyChar)
        Select Case keyascii
            Case 39, 34, 96, 45
                e.Handled = True
        End Select
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    'Samiksha : Change for Customer Master Authorization
    Private Sub cmdHelpGlobalCustomer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _cmdHelpGlobalCustomer_0.Click
        On Error GoTo ErrHandler
        Dim strSql As String
        Dim strHelp() As String
        If authorize_active = True Then
            strSql = " select Customer_Code ,Customer_Name,CUSTOMER_LOCATION ,CUSTOMER_WEBSITE  from [GLOBAL_CUSTOMER_MST] " &
          " inner join [GLOBAL_MASTER_MAPPING]" &
       " on [GLOBAL_MASTER_MAPPING].Global_slno=[GLOBAL_CUSTOMER_MST].SLNO " &
       " and ltrim(rtrim(TableName))=ltrim(rtrim('GLOBAL_CUSTOMER_MST'))" &
       " and ISACTIVE =1" &
       " WHERE UNIT_CODE = '" & gstrUNITID & "'" &
       " and [GLOBAL_CUSTOMER_MST].CUSTOMER_CODE not in(SELECT GLOBAL_CUATOMER_CODE from Customer_mst where Unit_Code ='" & gstrUNITID & "' And GLOBAL_CUATOMER_CODE is Not Null" &
       " UNION Select GLOBAL_CUATOMER_CODE from Customer_mst_authorization where Unit_Code ='" & gstrUNITID & "' And GLOBAL_CUATOMER_CODE is Not Null) "


        ElseIf authorize_active = False Then
            strSql = " select Customer_Code ,Customer_Name,CUSTOMER_LOCATION ,CUSTOMER_WEBSITE  from [GLOBAL_CUSTOMER_MST] " &
              " inner join [GLOBAL_MASTER_MAPPING]" &
           " on [GLOBAL_MASTER_MAPPING].Global_slno=[GLOBAL_CUSTOMER_MST].SLNO " &
           " and ltrim(rtrim(TableName))=ltrim(rtrim('GLOBAL_CUSTOMER_MST'))" &
           " and ISACTIVE =1" &
           " WHERE UNIT_CODE = '" & gstrUNITID & "'" &
           " and [GLOBAL_CUSTOMER_MST].CUSTOMER_CODE not in(SELECT GLOBAL_CUATOMER_CODE from Customer_mst where Unit_Code ='" & gstrUNITID & "' And GLOBAL_CUATOMER_CODE is Not Null)"
        End If

        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        strHelp = ctlCurrCode.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSql, "List of Global Customer(s)", 1)
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        If UBound(strHelp) = -1 Then Exit Sub
        If strHelp(0) = "0" Then
            MsgBox("No Global Customers are defined.", MsgBoxStyle.Information, ResolveResString(100))
            ' txtRecLoc.Focus()
        Else
            lblGlobalCustCodeDesc.Text = strHelp(0)
            lblGlobalCustCodeDesc.Tag = strHelp(1)
            ' lblGlobalCustNameDesc.Text = strHelp(1)
            If cmdgrpCustMst.Enabled(3) = True Then

            End If
            If cmdgrpCustMst.Enabled(1) = True Then
                txtCustCode.Text = strHelp(0)
                txtCustName.Text = strHelp(1)
                txtWebSite.Text = strHelp(3)

                txtCustLoc.Text = strHelp(2)

            End If
        End If
        Select Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                txtCustCode.Enabled = False
                txtCustName.Enabled = False
                txtWebSite.Enabled = False
                txtCustLoc.Enabled = False
        End Select
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub txtbillCity_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtbillCity.KeyPress
        On Error GoTo ErrHandler
        Dim keyascii As Short
        keyascii = Asc(e.KeyChar)
        Select Case keyascii
            Case 39, 34, 96, 45
                e.Handled = True
        End Select
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub txtBillCountry_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtBillCountry.KeyPress
        On Error GoTo ErrHandler
        Dim keyascii As Short
        keyascii = Asc(e.KeyChar)
        Select Case keyascii
            Case 39, 34, 96, 45
                e.Handled = True
        End Select
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub txtCustLoc_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCustLoc.KeyPress
        On Error GoTo ErrHandler
        Dim keyascii As Short
        keyascii = Asc(e.KeyChar)
        Select Case keyascii
            Case 39, 34, 96, 45
                e.Handled = True
        End Select
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub txtCustName_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCustName.KeyPress
        On Error GoTo ErrHandler
        Dim keyascii As Short
        keyascii = Asc(e.KeyChar)
        Select Case keyascii
            Case 39, 34, 96, 45
                e.Handled = True
        End Select
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub txtCustvendCode_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCustvendCode.KeyPress
        On Error GoTo ErrHandler
        Dim keyascii As Short
        keyascii = Asc(e.KeyChar)
        Select Case keyascii
            Case 39, 34, 96, 45
                e.Handled = True
        End Select
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub txtTinNo_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtTinNo.KeyPress
        On Error GoTo ErrHandler
        Dim keyascii As Short
        keyascii = Asc(e.KeyChar)
        Select Case keyascii
            Case 39, 34, 96, 45
                e.Handled = True
        End Select
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub txtOffCity1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtOffCity1.KeyPress
        On Error GoTo ErrHandler
        Dim keyascii As Short
        keyascii = Asc(e.KeyChar)
        Select Case keyascii
            Case 39, 34, 96, 45
                e.Handled = True
        End Select
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub txtOffCountry1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtOffCountry1.KeyPress
        On Error GoTo ErrHandler
        Dim keyascii As Short
        keyascii = Asc(e.KeyChar)
        Select Case keyascii
            Case 39, 34, 96, 45
                e.Handled = True
        End Select
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub txtOffState1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtOffState1.KeyPress
        On Error GoTo ErrHandler
        Dim keyascii As Short
        keyascii = Asc(e.KeyChar)
        Select Case keyascii
            Case 39, 34, 96, 45
                e.Handled = True
        End Select
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub txtBillState_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtBillState.KeyPress
        On Error GoTo ErrHandler
        Dim keyascii As Short
        keyascii = Asc(e.KeyChar)
        Select Case keyascii
            Case 39, 34, 96, 45
                e.Handled = True
        End Select
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub txtExciseRange_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtExciseRange.KeyPress
        On Error GoTo ErrHandler
        Dim keyascii As Short
        keyascii = Asc(e.KeyChar)
        Select Case keyascii
            Case 39, 34, 96, 45
                e.Handled = True
        End Select
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub txtComRate_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtComRate.KeyPress
        On Error GoTo ErrHandler
        Dim keyascii As Short
        keyascii = Asc(e.KeyChar)
        Select Case keyascii
            Case 39, 34, 96, 45
                e.Handled = True
        End Select
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub txtDivision_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDivision.KeyPress
        On Error GoTo ErrHandler
        Dim keyascii As Short
        keyascii = Asc(e.KeyChar)
        Select Case keyascii
            Case 39, 34, 96, 45
                e.Handled = True
        End Select
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub txtLst_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtLst.KeyPress
        On Error GoTo ErrHandler
        Dim keyascii As Short
        keyascii = Asc(e.KeyChar)
        Select Case keyascii
            Case 39, 34, 96, 45
                e.Handled = True
        End Select
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub txtCst_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCst.KeyPress
        On Error GoTo ErrHandler
        Dim keyascii As Short
        keyascii = Asc(e.KeyChar)
        Select Case keyascii
            Case 39, 34, 96, 45
                e.Handled = True
        End Select
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub txtEcc_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtEcc.KeyPress
        On Error GoTo ErrHandler
        Dim keyascii As Short
        keyascii = Asc(e.KeyChar)
        Select Case keyascii
            Case 39, 34, 96, 45
                e.Handled = True
        End Select
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub txtWebSite_KeyPress_1(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtWebSite.KeyPress
        On Error GoTo ErrHandler
        Dim keyascii As Short
        keyascii = Asc(e.KeyChar)
        Select Case keyascii
            Case 39, 34, 96, 45
                e.Handled = True
        End Select
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub Chkbarcodeprinting_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Chkbarcodeprinting.CheckedChanged
        CmbPrintMethod.Enabled = Chkbarcodeprinting.Checked
        CmbPrintMethod.BackColor = IIf(Chkbarcodeprinting.Checked = True, System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED), System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED))
    End Sub

    Private Sub txtShipDur_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtShipDur.Enter

        On Error GoTo ErrHandler
        txtShipDur.SelectionStart = 0
        txtShipDur.SelectionLength = Len(txtShipDur.Text)
        Exit Sub

ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)

    End Sub

    Private Sub txtShipDur_KeyPress(ByVal sender As System.Object, ByVal EventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtShipDur.KeyPress

        Dim KeyAscii As Short = Asc(EventArgs.KeyChar) ' Only 0-9 can be entered.
        On Error GoTo ErrHandler
        If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8) And (KeyAscii <> 127) Then
            KeyAscii = 0
        End If
        GoTo EventExitSub

ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        EventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            EventArgs.Handled = True
        End If
    End Sub

    ''' <summary>
    ''' Open Help Menu for KAM Name.
    ''' </summary>
    ''' <remarks>10688280 -Add KAM Code in Customer Master.</remarks>
    Private Sub BtnHelpKAM_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnHelpKAM.Click
        Dim strSQL As String = String.Empty
        Dim strHelp As String()
        Try
            If cmdgrpCustMst.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                Return
            End If

            strSQL = " select Employee_code As [EmployeeCode], Name As [EmployeeName] from Employee_mst(NOLOCK) where UNIT_CODE ='" & gstrUNITID & "'"
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
            strHelp = ctlCurrCode.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strSQL, "List of Employee(s)", 1)
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
            If Not IsNothing(strHelp) Then
                If strHelp.Length > 0 And Not IsNothing(strHelp(1)) Then
                    txtKAMName.Text = strHelp(0)
                    lblKAMDesc.Text = strHelp(1)
                Else
                    txtKAMName.Text = String.Empty
                    lblKAMDesc.Text = String.Empty
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    ''' <summary>
    ''' Open Help for KAM on F1 press
    ''' </summary>
    ''' <remarks>10688280 -Add KAM Code in Customer Master.</remarks>
    Private Sub txtKAMName_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtKAMName.KeyDown
        Try
            If e.KeyCode = Keys.F1 Then
                BtnHelpKAM_Click(BtnHelpKAM, New EventArgs())
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    '10688280 -Add KAM Code in Customer Master.
    Private Sub txtKAMName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtKAMName.TextChanged
        Try
            If txtKAMName.Text.Trim.Length = 0 Then
                lblKAMDesc.Text = String.Empty
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub txtA4orginal_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtA4orginal.KeyPress

        Dim KeyAscii As Short = Asc(e.KeyChar)
        Select Case KeyAscii
            Case 39, 34, 96
                KeyAscii = 0

        End Select

        e.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            e.Handled = True
        End If
        AllowNumericValueInTextBox(txtA4orginal, e)
    End Sub
    Private Sub txtA4orginal_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtA4orginal.Validating

        Try
            Dim intLoopCounter As Integer
            spA4grid.MaxRows = 0
            For intLoopCounter = 1 To Val(txtA4orginal.Text)
                Call AddBlankRow()
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
        End Try

    End Sub
    Private Sub AddBlankRow()
        ''PURPOSE       : ADD A NEW BLANK ROW IN THE GRID

        Try
            With Me.spA4grid
                .MaxRows = .MaxRows + 1
                .MaxCols = 2
                .Row = .MaxRows
                .set_RowHeight(.Row, 200)
                .Col = 1 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton : .TypeButtonText = .MaxRows
                .Col = 2 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit

            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
        End Try

    End Sub
    Private Sub AddBlankRow_Reprint()
        ''PURPOSE       : ADD A NEW BLANK ROW IN THE GRID

        Try
            With Me.spA4grid_Reprint
                .MaxRows = .MaxRows + 1
                .MaxCols = 2
                .Row = .MaxRows
                .set_RowHeight(.Row, 200)
                .Col = 1 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton : .TypeButtonText = .MaxRows
                .Col = 2 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit

            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
        End Try

    End Sub
    Private Sub SetGridHeading()
        ''PURPOSE       : SET THE GRID COLUMNS HEADINGS
        Try
            With spA4grid
                .MaxRows = 0
                .MaxCols = 2
                .Row = 0
                .Col = 0 : .set_ColWidth(0, 3)
                .Col = 1 : .Text = "Nos." : .set_ColWidth(1, 400) : .FontBold = True
                .Col = 2 : .Text = "Text to Print in Invoice" : .FontBold = True : .set_ColWidth(2, 5000)
            End With

            With spA4grid_Reprint
                .MaxRows = 0
                .MaxCols = 2
                .Row = 0
                .Col = 0 : .set_ColWidth(0, 3)
                .Col = 1 : .Text = "Nos." : .set_ColWidth(1, 400) : .FontBold = True
                .Col = 2 : .Text = "Text to Print in Invoice" : .FontBold = True : .set_ColWidth(2, 5000)
            End With

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
        End Try

    End Sub

    Private Sub InsertA4customer(ByVal pstrCustomerCode As String)
        '-------------------------------------------------------------------------------------------
        ' Function      : DELETE AND INSERT A4CUSTOMER_INVOICEPRINTINGTAG  
        '----------------------------------------------------------------------------------------------
        Dim intLoopCounter As Short
        Dim rsallowa4reports As ClsResultSetDB
        Dim strQuery As String

        Try

            If IsNothing(txtA4orginal.Text) Or Val(txtA4orginal.Text) > 0 And IsNothing(txtA4Reprint.Text) Or Val(txtA4Reprint.Text) > 0 Then
                rsallowa4reports = New ClsResultSetDB
                strQuery = "select customer_code from customer_mst (nolock)  Where Unit_Code='" & gstrUNITID & "' And customer_code='" & pstrCustomerCode & "' and allowa4reports=1 "
                rsallowa4reports.GetResult(strQuery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
                If rsallowa4reports.GetNoRows >= 0 Then
                    mstrA4insert = "DELETE FROM A4CUSTOMER_INVOICEPRINTINGTAG  WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & pstrCustomerCode & "'"
                    For intLoopCounter = 1 To Val(txtA4orginal.Text)
                        With spA4grid
                            .Row = intLoopCounter
                            .Col = 2
                            If .Text = "" Or .Text = Nothing Then
                                mstrA4insert = ""
                                MsgBox("Please Enter some Text ." & vbCrLf & "Changes can't be Commit !", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100))
                                Exit Sub
                            End If
                            mstrA4insert = mstrA4insert + "INSERT INTO A4CUSTOMER_INVOICEPRINTINGTAG(UNIT_CODE,CUSTOMER_CODE,ORIGINAL_REPRINT,SERIALNO,TEXTHEADING)"
                            mstrA4insert = mstrA4insert + "SELECT '" & gstrUNITID & "','" & pstrCustomerCode & "','O'," & intLoopCounter & ",'" & .Text & "'"
                        End With
                    Next

                    For intLoopCounter = 1 To Val(txtA4Reprint.Text)
                        With spA4grid_Reprint
                            .Row = intLoopCounter
                            .Col = 2
                            If .Text = "" Or .Text = Nothing Then
                                mstrA4insert = ""
                                MsgBox("Please Enter some Text ." & vbCrLf & "Changes can't be Commit !", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100))
                                Exit Sub
                            End If
                            mstrA4insert = mstrA4insert + "INSERT INTO A4CUSTOMER_INVOICEPRINTINGTAG(UNIT_CODE,CUSTOMER_CODE,ORIGINAL_REPRINT,SERIALNO,TEXTHEADING)"
                            mstrA4insert = mstrA4insert + "SELECT '" & gstrUNITID & "','" & pstrCustomerCode & "','R'," & intLoopCounter & ",'" & .Text & "'"
                        End With
                    Next
                End If
            Else
                MsgBox("NO. of Copies Cannot be Blank or Zero ." & vbCrLf & "Changes can't be Commit !", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub

            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
        End Try
    End Sub
    Private Sub DisplayA4customer(ByVal pstrCustomerCode As String)
        '-------------------------------------------------------------------------------------------
        ' Function      : TO DISPLAY A4CUSTOMER_INVOICEPRINTINGTAG  
        '----------------------------------------------------------------------------------------------
        Dim intLoopCounter As Short
        Dim rsallowa4reports As ClsResultSetDB
        Dim strQuery As String
        Dim strOrignaltext As String
        Dim STRSQLFUNCTION As String
        Try
            rsallowa4reports = New ClsResultSetDB
            strQuery = "select * from A4CUSTOMER_INVOICEPRINTINGTAG Where Unit_Code='" & gstrUNITID & "' And customer_code='" & pstrCustomerCode & "' AND ORIGINAL_REPRINT='O' "
            rsallowa4reports.GetResult(strQuery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
            If rsallowa4reports.GetNoRows > 0 Then
                txtA4orginal.Text = CInt(Find_Value("select MAX(SERIALNO) SERIALNO from A4CUSTOMER_INVOICEPRINTINGTAG  WHERE UNIT_CODE='" + gstrUNITID + "'AND CUSTOMER_CODE='" & txtCustCode.Text & "' AND ORIGINAL_REPRINT='O'"))
                spA4grid.MaxRows = 0
                For intLoopCounter = 1 To Val(txtA4orginal.Text)
                    Call AddBlankRow()
                    With spA4grid
                        .Row = intLoopCounter
                        strOrignaltext = Find_Value("SELECT TEXTHEADING FROM A4CUSTOMER_INVOICEPRINTINGTAG  WHERE UNIT_CODE='" + gstrUNITID + "' AND CUSTOMER_CODE='" & txtCustCode.Text & "' AND ORIGINAL_REPRINT='O' AND SERIALNO=" & intLoopCounter)
                        .SetText(2, intLoopCounter, strOrignaltext)
                    End With
                Next
            Else
                txtA4orginal.Text = ""
                spA4grid.MaxRows = 0

            End If

            strQuery = "select * from A4CUSTOMER_INVOICEPRINTINGTAG Where Unit_Code='" & gstrUNITID & "' And customer_code='" & pstrCustomerCode & "' AND ORIGINAL_REPRINT='R' "
            rsallowa4reports.GetResult(strQuery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
            If rsallowa4reports.GetNoRows > 0 Then
                txtA4Reprint.Text = CInt(Find_Value("select MAX(SERIALNO) SERIALNO from A4CUSTOMER_INVOICEPRINTINGTAG  WHERE UNIT_CODE='" + gstrUNITID + "'AND CUSTOMER_CODE='" & txtCustCode.Text & "' AND ORIGINAL_REPRINT='R'"))
                spA4grid_Reprint.MaxRows = 0
                For intLoopCounter = 1 To Val(txtA4Reprint.Text)
                    Call AddBlankRow_Reprint()
                    With spA4grid_Reprint
                        .Row = intLoopCounter
                        strOrignaltext = Find_Value("SELECT TEXTHEADING FROM A4CUSTOMER_INVOICEPRINTINGTAG  WHERE UNIT_CODE='" + gstrUNITID + "' AND CUSTOMER_CODE='" & txtCustCode.Text & "' AND ORIGINAL_REPRINT='R' AND SERIALNO=" & intLoopCounter)
                        .SetText(2, intLoopCounter, strOrignaltext)
                    End With
                Next
            Else
                txtA4Reprint.Text = ""
                spA4grid_Reprint.MaxRows = 0

            End If


        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
        End Try
    End Sub

    Public Function Find_Value(ByRef strField As String) As String
        On Error GoTo ErrHandler
        Dim Rs As New ADODB.Recordset
        Rs = New ADODB.Recordset
        Rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        'Rs.Open(strField, mP_Connection, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText)
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
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

    Private Sub txtA4Reprint_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtA4Reprint.KeyPress

        Dim KeyAscii As Short = Asc(e.KeyChar)
        Select Case KeyAscii
            Case 39, 34, 96
                KeyAscii = 0

        End Select
        e.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            e.Handled = True
        End If
        AllowNumericValueInTextBox(txtA4orginal, e)
    End Sub
    Private Sub txtA4Reprint_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtA4Reprint.Validating
        Try
            Dim intLoopCounter As Integer
            If Val(txtA4orginal.Text) < Val(txtA4Reprint.Text) Then
                MsgBox("Reprint cant be more than Orignal Copy !", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100))
                spA4grid_Reprint.MaxRows = 0
                txtA4Reprint.Text = "0"
                Exit Sub
            Else

                spA4grid_Reprint.MaxRows = 0
                For intLoopCounter = 1 To Val(txtA4Reprint.Text)
                    Call AddBlankRow_Reprint()
                Next
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, ResolveResString(100))
        End Try
    End Sub

    
    Private Sub cmdGSTBillStateHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdGSTBillStateHelp.Click
        'GST CHANGES
        On Error GoTo errHandler
        Dim strGSTBILLSTATE() As String
        Dim strString As String
        Dim strcode As String
        strString = txtGSTBillState.Text & "%"
        strGSTBILLSTATE = ctlCurrCode.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "SELECT STATE_CODE, STATE_NAME from VW_GST_STATE_MST " & strcode & " ", "GST state Listing", 1)
        If UBound(strGSTBILLSTATE) = -1 Then Me.txtGSTBillState.Focus() : Exit Sub
        If (strGSTBILLSTATE(0)) = "0" Then
            ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
            txtCreditTermId.Text = ""
            txtCreditTermId.Focus()
        Else
            txtGSTBillState.Text = Trim(strGSTBILLSTATE(0))
            lblGSTStateDesc.Text = Trim(strGSTBILLSTATE(1))
            txtGSTBillState.Focus()
        End If
        Exit Sub
errHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    '101482956
    Private Sub VisibleTaxRegistrationNumber()
        Try
            If plantName = "MTL" Then
                lblTRNNo.Visible = True
                txtTRNNo.Enabled = True
                txtTRNNo.Visible = True
            Else
                lblTRNNo.Visible = False
                txtTRNNo.Enabled = False
                txtTRNNo.Visible = False
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    '101482956
    Private Sub txtTRNNo_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtTRNNo.KeyPress
        On Error GoTo ErrHandler
        Dim keyascii As Short
        keyascii = Asc(e.KeyChar)
        Select Case keyascii
            Case 39, 34, 96, 45
                e.Handled = True
        End Select
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    '101482956
    Private Sub txtTRNNo_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTRNNo.Enter
        On Error GoTo ErrHandler
        Me.txtTRNNo.SelectionStart = 0
        Me.txtTRNNo.SelectionLength = Len(txtTRNNo.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub AddShippingDetailinNewRow(ByVal pstrCustomerCode As String)
        '-------------------------------------------------------------------------------------------
        ' Function      : Show the Customer Shipping Detail in View and Edit Mode
        '---------------------------------------------------------------------------------------------
        Dim rsgetDetail As ClsResultSetDB
        Dim strQuery As String
        Dim intMaxRecord As Short
        Dim intLoopCounter As Short
        On Error GoTo ErrHandler
        With spShippingAddess
            rsgetDetail = New ClsResultSetDB
            strQuery = "select Customer_Code,Shipping_Code,Shipping_Desc,Default_Address,InActive_Flag,InActive_date,Ship_Address1,Ship_Address2,Ship_City,Ship_dist,Ship_State,Ship_Country,Ship_Pin,Ship_Phone,Ship_Fax,Ship_email_id,Ship_Contact_person,Ship_Person_desig,GST_STATE_CODE,GST_STATE_DESC,GSTIN_ID,ISNULL(DISTANCE_FROM_UNIT_KM,0) DISTANCE_FROM_UNIT_KM from Customer_Shipping_Dtl Where Unit_Code='" & gstrUNITID & "' And customer_code='" & Trim(pstrCustomerCode) & "' and Default_Address=1 order by Shipping_code,default_Address,inactive_flag "
            rsgetDetail.GetResult(strQuery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
            If rsgetDetail.GetNoRows > 0 Then
                intMaxRecord = .MaxRows
                rsgetDetail.MoveFirst()
                intLoopCounter = intMaxRecord
                Call .SetText(enmshipdetail.VAL_DEFAULT, intLoopCounter, 0)
                Call .SetText(enmshipdetail.VAL_INACTIVE, intLoopCounter, IIf(rsgetDetail.GetValue("Inactive_Flag") = "True", 1, 0))
                If rsgetDetail.GetValue("Inactive_Flag") = True Then
                    .Row = intLoopCounter
                    .Col = enmshipdetail.VAL_INACTIVE
                    .Lock = True
                End If
                Call .SetText(enmshipdetail.VAL_SHIPCODE, intLoopCounter, rsgetDetail.GetValue("Shipping_Code"))
                Call .SetText(enmshipdetail.VAL_SHIPDESC, intLoopCounter, rsgetDetail.GetValue("Shipping_Desc"))
                Call .SetText(enmshipdetail.VAL_SHIPADD1, intLoopCounter, rsgetDetail.GetValue("Ship_Address1"))
                Call .SetText(enmshipdetail.VAL_SHIPADD2, intLoopCounter, rsgetDetail.GetValue("Ship_Address2"))
                Call .SetText(enmshipdetail.VAL_CITY, intLoopCounter, rsgetDetail.GetValue("Ship_City"))
                Call .SetText(enmshipdetail.VAL_DISCT, intLoopCounter, rsgetDetail.GetValue("Ship_Dist"))
                Call .SetText(enmshipdetail.VAL_STATE, intLoopCounter, rsgetDetail.GetValue("Ship_State"))
                Call .SetText(enmshipdetail.VAL_COUNTRY, intLoopCounter, rsgetDetail.GetValue("Ship_country"))
                Call .SetText(enmshipdetail.VAL_Pin, intLoopCounter, rsgetDetail.GetValue("Ship_pin"))
                Call .SetText(enmshipdetail.VAL_PHONE, intLoopCounter, rsgetDetail.GetValue("Ship_Phone"))
                Call .SetText(enmshipdetail.VAL_FAX, intLoopCounter, rsgetDetail.GetValue("Ship_fax"))
                Call .SetText(enmshipdetail.VAL_EMAILID, intLoopCounter, rsgetDetail.GetValue("Ship_email_id"))
                Call .SetText(enmshipdetail.VAL_CONTACTPERSON, intLoopCounter, rsgetDetail.GetValue("Ship_contact_person"))
                Call .SetText(enmshipdetail.VAL_DESIGNATION, intLoopCounter, rsgetDetail.GetValue("Ship_person_desig"))

                If (rsgetDetail.GetValue("Inactive_Flag")) = True Then
                    Call .SetText(enmshipdetail.VAL_INACTIVEDATE, intLoopCounter, VB6.Format(rsgetDetail.GetValue("Inactive_Date"), gstrDateFormat))
                End If

                Call .SetText(enmshipdetail.VAL_GSTSTATECODE, intLoopCounter, rsgetDetail.GetValue("GST_STATE_CODE"))
                Call .SetText(enmshipdetail.VAL_GSTSTATEDESC, intLoopCounter, rsgetDetail.GetValue("GST_STATE_DESC"))
                Call .SetText(enmshipdetail.VAL_SHIP_GSTIN_ID, intLoopCounter, rsgetDetail.GetValue("GSTIN_ID"))
                Call .SetText(enmshipdetail.VAL_DISTANCE_FROM_UNIT, intLoopCounter, rsgetDetail.GetValue("DISTANCE_FROM_UNIT_KM"))
                rsgetDetail.MoveNext()

            End If
            rsgetDetail.ResultSetClose()
            rsgetDetail = Nothing
        End With
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub FillASN_Status()
        Dim StatusResultset1 As ClsResultSetDB
        StatusResultset1 = New ClsResultSetDB
        Dim strQuery As String
        strQuery = "select ASN_ENABLED from VW_TATA_MAHINDRA_CUSTOMER Where Unit_Code='" & gstrUNITID & "' And customer_code='" & txtCustCode.Text & "' "
        StatusResultset1.GetResult(strQuery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
        If StatusResultset1.GetNoRows > 0 Then
            Me.GroupBox1.Visible = True
            Me.chkASNEnable.Enabled = True
            Me.btnASNEnabled.Enabled = True
            If StatusResultset1.GetValue("ASN_ENABLED") = True Then
                Me.chkASNEnable.Checked = True
            Else
                Me.chkASNEnable.Checked = False
            End If

        Else
            Me.GroupBox1.Visible = False

        End If
    End Sub

    Private Sub BtnASNEnabled_Click(sender As Object, e As EventArgs) Handles btnASNEnabled.Click

        Dim lstrSqL As String
        Dim ASN_bool As Boolean
        Dim status As String
        Dim strSQL As String = String.Empty

        If chkASNEnable.Checked Then
            ASN_bool = 1
            status = "Enabled"
        Else
            ASN_bool = 0
            status = "disabled"
        End If
        strSQL = "SELECT TOP 1 1 FROM CUSTOMER_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & txtCustCode.Text & "' AND ISNULL(PRINT_METHOD,'')='TATA'"
        If IsRecordExists(strSQL) Then
            lstrSqL = "UPDATE CUSTOMER_MST SET TML_ASN_ENABLED='" & ASN_bool & "'   Where Unit_Code='" & gstrUNITID & "' And customer_code='" & txtCustCode.Text & "' AND ISNULL(PRINT_METHOD,'')='TATA'  "
        End If

        strSQL = "SELECT TOP 1 1 FROM CUSTOMER_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & txtCustCode.Text & "' AND GROUP_MAHINDRA=1 "
        If IsRecordExists(strSQL) Then
            lstrSqL = "UPDATE CUSTOMER_MST SET mahindra_ASN_ENABLED='" & ASN_bool & "'   Where Unit_Code='" & gstrUNITID & "' And customer_code='" & txtCustCode.Text & "' AND GROUP_MAHINDRA=1  "
        End If

        If UpdateRecordInTableNew(lstrSqL) = True Then
            MsgBox("ASN is Successfully ' " & status & "'. ", MsgBoxStyle.Information, "eMPro")
        Else
            'gblnCancelUnload = True : gblnFormAddEdit = True
            ' SqlConnectionclass.RollbackTran()
            Exit Sub
        End If
    End Sub


End Class