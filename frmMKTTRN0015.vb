Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Friend Class frmMKTTRN0015
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
	'changes done by nisha on 22/02/2003
	'1.Financial Rollover
	'2.temp invoice series 99000000
    '3.Update String in Sale_dtl 24/03/200327/03/2003
    ' Revised By                 -   Roshan Singh
    ' Revision Date              -   01 JUN 2011
    ' Description                -   FOR MULTIUNIT FUNCTIONALITY
    '===================================================================================
    Dim mStrCustMst As String
    Dim mresult As ClsResultSetDB
    Dim mintFormIndex As Short
    Dim salesconf As String
    Dim msubTotal, mInvNo, mExDuty, mBasicAmt, mOtherAmt As Double
    Dim mFrAmt, mGrTotal, mStAmt, mCustmtrl As Double
    Dim mDoc_No As Short
    Dim mAccount_Code, mInvType, mSubCat, mlocation As String
    Dim mstrAnnex As String
    Dim arrQty() As Double 'used in BomCheck() insertupdateAnnex()
    Dim arrItem() As String 'used in BomCheck() insertupdateAnnex()
    Dim arrReqQty() As Double
    Dim arrCustAnnex(0, 0) As Object
    Dim ref57f4 As String 'used in BomCheck() insertupdateAnnex()
    Dim dblFinishedQty As Double 'To get Qty of Finished Item from Spread
    Dim strCustCode As String 'used in BomCheck() insertupdateAnnex()
    Dim StrItemCode As String 'used in BomCheck() insertupdateAnnex()
    Dim strUPdateSaleDtl As String
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
    Dim rsBomMst As ClsResultSetDB
    Dim mstrMasterString As String 'To store master string for passing to Dr Cr COM
    Dim mstrDetailString As String 'To store detail string for passing to Dr Cr COM
    Dim mstrPurposeCode As String 'To store the Purpose Code which will be used for the fetching of GL and SL
    Dim mblnAddCustomerMaterial As Boolean 'To decide whether to add customer material in basic or not
    Dim mblnSameSeries As Boolean 'To store the flag whether the selected invoice will have same series as others
    Dim mstrReportFilename As String 'To store the report filename
    Dim mblnInsuranceFlag As Boolean 'To store insurance flag
    Dim mblnpostinfin As Boolean
    Dim mSaleConfNo As Double

    Private Sub cmdUnitCodeList_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdUnitCodeList.Click
        On Error GoTo ErrHandler
        Call ShowCode_Desc("SELECT Unt_CodeID,unt_unitname FROM Gen_UnitMaster WHERE Unt_Status=1 and Unt_CodeID= '" & gstrUNITID & "'", txtUnitCode)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtUnitCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUnitCode.TextChanged
        If Trim(txtUnitCode.Text) = "" Then
            cmbInvType.Enabled = False
            cmbInvType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            cmbInvType.SelectedIndex = -1
            CmbCategory.Enabled = False
            CmbCategory.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            CmbCategory.SelectedIndex = -1
        End If
    End Sub
    Private Sub txtUnitCode_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUnitCode.Enter
        On Error GoTo ErrHandler
        'Selecting the text on focus
        With txtUnitCode
            .SelectionStart = 0 : .SelectionLength = Len(Trim(.Text))
        End With
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
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
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub txtUnitCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtUnitCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then
            System.Windows.Forms.SendKeys.Send("{TAB}")
            'Supressing ¬ ¤ ¦ » characters since these are being used as string delimiters
        ElseIf KeyAscii = 187 Or KeyAscii = 166 Or KeyAscii = 164 Or KeyAscii = 172 Then
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
        'Populate the details
        mobjGLTrans = New prj_GLTransactions.cls_GLTransactions(gstrUNITID, GetServerDate)
        strUnitDesc = mobjGLTrans.GetUnit(Trim(txtUnitCode.Text), ConnectionString:=gstrCONNECTIONSTRING)
        mobjGLTrans = Nothing
        If Trim(txtUnitCode.Text) = "" Then GoTo EventExitSub
        If CheckString(strUnitDesc) <> "Y" Then
            MsgBox(CheckString(strUnitDesc), MsgBoxStyle.Critical, "eMPro")
            txtUnitCode.Text = ""
            cmbInvType.Enabled = False
            cmbInvType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            cmbInvType.SelectedIndex = -1
            CmbCategory.Enabled = False
            CmbCategory.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            CmbCategory.SelectedIndex = -1
            Cancel = True
        Else
            If mblnEOUUnit = True Then
                Call selectDataFromSaleConf(Trim(txtUnitCode.Text), cmbInvType, "Description", "'INV'", "datediff(dd,getdate(),fin_start_date)<=0  and datediff(dd,fin_end_date,getdate())<=0")
            Else
                Call selectDataFromSaleConf(Trim(txtUnitCode.Text), cmbInvType, "Description", "'INV'", "datediff(dd,getdate(),fin_start_date)<=0  and datediff(dd,fin_end_date,getdate())<=0")
            End If
            cmbInvType.Enabled = True
            cmbInvType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        End If
        GoTo EventExitSub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
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
    Private Sub CmbCategory_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmbCategory.SelectedIndexChanged
        On Error GoTo Err_Handler
        If Len(Trim(CmbCategory.Text)) = 0 Or Trim(CmbCategory.Text) = "-None-" Or Len(Trim(CmbCategory.Text)) > 0 Then
            lblcategory.Text = ""
            Ctlinvoice.Text = ""
        End If
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
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
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
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
            If UCase(lbldescription.Text) = "SMP" Then
                If rsSalesConf.State = ADODB.ObjectStateEnum.adStateOpen Then rsSalesConf.Close()
                rsSalesConf.Open("SELECT * FROM fin_GlobalGl WHERE gbl_prpsCode='Sample_Expences' and UNIT_CODE = '" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
                If rsSalesConf.EOF Then
                    MsgBox("Please define Sample Expences Account in Global Gl Definition", MsgBoxStyle.Information, "eMPro")
                    Me.CmbCategory.SelectedIndex = 0
                    Me.lblcategory.Text = ""
                    Me.cmbInvType.SelectedIndex = 0
                    Me.lbldescription.Text = ""
                    Me.cmbInvType.Focus()
                    GoTo EventExitSub
                End If
            End If
            If rsSalesConf.State = ADODB.ObjectStateEnum.adStateOpen Then rsSalesConf.Close()
            rsSalesConf.CursorLocation = ADODB.CursorLocationEnum.adUseClient
            rsSalesConf.Open("SELECT * FROM SaleConf WHERE Invoice_Type='" & lbldescription.Text & "' AND Sub_Type_description ='" & CmbCategory.Text & "' and UNIT_CODE = '" & gstrUNITID & "' AND Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,getdate(),fin_start_date)<=0  and datediff(dd,fin_end_date,getdate())<=0", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
            If Not rsSalesConf.EOF Then
                mstrPurposeCode = Trim(IIf(IsDBNull(rsSalesConf.Fields("inv_GLD_prpsCode").Value), "", rsSalesConf.Fields("inv_GLD_prpsCode").Value))
                mblnSameSeries = rsSalesConf.Fields("Single_Series").Value
                mstrReportFilename = Trim(IIf(IsDBNull(rsSalesConf.Fields("Report_filename").Value), "", rsSalesConf.Fields("Report_filename").Value))
                If mstrPurposeCode = "" Then
                    MsgBox("Please select a Purpose Code in Sales Configuration", MsgBoxStyle.Information, "eMPro")
                    Me.CmbCategory.SelectedIndex = 0
                    Me.lblcategory.Text = ""
                    Me.cmbInvType.SelectedIndex = 0
                    Me.lbldescription.Text = ""
                    Me.cmbInvType.Focus()
                    mstrPurposeCode = ""
                    GoTo EventExitSub
                End If
            Else
                MsgBox("No record found in Sales Configuration for the selected Location, Invoice Type and Sub-Category", MsgBoxStyle.Information, "eMPro")
                Me.CmbCategory.SelectedIndex = 0
                Me.lblcategory.Text = ""
                Me.cmbInvType.SelectedIndex = 0
                Me.lbldescription.Text = ""
                Me.cmbInvType.Focus()
                mstrPurposeCode = ""
                GoTo EventExitSub
            End If
            mresult = New ClsResultSetDB
            mresult.GetResult("Select sub_type,Sub_Type_Description,Stock_Location,updateStock_Flag  from SaleConf where Invoice_type = '" & Trim(Me.lbldescription.Text) & "' and UNIT_CODE = '" & gstrUNITID & "'  and sub_Type_Description = '" & Trim(Me.CmbCategory.Text) & "' and Location_Code ='" & Trim(txtUnitCode.Text) & "' and datediff(dd,getdate(),fin_start_date)<=0  and datediff(dd,fin_end_date,getdate())<=0")
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
                        Me.cmbInvType.SelectedIndex = 0
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
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
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
            CmbCategory.Enabled = False
            lblcategory.Text = ""
            CmbCategory.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            Ctlinvoice.Enabled = False
            frachkRequired.Enabled = False
            Ctlinvoice.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            Call cmbInvType_Validating(cmbInvType, New System.ComponentModel.CancelEventArgs(False))
            Exit Sub
        End If
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
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
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
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
            If mblnEOUUnit = True Then
                mresult = New ClsResultSetDB
                mresult.GetResult("Select distinct(Invoice_type),Description from SaleConf where Invoice_Type in('INV')and Description = '" & cmbInvType.Text & "' and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,getdate(),fin_start_date)<=0  and datediff(dd,fin_end_date,getdate())<=0 and UNIT_CODE = '" & gstrUNITID & "'")
            Else
                mresult = New ClsResultSetDB
                mresult.GetResult("Select distinct(Invoice_type),Description from SaleConf where Invoice_Type in('INV')and Description = '" & cmbInvType.Text & "' and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,getdate(),fin_start_date)<=0  and datediff(dd,fin_end_date,getdate())<=0 and UNIT_CODE = '" & gstrUNITID & "'")
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
                CmbCategory.Enabled = True
                CmbCategory.BackColor = System.Drawing.Color.White
                CmbCategory.Items.Clear()
                Call selectDataFromSaleConf(Trim(txtUnitCode.Text), CmbCategory, "Sub_Type_Description", "'" & Trim(lbldescription.Text) & "'", "datediff(dd,getdate(),fin_start_date)<=0  and datediff(dd,fin_end_date,getdate())<=0", "'F'")
                mresult.ResultSetClose()
                Me.CmbCategory.Focus()
            End If
        End If
        GoTo EventExitSub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
        Dim Index As Short = cmdHelp.GetIndex(eventSender)
        Dim strHelp As Object
        On Error GoTo Err_Handler
        Select Case Index
            Case 2
                With Me.Ctlinvoice
                    If optInvYes(0).Checked = True Then
                        strHelp = ShowList(1, .Maxlength, "", "Doc_No", "Invoice_Type", "SalesChallan_dtl", " and Invoice_Type = '" & Me.lbldescription.Text & "' and Sub_category = '" & Me.lblcategory.Text & "' and Doc_No >99000000 and bill_flag = 0 and Location_Code='" & Trim(txtUnitCode.Text) & "'")
                    Else
                        strHelp = ShowList(1, .Maxlength, "", "Doc_No", "Invoice_Type", "SalesChallan_dtl", " and Invoice_Type = '" & Me.lbldescription.Text & "' and Sub_category = '" & Me.lblcategory.Text & "' and Doc_No < 99000000 and bill_flag = 1 and Location_Code='" & Trim(txtUnitCode.Text) & "'")
                    End If
                    .Focus()
                End With
                If Val(strHelp) = -1 Then ' No record
                    Call ConfirmWindow(10512, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION)
                Else
                    Me.Ctlinvoice.Text = strHelp
                    If optInvYes(0).Checked = True Then
                        gobjDB.GetResult("SELECT Doc_NO,Invoice_Type FROM SalesChallan_Dtl Where Invoice_Type = '" & Me.lbldescription.Text & "' and Sub_category = '" & Me.lblcategory.Text & "' and Doc_No >99000000 and bill_flag =0 and Doc_No = '" & strHelp & "' and Location_Code='" & Trim(txtUnitCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "'")
                    Else
                        gobjDB.GetResult("SELECT Doc_NO,Invoice_Type FROM SalesChallan_Dtl Where Invoice_Type = '" & Me.lbldescription.Text & "' and Sub_category = '" & Me.lblcategory.Text & "' and Doc_No <99000000 and bill_flag =1 and Doc_No = '" & strHelp & "'  and Location_Code='" & Trim(txtUnitCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "'")
                    End If
                    If gobjDB.GetNoRows > 0 Then 'RECORD FOUND
                    End If
                End If
        End Select
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
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
                txtPLA.Focus()
                Exit Sub
            End If
        End If
        If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
            KeyAscii = KeyAscii
        Else
            KeyAscii = 0
        End If
        DirectCast(Sender, CtlGeneral).KeyPressKeyascii = KeyAscii
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
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
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub CtlInvoice_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles Ctlinvoice.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim strSql As String
        On Error GoTo Err_Handler
        If Len(Ctlinvoice.Text) = 0 Then GoTo EventExitSub
        mStrCustMst = "Select Doc_No,Invoice_type from SalesChallan_Dtl where Invoice_Type = 'INV' and sub_category = 'F' and UNIT_CODE = '" & gstrUNITID & "' and "
        If optInvYes(0).Checked = True Then
            mStrCustMst = mStrCustMst & " Location_Code='" & Trim(txtUnitCode.Text) & "'" & " and bill_flag =0 and Doc_No ="
        Else
            mStrCustMst = mStrCustMst & " Location_Code='" & Trim(txtUnitCode.Text) & "'" & " and bill_flag =1 and Doc_No ="
        End If
        strSql = mStrCustMst & Ctlinvoice.Text
        Me.Ctlinvoice.ExistRecQry = mStrCustMst
        mresult = New ClsResultSetDB
        mresult.GetResult(strSql)
        mresult.ResultSetClose()
        If Len(Ctlinvoice.Text) > 0 Then 'Checking if Item Field is not Blank
            If Ctlinvoice.ExistsRec = True Then 'Checking if the Record Exists
                Me.Cmdinvoice.Focus()
            Else
                Cancel = True
                Ctlinvoice.Text = ""
                Ctlinvoice.Focus()
                GoTo EventExitSub
            End If
        End If
        GoTo EventExitSub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub frmMKTTRN0015_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo Err_Handler
        mdifrmMain.CheckFormName = mintFormIndex
        frmModules.NodeFontBold(Tag) = True
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0015_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        On Error GoTo Err_Handler
        frmModules.NodeFontBold(Tag) = False
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub Form_Initialize_Renamed()
        On Error GoTo Err_Handler
        gobjDB = New ClsResultSetDB
        gobjDB.GetResult("SELECT EOU_Flag, CustSupp_Inc,InsExc_Excise,postinfin FROM sales_parameter WHERE UNIT_CODE = '" & gstrUNITID & "'")
        If gobjDB.GetValue("EOU_Flag") = True Then
            mStrCustMst = "Select Doc_No,Invoice_type from SalesChallan_Dtl where Invoice_Type <> 'EXP' and Location_Code='" & Trim(txtUnitCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "'"
            mblnEOUUnit = True
        Else
            mStrCustMst = "Select Doc_No,Invoice_type from SalesChallan_Dtl where Location_Code='" & Trim(txtUnitCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "'"
            mblnEOUUnit = False
        End If
        mblnAddCustomerMaterial = gobjDB.GetValue("CustSupp_Inc")
        mblnInsuranceFlag = gobjDB.GetValue("InsExc_Excise")
        mblnpostinfin = gobjDB.GetValue("postinfin")
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0015_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0 Then Call ctlFormHeader1_Click(ctlFormHeader1, New System.EventArgs()) : Exit Sub
    End Sub
    Private Sub frmMKTTRN0015_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo Err_Handler
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
        optYes(1).Checked = True
        optInvYes(0).Checked = True
        Me.chkLockPrintingFlag.Enabled = True
        Call Form_Initialize_Renamed()
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub selectDataFromSaleConf(ByRef LocationCode As String, ByRef combo As System.Windows.Forms.ComboBox, ByRef feild As String, ByRef invoicetype As String, ByRef pstrCondition As String, Optional ByRef SubType As String = "")
        Dim strSql As String
        Dim rsSaleConf As ClsResultSetDB
        Dim intRowCount As Short
        Dim intLoopCount As Short
        On Error GoTo Err_Handler
        If Len(Trim(SubType)) > 0 Then
            strSql = "select Distinct(" & feild & ") from Saleconf where Location_Code='" & LocationCode & "' and Invoice_Type in(" & invoicetype & ") and sub_type in (" & SubType & ") and UNIT_CODE = '" & gstrUNITID & "' and " & pstrCondition
        Else
            strSql = "select Distinct(" & feild & ") from Saleconf where Location_Code='" & LocationCode & "' and UNIT_CODE = '" & gstrUNITID & "' and Invoice_Type in(" & invoicetype & ") and " & pstrCondition
        End If
        rsSaleConf = New ClsResultSetDB
        rsSaleConf.GetResult(strSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsSaleConf.GetNoRows > 0 Then
            intRowCount = rsSaleConf.GetNoRows
            VB6.SetItemString(combo, 0, "-None-")
            rsSaleConf.MoveFirst()
            For intLoopCount = 1 To intRowCount
                VB6.SetItemString(combo, intLoopCount, rsSaleConf.GetValue(feild))
                rsSaleConf.MoveNext()
            Next intLoopCount
        End If
        rsSaleConf.ResultSetClose()
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0015_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error GoTo Err_Handler
        'Removing the form name from list
        mdifrmMain.RemoveFormNameFromWindowList = mintFormIndex
        'Setting the corresponding node's tag
        frmModules.NodeFontBold(Tag) = False
        'Releasing the form reference
        Me.Dispose()
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Sub ValuetoVariables()
        Dim strSql As String
        Dim rsSalesChallan As ClsResultSetDB
        Dim strInvoiceDate As String
        On Error GoTo Err_Handler
        strSql = "select INVOICE_DATE from Saleschallan_Dtl where Doc_No =" & Me.Ctlinvoice.Text & "  and Location_Code='" & Trim(txtUnitCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "'"
        rsSalesChallan = New ClsResultSetDB
        rsSalesChallan.GetResult(strSql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        strInvoiceDate = VB6.Format(rsSalesChallan.GetValue("Invoice_Date"), gstrDateFormat)
        mInvType = Me.lbldescription.Text
        mSubCat = Me.lblcategory.Text
        mInvNo = CDbl(GenerateInvoiceNo(mInvType, mSubCat, strInvoiceDate))
        strSql = " Select Asseccable= isnull(SUM(Accessible_amount),0) from sales_dtl "
        strSql = strSql & " where Doc_No =" & Me.Ctlinvoice.Text & " and Location_Code='" & Trim(txtUnitCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "'"
        mresult = New ClsResultSetDB
        mresult.GetResult(strSql)
        mAssessableValue = mresult.GetValue("Asseccable")
        mresult.ResultSetClose()
        rsSalesChallan.ResultSetClose()
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Sub updatesalesconfandsaleschallan()
        Dim strSql As String
        Dim rsSalesChallan As ClsResultSetDB
        Dim dblInvoiceAmt As Double
        Dim strInvoiceDate As String
        On Error GoTo Err_Handler
        strSql = "select *  from Saleschallan_dtl where Doc_No = " & Me.Ctlinvoice.Text
        strSql = strSql & " and Invoice_type = '" & mInvType & "'  and  sub_category =  '" & mSubCat & "' and Location_Code='" & Trim(txtUnitCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "'"
        rsSalesChallan = New ClsResultSetDB
        rsSalesChallan.GetResult(strSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsSalesChallan.GetNoRows > 0 Then
            mAccount_Code = rsSalesChallan.GetValue("Account_Code")
            mCust_Ref = rsSalesChallan.GetValue("Cust_ref")
            mAmendment_No = rsSalesChallan.GetValue("Amendment_No")
            dblInvoiceAmt = rsSalesChallan.GetValue("total_amount")
            strInvoiceDate = getDateForDB(VB6.Format(rsSalesChallan.GetValue("Invoice_date"), gstrDateFormat))
        End If
        rsSalesChallan.ResultSetClose()
        If mblnEOUUnit = True Then
            If UCase(lbldescription.Text) <> "EXP" Then
                If Not mblnSameSeries Then
                    salesconf = "update saleconf set current_No = " & mSaleConfNo & ", OpenningBal = openningBal - " & mAssessableValue & " where  UNIT_CODE = '" & gstrUNITID & "' AND  Invoice_type <> 'EXP' and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & strInvoiceDate & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & strInvoiceDate & "')<=0 "
                Else
                    salesconf = "update saleconf set current_No = " & mSaleConfNo & " where UNIT_CODE = '" & gstrUNITID & "' AND Single_Series = 1 and Invoice_type <> 'EXP' and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & strInvoiceDate & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & strInvoiceDate & "')<=0" & vbCrLf
                    salesconf = salesconf & "update saleconf set OpenningBal = openningBal - " & mAssessableValue & " where UNIT_CODE = '" & gstrUNITID & "' AND Invoice_type <> 'EXP' and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & strInvoiceDate & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & strInvoiceDate & "')<=0"
                End If
            Else
                salesconf = "update saleconf set current_No = " & mSaleConfNo & " where  UNIT_CODE = '" & gstrUNITID & "' AND Invoice_type = 'EXP' and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & strInvoiceDate & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & strInvoiceDate & "')<=0"
            End If
        Else
            If Not mblnSameSeries Then
                salesconf = "update saleconf set current_No = " & mSaleConfNo & " where UNIT_CODE = '" & gstrUNITID & "' AND Invoice_type = '" & Me.lbldescription.Text & "' and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & strInvoiceDate & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & strInvoiceDate & "')<=0"
            Else
                salesconf = "update saleconf set current_No = " & mSaleConfNo & " where UNIT_CODE = '" & gstrUNITID & "' AND Single_Series = 1 and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & strInvoiceDate & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & strInvoiceDate & "')<=0"
            End If
        End If
        saleschallan = "UPDATE SalesChallan_Dtl SET doc_no=" & mInvNo & ", total_amount = " & System.Math.Round(dblInvoiceAmt, 0) & ", Bill_Flag=1,print_flag = 1 WHERE UNIT_CODE = '" & gstrUNITID & "' AND Doc_No=" & Ctlinvoice.Text & " and Location_Code='" & Trim(txtUnitCode.Text) & "'"
        strUPdateSaleDtl = "UPDATE Sales_Dtl SET doc_no=" & mInvNo & " WHERE UNIT_CODE = '" & gstrUNITID & "' AND Doc_No=" & Ctlinvoice.Text & " and Location_Code='" & Trim(txtUnitCode.Text) & "'"
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Function ValidSelection() As Boolean
        Dim blnInvalidData As Boolean
        Dim strErrMsg As String
        Dim ctlBlank As System.Windows.Forms.Control
        Dim lNo As Integer
        On Error GoTo Err_Handler
        ValidRecord = False
        lNo = 1
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
            Call MsgBox(strErrMsg, MsgBoxStyle.Information, "Error")
            ctlBlank.Focus()
            Exit Function
        End If
        ValidRecord = True
        ValidSelection = True
        Exit Function
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Public Sub RefreshForm()
        On Error GoTo ErrHandler
        Call EnableControls(False, Me, True)
        optInvYes(0).Enabled = True : optInvYes(1).Enabled = True : optInvYes(0).Checked = True
        Me.cmbInvType.Enabled = True : Me.cmbInvType.BackColor = System.Drawing.Color.White : Me.cmbInvType.Focus()
        Me.chkLockPrintingFlag.Enabled = True
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Public Function InvoiceGeneration(ByVal objRpt As ReportDocument, ByVal frmReportViewer As eMProCrystalReportViewer) As Boolean
        Dim rsCompMst As ClsResultSetDB
        Dim Phone, Range, RegNo, EccNo, Address, Invoice_Rule As String
        Dim CST, PLA, Fax, EMail, UPST, Division As String
        Dim Commissionerate As String
        Dim strSql As String
        Dim strCompMst, DeliveredAdd As String
        On Error GoTo Err_Handler
        rsCompMst = New ClsResultSetDB
        strSql = "{SalesChallan_Dtl.Location_Code}='" & Trim(txtUnitCode.Text) & "'  and {SalesChallan_Dtl.Doc_No} =" & Trim(Ctlinvoice.Text) & " AND {SalesChallan_Dtl.UNIT_CODE}='" & gstrUNITID & "'  and {SalesChallan_Dtl.Invoice_Type}"
        strSql = strSql & " = '" & Trim(Me.lbldescription.Text) & "'  and {SalesChallan_Dtl.Sub_Category} = '" & Trim(Me.lblcategory.Text) & "'"
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
            Call UpdateinSale_Dtl()
        End If
        If UCase(lbldescription.Text) = "REJ" Then
            If Len(Trim(mCust_Ref)) > 0 Then
                Call UpdateGrnHdr(CDbl(mCust_Ref), mInvNo)
            End If
        End If
        If UCase(lbldescription.Text) = "JOB" Then
            mP_Connection.Execute("DELETE FROM tempCustAnnex WHERE UNIT_CODE ='" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords) ' to delete all the records from table before inserting new one for selected invoice
            If BomCheck() = False Then
                InvoiceGeneration = False
                Exit Function
            End If
        End If
        Address = gstr_WRK_ADDRESS1 & gstr_WRK_ADDRESS2
        rsCompMst.ResultSetClose()
        '*******************To Calculate Value of Delivery Address in Case of Delivery Address requird
        rsCompMst = New ClsResultSetDB
        rsCompMst.GetResult("Select a.* from Customer_Mst a, saleschallan_dtl b where a.Customer_code = b.Account_code AND A.UNIT_CODE = B.UNIT_CODE AND A.UNIT_CODE = '" & gstrUNITID & "'  and b.Doc_No = " & Ctlinvoice.Text & " and b.Location_Code='" & Trim(txtUnitCode.Text) & "'")
        If rsCompMst.GetNoRows > 0 Then
            DeliveredAdd = Trim(rsCompMst.GetValue("Ship_address1"))
            If Len(Trim(DeliveredAdd)) Then
                DeliveredAdd = Trim(DeliveredAdd) & "," & Trim(rsCompMst.GetValue("Ship_address2"))
            Else
                DeliveredAdd = Trim(rsCompMst.GetValue("Ship_address2"))
            End If
        End If
        '********************************** to Check Invoice Cust supp Detail
        If InvoiceCustSuppBOM(CDbl(Ctlinvoice.Text)) = False Then
            InvoiceGeneration = False
            Exit Function
        Else
            If ToCheckDrgForCustSupp(CDbl(Ctlinvoice.Text)) = False Then
                InvoiceGeneration = False
                Exit Function
            Else
                If ToCheckItemRateForCustMtrl(CDbl(Ctlinvoice.Text)) = False Then
                    InvoiceGeneration = False
                    Exit Function
                Else
                    mP_Connection.Execute("Exec MKTCustSuppMat '" & gstrUNITID & "'," & Ctlinvoice.Text, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                End If
            End If
        End If
        rsCompMst.ResultSetClose()
        objRpt.Load(My.Application.Info.DirectoryPath & "\Reports\rptinvoiceCSI.rpt")
        If UCase(cmbInvType.Text) <> "JOBWORK INVOICE" Then
            objRpt.DataDefinition.FormulaFields("Category").Text = "'" & Me.lblcategory.Text & "'"
        End If
        objRpt.DataDefinition.FormulaFields("Registration").Text = "'" & RegNo & "'"
        objRpt.DataDefinition.FormulaFields("ECC").Text = "'" & EccNo & "'"
        objRpt.DataDefinition.FormulaFields("Range").Text = "'" & Range & "'"
        objRpt.DataDefinition.FormulaFields("CompanyName").Text = "'" & gstrCOMPANY & "'"
        objRpt.DataDefinition.FormulaFields("CompanyAddress").Text = "'" & Address & "'"
        objRpt.DataDefinition.FormulaFields("Phone").Text = "'" & Phone & "'"
        objRpt.DataDefinition.FormulaFields("Fax").Text = "'" & Fax & "'"
        If UCase(cmbInvType.Text) <> "JOBWORK INVOICE" Then
            objRpt.DataDefinition.FormulaFields("EMail").Text = "'" & EMail & "'"
        End If
        objRpt.DataDefinition.FormulaFields("PLA").Text = "'" & PLA & "'"
        objRpt.DataDefinition.FormulaFields("UPST").Text = "'" & UPST & "'"
        objRpt.DataDefinition.FormulaFields("CST").Text = "'" & CST & "'"
        objRpt.DataDefinition.FormulaFields("Division").Text = "'" & Division & "'"
        objRpt.DataDefinition.FormulaFields("commissionerate").Text = "'" & Commissionerate & "'"
        objRpt.DataDefinition.FormulaFields("InvoiceRule").Text = "'" & Invoice_Rule & "'"
        objRpt.DataDefinition.FormulaFields("EOUFlag").Text = "'" & mblnEOUUnit & "'"
        If optYes(0).Checked = True Then
            objRpt.DataDefinition.FormulaFields("DeliveredAt").Text = "'Delivered At'"
            objRpt.DataDefinition.FormulaFields("Address2").Text = "'" & DeliveredAdd & "'"
        Else
            objRpt.DataDefinition.FormulaFields("DeliveredAt").Text = "''"
            objRpt.DataDefinition.FormulaFields("Address2").Text = "''" 'to pass blanck Address in this case will overwrite this Formula written in Crystal Report for else case
        End If
        objRpt.DataDefinition.FormulaFields("PLADuty").Text = "'" & Trim(txtPLA.Text) & "'"
        objRpt.DataDefinition.FormulaFields("InsuranceFlag").Text = "'" & mblnInsuranceFlag & "'"
        objRpt.DataDefinition.FormulaFields("StringYear").Text = "'" & Year(GetServerDate) & "'"
        Dim strInvoiceDate As String
        Dim dblExistingInvNo As Double
        Dim strsql1 As String
        Dim rsSalesInvoiceDate As ClsResultSetDB
        Dim rsSalesConf As ClsResultSetDB
        Dim strSuffix As String
        If optInvYes(0).Checked = True Then
            objRpt.DataDefinition.FormulaFields("InvoiceNo").Text = "'" & mSaleConfNo & "'"
        Else
            strsql1 = "select * from Saleschallan_Dtl where Doc_No =" & Me.Ctlinvoice.Text & "  and Location_Code='" & Trim(txtUnitCode.Text) & "' AND UNIT_CODE='" & gstrUNITID & "'"
            rsSalesInvoiceDate = New ClsResultSetDB
            rsSalesInvoiceDate.GetResult(strsql1, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            strInvoiceDate = VB6.Format(rsSalesInvoiceDate.GetValue("Invoice_Date"), gstrDateFormat)
            rsSalesInvoiceDate.ResultSetClose()
            rsSalesConf = New ClsResultSetDB
            rsSalesConf.GetResult("Select Suffix from SaleConf Where Description ='" & cmbInvType.Text & "' AND Location_Code ='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0  AND UNIT_CODE='" & gstrUNITID & "'")
            strSuffix = rsSalesConf.GetValue("Suffix")
            rsSalesConf.ResultSetClose()
            If Len(Trim(strSuffix)) > 0 Then
                If Val(strSuffix) > 0 Then
                    dblExistingInvNo = Val(Mid(CStr(Ctlinvoice.Text), Len(Trim(strSuffix)) + 1))
                Else
                    dblExistingInvNo = CDbl(Ctlinvoice.Text)
                End If
            Else
                dblExistingInvNo = CDbl(Ctlinvoice.Text)
            End If
            objRpt.DataDefinition.FormulaFields("InvoiceNo").Text = "'" & dblExistingInvNo & "'"
        End If
        objRpt.RecordSelectionFormula = strSql
        InvoiceGeneration = True
        Exit Function
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Public Sub InitializeValues()
        On Error GoTo ErrHandler
        mExDuty = 0 : mInvNo = 0 : mBasicAmt = 0 : msubTotal = 0 : mOtherAmt = 0 : mGrTotal = 0 : mStAmt = 0 : mFrAmt = 0
        mDoc_No = 0 : mCustmtrl = 0 : mAmortization = 0 : mstrAnnex = "" : strupdateGrinhdr = "" : mblnCustSupp = False : strUPdateSaleDtl = ""
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Public Function ParentQty(ByRef pstrItemCode As String, ByRef pstrfinished As Object) As Double
        Dim rsParentQty As ClsResultSetDB
        Dim strParentQty As String
        On Error GoTo ErrHandler
        strParentQty = "select sum(required_qty + waste_Qty) as TotalQty from Bom_Mst where finished_Product_code ='"
        strParentQty = strParentQty & pstrfinished & "' and rawMaterial_Code ='" & pstrItemCode & "' and UNIT_CODE = '" & gstrUNITID & "'"
        rsParentQty = New ClsResultSetDB
        rsParentQty.GetResult(strParentQty, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        ParentQty = rsParentQty.GetValue("TotalQty")
        rsParentQty.ResultSetClose()
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
                strannex = strannex & " WHERE Item_code ='" & parrCustAnnex(0, intLoopCount) & "' and Customer_code ='"
                strannex = strannex & strCustCode & "' and UNIT_CODE = '" & gstrUNITID & "'"
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
                        mstrAnnex = mstrAnnex & " WHERE Item_code ='" & parrCustAnnex(0, intLoopCount) & "' and UNIT_CODE = '" & gstrUNITID & "' and Customer_code ='"
                        mstrAnnex = mstrAnnex & strCustCode & "' and Ref57f4_No ='" & strRef57F4 & "' "
                        mstrAnnex = mstrAnnex & "Insert into CustAnnex_dtl (Doc_Ty,"
                        mstrAnnex = mstrAnnex & "Invoice_No,Invoice_Date,ref57f4_Date,Ref57f4_No,"
                        mstrAnnex = mstrAnnex & "Item_Code,Quantity,"
                        mstrAnnex = mstrAnnex & "Customer_Code,"
                        mstrAnnex = mstrAnnex & "Location_Code,Product_Code,Ent_Userid,Ent_dt,"
                        mstrAnnex = mstrAnnex & "Upd_Userid,Upd_dt,UNIT_CODE) values ('O'," & mInvNo & ",GetDate(),'" & str57f4Date & "','"
                        mstrAnnex = mstrAnnex & ref57f4 & "','" & parrCustAnnex(0, intLoopCount) & "'," & parrCustAnnex(1, intLoopCount) & ","
                        mstrAnnex = mstrAnnex & "'" & strCustCode & "',"
                        mstrAnnex = mstrAnnex & "'SMIL','" & pstrFinishedItem & "','" & mP_User & "',GETDATE(),'"
                        mstrAnnex = mstrAnnex & mP_User & "',GETDATE(),'" & gstrUNITID & "')"
                        If dblbalanceqty < parrCustAnnex(1, intLoopCount) Then
                            mP_Connection.Execute(" insert into tempCustAnnex (Ref57f4_No,Annex_No,ref57f4_date,Item_code,Quantity,Balance_qty,finishedItem,UNIT_CODE) values ('" & strRef57F4 & "',0,'" & str57f4Date & "','" & parrCustAnnex(0, intLoopCount) & "'," & dblbalanceqty & ",0,'" & pstrFinishedItem & "','" & gstrUNITID & "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            parrCustAnnex(1, intLoopCount) = parrCustAnnex(1, intLoopCount) - dblbalanceqty
                        Else
                            mP_Connection.Execute(" insert into tempCustAnnex (Ref57f4_No,Annex_No,ref57f4_date,Item_code,Quantity,Balance_qty,finishedItem,UNIT_CODE) values ('" & strRef57F4 & "',0,'" & str57f4Date & "','" & parrCustAnnex(0, intLoopCount) & "'," & parrCustAnnex(1, intLoopCount) & "," & dblbalanceqty - parrCustAnnex(1, intLoopCount) & ",'" & pstrFinishedItem & "' ,'" & gstrUNITID & "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            parrCustAnnex(1, intLoopCount) = 0
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
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
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
        Dim intBomMaxItem As Short
        Dim rsCustAnnexDtl As ClsResultSetDB
        Dim rsSalesChallan As ClsResultSetDB
        Dim rsVandorBom As ClsResultSetDB
        Dim dblTotalReqQty As Double
        Dim strchallan As String
        Dim intAnnexMaxCount As Short
        On Error GoTo ErrHandler
        BomCheck = False
        rsSalesChallan = New ClsResultSetDB
        rsVandorBom = New ClsResultSetDB
        inti = 0
        intAnnexMaxCount = 0
        ReDim arrCustAnnex(3, intAnnexMaxCount)
        strchallan = " select a.Account_code,a.ref_Doc_No,a.Fifo_Flag,b.Item_Code,b.Sales_Quantity from "
        strchallan = strchallan & "salesChallan_dtl a,Sales_dtl b where a.Doc_No = " & Ctlinvoice.Text
        strchallan = strchallan & " and a.Location_Code = b.Location_Code and a.Doc_No = b.Doc_no and b.Location_Code='" & Trim(txtUnitCode.Text) & "' AND a.UNIT_CODE = b.UNIT_CODE AND a.UNIT_CODE = '" & gstrUNITID & "'"
        'Loop for Spread
        rsSalesChallan.GetResult(strchallan, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        intChallanMax = rsSalesChallan.GetNoRows
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
                strBomMst = strBomMst & " from Bom_Mst where Finished_Product_code ='"
                strBomMst = strBomMst & VarFinishedItem & "' and UNIT_CODE = '" & gstrUNITID & "' Order By Bom_Level"
                rsBomMst = New ClsResultSetDB
                rsBomMst.GetResult(strBomMst, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                intBomMaxItem = rsBomMst.GetNoRows
                rsBomMst.MoveFirst()
                If intBomMaxItem > 0 Then ' Item Found in Bom_Mst
                    rsVandorBom.GetResult("Select RawMaterial_Code from Vendor_bom where Finish_Product_code = '" & VarFinishedItem & "' and Vendor_code = '" & strCustCode & "' and UNIT_CODE = '" & gstrUNITID & "'")
                    If rsVandorBom.GetNoRows > 0 Then
                        'Loop for Parent Items of Items at First lavel
                        For intCurrentItem = 1 To intBomMaxItem
                            strBomItem = ""
                            strBomItem = rsBomMst.GetValue("RawMaterial_Code")
                            'String for CustAnnex_dtl
                            strCustAnnexDtl = "Select Item_Code,Balance_qty = sum(Balance_qty) from CustAnnex_hdr where Customer_code ='"
                            strCustAnnexDtl = strCustAnnexDtl & Trim(strCustCode) & "' and UNIT_CODE = '" & gstrUNITID & "'"
                            If blnFIFOFlag = False Then
                                strCustAnnexDtl = strCustAnnexDtl & " and ref57f4_no in("
                                strCustAnnexDtl = strCustAnnexDtl & Trim(strRef57F4) & ")"
                            End If
                            strCustAnnexDtl = strCustAnnexDtl & " and getdate() <= "
                            strCustAnnexDtl = strCustAnnexDtl & " DateAdd(d, 180, ref57f4_date)"
                            strCustAnnexDtl = strCustAnnexDtl & " and Item_code ='" & strBomItem & "' group by Item_code"
                            rsCustAnnexDtl = New ClsResultSetDB
                            rsCustAnnexDtl.GetResult(strCustAnnexDtl)
                            If rsCustAnnexDtl.GetNoRows >= 1 Then 'if item Found in Cust Annex
                                rsVandorBom.GetResult("Select RawMaterial_Code from Vendor_bom where Finish_Product_code = '" & VarFinishedItem & "'and RawMaterial_code = '" & strBomItem & "' and UNIT_CODE = '" & gstrUNITID & "'")
                                If rsVandorBom.GetNoRows > 0 Then
                                    'To Remove  that item from string will be used later for checking in case any item is not supplied
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
                                                arrReqQty(intArrCount) = arrReqQty(intArrCount) + (dblTotalReqQty * dblFinishedQty)
                                                If arrQty(intArrCount) < arrReqQty(intArrCount) Then 'in case if sum up is less then Quantity suplied in cust annex
                                                    MsgBox("Customer Supplied Materail for Item " & arrItem(inti) & "is" & arrQty(inti) & ".", MsgBoxStyle.OkOnly, "eMPro")
                                                    Cmdinvoice.Focus()
                                                    BomCheck = False
                                                    rsBomMst.ResultSetClose()
                                                    Exit Function
                                                Else
                                                    Call ToGetIteminAcustannex(arrCustAnnex, strBomItem, intAnnexMaxCount, dblTotalReqQty * dblFinishedQty)
                                                End If
                                            End If
                                        Next
                                        If blnItemFoundinArray = False Then
                                            'in case item not found in arrItem with help of blnItemFoundinarray = false then will add new value to Arrays
                                            arrItem(inti) = rsCustAnnexDtl.GetValue("Item_code")
                                            arrQty(inti) = rsCustAnnexDtl.GetValue("Balance_qty")
                                            arrReqQty(inti) = dblTotalReqQty * dblFinishedQty
                                            If arrQty(inti) < arrReqQty(inti) Then 'again  check for Quantity requird as compare to supplied in CustAnnex
                                                MsgBox("Customer Supplied Materail for Item " & arrItem(inti) & "is" & arrQty(inti) & ".", MsgBoxStyle.OkOnly, "eMPro")
                                                Cmdinvoice.Focus()
                                                BomCheck = False
                                                rsBomMst.ResultSetClose()
                                                Exit Function
                                            Else
                                                '*********** for Adding Values in CustAnnex Array
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
                                            MsgBox("Customer Supplied Materail for Item " & arrItem(inti) & "is" & arrQty(inti) & ".", MsgBoxStyle.OkOnly, "eMPro")
                                            Cmdinvoice.Focus()
                                            BomCheck = False
                                            rsBomMst.ResultSetClose()
                                            Exit Function
                                        Else
                                            '***********for Adding Values in CustAnnex Array
                                            If Len(Trim(arrCustAnnex(0, intAnnexMaxCount))) > 0 Then
                                                intAnnexMaxCount = intAnnexMaxCount + 1
                                                ReDim Preserve arrCustAnnex(3, intAnnexMaxCount)
                                            End If
                                            arrCustAnnex(0, intAnnexMaxCount) = rsCustAnnexDtl.GetValue("Item_code")
                                            arrCustAnnex(1, intAnnexMaxCount) = dblTotalReqQty * dblFinishedQty
                                            arrCustAnnex(2, intAnnexMaxCount) = (rsCustAnnexDtl.GetValue("Balance_qty") - (dblTotalReqQty * dblFinishedQty))
                                        End If
                                    End If
                                End If
                            Else ' if Item Not Found in Cust Annex
                                rsVandorBom = New ClsResultSetDB
                                rsVandorBom.GetResult("Select RawMaterial_Code from Vendor_bom where Finish_Product_code = '" & VarFinishedItem & "'and RawMaterial_code = '" & strBomItem & "' and UNIT_CODE = '" & gstrUNITID & "'")
                                If rsVandorBom.GetNoRows > 0 Then
                                    MsgBox("Item " & strBomItem & " is not supplied.", MsgBoxStyle.OkOnly, "eMPro")
                                    Cmdinvoice.Focus()
                                    BomCheck = False
                                    rsBomMst.ResultSetClose()
                                    Exit Function
                                Else ' if it'Process type is not I then Explore it Again in BOM_Mst
                                    dblFinishedQty = dblFinishedQty * rsBomMst.GetValue("TotalReqQty")
                                    If ExploreBom(strBomItem, dblFinishedQty, intSpCurrentRow, strCustCode, ref57f4, intAnnexMaxCount, CStr(VarFinishedItem)) = False Then
                                        BomCheck = False
                                        rsBomMst.ResultSetClose()
                                        Exit Function
                                    End If
                                End If
                            End If
                            rsBomMst.MoveNext()
                            inti = inti + 1
                        Next
                        intSpCurrentRow = intSpCurrentRow + 1 'for next spread item
                    Else
                        MsgBox("No BOM Defind for the Item (" & VarFinishedItem & ") defined in challan", MsgBoxStyle.Information, "eMPro")
                        BomCheck = False
                        rsBomMst.ResultSetClose()
                        Exit Function
                    End If
                Else
                    rsBomMst.ResultSetClose()
                    MsgBox("No Customer BOM Defind for Item (" & VarFinishedItem & ") defined in challan", MsgBoxStyle.Information, "eMPro")
                    BomCheck = False
                    Exit Function
                End If
                Call InsertUpdateAnnex(arrCustAnnex, VarFinishedItem, intAnnexMaxCount)
                rsBomMst.ResultSetClose()
            Next
        End If
        BomCheck = True
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Public Function ExploreBom(ByRef pstrItemCode As String, ByRef pstrFinishedQty As Object, ByRef pstrSPCurrentRow As Object, ByRef pstrCustCode As String, ByRef pstrRef As String, ByRef pintAnnexMaxCount As Short, ByRef pstrFinishedProduct As String) As Boolean
        '*****************************************************
        'Created By     -  Nisha
        'Description    -  to get the values of required items in Sub assambly bom
        'input Variable -  Item Code to Found, reqquantity of Finished Product,row in spread
        '*****************************************************
        Dim strBomMstRaw As String
        Dim rsBomMstRaw As ClsResultSetDB
        Dim rsCustAnnexDtl As ClsResultSetDB
        Dim rsVandorBom As ClsResultSetDB
        Dim intBomMaxRaw As Short
        Dim intCurrentRaw As Short
        Dim dblTotalReqQty As Double
        Dim strCustAnnexDtl As String
        Dim strref As String
        On Error GoTo ErrHandler
        strBomMstRaw = "Select RawMaterial_Code,Required_qty + Waste_qty "
        strBomMstRaw = strBomMstRaw & " As TotalReqQty,Process_Type from Bom_Mst where "
        strBomMstRaw = strBomMstRaw & " item_Code ='" & strBomItem
        strBomMstRaw = strBomMstRaw & "' and UNIT_CODE = '" & gstrUNITID & "' and  finished_product_code ='"
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
                strCustAnnexDtl = "Select Item_Code,Balance_qty,REF57F4_DATE from CustAnnex_hdr where  UNIT_CODE = '" & gstrUNITID & "' and Customer_code ='"
                strCustAnnexDtl = strCustAnnexDtl & Trim(pstrCustCode) & "'"
                If blnFIFOFlag = False Then
                    strref = Replace(pstrRef, "§", "','", 1)
                    strref = "'" & strref & "'"
                    strCustAnnexDtl = strCustAnnexDtl & " and ref57f4_no IN ("
                    strCustAnnexDtl = strCustAnnexDtl & Trim(strref) & ")"
                End If
                strCustAnnexDtl = strCustAnnexDtl & "  and getdate() <= "
                strCustAnnexDtl = strCustAnnexDtl & " DateAdd(d, 180, ref57f4_date)"
                strCustAnnexDtl = strCustAnnexDtl & " and Item_code ='" & strBomItem & "'"
                rsCustAnnexDtl = New ClsResultSetDB
                rsCustAnnexDtl.GetResult(strCustAnnexDtl, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                If rsCustAnnexDtl.GetNoRows >= 1 Then 'if item Found in CustAnnex then replace that item from Parant string
                    rsVandorBom = New ClsResultSetDB
                    rsVandorBom.GetResult("Select RawMaterial_Code from Vendor_bom where Finish_Product_code = '" & pstrFinishedProduct & "'and RawMaterial_code = '" & strBomItem & "' and UNIT_CODE = '" & gstrUNITID & "'")
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
                                ' if found then sum up Requird Quantity in array arrReqQty and assign value true to blnArrITemFound
                                blnArrItemFound = True
                                arrReqQty(intArrCount) = arrReqQty(intArrCount) + (dblTotalReqQty * pstrFinishedQty)
                                If arrQty(intArrCount) < arrReqQty(intArrCount) Then ' to Check with Quantity supplieded in Cust Annex
                                    MsgBox("Customer Supplied Materail for Item " & arrItem(inti) & " is " & arrQty(inti) & " .", MsgBoxStyle.OkOnly, "eMPro")
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
                                MsgBox("Customer Supplied Materail for Item " & arrItem(inti) & " is " & arrQty(inti) & " .", MsgBoxStyle.OkOnly, "eMPro")
                                Cmdinvoice.Focus()
                                ExploreBom = False
                                Exit Function
                            Else
                                '***********for Adding Values in CustAnnex Array
                                If Len(Trim(arrCustAnnex(0, pintAnnexMaxCount))) > 0 Then
                                    pintAnnexMaxCount = pintAnnexMaxCount + 1
                                    ReDim Preserve arrCustAnnex(3, pintAnnexMaxCount)
                                End If
                                arrCustAnnex(0, pintAnnexMaxCount) = rsCustAnnexDtl.GetValue("Item_code")
                                arrCustAnnex(1, pintAnnexMaxCount) = (dblTotalReqQty * pstrFinishedQty)
                                arrCustAnnex(2, pintAnnexMaxCount) = (rsCustAnnexDtl.GetValue("Balance_qty") - (dblTotalReqQty * pstrFinishedQty))
                                ExploreBom = True
                            End If
                        End If
                    End If
                Else
                    rsVandorBom = New ClsResultSetDB
                    rsVandorBom.GetResult("Select RawMaterial_Code from Vendor_bom where Finish_Product_code = '" & pstrItemCode & "'and RawMaterial_code = '" & strBomItem & "' and UNIT_CODE = '" & gstrUNITID & "'")
                    If rsVandorBom.GetNoRows > 0 Then
                        rsVandorBom.ResultSetClose()
                        MsgBox("Item " & strBomItem & " is not supplied.", MsgBoxStyle.OkOnly, "eMPro")
                        Cmdinvoice.Focus()
                        ExploreBom = False
                        rsCustAnnexDtl.ResultSetClose()
                        Exit Function
                    Else 'if not of Process type I then again Explore
                        pstrFinishedQty = pstrFinishedQty * dblTotalReqQty
                        Call ExploreBom(strBomItem, pstrFinishedQty, pstrSPCurrentRow, pstrCustCode, pstrRef, pintAnnexMaxCount, pstrFinishedProduct)
                    End If
                End If
                rsBomMstRaw.MoveNext()
            Next
        Else
            MsgBox("No BOM Defind for Item (" & strBomItem & ") defined in challan", MsgBoxStyle.Information, "eMPro")
            ExploreBom = False
            rsBomMstRaw.ResultSetClose()
            Exit Function
        End If
        rsBomMstRaw.ResultSetClose()
        rsCustAnnexDtl.ResultSetClose()
        rsVandorBom.ResultSetClose()
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
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
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
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
                chkLockPrintingFlag.Enabled = True
                chkLockPrintingFlag.Focus()
        End Select
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Public Function CheckDataFromGrin(ByRef pdblDocNo As Double, ByRef pstrCustCode As String) As Boolean
        Dim rsGrnDtl As ClsResultSetDB
        Dim rsSalesDtl As ClsResultSetDB
        Dim strSql As String
        Dim StrItemCode As String
        Dim dblItemQty As Double
        Dim dblRejQty As Double
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        On Error GoTo ErrHandler
        rsSalesDtl = New ClsResultSetDB
        rsSalesDtl.GetResult("Select Item_Code,Sales_Quantity from Sales_dtl where doc_No =" & Ctlinvoice.Text & " and Location_Code='" & Trim(txtUnitCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "'")
        intMaxLoop = rsSalesDtl.GetNoRows : rsSalesDtl.MoveFirst()
        CheckDataFromGrin = False
        For intLoopCounter = 1 To intMaxLoop
            StrItemCode = rsSalesDtl.GetValue("Item_code")
            dblItemQty = rsSalesDtl.GetValue("Sales_quantity")
            strSql = "select a.Doc_No,a.Item_code,a.Rejected_Quantity,"
            strSql = strSql & "Despatch_Quantity = isnull(a.Despatch_Quantity,0),"
            strSql = strSql & " Inspected_Quantity = isnull(Inspected_Quantity,0),"
            strSql = strSql & "RGP_Quantity = isnull(RGP_Quantity,0) from grn_Dtl a,grn_hdr b Where "
            strSql = strSql & "a.Doc_type = b.Doc_type And a.Doc_No = b.Doc_No AND a.UNIT_CODE=b.UNIT_CODE AND a.UNIT_CODE = '" & gstrUNITID & "' and "
            strSql = strSql & "a.From_Location = b.From_Location and a.From_Location ='01R1'"
            strSql = strSql & "and a.Rejected_quantity > 0 and b.Vendor_code = '" & pstrCustCode
            strSql = strSql & "' and a.Doc_No = " & pdblDocNo & " and a.Item_code = '" & StrItemCode & "'"
            rsGrnDtl = New ClsResultSetDB
            rsGrnDtl.GetResult(strSql)
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
            rsGrnDtl.ResultSetClose()
            rsSalesDtl.MoveNext()
        Next
        rsSalesDtl.ResultSetClose()
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Public Function UpdateGrnHdr(ByRef pdblGrinNo As Double, ByRef pdblinvoiceNo As Double) As Object
        Dim rsSalesDtl As ClsResultSetDB
        Dim intMaxLoop As Short
        Dim StrItemCode As String
        Dim dblqty As Double
        Dim intLoopCount As Short
        rsSalesDtl = New ClsResultSetDB
        rsSalesDtl.GetResult("select * from sales_dtl where Doc_No = " & Ctlinvoice.Text & " and Location_Code='" & Trim(txtUnitCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "'")
        If rsSalesDtl.GetNoRows > 0 Then
            intMaxLoop = rsSalesDtl.GetNoRows
            rsSalesDtl.MoveFirst()
            strupdateGrinhdr = ""
            For intLoopCount = 1 To intMaxLoop
                StrItemCode = rsSalesDtl.GetValue("ITem_code")
                dblqty = rsSalesDtl.GetValue("Sales_Quantity")
                If Len(Trim(strupdateGrinhdr)) = 0 Then
                    strupdateGrinhdr = "Update Grn_Dtl Set Despatch_Quantity = isnull(Despatch_Quantity,0) +" & dblqty
                    strupdateGrinhdr = strupdateGrinhdr & " Where ITem_Code = '" & StrItemCode & "' and UNIT_CODE = '" & gstrUNITID & "' and Doc_No = " & pdblGrinNo
                Else
                    strupdateGrinhdr = strupdateGrinhdr & vbCrLf & "Update Grn_Dtl Set Despatch_Quantity = isnull(Despatch_Quantity,0) + " & dblqty
                    strupdateGrinhdr = strupdateGrinhdr & " Where ITem_Code = '" & StrItemCode & "' and UNIT_CODE = '" & gstrUNITID & "' and Doc_No = " & pdblGrinNo
                End If
                rsSalesDtl.MoveNext()
            Next
        Else
            MsgBox("No Items Available in Invoice " & Ctlinvoice.Text)
        End If
        rsSalesDtl.ResultSetClose()
    End Function
    Private Function GetTaxGlSl(ByVal TaxType As String) As String
        Dim objRecordSet As New ADODB.Recordset
        On Error GoTo ErrHandler
        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close()
        objRecordSet.Open("SELECT tx_glCode, tx_slCode FROM fin_TaxGlRel where tx_rowType = 'ARTAX' and UNIT_CODE = '" & gstrUNITID & "' AND tx_taxId ='" & TaxType & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
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
    Private Function CreateStringForAccounts() As Boolean
        Dim objRecordSet As New ADODB.Recordset
        Dim objTmpRecordset As New ADODB.Recordset
        Dim strRetVal As String
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
        On Error GoTo ErrHandler
        objRecordSet.Open("SELECT * FROM  saleschallan_dtl WHERE Doc_No='" & Trim(Ctlinvoice.Text) & "' and UNIT_CODE = '" & gstrUNITID & "' and Location_Code='" & Trim(txtUnitCode.Text) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
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
        strCustCode = Trim(objRecordSet.Fields("Account_Code").Value)
        If UCase(lbldescription.Text) <> "SMP" Then 'if invoice type is not sample sales then
            'Retreiving the customer gl, sl and credit term id
            If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
            If UCase(Trim(cmbInvType.Text)) = "REJECTION" Then
                objTmpRecordset.Open("SELECT GL_AccountID, Ven_slCode, CrTrm_Termid FROM Pur_VendorMaster where Prty_PartyID='" & strCustCode & "' and UNIT_CODE = '" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
            Else
                objTmpRecordset.Open("SELECT Cst_ArCode, Cst_slCode, Cst_CreditTerm FROM Sal_CustomerMaster where Prty_PartyID='" & strCustCode & "' and UNIT_CODE = '" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
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
                strCreditTermsID = Trim(IIf(IsDBNull(objTmpRecordset.Fields("Cst_CreditTerm").Value), "", objTmpRecordset.Fields("Cst_CreditTerm").Value))
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
            Dim objCreditTerms As New prj_CreditTerm.clsCR_Term_Resolver
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
        mstrMasterString = "INV»" & strInvoiceNo & "»Dr»»" & strInvoiceDate & "»»»»»SAL»INV»" & strInvoiceNo & "»" & strInvoiceDate & "»" & Trim(strCustCode) & "»" & gstrUNITID & "»" & strCurrencyCode & "»" & System.Math.Round(dblInvoiceAmt, 0) & "»" & dblInvoiceAmt & "»" & dblExchangeRate & "»" & strCreditTermsID & "»" & strBasicDueDate & "»" & strPaymentDueDate & "»" & strExpectedDueDate & "»" & strCustomerGL & "»" & strCustomerSL & "»" & mP_User & "»getdate()»»"
        iCtr = 1
        'CST/LST Posting
        If objTmpRecordset.State = ADODB.ObjectStateEnum.adStateOpen Then objTmpRecordset.Close()
        If Trim(IIf(IsDBNull(objRecordSet.Fields("SalesTax_Type").Value), "", objRecordSet.Fields("SalesTax_Type").Value)) <> "" Then
            objTmpRecordset.Open("SELECT Tx_TaxeID FROM Gen_TaxRate WHERE UNIT_CODE = '" & gstrUNITID & "' and TxRt_Rate_No='" & Trim(IIf(IsDBNull(objRecordSet.Fields("SalesTax_Type").Value), "", objRecordSet.Fields("SalesTax_Type").Value)) & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
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
            If strTaxType = "LST" Or strTaxType = "CST" Then
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
                    mstrDetailString = mstrDetailString & "INV»" & strInvoiceNo & "»" & iCtr & "»TAX»" & strTaxType & "»0»" & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                    iCtr = iCtr + 1
                End If
            End If
        End If
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
            mstrDetailString = mstrDetailString & "INV»" & strInvoiceNo & "»" & iCtr & "»TAX»SST»0»" & Trim(objRecordSet.Fields("item_code").Value) & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
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
            mstrDetailString = mstrDetailString & "INV»" & strInvoiceNo & "»" & iCtr & "»TAX»INS»0»" & "»»0»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
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
            mstrDetailString = mstrDetailString & "INV»" & strInvoiceNo & "»" & iCtr & "»TAX»FRT»0»" & "»»0»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
            iCtr = iCtr + 1
        End If
        If objRecordSet.State = ADODB.ObjectStateEnum.adStateOpen Then objRecordSet.Close()
        objRecordSet.Open("SELECT sales_dtl.*, item_mst.GlGrp_code FROM sales_dtl, item_mst WHERE sales_dtl.Doc_No='" & Trim(Ctlinvoice.Text) & "' and sales_dtl.Item_Code=item_mst.Item_Code and sales_dtl.Location_Code='" & Trim(txtUnitCode.Text) & "' AND sales_dtl.UNIT_CODE=item_mst.UNIT_CODE AND sales_dtl.UNIT_CODE = '" & gstrUNITID & "'")
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
            'Basic Amount Posting
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
                objTmpRecordset.Open("SELECT * FROM invcc_dtl WHERE Invoice_Type='" & lbldescription.Text & "' AND Sub_Type = '" & lblcategory.Text & "' AND Location_Code ='" & Trim(txtUnitCode.Text) & "' AND ccM_cc_Percentage > 0 and UNIT_CODE = '" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockReadOnly, ADODB.CommandTypeEnum.adCmdText)
                If Not objTmpRecordset.EOF Then
                    While Not objTmpRecordset.EOF
                        dblCCShare = (dblBaseCurrencyAmount / 100) * objTmpRecordset.Fields("ccM_cc_Percentage").Value
                        mstrDetailString = mstrDetailString & "INV»" & strInvoiceNo & "»" & iCtr & "»ITM»SAL»" & iCtr & "»" & Trim(objRecordSet.Fields("item_code").Value) & "»" & strGlGroupId & "»0»" & strItemGL & "»" & strItemSL & "»" & dblCCShare & "»Cr»»" & Trim(objTmpRecordset.Fields("ccM_ccCode").Value) & "»»»»0»0»0»0»0¦"
                        objTmpRecordset.MoveNext()
                        iCtr = iCtr + 1
                    End While
                Else
                    mstrDetailString = mstrDetailString & "INV»" & strInvoiceNo & "»" & iCtr & "»ITM»SAL»" & iCtr & "»" & Trim(objRecordSet.Fields("item_code").Value) & "»" & strGlGroupId & "»0»" & strItemGL & "»" & strItemSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                    iCtr = iCtr + 1
                End If
                '*********************************************************
            End If
            'EXC Duty Posting
            dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("Excise_Tax").Value), 0, objRecordSet.Fields("Excise_Tax").Value)
            dblBaseCurrencyAmount = dblTaxAmt
            dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("Excise_per").Value), 0, objRecordSet.Fields("Excise_per").Value)
            If dblBaseCurrencyAmount > 0 Then
                'initializing the tax gl and sl here
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
                mstrDetailString = mstrDetailString & "INV»" & strInvoiceNo & "»" & iCtr & "»TAX»EXC»0»" & Trim(objRecordSet.Fields("item_code").Value) & "»»" & dblTaxRate & "»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
                iCtr = iCtr + 1
            End If
            'Others Posting
            dblTaxAmt = IIf(IsDBNull(objRecordSet.Fields("Others").Value), 0, objRecordSet.Fields("Others").Value)
            dblBaseCurrencyAmount = dblTaxAmt
            'initialize the tax gl and sl here
            If dblBaseCurrencyAmount > 0 Then
                mstrDetailString = mstrDetailString & "INV»" & strInvoiceNo & "»" & iCtr & "»TAX»OTH»0»" & Trim(objRecordSet.Fields("item_code").Value) & "»»0»" & strTaxGL & "»" & strTaxSL & "»" & dblBaseCurrencyAmount & "»Cr»»»»»»0»0»0»0»0" & "¦"
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
        dblBaseCurrencyAmount = dblInvoiceAmt - System.Math.Round(dblInvoiceAmt, 0)
        dblBaseCurrencyAmount = System.Math.Round(dblBaseCurrencyAmount, 4)
        If System.Math.Abs(dblBaseCurrencyAmount) > 0 Then
            mstrDetailString = mstrDetailString & "INV»" & strInvoiceNo & "»" & iCtr & "»»RND»0»" & "»»0»" & strItemGL & "»" & strItemSL & "»" & System.Math.Abs(dblBaseCurrencyAmount) & "»"
            If dblBaseCurrencyAmount < 0 Then
                mstrDetailString = mstrDetailString & "Cr»»»»»»0»0»0»0»0" & "¦"
            Else
                mstrDetailString = mstrDetailString & "Dr»»»»»»0»0»0»0»0" & "¦"
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
    Public Sub UpdateinSale_Dtl()
        Dim rssaledtl As ClsResultSetDB
        Dim rsSaleConf As ClsResultSetDB
        Dim strSql As String
        Dim strStockLocCode As String
        Dim strInvoiceDate As String
        Dim intRow, intLoopCount As Short
        Dim mItem_Code, mCust_Item_Code As String
        Dim mSales_Quantity As Double
        Dim rsSalesChallan As ClsResultSetDB
        strupdateitbalmst = ""
        strupdatecustodtdtl = ""
        On Error GoTo Err_Handler
        strSql = "select * from Saleschallan_Dtl where Doc_No =" & Me.Ctlinvoice.Text & "  and Location_Code='" & Trim(txtUnitCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "'"
        rsSalesChallan = New ClsResultSetDB
        rsSalesChallan.GetResult(strSql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        strInvoiceDate = VB6.Format(rsSalesChallan.GetValue("Invoice_Date"), gstrDateFormat)
        rsSaleConf = New ClsResultSetDB
        rsSaleConf.GetResult("Select Stock_Location from saleconf where Description = '" & Me.cmbInvType.Text & "' and Sub_Type_Description ='" & Me.CmbCategory.Text & "' and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "' ,fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0 and UNIT_CODE = '" & gstrUNITID & "'")
        strStockLocCode = rsSaleConf.GetValue("Stock_Location")
        strSql = "Select * from sales_Dtl where Doc_No = " & Me.Ctlinvoice.Text & " and Location_Code='" & Trim(txtUnitCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "'"
        rssaledtl = New ClsResultSetDB
        rssaledtl.GetResult(strSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rssaledtl.GetNoRows > 0 Then
            intRow = rssaledtl.GetNoRows
            rssaledtl.MoveFirst()
            For intLoopCount = 1 To intRow
                If Not rssaledtl.EOFRecord Then
                    mItem_Code = rssaledtl.GetValue("Item_Code")
                    mCust_Item_Code = rssaledtl.GetValue("Cust_Item_Code")
                    mSales_Quantity = IIf(rssaledtl.GetValue("Sales_Quantity") = "", 0, rssaledtl.GetValue("Sales_Quantity"))
                    strupdateitbalmst = Trim(strupdateitbalmst) & "Update Itembal_mst set cur_bal= cur_bal-"
                    strupdateitbalmst = strupdateitbalmst & mSales_Quantity & " where Location_code = '" & strStockLocation
                    strupdateitbalmst = strupdateitbalmst & "' and UNIT_CODE = '" & gstrUNITID & "' and item_code = '" & mItem_Code & "' "
                    strupdatecustodtdtl = Trim(strupdatecustodtdtl) & "Update Cust_ord_dtl set Despatch_Qty = Despatch_Qty + "
                    strupdatecustodtdtl = strupdatecustodtdtl & mSales_Quantity & " where Account_code ='"
                    strupdatecustodtdtl = strupdatecustodtdtl & mAccount_Code & "' and UNIT_CODE = '" & gstrUNITID & "' and Cust_DrgNo = '"
                    strupdatecustodtdtl = strupdatecustodtdtl & mCust_Item_Code & "' and Cust_ref = '" & mCust_Ref
                    strupdatecustodtdtl = strupdatecustodtdtl & "'and amendment_no = '" & mAmendment_No & "' and active_Flag ='A' "
                    rssaledtl.MoveNext()
                End If
            Next
        End If
        rssaledtl.ResultSetClose()
        rsSaleConf.ResultSetClose()
        rsSalesChallan.ResultSetClose()
        rsSaleConf = Nothing
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
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
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Function InvoiceCustSuppBOM(ByRef dblInvoiceNo As Double) As Boolean
        Dim rsSalesDtl As ClsResultSetDB
        Dim rsEnggBom As ClsResultSetDB
        Dim rsVendorBom As ClsResultSetDB
        Dim rsSalesChallan As ClsResultSetDB
        Dim strCustomer As String
        Dim strMsgVendITem As String
        Dim strMsgEnggITem As String
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        rsSalesChallan = New ClsResultSetDB
        rsSalesChallan.GetResult("Select Account_code from SalesChallan_dtl Where  UNIT_CODE = '" & gstrUNITID & "' and Doc_No =" & dblInvoiceNo)
        strCustomer = rsSalesChallan.GetValue("Account_code")
        rsSalesChallan.ResultSetClose()
        rsSalesDtl = New ClsResultSetDB
        rsSalesDtl.GetResult("Select * from Sales_dtl Where  UNIT_CODE = '" & gstrUNITID & "' and Doc_No =" & dblInvoiceNo)
        intMaxLoop = rsSalesDtl.GetNoRows
        rsSalesDtl.MoveFirst()
        'To Check Vendor Bom in For ITemMst
        strMsgVendITem = ""
        For intLoopCounter = 1 To intMaxLoop
            rsVendorBom = New ClsResultSetDB
            rsVendorBom.GetResult("Select * from Vendor_Bom where Vendor_Code = '" & strCustomer & "' and UNIT_CODE = '" & gstrUNITID & "' and Finish_Product_code = '" & rsSalesDtl.GetValue("Item_code") & "'")
            If rsVendorBom.GetNoRows = 0 Then
                If Len(Trim(strMsgVendITem)) = 0 Then
                    strMsgVendITem = " Following Item(s) Customer BOM is Not Defined :" & vbCrLf & " " & rsSalesDtl.GetValue("Item_code")
                Else
                    strMsgVendITem = strMsgVendITem & vbCrLf & " " & rsSalesDtl.GetValue("Item_code")
                End If
            End If
            rsVendorBom.ResultSetClose()
            rsSalesDtl.MoveNext()
        Next
        rsSalesDtl.MoveFirst()
        strMsgEnggITem = ""
        For intLoopCounter = 1 To intMaxLoop
            rsEnggBom = New ClsResultSetDB
            rsEnggBom.GetResult("Select * from Bom_Mst where  UNIT_CODE = '" & gstrUNITID & "' and Finished_Product_code = '" & rsSalesDtl.GetValue("Item_code") & "'")
            If rsEnggBom.GetNoRows = 0 Then
                If Len(Trim(strMsgVendITem)) = 0 Then
                    strMsgEnggITem = " Following Item(s) Engg BOM is Not Defined :" & vbCrLf & " " & rsSalesDtl.GetValue("Item_code")
                Else
                    strMsgEnggITem = strMsgEnggITem & vbCrLf & " " & rsSalesDtl.GetValue("Item_code")
                End If
            End If
            rsEnggBom.ResultSetClose()
            rsSalesDtl.MoveNext()
        Next
        If (Len(Trim(strMsgVendITem)) > 0) Or (Len(Trim(strMsgVendITem)) > 0) Then
            MsgBox(strMsgVendITem & vbCrLf & strMsgEnggITem, MsgBoxStyle.Information, "eMPro")
            InvoiceCustSuppBOM = False
        Else
            InvoiceCustSuppBOM = True
        End If
        rsSalesDtl.ResultSetClose()
    End Function
    Public Function ToCheckDrgForCustSupp(ByRef dblInvoiceNo As Double) As Boolean
        Dim rsSalesDtl As ClsResultSetDB
        Dim rsCustItem As ClsResultSetDB
        Dim rsVendorBom As ClsResultSetDB
        Dim rsSalesChallan As ClsResultSetDB
        Dim strCustomer As String
        Dim strMsgCustITem As String
        Dim strMsgEnggITem As String
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        rsSalesChallan = New ClsResultSetDB
        rsSalesChallan.GetResult("Select Account_code from SalesChallan_dtl Where  UNIT_CODE = '" & gstrUNITID & "' and Doc_No =" & dblInvoiceNo)
        strCustomer = rsSalesChallan.GetValue("Account_code")
        rsSalesChallan.ResultSetClose()
        rsSalesDtl = New ClsResultSetDB
        rsSalesDtl.GetResult("Select * from Sales_dtl Where  UNIT_CODE = '" & gstrUNITID & "' and Doc_No =" & dblInvoiceNo)
        intMaxLoop = rsSalesDtl.GetNoRows
        rsSalesDtl.MoveFirst()
        strMsgCustITem = ""
        Dim intloopcounter1 As Short
        Dim intMaxLoop1 As Short
        For intLoopCounter = 1 To intMaxLoop
            rsVendorBom = New ClsResultSetDB
            rsVendorBom.GetResult("Select RawMaterial_Code,Item_code from Vendor_Bom where Vendor_Code = '" & strCustomer & "' and UNIT_CODE = '" & gstrUNITID & "' and Finish_Product_code = '" & rsSalesDtl.GetValue("Item_code") & "'")
            intMaxLoop1 = rsVendorBom.GetNoRows
            rsVendorBom.MoveFirst()
            For intloopcounter1 = 1 To intMaxLoop1
                rsCustItem = New ClsResultSetDB
                rsCustItem.GetResult(" Select * from CustItem_Mst where Account_code = '" & strCustomer & "' and UNIT_CODE = '" & gstrUNITID & "' and Item_code = '" & rsVendorBom.GetValue("RawMaterial_Code") & "'")
                If rsCustItem.GetNoRows = 0 Then
                    If Len(Trim(strMsgCustITem)) > 0 Then
                        strMsgCustITem = strMsgCustITem & vbCrLf & " " & rsVendorBom.GetValue("RawMaterial_Code")
                    Else
                        strMsgCustITem = "Following Item(s) Customer Part Code is not Defined :" & vbCrLf & " " & rsVendorBom.GetValue("RawMaterial_Code")
                    End If
                End If
                rsCustItem.ResultSetClose()
                rsVendorBom.MoveNext()
            Next
            rsVendorBom.ResultSetClose()
            rsSalesDtl.MoveNext()
        Next
        If Len(Trim(strMsgCustITem)) > 0 Then
            MsgBox(strMsgCustITem, MsgBoxStyle.Information, "eMPro")
            ToCheckDrgForCustSupp = False
        Else
            ToCheckDrgForCustSupp = True
        End If
        rsSalesDtl.ResultSetClose()
    End Function
    Public Function ToCheckItemRateForCustMtrl(ByRef dblInvoiceNo As Double) As Boolean
        Dim rsSalesDtl As ClsResultSetDB
        Dim rsCustItem As ClsResultSetDB
        Dim rsItemRate As ClsResultSetDB
        Dim rsVendorBom As ClsResultSetDB
        Dim rsSalesChallan As ClsResultSetDB
        Dim rsEnggBom As ClsResultSetDB
        Dim strCustomer As String
        Dim strMsgITemRate As String
        Dim strMsgEnggITem As String
        Dim blnCustSuppMtrl As Boolean
        Dim dblCustSupp As Double
        Dim dblAllCustSupp As Double
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        rsSalesChallan = New ClsResultSetDB
        rsSalesChallan.GetResult("Select Account_code from SalesChallan_dtl Where  UNIT_CODE = '" & gstrUNITID & "' and Doc_No =" & dblInvoiceNo)
        strCustomer = rsSalesChallan.GetValue("Account_code")
        rsSalesChallan.ResultSetClose()
        rsSalesDtl = New ClsResultSetDB
        rsSalesDtl.GetResult("Select * from Sales_dtl Where   UNIT_CODE = '" & gstrUNITID & "' and Doc_No =" & dblInvoiceNo)
        intMaxLoop = rsSalesDtl.GetNoRows
        rsSalesDtl.MoveFirst()
        blnCustSuppMtrl = True
        strMsgITemRate = ""
        Dim intloopcounter1 As Short
        Dim intMaxLoop1 As Short
        For intLoopCounter = 1 To intMaxLoop
            dblCustSupp = rsSalesDtl.GetValue("Cust_Mtrl")
            dblAllCustSupp = 0
            rsVendorBom = New ClsResultSetDB
            rsVendorBom.GetResult("Select RawMaterial_Code,Item_code from Vendor_Bom where Vendor_Code = '" & strCustomer & "' and Finish_Product_code = '" & rsSalesDtl.GetValue("Item_code") & "' and UNIT_CODE = '" & gstrUNITID & "'")
            intMaxLoop1 = rsVendorBom.GetNoRows
            rsVendorBom.MoveFirst()
            For intloopcounter1 = 1 To intMaxLoop1
                rsCustItem = New ClsResultSetDB
                rsCustItem.GetResult(" Select Cust_DrgNo from CustItem_Mst where Account_code = '" & strCustomer & "' and Item_code = '" & rsVendorBom.GetValue("RawMaterial_Code") & "' and UNIT_CODE = '" & gstrUNITID & "'")
                If rsCustItem.GetNoRows > 0 Then
                    rsCustItem.MoveFirst()
                    rsItemRate = New ClsResultSetDB
                    rsItemRate.GetResult("SELECT Cust_Supplied_Material = isnull(Cust_Supplied_Material,0) FROM ItemRate_mst WHERE Custvend_flg = 'C' AND Party_code = '" & strCustomer & "' AND Item_code = '" & rsCustItem.GetValue("Cust_drgNo") & "' and UNIT_CODE = '" & gstrUNITID & "' AND serial_no = (SELECT MAX(Serial_No) FROM ItemRate_mst WHERE Custvend_flg = 'C' and UNIT_CODE = '" & gstrUNITID & "' AND Party_code = '" & strCustomer & "' AND Item_code = '" & rsCustItem.GetValue("Cust_drgNo") & "')")
                    '*******To Get Required Qty From BOM for Cust Supplied Material
                    rsEnggBom = New ClsResultSetDB
                    rsEnggBom.GetResult(" Select Req_Qty = isnull(Required_Qty,0) + isnull(Waste_Qty,0) from Bom_Mst where RawMaterial_Code = '" & rsVendorBom.GetValue("RawMaterial_Code") & "' and Finished_Product_code = '" & rsVendorBom.GetValue("Item_Code") & "' and UNIT_CODE = '" & gstrUNITID & "'")
                    If rsItemRate.GetNoRows > 0 Then
                        rsItemRate.MoveFirst()
                        dblAllCustSupp = dblAllCustSupp + rsItemRate.GetValue("Cust_Supplied_Material") * rsEnggBom.GetValue("Req_qty")
                    Else
                        If Len(Trim(strMsgITemRate)) > 0 Then
                            strMsgITemRate = strMsgITemRate & vbCrLf & " " & rsCustItem.GetValue("Cust_drgNo")
                        Else
                            strMsgITemRate = "Following Customer Part Code(s) are not Defined in Item Rate Master :" & vbCrLf & " " & rsCustItem.GetValue("Cust_drgNo")
                        End If
                    End If
                    rsEnggBom.ResultSetClose()
                    rsItemRate.ResultSetClose()
                End If
                rsCustItem.ResultSetClose()
                rsVendorBom.MoveNext()
            Next
            If blnCustSuppMtrl = False Then
                blnCustSuppMtrl = False
            Else
                If Val(CStr(dblAllCustSupp)) <> Val(CStr(dblCustSupp)) Then
                    blnCustSuppMtrl = False
                Else
                    blnCustSuppMtrl = True
                End If
            End If
            If Len(Trim(strMsgITemRate)) = 0 Then
                If blnCustSuppMtrl = False Then
                    MsgBox("Sum of Customer Supplied Material in Item Rate Master (Acc. To Engg. BOM calculations) is [ " & dblAllCustSupp & " ] and Finished Item (" & rsSalesDtl.GetValue("Item_code") & ") entered in SO is [ " & dblCustSupp & " ] is Not Same,can Not Print.", MsgBoxStyle.Information, "eMPro")
                    ToCheckItemRateForCustMtrl = False
                    rsVendorBom.ResultSetClose()
                    Exit Function
                End If
            End If
            rsVendorBom.ResultSetClose()
            rsSalesDtl.MoveNext()
        Next
        If Len(Trim(strMsgITemRate)) > 0 Then
            MsgBox(strMsgITemRate, MsgBoxStyle.Information, "eMPro")
            ToCheckItemRateForCustMtrl = False
        Else
            If blnCustSuppMtrl = False Then
                MsgBox("sum of Customer Supplied Material of (BOP's and Sub Assamblies) and Finished Item is Not Same can Not Print.", MsgBoxStyle.Information, "eMPro")
                ToCheckItemRateForCustMtrl = False
            Else
                ToCheckItemRateForCustMtrl = True
            End If
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
        Dim strSql As String 'String SQL Query
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        pstrRequiredDate = getDateForDB(pstrRequiredDate)
        If Len(Trim(pstrInvoiceType)) > 0 Then 'For Dated Docs
            strSql = "Select Current_No,Suffix,Fin_start_date,Fin_end_Date From saleConf Where "
            strSql = strSql & "Invoice_Type ='" & pstrInvoiceType & "' and  sub_type='" & pstrInvoiceSubType & "' AND Location_Code ='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & pstrRequiredDate & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & pstrRequiredDate & "')<=0 and UNIT_CODE = '" & gstrUNITID & "'"
            With clsInstEMPDBDbase.CConnection
                .OpenConnection(gstrDSNName, gstrDatabaseName)
                .ExecuteSQL("Set Dateformat 'dmy'")
            End With
            clsInstEMPDBDbase.CRecordset.OpenRecordset(strSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
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
        clsErrorInst.CError.RaiseError(20008, "[frmexptrn0006]", "[GenerateInvoiceNo]", "", "No. Not Generated For DocType = " & pstrInvoiceType & " due to [ " & Err.Description & " ].", My.Application.Info.DirectoryPath, gstrDSNName, gstrDatabaseName)
    End Function
    Private Sub Cmdinvoice_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.UCfraRepCmd.ButtonClickEventArgs) Handles Cmdinvoice.ButtonClick
        Dim rsSalesConf As ClsResultSetDB
        Dim rssaledtl As ClsResultSetDB
        Dim rsItembal As ClsResultSetDB
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
        Dim strInvoiceDate As String
        Dim intNoCopies As Short
        Dim strRetVal As String
        Dim objDrCr As prj_DrCrNote.cls_DrCrNote
        On Error GoTo Err_Handler
        rssaledtl = New ClsResultSetDB
        Dim objRpt As ReportDocument
        Dim frmReportViewer As New eMProCrystalReportViewer
        objRpt = frmReportViewer.GetReportDocument()
        objRpt = frmReportViewer.GetReportDocument()
        frmReportViewer.ReportHeader = Me.ctlFormHeader1.HeaderString()
        If e.Button = UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE Then
            Me.Dispose()
            Exit Sub
        Else
            If ValidSelection() = False Then Exit Sub
        End If
        SALEDTL = "select * from Saleschallan_Dtl where Doc_No =" & Me.Ctlinvoice.Text & "  and Location_Code='" & Trim(txtUnitCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "'"
        rssaledtl = New ClsResultSetDB
        rssaledtl.GetResult(SALEDTL, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        strAccountCode = rssaledtl.GetValue("Account_code")
        strCustRef = rssaledtl.GetValue("Cust_ref")
        StrAmendmentNo = rssaledtl.GetValue("Amendment_No")
        strInvoiceDate = VB6.Format(rssaledtl.GetValue("Invoice_Date"), gstrDateFormat)
        rssaledtl.ResultSetClose()
        strSalesconf = "Select UpdatePO_Flag,UpdateStock_Flag,Stock_Location,OpenningBal,Preprinted_Flag,NoCopies from saleconf where "
        strSalesconf = strSalesconf & "Invoice_type = '" & Me.lbldescription.Text & "' and sub_type = '"
        strSalesconf = strSalesconf & Me.lblcategory.Text & "' and Location_Code='" & Trim(txtUnitCode.Text) & "' and datediff(dd,'" & getDateForDB(strInvoiceDate) & "',fin_start_date)<=0  and datediff(dd,fin_end_date,'" & getDateForDB(strInvoiceDate) & "')<=0 and UNIT_CODE = '" & gstrUNITID & "'"
        rsSalesConf = New ClsResultSetDB
        rsSalesConf.GetResult(strSalesconf, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        updatePOflag = rsSalesConf.GetValue("UpdatePO_Flag")
        updatestockflag = rsSalesConf.GetValue("UpdateStock_Flag")
        strStockLocation = rsSalesConf.GetValue("Stock_Location")
        mOpeeningBalance = Val(rsSalesConf.GetValue("OpenningBal"))
        intNoCopies = Val(rsSalesConf.GetValue("NoCopies"))
        If Len(Trim(strStockLocation)) = 0 Then
            MsgBox("Please Define Stock Location in Sales Configuration. ")
            Exit Sub
        End If
        SALEDTL = "Select Sales_Quantity,Item_code,Cust_Item_Code from sales_Dtl where Doc_No = " & Me.Ctlinvoice.Text & " and Location_Code='" & Trim(txtUnitCode.Text) & "' and UNIT_CODE = '" & gstrUNITID & "'"
        rssaledtl = New ClsResultSetDB
        rssaledtl.GetResult(SALEDTL, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        intRow = rssaledtl.GetNoRows
        rssaledtl.MoveFirst()
        If optInvYes(0).Checked = True Then
            '******Check for balance & despatch in Cust_ord_dtl
            For intLoopCount = 1 To intRow
                ItemCode = rssaledtl.GetValue("Item_code")
                salesQuantity = rssaledtl.GetValue("Sales_quantity")
                strDrgNo = rssaledtl.GetValue("Cust_Item_code")
                rsItembal = New ClsResultSetDB
                rsItembal.GetResult("Select Cur_bal from Itembal_Mst where Item_code = '" & ItemCode & "'and Location_code ='" & strStockLocation & "' and UNIT_CODE = '" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                If rsItembal.GetNoRows > 0 Then
                    If salesQuantity > rsItembal.GetValue("Cur_Bal") Then
                        MsgBox("Balance for item " & ItemCode & " at Location " & strStockLocation & " not available. ", MsgBoxStyle.Information, "eMPro")
                        rsItembal.ResultSetClose()
                        Exit Sub
                    End If
                Else
                    MsgBox("No Balance for item " & ItemCode & " at Location " & strStockLocation & ".", MsgBoxStyle.OkOnly, "eMPro")
                    rsSalesConf.ResultSetClose()
                    rsItembal.ResultSetClose()
                    Exit Sub
                End If
                rsItembal.ResultSetClose()
                If Len(Trim(strCustRef)) > 0 Then
                    If UCase(cmbInvType.Text) <> "REJECTION" Then
                        rsItembal = New ClsResultSetDB
                        rsItembal.GetResult("Select balanceQty = order_qty - despatch_Qty,OpenSO from Cust_ord_dtl where account_code ='" & strAccountCode & "' and Cust_ref ='" & strCustRef & "' and Amendment_No = '" & StrAmendmentNo & "' and Item_code ='" & ItemCode & "' and Cust_drgNo ='" & strDrgNo & "' and Active_flag ='A' and UNIT_CODE = '" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                        If rsItembal.GetNoRows > 0 Then
                            If rsItembal.GetValue("OpenSO") = False Then
                                If salesQuantity > rsItembal.GetValue("BalanceQty") Then
                                    MsgBox("Balance Quantity in SO for item " & ItemCode & " is " & rsItembal.GetValue("BalanceQty") & ".Check Quantity of Item in Challan.", MsgBoxStyle.Information, "eMPro")
                                    rsItembal.ResultSetClose()
                                    Exit Sub
                                End If
                            End If
                        Else
                            MsgBox("No Item (" & StrItemCode & ") exist in SO - " & strCustRef & ".", MsgBoxStyle.Information, "eMPro")
                            rsItembal.ResultSetClose()
                            Exit Sub
                        End If
                        rsItembal.ResultSetClose()
                    End If
                End If
                rssaledtl.MoveNext()
            Next
            '****To Check in Rejection Invoice if Grin No Exist
            If UCase(cmbInvType.Text) = "REJECTION" Then
                If Len(Trim(strCustRef)) > 0 Then
                    If CheckDataFromGrin(CDbl(Trim(strCustRef)), strAccountCode) = False Then
                        Exit Sub
                    End If
                End If
            End If
        End If
        rssaledtl.ResultSetClose()
        rssaledtl = Nothing
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        Select Case e.Button
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                Me.Dispose()
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT_TO_WINDOW
                If InvoiceGeneration(objRpt, frmReportViewer) = True Then
                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
                    frmReportViewer.Show()
                    System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                Else
                    Exit Sub
                End If
                If chkLockPrintingFlag.CheckState = 1 And optInvYes(0).Checked = True Then
                    Sleep((5000))
                    If ConfirmWindow(10344, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                        mP_Connection.BeginTrans()
                        mP_Connection.Execute("set Dateformat 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        mP_Connection.Execute(salesconf, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        mP_Connection.Execute(saleschallan, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        mP_Connection.Execute(strUPdateSaleDtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        If updatePOflag = True Then
                            mP_Connection.Execute(strupdatecustodtdtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        End If
                        If updatestockflag = True Then
                            mP_Connection.Execute(strupdateitbalmst, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
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
                        If mblnpostinfin = True Then
                            'Accounts Posting is done here
                            objDrCr = New prj_DrCrNote.cls_DrCrNote(GetServerDate)
                            strRetVal = objDrCr.SetARInvoiceDocument(gstrUNITID, mstrMasterString, mstrDetailString, prj_GLTransactions.cls_GLTransactions.udtOperationType.optInsert, gstrCONNECTIONSTRING)
                            strRetVal = CheckString(strRetVal)
                            objDrCr = Nothing
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
                            Ctlinvoice.Text = ""
                        End If
                    End If
                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT_TO_PRINTER
                If InvoiceGeneration(objRpt, frmReportViewer) = True Then
                    '***TO PRINT userdefined no of copies stored in SalesConf on 10/07/02
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
                                    objRpt.DataDefinition.FormulaFields("CopyName").Text = "'ORIGINAL FOR BUYER'"
                                Else
                                    objRpt.DataDefinition.FormulaFields("CopyName").Text = "'DUPLICATE FOR TRANSPORTER'"
                                End If
                            Case 2
                                If optInvYes(0).Checked = True Then
                                    objRpt.DataDefinition.FormulaFields("CopyName").Text = "'DUPLICATE FOR TRANSPORTER'"
                                Else
                                    objRpt.DataDefinition.FormulaFields("CopyName").Text = "'TRIPLICATE FOR ASSESSEE'"
                                End If
                            Case 3
                                If optInvYes(0).Checked = True Then
                                    objRpt.DataDefinition.FormulaFields("CopyName").Text = "'TRIPLICATE FOR ASSESSEE'"
                                Else
                                    objRpt.DataDefinition.FormulaFields("CopyName").Text = "'EXTRA COPY'"
                                End If
                            Case Is >= 4
                                objRpt.DataDefinition.FormulaFields("CopyName").Text = "'EXTRA COPY'"
                        End Select
                        objRpt.DataDefinition.FormulaFields("InsuranceFlag").Text = "'" & mblnInsuranceFlag & "'"
                        frmReportViewer.Show()
                    Next
                    If chkLockPrintingFlag.CheckState = 0 Then
                        If optInvYes(0).Checked = True Then
                            If ConfirmWindow(10344, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                                mP_Connection.BeginTrans()
                                mP_Connection.Execute("set Dateformat 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                mP_Connection.Execute(strsaledetails, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                mP_Connection.Execute(saleschallan, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                mP_Connection.Execute(salesconf, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                If updatePOflag = True Then
                                    mP_Connection.Execute(strupdatecustodtdtl, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                End If
                                If updatestockflag = True Then
                                    mP_Connection.Execute(strupdateitbalmst, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
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
                                'Accounts Posting is done here
                                objDrCr = New prj_DrCrNote.cls_DrCrNote(GetServerDate)
                                strRetVal = objDrCr.SetARInvoiceDocument(gstrUNITID, mstrMasterString, mstrDetailString, prj_GLTransactions.cls_GLTransactions.udtOperationType.optInsert, gstrCONNECTIONSTRING)
                                strRetVal = CheckString(strRetVal)
                                objDrCr = Nothing
                                If Not strRetVal = "Y" Then
                                    MsgBox(strRetVal, MsgBoxStyle.Information, "eMPro")
                                    mP_Connection.RollbackTrans()
                                    Exit Sub
                                Else
                                    mP_Connection.CommitTrans()
                                    MsgBox("Invoice has been locked successfully with number " & mInvNo, MsgBoxStyle.Information, "eMPro")
                                    Ctlinvoice.Text = ""
                                End If
                            End If
                        End If
                        Call RefreshForm()
                    End If
                    frmReportViewer.Show()
                Else
                    Exit Sub
                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT_TO_FILE
                'Reset the mouse pointer
                If InvoiceGeneration(objRpt, frmReportViewer) = True Then
                    frmExport.ShowDialog()
                    Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
                    If gblnCancelExport Then Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default) : Exit Sub
                    frmReportViewer.Show()
                    Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
                Else
                    Exit Sub
                End If
        End Select
        rsItembal.ResultSetClose()
        rsSalesConf.ResultSetClose()
        Exit Sub
Err_Handler:
        If Err.Number = 20545 Then
            Resume Next
        Else
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub
    Private Sub ctlFormHeader1_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlFormHeader1.Click
        Call ShowHelp("HLPMKTTRN0008.htm")
    End Sub
End Class