
Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports CrystalDecisions.CrystalReports.Engine
Imports System.Data.SqlClient
Imports System.IO
Friend Class frmMKTTRN0002_SOUTH
	Inherits System.Windows.Forms.Form
	'---------------------------------------------------------------------------
	'(C) 2001 MIND, All rights reserved
	'
	'File Name          :   frmMKTTRN0002.frm
	'Function           :   Customer PO Authorization
	'Created By         :   Meenu Gupta
	'Created on         :   28,Mar 2001
	'Revision History   :
	'Revised date       : 24 Aug 2001
	'Revised By         : Nisha Rai
	'17-01-2002 internal issue log no 55-checked out from no =4008
	'25-01-2002 done to allow 4 decimal places in Rate for MSSL-ED - checked out form no =4017
	'28-02-2002 Changed lable Surcharge % on formNo 4052 in grid control
	'13/09/2002 changed by nisha for accounts Plugin
	'changed by nisha on 01/07/2003
	'=======================================================================================
	'Revised By     : Ashutosh Verma
	'Revised On     : 22-01-2007 ,Issue ID:19352
	'Revised Reason : Show Additional Excise Duty in grid.
	'=======================================================================================
	'Revised By     : Ashutosh Verma
	'Revised On     : 16 Apr 2007 ,Issue ID:19731
	'Revised Reason : Add Consignee Code .
    '=======================================================================================
    'Revised By        -    Vinod Singh
    'Revision Date     -    04/05/2011
    'Revision History  -    Changes for Multi Unit
    '=======================================================================================
    'Revised By        -    Vinod Singh
    'Revision Date     -    29/06/2012
    'Revision History  -    Authorized SO was appearing in SO Help
    'Issue Id          -    10243325 
    '--------------------------------------------------------------------------------------
    'REVISED BY        -    PRASHANT RAJPAL
    'REVISION DATE     -    18-APR-2013
    'REVISION HISTORY  -    NEW FUNCTIONALITY -ONE SALES ORDER ACTIVE FOR ONE CUSTOMER ,CONFIGURABLE FUNCTIONALITY
    'ISSUE ID          -    10375521  
    '----------------------------------------------------------------------------------------------
    'REVISED BY        -    Parveen Kumar
    'REVISION DATE     -    23-Dec-2014
    'REVISION HISTORY  -    Sale order Authorization Form - Cess on ED & Surch. OnRejection Report.
    'ISSUE ID          -    10552513  
    '***************************************************************************************
    'REVISED BY        -   Abhinav Kumar
    'REVISION DATE     -    09-JAN-2015
    'REVISION HISTORY  -    CHANGES DONE FOR CT2 ARE 3 FUNCTIONALITY - TO ENABLE CT2 REQD FLAG
    'ISSUE ID          -    10736222 
    '***************************************************************************************
    'REVISED BY         :   Abhinav Kumar
    'ISSUE ID           :   10869290 
    'DESCRIPTION        :   eMPro- Service Invoice Functionality
    'REVISION DATE      :   21 Sep 2015
    '***************************************************************************************

    'REVISED BY         :   SUMIT KUMAR
    'ISSUE ID           :   101377199
    'DESCRIPTION        :   Preview PO Report AND ALART THE EMAIL  
    'REVISION DATE      :   10-AUG-2018
    '***************************************************************************************

    Dim rsdb As New ClsResultSetDB
	Dim m_CloseButton As Boolean
	Dim m_blnhelp As Boolean
	Dim mintFormIndex, intRow As Short
	Dim m_strSql, strSQL As String
    Dim rsRefNo As ClsResultSetDB
	Dim m_ItemDesc, m_custItemDesc As String
	Dim strpotype As String
	Dim blnValidAmend As Boolean
	Dim blnValidCust As Boolean
	Dim blnValidref As Boolean

 
	Private Sub cmdchangetype_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdchangetype.Click
        m_pstrSql = "select Account_Code,Cust_Ref,Amendment_No,Order_Date,Amendment_Date,Active_Flag,Currency_Code,Valid_Date,Effect_Date,Term_Payment,Special_Remarks,Pay_Remarks,Price_Remarks,Packing_Remarks,Frieght_Remarks,Transport_Remarks,Octorai_Remarks,Mode_Despatch,Delivery,First_Authorized,Second_Authorized,Third_Authorized,Authorized_Flag,Reason,PO_Type,SalesTax_Type,OpenSO,AddCustSupp,PerValue,InternalSONo,RevisionNo,Surcharge_code,Future_SO,ECESS_Code,Consignee_Code,ADDVAT_TYPE from cust_ord_hdr where unit_code='" & gstrUNITID & "' and Cust_ref='" & txtReferenceNo.Text & "' and amendment_No='" & txtAmendmentNo.Text & "' AND ACCOUNT_CODE='" & txtCustomerCode.Text & "' and Consignee_Code='" & Trim(txtConsCode.Text) & "' "
		frmMKTTRNAdditionalDetails.ShowDialog()
	End Sub
	Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
		Dim Index As Short = cmdHelp.GetIndex(eventSender)
		On Error GoTo ErrHandler
		Dim varRetVal As Object
		Select Case Index
			Case 0
				With Me.txtCustomerCode
                    If Len(.Text) = 0 Then
                        'Samiksha SMRC start
                        If gstrUNITID = "STH" Then
                            varRetVal = ShowList(1, .MaxLength, "", "A.Customer_code", "B.Cust_Name", "CUSTOMER_CONSIGNEE_MAPPING A,Customer_mst B", " and A.CUSTOMER_CODE = B.Customer_Code AND A.UNIT_CODE=B.UNIT_CODE and ((isnull(B.deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= B.deactive_date))", , , , , , "A.UNIT_CODE")

                        Else
                            varRetVal = ShowList(1, .MaxLength, "", "Customer_Code", "Cust_Name", "Customer_mst", " and ((isnull(deactive_flag,0) <> 1) OR (convert(varchar(12),getdate(),106)<= convert(varchar(12),deactive_date,106)))")

                        End If

                        'Samiksha SMRC end

                        If varRetVal = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            'MsgBox ("There Are No Existing Records")
                        Else
                            .Text = varRetVal
                        End If
                    Else
                        'Samiksha SMRC start

                        If gstrUNITID = "STH" Then
                            varRetVal = ShowList(1, .MaxLength, .Text, "A.Customer_code", "B.Cust_Name", "CUSTOMER_CONSIGNEE_MAPPING A,Customer_mst B", " and A.CUSTOMER_CODE = B.Customer_Code AND A.UNIT_CODE=B.UNIT_CODE and ((isnull(B.deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= B.deactive_date))", , , , , , "A.UNIT_CODE")
                        Else
                            varRetVal = ShowList(1, .MaxLength, .Text, "Customer_Code", "Cust_Name", "Customer_mst", " and ((isnull(deactive_flag,0) <> 1) OR (convert(varchar(12),getdate(),106)<= convert(varchar(12),deactive_date,106)))")
                        End If

                        'Samiksha SMRC end


                        If varRetVal = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            'MsgBox ("There Are No Existing Records")
                        Else
                            .Text = varRetVal
                        End If
                    End If
                    .Focus()
                End With
            Case 1
                With Me.txtReferenceNo
                    If Len(.Text) = 0 Then
                        varRetVal = ShowList(1, .MaxLength, "", "Cust_Ref", DateColumnNameInShowList("Order_Date") & "As Order_Date", "cust_ord_hdr A", "and Account_Code='" & Trim(txtCustomerCode.Text) & "' and Consignee_Code='" & Trim(txtConsCode.Text) & "' and isnull(authorized_flag,0)=0 AND REVISIONNO = (SELECT MAX(REVISIONNO) FROM CUST_ORD_HDR  WHERE UNIT_CODE=A.UNIT_CODE AND ACCOUNT_CODE=A.ACCOUNT_CODE AND CUST_REF = A.CUST_REF)  ")
                        If varRetVal = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            'MsgBox ("There Are No Existing Records")
                        Else
                            .Text = varRetVal
                        End If
                    Else
                        varRetVal = ShowList(1, .MaxLength, .Text, "Cust_Ref", DateColumnNameInShowList("Order_Date") & "As Order_Date", "cust_ord_hdr A", "and Account_Code='" & Trim(txtCustomerCode.Text) & "' and Consignee_Code='" & Trim(txtConsCode.Text) & "' and isnull(authorized_flag,0) = 0 AND REVISIONNO = (SELECT MAX(REVISIONNO) FROM CUST_ORD_HDR  WHERE UNIT_CODE=A.UNIT_CODE AND ACCOUNT_CODE=A.ACCOUNT_CODE AND CUST_REF = A.CUST_REF)")
                        If varRetVal = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            'MsgBox ("There Are No Existing Records")
                        Else
                            .Text = varRetVal
                        End If
                    End If
                    .Focus()
                End With
            Case 2
                With txtConsCode
                    If Len(.Text) = 0 Then
                        varRetVal = ShowList(1, .MaxLength, "", "Customer_Code", "Cust_Name", "Customer_mst", " and ((isnull(deactive_flag,0) <> 1) OR (convert(varchar(12),getdate(),106)<= convert(varchar(12),deactive_date,106)))")
                        If varRetVal = "-1" Then
                            ''Call ConfirmWindow(10010, BUTTON_OK, IMG_INFO)
                            MsgBox("Invalid Consignee Code.", MsgBoxStyle.Information, "eMpower")
                        Else
                            .Text = varRetVal
                        End If
                    Else
                        varRetVal = ShowList(1, .MaxLength, .Text, "Customer_Code", "Cust_Name", "Customer_mst", " and ((isnull(deactive_flag,0) <> 1) OR (convert(varchar(12),getdate(),106)<= convert(varchar(12),deactive_date,106)))")
                        If varRetVal = "-1" Then
                            ''Call ConfirmWindow(10010, BUTTON_OK, IMG_INFO)
                            MsgBox("Invalid Consignee Code.", MsgBoxStyle.Information, "eMpower")
                        Else
                            .Text = varRetVal
                        End If
                    End If
                    .Focus()
                End With
            Case 3
                With txtAmendmentNo
                    If Len(.Text) = 0 Then
                        varRetVal = ShowList(1, .MaxLength, "", "Amendment_No", "Cust_Ref", "cust_ord_hdr", "and cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'and Account_Code='" & txtCustomerCode.Text & "' and Consignee_Code='" & Trim(txtConsCode.Text) & "' and Amendment_No <>' ' and authorized_flag = 0 or authorized_flag is null ")
                        If varRetVal = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            'MsgBox ("There Are No Existing Records")
                        Else
                            .Text = varRetVal
                        End If
                    Else
                        varRetVal = ShowList(1, .MaxLength, .Text, "Amendment_No", "Cust_Ref", "cust_ord_hdr", "and  cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'and Account_Code='" & txtCustomerCode.Text & "' and Consignee_Code='" & Trim(txtConsCode.Text) & "'  and Amendment_No <>' ' and authorized_flag = 0 or authorized_flag is null ")
                        If varRetVal = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            'MsgBox ("There Are No Existing Records")
                        Else
                            .Text = varRetVal
                        End If
                    End If
                    .Focus()
                End With

        End Select
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub cmdHelp_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles cmdHelp.MouseDown
        Dim Button As Short = eventArgs.Button \ &H100000
        Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim x As Single = VB6.PixelsToTwipsX(eventArgs.X)
        Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
        Dim Index As Short = cmdHelp.GetIndex(eventSender)
        Select Case Index
            Case 0
                m_blnhelp = True
        End Select
    End Sub
    Private Sub frmMKTTRN0002_SOUTH_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        On Error GoTo ErrHandler
        'Make the node normal font
        frmModules.NodeFontBold(Tag) = False
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0002_SOUTH_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0 Then Call ctlHeader_Click(ctlHeader, New System.EventArgs()) : Exit Sub
    End Sub
    Private Sub txtAmendmentNo_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAmendmentNo.TextChanged
        txtSTax.Text = "" : txtCreditTerms.Text = ""
        'chkAddCustSupp.Value = 0
        chkOpenSo.CheckState = System.Windows.Forms.CheckState.Unchecked
        If Len(Trim(txtAmendmentNo.Text)) = 0 Then
            Call RefreshForm()
        End If
    End Sub
    Private Sub txtAmendmentNo_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtAmendmentNo.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = 112 Then Call cmdHelp_Click(cmdHelp.Item(3), New System.EventArgs())
        If KeyCode = System.Windows.Forms.Keys.Return Then
            Call txtAmendmentNo_Validating(txtAmendmentNo, New System.ComponentModel.CancelEventArgs(False))
            If blnValidAmend = True Then
                If cmdchangetype.Enabled Then cmdchangetype.Focus() Else cmdAuthorize.Focus()
            End If
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub


    Private Sub txtConsCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtConsCode.TextChanged
        On Error GoTo ErrHandler
        Dim rsdbCons As ClsResultSetDB
        If Len(Trim(txtConsCode.Text)) <> 0 Then
            m_strSql = "Select top 1 1 from cust_Ord_hdr where unit_code='" & gstrUNITID & "' and account_code ='" & txtCustomerCode.Text & "' and consignee_code='" & Trim(txtConsCode.Text) & "' "
            rsdbCons = New ClsResultSetDB
            rsdbCons.GetResult(m_strSql)
            If rsdbCons.GetNoRows > 0 Then
                txtReferenceNo.Enabled = True
                txtReferenceNo.BackColor = System.Drawing.Color.White
                cmdHelp(2).Enabled = True
            Else
                lblConsigneeDesc.Text = ""
                If txtReferenceNo.Enabled = True Then txtReferenceNo.Text = ""
            End If
            rsdbCons.ResultSetClose()
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        rsdbCons.ResultSetClose()
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub txtConsCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtConsCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*******************************************************************************
        'Author             :   Ashutosh Verma
        'Argument(s)if any  :   Issue Id: 19731
        'Return Value       :   NA
        'Function           :
        'Comments           :   NA
        'Creation Date      :   16 Apr 2007
        '*******************************************************************************

        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        GoTo EventExitSub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        GoTo EventExitSub
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub

    Private Sub txtConsCode_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtConsCode.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '*******************************************************************************
        'Author             :   Ashutosh Verma
        'Argument(s)if any  :   Issue Id: 19731
        'Return Value       :   NA
        'Function           :   If F1 Key Press Then Display Help From for Consignee code.
        'Comments           :   NA
        'Creation Date      :   16 Apr 2007
        '*******************************************************************************
        On Error GoTo ErrHandler
        If KeyCode = 112 Then Call cmdHelp_Click(cmdHelp.Item(2), New System.EventArgs()) 'Help should be invoked if F1 key is pressed
        If KeyCode = System.Windows.Forms.Keys.Return Then
            Call txtConsCode_Validating(txtConsCode, New System.ComponentModel.CancelEventArgs(False))
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub

    Private Sub txtConsCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtConsCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        '*******************************************************************************
        'Author             :   Ashutosh Verma
        'Argument(s)if any  :   Issue Id: 19731
        'Return Value       :   NA
        'Function           :   Validate Consignee code.
        'Comments           :   NA
        'Creation Date      :   16 Apr 2007
        '*******************************************************************************

        On Error GoTo ErrHandler
        Dim rsCD As New ClsResultSetDB

        If Len(Trim(txtConsCode.Text)) = 0 Then
            GoTo EventExitSub
        Else
            m_strSql = "Select Cust_Name from Customer_mst where unit_code='" & gstrUNITID & "' and customer_Code='" & Trim(txtConsCode.Text) & "' and ((isnull(deactive_flag,0) <> 1) OR (convert(varchar(12),getdate(),106)<= convert(varchar(12),deactive_date,106)))"
            rsCD.GetResult(m_strSql)
            If rsCD.GetNoRows = 0 Then
                MsgBox("Invalid Consignee Code !!!", MsgBoxStyle.Information, "eMpower")
                ''Call ConfirmWindow(10145, BUTTON_OK, IMG_INFO)
                txtConsCode.Text = ""
                lblConsigneeDesc.Text = ""
                Cancel = True
                txtConsCode.Focus()
                GoTo EventExitSub
            Else
                lblConsigneeDesc.Text = IIf(UCase(rsCD.GetValue("Cust_Name")) = "UNKNOWN", "", rsCD.GetValue("Cust_Name"))
                If txtReferenceNo.Enabled = True Then txtReferenceNo.Focus()
            End If
            rsCD.ResultSetClose()
        End If
        GoTo EventExitSub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        GoTo EventExitSub
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub


    Private Sub txtCreditTerms_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCreditTerms.TextChanged
        Call FillLabel("CREDIT")
    End Sub
    Private Sub txtCurrencyType_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCurrencyType.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = 112 Then Call cmdHelp_Click(cmdHelp.Item(2), New System.EventArgs()) 'Help should be invoked if F1 key is pressed
    End Sub

    Private Sub txtCustomerCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCustomerCode.TextChanged
        On Error GoTo ErrHandler
        Call FillLabel("CUSTOMER")
        If Len(Trim(txtCustomerCode.Text)) <> 0 Then
            txtReferenceNo.Enabled = True
            txtReferenceNo.BackColor = System.Drawing.Color.White
            cmdHelp(2).Enabled = True
            cmdHelp(1).Enabled = True
            txtConsCode.Enabled = True
            txtConsCode.BackColor = System.Drawing.Color.White
            ''ADDED BY SUMIT KUMAR ON 17 JULY 2019 FOR CHECK EXTERNAL SO FLAG
            Call HideExternalSo()
            '''txtConsCode.Focus()
        Else
            Call RefreshForm()
            lblCustDesc.Text = ""
            txtConsCode.Text = ""
            lblConsigneeDesc.Text = ""
            txtReferenceNo.Text = ""
            cmdHelp(1).Enabled = False
            cmdHelp(3).Enabled = False
            txtReferenceNo.Enabled = False
            txtReferenceNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            txtAmendmentNo.Text = ""
            txtAmendmentNo.Enabled = False
            txtReferenceNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        End If

        ssPOEntry.maxRows = 0
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub txtCustomerCode_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCustomerCode.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = 112 Then Call cmdHelp_Click(cmdHelp.Item(0), New System.EventArgs()) 'Help should be invoked if F1 key is pressed
        If KeyCode = System.Windows.Forms.Keys.Return Then
            Call txtCustomerCode_Validating(txtCustomerCode, New System.ComponentModel.CancelEventArgs(False))
            If blnValidCust = True Then
                ''Changes done By ashutosh on 16 Apr 2007, Issue Id:19731
                txtConsCode.Enabled = True
                txtConsCode.BackColor = System.Drawing.Color.White
                lblConsigneeDesc.Enabled = True
                cmdHelp(2).Enabled = True
                txtConsCode.Text = txtCustomerCode.Text
                'Me.lblConsigneeDesc.Caption = Me.lblCustDesc.Caption
                If txtReferenceNo.Enabled = True Then
                    txtReferenceNo.Focus()
                Else
                    cmdAuthorize.Focus()
                End If
                ''Changes for Issue Id:19731 end here.
            End If
        End If
    End Sub
    Private Sub txtCustomerCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtCustomerCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        Dim rsCD As New ClsResultSetDB
        blnValidCust = False
        If m_CloseButton = True Then
            m_CloseButton = False
            GoTo EventExitSub
        End If
        If m_blnhelp = True Then
            m_blnhelp = False
            GoTo EventExitSub
        End If
        If Len(Trim(txtCustomerCode.Text)) <> 0 Then
            txtReferenceNo.Enabled = True
            txtReferenceNo.BackColor = System.Drawing.Color.White
            cmdHelp(1).Enabled = True
            Call FillLabel("CUSTOMER")

            m_strSql = "Select top 1 1 from Customer_Mst where unit_code='" & gstrUNITID & "' and Customer_Code='" & Trim(txtCustomerCode.Text) & "' and ((isnull(deactive_flag,0) <> 1) OR (convert(varchar(12),getdate(),106)<= convert(varchar(12),deactive_date,106)))"
            rsCD.GetResult(m_strSql)
            If rsCD.GetNoRows = 0 Then
                'MsgBox ("Customer Code Does Not Exist")
                Call ConfirmWindow(10145, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                Cancel = True
                txtConsCode.Text = ""
                txtCustomerCode.Text = ""
                blnValidCust = False
                GoTo EventExitSub
            Else
                blnValidCust = True
                
            End If
            rsCD.ResultSetClose()
        End If
        blnValidCust = True
        GoTo EventExitSub 'This is to avoid the execution of the error handler
ErrHandler:
        rsCD.ResultSetClose()
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        GoTo EventExitSub
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    ''ADDED BY SUMIT KUMAR ON 17 JULY 2019 FOR CHECK EXTERNAL SO FLAG
    Public Sub HideExternalSo()
        On Error GoTo ErrHandler

        If DataExist("SELECT TOP 1 ENABLE_EXTERNAL_SALESNO FROM CUSTOMER_MST(NOLOCK) WHERE  UNIT_CODE='" & gstrUNITID & "' AND  CUSTOMER_CODE='" & Trim(txtCustomerCode.Text) & "' AND ENABLE_EXTERNAL_SALESNO =1 AND ((ISNULL(DEACTIVE_FLAG,0) <> 1) OR (CONVERT(VARCHAR(12),GETDATE(),106)<= CONVERT(VARCHAR(12),DEACTIVE_DATE,106)))") = True Then
            With ssPOEntry
                .Col = 20
                .Col2 = 20
                .ColHidden = False

            End With
        Else
            With ssPOEntry
                .Col = 20
                .Col2 = 20
                .ColHidden = True

            End With
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub txtReferenceNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtReferenceNo.LostFocus
        On Error GoTo ErrHandler
        Dim inti As Short
        If m_CloseButton = True Then
            m_CloseButton = False
            Exit Sub
        End If
        blnValidref = False
        If Len(Trim(txtReferenceNo.Text)) > 0 Then
            ' Check if records for the entered reference no exist or not
            m_strSql = " Select 1 from cust_ord_hdr where unit_code='" & gstrUNITID & "' and  cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'and Account_Code='" & Trim(Me.txtCustomerCode.Text) & "' and Consignee_Code='" & Trim(txtConsCode.Text) & "' "
            rsRefNo = New ClsResultSetDB
            Call rsRefNo.GetResult(m_strSql)
            If rsRefNo.GetNoRows = 1 Then ' If there are records existing for the entered reference no
                rsRefNo.ResultSetClose()
                ' check whether the PO is active or not
                m_strSql = " Select top 1 1 from cust_ord_hdr where unit_code='" & gstrUNITID & "' and cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'and Account_Code='" & Trim(Me.txtCustomerCode.Text) & "' and Consignee_Code='" & Trim(txtConsCode.Text) & "' and Active_Flag='A'"
                rsRefNo = New ClsResultSetDB
                Call rsRefNo.GetResult(m_strSql)
                If rsRefNo.GetNoRows > 0 Then 'For reference no to which no amendment had been made
                    rsRefNo.ResultSetClose()
                    m_strSql = " Select top 1 1 from cust_ord_hdr where unit_code='" & gstrUNITID & "' and cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'and Account_Code='" & Trim(Me.txtCustomerCode.Text) & "' and Consignee_Code='" & Trim(txtConsCode.Text) & "' and Active_Flag='A' and Authorized_Flag =1"
                    rsRefNo = New ClsResultSetDB
                    Call rsRefNo.GetResult(m_strSql)
                    If rsRefNo.GetNoRows > 0 Then 'For reference no to which no amendment had been made
                        Call GetDetails()
                        Call ConfirmWindow(10161, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                        txtReferenceNo.Text = ""
                        txtReferenceNo.Focus()
                        blnValidref = False
                        cmdAuthorize.Enabled(0) = False
                        cmdAuthorize.Enabled(1) = False
                        cmdAuthorize.Enabled(2) = False
                        cmdAuthorize.Enabled(3) = True
                        btnPOPreview.Visible = True
                        btnPOPreview.Enabled = True
                        Exit Sub
                    Else
                        rsRefNo.ResultSetClose()
                        Call GetDetails()
                        cmdAuthorize.Enabled(0) = True
                        cmdAuthorize.Enabled(1) = True
                        cmdAuthorize.Enabled(2) = False
                        cmdAuthorize.Enabled(3) = True
                        cmdAuthorize.Focus()
                        btnPOPreview.Visible = True
                        btnPOPreview.Enabled = True
                        blnValidref = True
                        Exit Sub
                    End If


                Else
                    rsRefNo.ResultSetClose()
                    Call ConfirmWindow(10162, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                    txtReferenceNo.Text = ""
                    txtReferenceNo.Focus()
                    blnValidref = False
                    btnPOPreview.Visible = False
                    btnPOPreview.Enabled = False
                    Exit Sub
                End If
            ElseIf rsRefNo.GetNoRows > 1 Then  ' If An amendment exists for the reference no
                rsRefNo.ResultSetClose()
                txtAmendmentNo.Enabled = True
                txtAmendmentNo.BackColor = System.Drawing.Color.White
                cmdHelp(3).Enabled = True
                If txtAmendmentNo.Enabled = True Then
                    txtAmendmentNo.Focus()
                End If
                blnValidref = True
                btnPOPreview.Visible = True
                btnPOPreview.Enabled = True
            Else
                rsRefNo.ResultSetClose()
                Call ConfirmWindow(10129, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                txtReferenceNo.Text = ""
                txtReferenceNo.Focus()
                blnValidref = False
                btnPOPreview.Visible = False
                Exit Sub
            End If
        End If
        blnValidref = True
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        rsRefNo.ResultSetClose()
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub

    End Sub

    Private Sub txtReferenceNo_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtReferenceNo.TextChanged
        If Len(Trim(txtReferenceNo.Text)) = 0 Then
            Call RefreshForm()
            lblCustDesc.Text = ""
            txtAmendmentNo.Text = ""
            txtSTax.Text = "" : txtCreditTerms.Text = ""
            'chkAddCustSupp.Value = 0
            chkOpenSo.CheckState = System.Windows.Forms.CheckState.Unchecked
            cmdHelp(3).Enabled = False
            txtAmendmentNo.Enabled = False
            txtAmendmentNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        End If
    End Sub
    Private Sub TxtReferenceNo_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtReferenceNo.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = 112 Then Call cmdHelp_Click(cmdHelp.Item(1), New System.EventArgs())
        If KeyCode = System.Windows.Forms.Keys.Return Then
            If txtAmendmentNo.Enabled = True Then txtAmendmentNo.Focus() Else cmdAuthorize.Focus()
        End If
    End Sub

    Private Sub frmMKTTRN0002_SOUTH_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo ErrHandler
        'Checking the form name in the Windows list
        mdifrmMain.CheckFormName = mintFormIndex
        frmModules.NodeFontBold(Tag) = True
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0002_SOUTH_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        DTDate.Format = DateTimePickerFormat.Custom
        DTDate.CustomFormat = gstrDateFormat

        DTValidDate.Format = DateTimePickerFormat.Custom
        DTValidDate.CustomFormat = gstrDateFormat

        DTAmendmentDate.Format = DateTimePickerFormat.Custom
        DTAmendmentDate.CustomFormat = gstrDateFormat

        DTEffectiveDate.Format = DateTimePickerFormat.Custom
        DTEffectiveDate.CustomFormat = gstrDateFormat

        'Get the index of form in the Windows list
        mintFormIndex = mdifrmMain.AddFormNameToWindowList(ctlHeader.Tag)
        'Load the captions
        Call FillLabelFromResFile(Me)
        'Size the form to client workspace
        Call FitToClient(Me, fraContainer, ctlHeader, cmdAuthorize, 400)
        'Disabling the controls
        Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 0)
        Call EnableControls(False, Me, True)
        'Initialising the buttons
        'Disabling Authorize, Refresh  buttons
        cmdAuthorize.Enabled(0) = False
        cmdAuthorize.Enabled(1) = False
        cmdAuthorize.Enabled(2) = False
        cmdAuthorize.Enabled(3) = True
        '10736222-changes done for CT2 ARE 3 functionality - to enable CT2 Reqd flag
        ChkCT2reqd.Enabled = False
        ChkCT2reqd.Checked = False
        cmdHelp(0).Enabled = True
        txtCustomerCode.Enabled = True
        txtCustomerCode.BackColor = System.Drawing.Color.White
        Call AddPOType()
        cmbPOType.Text = ObsoleteManagement.GetItemString(cmbPOType, 0)
        Me.DTDate.Value = GetServerDate() 'VB6.Format(ServerDate(), "dd/mm/yyyy")
        Me.DTAmendmentDate.Value = GetServerDate() 'VB6.Format(ServerDate(), "dd/mm/yyyy")
        Me.DTEffectiveDate.Value = GetServerDate() 'VB6.Format(ServerDate(), "dd/mm/yyyy")
        Me.DTValidDate.Value = GetServerDate() 'VB6.Format(ServerDate(), "dd/mm/yyyy")
        ssPOEntry.Enabled = False
        m_CloseButton = False
        With ssPOEntry
           
            .Col = 20
            .Col2 = 20
            .ColHidden = True

        End With
    End Sub
    Private Sub frmMKTTRN0002_SOUTH_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        Dim enmValue As eMPowerFunctions.ConfirmWindowReturnEnum
        If UnloadMode >= 0 And UnloadMode <= 5 Then
        End If
        'Checking the status
        If gblnCancelUnload = True Then
            eventArgs.Cancel = True
            'txtAmendmentNo.SetFocus
        End If
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub frmMKTTRN0002_SOUTH_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        'Releasing the form reference
        Me.Dispose()
        'Removing the form name from list
        mdifrmMain.RemoveFormNameFromWindowList = mintFormIndex
        'Setting the corresponding node's tag
        frmModules.NodeFontBold(Tag) = False
    End Sub
    Private Sub ssSetFocus(ByRef Row As Integer, Optional ByRef Col As Integer = 3)
        '----------------------------------------------------------------------------
        'Argument       :   Row, Col
        'Return Value   :   Nil
        'Function       :   Set the focus according to row and col value
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        With ssPOEntry
            .Row = Row
            .Col = Col
            .Action = 0
        End With
    End Sub
    Private Sub txtAmendmentNo_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtAmendmentNo.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        Dim inti As Short
        Dim rsAmend As ClsResultSetDB
        blnValidAmend = False
        If m_CloseButton = True Then
            m_CloseButton = False
            GoTo EventExitSub
        End If
        If Len(Trim(txtAmendmentNo.Text)) > 0 Then
            m_strSql = " Select top 1 1 from cust_ord_hdr where unit_code='" & gstrUNITID & "' and cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "' and Account_Code='" & Trim(Me.txtCustomerCode.Text) & "' and Consignee_Code='" & Trim(txtConsCode.Text) & "' and amendment_No='" & txtAmendmentNo.Text & "'"
            rsAmend = New ClsResultSetDB
            Call rsAmend.GetResult(m_strSql)
            If rsAmend.GetNoRows > 0 Then ' If there are records existing for the entered Amendment no
                rsAmend.ResultSetClose()
                ' check whether the PO is active or not
                m_strSql = " Select top 1 1 from cust_ord_hdr where unit_code='" & gstrUNITID & "' and cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "' and Account_Code='" & Trim(Me.txtCustomerCode.Text) & "' and Consignee_Code='" & Trim(txtConsCode.Text) & "' and Active_Flag='A' and amendment_No='" & txtAmendmentNo.Text & "'"
                rsAmend = New ClsResultSetDB
                Call rsAmend.GetResult(m_strSql)
                If rsAmend.GetNoRows > 0 Then ' If An amendment exists for the reference no
                    rsAmend.ResultSetClose()
                    m_strSql = " Select top 1 1 from cust_ord_hdr where unit_code='" & gstrUNITID & "' and cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "' and Account_Code='" & Trim(Me.txtCustomerCode.Text) & "' and Consignee_Code='" & Trim(txtConsCode.Text) & "' and Active_Flag='A' and amendment_No='" & txtAmendmentNo.Text & "' and Authorized_Flag= 1"
                    rsAmend = New ClsResultSetDB
                    Call rsAmend.GetResult(m_strSql)
                    If rsAmend.GetNoRows > 0 Then ' If the amendment is Already Authorized
                        rsAmend.ResultSetClose()
                        Call GetAmendmentDetails()
                        cmdAuthorize.Enabled(0) = False
                        cmdAuthorize.Enabled(1) = False
                        cmdAuthorize.Enabled(2) = False
                        cmdAuthorize.Enabled(3) = True
                        'MsgBox (" This Amendment is Already Authorized")
                        Call ConfirmWindow(10161, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                        Cancel = True
                        blnValidAmend = False
                        GoTo EventExitSub
                    Else
                        rsAmend.ResultSetClose()
                        Call GetAmendmentDetails()
                        cmdAuthorize.Enabled(0) = True
                        cmdAuthorize.Enabled(1) = True
                        cmdAuthorize.Enabled(2) = False
                        cmdAuthorize.Enabled(3) = True
                        blnValidAmend = True
                        GoTo EventExitSub
                    End If
                Else
                    rsAmend.ResultSetClose()
                    'MsgBox "This Amendment Is No More Valid" ' If PO is not Valid. where active flag <> A
                    Call ConfirmWindow(10160, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                    'Call GetAmendmentDetails
                    txtAmendmentNo.Text = ""
                    Cancel = True
                    blnValidAmend = False
                    GoTo EventExitSub
                End If
            Else
                rsAmend.ResultSetClose()
                'MsgBox "There are no existing records for the Amendment No"
                Call ConfirmWindow(10128, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                txtAmendmentNo.Text = ""
                Cancel = True
                blnValidAmend = False
                GoTo EventExitSub
            End If
        End If
        blnValidAmend = True
        GoTo EventExitSub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        GoTo EventExitSub
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Public Function FillLabel(ByRef pstrCode As Object) As Object
        On Error GoTo ErrHandler
        '----------------------------------------------------------------------------
        'Argument       :   Nil
        'Return Value   :   Nil
        'Function       :   fills the customer detail label
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        Dim rsCust As ClsResultSetDB


        Select Case UCase(pstrCode)
            Case "CUSTOMER"
                m_strSql = "select Cust_Name,Credit_Days from Customer_mst where unit_code='" & gstrUNITID & "' and Customer_code='" & txtCustomerCode.Text & "' and ((isnull(deactive_flag,0) <> 1) OR (convert(varchar(12),getdate(),106)<= convert(varchar(12),deactive_date,106)))"
                rsCust = New ClsResultSetDB
                rsCust.GetResult(m_strSql)
                lblCustDesc.ForeColor = System.Drawing.Color.White
                lblCustDesc.Text = IIf(UCase(rsCust.GetValue("Cust_Name")) = "UNKNOWN", "", rsCust.GetValue("Cust_Name"))
                rsCust.ResultSetClose()
            Case "STAX"
                m_strSql = "select TxRt_Rate_No,TxRt_RateDesc from Gen_TaxRate where unit_code='" & gstrUNITID & "' and Txrt_Rate_No = '" & txtSTax.Text & "'"
                rsCust = New ClsResultSetDB
                rsCust.GetResult(m_strSql)
                lblSTaxDesc.ForeColor = System.Drawing.Color.White
                lblSTaxDesc.Text = IIf(UCase(rsCust.GetValue("TxRt_RateDesc")) = "UNKNOWN", "", rsCust.GetValue("TxRt_RateDesc"))
                rsCust.ResultSetClose()
            Case "CREDIT"
                m_strSql = "select crtrm_TermID,crtrm_Desc from Gen_CreditTrmMaster where unit_code='" & gstrUNITID & "' and crtrm_TermID = '" & txtCreditTerms.Text & "'"
                rsCust = New ClsResultSetDB
                rsCust.GetResult(m_strSql)
                lblCreditTermDesc.ForeColor = System.Drawing.Color.White
                lblCreditTermDesc.Text = IIf(UCase(rsCust.GetValue("crtrm_Desc")) = "UNKNOWN", "", rsCust.GetValue("crtrm_Desc"))
                rsCust.ResultSetClose()
        End Select
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Function
    End Function
    Public Sub GetAmendmentDetails()
        '----------------------------------------------------------------------------
        'Argument       :   Nil
        'Return Value   :   Nil
        'Function       :   Get the details if there is an amendment
        'Comments       :   Nil
        '=======================================================================================
        'Revised By     : Ashutosh Verma
        'Revised On     : 22-01-2007 ,Issue ID:19352
        'Revised Reason : Show Additional Excise Duty in grid.
        '=======================================================================================


        On Error GoTo ErrHandler
        Dim rsAD As New ClsResultSetDB
        Dim rscurrency As ClsResultSetDB

        m_strSql = "select Account_Code,Cust_Ref,Amendment_No,Order_Date,Amendment_Date,Active_Flag,Currency_Code,Valid_Date,Effect_Date,Term_Payment,Special_Remarks,Pay_Remarks,Price_Remarks,Packing_Remarks,Frieght_Remarks,Transport_Remarks,Octorai_Remarks,Mode_Despatch,Delivery,First_Authorized,Second_Authorized,Third_Authorized,Authorized_Flag,Reason,PO_Type,SalesTax_Type,OpenSO,AddCustSupp,PerValue,InternalSONo,RevisionNo,Surcharge_code,Future_SO,ECESS_Code,Consignee_Code,ADDVAT_TYPE,CT2_REQD_IN_SO from cust_ord_hdr where unit_code='" & gstrUNITID & "' and Cust_ref='" & txtReferenceNo.Text & "' and amendment_No='" & txtAmendmentNo.Text & "' AND ACCOUNT_CODE='" & txtCustomerCode.Text & "' and Consignee_Code='" & Trim(txtConsCode.Text) & "' and Active_Flag='A'"
        rsAD.GetResult(m_strSql)
        m_strSql = "select Account_Code,Cust_Ref,Amendment_No,Item_Code,Rate,Order_Qty,Despatch_Qty,Active_Flag,Cust_Mtrl,Cust_DrgNo,Packing,Others,Excise_Duty,Cust_Drg_Desc,Tool_Cost,Authorized_flag,Ent_dt,OpenSO,SalesTax_Type,PerValue,InternalSONo,RevisionNo,Remarks,MRP,Abantment_Code,AccessibleRateforMRP,Packing_Type,TOOL_AMOR_FLAG,ShowInAuth,ADD_Excise_Duty,HSNSACCODE,ISHSNORSAC,CGSTTXRT_TYPE,SGSTTXRT_TYPE,UTGSTTXRT_TYPE,IGSTTXRT_TYPE,COMPENSATION_CESS,EXTERNAL_SALESORDER_NO  from cust_ord_dtl where unit_code='" & gstrUNITID & "' and Cust_ref='" & txtReferenceNo.Text & "' and amendment_No='" & txtAmendmentNo.Text & "' AND ACCOUNT_CODE='" & txtCustomerCode.Text & "' and Active_Flag='A' and ShowInAuth = 1"
        rsdb = New ClsResultSetDB
        rsdb.GetResult(m_strSql)
        Dim intLoopCounter As Short
        Dim intDecimal As Short
        Dim strMin As String
        Dim strMax As String
        If rsAD.GetNoRows > 0 Then
            rsAD.MoveFirst()
            Me.DTDate.Value = rsAD.GetValue("Order_Date") 'VB6.Format(rsAD.GetValue("Order_Date"), "dd/mm/yyyy")
            DTAmendmentDate.Value = rsAD.GetValue("Amendment_Date") ' VB6.Format(rsAD.GetValue("Amendment_Date"), "dd/mm/yyyy")
            DTEffectiveDate.Value = rsAD.GetValue("Effect_Date") 'VB6.Format(rsAD.GetValue("Effect_Date"), "dd/mm/yyyy")
            DTValidDate.Value = rsAD.GetValue("Valid_Date") 'VB6.Format(rsAD.GetValue("Valid_Date"), "dd/mm/yyyy")
            txtCurrencyType.Text = rsAD.GetValue("Currency_Code")
            '10552513-Starts
            txtSChSTax.Text = rsAD.GetValue("Surcharge_code")
            txtECSSTaxType.Text = rsAD.GetValue("ECESS_Code")
            '10552513-Ends
            ctlPerValue.Text = rsAD.GetValue("PerValue")
            strpotype = rsAD.GetValue("PO_Type")
            'Select Case cmbPOType.Text
            Select Case UCase(strpotype)
                Case "O"
                    Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 1)
                Case "S"
                    Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 2)
                Case "J"
                    Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 3)
                Case "E"
                    Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 4)
                Case "V"
                    Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 5)
            End Select
            'To Display Value in Sales Tax,CreDit Terms Open SO Flag,Add CustSupp Flag
            txtSTax.Text = rsAD.GetValue("SalesTax_Type")
            txtCreditTerms.Text = rsAD.GetValue("term_Payment")
            '10736222-changes done for CT2 ARE 3 functionality - to enable CT2 Reqd flag
            ChkCT2reqd.Checked = rsAD.GetValue("CT2_REQD_IN_SO")

            If rsAD.GetValue("OpenSO") = False Then
                chkOpenSo.CheckState = System.Windows.Forms.CheckState.Unchecked
            Else
                chkOpenSo.CheckState = System.Windows.Forms.CheckState.Checked
            End If
            '****
            rsAD.MoveFirst()
            ssPOEntry.MaxRows = 0
            Do While Not rsdb.EOFRecord
                ssPOEntry.MaxRows = ssPOEntry.MaxRows + 1
                ssPOEntry.Col = 1
                ssPOEntry.Col2 = 1
                ssPOEntry.Row = 1
                ssPOEntry.Row2 = ssPOEntry.MaxRows
                ssPOEntry.BlockMode = True
                ssPOEntry.CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
                ssPOEntry.BlockMode = False
                ssPOEntry.Col = 10
                ssPOEntry.Col2 = 10
                ssPOEntry.Row = 1
                ssPOEntry.Row2 = ssPOEntry.MaxRows
                ssPOEntry.BlockMode = True
                ssPOEntry.CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                ssPOEntry.BlockMode = False
                '***** for accounts plug in
                'Changed to add open Item Falg in Grid
                If rsdb.GetValue("OpenSO") = False Then
                    ssPOEntry.Row = ssPOEntry.MaxRows
                    ssPOEntry.Col = 1
                    ssPOEntry.Value = 0
                Else
                    ssPOEntry.Row = ssPOEntry.MaxRows
                    ssPOEntry.Col = 1
                    ssPOEntry.Value = 1
                End If
                Call ssPOEntry.SetText(2, ssPOEntry.MaxRows, rsdb.GetValue("Cust_DrgNo"))
                Call ssPOEntry.SetText(3, ssPOEntry.MaxRows, rsdb.GetValue("Item_Code "))
                Call ssPOEntry.SetText(4, ssPOEntry.MaxRows, rsdb.GetValue("Cust_Drg_Desc"))
                Call ssPOEntry.SetText(5, ssPOEntry.MaxRows, rsdb.GetValue("Order_Qty"))
                '************** for add decimal places
                If Len(Trim(txtCurrencyType.Text)) Then
                    rscurrency = New ClsResultSetDB
                    rscurrency.GetResult("Select decimal_Place from Currency_Mst where unit_code='" & gstrUNITID & "' and currency_code ='" & Trim(txtCurrencyType.Text) & "'")
                    intDecimal = rscurrency.GetValue("Decimal_Place")
                    rscurrency.ResultSetClose()
                End If
                If intDecimal <= 0 Then
                    intDecimal = 2
                End If
                strMin = "0." : strMax = "99999999."
                For intLoopCounter = 1 To intDecimal
                    strMin = strMin & "0"
                    strMax = strMax & "9"
                Next
                '**************
                For intLoopCounter = 6 To 9
                    With Me.ssPOEntry
                        .Row = .MaxRows
                        .Col = intLoopCounter
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                        .TypeFloatDecimalPlaces = intDecimal
                        .TypeFloatMax = strMax
                        .TypeFloatMin = strMin
                    End With
                Next
                Call ssPOEntry.SetText(6, ssPOEntry.MaxRows, rsdb.GetValue("Rate") * CDbl(ctlPerValue.Text))
                Call ssPOEntry.SetText(7, ssPOEntry.MaxRows, rsdb.GetValue("Cust_Mtrl") * CDbl(ctlPerValue.Text))
                Call ssPOEntry.SetText(8, ssPOEntry.MaxRows, rsdb.GetValue("Tool_Cost") * CDbl(ctlPerValue.Text))
                Call ssPOEntry.SetText(9, ssPOEntry.MaxRows, rsdb.GetValue("Packing"))
                Call ssPOEntry.SetText(10, ssPOEntry.MaxRows, rsdb.GetValue("Excise_Duty"))
                Call ssPOEntry.SetText(11, ssPOEntry.MaxRows, rsdb.GetValue("Others") * CDbl(ctlPerValue.Text))
                ssPOEntry.Col = 13
                ssPOEntry.Col2 = 13
                ssPOEntry.Row = ssPOEntry.MaxRows
                ssPOEntry.Row2 = ssPOEntry.MaxRows
                ssPOEntry.BlockMode = True
                ssPOEntry.CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                ssPOEntry.BlockMode = False
                Call ssPOEntry.SetText(13, ssPOEntry.MaxRows, rsdb.GetValue("ADD_Excise_Duty"))
                'GST DETAILS
                If gblnGSTUnit = True Then
                    Call ssPOEntry.SetText(14, ssPOEntry.MaxRows, rsdb.GetValue("HSNSACCODE"))
                    Call ssPOEntry.SetText(15, ssPOEntry.MaxRows, rsdb.GetValue("CGSTTXRT_TYPE"))
                    Call ssPOEntry.SetText(16, ssPOEntry.MaxRows, rsdb.GetValue("SGSTTXRT_TYPE"))
                    Call ssPOEntry.SetText(17, ssPOEntry.MaxRows, rsdb.GetValue("IGSTTXRT_TYPE"))
                    Call ssPOEntry.SetText(18, ssPOEntry.MaxRows, rsdb.GetValue("UTGSTTXRT_TYPE"))
                    Call ssPOEntry.SetText(19, ssPOEntry.MaxRows, rsdb.GetValue("COMPENSATION_CESS"))
                End If
                'GST DETAILS
                ''ADDED BY SUMIT KUMAR ON 09 JULY 2019
                Call ssPOEntry.SetText(20, ssPOEntry.MaxRows, rsdb.GetValue("EXTERNAL_SALESORDER_NO"))
                ''Changes for Issue Id:19352 end here.
                rsdb.MoveNext()
            Loop
        End If
        With ssPOEntry
            '*****Changed by nisha for accounts plug in
            'Changed to add open Item Falg in Grid
            .BlockMode = True
            .Col = 1
            .Col2 = 12
            .Row = 1
            .Row2 = .MaxRows
            .Lock = True
            .BlockMode = False
            .Enabled = True
        End With
        cmdchangetype.Enabled = True
        rsAD.ResultSetClose()
        rsdb.ResultSetClose()
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Public Sub GetDetails()
        '----------------------------------------------------------------------------
        'Argument       :   Nil
        'Return Value   :   Nil
        'Function       :   Get the details if there is no amendment
        'Comments       :   Nil
        '=======================================================================================
        'Revised By     : Ashutosh Verma
        'Revised On     : 22-01-2007 ,Issue ID:19352
        'Revised Reason : Show Additional Excise Duty in grid.
        '=======================================================================================
        On Error GoTo ErrHandler
        Dim rsAD As New ClsResultSetDB
        Dim rscurrency As ClsResultSetDB
        Dim rsdb2 As ClsResultSetDB
        Dim strAuthFlg As String
        Dim VarCT2 As Object = Nothing
        m_strSql = "select Account_Code,Cust_Ref,Amendment_No,Order_Date,Amendment_Date,Active_Flag,Currency_Code,Valid_Date,Effect_Date,Term_Payment,Special_Remarks,Pay_Remarks,Price_Remarks,Packing_Remarks,Frieght_Remarks,Transport_Remarks,Octorai_Remarks,Mode_Despatch,Delivery,First_Authorized,Second_Authorized,Third_Authorized,Authorized_Flag,Reason,PO_Type,SalesTax_Type,OpenSO,AddCustSupp,PerValue,InternalSONo,RevisionNo,Surcharge_code,Future_SO,ECESS_Code,Consignee_Code,ADDVAT_TYPE,CT2_REQD_IN_SO,SERVICETAX_TYPE from cust_ord_hdr where unit_code='" & gstrUNITID & "' and Cust_ref='" & txtReferenceNo.Text & "' and Account_Code='" & txtCustomerCode.Text & "'  and Consignee_Code='" & Trim(txtConsCode.Text) & "' and active_Flag='A'"
        rsAD.GetResult(m_strSql)
        m_strSql = "select Account_Code,Cust_Ref,Amendment_No,Item_Code,Rate,Order_Qty,Despatch_Qty,Active_Flag,Cust_Mtrl,Cust_DrgNo,Packing,Others,Excise_Duty,Cust_Drg_Desc,Tool_Cost,Authorized_flag,Ent_dt,OpenSO,SalesTax_Type,PerValue,InternalSONo,RevisionNo,Remarks,MRP,Abantment_Code,AccessibleRateforMRP,Packing_Type,TOOL_AMOR_FLAG,ShowInAuth,ADD_Excise_Duty ,HSNSACCODE,ISHSNORSAC,CGSTTXRT_TYPE,SGSTTXRT_TYPE,UTGSTTXRT_TYPE,IGSTTXRT_TYPE,COMPENSATION_CESS,EXTERNAL_SALESORDER_NO  from cust_ord_dtl where unit_code='" & gstrUNITID & "' and Cust_ref='" & txtReferenceNo.Text & "' and Account_Code='" & txtCustomerCode.Text & "' and active_Flag='A' and ShowInAuth = 1 "
        rsdb2 = New ClsResultSetDB
        rsdb2.GetResult(m_strSql)
        Dim intLoopCounter As Short
        Dim intDecimal As Short
        Dim strMin As String
        Dim strMax As String
        If rsAD.GetNoRows > 0 Then
            rsAD.MoveFirst()
            Me.DTDate.Value = rsAD.GetValue("Order_Date") 'VB6.Format(rsAD.GetValue("Order_Date"), "dd/mm/yyyy")
            If Len(Trim(rsAD.GetValue("Amendment_No"))) > 0 Then
                DTAmendmentDate.Value = rsAD.GetValue("Amendment_Date") 'VB6.Format(rsAD.GetValue("Amendment_Date"), "dd/mm/yyyy")
            End If
            '10736222-changes done for CT2 ARE 3 functionality - to enable CT2 Reqd flag
            ChkCT2reqd.Checked = rsAD.GetValue("CT2_REQD_IN_SO")
          

            DTEffectiveDate.Value = rsAD.GetValue("Effect_Date") ' VB6.Format(rsAD.GetValue("Effect_Date"), "dd/mm/yyyy")
            DTValidDate.Value = rsAD.GetValue("Valid_Date") 'VB6.Format(rsAD.GetValue("Valid_Date"), "dd/mm/yyyy")
            txtCurrencyType.Text = rsAD.GetValue("Currency_Code")
            '10552513-Starts
            txtSChSTax.Text = rsAD.GetValue("Surcharge_code")
            txtECSSTaxType.Text = rsAD.GetValue("ECESS_Code")
            '10552513-Ends
            '***17/10/2002 add by nisha to add per value details
            ctlPerValue.Text = rsAD.GetValue("PerValue")
            '***
            strpotype = rsAD.GetValue("PO_Type")
            Select Case UCase(strpotype)
                Case "O"
                    Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 1)
                Case "S"
                    Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 2)
                Case "J"
                    Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 3)
                Case "E"
                    Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 4)
                Case "V"
                    Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 5)
            End Select
            '**** for accounts Plug in
            'to show the details of Sales Tax,Credit Days,AddCustSupplied Flag,Open SO Flag
            txtSTax.Text = rsAD.GetValue("SalesTax_Type")
            txtCreditTerms.Text = rsAD.GetValue("term_Payment")
            TxtServiceCode.Text = IIf(IsDBNull(rsAD.GetValue("SERVICETAX_TYPE")), "", rsAD.GetValue("SERVICETAX_TYPE"))

            If rsAD.GetValue("OpenSO") = False Then
                chkOpenSo.CheckState = System.Windows.Forms.CheckState.Unchecked
            Else
                chkOpenSo.CheckState = System.Windows.Forms.CheckState.Checked
            End If
            '****
            rsAD.MoveFirst()
            ssPOEntry.MaxRows = 0
            '***** for accounts plug in
            'Changed to add open Item Falg in Grid
            Do While Not rsdb2.EOFRecord
                ssPOEntry.MaxRows = ssPOEntry.MaxRows + 1
                ssPOEntry.Col = 1
                ssPOEntry.Col2 = 1
                ssPOEntry.Row = 1
                ssPOEntry.Row2 = ssPOEntry.MaxRows
                ssPOEntry.BlockMode = True
                ssPOEntry.CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
                ssPOEntry.BlockMode = False
                ssPOEntry.Col = 10
                ssPOEntry.Col2 = 10
                ssPOEntry.Row = ssPOEntry.MaxRows
                ssPOEntry.Row2 = ssPOEntry.MaxRows
                ssPOEntry.BlockMode = True
                ssPOEntry.CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                ssPOEntry.BlockMode = False
                '*****for accounts plug in
                'Changed to add open Item Falg in Grid
                If rsdb2.GetValue("OpenSO") = False Then
                    ssPOEntry.Row = ssPOEntry.MaxRows
                    ssPOEntry.Col = 1
                    ssPOEntry.Value = 0
                Else
                    ssPOEntry.Row = ssPOEntry.MaxRows
                    ssPOEntry.Col = 1
                    ssPOEntry.Value = 1
                End If
                Call ssPOEntry.SetText(2, ssPOEntry.MaxRows, rsdb2.GetValue("Cust_DrgNo"))
                Call ssPOEntry.SetText(3, ssPOEntry.MaxRows, rsdb2.GetValue("Item_Code"))
                Call ssPOEntry.SetText(4, ssPOEntry.MaxRows, rsdb2.GetValue("Cust_Drg_Desc"))
                Call ssPOEntry.SetText(5, ssPOEntry.MaxRows, rsdb2.GetValue("Order_Qty"))
                '*********by nisha for decimal places
                If Len(Trim(txtCurrencyType.Text)) Then
                    rscurrency = New ClsResultSetDB
                    rscurrency.GetResult("Select decimal_Place from Currency_Mst where unit_code='" & gstrUNITID & "' and currency_code ='" & Trim(txtCurrencyType.Text) & "'")
                    intDecimal = rscurrency.GetValue("Decimal_Place")
                    rscurrency.ResultSetClose()
                End If

                If intDecimal <= 0 Then
                    intDecimal = 2
                End If
                strMin = "0." : strMax = "99999999."
                For intLoopCounter = 1 To intDecimal
                    strMin = strMin & "0"
                    strMax = strMax & "9"
                Next
                '*********
                ' for accounts plug in
                'Changed to add open Item Falg in Grid
                For intLoopCounter = 6 To 9
                    With Me.ssPOEntry
                        .Row = .MaxRows
                        .Col = intLoopCounter
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                        .TypeFloatDecimalPlaces = intDecimal
                        .TypeFloatMax = strMax
                        .TypeFloatMin = strMin
                    End With
                Next
                Call ssPOEntry.SetText(6, ssPOEntry.MaxRows, rsdb2.GetValue("Rate") * CDbl(ctlPerValue.Text))
                Call ssPOEntry.SetText(7, ssPOEntry.MaxRows, rsdb2.GetValue("Cust_Mtrl") * CDbl(ctlPerValue.Text))
                Call ssPOEntry.SetText(8, ssPOEntry.MaxRows, rsdb2.GetValue("Tool_Cost") * CDbl(ctlPerValue.Text))
                Call ssPOEntry.SetText(9, ssPOEntry.MaxRows, rsdb2.GetValue("Packing"))
                '*********
                Call ssPOEntry.SetText(10, ssPOEntry.MaxRows, rsdb2.GetValue("Excise_Duty"))
                Call ssPOEntry.SetText(11, ssPOEntry.MaxRows, rsdb2.GetValue("Others") * CDbl(ctlPerValue.Text))
                ssPOEntry.Col = 13
                ssPOEntry.Col2 = 13
                ssPOEntry.Row = ssPOEntry.MaxRows
                ssPOEntry.Row2 = ssPOEntry.MaxRows
                ssPOEntry.BlockMode = True
                ssPOEntry.CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                ssPOEntry.BlockMode = False
                'Issue Id:19352, Show Add Excise Duty in grid.
                Call ssPOEntry.SetText(13, ssPOEntry.MaxRows, rsdb2.GetValue("ADD_Excise_Duty"))
                If gblnGSTUnit = True Then
                    Call ssPOEntry.SetText(14, ssPOEntry.MaxRows, rsdb2.GetValue("HSNSACCODE"))
                    Call ssPOEntry.SetText(15, ssPOEntry.MaxRows, rsdb2.GetValue("CGSTTXRT_TYPE"))
                    Call ssPOEntry.SetText(16, ssPOEntry.MaxRows, rsdb2.GetValue("SGSTTXRT_TYPE"))
                    Call ssPOEntry.SetText(17, ssPOEntry.MaxRows, rsdb2.GetValue("IGSTTXRT_TYPE"))
                    Call ssPOEntry.SetText(18, ssPOEntry.MaxRows, rsdb2.GetValue("UTGSTTXRT_TYPE"))
                    Call ssPOEntry.SetText(19, ssPOEntry.MaxRows, rsdb2.GetValue("COMPENSATION_CESS"))
                End If
                ''ADDED BY SUMIT KUMAR ON 09 JULY 2019
                Call ssPOEntry.SetText(20, ssPOEntry.MaxRows, rsdb2.GetValue("EXTERNAL_SALESORDER_NO"))
                ''Changes for Issue Id:19352 end here.
                rsdb2.MoveNext()
            Loop
        End If
        With ssPOEntry
            '*****Changed by nisha for accounts plug in
            'Changed to add open Item Falg in Grid
            .BlockMode = True
            .Col = 1
            .Col2 = 11
            .Row = 1
            .Row2 = .MaxRows
            .Lock = True
            .BlockMode = False
            .Enabled = True
        End With
        cmdchangetype.Enabled = True
        rsAD.ResultSetClose()
        rsdb2.ResultSetClose()
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Public Sub AddPOType()
        '10869290
        cmbPOType.Items.Insert(0, "None")
        cmbPOType.Items.Insert(1, "OEM")
        cmbPOType.Items.Insert(2, "SPARES")
        cmbPOType.Items.Insert(3, "JOB WORK")
        cmbPOType.Items.Insert(4, "Export")
        cmbPOType.Items.Insert(5, "V-SERVICE")
    End Sub
    Public Sub RefreshForm()
        txtCurrencyType.Text = ""
        cmbPOType.Text = ObsoleteManagement.GetItemString(cmbPOType, 0)
        Me.DTDate.Value = GetServerDate() 'VB6.Format(ServerDate(), "dd/mm/yyyy")
        Me.DTAmendmentDate.Value = GetServerDate() 'VB6.Format(ServerDate(), "dd/mm/yyyy")
        Me.DTEffectiveDate.Value = GetServerDate() 'VB6.Format(ServerDate(), "dd/mm/yyyy")
        Me.DTValidDate.Value = GetServerDate() 'VB6.Format(ServerDate(), "dd/mm/yyyy")
        cmdAuthorize.Enabled(0) = False
        cmdAuthorize.Enabled(1) = False
        cmdAuthorize.Enabled(2) = False
        cmdAuthorize.Enabled(3) = True
        '10736222-changes done for CT2 ARE 3 functionality - to enable CT2 Reqd flag
        ChkCT2reqd.Enabled = False
        ChkCT2reqd.Checked = False
        ssPOEntry.MaxRows = 0
        btnPOPreview.Visible = False
        btnPOPreview.Enabled = False
    End Sub
    Public Function ValidAuthorize() As Boolean
        Dim strActiveSalesorder As String = String.Empty
        ValidAuthorize = False
        Dim SALESORDERLIST As String = String.Empty
        Dim rsSoDtl As New ClsResultSetDB
        Dim STROPENCLOSEDSO As String

        strActiveSalesorder = Find_Value("SELECT DBO.UDF_CHECK_ACTIVE_SO_SINGLECUSTOMER('" & gstrUNITID & "','" & Me.txtCustomerCode.Text.Trim & "')")
        If strActiveSalesorder <> "" Then
            MsgBox("ALREADY ACTIVE SALES ORDER EXISTS ,KINDLY LOCK THE BELOW SALES ORDER FIRST (AMENDMENT NO ) : ." & strActiveSalesorder, MsgBoxStyle.Information, ResolveResString(100))
            ValidAuthorize = False
            Exit Function
        End If

        '10830216

        'rsSoDtl.GetResult("SELECT * FROM cust_ord_dtl Where Cust_ref = '" & txtReferenceNo.Text.Trim & "' and UNIT_CODE = '" & gstrUNITID & "' and Account_Code = '" & txtCustomerCode.Text.Trim & "' and Active_Flag='A' and amendment_no ='" & txtAmendmentNo.Text.Trim & "'")
        'If rsSoDtl.GetNoRows > 0 Then
        '    rsSoDtl.MoveFirst()
        '    While Not rsSoDtl.EOFRecord
        '        SALESORDERLIST = Find_Value("Select dbo.UDF_CHECK_NOOF_ACTIVESO('" & gstrUNITID & "','" & txtCustomerCode.Text & "','" & rsSoDtl.GetValue("Cust_ref") & "','" & rsSoDtl.GetValue("Cust_DrgNo") & "','" & rsSoDtl.GetValue("item_code") & "',CONVERT(VARCHAR(10),Convert(varchar(10),'" & DTDate.Value.ToString & "'),113))")
        '        If SALESORDERLIST <> "" Then
        '            If Len(SALESORDERLIST) >= 1 Then
        '                MsgBox("Already one sales order is Active for item code ." & rsSoDtl.GetValue("item_code") & vbCrLf & " Sales Order Details : " & SALESORDERLIST, MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100))
        '                ValidAuthorize = False
        '                Exit Function
        '            End If
        '        End If
        '        rsSoDtl.MoveNext()
        '    End While
        'End If
        'rsSoDtl.ResultSetClose()
        If chkOpenSo.CheckState = CheckState.Checked Then 'checked true if open sales order
            STROPENCLOSEDSO = 1
        Else
            STROPENCLOSEDSO = 0
        End If

        rsSoDtl.GetResult("SELECT * FROM CUST_ORD_DTL Where Cust_ref = '" & txtReferenceNo.Text.Trim & "' and UNIT_CODE = '" & gstrUNITID & "' and Account_Code = '" & txtCustomerCode.Text.Trim & "' and Active_Flag='A' and amendment_no ='" & txtAmendmentNo.Text.Trim & "'")
        If rsSoDtl.GetNoRows > 0 Then
            rsSoDtl.MoveFirst()
            While Not rsSoDtl.EOFRecord
                'SALESORDERLIST = Find_Value("Select dbo.UDF_CHECK_NOOF_OPEN_CLOSED_SALESORDER('" & gstrUNITID & "','" & txtCustomerCode.Text & "','" & rsSoDtl.GetValue("Cust_ref") & "','" & rsSoDtl.GetValue("Cust_DrgNo") & "','" & rsSoDtl.GetValue("item_code") & "','" & STROPENCLOSEDSO & "')")
                SALESORDERLIST = Find_Value("Select dbo.UDF_CHECK_NOOF_OPEN_CLOSED_SALESORDER('" & gstrUNITID & "','" & txtCustomerCode.Text & "','" & txtConsCode.Text & "','" & rsSoDtl.GetValue("Cust_ref") & "','" & rsSoDtl.GetValue("Cust_DrgNo") & "','" & rsSoDtl.GetValue("item_code") & "','" & STROPENCLOSEDSO & "')")
                If SALESORDERLIST <> "" Then
                    If Len(SALESORDERLIST) >= 1 Then
                        If STROPENCLOSEDSO = "0" Then
                            MsgBox("Already Closed sales order is Active for item code ." & rsSoDtl.GetValue("item_code") & vbCrLf & " Sales Order No : " & SALESORDERLIST, MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100))
                        Else
                            MsgBox("Already Open sales order is Active for item code ." & rsSoDtl.GetValue("item_code") & vbCrLf & " Sales Order No: " & SALESORDERLIST, MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100))

                        End If

                        ValidAuthorize = False
                        Exit Function
                    End If
                End If
                rsSoDtl.MoveNext()
            End While
        End If
        rsSoDtl.ResultSetClose()


        Call txtCustomerCode_Validating(txtCustomerCode, New System.ComponentModel.CancelEventArgs(False))
        If blnValidCust = True Then
            Call txtReferenceNo_LostFocus(txtReferenceNo, New System.EventArgs())
            If blnValidref = True Then
                If Len(Trim(txtAmendmentNo.Text)) <> 0 Then
                    Call txtAmendmentNo_Validating(txtAmendmentNo, New System.ComponentModel.CancelEventArgs(False))
                    If blnValidAmend = True Then
                        ValidAuthorize = True
                    Else
                        Exit Function
                    End If
                Else
                    ValidAuthorize = True
                End If
            Else
                Exit Function
            End If
        Else
            Exit Function
        End If
    End Function

    Private Sub txtSTax_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSTax.TextChanged
        Call FillLabel("STAX")
    End Sub

    Public Sub UpdateHdrActiveFlag()
        On Error GoTo ErrHandler
        Dim rsCustOrdHdr As ClsResultSetDB
        Dim rsCustOrdDtl As ClsResultSetDB
        Dim intMaxLoop As Short
        Dim intLoopCounter As Short
        Dim intmaxitems As Short
        Dim intMaxOverItem As Short
        Dim AmendmentNo As String


        ''Changes done by Ashutosh on 16 Apr 2007 , Issue Id:19731
        m_strSql = "select distinct(AmendMEnt_No)from cust_ord_hdr where unit_code='" & gstrUNITID & "' and Cust_ref='" & txtReferenceNo.Text & "' and account_Code='" & txtCustomerCode.Text & "' and Consignee_Code='" & Trim(txtConsCode.Text) & "' and active_Flag='A'"
        ''Changes for Issue Id:19731 end here.
        rsCustOrdHdr = New ClsResultSetDB
        rsCustOrdHdr.GetResult(m_strSql)
        intMaxLoop = rsCustOrdHdr.GetNoRows
        rsCustOrdHdr.MoveFirst()
        For intLoopCounter = 1 To intMaxLoop
            AmendmentNo = Trim(rsCustOrdHdr.GetValue("Amendment_No"))
            m_strSql = "select 1 from cust_ord_dtl where unit_code='" & gstrUNITID & "' and Cust_ref='" & txtReferenceNo.Text & "' and account_Code='" & txtCustomerCode.Text & "' and amendment_no ='" & AmendmentNo & "' and active_Flag='O'"
            rsCustOrdDtl = New ClsResultSetDB
            rsCustOrdDtl.GetResult(m_strSql)
            intMaxOverItem = rsCustOrdDtl.GetNoRows
            rsCustOrdDtl.ResultSetClose()
            m_strSql = "select 1 from cust_ord_dtl where unit_code='" & gstrUNITID & "' and Cust_ref='" & txtReferenceNo.Text & "' and account_Code='" & txtCustomerCode.Text & "' and amendment_no ='" & AmendmentNo & "'"
            rsCustOrdDtl = New ClsResultSetDB
            rsCustOrdDtl.GetResult(m_strSql)
            intmaxitems = rsCustOrdDtl.GetNoRows
            rsCustOrdDtl.ResultSetClose()
            If intmaxitems = intMaxOverItem Then
                m_strSql = "Update cust_ord_hdr set active_Flag='O' where unit_code='" & gstrUNITID & "' and Cust_ref='" & txtReferenceNo.Text & "' and account_Code='" & txtCustomerCode.Text & "' and  Consignee_Code='" & Trim(txtConsCode.Text) & "' and active_Flag='A' and amendment_no ='" & AmendmentNo & "'"
                mP_Connection.Execute(m_strSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End If
            rsCustOrdHdr.MoveNext()
        Next
        rsCustOrdHdr.ResultSetClose()
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        rsCustOrdHdr.ResultSetClose()
        rsCustOrdDtl.ResultSetClose()
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub

    Private Sub ctlPerValue_Change(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlPerValue.Change
        With ssPOEntry

            If Len(Trim(ctlPerValue.Text)) = 0 Then ctlPerValue.Text = 1
            If Val(ctlPerValue.Text) > 1 Then
                .Row = 0
                .Col = 6
                .Text = "Rate (Per " & Val(ctlPerValue.Text) & ")"
                .Row = 0
                .Col = 7
                .Text = "Cust Supp Mat. (Per " & Val(ctlPerValue.Text) & ")"
                .Row = 0
                .Col = 8
                .Text = "Tool Cost (Per " & Val(ctlPerValue.Text) & ")"
                .Row = 0
                .Col = 11
                .Text = "Others (Per " & Val(ctlPerValue.Text) & ")"
            Else
                .Row = 0
                .Col = 6 : .Text = "Rate (Per Unit)"
                .Row = 0
                .Col = 7 : .Text = "Cust Supp Mat (Per Unit)"
                .Row = 0
                .Col = 8 : .Text = "Tool Cost (Per Unit)"
                .Row = 0
                .Col = 11 : .Text = "Others (Per Unit)"
            End If
        End With
    End Sub

    Private Sub cmdAuthorize_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.cmdGrpAuthorise.ButtonClickEventArgs) Handles cmdAuthorize.ButtonClick
        On Error GoTo ErrHandler
        Dim strSQL As String
        Dim strErrMsg As String
        Dim strAns As MsgBoxResult
        Dim enmValue As eMPowerFunctions.ConfirmWindowReturnEnum
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        Dim varItemCode As Object
        Dim varCustDrgNo As Object
        Dim dblDespatch As Double
        Dim rsDespatch As ClsResultSetDB
        Dim strRetMsg As String = String.Empty

        Select Case e.Button
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_AUTHORIZE 'Authorize PO
                If ValidAuthorize() = True Then
                    enmValue = ConfirmWindow(10163, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) 'Are You Sure To Authorize the PO
                    If enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                        If DTEffectiveDate.Value > GetServerDate() Then 'VB6.Format(ServerDate(), "dd/mm/yyyy") Then
                            Call ConfirmWindow(10198, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                        Else
                            Dim CmdRP As New ADODB.Command
                            With CmdRP
                                .CommandText = "USP_VALIDATE_RELATED_PARTY_BUDGET_VALUE_SALESORDER"
                                .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                                .ActiveConnection = mP_Connection
                                .CommandTimeout = 0

                                .Parameters.Append(.CreateParameter("@UNIT_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                                .Parameters.Append(.CreateParameter("@SO_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 25, txtReferenceNo.Text))
                                .Parameters.Append(.CreateParameter("@AMENDMENT_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 25, txtAmendmentNo.Text))
                                .Parameters.Append(.CreateParameter("@CUSTOMER_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 8, txtCustomerCode.Text))
                                .Parameters.Append(.CreateParameter("@IP_ADDRESS", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 20, mP_User))
                                .Parameters.Append(.CreateParameter("@SOURCE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 4, "SO"))
                                .Parameters.Append(.CreateParameter("@MSG_OUT", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInputOutput, 8000, ""))

                                .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                strRetMsg = Convert.ToString(.Parameters("@MSG_OUT").Value)
                                If String.IsNullOrEmpty(strRetMsg) = False Then

                                    strRetMsg += vbCrLf + "Transaction cannot save !"
                                    MsgBox(strRetMsg, MsgBoxStyle.Exclamation, "RELATED PARTY BUDGET VALIDATION")
                                    Exit Sub
                                End If

                            End With

                            CmdRP = Nothing

                            'Update Cust_ord_hdr table
                            m_strSql = "Update cust_ord_hdr set First_Authorized='"
                            m_strSql = m_strSql & mP_User & "', Second_Authorized='"
                            m_strSql = m_strSql & mP_User & "', Third_Authorized ='"
                            m_strSql = m_strSql & mP_User & "', Authorized_Flag =1 where unit_code='" & gstrUNITID & "' and "
                            m_strSql = m_strSql & " cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'"
                            m_strSql = m_strSql & " and Account_Code='"
                            ' issue Id:19731
                            m_strSql = m_strSql & Trim(Me.txtCustomerCode.Text) & "' and  Consignee_Code='" & Trim(txtConsCode.Text) & "' and "
                            ''Changes for issue id:19731 end here.
                            m_strSql = m_strSql & " amendment_No='" & Trim(txtAmendmentNo.Text) & "'"
                            mP_Connection.Execute(m_strSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            'Update Cust_ord_dtl table
                            m_strSql = "Update cust_ord_dtl set Authorized_Flag =1 where unit_code='" & gstrUNITID & "' and "
                            m_strSql = m_strSql & " cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'"
                            m_strSql = m_strSql & " and Account_Code='"
                            m_strSql = m_strSql & Trim(Me.txtCustomerCode.Text) & "' and "
                            m_strSql = m_strSql & " amendment_No='" & Trim(txtAmendmentNo.Text) & "'"
                            mP_Connection.Execute(m_strSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            '************
                            'To update despatch Qty from Previous active items on 01/07/2003
                            If ssPOEntry.MaxRows > 0 Then
                                intMaxLoop = ssPOEntry.MaxRows
                                For intLoopCounter = 1 To intMaxLoop
                                    With ssPOEntry
                                        varCustDrgNo = Nothing
                                        Call .GetText(2, intLoopCounter, varCustDrgNo)

                                        varItemCode = Nothing
                                        Call .GetText(3, intLoopCounter, varItemCode)
                                        m_strSql = "Select Despatch_Qty from cust_ord_dtl Where unit_code='" & gstrUNITID & "' and Account_code = '" & Trim(txtCustomerCode.Text) & "'"
                                        m_strSql = m_strSql & " and Cust_ref = '" & Trim(txtReferenceNo.Text) & "'  "
                                        m_strSql = m_strSql & " and amendment_no <> '" & Trim(txtAmendmentNo.Text) & "' and active_flag = 'A'"
                                        m_strSql = m_strSql & " and Authorized_flag =1 and ITem_code = '" & Trim(varItemCode) & "'"
                                        m_strSql = m_strSql & " and Cust_drgno = '" & Trim(varCustDrgNo) & "'"
                                        rsDespatch = New ClsResultSetDB
                                        rsDespatch.GetResult(m_strSql)
                                        If rsDespatch.GetNoRows > 0 Then
                                            rsDespatch.MoveFirst()
                                            dblDespatch = rsDespatch.GetValue("Despatch_Qty")
                                            rsDespatch.ResultSetClose()
                                            m_strSql = "update cust_ord_dtl set Despatch_qty = " & dblDespatch & " Where unit_code='" & gstrUNITID & "' and Account_code = '" & Trim(txtCustomerCode.Text) & "'"
                                            m_strSql = m_strSql & " and Cust_ref = '" & Trim(txtReferenceNo.Text) & "'  "
                                            m_strSql = m_strSql & " and amendment_no = '" & Trim(txtAmendmentNo.Text) & "' and active_flag = 'A'"
                                            m_strSql = m_strSql & " and ITem_code = '" & Trim(varItemCode) & "' and Cust_drgno = '" & Trim(varCustDrgNo) & "'"
                                            mP_Connection.Execute(m_strSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                        End If
                                    End With
                                Next
                            End If
                            '************
                            m_strSql = "Update cust_ord_dtl set Active_Flag = 'O' where unit_code='" & gstrUNITID & "' and cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'"
                            m_strSql = m_strSql & " and Account_Code = '" & Trim(Me.txtCustomerCode.Text) & "' and  amendment_No <> '" & Trim(Me.txtAmendmentNo.Text) & "' and authorized_flag =1 "
                            m_strSql = m_strSql & " and cust_drgno in(select cust_drgno from   cust_ord_dtl  where cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "' AND UNIT_CODE = '" & gstrUNITID & "' and Account_Code= '" & Trim(Me.txtCustomerCode.Text) & "' and  amendment_No='" & Trim(Me.txtAmendmentNo.Text) & "')"
                            mP_Connection.Execute(m_strSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            'function added by nisha to add upadetion cust_ord_hdr ActiveFlag on 01/07/2003
                            Call UpdateHdrActiveFlag()
                            'changes Ends here
                        End If
                        Call SendMail()
                        Call MsgBox("Sales Order Authorized successfully", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "eMPro")
                        Call EnableControls(False, Me, True)
                        txtCustomerCode.Enabled = True
                        cmdHelp(0).Enabled = True

                        ssPOEntry.MaxRows = 0
                        Me.txtCustomerCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        Call RefreshForm()
                        txtCustomerCode.Focus()
                    Else
                        If txtAmendmentNo.Enabled = True Then
                            txtAmendmentNo.Focus()
                            Exit Sub
                        Else
                            txtReferenceNo.Focus()
                            Exit Sub
                        End If
                    End If
                Else
                    Exit Sub
                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_REFRESH 'Refresh Screen
                Call EnableControls(False, Me, True)
                ssPOEntry.MaxRows = 0
                cmdHelp(0).Enabled = True
                txtCustomerCode.Enabled = True
                txtCustomerCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                cmdHelp(0).Enabled = True
                lblCustDesc.Text = ""
                txtCustomerCode.Focus()
                cmdAuthorize.Enabled(0) = False
                cmdAuthorize.Enabled(1) = False
                cmdAuthorize.Enabled(2) = False
                cmdAuthorize.Enabled(3) = True
                btnPOPreview.Visible = False
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE 'Close
                Me.Close()

        End Select
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Public Sub SendMail()
       

        Dim oSqlCommand As New SqlCommand
        Dim ds As New DataSet
        Dim dt As DataTable
        Dim dtmailserver As DataTable
        Dim dtsendto As DataTable
        Dim dtccto As DataTable
        Dim sqlda As New SqlDataAdapter
        Dim strUnitDesc As String
        Dim strGroup_Id As String
        Dim strFormDesc As String
        Dim rcount As Integer
        Dim dtMail As New DataTable
        Dim SEND_TO, CC_TO, BCC_TO, SUBJECT, MESSAGE, BODY_FORMAT As String
        Dim StrAttachedFiles() As String
        Dim strFileName As String
        Try

            dtMail = SqlConnectionclass.GetDataTable("SELECT top 1 SEND_TO,CCTO,MailSubject FROM tbl_AutoGeneratedReports With(NoLock) WHERE UNIT_CODE='" & gstrUNITID & "' And Report_Code='-SODocument'")
            If dtMail.Rows.Count > 0 Then
                SEND_TO = dtMail.Rows(0)("SEND_TO")
                CC_TO = dtMail.Rows(0)("CCTO")
                SUBJECT = dtMail.Rows(0)("MailSubject")
            Else
                MsgBox("Document Authorization successfully!,Auto Mail Receiver for Sales Order Authorization Not defined.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            End If
            dtMail.Dispose()

            dtmailserver = SqlConnectionclass.GetDataTable("SELECT HOSTNAME,PORT_NO,MAILADDRESS FROM MAILSERVER_CFG WHERE UNIT_CODE='" & gstrUNITID & "'")
            Dim Strmailserver As String = ""
            If dtmailserver.Rows.Count > 0 Then
                Strmailserver = dtmailserver.Rows(0).Item("MAILADDRESS").ToString()
                If Strmailserver = "" Then
                    Exit Sub
                End If
            Else
                Exit Sub
            End If

            Dim msg As System.Net.Mail.MailMessage = New System.Net.Mail.MailMessage()
            msg.Subject = SUBJECT

            Dim strSendto() As String = SEND_TO.Split(";")
            For Each MailId As String In strSendto
                If MailId.Length > 0 Then msg.To.Add(New System.Net.Mail.MailAddress(MailId))
            Next

            Dim strCCTo() As String = CC_TO.Split(";")
            For Each MailId As String In strCCTo
                If MailId.Length > 0 Then msg.CC.Add(New System.Net.Mail.MailAddress(MailId))
            Next

            Dim fromadx As New System.Net.Mail.MailAddress(Strmailserver, "eMPro")
            msg.From = fromadx


            Dim MainBody As String = String.Empty
            MainBody = MainBody + "Dear Concern," + vbNewLine + vbNewLine
            MainBody = MainBody + vbTab + vbTab + vbTab + vbTab + ":: SO Authorization MAIL NOTIFICATION:: " + vbTab + vbTab
            MainBody = MainBody + vbNewLine + vbNewLine
            MainBody = MainBody + "REQ. BY: " + GetUserNameByEmployeeCodeAndUnit(mP_User, gstrUNITID) + "" + vbTab + vbTab + "DEPT: " + gstrUNITID.ToString() + ""
            MainBody = MainBody + vbNewLine + "Sale Order Ref #: " + txtReferenceNo.Text.Trim + " IS AUTHORIZED."
            MainBody = MainBody + vbNewLine + vbNewLine + "AUTH. DATE:" + DateTime.Now.ToString("dd/MM/yyyy hh:mm tt") + "" + vbTab + vbTab + vbTab + "AUTH. BY   :" + GetUserNameByEmployeeCodeAndUnit(mP_User, gstrUNITID) + " " + vbNewLine + vbNewLine + vbNewLine
            MainBody = MainBody + "Regards," + vbNewLine + "eMProADMIN"
            MESSAGE = MainBody
            msg.Body = MESSAGE
            Dim smtpClient As System.Net.Mail.SmtpClient = New System.Net.Mail.SmtpClient(dtmailserver.Rows(0).Item("HOSTNAME").ToString(), dtmailserver.Rows(0).Item("PORT_NO"))
            smtpClient.Send(msg)

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            If Not IsNothing(oSqlCommand) Then
                oSqlCommand = Nothing
            End If
            If Not IsNothing(ds) Then
                ds = Nothing
            End If
            dt = Nothing
            dtmailserver = Nothing
            dtsendto = Nothing
            dtccto = Nothing

        End Try
    End Sub

    Private Sub cmdAuthorize_MouseDown(ByVal Sender As Object, ByVal e As UCActXCtl.cmdGrpAuthorise.MouseDownEventArgs) Handles cmdAuthorize.MouseDown
        Select Case e.Index
            Case 3
                m_CloseButton = True
        End Select
    End Sub

    Private Sub ctlHeader_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlHeader.Click
        On Error GoTo ErrHandler
        Call ShowHelp("HLP" & Mid(Me.Name, 4, Len(Me.Name)) & ".htm") '("HLPCSTMS0001.htm")
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
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

    Private Sub btnPOPreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPOPreview.Click
        If txtReferenceNo.Text.Trim() = "" Then
            Exit Sub
        End If
        Dim strSQL As String = String.Empty
        Dim sqlRdr As SqlDataReader = Nothing
        Dim fileBytes() As Byte
        Dim strFileName As String = String.Empty
        Try
            If txtAmendmentNo.Text.Trim <> "" Then
                strSQL = "SELECT DOC_NAME,MSMEDOC FROM Cust_Ord_DOC_Dtl WHERE UNIT_CODE ='" & gstrUNITID & "' AND Account_Code ='" & txtCustomerCode.Text.Trim & "' AND Cust_Ref='" & txtReferenceNo.Text.Trim & "' AND ISACTIVE=1 AND Amendment_No='" & txtAmendmentNo.Text.Trim & "'"

            Else
                strSQL = "SELECT DOC_NAME,MSMEDOC FROM Cust_Ord_DOC_Dtl WHERE UNIT_CODE ='" & gstrUNITID & "' AND Account_Code ='" & txtCustomerCode.Text.Trim & "' AND Cust_Ref='" & txtReferenceNo.Text.Trim & "' AND ISACTIVE=1"

            End If
            sqlRdr = SqlConnectionclass.ExecuteReader(strSQL)
            If sqlRdr.HasRows = True Then
                sqlRdr.Read()
                If sqlRdr("MSMEDOC").Equals(DBNull.Value) = True Then
                    MsgBox("No Attachment Found.", MsgBoxStyle.Information, ResolveResString(100))
                    sqlRdr.Close()

                    Exit Sub
                End If


                strFileName = Path.GetTempPath() & sqlRdr("DOC_NAME").ToString()
                fileBytes = sqlRdr("MSMEDOC")
                MakeFileFromBytes(strFileName, fileBytes)
                If sqlRdr("DOC_NAME").ToString().Trim.Length > 0 Then
                    Try
                        System.Diagnostics.Process.Start(strFileName)
                    Catch ex As Exception
                        RaiseException(ex)
                    End Try
                End If
            Else
                MsgBox("No Attachment Found.", MsgBoxStyle.Information, ResolveResString(100))
            End If
            sqlRdr.Close()

        Catch ex As Exception
            RaiseException(ex)
        Finally
            If IsNothing(sqlRdr) = False AndAlso sqlRdr.IsClosed = False Then sqlRdr.Close()
            sqlRdr = Nothing
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

End Class