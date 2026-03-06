Option Strict Off
Option Explicit On
Imports System.Data.io
Imports System.Data.SqlClient
Friend Class frmMKTTRN0002
	Inherits System.Windows.Forms.Form
	'---------------------------------------------------------------------------
	'(C) 2001 MIND, All rights reserved
	'
	'File Name          :   frmMKTTRN0002.frm
	'Function           :   Customer PO Authorization
	'Created By         :   Meenu Gupta
	'Created on         :   28,Mar 2001
	'Revision History   :
	'Revised date       :   24 Aug 2001
	'Revised By         :   Nisha Rai
	'17-01-2002         :   By Nisha internal issue log no 55-checked out from no =4008
	'25-01-2002         :   By nisha done to allow 4 decimal places in Rate for MSSL-ED -
	'                       checked out form no =4017
	'28-02-2002         :   By Nisha Changed lable Surcharge % on formNo 4052 in grid control
	'13/09/2002         :   By Nisha changed by nisha for accounts Plugin
	'****Field Added by Ajay on 21/07/2003
	'1.Surcharge on S.Tax
	'07-09-2003         :   1.Changed by Nisha To correct The Excise Duty Display in
	'                         GRID in Case of New SO Entry
	'---------------------------------------------------------------------------
	'08/07/2004 order by clause Added by Arshad Ali when selecting data from cust_ord_dtl
	'to fill grid(order by cust_drgNo)
	'---------------------------------------------------------------------------
	'---------------------------------------------------------------------------
	'17/12/2004 changes done by NR for future SO updations and To Show Remarks Column
	'---------------------------------------------------------------------------\
	
	
	'Revised By        -    Amit Kumar
	'Revision Date     -    13 Jan 2006
	'Issue Id          -    PRJ-2004-04-003-16804  One check is needed in Sales Order Authorization Screen & Invoice Entry Screen (in SO selection dialog box). In these screen only open SO no should be display. So please put this check into next exe.
	'Revision History  -    Criteria For Population The Reference Number Has Been Modified So That Records Having Active (Flag Lock(L) And Over(O)) Will Not Come.
	'Revised By        -    Jogender
	'Revision Date     -    24/04/2006
	'Issue Id          -    PRJ-2004-04-003-17432  changes Active_Flag <> 'L' while updating SO Active Flag to 'O' added against Issue ID 17432 of Auto SO Locking
	'Revised By        -    Jogender
	'Revision Date     -    31/05/2006
	'Issue Id          -    PRJ-2004-04-003-17975  MRP & Abatment fields added
	'Revised By        -    Jogender
	'Revision Date     -    06/06/2006
	'Issue Id          -    PRJ-2004-04-003-18021  Accessible Rate column added
	
    'Revised By        -    Manoj Vaish
    'Revision Date     -    06 Aug 2008
    'Issue ID          -    eMpro-20080805-20745
    'Revision History  -    Rectifcation of .Net Conversion Issues
    '-----------------------------------------------------------------------------
    'Revised By        -    Manoj Vaish
    'Revision Date     -    21 Jul 2009
    'Issue ID          -    eMpro-20090720-33879
    'Revision History  -    Addition of Additional VAT 
    '-----------------------------------------------------------------------------
    'Revised By        -    Vinod Singh
    'Revision Date     -    04/05/2011
    'Revision History  -    Changes for Multi Unit
    '-----------------------------------------------------------------------------
    'Revised By        -    Vinod Singh
    'Revision Date     -    29/06/2012
    'Revision History  -    Authorized SO was appearing in SO Help
    'Issue Id          -    10243325 
    '-----------------------------------------------------------------------------
    'REVISED BY        -    PRASHANT RAJPAL
    'REVISION DATE     -    18-APR-2013
    'REVISION HISTORY  -    NEW FUNCTIONALITY -ONE SALES ORDER ACTIVE FOR ONE CUSTOMER ,CONFIGURABLE FUNCTIONALITY
    'ISSUE ID          -    10375521  
    '--------------------------------------------------------------------------------------
    'REVISED BY         :   PRASHANT RAJPAL
    'ISSUE ID           :   10439883
    'DESCRIPTION        :   CHANGES ON DISCOUNT INVOICE FUNCTIONLITY 
    'REVISION DATE      :    25 -AUG 2013-03 SEP 2013
    '***********************************************************************************************
    'REVISED BY        -    Parveen Kumar
    'REVISION DATE     -    23-Dec-2014
    'REVISION HISTORY  -    Sale order Authorization Form - Cess on ED & Surch. OnRejection Report.
    'ISSUE ID          -    10552513  
    '************************************************************************************************
    'REVISED BY         :   Abhinav Kumar
    'ISSUE ID           :   10736222
    'DESCRIPTION        :   CHANGES DONE FOR CT2 ARE 3 FUNCTIONALITY - TO ENABLE CT2 REQD FLAG
    'REVISION DATE      :   09-JAN-2015
    '************************************************************************************************
    'REVISED BY         :   Abhinav Kumar
    'ISSUE ID           :   10869290 
    'DESCRIPTION        :   eMPro- Service Invoice Functionality
    'REVISION DATE      :   21 Sep 2015
    '*************************************************************************************************
    'REVISED BY         :   Abhinav Kumar
    'ISSUE ID           :   10830216 
    'DESCRIPTION        :   Customer Item rate shows many active Sales order
    'REVISION DATE      :   26 Oct 2015
    '*************************************************************************************************
    'REVISED BY         :   Parveen Kumar
    'ISSUE ID           :   10844039
    'DESCRIPTION        :   Change in SO Authorization. 
    'REVISION DATE      :   18-May-2016
    '***************************************************************************************

    Dim rsdb As ClsResultSetDB
	Dim m_CloseButton As Boolean
	Dim m_blnhelp As Boolean
	Dim mintFormIndex, intRow As Short
	Dim m_strSql, strSql As String
    Dim rsRefNo As ClsResultSetDB
	Dim m_ItemDesc, m_custItemDesc As String
	Dim strpotype As String
	Dim blnValidAmend As Boolean
	Dim blnValidCust As Boolean
    Dim blnValidref As Boolean
    Dim mblnDiscountFunctionality As Boolean
	
	Private Sub cmbPOType_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles cmbPOType.Validating
		Dim Cancel As Boolean = eventArgs.Cancel
		On Error GoTo ErrHandler
		If cmbPOType.Text = "MRP-SPARES" Then

			ssPOEntry.BlockMode = True

            ssPOEntry.Col = 13

            ssPOEntry.Col2 = 13

            ssPOEntry.ColHidden = False

            ssPOEntry.Col = 14

            ssPOEntry.Col2 = 14

            ssPOEntry.ColHidden = False

            ssPOEntry.BlockMode = False

            ssPOEntry.Col = 15

            ssPOEntry.Col2 = 15

            ssPOEntry.ColHidden = False

            ssPOEntry.BlockMode = False
        Else

            ssPOEntry.BlockMode = True

            ssPOEntry.Col = 13

            ssPOEntry.Col2 = 13

            ssPOEntry.ColHidden = True

            ssPOEntry.Col = 14

            ssPOEntry.Col2 = 14

            ssPOEntry.ColHidden = True

            ssPOEntry.Col = 15

            ssPOEntry.Col2 = 15

            ssPOEntry.ColHidden = True

            ssPOEntry.BlockMode = False
        End If
        GoTo EventExitSub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        GoTo EventExitSub

EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub

    ''Private Sub cmdAuthorize_ButtonClick(ByVal eventSender As System.Object, ByVal eventArgs As AxActXCtl.__cmdGrpAuthorise_ButtonClickEvent)
    ''End Sub
    ''Private Sub cmdAuthorize_MouseDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxActXCtl.__cmdGrpAuthorise_MouseDownEvent)
    ''End Sub
    Private Sub cmdchangetype_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdchangetype.Click
        '-----------------------------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/03/2001
        'Arguments           - None
        'Return Value        - None
        'Function            - To Display Sub Form for Payment Terms
        '-----------------------------------------------------------------------
        m_pstrSql = "select * from cust_ord_hdr where unit_code='" & gstrUNITID & "' and  Cust_ref='" & txtReferenceNo.Text & "' and amendment_No='" & txtAmendmentNo.Text & "' AND ACCOUNT_CODE='" & txtCustomerCode.Text & "'"
        frmMKTTRNAdditionalDetails.ShowDialog()
    End Sub
    Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
        Dim Index As Short = cmdHelp.GetIndex(eventSender)
        '-----------------------------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/03/2001
        'Arguments           - None
        'Return Value        - None
        'Function            - To Show the Help Form
        '-----------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim varRetVal As Object
        Dim strsql As String = String.Empty
        Dim VarRef As Object = Nothing
        Select Case Index
            Case 0
                With Me.txtCustomerCode
                    If Len(.Text) = 0 Then
                        varRetVal = ShowList(1, .MaxLength, "", "Customer_Code", "Cust_Name", "Customer_mst", " and ((isnull(deactive_flag,0) <> 1) OR (convert(varchar(12),getdate(),106)<= convert(varchar(12),deactive_date,106)))")
                        If varRetVal = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                        Else
                            .Text = varRetVal
                        End If
                    Else
                        varRetVal = ShowList(1, .MaxLength, .Text, "Customer_Code", "Cust_Name", "Customer_mst", " and ((isnull(deactive_flag,0) <> 1) OR (convert(varchar(12),getdate(),106)<= convert(varchar(12),deactive_date,106)))")
                        If varRetVal = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                        Else
                            .Text = varRetVal
                        End If
                    End If
                    .Focus()
                End With
            Case 1
                With Me.txtReferenceNo
                    If Len(.Text) = 0 Then
                        varRetVal = ShowList(1, .MaxLength, "", "Cust_Ref", DateColumnNameInShowList("order_date") & "As order_date", "cust_ord_hdr A", "and Account_Code='" & Trim(txtCustomerCode.Text) & "' and future_so = 0 and (authorized_flag = 0 or authorized_flag is null) and Rtrim(Active_Flag) Not In('L','O') AND REVISIONNO = (SELECT MAX(REVISIONNO) FROM CUST_ORD_HDR  WHERE UNIT_CODE=A.UNIT_CODE AND ACCOUNT_CODE=A.ACCOUNT_CODE AND CUST_REF = A.CUST_REF) ")

                        If varRetVal = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                        Else
                            .Text = varRetVal
                        End If
                    Else
                        varRetVal = ShowList(1, .MaxLength, .Text, "Cust_Ref", DateColumnNameInShowList("order_date") & "As order_date", "cust_ord_hdr A", "and Account_Code='" & Trim(txtCustomerCode.Text) & "' and future_so = 0 and (authorized_flag = 0 or future_so = 0) And Rtrim(Active_Flag) In ('L','O') AND REVISIONNO = (SELECT MAX(REVISIONNO) FROM CUST_ORD_HDR  WHERE UNIT_CODE=A.UNIT_CODE AND ACCOUNT_CODE=A.ACCOUNT_CODE AND CUST_REF = A.CUST_REF)")

                        If varRetVal = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                        Else
                            .Text = varRetVal
                        End If
                    End If
                    .Enabled = True
                    .Focus()
                End With
            Case 2
                With txtCurrencyType
                    If Len(.Text) = 0 Then
                        varRetVal = ShowList(1, .MaxLength, "", "Currency_Code", "Description", "Currency_mst")

                        If varRetVal = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                        Else

                            .Text = varRetVal
                        End If
                    Else
                        varRetVal = ShowList(1, .MaxLength, .Text, "Currency_Code", "Description", "Currency_mst")
                        If varRetVal = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                        Else
                            .Text = varRetVal
                        End If
                    End If
                    .Focus()
                End With
            Case 3
                With txtAmendmentNo
                    If Len(.Text) = 0 Then

                        varRetVal = ShowList(1, .MaxLength, "", "Amendment_No", "Cust_Ref", "cust_ord_hdr", "and cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'and Account_Code='" & txtCustomerCode.Text & "' and Amendment_No <>' ' and future_so = 0 and (authorized_flag = 0 or authorized_flag is null) and isnull(Active_Flag,'') Not In('L','O') ")

                        If varRetVal = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                        Else
                            .Text = varRetVal
                        End If
                    Else

                        varRetVal = ShowList(1, .MaxLength, .Text, "Amendment_No", "Cust_Ref", "cust_ord_hdr", "and  cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'and Account_Code='" & txtCustomerCode.Text & "'  and Amendment_No <>' ' and future_so = 0 and (authorized_flag = 0 or authorized_flag is null) And isnull(Active_Flag,'') Not In('L','O') ")

                        If varRetVal = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
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
        '-----------------------------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/03/2001
        'Arguments           - None
        'Return Value        - None
        'Function            - To Set Value of Public Variable
        '-----------------------------------------------------------------------
        Select Case Index
            Case 0
                m_blnhelp = True
        End Select
    End Sub

    Private Sub frmMKTTRN0002_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        '-----------------------------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/03/2001
        'Arguments           - None
        'Return Value        - None
        'Function            - To Add Required code on Form Deactivate
        '-----------------------------------------------------------------------
        On Error GoTo ErrHandler
        'Make the node normal font
        frmModules.NodeFontBold(Tag) = False
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0002_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '-----------------------------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/03/2001
        'Arguments           - None
        'Return Value        - None
        'Function            - To Call empower help on F4 Click
        '-----------------------------------------------------------------------
        If KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0 Then Call ctlHeader_Click(ctlHeader, New System.EventArgs()) : Exit Sub
    End Sub
    Private Sub txtAmendmentNo_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAmendmentNo.TextChanged
        '-----------------------------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/03/2001
        'Arguments           - None
        'Return Value        - None
        'Function            - To Refresh Values on Change in Amendment Text Box
        '-----------------------------------------------------------------------
        txtSTax.Text = "" : txtCreditTerms.Text = "" : txtSChSTax.Text = ""
        chkOpenSo.CheckState = System.Windows.Forms.CheckState.Unchecked
        If Len(Trim(txtAmendmentNo.Text)) = 0 Then
            Call RefreshForm()
        End If
    End Sub
    Private Sub txtAmendmentNo_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtAmendmentNo.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '-----------------------------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/03/2001
        'Arguments           - None
        'Return Value        - None
        'Function            - To Call Help on F1 Click
        '-----------------------------------------------------------------------
        If KeyCode = 112 Then Call cmdHelp_Click(cmdHelp.Item(3), New System.EventArgs())
        If KeyCode = System.Windows.Forms.Keys.Return Then
            Call txtAmendmentNo_Validating(txtAmendmentNo, New System.ComponentModel.CancelEventArgs(False))
            If blnValidAmend = True Then
                If cmdchangetype.Enabled Then cmdchangetype.Focus() Else cmdAuthorize.Focus()
            End If
        End If
    End Sub

    Private Sub txtCreditTerms_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCreditTerms.TextChanged
        '-----------------------------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/03/2001
        'Arguments           - None
        'Return Value        - None
        'Function            - To Refresh Label of Credit Description
        '-----------------------------------------------------------------------
        Call FillLabel("CREDIT")
    End Sub
    Private Sub txtCurrencyType_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCurrencyType.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '-----------------------------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/03/2001
        'Arguments           - None
        'Return Value        - None
        'Function            - To Call Help on F1 Click
        '-----------------------------------------------------------------------
        If KeyCode = 112 Then Call cmdHelp_Click(cmdHelp.Item(2), New System.EventArgs()) 'Help should be invoked if F1 key is pressed
    End Sub

    Private Sub txtCustomerCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCustomerCode.TextChanged
        '-----------------------------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/03/2001
        'Arguments           - None
        'Return Value        - None
        'Function            - To Refresh The Data on Change of Customer Code
        '-----------------------------------------------------------------------
        On Error GoTo ErrHandler
        Call FillLabel("CUSTOMER")
        If Len(Trim(txtCustomerCode.Text)) <> 0 Then
            txtReferenceNo.Enabled = True
            txtReferenceNo.BackColor = System.Drawing.Color.White
            cmdHelp(1).Enabled = True
            ''ADDED BY SUMIT KUMAR ON 17 JULY 2019 FOR CHECK EXTERNAL SO FLAG
            Call HideExternalSo()
        Else
            Call RefreshForm()
            lblCustDesc.Text = ""
            txtReferenceNo.Text = ""
            cmdHelp(1).Enabled = False
            cmdHelp(3).Enabled = False
            txtReferenceNo.Enabled = False
            txtReferenceNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            txtAmendmentNo.Text = ""
            txtAmendmentNo.Enabled = False
            txtReferenceNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        End If

        ssPOEntry.MaxRows = 0
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub txtCustomerCode_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCustomerCode.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '-----------------------------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/03/2001
        'Arguments           - None
        'Return Value        - None
        'Function            - To Add Code to Validate Entered Customer Code
        '-----------------------------------------------------------------------
        If KeyCode = 112 Then Call cmdHelp_Click(cmdHelp.Item(0), New System.EventArgs()) 'Help should be invoked if F1 key is pressed
        If KeyCode = System.Windows.Forms.Keys.Return Then
            Call txtCustomerCode_Validating(txtCustomerCode, New System.ComponentModel.CancelEventArgs(False))
            If blnValidCust = True Then
                If txtReferenceNo.Enabled = True Then txtReferenceNo.Focus() Else cmdAuthorize.Focus()
            End If
        End If
    End Sub
    Private Sub txtCustomerCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtCustomerCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        '-----------------------------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/03/2001
        'Arguments           - None
        'Return Value        - None
        'Function            - To Validate Customer Code Entered by User
        '-----------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim rsCD As ClsResultSetDB
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
            rsCD = New ClsResultSetDB
            m_strSql = "Select top 1 1 from Customer_Mst where  unit_code='" & gstrUNITID & "' and  Customer_Code='" & Trim(txtCustomerCode.Text) & "' and ((isnull(deactive_flag,0) <> 1) OR (convert(varchar(12),getdate(),106)<= convert(varchar(12),deactive_date,106)))"
            rsCD.GetResult(m_strSql)
            If rsCD.GetNoRows = 0 Then
                'MsgBox ("Customer Code Does Not Exist")
                Call ConfirmWindow(10145, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                Cancel = True
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
                .Col = 24
                .Col2 = 24
                .ColHidden = False

            End With
        Else
            With ssPOEntry
                .Col = 24
                .Col2 = 24
                .ColHidden = True

            End With
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub

    ''Private Sub txtReferenceNo_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtReferenceNo.TextChanged
    ''   End Sub
    Private Sub TxtReferenceNo_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtReferenceNo.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '-----------------------------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/03/2001
        'Arguments           - None
        'Return Value        - None
        'Function            - To Call Help on F1 Click
        '-----------------------------------------------------------------------
        If KeyCode = 112 Then Call cmdHelp_Click(cmdHelp.Item(1), New System.EventArgs())
        If KeyCode = System.Windows.Forms.Keys.Return Then
            If txtAmendmentNo.Enabled = True Then txtAmendmentNo.Focus() Else cmdAuthorize.Focus()
        End If
    End Sub

    Private Sub frmMKTTRN0002_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        '-----------------------------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/03/2001
        'Arguments           - None
        'Return Value        - None
        'Function            - To Add Code on Form Activate
        '-----------------------------------------------------------------------
        On Error GoTo ErrHandler
        'Checking the form name in the Windows list
        mdifrmMain.CheckFormName = mintFormIndex
        frmModules.NodeFontBold(Tag) = True
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0002_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        '-----------------------------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/03/2001
        'Arguments           - None
        'Return Value        - None
        'Function            - To Add Code on Form Load
        '-----------------------------------------------------------------------
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
        '10736222-changes done for CT2 ARE 3 functionality - to enable CT2 Reqd flag
        ChkCT2reqd.Enabled = False
        ChkCT2reqd.Checked = False
        cmdAuthorize.Enabled(3) = True
        cmdHelp(0).Enabled = True
        txtCustomerCode.Enabled = True
        txtCustomerCode.BackColor = System.Drawing.Color.White
        Call AddPOType()
        cmbPOType.Text = ObsoleteManagement.GetItemString(cmbPOType, 0)

        Me.DTDate.Value = GetServerDate()
        Me.DTAmendmentDate.Value = GetServerDate()
        Me.DTEffectiveDate.Value = GetServerDate()
        Me.DTValidDate.Value = GetServerDate()

        ssPOEntry.Enabled = False
        m_CloseButton = False
        mblnDiscountFunctionality = CBool(Find_Value("SELECT DISCOUNT_ON_INVOICE  FROM SALES_PARAMETER WHERE UNIT_CODE='" & gstrUNITID & "'"))
        Call InitializeSpreed()

        With ssPOEntry
            .Col = 13
            .Col2 = 13
            .ColHidden = True
            .Col = 14
            .Col2 = 14
            .ColHidden = True
            .Col = 15
            .Col2 = 15
            .ColHidden = True
            .Col = 24
            .Col2 = 24
            .ColHidden = True

        End With

        If mblnDiscountFunctionality = True Then
            With ssPOEntry
                .Col = 16
                .Col2 = 16
                .ColHidden = False
                .Col = 17
                .Col2 = 17
                .ColHidden = False
            End With
        End If

    End Sub
    Private Sub frmMKTTRN0002_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        '-----------------------------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/03/2001
        'Arguments           - None
        'Return Value        - None
        'Function            - To Add Code on Form_QueryUnload
        '-----------------------------------------------------------------------
        Dim enmValue As eMPowerFunctions.ConfirmWindowReturnEnum
        If UnloadMode >= 0 And UnloadMode <= 5 Then
        End If
        'Checking the status
        If gblnCancelUnload = True Then
            eventArgs.Cancel = True
        End If
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub frmMKTTRN0002_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '-----------------------------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/03/2001
        'Arguments           - None
        'Return Value        - None
        'Function            - To Add Code on Form_Unload
        '-----------------------------------------------------------------------
        'Releasing the form reference

        'Me = Nothing
        Me.Dispose()
        'Removing the form name from list
        mdifrmMain.RemoveFormNameFromWindowList = mintFormIndex
        'Setting the corresponding node's tag
        frmModules.NodeFontBold(Tag) = False
    End Sub
    Private Sub ssSetFocus(ByRef Row As Integer, Optional ByRef Col As Integer = 3)
        '-----------------------------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/03/2001
        'Arguments           - None
        'Return Value        - None
        'Function            - Set the focus according to row and col value
        '-----------------------------------------------------------------------
        With ssPOEntry

            .Row = Row

            .Col = Col

            .Action = 0
        End With
    End Sub
    Private Sub txtAmendmentNo_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtAmendmentNo.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        '-----------------------------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/03/2001
        'Arguments           - None
        'Return Value        - None
        'Function            - To Display Details of Amendment No on Validate
        '-----------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim inti As Short
        Dim rsAmend As ClsResultSetDB
        blnValidAmend = False
        If m_CloseButton = True Then
            m_CloseButton = False
            GoTo EventExitSub
        End If
        If Len(Trim(txtAmendmentNo.Text)) > 0 Then
            m_strSql = " Select top 1 1  from cust_ord_hdr where  unit_code='" & gstrUNITID & "' and cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'and Account_Code='" & Trim(Me.txtCustomerCode.Text) & "' and amendment_No='" & txtAmendmentNo.Text & "'"
            rsAmend = New ClsResultSetDB
            Call rsAmend.GetResult(m_strSql)
            If rsAmend.GetNoRows > 0 Then ' If there are records existing for the entered Amendment no
                rsAmend.ResultSetClose()
                ' check whether the PO is active or not
                rsAmend = New ClsResultSetDB
                m_strSql = " Select top 1 1 from cust_ord_hdr where  unit_code='" & gstrUNITID & "' and cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'and Account_Code='" & Trim(Me.txtCustomerCode.Text) & "' and Active_Flag='A' and amendment_No='" & txtAmendmentNo.Text & "'"
                Call rsAmend.GetResult(m_strSql)
                If rsAmend.GetNoRows > 0 Then ' If An amendment exists for the reference no
                    rsAmend.ResultSetClose()
                    rsAmend = New ClsResultSetDB
                    m_strSql = " Select top 1 1  from cust_ord_hdr where  unit_code='" & gstrUNITID & "' and cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'and Account_Code='" & Trim(Me.txtCustomerCode.Text) & "' and Active_Flag='A' and amendment_No='" & txtAmendmentNo.Text & "' and (Authorized_Flag= 1 or future_so =1)"
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
                        Call GetAmendmentDetails()
                        cmdAuthorize.Enabled(0) = True
                        cmdAuthorize.Enabled(1) = True
                        cmdAuthorize.Enabled(2) = False
                        cmdAuthorize.Enabled(3) = True
                        blnValidAmend = True
                        rsAmend.ResultSetClose()
                        GoTo EventExitSub
                    End If
                Else
                    'MsgBox "This Amendment Is No More Valid" ' If PO is not Valid. where active flag <> A
                    Call ConfirmWindow(10160, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                    'Call GetAmendmentDetails
                    txtAmendmentNo.Text = ""
                    Cancel = True
                    blnValidAmend = False
                    rsAmend.ResultSetClose()
                    GoTo EventExitSub
                End If
            Else
                'MsgBox "There are no existing records for the Amendment No"
                Call ConfirmWindow(10128, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                txtAmendmentNo.Text = ""
                Cancel = True
                blnValidAmend = False
                rsAmend.ResultSetClose()
                GoTo EventExitSub
            End If
        End If
        blnValidAmend = True
        GoTo EventExitSub 'This is to avoid the execution of the error handler
ErrHandler:
        rsAmend.ResultSetClose()
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        GoTo EventExitSub
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Public Function FillLabel(ByRef pstrCode As Object) As Object
        '-----------------------------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/03/2001
        'Arguments           - None
        'Return Value        - None
        'Function            - fills the customer detail label
        '-----------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim rsCust As ClsResultSetDB

        Select Case UCase(pstrCode)
            Case "CUSTOMER"
                m_strSql = "select Cust_Name,Credit_Days from Customer_mst where  unit_code='" & gstrUNITID & "' and Customer_code='" & txtCustomerCode.Text & "' and ((isnull(deactive_flag,0) <> 1) OR (convert(varchar(12),getdate(),106)<= convert(varchar(12),deactive_date,106)))"
                rsCust = New ClsResultSetDB
                rsCust.GetResult(m_strSql)
                lblCustDesc.ForeColor = System.Drawing.Color.White

                lblCustDesc.Text = IIf(UCase(rsCust.GetValue("Cust_Name")) = "UNKNOWN", "", rsCust.GetValue("Cust_Name"))
                rsCust.ResultSetClose()
            Case "STAX"
                m_strSql = "select TxRt_Rate_No,TxRt_RateDesc from Gen_TaxRate where  unit_code='" & gstrUNITID & "' and  Txrt_Rate_No = '" & txtSTax.Text & "'"
                rsCust = New ClsResultSetDB
                rsCust.GetResult(m_strSql)
                lblSTaxDesc.ForeColor = System.Drawing.Color.White

                lblSTaxDesc.Text = IIf(UCase(rsCust.GetValue("TxRt_RateDesc")) = "UNKNOWN", "", rsCust.GetValue("TxRt_RateDesc"))
                rsCust.ResultSetClose()
                '1.Surcharge on S.Tax
            Case "SSCHTAX"
                m_strSql = "SELECT TxRt_Rate_no, TxRt_RateDesc FROM gen_taxrate where  unit_code='" & gstrUNITID & "' and Txrt_Rate_No= '" & txtSChSTax.Text & "'"
                rsCust = New ClsResultSetDB
                rsCust.GetResult(m_strSql)
                lblSChSTaxDesc.ForeColor = System.Drawing.Color.White

                lblSChSTaxDesc.Text = IIf(UCase(rsCust.GetValue("TxRt_RateDesc")) = "UNKNOWN", "", rsCust.GetValue("TxRt_RateDesc"))
                rsCust.ResultSetClose()
                '------------------->>
            Case "CREDIT"
                m_strSql = "select crtrm_TermID,crtrm_Desc from Gen_CreditTrmMaster where  unit_code='" & gstrUNITID & "' and crtrm_TermID = '" & txtCreditTerms.Text & "'"
                rsCust = New ClsResultSetDB
                rsCust.GetResult(m_strSql)
                lblCreditTermDesc.ForeColor = System.Drawing.Color.White

                lblCreditTermDesc.Text = IIf(UCase(rsCust.GetValue("crtrm_Desc")) = "UNKNOWN", "", rsCust.GetValue("crtrm_Desc"))
                rsCust.ResultSetClose()
            Case "ADDVAT"
                m_strSql = "SELECT TxRt_Rate_no, TxRt_RateDesc FROM gen_taxrate where  unit_code='" & gstrUNITID & "' and Txrt_Rate_No= '" & txtAddVAT.Text & "'"
                rsCust = New ClsResultSetDB
                rsCust.GetResult(m_strSql)
                lbladdvatdesc.ForeColor = System.Drawing.Color.White

                lbladdvatdesc.Text = IIf(UCase(rsCust.GetValue("TxRt_RateDesc")) = "UNKNOWN", "", rsCust.GetValue("TxRt_RateDesc"))
                rsCust.ResultSetClose()
        End Select
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        rsCust.ResultSetClose()
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Function
    End Function
    Public Sub GetAmendmentDetails()
        '-----------------------------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/03/2001
        'Arguments           - None
        'Return Value        - None
        'Function            - Get the details if there is an amendment
        '-----------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim rsAD As ClsResultSetDB
        Dim rscurrency As ClsResultSetDB

        m_strSql = "select Account_Code,Cust_Ref,Amendment_No,Order_Date,Amendment_Date,Active_Flag,Currency_Code,Valid_Date,Effect_Date,Term_Payment,Special_Remarks,Pay_Remarks,Price_Remarks,Packing_Remarks,Frieght_Remarks,Transport_Remarks,Octorai_Remarks,Mode_Despatch,Delivery,First_Authorized,Second_Authorized,Third_Authorized,Authorized_Flag,Reason,PO_Type,SalesTax_Type,OpenSO,AddCustSupp,PerValue,InternalSONo,RevisionNo,Surcharge_code,Future_SO,ECESS_Code,Consignee_Code,ADDVAT_TYPE,CT2_REQD_IN_SO  from cust_ord_hdr where  unit_code='" & gstrUNITID & "' and Cust_ref='" & txtReferenceNo.Text & "' and amendment_No='" & txtAmendmentNo.Text & "'AND ACCOUNT_CODE='" & txtCustomerCode.Text & "' and Active_Flag='A'"
        rsAD = New ClsResultSetDB
        rsAD.GetResult(m_strSql)
        m_strSql = "select Account_Code,Cust_Ref,Amendment_No,Item_Code,Rate,Order_Qty,Despatch_Qty,Active_Flag,Cust_Mtrl,Cust_DrgNo,Packing,Others,Excise_Duty,Cust_Drg_Desc,Tool_Cost,Authorized_flag,Ent_dt,OpenSO,SalesTax_Type,PerValue,InternalSONo,RevisionNo,Remarks,MRP,Abantment_Code,AccessibleRateforMRP,Packing_Type,TOOL_AMOR_FLAG,ShowInAuth,ADD_Excise_Duty , DISCOUNT_TYPE ,DISCOUNT_VALUE ,HSNSACCODE,ISHSNORSAC,CGSTTXRT_TYPE,SGSTTXRT_TYPE,UTGSTTXRT_TYPE,IGSTTXRT_TYPE,COMPENSATION_CESS,EXTERNAL_SALESORDER_NO from cust_ord_dtl where  unit_code='" & gstrUNITID & "' and Cust_ref='" & txtReferenceNo.Text & "' and amendment_No='" & txtAmendmentNo.Text & "'AND ACCOUNT_CODE='" & txtCustomerCode.Text & "' and Active_Flag='A' order by cust_drgno"
        rsdb = New ClsResultSetDB
        rsdb.GetResult(m_strSql)

        Dim intLoopCounter As Short
        Dim intDecimal As Short
        Dim strMin As String
        Dim strMax As String
        If rsAD.GetNoRows > 0 Then
            rsAD.MoveFirst()

            Me.DTDate.Value = rsAD.GetValue("Order_Date")

            DTAmendmentDate.Value = rsAD.GetValue("Amendment_Date")

            DTEffectiveDate.Value = rsAD.GetValue("Effect_Date")

            DTValidDate.Value = rsAD.GetValue("Valid_Date")

            txtCurrencyType.Text = rsAD.GetValue("Currency_Code")
            '10552513-Starts
            txtSChSTax.Text = rsAD.GetValue("Surcharge_code")
            txtECSSTaxType.Text = rsAD.GetValue("ECESS_Code")
            '10552513-Ends

            ctlPerValue.Text = rsAD.GetValue("PerValue")

            strpotype = rsAD.GetValue("PO_Type")
            With ssPOEntry

                .Col = 13

                .Col2 = 13

                .ColHidden = True

                .Col = 14

                .Col2 = 14

                .ColHidden = True

                .Col = 15

                .Col2 = 15

                .ColHidden = True
            End With
            Select Case UCase(strpotype)
                Case "O"
                    Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 1)
                Case "S"
                    Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 2)
                Case "J"
                    Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 3)
                Case "E"
                    Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 4)
                Case "M"
                    Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 5)
                    With ssPOEntry

                        .Col = 13

                        .Col2 = 13

                        .ColHidden = False

                        .Col = 14

                        .Col2 = 14

                        .ColHidden = False

                        .Col = 15

                        .Col2 = 15

                        .ColHidden = False
                    End With
                Case "V"
                    Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 6)
            End Select
            'To Display Value in Sales Tax,CreDit Terms Open SO Flag,Add CustSupp Flag
            txtSTax.Text = rsAD.GetValue("SalesTax_Type")
            txtSChSTax.Text = rsAD.GetValue("Surcharge_code")
            txtAddVAT.Text = IIf(IsDBNull(rsAD.GetValue("AddVAT_Type")), "", rsAD.GetValue("AddVAT_Type"))
            txtCreditTerms.Text = rsAD.GetValue("term_Payment")
            '10736222-changes done for CT2 ARE 3 functionality - to enable CT2 Reqd flag
            ChkCT2reqd.Checked = rsAD.GetValue("CT2_REQD_IN_SO")

            If rsAD.GetValue("OpenSO") = False Then
                chkOpenSo.CheckState = System.Windows.Forms.CheckState.Unchecked
            Else
                chkOpenSo.CheckState = System.Windows.Forms.CheckState.Checked
            End If
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
                If Len(Trim(txtCurrencyType.Text)) Then
                    rscurrency = New ClsResultSetDB
                    rscurrency.GetResult("Select decimal_Place from Currency_Mst where unit_code='" & gstrUNITID & "' and  currency_code ='" & Trim(txtCurrencyType.Text) & "'")
                    intDecimal = rscurrency.GetValue("Decimal_Place")
                End If
                If intDecimal <= 0 Then
                    intDecimal = 2
                End If
                strMin = "0." : strMax = "99999999."
                For intLoopCounter = 1 To intDecimal
                    strMin = strMin & "0"
                    strMax = strMax & "9"
                Next

                For intLoopCounter = 6 To 8
                    With Me.ssPOEntry
                        .Row = .MaxRows
                        .Col = intLoopCounter
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat

                        .TypeFloatDecimalPlaces = intDecimal

                        .TypeFloatMax = strMax

                        .TypeFloatMin = strMin
                    End With
                Next
                With Me.ssPOEntry
                    .Row = .MaxRows
                    .Col = 9
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                    .Lock = True
                    .Row = .MaxRows
                    .Col = 13
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                    .TypeFloatDecimalPlaces = intDecimal
                    .TypeFloatMax = strMax
                    .TypeFloatMin = strMin
                    .Row = .MaxRows
                    .Col = 14
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                    .Row = .MaxRows
                    .Col = 16
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                    .Lock = True
                    .Row = .MaxRows
                    .Col = 17
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                    .Lock = True

                End With
                Call ssPOEntry.SetText(6, ssPOEntry.MaxRows, rsdb.GetValue("Rate") * CDbl(ctlPerValue.Text))
                Call ssPOEntry.SetText(7, ssPOEntry.MaxRows, rsdb.GetValue("Cust_Mtrl") * CDbl(ctlPerValue.Text))
                Call ssPOEntry.SetText(8, ssPOEntry.MaxRows, rsdb.GetValue("Tool_Cost") * CDbl(ctlPerValue.Text))

                Call ssPOEntry.SetText(9, ssPOEntry.MaxRows, rsdb.GetValue("Packing_Type"))
                Call ssPOEntry.SetText(10, ssPOEntry.MaxRows, rsdb.GetValue("Excise_Duty"))
                Call ssPOEntry.SetText(11, ssPOEntry.MaxRows, rsdb.GetValue("Others") * CDbl(ctlPerValue.Text))
                Call ssPOEntry.SetText(12, ssPOEntry.MaxRows, rsdb.GetValue("Remarks"))
                Call ssPOEntry.SetText(13, ssPOEntry.MaxRows, rsdb.GetValue("MRP"))
                Call ssPOEntry.SetText(14, ssPOEntry.MaxRows, rsdb.GetValue("Abantment_code"))
                Call ssPOEntry.SetText(15, ssPOEntry.MaxRows, rsdb.GetValue("AccessibleRateforMRP"))
                Call ssPOEntry.SetText(16, ssPOEntry.MaxRows, rsdb.GetValue("DISCOUNT_TYPE"))
                Call ssPOEntry.SetText(17, ssPOEntry.MaxRows, rsdb.GetValue("DISCOUNT_VALUE"))
                'GST DETAILS
                If gblnGSTUnit = True Then
                    Call ssPOEntry.SetText(18, ssPOEntry.MaxRows, rsdb.GetValue("HSNSACCODE"))
                    Call ssPOEntry.SetText(19, ssPOEntry.MaxRows, rsdb.GetValue("CGSTTXRT_TYPE"))
                    Call ssPOEntry.SetText(20, ssPOEntry.MaxRows, rsdb.GetValue("SGSTTXRT_TYPE"))
                    Call ssPOEntry.SetText(21, ssPOEntry.MaxRows, rsdb.GetValue("IGSTTXRT_TYPE"))
                    Call ssPOEntry.SetText(22, ssPOEntry.MaxRows, rsdb.GetValue("UTGSTTXRT_TYPE"))
                    Call ssPOEntry.SetText(23, ssPOEntry.MaxRows, rsdb.GetValue("COMPENSATION_CESS"))

                End If
                ''ADDED BY SUMIT KUMAR ON 9 JULY 2019
                Call ssPOEntry.SetText(24, ssPOEntry.MaxRows, rsdb.GetValue("EXTERNAL_SALESORDER_NO"))
                'GST DETAILS
                rsdb.MoveNext()
            Loop
        End If

        With ssPOEntry
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
        '-----------------------------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/03/2001
        'Arguments           - None
        'Return Value        - None
        'Function            - Get the details if there is no amendment
        '-----------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim rsAD As New ClsResultSetDB
        Dim rscurrency As ClsResultSetDB
        Dim strAuthFlg As String

        m_strSql = "select Account_Code,Cust_Ref,Amendment_No,Order_Date,Amendment_Date,Active_Flag,Currency_Code,Valid_Date,Effect_Date,Term_Payment,Special_Remarks,Pay_Remarks,Price_Remarks,Packing_Remarks,Frieght_Remarks,Transport_Remarks,Octorai_Remarks,Mode_Despatch,Delivery,First_Authorized,Second_Authorized,Third_Authorized,Authorized_Flag,Reason,PO_Type,SalesTax_Type,OpenSO,AddCustSupp,PerValue,InternalSONo,RevisionNo,Surcharge_code,Future_SO,ECESS_Code,Consignee_Code,ADDVAT_TYPE,CT2_REQD_IN_SO,SERVICETAX_TYPE from cust_ord_hdr where unit_code='" & gstrUNITID & "' and Cust_ref='" & txtReferenceNo.Text & "' and Account_Code='" & txtCustomerCode.Text & "' and active_Flag='A'"
        rsAD.GetResult(m_strSql)
        m_strSql = "select Account_Code,Cust_Ref,Amendment_No,Item_Code,Rate,Order_Qty,Despatch_Qty,Active_Flag,Cust_Mtrl,Cust_DrgNo,Packing,Others,Excise_Duty,Cust_Drg_Desc,Tool_Cost,Authorized_flag,Ent_dt,OpenSO,SalesTax_Type,PerValue,InternalSONo,RevisionNo,Remarks,MRP,Abantment_Code,AccessibleRateforMRP,Packing_Type,TOOL_AMOR_FLAG,ShowInAuth,ADD_Excise_Duty,discount_type,discount_value ,HSNSACCODE,ISHSNORSAC,CGSTTXRT_TYPE,SGSTTXRT_TYPE,UTGSTTXRT_TYPE,IGSTTXRT_TYPE,COMPENSATION_CESS,EXTERNAL_SALESORDER_NO from cust_ord_dtl where unit_code='" & gstrUNITID & "' and Cust_ref='" & txtReferenceNo.Text & "' and Account_Code='" & txtCustomerCode.Text & "' and active_Flag='A'"
        rsdb = New ClsResultSetDB
        rsdb.GetResult(m_strSql)
        Dim intLoopCounter As Short
        Dim intDecimal As Short
        Dim strMin As String
        Dim strMax As String
        If rsAD.GetNoRows > 0 Then
            rsAD.MoveFirst()
            Me.DTDate.Value = rsAD.GetValue("Order_Date")
            If Len(Trim(rsAD.GetValue("Amendment_No"))) > 0 Then
                DTAmendmentDate.Value = rsAD.GetValue("Amendment_Date")
            End If
            '10736222-changes done for CT2 ARE 3 functionality - to enable CT2 Reqd flag
            ChkCT2reqd.Checked = rsAD.GetValue("CT2_REQD_IN_SO")


            DTEffectiveDate.Value = rsAD.GetValue("Effect_Date")
            DTValidDate.Value = rsAD.GetValue("Valid_Date")
            txtCurrencyType.Text = rsAD.GetValue("Currency_Code")
            txtCurrencyType.Text = rsAD.GetValue("Currency_Code")
            '10552513-Starts
            txtSChSTax.Text = rsAD.GetValue("Surcharge_code")
            txtECSSTaxType.Text = rsAD.GetValue("ECESS_Code")
            '10552513-Ends
            ctlPerValue.Text = rsAD.GetValue("PerValue")
            strpotype = rsAD.GetValue("PO_Type")
            With ssPOEntry
                .Col = 13
                .Col2 = 13
                .ColHidden = True
                .Col = 14
                .Col2 = 14
                .ColHidden = True
                .Col = 15
                .Col2 = 15
                .ColHidden = True
            End With
            '10869290
            Select Case UCase(strpotype)
                Case "O"
                    Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 1)
                Case "S"
                    Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 2)
                Case "J"
                    Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 3)
                Case "E"
                    Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 4)
                Case "M"
                    Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 5)
                    With ssPOEntry
                        .Col = 13
                        .Col2 = 13
                        .ColHidden = False
                        .Col = 14
                        .Col2 = 14
                        .ColHidden = False
                        .Col = 15
                        .Col2 = 15
                        .ColHidden = False
                    End With
                Case "V"
                    Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 6)
            End Select
            'to show the details of Sales Tax,Credit Days,AddCustSupplied Flag,Open SO Flag
            txtSTax.Text = rsAD.GetValue("SalesTax_Type")
            txtSChSTax.Text = rsAD.GetValue("Surcharge_code")
            txtAddVAT.Text = IIf(IsDBNull(rsAD.GetValue("AddVAT_Type")), "", rsAD.GetValue("AddVAT_Type"))
            txtCreditTerms.Text = rsAD.GetValue("term_Payment")
            TxtServiceCode.Text = IIf(IsDBNull(rsAD.GetValue("SERVICETAX_TYPE")), "", rsAD.GetValue("SERVICETAX_TYPE"))

            If rsAD.GetValue("OpenSO") = False Then
                chkOpenSo.CheckState = System.Windows.Forms.CheckState.Unchecked
            Else
                chkOpenSo.CheckState = System.Windows.Forms.CheckState.Checked
            End If
            rsAD.MoveFirst()
            ssPOEntry.MaxRows = 0
            'Changed to add open Item Falg in Grid
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
                Call ssPOEntry.SetText(3, ssPOEntry.MaxRows, rsdb.GetValue("Item_Code"))
                Call ssPOEntry.SetText(4, ssPOEntry.MaxRows, rsdb.GetValue("Cust_Drg_Desc"))
                Call ssPOEntry.SetText(5, ssPOEntry.MaxRows, rsdb.GetValue("Order_Qty"))
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
                'Changed to add open Item Falg in Grid
                For intLoopCounter = 6 To 8
                    With Me.ssPOEntry
                        .Row = .MaxRows
                        .Col = intLoopCounter
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                        .TypeFloatDecimalPlaces = intDecimal
                        .TypeFloatMax = strMax
                        .TypeFloatMin = strMin
                    End With
                Next
                With Me.ssPOEntry
                    .Row = .MaxRows
                    .Col = 9
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                    .Lock = True
                    .Row = .MaxRows
                    .Col = 13
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                    .TypeFloatDecimalPlaces = intDecimal
                    .TypeFloatMax = strMax
                    .TypeFloatMin = strMin
                    .Lock = True
                    .Row = .MaxRows
                    .Col = 14
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                    .Lock = True
                    .Row = .MaxRows
                    .Col = 15
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                    .TypeFloatDecimalPlaces = intDecimal
                    .TypeFloatMax = strMax
                    .TypeFloatMin = strMin
                    .Lock = True


                    .Row = .MaxRows
                    .Col = 16
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                    .Lock = True

                    .Row = .MaxRows
                    .Col = 17
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                    .Lock = True

                    If gblnGSTUnit = True Then
                        .Row = .MaxRows
                        .Col = 18
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                        .Lock = True
                        .Row = .MaxRows
                        .Col = 19
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                        .Lock = True
                        .Row = .MaxRows
                        .Col = 20
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                        .Lock = True
                        .Row = .MaxRows
                        .Col = 21
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                        .Lock = True
                        .Row = .MaxRows
                        .Col = 22
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                        .Lock = True
                        .Row = .MaxRows
                        .Col = 23
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                        .Lock = True
                    End If

                    .Row = .MaxRows
                    .Col = 24
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                    .Lock = True
                End With


                Call ssPOEntry.SetText(6, ssPOEntry.MaxRows, rsdb.GetValue("Rate") * CDbl(ctlPerValue.Text))
                Call ssPOEntry.SetText(7, ssPOEntry.MaxRows, rsdb.GetValue("Cust_Mtrl") * CDbl(ctlPerValue.Text))
                Call ssPOEntry.SetText(8, ssPOEntry.MaxRows, rsdb.GetValue("Tool_Cost") * CDbl(ctlPerValue.Text))
                Call ssPOEntry.SetText(9, ssPOEntry.MaxRows, rsdb.GetValue("Packing_Type"))
                Call ssPOEntry.SetText(10, ssPOEntry.MaxRows, rsdb.GetValue("Excise_Duty"))
                Call ssPOEntry.SetText(11, ssPOEntry.MaxRows, rsdb.GetValue("Others") * CDbl(ctlPerValue.Text))
                Call ssPOEntry.SetText(12, ssPOEntry.MaxRows, rsdb.GetValue("Remarks"))
                Call ssPOEntry.SetText(13, ssPOEntry.MaxRows, rsdb.GetValue("MRP"))
                Call ssPOEntry.SetText(14, ssPOEntry.MaxRows, rsdb.GetValue("Abantment_code"))
                Call ssPOEntry.SetText(15, ssPOEntry.MaxRows, rsdb.GetValue("AccessibleRateforMRP"))
                Call ssPOEntry.SetText(16, ssPOEntry.MaxRows, rsdb.GetValue("discount_type"))
                Call ssPOEntry.SetText(17, ssPOEntry.MaxRows, rsdb.GetValue("discount_value"))
                If gblnGSTUnit = True Then
                    Call ssPOEntry.SetText(18, ssPOEntry.MaxRows, rsdb.GetValue("HSNSACCODE"))
                    Call ssPOEntry.SetText(19, ssPOEntry.MaxRows, rsdb.GetValue("CGSTTXRT_TYPE"))
                    Call ssPOEntry.SetText(20, ssPOEntry.MaxRows, rsdb.GetValue("SGSTTXRT_TYPE"))
                    Call ssPOEntry.SetText(21, ssPOEntry.MaxRows, rsdb.GetValue("IGSTTXRT_TYPE"))
                    Call ssPOEntry.SetText(22, ssPOEntry.MaxRows, rsdb.GetValue("UTGSTTXRT_TYPE"))
                    Call ssPOEntry.SetText(23, ssPOEntry.MaxRows, rsdb.GetValue("COMPENSATION_CESS"))

                End If
                ''ADDED BY SUMIT KUMAR ON 9 JULY 2019
                Call ssPOEntry.SetText(24, ssPOEntry.MaxRows, rsdb.GetValue("EXTERNAL_SALESORDER_NO"))
                rsdb.MoveNext()
            Loop
        End If
        With ssPOEntry
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
        rsdb.ResultSetClose()
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Public Sub AddPOType()
        '-----------------------------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/03/2001
        'Arguments           - None
        'Return Value        - None
        'Function            - To Add SO Type in Combo
        '-----------------------------------------------------------------------
        '10869290
        cmbPOType.Items.Insert(0, "None")
        cmbPOType.Items.Insert(1, "OEM")
        cmbPOType.Items.Insert(2, "SPARES")
        cmbPOType.Items.Insert(3, "JOB WORK")
        cmbPOType.Items.Insert(4, "Export")
        cmbPOType.Items.Insert(5, "MRP-SPARES")
        cmbPOType.Items.Insert(6, "V-SERVICE")
    End Sub
    Public Sub RefreshForm()
        '-----------------------------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/03/2001
        'Arguments           - None
        'Return Value        - None
        'Function            - To Refresh The Form
        '-----------------------------------------------------------------------
        txtCurrencyType.Text = ""
        cmbPOType.Text = ObsoleteManagement.GetItemString(cmbPOType, 0)
        Me.DTDate.Value = GetServerDate()
        Me.DTAmendmentDate.Value = GetServerDate()
        Me.DTEffectiveDate.Value = GetServerDate()
        Me.DTValidDate.Value = GetServerDate()
        cmdAuthorize.Enabled(0) = False
        cmdAuthorize.Enabled(1) = False
        cmdAuthorize.Enabled(2) = False
        cmdAuthorize.Enabled(3) = True
        '10736222-changes done for CT2 ARE 3 functionality - to enable CT2 Reqd flag
        ChkCT2reqd.Enabled = False
        ChkCT2reqd.Checked = False
        ssPOEntry.MaxRows = 0

    End Sub
    Public Function ValidAuthorize() As Boolean
        '-----------------------------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/03/2001
        'Arguments           - None
        'Return Value        - None
        'Function            - To Check Before Authorization
        '-----------------------------------------------------------------------
        Dim strActiveSalesorder As String = String.Empty
        Dim SALESORDERLIST As String = String.Empty
        Dim rsSoDtl As New ClsResultSetDB
        Dim STROPENCLOSEDSO As String
        ValidAuthorize = False

        strActiveSalesorder = Find_Value("SELECT DBO.UDF_CHECK_ACTIVE_SO_SINGLECUSTOMER('" & gstrUNITID & "','" & Me.txtCustomerCode.Text.Trim & "')")
        If strActiveSalesorder <> "" Then
            MsgBox("ALREADY ACTIVE SALES ORDER EXISTS ,KINDLY LOCK THE BELOW SALES ORDER FIRST (AMENDMENT NO ) : ." & strActiveSalesorder, MsgBoxStyle.Information, ResolveResString(100))
            ValidAuthorize = False
            Exit Function
        End If



        Call txtCustomerCode_Validating(txtCustomerCode, New System.ComponentModel.CancelEventArgs(False))

        '10830216


        'rsSoDtl.GetResult("SELECT * FROM CUST_ORD_DTL Where Cust_ref = '" & txtReferenceNo.Text.Trim & "' and UNIT_CODE = '" & gstrUNITID & "' and Account_Code = '" & txtCustomerCode.Text.Trim & "' and Active_Flag='A' and amendment_no ='" & txtAmendmentNo.Text.Trim & "'")
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
        'If CBool(Find_Value("SELECT OPENSO from cust_ord_hdr where unit_code='" & gstrUNITID & "' and account_code='" & Me.txtCustomerCode.Text.Trim & "'" & " and cust_Ref='" & txtReferenceNo.Text & "' and amendment_no='" & txtAmendmentNo.Text & "' and ")) = False Then '1 MEANS CLOSED SO

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
                SALESORDERLIST = Find_Value("Select dbo.UDF_CHECK_NOOF_OPEN_CLOSED_SALESORDER('" & gstrUNITID & "','" & txtCustomerCode.Text & "','" & txtCustomerCode.Text & "','" & rsSoDtl.GetValue("Cust_ref") & "','" & rsSoDtl.GetValue("Cust_DrgNo") & "','" & rsSoDtl.GetValue("item_code") & "','" & STROPENCLOSEDSO & "')")
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
        '-----------------------------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/03/2001
        'Arguments           - None
        'Return Value        - None
        'Function            - To Get The Label STAX
        '-----------------------------------------------------------------------
        Call FillLabel("STAX")
    End Sub

    Private Sub txtSChSTax_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSChSTax.TextChanged
        '-----------------------------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/03/2001
        'Arguments           - None
        'Return Value        - None
        'Function            - To Get The Label SSTAX
        '-----------------------------------------------------------------------
        Call FillLabel("SSCHTAX")
    End Sub
    '----------------->>
    Public Sub UpdateHdrActiveFlag()
        '-----------------------------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/03/2001
        'Arguments           - None
        'Return Value        - None
        'Function            - To update Active Flag To "O" on Authorization
        '-----------------------------------------------------------------------
        Dim rsCustOrdHdr As ClsResultSetDB
        Dim rsCustOrdDtl As ClsResultSetDB
        Dim intMaxLoop As Short
        Dim intLoopCounter As Short
        Dim intmaxitems As Short
        Dim intMaxOverItem As Short
        Dim AmendmentNo As String


        m_strSql = "select distinct(AmendMEnt_No) from cust_ord_hdr where unit_code='" & gstrUNITID & "' and Cust_ref='" & txtReferenceNo.Text & " 'and account_Code='" & txtCustomerCode.Text & "' and active_Flag='A'"
        rsCustOrdHdr = New ClsResultSetDB
        rsCustOrdHdr.GetResult(m_strSql)
        intMaxLoop = rsCustOrdHdr.GetNoRows
        rsCustOrdHdr.MoveFirst()
        For intLoopCounter = 1 To intMaxLoop

            AmendmentNo = Trim(rsCustOrdHdr.GetValue("Amendment_No"))
            m_strSql = "select 1 from cust_ord_dtl where unit_code='" & gstrUNITID & "' and Cust_ref='" & txtReferenceNo.Text & " 'and account_Code='" & txtCustomerCode.Text & "' and amendment_no ='" & AmendmentNo & "' and active_Flag='O'"
            rsCustOrdDtl = New ClsResultSetDB
            rsCustOrdDtl.GetResult(m_strSql)
            intMaxOverItem = rsCustOrdDtl.GetNoRows
            rsCustOrdDtl.ResultSetClose()

            m_strSql = "select 1 from cust_ord_dtl where unit_code='" & gstrUNITID & "' and Cust_ref='" & txtReferenceNo.Text & " 'and account_Code='" & txtCustomerCode.Text & "' and amendment_no ='" & AmendmentNo & "'"
            rsCustOrdDtl = New ClsResultSetDB
            rsCustOrdDtl.GetResult(m_strSql)
            intmaxitems = rsCustOrdDtl.GetNoRows
            rsCustOrdDtl.ResultSetClose()

            If intmaxitems = intMaxOverItem Then
                m_strSql = "Update cust_ord_hdr set active_Flag='O' where unit_code='" & gstrUNITID & "' and Cust_ref='" & txtReferenceNo.Text & " 'and account_Code='" & txtCustomerCode.Text & "'and active_Flag='A' and amendment_no ='" & AmendmentNo & "'"
                mP_Connection.Execute(m_strSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End If
            rsCustOrdHdr.MoveNext()
        Next
        rsCustOrdHdr.ResultSetClose()
    End Sub
    Private Sub InitializeSpreed()
        On Error GoTo ErrHandler
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        With Me.ssPOEntry
            If gblnGSTUnit = True Then
                .MaxCols = 24
            Else
                .MaxCols = 17
            End If
            .Row = 0
            .Col = 13 : .Text = "MRP"
            .Row = 0
            .Col = 14 : .Text = "Abantment"
            .Row = 0
            .Col = 15 : .Text = "Accessible Rate"
            .Row = 0
            .Col = 16 : .Text = "Discount Type"
            .Row = 0
            .Col = 17 : .Text = "Discount Per"

            .CursorStyle = FPSpreadADO.CursorStyleConstants.CursorStyleArrow
        End With
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Exit Sub
ErrHandler:
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub

    Private Sub ctlPerValue_Change(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlPerValue.Change
        '-----------------------------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/03/2001
        'Arguments           - None
        'Return Value        - None
        'Function            - To Change The Values Grid Labeles on Change of
        '                    - Per Value
        '-----------------------------------------------------------------------
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
        '-----------------------------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/03/2001
        'Arguments           - None
        'Return Value        - None
        'Function            - To Add Functionality of Authorize/Refresh/Close
        '-----------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim strSql As String
        Dim strSql1 As String
        Dim strSql2 As String
        Dim strErrMsg As String
        Dim strAns As MsgBoxResult
        Dim enmValue As eMPowerFunctions.ConfirmWindowReturnEnum
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        Dim varItemCode As Object
        Dim varCustDrgNo As Object
        Dim dblDespatch As Double
        Dim rsDespatch As ClsResultSetDB
        Dim mblnappendsoitem_unit As Boolean = False
        Dim mblnappendsoitem_customer As Boolean = False
        Dim strRetMsg As String = String.Empty

        Select Case e.Button
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_AUTHORIZE 'Authorize PO
                If ValidAuthorize() = True Then
                    enmValue = ConfirmWindow(10163, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) 'Are You Sure To Authorize the PO
                    If enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                        If DTEffectiveDate.Value > GetServerDate() Then
                            Call ConfirmWindow(10198, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                        Else
                            If DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(FormatDateTime(DTEffectiveDate.Value, DateFormat.LongDate)), CDate(FormatDateTime(GetServerDate, DateFormat.LongDate))) >= 0 Then
                                'Update Cust_ord_hdr table
                                '10844039--Starts
                                strSql1 = "Select AppendSOItem from sales_parameter where unit_code='" & gstrUNITID & "'"
                                mblnappendsoitem_unit = Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSql1))
                                If mblnappendsoitem_unit = True Then
                                    strSql2 = "Select appendsoitem from customer_mst where unit_code='" & gstrUNITID & "' and customer_code='" & txtCustomerCode.Text.Trim & "'"
                                    mblnappendsoitem_customer = Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSql2))
                                    If mblnappendsoitem_customer = True Then

                                        Call InsertPreviousSODetails(Trim(txtCustomerCode.Text), Trim(txtReferenceNo.Text), Trim(txtAmendmentNo.Text))
                                    End If
                                End If
                                '10844039--Ends
                                Dim CmdRP As New ADODB.Command
                                With CmdRP
                                    .CommandText = "USP_VALIDATE_RELATED_PARTY_BUDGET_VALUE_SALESORDER"
                                    .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                                    .ActiveConnection = mP_Connection
                                    .CommandTimeout = 0

                                    .Parameters.Append(.CreateParameter("@UNIT_CODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                                    .Parameters.Append(.CreateParameter("@SO_NO", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 34, txtReferenceNo.Text))
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

                                m_strSql = "Update cust_ord_hdr set First_Authorized='"
                                m_strSql = m_strSql & mP_User & "', Second_Authorized='"
                                m_strSql = m_strSql & mP_User & "', Third_Authorized ='"
                                m_strSql = m_strSql & mP_User & "', Authorized_Flag =1 where unit_code='" & gstrUNITID & "' and "
                                m_strSql = m_strSql & " cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'"
                                m_strSql = m_strSql & " and Account_Code='"
                                m_strSql = m_strSql & Trim(Me.txtCustomerCode.Text) & "' and "
                                m_strSql = m_strSql & " amendment_No='" & Trim(txtAmendmentNo.Text) & "'"
                                mP_Connection.Execute(m_strSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                'Update Cust_ord_dtl table
                                m_strSql = "Update cust_ord_dtl set Authorized_Flag =1 where unit_code='" & gstrUNITID & "' and "
                                m_strSql = m_strSql & " cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'"
                                m_strSql = m_strSql & " and Account_Code='"
                                m_strSql = m_strSql & Trim(Me.txtCustomerCode.Text) & "' and "
                                m_strSql = m_strSql & " amendment_No='" & Trim(txtAmendmentNo.Text) & "'"
                                mP_Connection.Execute(m_strSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
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
                                            m_strSql = m_strSql & " and Cust_ref = '" & Trim(txtReferenceNo.Text) & "'"
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
                                                m_strSql = m_strSql & " and Cust_ref = '" & Trim(txtReferenceNo.Text) & "'"
                                                m_strSql = m_strSql & " and amendment_no = '" & Trim(txtAmendmentNo.Text) & "' and active_flag = 'A'"
                                                m_strSql = m_strSql & " and ITem_code = '" & Trim(varItemCode) & "' and Cust_drgno = '" & Trim(varCustDrgNo) & "'"
                                                mP_Connection.Execute(m_strSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                            End If
                                        End With
                                    Next
                                End If
                                m_strSql = "Update cust_ord_dtl set Active_Flag = 'O' where unit_code='" & gstrUNITID & "' and cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'"
                                m_strSql = m_strSql & " and Account_Code = '" & Trim(Me.txtCustomerCode.Text) & "' and  amendment_No <> '" & Trim(Me.txtAmendmentNo.Text) & "' and authorized_flag =1 and Active_Flag <> 'L' " 'Active_Flag <> 'L' added against Issue ID 17432 of Auto SO Locking
                                m_strSql = m_strSql & " and cust_drgno in(select cust_drgno from cust_ord_dtl where unit_code='" & gstrUNITID & "' and cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "' and Account_Code= '" & Trim(Me.txtCustomerCode.Text) & "' and  amendment_No='" & Trim(Me.txtAmendmentNo.Text) & "')"
                                mP_Connection.Execute(m_strSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                'function to add upadetion cust_ord_hdr ActiveFlag on 01/07/2003
                                Call UpdateHdrActiveFlag()
                            Else
                                m_strSql = "Update cust_ord_hdr set First_Authorized='"
                                m_strSql = m_strSql & mP_User & "', Second_Authorized='"
                                m_strSql = m_strSql & mP_User & "', Third_Authorized ='"
                                m_strSql = m_strSql & mP_User & "', future_so =1 where unit_code='" & gstrUNITID & "' and "
                                m_strSql = m_strSql & " cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'"
                                m_strSql = m_strSql & " and Account_Code='"
                                m_strSql = m_strSql & Trim(Me.txtCustomerCode.Text) & "' and "
                                m_strSql = m_strSql & " amendment_No='" & Trim(txtAmendmentNo.Text) & "'"
                                mP_Connection.Execute(m_strSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            End If
                        End If
                        MsgBox("SO Authorized successfully.", MsgBoxStyle.Information, "eMpro")
                        If GetPlantName() = "HILEX" Then
                            Call automailer()
                        End If


                        txtReferenceNo.Text = ""
                        Call RefreshForm()
                        Call EnableControls(False, Me, True)
                        txtCustomerCode.Enabled = True
                        cmdHelp(0).Enabled = True

                        ssPOEntry.MaxRows = 0
                        Me.txtCustomerCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        'Call RefreshForm()
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
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE 'Close
                Me.Close()

        End Select
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub

    Private Sub cmdAuthorize_MouseDown(ByVal Sender As Object, ByVal e As UCActXCtl.cmdGrpAuthorise.MouseDownEventArgs) Handles cmdAuthorize.MouseDown
        '-----------------------------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/03/2001
        'Arguments           - None
        'Return Value        - None
        'Function            - To Add Value of Public Variable Used in Diffrent
        '                    - Cases.
        '-----------------------------------------------------------------------
        Select Case e.Index
            Case 3
                m_CloseButton = True
        End Select
    End Sub

    Private Sub txtReferenceNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtReferenceNo.LostFocus
        '-----------------------------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/03/2001
        'Arguments           - None
        'Return Value        - None
        'Function            - To Dispaly Details on Lost Focus of Referance No
        '-----------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim inti As Short
        If m_CloseButton = True Then
            m_CloseButton = False
            Exit Sub
        End If
        blnValidref = False
        If Len(Trim(txtReferenceNo.Text)) > 0 Then
            ' Check if records for the entered reference no exist or not
            m_strSql = " Select 1 from cust_ord_hdr where unit_code='" & gstrUNITID & "' and cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'and Account_Code='" & Trim(Me.txtCustomerCode.Text) & "'"
            rsRefNo = New ClsResultSetDB
            Call rsRefNo.GetResult(m_strSql)
            If rsRefNo.GetNoRows = 1 Then ' If there are records existing for the entered reference no
                rsRefNo.ResultSetClose()
                ' check whether the PO is active or not
                m_strSql = " Select 1 from cust_ord_hdr where unit_code='" & gstrUNITID & "' and cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'and Account_Code='" & Trim(Me.txtCustomerCode.Text) & "' and Active_Flag='A'"
                rsRefNo = New ClsResultSetDB
                Call rsRefNo.GetResult(m_strSql)
                If rsRefNo.GetNoRows > 0 Then 'For reference no to which no amendment had been made
                    rsRefNo.ResultSetClose()
                    m_strSql = " Select 1 from cust_ord_hdr where unit_code='" & gstrUNITID & "' and cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'and Account_Code='" & Trim(Me.txtCustomerCode.Text) & "' and Active_Flag='A' and (Authorized_Flag =1 or future_so = 1) "
                    rsRefNo = New ClsResultSetDB
                    Call rsRefNo.GetResult(m_strSql)
                    If rsRefNo.GetNoRows > 0 Then 'For reference no to which no amendment had been made
                        Call GetDetails()
                        'MsgBox ("This PO Is Already Authorized")
                        Call ConfirmWindow(10161, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                        blnValidref = False
                        cmdAuthorize.Enabled(0) = False
                        cmdAuthorize.Enabled(1) = False
                        cmdAuthorize.Enabled(2) = False
                        cmdAuthorize.Enabled(3) = True
                        Exit Sub
                    Else
                        Call GetDetails()
                        cmdAuthorize.Enabled(0) = True
                        cmdAuthorize.Enabled(1) = True
                        cmdAuthorize.Enabled(2) = False
                        cmdAuthorize.Enabled(3) = True
                        cmdAuthorize.Focus()
                        blnValidref = True
                        Exit Sub
                    End If
                    rsRefNo.ResultSetClose()
                Else
                    Call ConfirmWindow(10162, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                    txtReferenceNo.Text = ""
                    txtReferenceNo.Focus()
                    blnValidref = False
                    rsRefNo.ResultSetClose()
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
            Else
                rsRefNo.ResultSetClose()
                'MsgBox "There are no existing records for the Reference No"
                Call ConfirmWindow(10129, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                txtReferenceNo.Text = ""
                txtReferenceNo.Focus()
                blnValidref = False
                Exit Sub
            End If
        End If
        blnValidref = True
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub

    Private Sub txtReferenceNo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtReferenceNo.TextChanged
        '-----------------------------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/03/2001
        'Arguments           - None
        'Return Value        - None
        'Function            - To Refresh all the Data on Changeing Referance
        '                    - No.
        '-----------------------------------------------------------------------
        If Len(Trim(txtReferenceNo.Text)) = 0 Then
            Call RefreshForm()
            txtAmendmentNo.Text = ""
            txtSTax.Text = "" : txtCreditTerms.Text = "" : txtSChSTax.Text = ""
            chkOpenSo.CheckState = System.Windows.Forms.CheckState.Unchecked
            cmdHelp(3).Enabled = False
            txtAmendmentNo.Enabled = False
            txtAmendmentNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        End If

    End Sub

    Private Sub ctlHeader_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlHeader.Click
        '-----------------------------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/03/2001
        'Arguments           - None
        'Return Value        - None
        'Function            - To Show Empower Help on F4 Click
        '-----------------------------------------------------------------------
        Call ShowHelp("HLP" & Mid(Me.Name, 4, Len(Me.Name)) & ".htm") '("salesorderauth_trans(mkt).htm")
    End Sub

    Private Sub ctlHeader_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ctlHeader.Load

    End Sub

    Private Sub txtAddVAT_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAddVAT.TextChanged
        '-----------------------------------------------------------------------
        'Author              - Manoj Vaish
        'Create Date         - 21 Jul 2009
        'Arguments           - None
        'Return Value        - None
        'Function            - To Get The Label Additonal VAT
        '-----------------------------------------------------------------------
        Call FillLabel("ADDVAT")
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
    '10844039--Starts
    Public Sub InsertPreviousSODetails(ByRef pstrAccountCode As String, ByRef pstrRef As String, ByRef pstrAmendment As String)

        Dim strSql As String
        Dim strDrgItem As String
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        Dim varItemCode As Object
        Dim varDrgNo As Object
        Dim VarDelete As Object
        Dim rsCustOrdDtl As ClsResultSetDB
        Dim STRUPDATERATESQL As String
        On Error GoTo ErrHandler
        strSql = "insert into cust_ord_dtl (rate,unit_code,Account_Code,Cust_Ref, Amendment_No,Item_Code,Order_Qty,Despatch_Qty,Active_Flag,Cust_Mtrl,"
        strSql = strSql & " Cust_DrgNo,Packing,Others,Excise_Duty,Cust_Drg_Desc,Tool_Cost,Authorized_flag,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,"
        strSql = strSql & " OpenSO,SalesTax_Type,PerValue,InternalSONo,RevisionNo,Packing_Type)"
        strSql = strSql & " (Select DISTINCT rate,'" & gstrUNITID & "',Account_Code,Cust_Ref, Amendment_No = '" & pstrAmendment & "',Item_Code,Order_Qty,Despatch_Qty = 0 ,"
        strSql = strSql & " Active_Flag ,Cust_Mtrl,Cust_DrgNo,Packing,Others, Excise_Duty,Cust_Drg_Desc,Tool_Cost,Authorized_flag = 0 "
        strSql = strSql & " ,getdate(),'" & mP_User & "',getdate(),'" & mP_User & "',OpenSO,SalesTax_Type,PerValue,InternalSONo ,"
        strSql = strSql & " isnull(RevisionNo,0)+1 ,Packing_Type from Cust_ord_dtl where  unit_code='" & gstrUNITID & "' and Account_code = '" & pstrAccountCode & "' "
        strSql = strSql & " and cust_ref = '" & pstrRef & "' and Active_flag = 'A' and authorized_flag = 1 "
        strSql = strSql & " and amendment_no <> '" & pstrAmendment & "' "
        If ssPOEntry.MaxRows > 0 Then
            intMaxLoop = ssPOEntry.MaxRows
            strDrgItem = ""
            For intLoopCounter = 1 To intMaxLoop
                With ssPOEntry
                    .Row = intLoopCounter
                    VarDelete = Nothing
                    Call .GetText(0, intLoopCounter, VarDelete)
                    varDrgNo = Nothing
                    Call .GetText(2, intLoopCounter, varDrgNo)
                    varItemCode = Nothing
                    Call .GetText(3, intLoopCounter, varItemCode)
                    If VarDelete <> "*" Then
                        If Len(Trim(strDrgItem)) > 0 Then
                            strDrgItem = Trim(strDrgItem) & " and (Cust_drgNo <> '" & varDrgNo & "' or Item_code <> '" & varItemCode & "')"
                        Else
                            strDrgItem = " and (Cust_drgNo <> '" & varDrgNo & "' or Item_code <> '" & varItemCode & "')"
                        End If
                    End If
                End With
            Next
            If Len(Trim(strDrgItem)) > 0 Then
                strSql = strSql & strDrgItem & ")"
            End If
        End If

        mP_Connection.BeginTrans()
        mP_Connection.Execute(strSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        'mP_Connection.Execute(STRUPDATERATESQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        mP_Connection.CommitTrans()
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Exit Sub
    End Sub
    '10844039--Ends
    Public Sub automailer()
        '04 oct 2023'
        Dim Sqlcmd As New SqlCommand
        Sqlcmd.CommandType = CommandType.StoredProcedure
        Sqlcmd.CommandText = "USP_SENDAUTOMAILER_SO_AUTHORIZATION"
        Sqlcmd.Parameters.Clear()

        Try

            Sqlcmd.Parameters.Add("@unit_code", SqlDbType.VarChar, 20).Value = gstrUNITID
            Sqlcmd.Parameters.Add("@ACCOUNT_CODE", SqlDbType.VarChar, 20).Value = txtCustomerCode.Text.Trim.ToString
            Sqlcmd.Parameters.Add("@CUST_REF", SqlDbType.VarChar, 34).Value = txtReferenceNo.Text.Trim.ToString
            Sqlcmd.Parameters.Add("@AMENDMENT_NO", SqlDbType.VarChar, 25).Value = txtAmendmentNo.Text.Trim.ToString
            Sqlcmd.Parameters.Add("@ERRMSG", SqlDbType.VarChar, 8000).Value = String.Empty
            Sqlcmd.Parameters("@ERRMSG").Direction = ParameterDirection.Output

            SqlConnectionclass.ExecuteNonQuery(Sqlcmd)

        Catch ex As Exception
            RaiseException(ex)
        Finally
            Sqlcmd.Dispose()
        End Try

        '04 oct 2023

    End Sub
    
End Class