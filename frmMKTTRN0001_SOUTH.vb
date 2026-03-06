Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports CrystalDecisions.CrystalReports.Engine
Imports System.Data.SqlClient
Imports System.IO


Friend Class frmMKTTRN0001_SOUTH
	Inherits System.Windows.Forms.Form
	'---------------------------------------------------------------------------
	'(C) 2001 MIND, All rights reserved
	'
	'File Name          :   frmMKTTRN0001.frm
	'Function           :   Customer PO Entry
	'Created By         :   Meenu Gupta
	'Created on         :   28,Mar 2001
	'Revision History   :
	'Revised date       : 24 Aug 2001
	' Revision History  :   Nisha Rai
	'21/09/2001 MARKED CHECKED BY BCs changed on version 9
	'30/10/2001 CHANGED IN GET REF.DETAIL & GET AMM.DETAIL FUNCTION FOR CHAECKING ACTIVE_FLAG IN CUST_ORD_DTL on version 12
	'17-01-2002 done for internal issue log no 48,49,51,52,53,54 - checked out form no =4007
	'25-01-2002 done to allow 4 decimal places in Rate for MSSL-ED - checked out form no =4016
	'29-01-2002 enable Scrolling in Veiw mode for internel issue log no-198 & form No -4034
	'To in case of new item in authorizes SO ,to allow to change its values except existing
	'Item Code & DrgNo
	'29-01-2002 enable Internel issue log no-198 & form No -4035 set property of customerdes's
	'UseMnemonic to False
	'29-01-2002 form No -4036 Problem in POType Validate Event Message in Case of
	'Invalid PO Type
	'7-02-2002 form No -4042 changed in getamendmentNo() set formate of all dtp
	'controls to gstr date format error reprted from MSSL-ED Min & Max Date
	'28-02-02 Changed lable Surcharge % on formNo 4052 in grid control
	'11-03-02  changed in Update Query in case of Export SO.
	'01/04/2002 changed for financial year on form no 4079
	'23/04/2002 decimal places changes from currency Master
	'04/05/2002 for back date SO & Amendment Entry from MATE
	'24/06/2002 ADDED PROGRESS BAR TO VEIW THE DETAILS
	'25/06/2002 ADDED PROGRESS BAR TO VEIW THE DETAILS
	'05/08/2002 changed help of Item in Grid
	'*****Commented by Nisha on 30/08/2002
	'to remove limit of financial startdate
	'NOW IT CAN TAKE ANY BACK DATE
	'****Changed by nisha on 20/09/2002
	'1.for excise from tariff in case of button click in Grid
	'2.type mismatch in case of no item selected and click on zeroth column
	'****
	'12/10/2002 chaged ssSetFocus() by nisha
	'17/10/2002 changes done by nisha to addper value item
	'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
	'08/11/2002 Changed by nisha to add
	'AutoGeneration No in SO
	'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
	'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
	'12/12/2002 Changed by nisha to add
	'1.Excise Editable
	'2. Sorting on Customer Part Code
	'3. Print Button Avalable on SO Printing
	'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
	'parameter check added by nisha on 16/16/20031
	'changes done by nisha for SO Parameterized on 27/06/2003 01/07/2003
	'****Field Added by Ajay on 18/07/2003
	'1.Surcharge on S.Tax
	'---------------------------------------->>>>
	'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
	'Changes done by nisha on 25/11/2003
	'To Add Drg Desc
	'---------------------------------------->>>>
	'05/07/2004 Arshad Ali :- No column "Cust_drg_desc" exist in itemRate_mst so its is commented
	'In RowDetailsfromKeyBoard Procedure
	'---------------------------------------------------------------------------
	'Query replaced By Arshad Ali on 06/07/2004 to include amendment_no condition in query
	'In RowDetailsfromKeyBoard Procedure and ssPOEntry_LeaveCell Procedure
	'This solves the problem in editing of SO with hundreds of records
	'---------------------------------------------------------------------------
	'08/07/2004
    'chkOpenSo.value = 1 is replaced with chkOpenSo.value = CheckState.Checked
	'---------------------------------------------------------------------------
	'Revised By     : Arul Mozhi
	'Revised On     : 11-01-2005
	'Revision History : Cess On ED added in this Form
	'--------------------------------------------------------------------------------
	'Revised By     : Ashutosh Verma
	'Revised On     : 16 Apr 2007 ,Issue ID:19731
	'Revised Reason : Add Consignee Code changes.
	'=======================================================================================
    'Revised By     : Manoj Vaish
    'Revised On     : 13 Apr 2009 ,Issue ID:  eMpro-20090413-30069
    'Revised Reason : Validating Additional Excise Duty 
    '=======================================================================================
    'Revised By        : Siddharth Ranjan
    'Issue ID          : eMpro-20090709-33428
    'Revision Date     : 13 Jul 2009
    'History           : CSI functionality
    '=======================================================================================
    'Revised By        : Siddharth Ranjan
    'Issue ID          : eMpro-20091104-38405 
    'Revision Date     : 04 NOV 2009
    'History           : ERROR IN PER VALUE WHEN BLANK COMPARED WITH DOUBLE
    '=======================================================================================
    'Revised By         : Vinod Singh
    'Revision Date      : 02/05/2011
    'Reason             : Multi Unit 
    '=======================================================================================
    'Revised By        : Prashant Rajpal
    'Issue ID          : 10117810
    'Revision Date     : 22 JULY 2011
    'History           : Customer item description Wrong Saved. 
    '=======================================================================================
    'Modified By Roshan Singh on 09 Nov 2011 For multi unit Change Management.
    '=======================================================================================
    'REVISED BY         :   SHUBHRA VERMA
    'REVISED ON         :   12 MAR 2012
    'ISSUE ID           :   10354980  
    'DESCRIPTION        :   VALIDATION ADDED, SO THAT "SO" WILL NOT GET SAVED WITHOUT ITEM DETAILS
    '***************************************************************************************
    'REVISED BY         :   PRASHANT RAJPAL
    'REVISED ON         :   31 MAY 2013
    'ISSUE ID           :   10390304  
    'DESCRIPTION        :   CUSTOMER SUPPLIED MATERIAL PICKED 
    '***************************************************************************************
    'Revised By      : PRASHANT RAJPAL
    'Issue ID        : 10229989
    'Revision Date   : 10-aug -2013- 31-aug -2013
    'History         : Multiple So Functionlity 
    '****************************************************************************************
    'Revised By      : Shubhra Verma
    'Issue ID        : 10532789 
    'Revision Date   : 05/Feb/2013
    'History         : Inactive Drawing No Entry was allowed in SO Entry Form.
    '****************************************************************************************
    'REVISED BY         :   Abhinav Kumar
    'ISSUE ID           :   10736222   
    'DESCRIPTION        :   CHANGES DONE FOR CT2 ARE 3 FUNCTIONALITY - TO ENABLE CT2 REQD FLAG
    'REVISION DATE      :   09 JAN 2015
    '***************************************************************************************
    'REVISED BY         :   Abhinav Kumar
    'ISSUE ID           :   10797956  
    'DESCRIPTION        :   EMPRO-CHANGES IN CT2 ARE-3 FUNCTIONALITY 
    'REVISION DATE      :   30 APR 2015
    '***************************************************************************************
    'REVISED BY         :   Prashant Rajpal
    'ISSUE ID           :   10763705 
    'DESCRIPTION        :   Amendment Tracking
    'REVISION DATE      :   27 MAY 2015
    '***************************************************************************************
    'REVISED BY         :   Parveen Kumar
    'ISSUE ID           :   10808160
    'DESCRIPTION        :   eMPro-New functionality of EOP
    'REVISION DATE      :   23 JUN 2015
    '***************************************************************************************
    'REVISED BY     -  PRASHANT RAJPAL
    'REVISED ON     -  27 JULY 2015
    'PURPOSE        -  10856126 -ASN DOCK CODE FUNCTIONALITY
    '***************************************************************************************
    'REVISED BY         :   Abhinav Kumar
    'ISSUE ID           :   10869290 
    'DESCRIPTION        :   eMPro- Service Invoice Functionality
    'REVISION DATE      :   21 Sep 2015
    '***************************************************************************************
    'REVISED BY         :   PRASHANT RAJPAL
    'ISSUE ID           :   
    'DESCRIPTION        :   GST ISSUE
    'REVISION DATE      :   18-May-2016



    'REVISED BY         :   ABHIJIT KUMAR SINGH
    'ISSUE ID           :   101342142
    'DESCRIPTION        :   SHIP ADDRESS FIELD ADDTION 
    'REVISION DATE      :   22 AUG 2017
    '***************************************************************************************

    'REVISED BY         :   SUMIT KUMAR
    'ISSUE ID           :   101377199
    'DESCRIPTION        :   UPLOAD PO PDF AGAINST SO AND ALART THE EMAIL  
    'REVISION DATE      :   10-AUG-2018
    '***************************************************************************************
    '***************************************************************************************

    'REVISED BY         :   SUMIT KUMAR
    'ISSUE ID           :   101377199
    'DESCRIPTION        :   UPLOAD PO PDF AGAINST SO AND ALART THE EMAIL  
    'REVISION DATE      :   05-Dec-2018
    '***************************************************************************************


    Dim m_blnHelpFlag, m_blnCloseFlag As Boolean
    Dim rsdb As ClsResultSetDB
    Dim mintFormIndex, intRow As Short
    Dim mstrCode As String
    Dim m_Item_Code As String
    Dim m_blnChangeFormFlg As Boolean
    Dim mvalid As Boolean
    Dim m_strSql, strSQL As String
    Dim m_blnGetAmendmentDetails As Boolean
    Dim rsRefNo As ClsResultSetDB
    Dim m_ItemDesc, m_custItemDesc As String
    Dim strdt As String
    Dim strpotype As String
    Dim strSOType As String
    Dim blnInvalidData As Boolean
    Dim blnValidCurrency As Boolean
    Dim mstrPrevAccountCode As String
    Dim mstrPrevRefNo As String
    Dim blnValidAmendDate As Boolean
    Dim ArrDispatchQty() As Double
    Dim dtSODate As Date
    Dim m_strSqlquery As String
    'ISSUE ID : 10763705
    Dim mblnappendsoitem_customer As Boolean = False
    Dim MBLNALLOW_MULTIPLEHSNITEMS_CUSTOMER As Boolean = False
    Dim strexportType As String
    Dim blnNoneditableCreditTerms_onSO As Boolean = False
    Dim blnSO_EDITABLE As Boolean = False




    Private Sub chkAddCustSupp_KeyPress(ByRef KeyAscii As Short)
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                chkOpenSo.Focus()
            Case 39, 34, 96
                KeyAscii = 0
        End Select
    End Sub

    Private Sub chkOpenSo_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkOpenSo.CheckedChanged
        ' added by priti sharma on 25.03.2019 
        If chkOpenSo.Checked Then
            'strSQL = "SELECT top 1 Party_Code FROM DBO.FIN_GET_RELATED_PARTIES ('" & gstrUNITID & "') WHERE CUST_VEND='C' and Party_Code='" & Trim(txtCustomerCode.Text) & "'"
            'If IsRecordExists(strSQL) = True Then
            '    MsgBox("Open SO is not allowed for Related Parties ", MsgBoxStyle.Information, ResolveResString(100))
            '    chkOpenSo.Checked = False
            '    Exit Sub
            'End If
        End If
        'ends 
    End Sub
    Private Sub chkOpenSo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles chkOpenSo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Select Case KeyAscii
            Case 39, 34, 96
                KeyAscii = 0
            Case System.Windows.Forms.Keys.Return
                If ctlPerValue.Enabled = True Then
                    ctlPerValue.Focus()
                End If
        End Select
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub chkOpenSo_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkOpenSo.Leave
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        intMaxLoop = ssPOEntry.MaxRows
        With ssPOEntry
            If chkOpenSo.CheckState = 1 Then
                For intLoopCounter = 1 To intMaxLoop
                    .Col = 1
                    .Col2 = 1
                    .Row = intLoopCounter
                    .Row2 = intLoopCounter
                    .Value = System.Windows.Forms.CheckState.Checked
                    .Col = 5
                    .Col2 = 5
                    .Row = intLoopCounter
                    .Row2 = intLoopCounter
                    .Text = CStr(0)
                Next
                .Col = 5
                .Col2 = 5
                .Row = 1
                .Row2 = .MaxRows
                .BlockMode = True
                .Lock = True
                .BlockMode = False
                'tO LOCK Open Item Flag
                .Col = 1
                .Col2 = 1
                .Row = 1
                .Row2 = .MaxRows
                .BlockMode = True
                .Lock = True
                .BlockMode = False
            Else
                'tO UnLOCK Open Item Flag
                .Col = 1
                .Col2 = 1
                .Row = 1
                .Row2 = .MaxRows
                .BlockMode = True
                .Lock = False
                .BlockMode = False
            End If
        End With
    End Sub
    Private Sub cmbPOType_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbPOType.Enter
        Call Dropdown(Me.cmbPOType.Handle.ToInt32)
    End Sub
    Private Sub cmbPOType_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles cmbPOType.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Return Then
            txtSTax.Focus()
        End If
    End Sub
    Private Sub cmbPOType_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles cmbPOType.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Select Case KeyAscii
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub cmdchangetype_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdchangetype.Click
        On Error GoTo ErrHandler
        Dim strsalesTerms As String
        Dim rssalesTerms As ClsResultSetDB
        Dim strSQL As String
        If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Or cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            strSQL = "select a.*,b.* from cust_ord_hdr a,cust_ord_dtl b where a.unit_code=b.unit_code AND a.Amendment_No=b.Amendment_No and a.cust_Ref=b.Cust_Ref and a.Account_Code=b.Account_Code and a.Cust_ref='" & txtReferenceNo.Text & "' and a.amendment_No='" & txtAmendmentNo.Text & "' AND A.UNIT_CODE='" & gstrUNITID & "'"
        ElseIf cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            strSQL = "select Account_Code,Cust_Ref,Amendment_No,Order_Date,Amendment_Date,Active_Flag,Currency_Code,Valid_Date,Effect_Date,Term_Payment,Special_Remarks,Pay_Remarks,Price_Remarks,Packing_Remarks,Frieght_Remarks,Transport_Remarks,Octorai_Remarks,Mode_Despatch,Delivery,First_Authorized,Second_Authorized,Third_Authorized,Authorized_Flag,Reason,PO_Type,SalesTax_Type,OpenSO,AddCustSupp,PerValue,InternalSONo,RevisionNo,Surcharge_code,Future_SO,ECESS_Code,Consignee_Code,ADDVAT_TYPE from cust_ord_hdr  where Cust_ref='" & txtReferenceNo.Text & "' and account_code='" & txtCustomerCode.Text & "' Consignee_Code= '" & Trim(txtConsCode.Text) & "' AND UNIT_CODE='" & gstrUNITID & "'"
        End If
        m_pstrSql = strSQL
        rssalesTerms = New ClsResultSetDB
        strsalesTerms = "Select Description From SaleTerms_Mst Where SaleTerms_Type ='PY' AND UNIT_CODE='" & gstrUNITID & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
        rssalesTerms.GetResult(strsalesTerms, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        If rssalesTerms.GetNoRows > 0 Then
            rssalesTerms.ResultSetClose()
            strsalesTerms = "Select Description From SaleTerms_Mst Where UNIT_CODE='" & gstrUNITID & "' AND SaleTerms_Type ='PR' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
            rssalesTerms = New ClsResultSetDB
            rssalesTerms.GetResult(strsalesTerms, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            If rssalesTerms.GetNoRows > 0 Then
                rssalesTerms.ResultSetClose()
                strsalesTerms = "Select Description From SaleTerms_Mst Where UNIT_CODE='" & gstrUNITID & "' AND SaleTerms_Type ='PK' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
                rssalesTerms = New ClsResultSetDB
                rssalesTerms.GetResult(strsalesTerms, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                If rssalesTerms.GetNoRows > 0 Then
                    rssalesTerms.ResultSetClose()
                    strsalesTerms = "Select Description From SaleTerms_Mst Where UNIT_CODE='" & gstrUNITID & "' AND SaleTerms_Type ='FR' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
                    rssalesTerms = New ClsResultSetDB
                    rssalesTerms.GetResult(strsalesTerms, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                    If rssalesTerms.GetNoRows > 0 Then
                        rssalesTerms.ResultSetClose()
                        strsalesTerms = "Select Description From SaleTerms_Mst Where UNIT_CODE='" & gstrUNITID & "' AND SaleTerms_Type ='TR' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
                        rssalesTerms = New ClsResultSetDB
                        rssalesTerms.GetResult(strsalesTerms, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                        If rssalesTerms.GetNoRows > 0 Then
                            rssalesTerms.ResultSetClose()
                            strsalesTerms = "Select Description From SaleTerms_Mst Where UNIT_CODE='" & gstrUNITID & "' AND SaleTerms_Type ='OC' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
                            rssalesTerms = New ClsResultSetDB
                            rssalesTerms.GetResult(strsalesTerms, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                            If rssalesTerms.GetNoRows > 0 Then
                                rssalesTerms.ResultSetClose()
                                strsalesTerms = "Select Description From SaleTerms_Mst Where UNIT_CODE='" & gstrUNITID & "' AND SaleTerms_Type ='MO' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
                                rssalesTerms = New ClsResultSetDB
                                rssalesTerms.GetResult(strsalesTerms, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                                If rssalesTerms.GetNoRows > 0 Then
                                    rssalesTerms.ResultSetClose()
                                    strsalesTerms = "Select Description From SaleTerms_Mst Where UNIT_CODE='" & gstrUNITID & "' AND SaleTerms_Type ='DL' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
                                    rssalesTerms = New ClsResultSetDB
                                    rssalesTerms.GetResult(strsalesTerms, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                                    If rssalesTerms.GetNoRows > 0 Then
                                        rssalesTerms.ResultSetClose()
                                        Select Case cmdButtons.Mode
                                            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                                                frmMKTTRN0010.formload("MODE_ADD")
                                            Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                                                frmMKTTRN0010.formload("MODE_EDIT")
                                            Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                                                frmMKTTRN0010.formload("MODE_VIEW")
                                        End Select
                                        Call frmMKTTRN0010.Show()
                                    Else
                                        Call ConfirmWindow(10480, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                                        cmdchangetype.Focus()
                                        rssalesTerms.ResultSetClose()
                                        Exit Sub
                                    End If
                                Else
                                    Call ConfirmWindow(10480, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                                    cmdchangetype.Focus()
                                    rssalesTerms.ResultSetClose()
                                    Exit Sub
                                End If
                            Else
                                Call ConfirmWindow(10480, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                                cmdchangetype.Focus()
                                rssalesTerms.ResultSetClose()
                                Exit Sub
                            End If
                        Else
                            Call ConfirmWindow(10480, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            cmdchangetype.Focus()
                            rssalesTerms.ResultSetClose()
                            Exit Sub
                        End If
                    Else
                        Call ConfirmWindow(10480, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                        cmdchangetype.Focus()
                        rssalesTerms.ResultSetClose()
                        Exit Sub
                    End If
                Else
                    Call ConfirmWindow(10480, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                    cmdchangetype.Focus()
                    rssalesTerms.ResultSetClose()
                    Exit Sub
                End If
            Else
                Call ConfirmWindow(10480, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                cmdchangetype.Focus()
                rssalesTerms.ResultSetClose()
                Exit Sub
            End If
        Else
            Call ConfirmWindow(10480, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
            cmdchangetype.Focus()
            rssalesTerms.ResultSetClose()
            Exit Sub
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub cmdchangetype_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles cmdchangetype.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Return Then
            System.Windows.Forms.SendKeys.SendWait("{TAB}")
        End If
    End Sub
    Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
        Dim Index As Short = cmdHelp.GetIndex(eventSender)
        On Error GoTo ErrHandler
        Dim varRetVal As Object
        On Error GoTo ErrHandler
        Dim strAmend, strString As String
        Select Case Index
            Case 0
                Select Case cmdButtons.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        With Me.txtCustomerCode
                            If Len(.Text) = 0 Then
                                'Samiksha SMRC start
                                If gstrUNITID = "STH" Then
                                    varRetVal = ShowList(1, .MaxLength, "", "A.Customer_code", "B.Cust_Name", "CUSTOMER_CONSIGNEE_MAPPING A,Customer_mst B", " and A.CUSTOMER_CODE = B.Customer_Code AND A.UNIT_CODE=B.UNIT_CODE and ((isnull(B.deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= B.deactive_date))", , , , , , "A.UNIT_CODE")
                                Else
                                    varRetVal = ShowList(1, .MaxLength, "", "Customer_Code", "Cust_Name", "Customer_Mst ", " and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))")
                                End If
                                'Samiksha SMRC end
                                If varRetVal = "-1" Then
                                    Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                                    .Text = ""
                                Else
                                    .Text = varRetVal
                                    '' added by priti on 11 May 2020 for OPEN SO Issue
                                    'strSQL = "SELECT top 1 Party_Code FROM DBO.FIN_GET_RELATED_PARTIES ('" & gstrUNITID & "') WHERE CUST_VEND='C' and Party_Code='" & Trim(txtCustomerCode.Text) & "'"
                                    'If IsRecordExists(strSQL) = True Then
                                    '    MsgBox("THIS CUSTOMER IS RELATED PARTIES.OPEN SO IS NOT ALLOWED FOR RELATED PARTIES. ", MsgBoxStyle.Information, ResolveResString(100))
                                    'End If
                                    '' code ends by priti on 11 May 2020 for OPEN SO Issue
                                End If
                            Else
                                'Samiksha SMRC start
                                If gstrUNITID = "STH" Then
                                    varRetVal = ShowList(1, .MaxLength, .Text, "A.Customer_code", "B.Cust_Name", "CUSTOMER_CONSIGNEE_MAPPING A,Customer_mst B", " and A.CUSTOMER_CODE = B.Customer_Code AND A.UNIT_CODE=B.UNIT_CODE and ((isnull(B.deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= B.deactive_date))", , , , , , "A.UNIT_CODE")
                                Else
                                    varRetVal = ShowList(1, .MaxLength, .Text, "Customer_code", "Cust_Name", "Customer_Mst", " and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))")
                                End If
                                'Samiksha SMRC end
                                If varRetVal = "-1" Then
                                    Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                                    .Text = ""
                                Else
                                    .Text = varRetVal
                                    '' added by priti on 11 May 2020 for OPEN SO Issue
                                    'strSQL = "SELECT top 1 Party_Code FROM DBO.FIN_GET_RELATED_PARTIES ('" & gstrUNITID & "') WHERE CUST_VEND='C' and Party_Code='" & Trim(txtCustomerCode.Text) & "'"
                                    'If IsRecordExists(strSQL) = True Then
                                    '    MsgBox("THIS CUSTOMER IS RELATED PARTIES.OPEN SO IS NOT ALLOWED FOR RELATED PARTIES. ", MsgBoxStyle.Information, ResolveResString(100))
                                    'End If
                                    ' '' code ends by priti on 11 May 2020 for OPEN SO Issue
                                End If
                            End If
                        End With
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT, UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                        With Me.txtCustomerCode
                            If Len(.Text) = 0 Then
                                'Samiksha SMRC start
                                If gstrUNITID = "STH" Then
                                    varRetVal = ShowList(1, .MaxLength, "", "A.Customer_code", "B.Cust_Name", "CUSTOMER_CONSIGNEE_MAPPING A,Customer_mst B", " and A.CUSTOMER_CODE = B.Customer_Code AND A.UNIT_CODE=B.UNIT_CODE and ((isnull(B.deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= B.deactive_date))", , , , , , "A.UNIT_CODE")
                                Else
                                    varRetVal = ShowList(1, .MaxLength, "", "a.Account_Code", "b.Cust_Name", "Cust_ord_hdr a,Customer_mst b", " and a.Account_code = b.Customer_code AND A.UNIT_CODE=B.UNIT_CODE and ((isnull(b.deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= b.deactive_date))", , , , , , "A.UNIT_CODE")
                                End If
                                'Samiksha SMRC end

                                If varRetVal = "-1" Then
                                    Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                                    .Text = ""
                                Else
                                    .Text = varRetVal
                                    '' added by priti on 11 May 2020 for OPEN SO Issue
                                    'strSQL = "SELECT top 1 Party_Code FROM DBO.FIN_GET_RELATED_PARTIES ('" & gstrUNITID & "') WHERE CUST_VEND='C' and Party_Code='" & Trim(txtCustomerCode.Text) & "'"
                                    'If IsRecordExists(strSQL) = True Then
                                    '    MsgBox("THIS CUSTOMER IS RELATED PARTIES.OPEN SO IS NOT ALLOWED FOR RELATED PARTIES. ", MsgBoxStyle.Information, ResolveResString(100))
                                    'End If
                                    '' code ends by priti on 11 May 2020 for OPEN SO Issue
                                End If
                            Else
                                'Samiksha SMRC start
                                If gstrUNITID = "STH" Then
                                    varRetVal = ShowList(1, .MaxLength, .Text, "A.Customer_code", "B.Cust_Name", "CUSTOMER_CONSIGNEE_MAPPING A,Customer_mst B", " and A.CUSTOMER_CODE = B.Customer_Code AND A.UNIT_CODE=B.UNIT_CODE and ((isnull(B.deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= B.deactive_date))", , , , , , "A.UNIT_CODE")
                                Else
                                    varRetVal = ShowList(1, .MaxLength, .Text, "a.Account_Code", "b.Cust_Name", "Cust_ord_hdr a,Customer_mst b", " and a.Account_code =b.Customer_code AND A.UNIT_CODE=B.UNIT_CODE and ((isnull(b.deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= b.deactive_date))", , , , , , "A.UNIT_CODE")
                                End If

                                'Samiksha SMRC end

                                If varRetVal = "-1" Then
                                    Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                                    .Text = ""
                                Else
                                    .Text = varRetVal
                                    '' added by priti on 11 May 2020 for OPEN SO Issue
                                    'strSQL = "SELECT top 1 Party_Code FROM DBO.FIN_GET_RELATED_PARTIES ('" & gstrUNITID & "') WHERE CUST_VEND='C' and Party_Code='" & Trim(txtCustomerCode.Text) & "'"
                                    'If IsRecordExists(strSQL) = True Then
                                    '    MsgBox("THIS CUSTOMER IS RELATED PARTIES.OPEN SO IS NOT ALLOWED FOR RELATED PARTIES. ", MsgBoxStyle.Information, ResolveResString(100))
                                    'End If
                                    '' code ends by priti on 11 May 2020 for OPEN SO Issue
                                End If
                            End If
                            .Focus()
                        End With
                End Select
            Case 1
                Select Case cmdButtons.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        With Me.txtReferenceNo
                            If Len(.Text) = 0 Then
                                varRetVal = ShowList(1, .MaxLength, "", "Cust_Ref", DateColumnNameInShowList("order_date") & "As order_date", "cust_ord_hdr", "and Account_Code='" & Trim(txtCustomerCode.Text) & "' and Consignee_Code='" & Trim(txtConsCode.Text) & "' and Active_Flag ='A' ")
                                If varRetVal = "-1" Then
                                    Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                                    .Text = ""
                                Else
                                    .Text = varRetVal
                                End If
                            Else
                                varRetVal = ShowList(1, .MaxLength, .Text, "Cust_Ref", DateColumnNameInShowList("order_date") & "As order_date", "cust_ord_hdr", "and Account_Code='" & Trim(txtCustomerCode.Text) & "' and Consignee_Code='" & Trim(txtConsCode.Text) & "' and active_flag ='A'")
                                If varRetVal = "-1" Then
                                    Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                                    .Text = ""
                                Else
                                    .Text = varRetVal
                                End If
                            End If
                            .Focus()
                        End With
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT, UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                        With Me.txtReferenceNo
                            If Len(.Text) = 0 Then
                                varRetVal = ShowList(1, .MaxLength, "", "Cust_Ref", DateColumnNameInShowList("order_date") & "As order_date", "cust_ord_hdr", "and Account_Code='" & Trim(txtCustomerCode.Text) & "' and Consignee_Code='" & Trim(txtConsCode.Text) & "' ")
                                If varRetVal = "-1" Then
                                    Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                                    .Text = ""
                                Else
                                    .Text = varRetVal
                                End If
                            Else
                                varRetVal = ShowList(1, .MaxLength, .Text, "Cust_Ref", DateColumnNameInShowList("order_date") & "As order_date", "cust_ord_hdr", "and Account_Code='" & Trim(txtCustomerCode.Text) & "' and Consignee_Code='" & Trim(txtConsCode.Text) & "' ")
                                If varRetVal = "-1" Then
                                    Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                                    .Text = ""
                                Else
                                    .Text = varRetVal
                                End If
                            End If
                            .Focus()
                        End With
                End Select
            Case 2
                With txtCurrencyType
                    If Len(.Text) = 0 Then
                        varRetVal = ShowList(1, .MaxLength, "", "Currency_Code", "Description", "Currency_mst")
                        If varRetVal = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            .Text = ""
                        Else
                            .Text = varRetVal
                        End If
                    Else
                        varRetVal = ShowList(1, .MaxLength, .Text, "Currency_Code", "Description", "Currency_mst")
                        If varRetVal = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            .Text = ""
                        Else
                            .Text = varRetVal
                        End If
                    End If
                    .Focus()
                End With
                '******Amendment No Help
            Case 3
                If cmdButtons.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                    With txtAmendmentNo
                        strString = txtAmendmentNo.Text & "%"
                        If txtAmendmentNo.Text <> "" Then
                            strAmend = " where UNIT_CODE='" & gstrUNITID & "' AND cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "' and Account_Code='" & txtCustomerCode.Text & "' and Consignee_Code='" & Trim(txtConsCode.Text) & "' and Amendment_No like '" & strString & "' "
                        Else
                            strAmend = " where UNIT_CODE='" & gstrUNITID & "' AND cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'and Account_Code='" & txtCustomerCode.Text & "' and Consignee_Code='" & Trim(txtConsCode.Text) & "' "
                        End If
                        varRetVal = ctlEMPHelpSOEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "SELECT Amendment_No," & DateColumnNameInShowList("Amendment_date") & "As Amendment_date, Cust_Ref  FROM cust_ord_hdr " & strAmend & " ", "List of All Amendments", 1)
                        If UBound(varRetVal) = -1 Then Exit Sub
                        If varRetVal(0) = "0" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            .Text = ""
                        Else
                            .Text = varRetVal(0)
                        End If
                        .Focus()
                    End With
                End If
            Case 4
                With txtSTax
                    If Len(.Text) = 0 Then
                        varRetVal = ShowList(1, .MaxLength, "", "txrt_Rate_No", "txrt_rateDesc", "Gen_TaxRate", "and tx_TaxeID in ('LST','CST','VAT') ")
                        If varRetVal = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            .Text = ""
                        Else
                            .Text = varRetVal
                        End If
                    Else
                        varRetVal = ShowList(1, .MaxLength, .Text, "txrt_Rate_No", "txrt_rateDesc", "Gen_TaxRate", "and tx_TaxeID in ('LST','CST') ")
                        If varRetVal = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            .Text = ""
                        Else
                            .Text = varRetVal
                        End If
                    End If
                    .Focus()
                End With
                '*****CreditTerms Help
            Case 5
                With txtCreditTerms
                    If Len(.Text) = 0 Then
                        varRetVal = ShowList(1, .MaxLength, "", "crtrm_termID", "crTrm_desc", "Gen_CreditTrmMaster", " and crtrm_status =1")
                        If varRetVal = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            .Text = ""
                        Else
                            .Text = varRetVal
                        End If
                    Else
                        varRetVal = ShowList(1, .MaxLength, .Text, "crtrm_termID", "crTrm_desc", "Gen_CreditTrmMaster", " and crtrm_status =1")
                        If varRetVal = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            .Text = ""
                        Else
                            .Text = varRetVal
                        End If
                    End If
                    .Focus()
                End With
                '1.Surcharge on S.Tax
            Case 6
                With txtSChSTax
                    If Len(.Text) = 0 Then
                        varRetVal = ShowList(1, .MaxLength, "", "txrt_Rate_No", "txrt_rateDesc", "Gen_TaxRate", " and tx_TaxeID ='SST'")
                        If varRetVal = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            .Text = ""
                        Else
                            .Text = varRetVal
                        End If
                    Else
                        varRetVal = ShowList(1, .MaxLength, .Text, "TxRt_Rate_no", "TxRt_RateDesc", "Gen_TaxRate", " and tx_TaxeID ='SST'")
                        If varRetVal = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            .Text = ""
                        Else
                            .Text = varRetVal
                        End If
                    End If
                    .Focus()
                End With
            Case 7
                Select Case cmdButtons.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        With txtConsCode
                            If Len(.Text) = 0 Then
                                varRetVal = ShowList(1, .MaxLength, "", "Customer_Code", "Cust_Name", "Customer_Mst ", " and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))")
                                If varRetVal = "-1" Then
                                    MsgBox("Invalid Consignee Code", MsgBoxStyle.Information, "eMpower")
                                    .Text = ""
                                Else
                                    .Text = varRetVal
                                End If
                            Else
                                varRetVal = ShowList(1, .MaxLength, .Text, "Customer_code", "Cust_Name", "Customer_Mst", " and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))")
                                If varRetVal = "-1" Then
                                    MsgBox("Invalid Consignee Code", MsgBoxStyle.Information, "eMpower")
                                    .Text = ""
                                Else
                                    .Text = varRetVal
                                End If
                            End If
                        End With
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT, UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                        With txtConsCode
                            If Len(.Text) = 0 Then
                                varRetVal = ShowList(1, .MaxLength, "", "a.Consignee_Code", "b.Cust_Name", "Cust_ord_hdr a,Customer_mst b", " and a.Consignee_Code = b.Customer_code and a.unit_code=b.unit_code and ((isnull(b.deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= b.deactive_date))", , , , , , "a.unit_code")
                                If varRetVal = "-1" Then
                                    MsgBox("Invalid Consignee Code", MsgBoxStyle.Information, "eMpower")
                                    .Text = ""
                                Else
                                    .Text = varRetVal
                                End If
                            Else
                                varRetVal = ShowList(1, .MaxLength, .Text, "a.Consignee_Code", "b.Cust_Name", "Cust_ord_hdr a,Customer_mst b", " and a.Consignee_Code =b.Customer_code and a.unit_code=b.unit_code and ((isnull(b.deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= b.deactive_date))", , , , , , "a.unit_code")
                                If varRetVal = "-1" Then
                                    MsgBox("Invalid Consignee Code", MsgBoxStyle.Information, "eMpower")
                                    .Text = ""
                                Else
                                    .Text = varRetVal
                                End If
                            End If
                            .Focus()
                        End With
                End Select
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
        On Error GoTo ErrHandler
        Select Case Index
            Case 1
                m_blnHelpFlag = True
            Case 3
                m_blnHelpFlag = True
        End Select
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0001_SOUTH_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        On Error GoTo ErrHandler
        'Make the node normal font
        frmModules.NodeFontBold(Tag) = False
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub txtAmendmentNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAmendmentNo.LostFocus
        On Error GoTo ErrHandler
        Dim inti As Short
        Dim rsAmend As ClsResultSetDB
        mvalid = False
        If m_blnHelpFlag = True Then
            m_blnHelpFlag = False
            Exit Sub
        End If
        If m_blnCloseFlag = True Then
            m_blnCloseFlag = False
            Exit Sub
        End If
        If Len(Trim(txtAmendmentNo.Text)) > 0 Then
            m_strSql = "Select top 1 1 from cust_ord_hdr where unit_code='" & gstrUNITID & "' and Amendment_No='" & Trim(txtAmendmentNo.Text) & "' and Cust_Ref ='" & Trim(txtReferenceNo.Text) & "'and Account_Code='" & Trim(txtCustomerCode.Text) & "'  and Consignee_Code='" & Trim(Me.txtConsCode.Text) & "' "
            rsAmend = New ClsResultSetDB
            rsAmend.GetResult(m_strSql)
            If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                If rsAmend.GetNoRows > 0 Then
                    rsAmend.ResultSetClose()
                    cmdButtons.Focus()
                    Call ConfirmWindow(10141, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                    txtAmendmentNo.Text = ""
                    mvalid = True
                    txtAmendmentNo.Focus()
                    Exit Sub
                Else
                    rsAmend.ResultSetClose()
                    DTAmendmentDate.Enabled = True
                    DTEffectiveDate.Enabled = True
                    DTValidDate.Enabled = True
                    txtAmendReason.Enabled = True
                    txtAmendReason.BackColor = System.Drawing.Color.White
                    cmdchangetype.Enabled = True
                    'for account Plug in
                    '1.S.Tax at Header Lavel
                    '2.Credit Terms at Main Screen
                    txtSTax.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : txtSTax.Enabled = True : cmdHelp(4).Enabled = True
                    If blnNoneditableCreditTerms_onSO Then
                        txtCreditTerms.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : txtCreditTerms.Enabled = False : cmdHelp(5).Enabled = False
                    Else
                        txtCreditTerms.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : txtCreditTerms.Enabled = True : cmdHelp(5).Enabled = True
                    End If
                    '  chkAddCustSupp.Enabled = True
                    chkOpenSo.Enabled = True
                    '1.Surcharge on S.Tax
                    txtSChSTax.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : txtSChSTax.Enabled = True : cmdHelp(6).Enabled = True
                    txtECSSTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : txtECSSTaxType.Enabled = True : CmdECSSTaxType.Enabled = True
                    With Me.ssPOEntry
                        .Enabled = True
                        .Row = 1
                        .Row2 = .MaxRows
                        .Col = 5
                        .Col2 = .MaxCols
                        .BlockMode = True
                        .Lock = False
                        .BlockMode = False
                    End With
                End If
            ElseIf cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                rsAmend.ResultSetClose()
                m_strSql = "Select top 1 1 from cust_ord_hdr where unit_code='" & gstrUNITID & "' and Amendment_No='" & Trim(txtAmendmentNo.Text) & "' and Cust_Ref ='" & Trim(txtReferenceNo.Text) & "'and Account_Code='" & Trim(txtCustomerCode.Text) & "' and Consignee_Code='" & Trim(Me.txtConsCode.Text) & "'  "
                rsAmend = New ClsResultSetDB
                rsAmend.GetResult(m_strSql)
                If rsAmend.GetNoRows > 0 Then
                    rsAmend.ResultSetClose()
                    m_strSql = "Select top 1 1 from cust_ord_hdr where unit_code='" & gstrUNITID & "' and Amendment_No='" & Trim(txtAmendmentNo.Text) & "' and Cust_Ref ='" & Trim(txtReferenceNo.Text) & "'and Account_Code='" & Trim(txtCustomerCode.Text) & "' and Consignee_Code='" & Trim(Me.txtConsCode.Text) & "'  and Active_Flag='A' "
                    rsAmend = New ClsResultSetDB
                    rsAmend.GetResult(m_strSql)
                    If rsAmend.GetNoRows > 0 Then
                        rsAmend.ResultSetClose()
                        m_strSql = "Select top 1 1  from cust_ord_hdr where unit_code='" & gstrUNITID & "' and Amendment_No='" & Trim(txtAmendmentNo.Text) & "' and Cust_Ref ='" & Trim(txtReferenceNo.Text) & "'and Account_Code='" & Trim(txtCustomerCode.Text) & "' and Consignee_Code='" & Trim(Me.txtConsCode.Text) & "'  and Active_Flag='A' and authorized_Flag=1 "
                        rsAmend = New ClsResultSetDB
                        rsAmend.GetResult(m_strSql)
                        If rsAmend.GetNoRows > 0 Then
                            rsAmend.ResultSetClose()
                            Call ConfirmWindow(10142, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            cmdButtons.Focus()
                            Call GetAmendmentDetails()
                            mvalid = True
                            cmdButtons.Focus()
                            Exit Sub
                        Else
                            rsAmend.ResultSetClose()
                            Call GetAmendmentDetails()
                            'change account Plug in
                            With ssPOEntry
                                .Col = 1
                                .Col2 = .MaxCols
                                .BlockMode = True
                                .Lock = True
                                .BlockMode = False
                            End With
                            cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD) = True
                            cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = True
                            cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = True
                            cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE) = True
                            cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL) = False
                            cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                        End If
                    Else
                        rsAmend.ResultSetClose()
                        cmdButtons.Focus()
                        Call ConfirmWindow(10143, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                        Call GetAmendmentDetails()
                        mvalid = True
                        cmdButtons.Focus()
                        Exit Sub
                    End If
                Else
                    rsAmend.ResultSetClose()
                    Call ConfirmWindow(10128, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                    txtAmendmentNo.Text = ""
                    mvalid = True
                    txtAmendmentNo.Focus()
                    Exit Sub
                End If
            End If
        Else
            Exit Sub
        End If
        mvalid = False
        cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = True
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub txtAmendmentNo_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtAmendmentNo.TextChanged
        If Len(Trim(txtAmendmentNo.Text)) = 0 Then
            If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                Call RefreshForm()
                cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
            End If
        End If
    End Sub
    Private Sub txtAmendmentNo_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtAmendmentNo.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = 112 Then Call cmdHelp_Click(cmdHelp.Item(3), New System.EventArgs())
        If KeyCode = System.Windows.Forms.Keys.Return Then
            'Call txtAmendmentNo_LostFocus
            If mvalid = False Then
                If txtAmendReason.Enabled Then
                    txtAmendReason.Focus()
                Else
                    cmdButtons.Focus()
                End If
            End If
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub txtAmendmentNo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtAmendmentNo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
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
    Private Sub txtAmendReason_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtAmendReason.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.Return Then
            If DTDate.Enabled = True Then
                DTDate.Focus()
            Else
                If DTAmendmentDate.Enabled = True Then
                    DTAmendmentDate.Focus()
                Else
                    If DTEffectiveDate.Enabled Then
                        DTEffectiveDate.Focus()
                    Else
                        cmdButtons.Focus()
                    End If
                End If
            End If
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub txtAmendReason_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtAmendReason.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
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
    Private Sub txtConsCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtConsCode.TextChanged
        Dim rsdb3 As ClsResultSetDB
        If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            lblConsDesc.Text = ""
            If txtConsCode.Enabled = True Then txtConsCode.Focus()
        ElseIf cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            If Len(Trim(txtCustomerCode.Text)) <> 0 Then
                m_strSql = "Select top 1 1 from cust_Ord_hdr where unit_code='" & gstrUNITID & "' and account_code ='" & txtCustomerCode.Text & "' and consignee_code='" & Trim(txtConsCode.Text) & "' "
                rsdb3 = New ClsResultSetDB
                rsdb3.GetResult(m_strSql)
                If rsdb3.GetNoRows > 0 Then
                    txtReferenceNo.Enabled = True
                    txtReferenceNo.BackColor = System.Drawing.Color.White
                    cmdHelp(1).Enabled = True
                Else
                    lblConsDesc.Text = ""
                    If txtReferenceNo.Enabled = True Then txtReferenceNo.Text = ""
                End If
                rsdb3.ResultSetClose()
            End If
        End If
    End Sub
    Private Sub txtConsCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtConsCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
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
        On Error GoTo ErrHandler
        If KeyCode = 112 Then Call cmdHelp_Click(cmdHelp.Item(7), New System.EventArgs()) 'Help should be invoked if F1 key is pressed
        If KeyCode = System.Windows.Forms.Keys.Return Then
            Call txtConsCode_Validating(txtConsCode, New System.ComponentModel.CancelEventArgs(False))
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtConsCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtConsCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        Dim rsCD As ClsResultSetDB
        If Len(Trim(txtConsCode.Text)) = 0 Then
            GoTo EventExitSub
        Else
            m_strSql = "Select Cust_Name from Customer_mst where unit_code='" & gstrUNITID & "' and customer_Code='" & Trim(txtConsCode.Text) & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
            rsCD = New ClsResultSetDB
            rsCD.GetResult(m_strSql)
            If rsCD.GetNoRows = 0 Then
                rsCD.ResultSetClose()
                MsgBox("Invalid Consignee Code !!!", MsgBoxStyle.Information, "eMpower")
                txtConsCode.Text = ""
                lblConsDesc.Text = ""
                Cancel = True
                txtConsCode.Focus()
                GoTo EventExitSub
            Else
                lblConsDesc.Text = IIf(UCase(rsCD.GetValue("Cust_Name")) = "UNKNOWN", "", rsCD.GetValue("Cust_Name"))
                rsCD.ResultSetClose()
                If txtReferenceNo.Enabled = True Then txtReferenceNo.Focus()
            End If
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
    Private Sub txtCreditTerms_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCreditTerms.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case 112
                Call cmdHelp_Click(cmdHelp.Item(5), New System.EventArgs())
        End Select
    End Sub
    Private Sub txtCreditTerms_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCreditTerms.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                chkOpenSo.Focus()
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtCreditTerms_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCreditTerms.Leave
        If Len(Trim(txtCreditTerms.Text)) <> 0 Then
            m_strSql = "select crtrm_TermID,crtrm_Desc from Gen_CreditTrmMaster where unit_code='" & gstrUNITID & "' and crtrm_TermID = '" & txtCreditTerms.Text & "'"
            rsdb = New ClsResultSetDB
            rsdb.GetResult(m_strSql)
            If rsdb.GetNoRows > 0 Then
                Call FillLabel("CREDIT")
            Else
                MsgBox("Entered Credit Term does not exist", MsgBoxStyle.Information, "empower")
                txtCreditTerms.Text = ""
                txtCreditTerms.Focus()
            End If
            rsdb.ResultSetClose()
        End If
    End Sub
    Private Sub txtCurrencyType_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCurrencyType.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = 112 Then Call cmdHelp_Click(cmdHelp.Item(2), New System.EventArgs()) 'Help should be invoked if F1 key is pressed
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub txtCurrencyType_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCurrencyType.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case 39, 34, 96
                KeyAscii = 0
            Case System.Windows.Forms.Keys.Return
                If Len(Trim(txtCurrencyType.Text)) > 0 Then
                    Call txtCurrencyType_Validating(txtCurrencyType, New System.ComponentModel.CancelEventArgs(False))
                    If blnValidCurrency = False Then
                        txtCurrencyType.Focus()
                    Else
                        cmbPOType.Focus()
                    End If
                Else
                    cmbPOType.Focus()
                End If
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
    Private Sub txtCurrencyType_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtCurrencyType.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        Dim rsdb As ClsResultSetDB
        If Trim(txtCurrencyType.Text) = "" Or Len(Trim(txtCurrencyType.Text)) = 0 Then
            GoTo EventExitSub
        End If
        m_strSql = "Select top 1 1 from Currency_mst where unit_code='" & gstrUNITID & "' and Currency_code = '" & txtCurrencyType.Text & "'"
        rsdb = New ClsResultSetDB
        Call rsdb.GetResult(m_strSql)
        If rsdb.GetNoRows = 0 Then
            rsdb.ResultSetClose()
            Call ConfirmWindow(10144, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
            blnValidCurrency = False
            Cancel = True
            GoTo EventExitSub
        Else
            rsdb.ResultSetClose()
            Call SSMaxLength()
        End If
        blnValidCurrency = True
        GoTo EventExitSub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        GoTo EventExitSub
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtCustomerCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCustomerCode.LostFocus
        On Error GoTo ErrHandler
        mvalid = False
        If Len(Trim(txtCustomerCode.Text)) <> 0 Then
            Select Case cmdButtons.Mode
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT, UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                    m_strSql = "Select top 1 1 from cust_Ord_hdr where unit_code='" & gstrUNITID & "' and Account_code ='" & Trim(txtCustomerCode.Text) & "'"
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                    m_strSql = "Select top 1 1 from Customer_Mst where unit_code='" & gstrUNITID & "' and Customer_code ='" & Trim(txtCustomerCode.Text) & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
            End Select
            rsdb = New ClsResultSetDB
            rsdb.GetResult(m_strSql)
            If rsdb.GetNoRows > 0 Then
                txtReferenceNo.Enabled = True
                txtReferenceNo.BackColor = Color.White
                txtConsCode.Enabled = True
                txtConsCode.BackColor = Color.White
                cmdHelp(7).Enabled = True
                txtConsCode.Text = txtCustomerCode.Text
                lblConsDesc.Text = lblCustDesc.Text
                cmdHelp(1).Enabled = True
                Call FillLabel("CUSTOMER")
            Else
                cmdButtons.Focus()
                MsgBox("Customer Code does not exist", vbInformation, "empower")
                txtCustomerCode.Text = ""
                txtCustomerCode.Focus()
                mvalid = True
            End If
            rsdb.ResultSetClose()
        End If
        m_blnCloseFlag = False
        m_blnHelpFlag = False
        If StrComp(Trim(txtCustomerCode.Text), mstrPrevAccountCode, vbTextCompare) <> 0 Then
            Call CLEARVAR()
        End If
        'mvalid = False
        Exit Sub            'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub txtCustomerCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCustomerCode.TextChanged
        On Error GoTo ErrHandler
        Dim rsdb2 As ClsResultSetDB
        Call FillLabel("CUSTOMER")
        Call FillLabel("CURRENCY")
        If Len(Trim(txtCustomerCode.Text)) <> 0 Then
            m_strSql = "Select top 1 1 from cust_Ord_hdr where unit_code='" & gstrUNITID & "' and account_code ='" & txtCustomerCode.Text & "'"
            rsdb2 = New ClsResultSetDB
            rsdb2.GetResult(m_strSql)
            If rsdb2.GetNoRows > 0 Then
                txtReferenceNo.Enabled = True
                txtReferenceNo.BackColor = System.Drawing.Color.White
                cmdHelp(1).Enabled = True
            Else
                txtConsCode.Text = ""
                lblConsDesc.Text = ""
                If txtReferenceNo.Enabled = True Then txtReferenceNo.Text = ""
                txtReferenceNo.Enabled = False
                txtReferenceNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                cmdHelp(1).Enabled = False
            End If
            rsdb2.ResultSetClose()
        Else
            If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                Call RefreshForm()
                lblCustDesc.Text = ""
                txtReferenceNo.Text = ""
                txtAmendmentNo.Text = ""
                cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
            End If
        End If
        ssPOEntry.MaxRows = 0

        txtShipAddress.Text = ""
        lblShipAddress_Details.Text = ""
        chkShipAddress.Checked = False

        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub txtCustomerCode_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCustomerCode.Enter
        mstrPrevAccountCode = Trim(txtCustomerCode.Text)
    End Sub
    Private Sub txtCustomerCode_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCustomerCode.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = 112 Then Call cmdHelp_Click(cmdHelp.Item(0), New System.EventArgs()) 'Help should be invoked if F1 key is pressed
        If KeyCode = System.Windows.Forms.Keys.Return Then
            Call txtCustomerCode_Validating(txtCustomerCode, New System.ComponentModel.CancelEventArgs())
            If mvalid = False Then
                If txtReferenceNo.Enabled = True Then txtReferenceNo.Focus() Else cmdButtons.Focus()
            Else
                txtCustomerCode.Focus()
            End If
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub txtCustomerCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCustomerCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
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
    Private Sub txtCustomerCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtCustomerCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        Dim rsCD As ClsResultSetDB
        If m_blnCloseFlag = True Then
            m_blnCloseFlag = False
            GoTo EventExitSub
        End If
        If Len(Trim(txtCustomerCode.Text)) = 0 Then
            GoTo EventExitSub
        Else
            m_strSql = "Select top 1 1 from Customer_mst where unit_code='" & gstrUNITID & "' and customer_Code='" & Trim(txtCustomerCode.Text) & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
            rsCD = New ClsResultSetDB
            rsCD.GetResult(m_strSql)
            If rsCD.GetNoRows = 0 Then
                rsCD.ResultSetClose()
                Call ConfirmWindow(10145, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                txtCustomerCode.Text = ""
                txtReferenceNo.Text = ""
                Cancel = True
                txtCustomerCode.Focus()
                GoTo EventExitSub
            Else
                '10736222-changes done for CT2 ARE 3 functionality - to enable CT2 Reqd flag
                If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                    strSQL = "Select dbo.UDF_IsCT2Customer('" & gstrUNITID & "','" & txtCustomerCode.Text.Trim & "')"
                    ChkCT2Reqd.Enabled = Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSQL))
                End If
                'ISSUE ID : 10763705 
                strSQL = "Select appendsoitem from customer_mst where unit_code='" & gstrUNITID & "' and customer_code='" & txtCustomerCode.Text.Trim & "'"
                mblnappendsoitem_customer = Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSQL))
                'ISSUE ID  :10763705 
                'GST CHANGES
                strSQL = "Select ALLOW_MULTIPLE_HSN_ITEMS from customer_mst where unit_code='" & gstrUNITID & "' and customer_code='" & txtCustomerCode.Text.Trim & "'"
                MBLNALLOW_MULTIPLEHSNITEMS_CUSTOMER = Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSQL))
                'GST CHANGES
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
    Private Sub txtReferenceNo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtReferenceNo.LostFocus
        On Error GoTo ErrHandler
        Dim rsRefNo As ClsResultSetDB
        Dim rsSalesParameter As ClsResultSetDB
        Dim strAns As MsgBoxResult
        Dim varPOFlg As Object
        Dim intAnswer As Short
        Dim intbase As Short
        mvalid = False
        If m_blnCloseFlag = True Then 'incase close button is clicked then exit
            m_blnCloseFlag = False
            Exit Sub
        End If
        If m_blnHelpFlag = True Then 'incase help button is clicked then exit
            m_blnHelpFlag = False
            Exit Sub
        End If
        If Len(Trim(txtReferenceNo.Text)) > 0 Then 'if reference no is not blank
            m_strSql = " Select Account_Code,Cust_Ref,Amendment_No,Order_Date,Amendment_Date,Active_Flag,Currency_Code,Valid_Date,Effect_Date,Term_Payment,Special_Remarks,Pay_Remarks,Price_Remarks,Packing_Remarks,Frieght_Remarks,Transport_Remarks,Octorai_Remarks,Mode_Despatch,Delivery,First_Authorized,Second_Authorized,Third_Authorized,Authorized_Flag,Reason,PO_Type,SalesTax_Type,OpenSO,AddCustSupp,PerValue,InternalSONo,RevisionNo,Surcharge_code,Future_SO,ECESS_Code,Consignee_Code,ADDVAT_TYPE,exportsotype from cust_ord_hdr where unit_code='" & gstrUNITID & "' and cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "' and Account_Code='" & Trim(Me.txtCustomerCode.Text) & "' and Consignee_Code='" & Trim(txtConsCode.Text) & "' "
            rsRefNo = New ClsResultSetDB
            Call rsRefNo.GetResult(m_strSql)
            If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                If rsRefNo.GetNoRows = 1 Then
                    intbase = 1
                End If
                If rsRefNo.GetNoRows > 1 Then
                    intAnswer = MsgBox("Would You Like to Veiw Base SO", MsgBoxStyle.YesNo, "emower")
                    If intAnswer = 6 Then
                        txtAmendmentNo.Enabled = False : txtAmendmentNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        intbase = 1
                    Else
                        txtAmendmentNo.Enabled = True : txtAmendmentNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        intbase = rsRefNo.GetNoRows
                        ''
                        If rsRefNo.GetValue("exportsotype") = "With Pay" Or rsRefNo.GetValue("exportsotype") = "Without Pay" Then
                            cmbExporttype.Enabled = True
                            cmbExporttype.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        End If
                        ''
                        End If
                End If
            Else
                intbase = rsRefNo.GetNoRows
            End If
            rsRefNo.ResultSetClose()
            If intbase = 1 Then 'if only one record for the reference no is existing
                If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                    m_strSql = " Select top 1 1 from cust_ord_hdr where unit_code='" & gstrUNITID & "' and cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "' and Account_Code='" & Trim(Me.txtCustomerCode.Text) & "' and Consignee_Code='" & Trim(txtConsCode.Text) & "'  and Active_Flag='A' and Authorized_flag=1"
                    rsRefNo = New ClsResultSetDB
                    Call rsRefNo.GetResult(m_strSql)
                    If rsRefNo.GetNoRows > 0 Then
                        rsRefNo.ResultSetClose()
                        strAns = ConfirmWindow(10131, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                        If strAns = MsgBoxResult.Yes Then
                            txtAmendmentNo.Enabled = True
                            txtAmendmentNo.BackColor = System.Drawing.Color.White
                            cmdHelp(3).Enabled = False
                            DTAmendmentDate.Value = GetServerDate()
                            Call GetReferenceDetails()
                            ssPOEntry.MaxRows = 1
                            Call SSMaxLength()
                            cmdchangetype.Enabled = True
                            If txtAmendmentNo.Enabled Then txtAmendmentNo.Focus()
                            mvalid = False

                            Exit Sub
                        Else
                            Call GetReferenceDetails()
                            cmdButtons.Revert()
                            cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD) = True
                            cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                            cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                            cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE) = True
                            cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL) = False
                            cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = True
                            Exit Sub
                        End If
                    Else
                        rsRefNo.ResultSetClose()
                        Call ConfirmWindow(10132, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                        Call GetReferenceDetails()
                        mvalid = True
                        cmdButtons.Revert()
                        cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD) = True
                        cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                        cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                        cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE) = False
                        cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL) = False
                        cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                        cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE) = True
                        Exit Sub
                    End If
                ElseIf cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                    m_strSql = " Select top 1 1 from cust_ord_hdr where unit_code='" & gstrUNITID & "' and cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'and Account_Code='" & Trim(Me.txtCustomerCode.Text) & "'  and Consignee_Code='" & Trim(txtConsCode.Text) & "' and Active_Flag='A' and amendment_No =''"
                    rsRefNo = New ClsResultSetDB
                    Call rsRefNo.GetResult(m_strSql)
                    If rsRefNo.GetNoRows > 0 Then
                        rsRefNo.ResultSetClose()
                        m_strSql = " Select top 1 1 from cust_ord_hdr where unit_code='" & gstrUNITID & "' and cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'and Account_Code='" & Trim(Me.txtCustomerCode.Text) & "' and Consignee_Code='" & Trim(txtConsCode.Text) & "' and Active_Flag ='A' and Authorized_flag=0 and amendment_No =''"
                        rsRefNo = New ClsResultSetDB
                        Call rsRefNo.GetResult(m_strSql)
                        If rsRefNo.GetNoRows > 0 Then
                            rsRefNo.ResultSetClose()
                            Call GetReferenceDetails()
                            cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD) = True
                            cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = True
                            cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = True
                            cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE) = True
                            cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL) = False
                            cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = True
                            GrpSODoc.Visible = True
                            GrpSODoc.Enabled = True
                            lblFilePath.Text = ""
                            BtnUpload.Enabled = False
                            btnViewDoc.Enabled = True
                        Else
                            rsRefNo.ResultSetClose()
                            Call ConfirmWindow(10133, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            Call GetReferenceDetails()
                            mvalid = True
                            GrpSODoc.Visible = True
                            GrpSODoc.Enabled = True
                            lblFilePath.Text = ""
                            BtnUpload.Enabled = False
                            btnViewDoc.Enabled = True
                            Exit Sub
                        End If
                    Else
                        rsRefNo.ResultSetClose()
                        Call ConfirmWindow(10134, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                        Call GetReferenceDetails()
                        mvalid = True
                        Exit Sub
                    End If
                End If
                ' If No of records for the reference no is more than 1 that means amendment no exists
            ElseIf intbase > 1 Then
                If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                    rsSalesParameter = New ClsResultSetDB
                    rsSalesParameter.GetResult("Select AppendSOItem from Sales_parameter where unit_code='" & gstrUNITID & "' ")
                    'ISSUE ID : 10763705
                    If rsSalesParameter.GetValue("AppendSOItem") = True And mblnappendsoitem_customer = True Then
                        'ISSUE ID : 10763705
                        rsSalesParameter.ResultSetClose()
                        m_strSql = " Select top 1 1 from cust_ord_hdr where unit_code='" & gstrUNITID & "' and cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "' and Account_Code='" & Trim(Me.txtCustomerCode.Text) & "'  and Consignee_Code='" & Trim(txtConsCode.Text) & "' and Active_Flag ='A' and Authorized_flag=0"
                        rsRefNo = New ClsResultSetDB
                        Call rsRefNo.GetResult(m_strSql)
                        If rsRefNo.GetNoRows > 0 Then 'incase a not authorized amendment exists
                            rsRefNo.ResultSetClose()
                            cmdButtons.Focus()
                            Call ConfirmWindow(10135, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            txtReferenceNo.Text = ""
                            txtReferenceNo.Focus()
                            mvalid = True
                            Exit Sub
                        Else
                            rsRefNo.ResultSetClose()
                            Me.txtAmendmentNo.Enabled = True
                            txtAmendmentNo.BackColor = System.Drawing.Color.White
                            Call GetReferenceDetails()
                            ssPOEntry.MaxRows = 1
                            Call SSMaxLength()
                            Exit Sub
                        End If
                    Else
                        rsSalesParameter.ResultSetClose()
                        Me.txtAmendmentNo.Enabled = True
                        txtAmendmentNo.BackColor = System.Drawing.Color.White
                        Call GetReferenceDetails()
                        ssPOEntry.MaxRows = 1
                        Call SSMaxLength()
                        Exit Sub
                    End If
                Else
                    txtAmendmentNo.Enabled = True
                    txtAmendmentNo.BackColor = System.Drawing.Color.White
                    cmdHelp(3).Enabled = True
                    If txtAmendmentNo.Enabled Then
                        txtAmendmentNo.Focus()
                    End If
                    Exit Sub
                End If
            Else 'If There are no records existing
                If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then 'If Mode is new,then
                    'DTDate.Enabled = True
                    If blnSO_EDITABLE = True Then  '' Added by priti on 25 Jul 2025 
                        DTDate.Enabled = True
                    Else
                        DTDate.Enabled = False
                    End If
                    DTValidDate.Enabled = True
                    DTEffectiveDate.Enabled = True
                    txtCurrencyType.Enabled = True
                    txtCurrencyType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    cmdHelp(2).Enabled = True
                    cmbPOType.Enabled = True
                    cmbPOType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    ssPOEntry.Enabled = True
                    cmdchangetype.Enabled = True
                    txtAmendmentNo.Enabled = False
                    txtAmendmentNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    cmdHelp(3).Enabled = False
                    'for account Plug in
                    '1.S.Tax at Header Lavel
                    '2.Credit Terms at Main Screen
                    txtSTax.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : txtSTax.Enabled = True : cmdHelp(4).Enabled = True
                    If gblnGSTUnit = False Then
                        txtSTax.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : txtSTax.Enabled = True : _cmdHelp_4.Enabled = True
                        txtSChSTax.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : txtSChSTax.Enabled = True : cmdHelp(6).Enabled = True
                        txtECSSTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : txtECSSTaxType.Enabled = True : CmdECSSTaxType.Enabled = True
                    Else
                        txtSTax.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : txtSTax.Enabled = False : _cmdHelp_4.Enabled = False
                        txtSChSTax.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : txtSChSTax.Enabled = False : cmdHelp(6).Enabled = False
                        txtECSSTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : txtECSSTaxType.Enabled = False : CmdECSSTaxType.Enabled = False
                    End If

                    txtCreditTerms.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : txtCreditTerms.Enabled = True : cmdHelp(5).Enabled = True
                    'txtSChSTax.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : txtSChSTax.Enabled = True : cmdHelp(6).Enabled = True
                    'ctlPerValue.Enabled = True
                    If blnNoneditableCreditTerms_onSO Then
                        txtCreditTerms.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : txtCreditTerms.Enabled = False : cmdHelp(5).Enabled = False
                    Else
                        txtCreditTerms.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : txtCreditTerms.Enabled = True : cmdHelp(5).Enabled = True
                    End If
                    ctlPerValue.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    ctlPerValue.Text = 1
                    'txtECSSTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : txtECSSTaxType.Enabled = True : CmdECSSTaxType.Enabled = True
                    chkOpenSo.Enabled = True
                    '10869290
                    If cmbPOType.Text = "V-SERVICE" And (gblnGSTUnit = False Or gstrUNITID = "STH") Then
                        txtService.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        txtService.Enabled = True
                        cmdServiceTax.Enabled = True
                        cmdSBCtax.Enabled = True
                        txtSBC.Enabled = True
                        txtSBC.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        cmdKKCtax.Enabled = True
                        txtKKC.Enabled = True
                        txtKKC.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    End If
                    Call ADDRow()
                    DTDate.Focus()
                    Exit Sub
                Else
                    DTDate.Enabled = False
                    Call ConfirmWindow(10136, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                    txtReferenceNo.Text = ""
                    If txtReferenceNo.Enabled Then txtReferenceNo.Focus()
                    mvalid = True
                    Exit Sub
                End If
            End If
        Else
            Exit Sub
        End If
        mvalid = False
        If StrComp(mstrPrevRefNo, Trim(txtReferenceNo.Text), CompareMethod.Text) <> 0 Then
            Call CLEARVAR()
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub txtReferenceNo_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtReferenceNo.TextChanged
        On Error GoTo ErrHandler
        ssPOEntry.MaxRows = 0
        txtAmendmentNo.Text = ""
        txtAmendmentNo.Enabled = False : txtAmendmentNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        If Len(Trim(txtReferenceNo.Text)) = 0 Then
            If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                Call RefreshForm()
                txtAmendmentNo.Text = ""
                txtAmendmentNo.Enabled = False : txtAmendmentNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                cmdHelp(3).Enabled = False
                cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
            End If
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub txtReferenceNo_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtReferenceNo.Enter
        mstrPrevRefNo = Trim(txtReferenceNo.Text)
    End Sub
    Private Sub TxtReferenceNo_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtReferenceNo.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = 112 Then Call cmdHelp_Click(cmdHelp.Item(1), New System.EventArgs())
        If KeyCode = System.Windows.Forms.Keys.Return Then
            If mvalid = False Then
                If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                    If txtAmendmentNo.Enabled = True Then txtAmendmentNo.Focus() Else cmdButtons.Focus()
                Else
                    If txtAmendmentNo.Enabled = True Then txtAmendmentNo.Focus() Else cmdButtons.Focus()
                End If
            End If
        ElseIf (KeyCode = 39) Or (KeyCode = 34) Or (KeyCode = 96) Then
            KeyCode = 0
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0001_SOUTH_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo ErrHandler
        'Checking the form name in the Windows list
        mdifrmMain.CheckFormName = mintFormIndex
        frmModules.NodeFontBold(Tag) = True
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0001_SOUTH_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Or cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
            If KeyCode = System.Windows.Forms.Keys.Escape Then
                Call cmdButtons_ButtonClick(cmdButtons, New UCActXCtl.UCbtndgrp.ButtonClickEventArgs(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL))
            End If
        End If
        If KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0 Then
            Call ctlHeader_Click(ctlHeader, New System.EventArgs())
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0001_SOUTH_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo ErrHandler
        Dim rsGetDate As ClsResultSetDB
        'Get the index of form in the Windows list
        mintFormIndex = mdifrmMain.AddFormNameToWindowList(ctlHeader.Tag)
        prgItemDetails.Visible = False
        'Load the captions
        Call FillLabelFromResFile(Me)
        'Size the form to client workspace
        Call FitToClient(Me, fraContainer, ctlHeader, cmdButtons, 500)
        prgItemDetails.Left = fraContainer.Left
        'Disabling the controls
        Call EnableControls(False, Me, True)
        'Initialising the buttons
        cmdButtons.Revert()
        'Disabling Edit, Delete and Print buttons
        cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
        cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
        cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
        cmdHelp(0).Enabled = True
        txtCustomerCode.Enabled = True
        '10736222-changes done for CT2 ARE 3 functionality
        ChkCT2Reqd.Enabled = False
        txtCustomerCode.BackColor = System.Drawing.Color.White
        ssPOEntry.Enabled = False
        m_strSql = "Select Financial_EndDate from company_mst where unit_code='" & gstrUNITID & "'"
        rsGetDate = New ClsResultSetDB
        rsGetDate.GetResult(m_strSql)
        DTDate.CustomFormat = gstrDateFormat
        DTDate.Format = DateTimePickerFormat.Custom
        DTEffectiveDate.CustomFormat = gstrDateFormat
        DTEffectiveDate.Format = DateTimePickerFormat.Custom
        DTValidDate.CustomFormat = gstrDateFormat
        DTValidDate.Format = DateTimePickerFormat.Custom
        DTAmendmentDate.CustomFormat = gstrDateFormat
        DTAmendmentDate.Format = DateTimePickerFormat.Custom
        With ssPOEntry
            .Col = 13
            .Col2 = 12
            .ColHidden = True
            .Col = 14
            .Col2 = 13
            .ColHidden = True
            .Col = 15
            .Col2 = 14
            .ColHidden = True
            .Col = 16
            .Col2 = 15
            .ColHidden = True
            .Col = 17
            .Col2 = 16
            .ColHidden = True
            .Col = 18
            .Col2 = 17
            .ColHidden = True
        End With
        DTDate.Value = GetServerDate()
        DTEffectiveDate.Value = GetServerDate()
        DTAmendmentDate.Value = GetServerDate()
        DTValidDate.Value = rsGetDate.GetValue("Financial_EndDate")
        rsGetDate.ResultSetClose()
        Call SSMaxLength()
        m_blnHelpFlag = False
        m_blnCloseFlag = False
        Call AddPOType()
        Call addExportSotype()
        cmbPOType.Text = ObsoleteManagement.GetItemString(cmbPOType, 0)
        m_strSalesTaxType = ""
        Me.KeyPreview = True
        If gblnGSTUnit = True Then
            _cmdHelp_4.Enabled = False : txtSTax.Enabled = False : txtSTax.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            txtSChSTax.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : txtSChSTax.Enabled = False : cmdHelp(6).Enabled = False
        End If
        blnSO_EDITABLE = Convert.ToBoolean(SqlConnectionclass.ExecuteScalar("Select isnull(SO_EDITABLE,0) from Sales_parameter where  unit_code='" & gstrUNITID & "'"))
        If blnSO_EDITABLE = True Then
            DTDate.Enabled = True
        Else
            DTDate.Enabled = False
        End If

        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0001_SOUTH_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        On Error GoTo ErrHandler
        Dim enmValue As eMPowerFunctions.ConfirmWindowReturnEnum
        If UnloadMode >= 0 And UnloadMode <= 5 Then
            If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Or cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                enmValue = ConfirmWindow(10055, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNOCANCEL, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION)
                If enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_NO Or enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                    If enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                        'Save the data before closing
                        Call cmdButtons_ButtonClick(cmdButtons, New UCActXCtl.UCbtndgrp.ButtonClickEventArgs(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE))
                    Else
                        gblnCancelUnload = False : gblnFormAddEdit = False
                    End If
                Else
                    'Set the global variable
                    gblnCancelUnload = True : gblnFormAddEdit = True
                End If
            End If
        End If
        'Checking the status
        If gblnCancelUnload = True Then
            eventArgs.Cancel = True
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub frmMKTTRN0001_SOUTH_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error GoTo ErrHandler
        Me.Dispose()
        'Removing the form name from list
        mdifrmMain.RemoveFormNameFromWindowList = mintFormIndex
        'Setting the corresponding node's tag
        frmModules.NodeFontBold(Tag) = False
        'Releasing the form reference
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Function ValidRowData(ByVal Row As Integer, Optional ByRef Col As Integer = 0) As Boolean
        '---------------------------------------------------------------
        'Arguments      :Nil
        'Return Value   :Nil
        'Function       : To validate the row data
        '---------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim rsitem As ClsResultSetDB
        Dim varQty, varDrawPartNo, varItemCode, varOpenSO As Object
        Dim varCustSuppMat, varRate, varToolCost As Object
        Dim varflg, varPF, varStax, varExd, varSSt, varOthers, varDespatch As Object
        Dim dummyVarItem As Object
        Dim varDelFlag As Object
        Dim intRow As Short ' to get the values in the grid
        Dim strMeasure As String
        Dim rsMeasure As ClsResultSetDB
        Dim emptyvar As Object
        Dim strtest As String
        Dim valAddExcise As Object
        Dim VARfirstitemcode As Object
        Dim VARfirstCUSTITEMCODE As Object
        'GST CHANGE
        Dim varHSNCode As Object
        Dim varHSNSACCode As Object
        Dim VARCGSTTAX As Object
        Dim VARSGSTTAX As Object
        Dim VARUTGSTTAX As Object
        Dim VARIGSTTAX As Object

        'GST CHANGE

        'change account Plug in
        varDelFlag = Nothing
        Call ssPOEntry.GetText(0, Row, varDelFlag)
        ssPOEntry.Row = Row
        ssPOEntry.Col = 1
        varOpenSO = ssPOEntry.Value
        varDrawPartNo = Nothing
        Call ssPOEntry.GetText(2, Row, varDrawPartNo)
        varItemCode = Nothing
        Call ssPOEntry.GetText(4, Row, varItemCode)
        varQty = Nothing
        Call ssPOEntry.GetText(5, Row, varQty)
        varRate = Nothing
        Call ssPOEntry.GetText(6, Row, varRate)
        varCustSuppMat = Nothing
        Call ssPOEntry.GetText(7, Row, varCustSuppMat)
        varToolCost = Nothing
        Call ssPOEntry.GetText(8, Row, varToolCost)
        varPF = Nothing
        Call ssPOEntry.GetText(9, Row, varPF)
        varExd = Nothing
        Call ssPOEntry.GetText(10, Row, varExd)
        'GST CHANGES
        varHSNSACCode = Nothing
        Call ssPOEntry.GetText(20, Row, varHSNSACCode)
        VARCGSTTAX = Nothing
        Call ssPOEntry.GetText(21, Row, VARCGSTTAX)
        VARSGSTTAX = Nothing
        Call ssPOEntry.GetText(22, Row, VARSGSTTAX)
        VARUTGSTTAX = Nothing
        Call ssPOEntry.GetText(23, Row, VARUTGSTTAX)
        VARIGSTTAX = Nothing
        Call ssPOEntry.GetText(24, Row, VARIGSTTAX)
        'GST CHANGES


        If gblnGSTUnit = True And ssPOEntry.MaxRows >= 1 And UCase(Trim(cmbPOType.Text)) <> "EXPORT" And varDrawPartNo <> "" Then
            If (varHSNSACCode = "") Then
                ValidRowData = False
                MsgBox("HSN/SAC CODE SHOULD NOT BE BLANK", MsgBoxStyle.OkOnly, ResolveResString(100))
                Call ssSetFocus(ssPOEntry.MaxRows, 3)
                ssPOEntry.Focus()
                Exit Function
            End If
            If (VARCGSTTAX = "" And VARSGSTTAX = "" And VARUTGSTTAX = "" And VARIGSTTAX = "") And varDrawPartNo <> "" Then
                ValidRowData = False
                MsgBox("ATLEAST ONE TAX SHOULD BE CONSIDERED", MsgBoxStyle.OkOnly, ResolveResString(100))
                Call ssSetFocus(ssPOEntry.MaxRows, 3)
                ssPOEntry.Focus()
                Exit Function
            End If
        End If

        'GST CHANGES

        valAddExcise = Nothing
        Call ssPOEntry.GetText(11, Row, valAddExcise)
        If (gblnGSTUnit = False And gstrUNITID <> "STH") And Val(varCustSuppMat) > 0 And Len((Trim(valAddExcise))) = 0 And ssPOEntry.ActiveCol = 12 Then
            ValidRowData = False
            MsgBox("Please Enter Additional Excise Duty.")
            ssPOEntry.Col = 11
            ssPOEntry.Row = ssPOEntry.ActiveRow
            ssPOEntry.Action = FPSpreadADO.ActionConstants.ActionActiveCell
            Exit Function
        End If
        varOthers = Nothing
        Call ssPOEntry.GetText(11, Row, varOthers)
        If varDelFlag = "*" Then
            ValidRowData = True
            Exit Function
        End If
        If Col = 0 Or Col = 2 Then ' if col is 2 or entire row
            If Len(Trim(varDrawPartNo)) = 0 Then
                ValidRowData = False
                Call ssPOEntry.SetText(4, Row, "")
                Call ssPOEntry.SetText(2, Row, "")
                Call ssSetFocus(Row)
                ssPOEntry.Focus()
                Exit Function
            End If
            '10869290
            If Mid(cmbPOType.Text, 1, 1) = "V" Then
                m_strSql = "Select top 1 1 from custitem_mst A, Item_Mst B where A.Unit_Code=B.Unit_Code and A.Item_Code=B.Item_Code and B.Status='A' and B.Hold_Flag=0 and B.Item_Main_Grp='M' and A.unit_code='" & gstrUNITID & "' and  Cust_Drgno = '" & Trim(varDrawPartNo) & "' and account_code='" & Trim(txtCustomerCode.Text) & "' and A.active = 1"
            Else
                m_strSql = "Select top 1 1 from custitem_mst where unit_code='" & gstrUNITID & "' and  Cust_Drgno = '" & Trim(varDrawPartNo) & "' and account_code='" & Trim(txtCustomerCode.Text) & "' and active = 1" '10532789 active = 1 added
            End If

            If IsNothing(rsitem) = False Then rsitem.ResultSetClose()
            rsitem = New ClsResultSetDB
            rsitem.GetResult(m_strSql)
            If rsitem.GetNoRows <= 0 Then
                rsitem.ResultSetClose()
                'MsgBox("Please check reference of this  item in  " & vbCrLf & "Item Master or Customer Item Master or in Item Rate Master")
                MsgBox("Please check reference of this  Cust. Part: " + " '" & Trim(varDrawPartNo) & "' " + "  " & vbCrLf & "Item Code:" + " '" & Trim(varItemCode) & "' " + " " & vbCrLf & "in Item Master or Customer Item Master or in Item Rate Master")
                ValidRowData = False
                Call ssPOEntry.SetText(2, Row, "")
                Call ssPOEntry.SetText(4, Row, "")
                Call ssSetFocus(Row)
                ssPOEntry.Focus()
                Exit Function
            End If
            m_strSql = "Select top 1 1 from cust_ord_dtl where unit_code='" & gstrUNITID & "' and Cust_Drgno = '" & Trim(varDrawPartNo) & "'and account_code='" & txtCustomerCode.Text & "' and cust_ref='" & txtReferenceNo.Text & "'and active_Flag='A' and ITem_code = '" & varItemCode & "' and amendment_no = '" & txtAmendmentNo.Text & "'"
            If IsNothing(rsitem) = False Then rsitem.ResultSetClose()
            rsitem = New ClsResultSetDB
            rsitem.GetResult(m_strSql)
            If rsitem.GetNoRows > 1 Then
                rsitem.ResultSetClose()
                'Call ConfirmWindow(10069, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                MsgBox("Item Code:" + " '" & Trim(varItemCode) & "' " + " Already Exists please enter another Item Code")
                ValidRowData = False
                Call ssPOEntry.SetText(2, Row, "")
                Call ssPOEntry.SetText(4, Row, "")
                Call ssSetFocus(Row)
                ssPOEntry.Focus()
                Exit Function
            End If
            For intRow = 1 To ssPOEntry.MaxRows
                'change account Plug in
                dummyVarItem = Nothing
                Call ssPOEntry.GetText(2, intRow, dummyVarItem)
                If dummyVarItem = varDrawPartNo And intRow <> Row Then
                    'Call ConfirmWindow(10156, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                    MsgBox("Customer Item Code:" + " '" & Trim(dummyVarItem) & "' " + "This Item Already Exists in the List ")
                    ValidRowData = False
                    Call ssSetFocus(Row)
                    ssPOEntry.Focus()
                    emptyvar = ""
                    If intRow > Row Then
                        'change account Plug in
                        Call ssPOEntry.SetText(2, intRow, emptyvar)
                        Call ssPOEntry.SetText(4, intRow, emptyvar)
                    Else
                        Call ssPOEntry.SetText(2, Row, emptyvar)
                        Call ssPOEntry.SetText(4, Row, emptyvar)
                    End If
                    Call ssSetFocus(Row, 2)
                    Exit Function
                End If
            Next

            'GST CHANGES
            If gblnGSTUnit = True Then
                VARfirstitemcode = Nothing
                Call ssPOEntry.GetText(4, 1, VARfirstitemcode)
                varHSNCode = Find_Value("SELECT HSN_SAC_CODE FROM ITEM_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND ITEM_CODE='" & VARfirstitemcode & "'")
            End If
            'GST CHANGES

            Dim dummyVarItemHSNCODE As Object
            strSQL = "Select ALLOW_MULTIPLE_HSN_ITEMS from customer_mst where unit_code='" & gstrUNITID & "' and customer_code='" & txtCustomerCode.Text.Trim & "'"
            MBLNALLOW_MULTIPLEHSNITEMS_CUSTOMER = Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSQL))

            If gblnGSTUnit = True And MBLNALLOW_MULTIPLEHSNITEMS_CUSTOMER = False And ssPOEntry.MaxRows > 1 Then
                For intRow = 1 To ssPOEntry.MaxRows
                    dummyVarItemHSNCODE = Nothing
                    Call ssPOEntry.GetText(4, intRow, dummyVarItemHSNCODE)
                    dummyVarItemHSNCODE = Find_Value("SELECT HSN_SAC_CODE FROM ITEM_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND ITEM_CODE='" & dummyVarItemHSNCODE & "'")
                    If dummyVarItemHSNCODE <> "" And dummyVarItemHSNCODE <> varHSNCode Then
                        MsgBox("MULTPLE HSN NOT ALLOWED , PREVIOUS ITEM LINKED WITH HSN CODE :" + varHSNCode)
                        ValidRowData = False
                        Call ssSetFocus(Row)
                        ssPOEntry.Focus()
                        emptyvar = ""
                        Exit Function
                    End If
                Next
            End If
            'GST CHANGES



            'GST CHANGES

            'DOCKCODE 
            Dim strdockcode As String
            strSQL = "select dbo.UDF_ISDOCKCODE_ASN( '" & gstrUNITID & "','" & txtCustomerCode.Text.Trim & "')"
            If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSQL)) = True Then
                If Col = 0 Or Col = 2 Then ' if col is 3rd
                    For intRow = 1 To ssPOEntry.MaxRows
                        If intRow = 1 Then
                            VARfirstCUSTITEMCODE = Nothing
                            Call ssPOEntry.GetText(2, 1, VARfirstCUSTITEMCODE)
                            VARfirstitemcode = Nothing
                            Call ssPOEntry.GetText(4, 1, VARfirstitemcode)
                        End If

                        varDrawPartNo = Nothing
                        Call ssPOEntry.GetText(2, Row, varDrawPartNo)

                        If intRow = 1 Then 'FUNCTIONALITY IS ON 
                            strSQL = "SELECT DOCKCODE FROM CUSTITEM_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND ACCOUNT_CODE='" & txtCustomerCode.Text & "'" & _
                                             " AND ITEM_CODE='" & VARfirstitemcode & "' AND CUST_DRGNO='" & VARfirstCUSTITEMCODE & "' AND ACTIVE=1 "
                            strdockcode = Convert.ToString(SqlConnectionclass.ExecuteScalar(strSQL))
                        Else
                            strSQL = "SELECT TOP 1 1 FROM CUSTITEM_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND ACCOUNT_CODE='" & txtCustomerCode.Text & "'" & _
                                     " AND CUST_DRGNO='" & varDrawPartNo & "' AND ACTIVE=1 " & _
                                    "AND DOCKCODE IN ( '" & strdockcode & "' )"
                            If Not IsRecordExists(strSQL) Then
                                MsgBox("Part Code:" + " '" & Trim(varItemCode) & "' " + " Should be same Dock Code : (" & strdockcode & ")")
                                ValidRowData = False
                                Call ssPOEntry.SetText(2, Row, "")
                                Call ssPOEntry.SetText(4, Row, "")
                                Call ssSetFocus(Row)
                                ssPOEntry.Focus()
                                Exit Function
                            End If
                        End If

                    Next
                End If

            End If


        End If

        'DOCKCODE END 
        If Col = 0 Or Col = 4 Then ' if col is 3rd
            If Len(Trim(varItemCode)) = 0 Then
                ValidRowData = False
                Call ssPOEntry.SetText(2, Row, "")
                Call ssSetFocus(Row)
                ssPOEntry.Focus()
                Exit Function
            End If

            If IsNothing(rsitem) = False Then rsitem.ResultSetClose()
            m_strSql = "Select top 1 1 from custitem_mst (nolock) where unit_code='" & gstrUNITID & "' and Cust_Drgno = '" & Trim(varDrawPartNo) & "' and Account_Code='" & Trim(txtCustomerCode.Text) & "' and Item_code ='" & Trim(varItemCode) & "' and active = 1" '10532789 active = 1 added
            rsitem = New ClsResultSetDB
            rsitem.GetResult(m_strSql)
            If rsitem.GetNoRows <= 0 Then
                rsitem.ResultSetClose()
                'MsgBox("Please check refrence of this  item in  " & vbCrLf & "Item Master or Customer Item Master or in Item Rate Master")
                MsgBox("Please check reference of this  Cust. Part: " + " '" & Trim(varDrawPartNo) & "' " + "  " & vbCrLf & "Item Code:" + " '" & Trim(varItemCode) & "' " + " " & vbCrLf & "in Item Master or Customer Item Master or in Item Rate Master")
                ValidRowData = False
                Call ssSetFocus(Row)
                ssPOEntry.Focus()
                Exit Function
            End If
            For intRow = 1 To ssPOEntry.MaxRows
                dummyVarItem = Nothing
                Call ssPOEntry.GetText(4, intRow, dummyVarItem)
            Next
        End If
        If (Col = 0 Or Col = 5) Then
            If chkOpenSo.CheckState = 0 Then
                If varOpenSO = 0 Then
                    If varQty <= 0 Or Val(Trim(varQty)) <= 0 Then
                        'Call ConfirmWindow(10224, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                        MsgBox("For Item: " + " '" & Trim(dummyVarItem) & "' " + "Quantity is Zero which can't be Zero")
                        ValidRowData = False
                        ssPOEntry.Col = 5
                        Call ssSetFocus(Row, 5)
                        ssPOEntry.Focus()
                        Exit Function
                    End If
                    If Col = 5 Then
                        varDespatch = Nothing
                        Call ssPOEntry.GetText(17, Row, varDespatch)
                        If Val(varDespatch) > 0 Then
                            If MsgBox("Despatch Qty for item: " + " '" & Trim(dummyVarItem) & "' " + " is [ " & varDespatch & " ] would you like to add this Quantity You have entered.", MsgBoxStyle.YesNo, "empower") = MsgBoxResult.Yes Then
                                Call ssPOEntry.SetText(5, Row, varQty + varDespatch)
                                varQty = Nothing
                                Call ssPOEntry.GetText(5, Row, varQty)
                            End If
                        End If
                    End If
                    'CHECK FOR MEASURMENT UNIT
                    strMeasure = "select a.Decimal_allowed_flag from Measure_Mst a,Item_Mst b"
                    strMeasure = strMeasure & " where a.unit_code=b.unit_code and b.cons_Measure_Code=a.Measure_Code and b.Item_Code = '" & varItemCode & "' and a.unit_code='" & gstrUNITID & "'"
                    If IsNothing(rsMeasure) = False Then rsMeasure.ResultSetClose()
                    rsMeasure = New ClsResultSetDB
                    rsMeasure.GetResult(strMeasure, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                    If rsMeasure.GetValue("Decimal_allowed_flag") = False Then
                        rsMeasure.ResultSetClose()
                        If System.Math.Round(varQty, 3) - Val(varQty) <> 0 Then
                            'Call ConfirmWindow(10455, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            MsgBox("The Quantity of Item:" + " '" & Trim(varItemCode) & "' " + "is not defined in Decimals/Fractions.")
                            ValidRowData = False
                            'change account Plug in
                            Call ssPOEntry.SetText(5, Row, CShort(varQty))
                            ssPOEntry.Col = 5
                            Call ssSetFocus(Row, 5)
                            ssPOEntry.Focus()
                            Exit Function
                        End If
                    End If
                    If varQty > 9999999 Then
                        MsgBox("Enter value less than 9999999 OR Make it Open Item." & vbCrLf & "Item Code:" + " '" & Trim(varItemCode) & "' ")
                        ValidRowData = False
                        ssPOEntry.Col = 5
                        ssPOEntry.Row = Row : ssPOEntry.Text = CStr(0)
                        Call ssSetFocus(Row, 5)
                        ssPOEntry.Focus()
                        Exit Function
                    End If
                End If
            End If 'Flag Check
        End If
        'change account Plug in
        If Col = 0 Or Col = 6 Then
            If (varRate = 0 Or Len(Trim(varRate)) = 0) And Len(varItemCode) > 1 Then
                MsgBox("Enter Rate Greater then 0", MsgBoxStyle.OkOnly, "empoer")
                ValidRowData = False
                Call ssSetFocus(Row, 6)
                ssPOEntry.Focus()
                Exit Function
            End If
            If varQty < 0 And Len(varItemCode) > 1 Then
                'change account Plug in
                MsgBox("Enter Rate Greater then 0", MsgBoxStyle.OkOnly, "empoer")
                ValidRowData = False
                ssPOEntry.Col = 6
                Call ssSetFocus(Row, 6)
                ssPOEntry.Focus()
                Exit Function
            End If
        End If

        If Col = 0 Or Col = 10 Then ' if col is 5
            'gst changes
            If Len(Trim(varExd)) = 0 Then
                If gblnGSTUnit = False And gstrUNITID <> "STH" Then
                    'gst changes
                    MsgBox("Excise Duty cannot be blank", MsgBoxStyle.Information, "empower")
                    ValidRowData = False
                    ssPOEntry.Col = 10
                    Call ssSetFocus(Row, 10)
                    ssPOEntry.Focus()
                    Exit Function
                End If
            End If
        End If

        Dim rsSalesParameter As ClsResultSetDB
        If Col = 0 Or Col = 10 Then
            With ssPOEntry
                .Row = Row : .Col = 10
                If Len(Trim(.Text)) > 0 Then
                    If gblnGSTUnit = False Then
                        If IsNothing(rsSalesParameter) = False Then rsSalesParameter.ResultSetClose()
                        rsSalesParameter = New ClsResultSetDB
                        rsSalesParameter.GetResult("Select Txrt_Rate_no,TxRt_Percentage from Gen_TaxRate where unit_code='" & gstrUNITID & "' and Tx_TaxeID = 'EXC' and Txrt_Rate_no = '" & Replace(Trim(.Text), "'", "") & "'")
                        If rsSalesParameter.GetNoRows = 0 Then
                            MsgBox("Invalid Excise Code.", vbInformation, "empower")
                            .Row = Row : .Col = 10
                            .Text = ""
                            .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                            .Focus()
                            ValidRowData = False
                            rsSalesParameter.ResultSetClose()
                            rsSalesParameter = Nothing
                            Exit Function
                        End If
                        rsSalesParameter.ResultSetClose()
                        rsSalesParameter = Nothing
                    End If
                End If

            End With
        End If
        If Col = 0 Or Col = 11 Then
            With ssPOEntry
                .Row = Row : .Col = 11
                If Len(Trim(.Text)) > 0 And (gblnGSTUnit = False And gstrUNITID <> "STH") Then
                    If IsNothing(rsSalesParameter) = False Then rsSalesParameter.ResultSetClose()
                    rsSalesParameter = New ClsResultSetDB
                    rsSalesParameter.GetResult("Select Txrt_Rate_no,TxRt_Percentage from Gen_TaxRate where unit_code='" & gstrUNITID & "' and Tx_TaxeID = 'AED' and Txrt_Rate_no = '" & Replace(Trim(.Text), "'", "") & "'")
                    If rsSalesParameter.GetNoRows = 0 Then
                        MsgBox("Invalid Addtional Excise Code.", vbInformation, "empower")
                        .Row = Row : .Col = 11
                        .Text = ""
                        .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                        .Focus()
                        ValidRowData = False
                        rsSalesParameter.ResultSetClose()
                        rsSalesParameter = Nothing
                        Exit Function
                    End If
                    rsSalesParameter.ResultSetClose()
                    rsSalesParameter = Nothing
                End If
            End With
        End If
        ValidRowData = True
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Function
    End Function
    Private Sub ssSetFocus(ByRef Row As Integer, Optional ByRef Col As Integer = 3)
        '----------------------------------------------------------------------------
        'Argument       :   Row, Col
        'Return Value   :   Nil
        'Function       :   Set the focus according to row and col value
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        With ssPOEntry
            '12/10/2002 if condition added by nisha
            If .Enabled = True Then
                .Row = Row
                .Col = Col
                .Action = 0
                .Focus()
            End If
        End With
    End Sub
    Public Sub FillLabel(ByRef pstrCode As String)
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
                m_strSql = "select Cust_Name,Credit_Days from Customer_mst where unit_code='" & gstrUNITID & "' and Customer_code='" & txtCustomerCode.Text & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
                rsCust = New ClsResultSetDB
                rsCust.GetResult(m_strSql)
                lblCustDesc.ForeColor = System.Drawing.Color.White
                lblCustDesc.Text = IIf(UCase(rsCust.GetValue("Cust_Name")) = "UNKNOWN", "", rsCust.GetValue("Cust_Name"))
                txtCreditTerms.Text = IIf(UCase(rsCust.GetValue("Credit_Days")) = "UNKNOWN", "", rsCust.GetValue("Credit_Days"))
                rsCust.ResultSetClose()
                blnNoneditableCreditTerms_onSO = Convert.ToBoolean(SqlConnectionclass.ExecuteScalar("Select isnull(NoneditableCreditTerms_onSO,0) from customer_mst where  unit_code='" & gstrUNITID & "' and Customer_code='" & txtCustomerCode.Text & "'"))
            Case "STAX"
                m_strSql = "select TxRt_Rate_No,TxRt_RateDesc from Gen_TaxRate where unit_code='" & gstrUNITID & "' and Txrt_Rate_No = '" & txtSTax.Text & "'"
                rsCust = New ClsResultSetDB
                rsCust.GetResult(m_strSql)
                lblSTaxDesc.ForeColor = System.Drawing.Color.White
                lblSTaxDesc.Text = IIf(UCase(rsCust.GetValue("TxRt_RateDesc")) = "UNKNOWN", "", rsCust.GetValue("TxRt_RateDesc"))
                rsCust.ResultSetClose()
                '1.Surcharge on S.Tax
            Case "SSCHTAX"
                m_strSql = "SELECT TxRt_Rate_no, TxRt_RateDesc FROM gen_taxrate where unit_code='" & gstrUNITID & "' and Txrt_Rate_No= '" & txtSChSTax.Text & "'"
                rsCust = New ClsResultSetDB
                rsCust.GetResult(m_strSql)
                lblSChSTaxDesc.ForeColor = System.Drawing.Color.White
                lblSChSTaxDesc.Text = IIf(UCase(rsCust.GetValue("TxRt_RateDesc")) = "UNKNOWN", "", rsCust.GetValue("TxRt_RateDesc"))
                rsCust.ResultSetClose()
            Case "CREDIT"
                m_strSql = "select crtrm_TermID,crtrm_Desc from Gen_CreditTrmMaster where unit_code='" & gstrUNITID & "' and crtrm_TermID = '" & txtCreditTerms.Text & "'"
                rsCust = New ClsResultSetDB
                rsCust.GetResult(m_strSql)
                lblCreditTermDesc.ForeColor = System.Drawing.Color.White
                lblCreditTermDesc.Text = IIf(UCase(rsCust.GetValue("crtrm_Desc")) = "UNKNOWN", "", rsCust.GetValue("crtrm_Desc"))
                rsCust.ResultSetClose()
            Case "CURRENCY"
                m_strSql = "select Cust_Name,currency_code from Customer_mst where unit_code='" & gstrUNITID & "' and Customer_code='" & txtCustomerCode.Text & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
                rsCust = New ClsResultSetDB
                rsCust.GetResult(m_strSql)
                lblCustDesc.ForeColor = System.Drawing.Color.White
                txtCurrencyType.Text = IIf(UCase(rsCust.GetValue("currency_code")) = "UNKNOWN", "", rsCust.GetValue("currency_code"))
                rsCust.ResultSetClose()
                '10869290
            Case "SRT"
                m_strSql = "select TxRt_Rate_No,TxRt_RateDesc from Gen_TaxRate where unit_code='" & gstrUNITID & "' and Txrt_Rate_No = '" & txtService.Text & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
                rsCust = New ClsResultSetDB
                rsCust.GetResult(m_strSql)
                lblservicedesc.ForeColor = System.Drawing.Color.White
                lblservicedesc.Text = IIf(UCase(rsCust.GetValue("TxRt_RateDesc")) = "UNKNOWN", "", rsCust.GetValue("TxRt_RateDesc"))
                rsCust.ResultSetClose()
        End Select
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub cmbPOType_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmbPOType.Leave
        On Error GoTo ErrHandler
        If m_blnChangeFormFlg = True Then
            frmMKTTRN0010.ShowDialog()
        End If
        If cmbPOType.Text = "" Or Len(Trim(cmbPOType.Text)) = 0 Then
            Call ConfirmWindow(10001, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
            cmbPOType.Enabled = True
            cmbPOType.BackColor = System.Drawing.Color.White
            cmbPOType.Focus()
            Exit Sub
        End If
        '10869290
        If Not (UCase(Trim(cmbPOType.Text)) = "OEM" Or UCase(Trim(cmbPOType.Text)) = "JOB WORK" Or UCase(Trim(cmbPOType.Text)) = "SPARES" Or UCase(Trim(cmbPOType.Text)) = "EXPORT" Or UCase(Trim(cmbPOType.Text)) = "V-SERVICE") Then
            MsgBox("Please Enter valid P.O. Type (OEM,J,S,E,V)", MsgBoxStyle.Information, "empower")
            cmbPOType.Enabled = True
            cmbPOType.BackColor = System.Drawing.Color.White
            cmbPOType.Focus()
            Exit Sub
        End If
        If cmbPOType.SelectedIndex = 4 Then
            cmdchangetype.Visible = False
        Else
            cmdchangetype.Visible = True
        End If
        If cmbPOType.Text.ToUpper = "EXPORT" Then
            cmbExporttype.Enabled = True
            cmbExporttype.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        Else
            cmbExporttype.Enabled = False
            cmbExporttype.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Public Sub GetAmendmentDetails()
        '----------------------------------------------------------------------------
        'Argument       :   Nil
        'Return Value   :   Nil
        'Function       :   Get the details if there is an amendment
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        m_blnGetAmendmentDetails = True
        Dim rsAD As ClsResultSetDB
        Dim strAuthFlg As String
        m_strSql = "select Account_Code,Cust_Ref,Amendment_No,Item_Code,Rate,Order_Qty,Despatch_Qty,Active_Flag,Cust_Mtrl,Cust_DrgNo,Packing,Others,Excise_Duty,Cust_Drg_Desc,Tool_Cost,Authorized_flag,OpenSO,SalesTax_Type,PerValue,InternalSONo,RevisionNo,Remarks,MRP,Abantment_Code,AccessibleRateforMRP,Packing_Type,TOOL_AMOR_FLAG,ShowInAuth,ADD_Excise_Duty ,external_salesorder_no,HSNSACCODE,  ISHSNORSAC, CGSTTXRT_TYPE,SGSTTXRT_TYPE, UTGSTTXRT_TYPE,IGSTTXRT_TYPE,COMPENSATION_CESS  from cust_ord_dtl where unit_code='" & gstrUNITID & "' and Cust_ref='" & txtReferenceNo.Text & " 'and account_Code='" & txtCustomerCode.Text & "'and amendment_No='" & txtAmendmentNo.Text & "' order by Cust_Drgno"
        rsdb = New ClsResultSetDB
        rsdb.GetResult(m_strSql)
        m_strSql = "select Account_Code,Cust_Ref,Amendment_No,Order_Date,Amendment_Date,Active_Flag,Currency_Code,Valid_Date,Effect_Date,Term_Payment,Special_Remarks,Pay_Remarks,Price_Remarks,Packing_Remarks,Frieght_Remarks,Transport_Remarks,Octorai_Remarks,Mode_Despatch,Delivery,First_Authorized,Second_Authorized,Third_Authorized,Authorized_Flag,Reason,PO_Type,SalesTax_Type,OpenSO,AddCustSupp,PerValue,InternalSONo,RevisionNo,Surcharge_code,Future_SO,ECESS_Code,Consignee_Code,ADDVAT_TYPE,CT2_Reqd_In_SO,exportsotype from cust_ord_hdr where unit_code='" & gstrUNITID & "' and Cust_ref='" & txtReferenceNo.Text & " 'and account_Code='" & txtCustomerCode.Text & "' and Consignee_Code='" & Trim(Me.txtConsCode.Text) & "'  and amendment_No='" & txtAmendmentNo.Text & "'"
        rsAD = New ClsResultSetDB
        rsAD.GetResult(m_strSql)
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        If rsAD.GetNoRows > 0 Then
            rsAD.MoveFirst()
            If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                DTDate.Value = rsAD.GetValue("Order_Date")
            End If

            '10736222-changes done for CT2 ARE 3 functionality
            ChkCT2Reqd.Checked = rsAD.GetValue("CT2_Reqd_In_SO")

            lblIntSONoDes.Text = rsAD.GetValue("InternalSONo")
            lblRevisionNo.Text = rsAD.GetValue("RevisionNo")
            DTAmendmentDate.Value = rsAD.GetValue("Amendment_Date")
            DTEffectiveDate.Value = rsAD.GetValue("Effect_Date")
            DTValidDate.Value = rsAD.GetValue("Valid_Date")
            txtCurrencyType.Text = rsAD.GetValue("Currency_Code")
            ctlPerValue.Text = rsAD.GetValue("PerValue")
            txtAmendReason.Text = rsAD.GetValue("Reason")
            strpotype = rsAD.GetValue("PO_Type")
            strexportType = rsAD.GetValue("ExportSoType")
            If strexportType = "WP" Then
                cmbExporttype.SelectedIndex = 1
            ElseIf strexportType = "WOP" Then
                cmbExporttype.SelectedIndex = 2
            Else
                cmbExporttype.SelectedIndex = 0
            End If
            Select Case strpotype
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
            txtSTax.Text = rsAD.GetValue("salestax_Type")
            txtSChSTax.Text = rsAD.GetValue("Surcharge_code")
            txtECSSTaxType.Text = rsAD.GetValue("ECESS_Code")
            If rsAD.GetValue("OpenSO") = False Then
                chkOpenSo.CheckState = System.Windows.Forms.CheckState.Unchecked
            Else
                chkOpenSo.CheckState = System.Windows.Forms.CheckState.Checked
            End If
            txtCreditTerms.Text = rsAD.GetValue("Term_Payment")
            rsdb.MoveFirst()
            ssPOEntry.MaxRows = 0
            'we add rows to it
            intMaxLoop = rsdb.RowCount : rsdb.MoveFirst()
            prgItemDetails.Minimum = 0 : prgItemDetails.Value = 0 : prgItemDetails.Maximum = intMaxLoop
            prgItemDetails.Visible = True
            ssPOEntry.Visible = False
            ssPOEntry.MaxRows = intMaxLoop
            Call SSMaxLength()
            For intLoopCounter = 1 To intMaxLoop
                If rsdb.GetValue("OpenSO") = False Then
                    ssPOEntry.Row = intLoopCounter
                    ssPOEntry.Col = 1
                    ssPOEntry.Value = 0
                Else
                    ssPOEntry.Row = intLoopCounter
                    ssPOEntry.Col = 1
                    ssPOEntry.Value = 1
                End If
                Call ssPOEntry.SetText(2, intLoopCounter, rsdb.GetValue("Cust_DrgNo"))
                Call ssPOEntry.SetText(4, intLoopCounter, rsdb.GetValue("Item_Code "))
                Call ssPOEntry.SetText(5, intLoopCounter, rsdb.GetValue("Order_Qty"))
                Call ssPOEntry.SetText(13, intLoopCounter, rsdb.GetValue("Rate"))
                Call ssPOEntry.SetText(6, intLoopCounter, rsdb.GetValue("Rate") * CDbl(ctlPerValue.Text))
                Call ssPOEntry.SetText(14, intLoopCounter, rsdb.GetValue("Cust_Mtrl"))
                Call ssPOEntry.SetText(7, intLoopCounter, rsdb.GetValue("Cust_Mtrl") * CDbl(ctlPerValue.Text))
                Call ssPOEntry.SetText(15, intLoopCounter, rsdb.GetValue("Tool_Cost"))
                Call ssPOEntry.SetText(8, intLoopCounter, rsdb.GetValue("Tool_Cost") * CDbl(ctlPerValue.Text))
                Call ssPOEntry.SetText(9, intLoopCounter, rsdb.GetValue("Packing"))
                Call ssPOEntry.SetText(10, intLoopCounter, rsdb.GetValue("Excise_Duty"))
                Call ssPOEntry.SetText(16, intLoopCounter, rsdb.GetValue("Others"))
                Call ssPOEntry.SetText(11, intLoopCounter, rsdb.GetValue("ADD_Excise_Duty"))
                Call ssPOEntry.SetText(19, intLoopCounter, rsdb.GetValue("external_salesorder_no"))
                'GST DETAILS
                Call ssPOEntry.SetText(20, intLoopCounter, rsdb.GetValue("HSNSACCODE"))
                Call ssPOEntry.SetText(21, intLoopCounter, rsdb.GetValue("CGSTTXRT_TYPE"))
                Call ssPOEntry.SetText(22, intLoopCounter, rsdb.GetValue("SGSTTXRT_TYPE"))
                Call ssPOEntry.SetText(23, intLoopCounter, rsdb.GetValue("UTGSTTXRT_TYPE"))
                Call ssPOEntry.SetText(24, intLoopCounter, rsdb.GetValue("IGSTTXRT_TYPE"))
                Call ssPOEntry.SetText(25, intLoopCounter, rsdb.GetValue("COMPENSATION_CESS"))
                Call ssPOEntry.SetText(26, intLoopCounter, rsdb.GetValue("ISHSNORSAC"))
                'GST DETAILS

                rsdb.MoveNext()
                prgItemDetails.Value = prgItemDetails.Value + 1
            Next
            prgItemDetails.Visible = False
            ssPOEntry.Visible = True
            '**************
        Else
            Call ConfirmWindow(10128, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
            ssPOEntry.MaxRows = 0
            rsAD.ResultSetClose()
            rsdb.ResultSetClose()
            Exit Sub
        End If
        With ssPOEntry
            .Enabled = True
            .BlockMode = True
            .Col = 3
            .Col2 = 3
            .Row = 1
            .Row2 = .MaxRows
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
            .TypeButtonPicture = My.Resources.resEmpower.ico111.ToBitmap()
            .BlockMode = False
            .Row = 1
            .Row2 = .MaxRows
            .Col = 1
            .Col2 = .MaxCols
            .BlockMode = True
            .Lock = True
            .BlockMode = False
        End With
        rsAD.ResultSetClose()
        rsdb.ResultSetClose()
        cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = True
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Public Sub ADDRow()
        '----------------------------------------------------------------------------
        'Argument       :   Nil
        'Return Value   :   Nil
        'Function       :   Adds a new row in the grid
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim inti As Short
        If ssPOEntry.MaxRows > 0 Then
            If ValidRowData(ssPOEntry.MaxRows, 0) = True Then
                ssPOEntry.MaxRows = ssPOEntry.MaxRows + 1
            Else
                Exit Sub
            End If
        Else
            ssPOEntry.MaxRows = ssPOEntry.MaxRows + 1
        End If
        With ssPOEntry
            'change account Plug in
            .BlockMode = True
            .Col = 3
            .Col2 = 3
            .Row = 1
            .Row2 = .MaxRows
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
            .TypeButtonPicture = My.Resources.resEmpower.ico111.ToBitmap()
            .BlockMode = False
        End With
        For inti = 5 To 9
            Call ssPOEntry.SetText(inti, ssPOEntry.MaxRows, 0)
        Next
        For inti = 11 To 12
            If inti = 11 Then
                Call ssPOEntry.SetText(inti, ssPOEntry.MaxRows, Nothing)
            Else
                Call ssPOEntry.SetText(inti, ssPOEntry.MaxRows, 0)
            End If
        Next
        Call SSMaxLength()
        With ssPOEntry
            .Col = 1
            If .MaxRows > 1 Then
                .Row = .MaxRows
                .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
            End If
            If chkOpenSo.Checked = True Then
                With ssPOEntry
                    .Row = .MaxRows
                    .Row2 = .MaxRows
                    .Col = 1
                    .Col2 = 1
                    .BlockMode = True
                    ssPOEntry.Value = System.Windows.Forms.CheckState.Checked
                    .BlockMode = False
                End With
            End If
            .Col = 7
            .Col2 = 7
            .Row = 1
            .Row2 = .MaxRows
            .BlockMode = True
            .Lock = False
            .BlockMode = False

            .Col = 29
            .Col2 = 29
            .Row = 1
            .Row2 = .MaxRows
            .BlockMode = True
            .Lock = True
            .BlockMode = False
        End With
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Public Function GetFormDetails() As Boolean
        '----------------------------------------------------------------------------
        'Argument       :   Nil
        'Return Value   :   Nil
        'Function       :   Get the additional details
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim rstAD As ClsResultSetDB
        Dim strSalesTaxType As String
        If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Or cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            m_strSql = "select SalesTax_Type from cust_ord_hdr a,cust_ord_dtl b where a.unit_code=b.unit_code and a.Amendment_No=b.Amendment_No and a.cust_Ref=b.Cust_Ref and a.Account_Code=b.Account_Code and a.Cust_ref='" & txtReferenceNo.Text & "' and a.amendment_No='" & txtAmendmentNo.Text & "' and a.unit_code='" & gstrUNITID & "'"
            rstAD = New ClsResultSetDB
            rstAD.GetResult(m_strSql)
            If rstAD.GetNoRows > 0 Then
                strSalesTaxType = rstAD.GetValue("SalesTax_Type")
            End If
            rstAD.ResultSetClose()
            If IsDBNull(strSalesTaxType) Then
                GetFormDetails = False
            Else
                GetFormDetails = True
            End If
        ElseIf cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            GetFormDetails = False
        End If
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Function
    End Function
    Public Sub GetReferenceDetails()
        '----------------------------------------------------------------------------
        'Argument       :   Nil
        'Return Value   :   Nil
        'Function       :   Get the additional details
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim rsAD As ClsResultSetDB
        Dim rscurrency As ClsResultSetDB
        Dim intLoopCounter As Short
        Dim intMaxCounter As Short
        Dim intDecimal As Short
        Dim strMax As String
        Dim strMin As String


        m_strSql = "select Account_Code,Cust_Ref,Amendment_No,Item_Code,Rate,Order_Qty,Despatch_Qty,Active_Flag,Cust_Mtrl,Cust_DrgNo,Packing,Others,Excise_Duty,Cust_Drg_Desc,Tool_Cost,Authorized_flag,OpenSO,SalesTax_Type,PerValue,InternalSONo,RevisionNo,Remarks,MRP,Abantment_Code,AccessibleRateforMRP,Packing_Type,TOOL_AMOR_FLAG,ShowInAuth,ADD_Excise_Duty ,EXTERNAL_SALESORDER_NO,HSNSACCODE,  ISHSNORSAC, CGSTTXRT_TYPE,SGSTTXRT_TYPE, UTGSTTXRT_TYPE,IGSTTXRT_TYPE,COMPENSATION_CESS  from cust_ord_dtl where unit_code='" & gstrUNITID & "' and  Cust_ref='" & txtReferenceNo.Text & " 'and account_Code='" & txtCustomerCode.Text & "' and amendment_No = '" & txtAmendmentNo.Text & "'"
        rsAD = New ClsResultSetDB
        rsAD.GetResult(m_strSql)
        '10736222-changes done for CT2 ARE 3 functionality - to enable CT2 Reqd flag
        m_strSql = "select Account_Code,Cust_Ref,Amendment_No,Order_Date,Amendment_Date,Active_Flag,Currency_Code,Valid_Date,Effect_Date,Term_Payment,Special_Remarks,Pay_Remarks,Price_Remarks,Packing_Remarks,Frieght_Remarks,Transport_Remarks,Octorai_Remarks,Mode_Despatch,Delivery,First_Authorized,Second_Authorized,Third_Authorized,Authorized_Flag,Reason,PO_Type,SalesTax_Type,OpenSO,AddCustSupp,PerValue,InternalSONo,RevisionNo,Surcharge_code,Future_SO,ECESS_Code,Consignee_Code,ADDVAT_TYPE,CT2_Reqd_In_SO,SERVICETAX_TYPE , SBCTAX_TYPE,KKCTAX_TYPE,exportsotype from cust_ord_hdr where unit_code='" & gstrUNITID & "' and Cust_ref='" & txtReferenceNo.Text & " 'and account_Code='" & txtCustomerCode.Text & "' and Consignee_Code='" & Trim(Me.txtConsCode.Text) & "'  and active_Flag in ('A','L')"
        rsRefNo = New ClsResultSetDB
        rsRefNo.GetResult(m_strSql)
        If rsAD.GetNoRows > 0 Then
            rsAD.MoveFirst()
            txtAmendmentNo.Text = " "
            lblIntSONoDes.Text = rsRefNo.GetValue("InternalSONo")
            lblRevisionNo.Text = rsRefNo.GetValue("RevisionNo")
            If Len(Trim(txtAmendmentNo.Text)) > 0 Then
                DTAmendmentDate.Value = rsRefNo.GetValue("Amendment_Date")
            Else
                DTAmendmentDate.Value = GetServerDate()
            End If
            '10736222-changes done for CT2 ARE 3 functionality - to enable CT2 Reqd flag
            ChkCT2Reqd.Checked = rsRefNo.GetValue("CT2_Reqd_In_SO")
            ChkCT2Reqd.Enabled = False
            txtCreditTerms.Text = rsRefNo.GetValue("Term_Payment")
            DTEffectiveDate.Value = rsRefNo.GetValue("Effect_Date")
            DTValidDate.Value = rsRefNo.GetValue("Valid_Date")
            txtCurrencyType.Text = rsRefNo.GetValue("Currency_Code")
            ctlPerValue.Text = rsRefNo.GetValue("PerValue")
            txtAmendReason.Text = rsRefNo.GetValue("Reason")
            DTDate.Value = rsRefNo.GetValue("Order_date")
            strpotype = rsRefNo.GetValue("PO_Type")
            strSOType = rsRefNo.GetValue("salestax_Type")
            strexportType = rsRefNo.GetValue("exportsotype")
            If strexportType = "WP" Then
                cmbExporttype.SelectedIndex = 1
            ElseIf strexportType = "WOP" Then
                cmbExporttype.SelectedIndex = 2
            Else
                cmbExporttype.SelectedIndex = 0
            End If
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
                Case "V"
                    Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 5)
            End Select
            txtSTax.Text = strSOType
            txtSChSTax.Text = rsRefNo.GetValue("Surcharge_code")
            txtECSSTaxType.Text = rsRefNo.GetValue("ECESS_Code")
            '10869290
            txtService.Text = IIf(IsDBNull(rsRefNo.GetValue("SERVICETAX_TYPE")), "", rsRefNo.GetValue("SERVICETAX_TYPE"))

            If txtService.Text.Trim.Length > 1 Then
                Call txtService_Validating(txtService, New System.ComponentModel.CancelEventArgs(False))
            End If
            txtSBC.Text = IIf(IsDBNull(rsRefNo.GetValue("SBCTAX_TYPE")), "", rsRefNo.GetValue("SBCTAX_TYPE"))
            txtKKC.Text = IIf(IsDBNull(rsRefNo.GetValue("KKCTAX_TYPE")), "", rsRefNo.GetValue("KKCTAX_TYPE"))

            If rsRefNo.GetValue("OpenSO") = False Then
                chkOpenSo.CheckState = System.Windows.Forms.CheckState.Unchecked
            Else
                chkOpenSo.CheckState = System.Windows.Forms.CheckState.Checked
            End If
            rsAD.MoveFirst()
            ssPOEntry.MaxRows = 0
            If cmdButtons.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                If Len(Trim(txtCurrencyType.Text)) Then
                    rscurrency = New ClsResultSetDB
                    rscurrency.GetResult("Select decimal_Place from Currency_Mst where unit_code='" & gstrUNITID & "' and currency_code ='" & Trim(txtCurrencyType.Text) & "'")
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
                '****************************
                '****************************
                intMaxCounter = rsAD.GetNoRows
                prgItemDetails.Value = 0 : prgItemDetails.Minimum = 0 : prgItemDetails.Maximum = intMaxCounter
                prgItemDetails.Visible = True
                rsAD.MoveFirst()
                With ssPOEntry
                    For intLoopCounter = 1 To intMaxCounter
                        ssPOEntry.MaxRows = ssPOEntry.MaxRows + 1
                        m_custItemDesc = rsAD.GetValue("Cust_Drg_Desc")
                        If rsAD.GetValue("OpenSO") = False Then
                            .Col = 1
                            .Row = intLoopCounter
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
                            ssPOEntry.Value = 0
                        Else
                            .Col = 1
                            .Row = intLoopCounter
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
                            ssPOEntry.Value = 1
                        End If
                        .Col = 2
                        .Row = intLoopCounter
                        .TypeMaxEditLen = 30
                        Call ssPOEntry.SetText(2, ssPOEntry.MaxRows, rsAD.GetValue("Cust_DrgNo"))
                        .Col = 4
                        .Row = intLoopCounter
                        .TypeMaxEditLen = 16
                        Call ssPOEntry.SetText(4, ssPOEntry.MaxRows, rsAD.GetValue("Item_Code"))
                        .Col = 5
                        .Row = intLoopCounter
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                        .TypeFloatDecimalPlaces = 2
                        .TypeFloatMin = "0.00"
                        .TypeFloatMax = "9999999.99"
                        Call ssPOEntry.SetText(5, ssPOEntry.MaxRows, rsAD.GetValue("Order_Qty"))
                        .Col = 14
                        .Row = intLoopCounter
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                        .TypeFloatDecimalPlaces = intDecimal
                        .TypeFloatMin = strMin
                        .TypeFloatMax = strMax
                        Call ssPOEntry.SetText(13, ssPOEntry.MaxRows, rsAD.GetValue("Rate"))
                        .Col = 6
                        .Row = intLoopCounter
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                        .TypeFloatDecimalPlaces = intDecimal
                        .TypeFloatMin = strMin
                        .TypeFloatMax = strMax
                        Call ssPOEntry.SetText(6, ssPOEntry.MaxRows, rsAD.GetValue("Rate") * CDbl(ctlPerValue.Text))
                        .Col = 15
                        .Row = intLoopCounter
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                        .TypeFloatDecimalPlaces = intDecimal
                        .TypeFloatMin = strMin
                        .TypeFloatMax = strMax
                        Call ssPOEntry.SetText(14, ssPOEntry.MaxRows, rsAD.GetValue("Cust_Mtrl"))
                        .Col = 7
                        .Row = intLoopCounter
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                        .TypeFloatDecimalPlaces = intDecimal
                        .TypeFloatMin = strMin
                        .TypeFloatMax = strMax
                        Call ssPOEntry.SetText(7, ssPOEntry.MaxRows, rsAD.GetValue("Cust_Mtrl") * CDbl(ctlPerValue.Text))
                        .Col = 16
                        .Row = intLoopCounter
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                        .TypeFloatDecimalPlaces = 4
                        .TypeFloatMin = strMin
                        .TypeFloatMax = strMax
                        Call ssPOEntry.SetText(15, ssPOEntry.MaxRows, rsAD.GetValue("Tool_Cost"))
                        .Col = 8
                        .Row = intLoopCounter
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                        .TypeFloatDecimalPlaces = 4
                        .TypeFloatMin = strMin
                        .TypeFloatMax = "99999999.9999"
                        Call ssPOEntry.SetText(8, ssPOEntry.MaxRows, rsAD.GetValue("Tool_Cost") * CDbl(ctlPerValue.Text))
                        .Col = 9
                        .Row = intLoopCounter
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                        .TypeFloatDecimalPlaces = 2
                        .TypeFloatMin = "0.00"
                        '.TypeFloatMax = "100.00"
                        .TypeFloatMax = "99999999.9999"
                        Call ssPOEntry.SetText(9, ssPOEntry.MaxRows, rsAD.GetValue("Packing"))
                        .Col = 10
                        .Row = intLoopCounter
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                        Call ssPOEntry.SetText(10, ssPOEntry.MaxRows, rsAD.GetValue("Excise_Duty"))
                        .Col = 11
                        .Row = intLoopCounter
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                        Call ssPOEntry.SetText(11, ssPOEntry.MaxRows, rsAD.GetValue("ADD_Excise_Duty") & "")
                        .Col = 12
                        .Row = intLoopCounter
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                        .TypeFloatDecimalPlaces = 2
                        .TypeFloatMin = "0.00"
                        .TypeFloatMax = "99999999.99"
                        Call ssPOEntry.SetText(17, ssPOEntry.MaxRows, rsAD.GetValue("Others") * CDbl(ctlPerValue.Text))
                        .Col = 12
                        .Row = intLoopCounter
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                        .TypeFloatDecimalPlaces = 2
                        .TypeFloatMin = "0.00"
                        .TypeFloatMax = "99999999.99"
                        Call ssPOEntry.SetText(12, ssPOEntry.MaxRows, rsAD.GetValue("Others"))
                        Call ssPOEntry.SetText(19, ssPOEntry.MaxRows, rsAD.GetValue("EXTERNAL_SALESORDER_NO"))

                        'GST CHANGE
                        .Col = 20
                        .Row = intLoopCounter
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                        Call ssPOEntry.SetText(20, ssPOEntry.MaxRows, rsAD.GetValue("HSNSACCODE"))
                        .Col = 21
                        .Row = intLoopCounter
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                        Call ssPOEntry.SetText(21, ssPOEntry.MaxRows, rsAD.GetValue("CGSTTXRT_TYPE"))
                        .Col = 22
                        .Row = intLoopCounter
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                        Call ssPOEntry.SetText(22, ssPOEntry.MaxRows, rsAD.GetValue("SGSTTXRT_TYPE"))
                        .Col = 23
                        .Row = intLoopCounter
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                        Call ssPOEntry.SetText(23, ssPOEntry.MaxRows, rsAD.GetValue("UTGSTTXRT_TYPE"))
                        .Col = 24
                        .Row = intLoopCounter
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                        Call ssPOEntry.SetText(24, ssPOEntry.MaxRows, rsAD.GetValue("IGSTTXRT_TYPE"))
                        .Col = 25
                        .Row = intLoopCounter
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                        Call ssPOEntry.SetText(25, ssPOEntry.MaxRows, rsAD.GetValue("COMPENSATION_CESS"))
                        .Col = 26
                        .Row = intLoopCounter
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                        Call ssPOEntry.SetText(26, ssPOEntry.MaxRows, rsAD.GetValue("ISHSNORSAC"))

                        'GST CHANGE
                        '****
                        rsAD.MoveNext()
                        prgItemDetails.Value = prgItemDetails.Value + 1
                    Next
                    prgItemDetails.Visible = False
                End With
            End If
        Else

            Call ConfirmWindow(10129, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
            rsRefNo.ResultSetClose()
            rsAD.ResultSetClose()
            Exit Sub
        End If
        With ssPOEntry
            .Enabled = True
            .BlockMode = True
            .Col = 3
            .Col2 = 3
            .Row = 1
            .Row2 = .MaxRows
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
            .TypeButtonPicture = My.Resources.resEmpower.ico111.ToBitmap()
            .BlockMode = False
            If cmdButtons.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                .Row = 1
                .Row2 = .MaxRows
                .Col = 1
                .Col2 = .MaxCols
                .BlockMode = True
                .Lock = True
                .BlockMode = False
            End If
        End With
        '****************Anan
        cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = True
        rsRefNo.ResultSetClose()
        rsAD.ResultSetClose()
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Public Sub DeleteDocumentPO()
        If txtReferenceNo.Text.Trim = "" Then
            Exit Sub
        End If
        Dim strSQL As String = String.Empty

        Try
            If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                lblFilePath.Tag = ""
                lblFilePath.Text = ""
                If txtAmendmentNo.Text.Trim <> "" Then
                    strSQL = "Delete FROM Cust_Ord_DOC_Dtl WHERE UNIT_CODE ='" & gstrUNITID & "' AND Account_Code ='" & txtCustomerCode.Text.Trim & "' AND Cust_Ref='" & txtReferenceNo.Text.Trim & "' AND ISACTIVE=1 AND Amendment_No='" & txtAmendmentNo.Text.Trim & "'"

                Else
                    strSQL = "Delete FROM Cust_Ord_DOC_Dtl WHERE UNIT_CODE ='" & gstrUNITID & "' AND Account_Code ='" & txtCustomerCode.Text.Trim & "' AND Cust_Ref='" & txtReferenceNo.Text.Trim & "' AND ISACTIVE=1"

                End If
                mP_Connection.Execute(strSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub
    Public Sub DeleteRow()
        '----------------------------------------------------------------------------
        'Argument       :   Nil
        'Return Value   :   Nil
        'Function       :   Deletes a row
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim intRow As Short
        Dim rsDespatchQuantity As ClsResultSetDB
        Dim varItemCode As Object
        Dim varDrawCode As Object
        ReDim ArrDispatchQty(ssPOEntry.MaxRows - 1)
        For intRow = 1 To ssPOEntry.MaxRows
            varItemCode = Nothing
            Call ssPOEntry.GetText(4, intRow, varItemCode)
            varDrawCode = Nothing
            Call ssPOEntry.GetText(2, intRow, varDrawCode)
            rsDespatchQuantity = New ClsResultSetDB
            rsDespatchQuantity.GetResult("Select Despatch_Qty from cust_ord_dtl where unit_code='" & gstrUNITID & "' and Account_Code='" & txtCustomerCode.Text & "'and Cust_Ref='" & txtReferenceNo.Text & "' and Amendment_No= '" & txtAmendmentNo.Text & "' and Item_Code= '" & varItemCode & "' and Cust_drgNo ='" & varDrawCode & "' and Authorized_flag = 0")
            If rsDespatchQuantity.GetNoRows > 0 Then
                ArrDispatchQty(intRow - 1) = rsDespatchQuantity.GetValue("Despatch_Qty")
            Else
                ArrDispatchQty(intRow - 1) = 0
            End If
            rsDespatchQuantity.ResultSetClose()
            strSQL = "delete cust_ord_dtl where unit_code='" & gstrUNITID & "' and Account_Code='"
            strSQL = strSQL & txtCustomerCode.Text & "'and Cust_Ref='"
            strSQL = strSQL & txtReferenceNo.Text & "' and Amendment_No= '"
            strSQL = strSQL & txtAmendmentNo.Text & "' and Item_Code= '"
            strSQL = strSQL & varItemCode & "' and Cust_drgNo ='" & varDrawCode & "' and Authorized_flag = 0"
            mP_Connection.Execute(strSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        Next
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
                MsgBox("Document uploaded successfully!,Auto Mail Receiver for Sales Order entry Not defined.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100))

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
                MsgBox("Document uploaded successfully!,Auto Mail Receiver for Sales Order Entry Not defined.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100))
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
            MainBody = MainBody + "A new Sale order for " + lblCustDesc.Text.Trim() + " and related document’s has been uploaded & sent for your approval as per below details:"
            MainBody = MainBody + vbNewLine + vbNewLine
            MainBody = MainBody + "REQ. BY: " + GetUserNameByEmployeeCodeAndUnit(mP_User, gstrUNITID) + "" + vbTab + vbTab + "DEPT: " + gstrUNITID.ToString() + ""
            MainBody = MainBody + vbNewLine + "Customer Code: " + txtCustomerCode.Text.Trim() + "" + vbTab + vbTab + vbTab + "Customer Consignee Code: " + txtConsCode.Text.Trim + "" + vbNewLine
            MainBody = MainBody + "Sale Order Date: " + DateTime.Now.Date.ToString("dd/MM/yyyy") + "" + vbTab + vbTab + vbTab + "Internal SO #: " + lblIntSONoDes.Text.Trim + " " + vbNewLine
            MainBody = MainBody + "Sale Order No: " + txtReferenceNo.Text.Trim() + "" + vbNewLine + vbNewLine + vbNewLine

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
    Public Sub UPLoadPODocument()
        '----------------------------------------------------------------------------
        'Argument       :   Nil
        'Return Value   :   Nil
        'Function       :   Inserts a row
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler

        Dim strSQL As String




        ''Upload the PO if required

        If Not IsNothing(lblFilePath.Tag) Or lblFilePath.Text.Trim.Length > 0 Then
            Dim oCmd As ADODB.Command
            oCmd = New ADODB.Command
            Dim FlagModified As Integer
            If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                FlagModified = 1
            Else
                FlagModified = 0
            End If
            With oCmd
                .let_ActiveConnection(mP_Connection)
                .CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                .CommandText = "USP__Customer_Order_DocDetails"
                .Parameters.Append(.CreateParameter("@UnitCode", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                .Parameters.Append(.CreateParameter("@CustomerCode", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 8, Trim(txtCustomerCode.Text)))
                .Parameters.Append(.CreateParameter("@ReferenceNo", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 25, Trim(txtReferenceNo.Text)))
                .Parameters.Append(.CreateParameter("@AmendmentNo", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 25, Trim(txtAmendmentNo.Text)))
                .Parameters.Append(.CreateParameter("@DOC_NAME", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 250, Trim(Convert.ToString(lblFilePath.Tag).Trim)))
                .Parameters.Append(.CreateParameter("@ISACTIVE", ADODB.DataTypeEnum.adBoolean, ADODB.ParameterDirectionEnum.adParamInput, , 1))
                .Parameters.Append(.CreateParameter("@MUser", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, mP_User))
                .Parameters.Append(.CreateParameter("@FlagModified", ADODB.DataTypeEnum.adTinyInt, ADODB.ParameterDirectionEnum.adParamInput, , FlagModified))
                .Parameters.Append(.CreateParameter("@MSMEDOC", ADODB.DataTypeEnum.adLongVarBinary, ADODB.ParameterDirectionEnum.adParamInput, 2147483647, GetFileBytes(lblFilePath.Text)))
                .Parameters.Append(.CreateParameter("@ERRCODE", ADODB.DataTypeEnum.adInteger, ADODB.ParameterDirectionEnum.adParamOutput))
                .Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End With
            If oCmd.Parameters(oCmd.Parameters.Count - 1).Value <> 0 Then
                MsgBox("Error while inserting Document Sale Order details", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100))
                oCmd = Nothing
                Exit Sub
            End If

            oCmd = Nothing




        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Public Sub InsertRow()
        '----------------------------------------------------------------------------
        'Argument       :   Nil
        'Return Value   :   Nil
        'Function       :   Inserts a row
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim varDeleteFlag As Object
        Dim clsDrgDes As ClsResultSetDB
        Dim strSQL As String
        Dim varExtraExciseDuty, varSalestax, varToolCost, varordqty, varItemCode, varRate, varCustSuppMaterial, varPkg, varExiseDuty, varSurchargeSalesTax As Object
        Dim varOpenSO, varDespatchQty, varCustItemCode, varCustItemDesc, varOthers, valAddExciseDuty, varexternalsalesorder As Object
        'GST CHANGES
        Dim VARHSNSACCODE As Object
        Dim VARISHSNORSAC As Object
        Dim VARCGSTTXRT_HEAD As Object
        Dim VARSGSTTXRT_HEAD As Object
        Dim VARIGSTTXRT_HEAD As Object
        Dim VARUGSTTXRT_HEAD As Object
        Dim VARCOMPENSATIONCESS_HEAD As Object

        'GST CHANGES

        For intRow = 1 To ssPOEntry.MaxRows
            varDeleteFlag = Nothing
            Call ssPOEntry.GetText(0, intRow, varDeleteFlag)
            If varDeleteFlag <> "*" Then 'to get the values from the grid
                'change account Plug in
                varCustItemCode = Nothing
                Call ssPOEntry.GetText(2, intRow, varCustItemCode)
                'issue id 10117810
                varItemCode = Nothing
                Call ssPOEntry.GetText(4, intRow, varItemCode)
                'Getting the Drawing No. Description
                'strSQL = "SELECT drg_desc FROM  custitem_mst WHERE unit_code='" & gstrUNITID & "' and Cust_drgno = '" & Trim(varCustItemCode) & "'"
                strSQL = "SELECT drg_desc FROM  custitem_mst WHERE Unit_code = '" & gstrUNITID & "' and account_code='" & txtCustomerCode.Text.Trim & "' and Cust_drgno = '" & Trim(varCustItemCode) & "' and item_code= '" & Trim(varItemCode) & "' and active=1 "

                clsDrgDes = New ClsResultSetDB
                If clsDrgDes.GetResult(strSQL, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly) And clsDrgDes.GetNoRows > 0 Then
                    clsDrgDes.MoveFirst()
                    m_custItemDesc = Trim(clsDrgDes.GetValue("drg_desc"))
                    clsDrgDes.ResultSetClose()
                End If
                'change account Plug in
                varOpenSO = Nothing
                Call ssPOEntry.GetText(1, intRow, varOpenSO)
                'issue id 10117810 
                'varItemCode = Nothing
                'Call ssPOEntry.GetText(4, intRow, varItemCode)
                'issue id 10117810 end 
                varordqty = Nothing
                Call ssPOEntry.GetText(5, intRow, varordqty)
                varRate = Nothing
                Call ssPOEntry.GetText(6, intRow, varRate)
                If Len(ctlPerValue.Text.Trim) >= 1 Then
                    varRate = varRate / CDbl(ctlPerValue.Text)
                End If
                '****
                varCustSuppMaterial = Nothing
                Call ssPOEntry.GetText(7, intRow, varCustSuppMaterial)
                varToolCost = Nothing
                Call ssPOEntry.GetText(8, intRow, varToolCost)
                varPkg = Nothing
                Call ssPOEntry.GetText(9, intRow, varPkg)
                varExiseDuty = Nothing
                Call ssPOEntry.GetText(10, intRow, varExiseDuty)
                valAddExciseDuty = Nothing
                Call ssPOEntry.GetText(11, intRow, valAddExciseDuty)
                varOthers = Nothing
                Call ssPOEntry.GetText(12, intRow, varOthers)
                varexternalsalesorder = Nothing
                Call ssPOEntry.GetText(19, intRow, varexternalsalesorder)

                'GST CHANGES
                If gblnGSTUnit = True Then
                    VARHSNSACCODE = Nothing
                    Call ssPOEntry.GetText(20, intRow, VARHSNSACCODE)
                    VARCGSTTXRT_HEAD = Nothing
                    Call ssPOEntry.GetText(21, intRow, VARCGSTTXRT_HEAD)
                    VARSGSTTXRT_HEAD = Nothing
                    Call ssPOEntry.GetText(22, intRow, VARSGSTTXRT_HEAD)
                    VARUGSTTXRT_HEAD = Nothing
                    Call ssPOEntry.GetText(23, intRow, VARUGSTTXRT_HEAD)
                    VARIGSTTXRT_HEAD = Nothing
                    Call ssPOEntry.GetText(24, intRow, VARIGSTTXRT_HEAD)
                    VARCOMPENSATIONCESS_HEAD = Nothing
                    Call ssPOEntry.GetText(25, intRow, VARCOMPENSATIONCESS_HEAD)
                    VARISHSNORSAC = Nothing
                    Call ssPOEntry.GetText(26, intRow, VARISHSNORSAC)
                End If
                'GST CHANGES

                strSQL = "Insert into Cust_Ord_Dtl (unit_code,Account_Code, Cust_Ref, Amendment_No,InternalSONo,RevisionNo, "
                strSQL = strSQL & "Item_Code , Rate, Order_Qty, Despatch_Qty, "
                strSQL = strSQL & "Active_Flag, Cust_Mtrl, Cust_DrgNo, Packing, Others,"
                strSQL = strSQL & "Excise_Duty,"
                strSQL = strSQL & "Cust_Drg_Desc,"
                strSQL = strSQL & "Tool_Cost, Authorized_flag, openSO,Ent_dt, Ent_UserId, Upd_dt, Upd_UserId,PerValue,ShowInAuth,ADD_Excise_Duty,external_salesorder_no,ISHSNORSAC,HSNSACCODE,CGSTTXRT_TYPE,SGSTTXRT_TYPE,UTGSTTXRT_TYPE,IGSTTXRT_TYPE,COMPENSATION_CESS)"
                strSQL = strSQL & " values('" & gstrUNITID & "','" & Trim(txtCustomerCode.Text) & "','"
                strSQL = strSQL & Trim(txtReferenceNo.Text) & "','"
                strSQL = strSQL & Trim(txtAmendmentNo.Text) & "','"
                strSQL = strSQL & Trim(lblIntSONoDes.Text) & "'," & Trim(lblRevisionNo.Text) & ",'"
                strSQL = strSQL & Trim(varItemCode) & "',"
                strSQL = strSQL & Trim(varRate) & ","
                strSQL = strSQL & IIf(IsNothing(varordqty), 0, varordqty) & "," & ArrDispatchQty(intRow - 1) & ",'A',"
                strSQL = strSQL & IIf(IsNothing(varCustSuppMaterial), 0, varCustSuppMaterial) & ",'"
                strSQL = strSQL & varCustItemCode & "'," & IIf(IsNothing(varPkg), 0, varPkg) & ","
                strSQL = strSQL & IIf(IsNothing(varOthers), 0, Val(varOthers)) & ",'"
                strSQL = strSQL & IIf(IsNothing(varExiseDuty), 0, varExiseDuty) & "','"
                strSQL = strSQL & Trim(m_custItemDesc) & "',"
                strSQL = strSQL & IIf(IsNothing(varToolCost), 0, varToolCost) & ",0,"
                If chkOpenSo.CheckState = System.Windows.Forms.CheckState.Checked Then
                    strSQL = strSQL & "1,"
                Else
                    If CBool(Val(varOpenSO)) = False Then
                        strSQL = strSQL & "0,"
                    Else
                        strSQL = strSQL & "1,"
                    End If
                End If
                strSQL = strSQL & " getdate()" & ",'" & mP_User & "'," & "getdate() "
                strSQL = strSQL & ",'" & mP_User & "'"
                If Len(ctlPerValue.Text.Trim) >= 1 Then
                    strSQL = strSQL & "," & ctlPerValue.Text & ",1,'" & valAddExciseDuty & "','" & varexternalsalesorder & "'"
                Else
                    strSQL = strSQL & ", 1 ,1,'" & valAddExciseDuty & "','" & varexternalsalesorder & "'"
                End If
                'GST CHANGES 
                strSQL = strSQL & ",'" & VARISHSNORSAC & "','" & VARHSNSACCODE & "','" & VARCGSTTXRT_HEAD & "','" & VARSGSTTXRT_HEAD & "','" & VARUGSTTXRT_HEAD & "','" & VARIGSTTXRT_HEAD & "','" & VARCOMPENSATIONCESS_HEAD & "' )"
                mP_Connection.Execute(strSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End If
        Next


        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
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
    Public Sub UpdateRow()
        '----------------------------------------------------------------------------
        'Argument       :   Nil
        'Return Value   :   Nil
        'Function       :   Updates the header table
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim varExiseDuty As Object
        varExiseDuty = Nothing
        Call ssPOEntry.GetText(10, 1, varExiseDuty)
        strSQL = "update cust_ord_hdr set Order_Date='"
        strSQL = strSQL & getDateForDB(DTDate.Value) & "',"
        If Len(Trim(txtAmendmentNo.Text)) > 0 Then
            strSQL = strSQL & "Amendment_Date='"
            strSQL = strSQL & getDateForDB(DTAmendmentDate.Value) & "',"
        End If
        strSQL = strSQL & "Currency_Code='"
        strSQL = strSQL & txtCurrencyType.Text & "',Valid_Date='"
        strSQL = strSQL & getDateForDB(DTValidDate.Value) & "',Effect_Date='"
        strSQL = strSQL & getDateForDB(DTEffectiveDate.Value) & "',Term_Payment='"
        strSQL = strSQL & Trim(txtCreditTerms.Text) & "',Special_Remarks='" & m_strSpecialNotes & "',Pay_Remarks='"
        strSQL = strSQL & m_strPaymentTerms & "',Price_Remarks='" & m_strPricesAre & "',Packing_Remarks='"
        strSQL = strSQL & m_strPkgAndFwd & "',Frieght_Remarks='" & m_strFreight & "',Transport_Remarks='"
        strSQL = strSQL & m_strTransitInsurance & "',Octorai_Remarks='" & m_strOctroi & "',Mode_Despatch='"
        strSQL = strSQL & m_strModeOfDespatch & "',Delivery='" & m_strDeliverySchedule & "',"
        strSQL = strSQL & "Reason='" & txtAmendReason.Text & "',PO_Type='"
        strSQL = strSQL & Mid(cmbPOType.Text, 1, 1) & "',"
        If chkOpenSo.CheckState = System.Windows.Forms.CheckState.Checked Then
            strSQL = strSQL & " OpenSO = 1,"
        Else
            strSQL = strSQL & " OpenSO = 0,"
        End If

        '10736222-changes done for CT2 ARE 3 functionality - to enable CT2 Reqd flag
        If ChkCT2Reqd.CheckState = System.Windows.Forms.CheckState.Checked Then
            strSQL = strSQL & " CT2_Reqd_In_SO = 1,"
        Else
            strSQL = strSQL & " CT2_Reqd_In_SO = 0,"
        End If

        strSQL = strSQL & " SalesTax_type = '" & Trim(txtSTax.Text) & "',"
        '****
        strSQL = strSQL & " Ent_dt="
        strSQL = strSQL & " getdate() " & ",Ent_UserId='" & mP_User & "',Upd_dt="
        strSQL = strSQL & " getdate() " & ",Upd_UserId='" & mP_User & "'"
        strSQL = strSQL & " , Surcharge_Code = '" & Trim(txtSChSTax.Text) & "'"
        Dim strConsCode As String
        If Len(Trim(txtConsCode.Text)) = 0 Then
            strConsCode = Trim(Me.txtCustomerCode.Text)
        Else
            strConsCode = Trim(txtConsCode.Text)
        End If
        strSQL = strSQL & " , ECESS_Code = '" & Trim(txtECSSTaxType.Text) & "',Consignee_Code ='" & strConsCode & "' "
        strSQL = strSQL & " , SERVICETAX_TYPE = '" & Trim(txtService.Text) & "' "
        strSQL = strSQL & " , SBCTAX_TYPE = '" & Trim(txtSBC.Text) & "'"
        strSQL = strSQL & " , KKCTAX_TYPE = '" & Trim(txtKKC.Text) & "'"
        If CDbl(ctlPerValue.Text) >= 1 Then
            strSQL = strSQL & ", PerValue = " & ctlPerValue.Text & " where Account_Code='"
        Else
            strSQL = strSQL & ", PerValue = 1 where Account_Code='"
        End If
        strSQL = strSQL & txtCustomerCode.Text & "'and Cust_Ref='"
        strSQL = strSQL & txtReferenceNo.Text & "'and Amendment_No='" & txtAmendmentNo.Text & "' and unit_code='" & gstrUNITID & "'"
        mP_Connection.Execute("Set dateformat 'mdy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        mP_Connection.Execute(strSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Public Sub InsertRowCustOrdHdr()
        '----------------------------------------------------------------------------
        'Argument       :   Nil
        'Return Value   :   Nil
        'Function       :   Inserts Row in the header table
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim varExiseDuty As Object
        Dim varSalestax As Object
        Dim varSurchargeSalesTax As Object

        Dim ship_address_code As String = ""
        Dim ship_address_desc As String = ""
        Dim strexportsotext As String

        'change account Plug in
        varExiseDuty = Nothing
        Call ssPOEntry.GetText(10, 1, varExiseDuty)

        If chkShipAddress.Checked = True Then
            If Len(txtShipAddress.Text) > 0 Then
                ship_address_code = Trim(txtShipAddress.Text)
                ship_address_desc = Trim(lblShipAddress_Details.Text)
            Else
                MsgBox("Please Select Ship Address", MsgBoxStyle.Information, "EMPRO")
                Exit Sub
            End If
        Else
            ship_address_code = ""
            ship_address_desc = ""
        End If


        If Len(Trim(txtAmendmentNo.Text)) = 0 Then
            lblIntSONoDes.Text = GenerateDocumentNumber("Cust_ord_hdr", "InternalSONo", "ent_dt", CStr(GetServerDate()))
        End If
        strSQL = "Insert into Cust_Ord_Hdr (Unit_code,Account_Code, Cust_Ref, Amendment_No,InternalSONo,RevisionNo, Order_Date, "
        strSQL = strSQL & "Amendment_Date, Active_Flag, "
        strSQL = strSQL & " Currency_Code, Valid_Date,"
        strSQL = strSQL & "Effect_Date, Term_Payment, Special_Remarks, Pay_Remarks, "
        strSQL = strSQL & "Price_Remarks, Packing_Remarks, Frieght_Remarks, Transport_Remarks,"
        strSQL = strSQL & "Octorai_Remarks, Mode_Despatch, Delivery, First_Authorized,"
        strSQL = strSQL & "Second_Authorized, Third_Authorized, Authorized_Flag, Reason, "
        '10736222-changes done for CT2 ARE 3 functionality - to enable CT2 Reqd flag
        'If cmbPOType.Text.ToUpper = "EXPORT" Then
        strSQL = strSQL & "PO_Type, SalesTax_Type,OpenSO,CT2_Reqd_In_SO,Ent_dt, Ent_UserId, Upd_dt, Upd_UserId,PerValue, Surcharge_Code,ECESS_Code,Consignee_Code,SERVICETAX_TYPE,SBCTAX_TYPE,KKCTAX_TYPE,ShipAddress_Code,ShipAddress_Desc,ExportSotype)"
        'Else
        ''strSQL = strSQL & "PO_Type, SalesTax_Type,OpenSO,CT2_Reqd_In_SO,Ent_dt, Ent_UserId, Upd_dt, Upd_UserId,PerValue, Surcharge_Code,ECESS_Code,Consignee_Code,SERVICETAX_TYPE,SBCTAX_TYPE,KKCTAX_TYPE,ShipAddress_Code,ShipAddress_Desc)"
        'End If

        strSQL = strSQL & " Values('" & gstrUNITID & "','" & Trim(txtCustomerCode.Text) & "','" & Trim(txtReferenceNo.Text) & "','" & Trim(txtAmendmentNo.Text) & "',"
        strSQL = strSQL & "'" & lblIntSONoDes.Text & "',"
        If Len(Trim(txtAmendmentNo.Text)) > 0 Then
            lblRevisionNo.Text = CStr(GenerateRevisionNo())
            strSQL = strSQL & lblRevisionNo.Text & ",'"
        Else
            lblRevisionNo.Text = "0"
            strSQL = strSQL & "0,'"
        End If
        strSQL = strSQL & getDateForDB(DTDate.Value) & "','" & IIf(Len(Me.txtAmendmentNo.Text) = 0, System.DBNull.Value, getDateForDB(DTAmendmentDate.Value)) & "','A' ,'"
        strSQL = strSQL & Trim(txtCurrencyType.Text) & "','" & getDateForDB(DTValidDate.Value) & "','" & getDateForDB(DTEffectiveDate.Value) & "','"
        strSQL = strSQL & IIf(Len(Trim(txtCreditTerms.Text)) = 0, 0, Trim(txtCreditTerms.Text)) & "','"
        strSQL = strSQL & Trim(m_strSpecialNotes) & "','"
        strSQL = strSQL & Trim(m_strPaymentTerms) & "','" & Trim(m_strPricesAre) & "','" & Trim(m_strPkgAndFwd) & "','" & Trim(m_strFreight) & "','"
        strSQL = strSQL & Trim(m_strTransitInsurance) & "','" & Trim(m_strOctroi) & "','" & Trim(m_strModeOfDespatch) & "','" & Trim(m_strDeliverySchedule) & "','',"
        strSQL = strSQL & "'','','','" & Trim(txtAmendReason.Text) & "','" & Mid(cmbPOType.Text, 1, 1) & "','" & Trim(txtSTax.Text) & "',"
        If chkOpenSo.CheckState = System.Windows.Forms.CheckState.Checked Then
            strSQL = strSQL & "1,"
        Else
            strSQL = strSQL & "0,"
        End If

        '10736222-changes done for CT2 ARE 3 functionality - to enable CT2 Reqd flag
        If ChkCT2Reqd.CheckState = System.Windows.Forms.CheckState.Checked Then
            strSQL = strSQL & "1,"
        Else
            strSQL = strSQL & "0,"
        End If

        strSQL = strSQL & " getdate() " & ",'" & mP_User & "'," & " getdate() " & ",'" & mP_User & "'"
        If CDbl(Val(ctlPerValue.Text)) >= 1 Then
            strSQL = strSQL & "," & ctlPerValue.Text & ",'" & Trim(txtSChSTax.Text) & "'"
        Else
            strSQL = strSQL & ", 1,'" & Trim(txtSChSTax.Text) & "'"
        End If

        Dim strConsCode As String
        If Len(Trim(txtConsCode.Text)) = 0 Then
            strConsCode = Trim(Me.txtCustomerCode.Text)
        Else
            strConsCode = Trim(txtConsCode.Text)
        End If


        strSQL = strSQL & ",'" & Trim(txtECSSTaxType.Text) & "', '" & strConsCode & "','"
        '10869290
        'strSql = strSql & Trim(txtService.Text) & "')"
        If UCase(cmbExporttype.Text) = "WITHOUT PAY" Then
            strexportsotext = "WOP"
        ElseIf UCase(cmbExporttype.Text) = "WITH PAY" Then
            strexportsotext = "WP"
        Else
            strexportsotext = ""
        End If

        'If cmbPOType.Text.ToUpper = "EXPORT" Then
        strSQL = strSQL & Trim(txtService.Text) & "','" & Trim(txtSBC.Text) & "','" & Trim(txtKKC.Text) & "','" & ship_address_code & "','" & ship_address_desc & "','" & strexportsotext & "')"
        ''Else
        ''strSQL = strSQL & Trim(txtService.Text) & "','" & Trim(txtSBC.Text) & "','" & Trim(txtKKC.Text) & "','" & ship_address_code & "','" & ship_address_desc & "')"
        ''End If

        mP_Connection.Execute("Set dateformat 'mdy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        mP_Connection.Execute(strSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Public Function CheckFormDetails() As Boolean
        '----------------------------------------------------------------------------
        'Argument       :   Nil
        'Return Value   :   Nil
        'Function       :   Check The Additional Details that they have been entered or not
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim rstAD As ClsResultSetDB
        If Len(Trim(m_strPaymentTerms)) = 0 Then
            If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Or cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                m_strSql = "select Special_Remarks,Pay_Remarks,Price_Remarks,Packing_Remarks,Frieght_Remarks,Transport_Remarks,Octorai_Remarks, Mode_Despatch,Delivery from cust_ord_hdr a,cust_ord_dtl b where a.unit_code=b.unit_code and a.Amendment_No=b.Amendment_No and a.cust_Ref=b.Cust_Ref and a.Account_Code=b.Account_Code and a.Cust_ref='" & txtReferenceNo.Text & "' and a.amendment_No='" & txtAmendmentNo.Text & "' and a.unit_code='" & gstrUNITID & "'"
            ElseIf cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                m_strSql = "select Special_Remarks,Pay_Remarks,Price_Remarks,Packing_Remarks,Frieght_Remarks,Transport_Remarks,Octorai_Remarks, Mode_Despatch,Delivery from cust_ord_hdr  where Cust_ref='" & txtReferenceNo.Text & "' and account_code='" & txtCustomerCode.Text & "' and Consignee_Code='" & Trim(Me.txtConsCode.Text) & "' and unit_code='" & gstrUNITID & "'"
            End If
            rstAD = New ClsResultSetDB
            rstAD.GetResult(m_strSql)
            If rstAD.GetNoRows > 0 Then
                m_strSpecialNotes = rstAD.GetValue("Special_Remarks")
                m_strPaymentTerms = rstAD.GetValue("Pay_Remarks")
                m_strPricesAre = rstAD.GetValue("Price_Remarks")
                m_strPkgAndFwd = rstAD.GetValue("Packing_Remarks")
                m_strFreight = rstAD.GetValue("Frieght_Remarks")
                m_strTransitInsurance = rstAD.GetValue("Transport_Remarks")
                m_strOctroi = rstAD.GetValue("Octorai_Remarks")
                m_strModeOfDespatch = rstAD.GetValue("Mode_Despatch")
                m_strDeliverySchedule = rstAD.GetValue("Delivery")
                CheckFormDetails = True
                rstAD.ResultSetClose()
            Else
                CheckFormDetails = False
                rstAD.ResultSetClose()
                Exit Function
            End If
        Else
            CheckFormDetails = True
        End If
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Function
    End Function
    Public Sub GetAmendmentDetailsForAuthorizedPO()
        '----------------------------------------------------------------------------
        'Argument       :   Nil
        'Return Value   :   Nil
        'Function       :   Get the details if there is an amendment
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        m_blnGetAmendmentDetails = True
        Dim rsAD As ClsResultSetDB
        Dim strAuthFlg As String
        m_strSql = "select Account_Code,Cust_Ref,Amendment_No,Item_Code,Rate,Order_Qty,Despatch_Qty,Active_Flag,Cust_Mtrl,Cust_DrgNo,Packing,Others,Excise_Duty,Cust_Drg_Desc,Tool_Cost,Authorized_flag,OpenSO,SalesTax_Type,PerValue,InternalSONo,RevisionNo,Remarks,MRP,Abantment_Code,AccessibleRateforMRP,Packing_Type,TOOL_AMOR_FLAG,ShowInAuth,ADD_Excise_Duty,HSNSACCODE,  ISHSNORSAC, CGSTTXRT_TYPE,SGSTTXRT_TYPE, UTGSTTXRT_TYPE,IGSTTXRT_TYPE,COMPENSATION_CESS  from cust_ord_dtl where unit_code='" & gstrUNITID & "' and  Cust_ref='" & txtReferenceNo.Text & " 'and account_Code='" & txtCustomerCode.Text & "'and Authorized_Flag=1 order by Cust_drgNo"
        rsdb = New ClsResultSetDB
        rsdb.GetResult(m_strSql)
        m_strSql = "select Account_Code,Cust_Ref,Amendment_No,Order_Date,Amendment_Date,Active_Flag,Currency_Code,Valid_Date,Effect_Date,Term_Payment,Special_Remarks,Pay_Remarks,Price_Remarks,Packing_Remarks,Frieght_Remarks,Transport_Remarks,Octorai_Remarks,Mode_Despatch,Delivery,First_Authorized,Second_Authorized,Third_Authorized,Authorized_Flag,Reason,PO_Type,SalesTax_Type,OpenSO,AddCustSupp,PerValue,InternalSONo,RevisionNo,Surcharge_code,Future_SO,ECESS_Code,Consignee_Code,ADDVAT_TYPE from cust_ord_hdr where unit_code='" & gstrUNITID & "' and Cust_ref='" & txtReferenceNo.Text & " 'and account_Code='" & txtCustomerCode.Text & "' and Consignee_Code='" & Trim(Me.txtConsCode.Text) & "' and authorized_Flag=1"
        rsAD = New ClsResultSetDB
        rsAD.GetResult(m_strSql)
        If rsAD.GetNoRows > 0 Then
            rsAD.MoveFirst()
            lblIntSONoDes.Text = rsAD.GetValue("InternalSONo")
            lblRevisionNo.Text = rsAD.GetValue("RevisionNo")
            DTAmendmentDate.Value = rsAD.GetValue("Amendment_Date")
            DTEffectiveDate.Value = rsAD.GetValue("Effect_Date")
            DTValidDate.Value = rsAD.GetValue("Valid_Date")
            txtCurrencyType.Text = rsAD.GetValue("Currency_Code")
            ctlPerValue.Text = rsAD.GetValue("PerValue")
            txtAmendReason.Text = rsAD.GetValue("Reason")
            strpotype = rsAD.GetValue("PO_Type")
            Select Case strpotype
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
            If rsAD.GetValue("OpenSO") = False Then
                chkOpenSo.CheckState = System.Windows.Forms.CheckState.Unchecked
            Else
                chkOpenSo.CheckState = System.Windows.Forms.CheckState.Checked
            End If
            txtCreditTerms.Text = rsAD.GetValue("Term_Payment")
            rsdb.MoveFirst()
            ssPOEntry.MaxRows = 0
            Do While Not rsdb.EOFRecord
                ssPOEntry.MaxRows = ssPOEntry.MaxRows + 1
                'change account Plug in
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
                Call ssPOEntry.SetText(4, ssPOEntry.MaxRows, rsdb.GetValue("Item_Code "))
                Call ssPOEntry.SetText(5, ssPOEntry.MaxRows, rsdb.GetValue("Order_Qty"))
                Call ssPOEntry.SetText(13, ssPOEntry.MaxRows, rsdb.GetValue("Rate"))
                Call ssPOEntry.SetText(6, ssPOEntry.MaxRows, rsdb.GetValue("Rate") * CDbl(ctlPerValue.Text))
                Call ssPOEntry.SetText(14, ssPOEntry.MaxRows, rsdb.GetValue("Cust_Mtrl"))
                Call ssPOEntry.SetText(7, ssPOEntry.MaxRows, rsdb.GetValue("Cust_Mtrl") * CDbl(ctlPerValue.Text))
                Call ssPOEntry.SetText(15, ssPOEntry.MaxRows, rsdb.GetValue("Tool_Cost"))
                Call ssPOEntry.SetText(8, ssPOEntry.MaxRows, rsdb.GetValue("Tool_Cost") * CDbl(ctlPerValue.Text))
                Call ssPOEntry.SetText(9, ssPOEntry.MaxRows, rsdb.GetValue("Packing"))
                Call ssPOEntry.SetText(10, ssPOEntry.MaxRows, rsdb.GetValue("Excise_Duty"))
                Call ssPOEntry.SetText(16, ssPOEntry.MaxRows, rsdb.GetValue("Others"))
                Call ssPOEntry.SetText(11, ssPOEntry.MaxRows, rsdb.GetValue("Others") * CDbl(ctlPerValue.Text))
                'GST DETAILS
                Call ssPOEntry.SetText(21, ssPOEntry.MaxRows, rsdb.GetValue("CGSTTXRT_TYPE"))
                Call ssPOEntry.SetText(22, ssPOEntry.MaxRows, rsdb.GetValue("SGSTTXRT_TYPE"))
                Call ssPOEntry.SetText(23, ssPOEntry.MaxRows, rsdb.GetValue("UTGSTTXRT_TYPE"))
                Call ssPOEntry.SetText(24, ssPOEntry.MaxRows, rsdb.GetValue("IGSTTXRT_TYPE"))
                Call ssPOEntry.SetText(25, ssPOEntry.MaxRows, rsdb.GetValue("COMPENSATION_CESS"))
                'GST DETAILS

                '********
                rsdb.MoveNext()
            Loop
        Else
            Call ConfirmWindow(10130, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
            ssPOEntry.MaxRows = 0
            rsAD.ResultSetClose()
            rsdb.ResultSetClose()
            Exit Sub
        End If
        With ssPOEntry
            .Enabled = True
            .BlockMode = True
            .Col = 3
            .Col2 = 3
            .Row = 1
            .Row2 = .MaxRows
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
            .TypeButtonPicture = My.Resources.resEmpower.ico111.ToBitmap()
            .BlockMode = False
            .Row = 1
            .Row2 = .MaxRows
            .Col = 1
            .Col2 = .MaxCols
            .BlockMode = True
            .Lock = True
            .BlockMode = False
        End With
        rsAD.ResultSetClose()
        rsdb.ResultSetClose()
        '***********Anan
        cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = True
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub TxtReferenceNo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtReferenceNo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
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
    Public Sub UpdateActiveFlag()
        '----------------------------------------------------------------------------
        'Argument       :   Nil
        'Return Value   :   Nil
        'Function       :   fills the customer detail label
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim rsAD As ClsResultSetDB
        Dim varItemCode As Object
        Dim varDrgNo As Object
        Dim AmendmentNo As String
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        Dim intmaxitems As Short
        Dim rsCustOrdHdr As ClsResultSetDB
        m_strSql = "select Account_Code,Cust_Ref,Amendment_No,Item_Code,Rate,Order_Qty,Despatch_Qty,Active_Flag,Cust_Mtrl,Cust_DrgNo,Packing,Others,Excise_Duty,Cust_Drg_Desc,Tool_Cost,Authorized_flag,OpenSO,SalesTax_Type,PerValue,InternalSONo,RevisionNo,Remarks,MRP,Abantment_Code,AccessibleRateforMRP,Packing_Type,TOOL_AMOR_FLAG,ShowInAuth,ADD_Excise_Duty,HSNSACCODE,  ISHSNORSAC, CGSTTXRT_TYPE,SGSTTXRT_TYPE, UTGSTTXRT_TYPE,IGSTTXRT_TYPE,COMPENSATION_CESS  from cust_ord_dtl where unit_code='" & gstrUNITID & "' and Cust_ref='" & txtReferenceNo.Text & "' and account_Code='" & txtCustomerCode.Text & "' and active_Flag='A'"
        rsAD = New ClsResultSetDB
        rsAD.GetResult(m_strSql)
        intmaxitems = rsAD.GetNoRows
        rsAD.ResultSetClose()
        intMaxLoop = ssPOEntry.MaxRows
        ReDim ArrDispatchQty(intMaxLoop - 1)
        For intLoopCounter = 1 To intMaxLoop
            'change account Plug in
            varItemCode = Nothing
            Call ssPOEntry.GetText(4, intLoopCounter, varItemCode)
            varDrgNo = Nothing
            Call ssPOEntry.GetText(2, intLoopCounter, varDrgNo)
            m_strSql = "select Despatch_qty from cust_ord_dtl where unit_code='" & gstrUNITID & "' and Cust_ref='" & txtReferenceNo.Text & " 'and account_Code='" & txtCustomerCode.Text & "' and active_Flag='A' and Item_Code ='" & varItemCode & "' and Cust_drgNo ='" & varDrgNo & "'"
            rsAD = New ClsResultSetDB
            rsAD.GetResult(m_strSql)
            If rsAD.GetNoRows >= 1 Then
                ArrDispatchQty(intLoopCounter - 1) = rsAD.GetValue("Despatch_qty")
                m_strSql = "update cust_ord_dtl set Active_Flag='O' where unit_code='" & gstrUNITID & "' and Cust_ref='" & txtReferenceNo.Text & " 'and account_Code='" & txtCustomerCode.Text & "' and active_Flag='A' and Item_Code='" & varItemCode & "' and Cust_drgNo ='" & varDrgNo & "'"
                mP_Connection.Execute(m_strSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            Else
                ArrDispatchQty(intLoopCounter - 1) = 0
            End If
            rsAD.ResultSetClose()
        Next
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Public Function ConfirmDeletion() As Boolean
        '----------------------------------------------------------------------------
        'Argument       :   Nil
        'Return Value   :   Nil
        'Function       :   confirms the deletion if some rows are marked for deletion
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim intRow As Short
        Dim blndelrow As Boolean
        Dim strAns As MsgBoxResult
        Dim vardelrow As Object
        ConfirmDeletion = True
        For intRow = 1 To ssPOEntry.MaxRows
            vardelrow = Nothing
            Call ssPOEntry.GetText(0, intRow, vardelrow)
            If vardelrow = "*" Then
                blndelrow = True
                Exit Function
            Else
                blndelrow = False
            End If
        Next
        If blndelrow = True Then
            Call ConfirmWindow(10101, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
            strAns = ConfirmWindow(10099, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNOCANCEL, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_EXCLAMATION)
            If strAns = MsgBoxResult.Yes Then
                Exit Function
            ElseIf strAns = MsgBoxResult.No Then
                For intRow = 1 To ssPOEntry.MaxRows
                    Call ssPOEntry.SetText(0, intRow, "")
                Next
                Exit Function
            Else
                Call cmdButtons_ButtonClick(cmdButtons, New UCActXCtl.UCbtndgrp.ButtonClickEventArgs(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL))
                ConfirmDeletion = False
                Exit Function
            End If
        End If
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Function
    End Function
    Public Sub SSMaxLength()
        On Error GoTo ErrHandler
        Dim intRow As Short
        Dim strMin As String
        Dim strMax As String
        Dim intDecimal As Short
        Dim intLoopCounter As Short
        Dim rscurrency As ClsResultSetDB
        With Me.ssPOEntry
            For intRow = 1 To .MaxRows
                .Row = intRow
                .Col = 1
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
                .Col = 2
                .TypeMaxEditLen = 30
                .Col = 3
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                .TypeButtonPicture = My.Resources.resEmpower.ico111.ToBitmap()
                .Col = 4
                .TypeMaxEditLen = 16
                .Col = 5
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                .TypeFloatDecimalPlaces = 2
                .TypeFloatMin = "0.00"
                .TypeFloatMax = "9999999.99"
                If chkOpenSo.CheckState = 1 Then
                    .Row = 1
                    .Row2 = .MaxRows
                    .Col = 1
                    .Col2 = 1
                    .BlockMode = True
                    .Lock = True
                    .BlockMode = False
                    .Row = 1
                    .Row2 = .MaxRows
                    .Col = 5
                    .Col2 = 5
                    .BlockMode = True
                    .Lock = True
                    .BlockMode = False
                Else
                    .Row = 1
                    .Row2 = .MaxRows
                    .Col = 1
                    .Col2 = 1
                    .BlockMode = True
                    .Lock = False
                    .BlockMode = False
                    .Row = 1
                    .Row2 = .MaxRows
                    .Col = 5
                    .Col2 = 5
                    .BlockMode = True
                    .Lock = False
                    .BlockMode = False
                End If
                If Len(Trim(txtCurrencyType.Text)) Then
                    rscurrency = New ClsResultSetDB
                    rscurrency.GetResult("Select decimal_Place from Currency_Mst Where unit_code='" & gstrUNITID & "' and currency_code ='" & Trim(txtCurrencyType.Text) & "'")
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
                .Row = intRow
                .Col = 5
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                .TypeFloatDecimalPlaces = intDecimal
                .TypeFloatMin = strMin
                .TypeFloatMax = strMax
                .Col = 6
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                .TypeFloatDecimalPlaces = intDecimal
                .TypeFloatMin = strMin
                .TypeFloatMax = strMax
                .Col = 7
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                .TypeFloatDecimalPlaces = intDecimal
                .TypeFloatMin = strMin
                .TypeFloatMax = strMax
                .Col = 8
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                .TypeFloatDecimalPlaces = 4
                .TypeFloatMin = "0.0000"
                .TypeFloatMax = "9999999.9999"
                .Col = 9
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                .TypeFloatDecimalPlaces = 2
                .TypeFloatMin = "0.00"
                .TypeFloatMax = "99999999.99"
                .Col = 10
                .Row = intRow
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            Next intRow
        End With
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Public Sub RefreshForm()
        Dim rsGetDate As ClsResultSetDB
        txtCurrencyType.Text = ""
        txtAmendReason.Text = ""
        lblIntSONoDes.Text = ""
        lblRevisionNo.Text = ""
        lblCustPartDesc.Text = ""
        cmbPOType.Text = ObsoleteManagement.GetItemString(cmbPOType, 0)
        DTDate.Value = GetServerDate()
        DTAmendmentDate.Value = GetServerDate()
        DTEffectiveDate.Value = GetServerDate()
        rsGetDate = New ClsResultSetDB
        rsGetDate.GetResult("select Financial_EndDate from Company_Mst where unit_code='" & gstrUNITID & "'")
        DTValidDate.Value = rsGetDate.GetValue("Financial_EndDate")
        ssPOEntry.MaxRows = 0
        m_strSalesTaxType = ""
        '10736222-changes done for CT2 ARE 3 functionality
        ChkCT2Reqd.Enabled = False
        ChkCT2Reqd.Checked = False
        rsGetDate.ResultSetClose()
        GrpSODoc.Visible = True
        GrpSODoc.Enabled = False
        lblFilePath.Text = ""
        lblFilePath.Tag = ""
        BtnUpload.Enabled = False
        btnViewDoc.Enabled = False
        btnRemoveDoc.Enabled = False
        cmbExporttype.SelectedIndex = 0
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
    Private Sub addExportSotype()

        Try
            cmbExporttype.Items.Insert(0, "None")
            cmbExporttype.Items.Insert(1, "With Pay")
            cmbExporttype.Items.Insert(2, "Without Pay")

        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try

    End Sub

    Private Function Validate_CSMRate() As Boolean
        On Error GoTo ErrHandler
        Dim strQry As String
        Dim Rs As ClsResultSetDB
        Dim kount As Integer
        Dim varItem_code, varRate As Object
        Validate_CSMRate = False
        With ssPOEntry
            For kount = 1 To .MaxRows
                varItem_code = Nothing
                Call .GetText(4, kount, varItem_code)
                varRate = Nothing
                Call .GetText(7, kount, varRate)
                strQry = "SELECT DBO.UDF_VALIDATE_CSM_RATE('" & gstrUNITID & "','" & txtCustomerCode.Text & "','" & varItem_code & "'," & Val(varRate) & ") AS A"
                Rs = New ClsResultSetDB
                If Rs.GetResult(strQry) = False Then GoTo ErrHandler
                If Len(Rs.GetValue("A")) = 0 Then
                    Validate_CSMRate = True
                Else
                    Validate_CSMRate = False
                    MsgBox(Rs.GetValue("A"), MsgBoxStyle.Information, ResolveResString(100))
                    Rs.ResultSetClose()
                    Rs = Nothing
                    Exit Function
                End If
                Rs.ResultSetClose()
            Next
        End With
        Rs = Nothing
        Exit Function
ErrHandler:
        Rs = Nothing
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Public Function ValidRecord() As Boolean
        Dim strErrMsg As String
        Dim ctlBlank As System.Windows.Forms.Control
        Dim lNo As Integer
        Dim varDrawPartNo1, varItemCode1, varOpenSO As Object
        Dim strSql As String
        Dim strCustPartNo As String = String.Empty
        Dim blnIsEopRequired As Boolean = False
        Dim StrDockcode As String
        Dim StrnewDockCode As String
        On Error GoTo Err_Handler
        blnInvalidData = False
        ValidRecord = False
        lNo = 1
        strErrMsg = ResolveResString(10059) & vbCrLf & vbCrLf
        If Validate_CSMRate() = False Then
            ValidRecord = False
            Exit Function
        End If
        If Len(Trim(txtCustomerCode.Text)) = 0 Then
            blnInvalidData = True
            strErrMsg = strErrMsg & vbCrLf & lNo & "." & "Customer Code "
            lNo = lNo + 1
            If ctlBlank Is Nothing Then ctlBlank = txtCustomerCode
        End If
        If Len(Trim(txtReferenceNo.Text)) = 0 Then
            blnInvalidData = True
            strErrMsg = strErrMsg & vbCrLf & lNo & "." & "Reference No "
            lNo = lNo + 1
            If ctlBlank Is Nothing Then ctlBlank = txtReferenceNo
        End If
        If UCase(Trim(cmbPOType.Text)) = "EXPORT" And cmbExporttype.SelectedIndex = 0 Then
            blnInvalidData = True
            strErrMsg = strErrMsg & vbCrLf & lNo & "." & "Export type "
            lNo = lNo + 1
            If ctlBlank Is Nothing Then ctlBlank = cmbExporttype
        End If
        If (UCase(Trim(gstrUNITID)) = "MST" Or UCase(Trim(gstrUNITID)) = "MAR") Then
            If txtShipAddress.Text = "" Then
                If MsgBox("Do You Want To create Sale Order without Ship address Code?", MsgBoxStyle.YesNo, ResolveResString(100)) = MsgBoxResult.No Then
                    blnInvalidData = True
                    Exit Function
                End If
            End If
        End If
       
        ' added by priti sharma on 25.03.2019 
        If chkOpenSo.Checked Then
            'strSql = "SELECT top 1 Party_Code FROM DBO.FIN_GET_RELATED_PARTIES ('" & gstrUNITID & "') WHERE CUST_VEND='C' and Party_Code='" & Trim(txtCustomerCode.Text) & "'"
            'If IsRecordExists(strSql) = True Then
            '    MsgBox("Open SO is not allowed for Related Parties ", MsgBoxStyle.Information, ResolveResString(100))
            '    Exit Function
            'End If
        End If
        'ends 
        '' added by priti on 11 May 2020 for OPEN SO Issue
        'strSql = "SELECT top 1 Party_Code FROM DBO.FIN_GET_RELATED_PARTIES ('" & gstrUNITID & "') WHERE CUST_VEND='C' and Party_Code='" & Trim(txtCustomerCode.Text) & "'"
        'If IsRecordExists(strSql) = True Then
        '    For intRow = 1 To ssPOEntry.MaxRows
        '        varOpenSO = Nothing
        '        Call ssPOEntry.GetText(1, intRow, varOpenSO)
        '        If Val(varOpenSO) = 1 Then
        '            MsgBox("Open SO is not allowed for Related Parties ", MsgBoxStyle.Information, ResolveResString(100))
        '            Exit Function
        '        End If
        '    Next
        'End If
        '' code ends here
        If txtAmendmentNo.Enabled Then
            If Len(Trim(txtAmendmentNo.Text)) = 0 Then
                blnInvalidData = True
                strErrMsg = strErrMsg & vbCrLf & lNo & "." & "Amendment No "
                lNo = lNo + 1
                If ctlBlank Is Nothing Then ctlBlank = txtAmendmentNo
            End If
        End If
        If Len(Trim(txtCurrencyType.Text)) = 0 Then
            blnInvalidData = True
            strErrMsg = strErrMsg & vbCrLf & lNo & "." & "Currency Type "
            lNo = lNo + 1
            If ctlBlank Is Nothing Then ctlBlank = txtCurrencyType
        End If
        If Len(Trim(cmbPOType.Text)) = 0 Then
            blnInvalidData = True
            strErrMsg = strErrMsg & vbCrLf & lNo & "." & "SO Type "
            lNo = lNo + 1
            If ctlBlank Is Nothing Then ctlBlank = cmbPOType
        End If
        If Len(Trim(txtCreditTerms.Text)) = 0 Then
            blnInvalidData = True
            strErrMsg = strErrMsg & vbCrLf & lNo & "." & "Credit Terms "
            lNo = lNo + 1
            If ctlBlank Is Nothing Then ctlBlank = txtCreditTerms
        End If
        If cmbPOType.SelectedIndex <> 4 Then
            If CheckFormDetails() = False Then
                blnInvalidData = True
                strErrMsg = strErrMsg & vbCrLf & lNo & "." & "Sales Terms "
                lNo = lNo + 1
                If ctlBlank Is Nothing Then ctlBlank = cmdchangetype
            End If
        End If
        '10869290

        If UCase(Trim(cmbPOType.Text)) = "V-SERVICE".ToUpper And gstrUNITID <> "STH" Then
            If Len(Trim(txtService.Text)) = 0 And gblnGSTUnit = False Then
                blnInvalidData = True
                strErrMsg = strErrMsg & vbCrLf & lNo & "." & "Service Tax "
                lNo = lNo + 1
                If ctlBlank Is Nothing Then ctlBlank = txtService
            End If
        End If



        If UCase(Trim(cmbPOType.Text)) = "V-SERVICE".ToUpper And gstrUNITID <> "STH" Then
            If Len(Trim(txtSBC.Text)) = 0 And (gblnGSTUnit = False) Then
                blnInvalidData = True
                strErrMsg = strErrMsg & vbCrLf & lNo & "." & "SBC Tax Code "
                lNo = lNo + 1
                If ctlBlank Is Nothing Then ctlBlank = txtSBC
            End If

            If Len(Trim(txtKKC.Text)) = 0 And (gblnGSTUnit = False) Then
                blnInvalidData = True
                strErrMsg = strErrMsg & vbCrLf & lNo & "." & "KKC Tax Code "
                lNo = lNo + 1
                If ctlBlank Is Nothing Then ctlBlank = txtKKC
            End If
        End If


        strErrMsg = VB.Left(strErrMsg, Len(strErrMsg) - 1)
        strErrMsg = strErrMsg & " ."
        lNo = lNo + 1
        If blnInvalidData = True Then
            gblnCancelUnload = True
            Call MsgBox(strErrMsg, MsgBoxStyle.Information, ResolveResString(100))
            ctlBlank.Focus()
            Exit Function
        End If

        '10808160--Starts
        If UCase(Trim(cmbPOType.Text)) <> "SPARES".ToUpper Then
            strSql = "Select dbo.UDF_ISEOPREQUIRED('" & gstrUNITID & "','" & Trim(txtCustomerCode.Text) & "')"
            blnIsEopRequired = Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSql))
            If blnIsEopRequired = True Then
                With ssPOEntry
                    For intRow = 1 To ssPOEntry.MaxRows
                        varDrawPartNo1 = Nothing
                        varItemCode1 = Nothing
                        Call ssPOEntry.GetText(2, intRow, varDrawPartNo1)
                        Call ssPOEntry.GetText(4, intRow, varItemCode1)

                        strSql = "SELECT TOP 1 CUST_DRGNO FROM BUDGETITEM_MST WHERE UNIT_CODE= '" & gstrUNITID & "' AND ACCOUNT_CODE = '" & Trim(txtCustomerCode.Text) & "' " & _
                                        " AND ITEM_CODE = '" & Trim(varItemCode1) & "' AND CUST_DRGNO = '" & Trim(varDrawPartNo1) & "' AND ENDDATE < '" & Convert.ToDateTime(DTValidDate.Value).ToString("dd MMM yyyy") & "' "
                        If IsRecordExists(strSql) = True Then
                            strCustPartNo = strCustPartNo + varDrawPartNo1 + " , "
                        End If
                    Next

                    If strCustPartNo <> "" Then
                        If MsgBox("End Date of Following Cust. Part: " + " '" & Trim(strCustPartNo) & "' " + "  " & vbCrLf & "In this SO are falling before the Validity Date of the SO." & vbCrLf & "Do you want to continue?", MsgBoxStyle.YesNo, ResolveResString(100)) = MsgBoxResult.No Then
                            Exit Function
                        End If
                    End If
                End With
            End If
        End If

        '10808160--Ends
        'strSql = "select dbo.UDF_ISDOCKCODE_ASN( '" & gstrUNITID & "','" & txtCustomerCode.Text.Trim & "')"
        'If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSql)) = True Then
        '    strCustPartNo = ""
        '    With ssPOEntry
        '        For intRow = 1 To ssPOEntry.MaxRows
        '            varDrawPartNo1 = Nothing
        '            varItemCode1 = Nothing
        '            Call ssPOEntry.GetText(2, intRow, varDrawPartNo1)
        '            Call ssPOEntry.GetText(4, intRow, varItemCode1)

        '            strSql = "SELECT ISNULL(DOCKCODE,'') DOCKCODE FROM CUSTITEM_MST WHERE UNIT_CODE= '" & gstrUNITID & "' AND ACCOUNT_CODE = '" & Trim(txtCustomerCode.Text) & "' " & _
        '                            " AND ITEM_CODE = '" & Trim(varItemCode1) & "' AND CUST_DRGNO = '" & Trim(varDrawPartNo1) & "' AND ACTIVE= 1 AND DOCKCODE >0  "

        '            If IsRecordExists(strSql) = True Then
        '                strSql = "SELECT TOP 1 CUST_DRGNO FROM CUSTITEM_MST WHERE UNIT_CODE= '" & gstrUNITID & "' AND ACCOUNT_CODE = '" & Trim(txtCustomerCode.Text) & "' " & _
        '                        " AND ITEM_CODE = '" & Trim(varItemCode1) & "' AND CUST_DRGNO = '" & Trim(varDrawPartNo1) & "' AND ACTIVE=1 "

        '                If intRow = ssPOEntry.MaxRows Then 'for last record
        '                    strCustPartNo = strCustPartNo + varDrawPartNo1
        '                Else
        '                    strCustPartNo = strCustPartNo + varDrawPartNo1 + " , "
        '                End If

        '            End If
        '        Next
        '    End With
        '    strSql = "SELECT top 1 1 FROM CUSTITEM_MST WHERE UNIT_CODE='" + gstrUNITID + "' AND ACTIVE=1 AND ACCOUNT_CODE='" + txtCustomerCode.Text.Trim() + "'" & _
        '                                                        " AND CUST_DRGNO in('" & Trim(varDrawPartNo1 & "')"  & " HAVING COUNT(DISTINCT DOCKCODE )>1 " 

        '    If IsRecordExists(strSql) = True Then
        '        blnInvalidData = True
        '        strErrMsg = strErrMsg & vbCrLf & lNo & "." & "TWO DOCKCODE ,SAME SALES ORDER CANT BE SAVED"
        '        lNo = lNo + 1
        '    End If
        'End If


        ''10856126 ASN DOCK CODE 

        '10856126 END 
        ValidRecord = True
        gblnCancelUnload = True
        gblnFormAddEdit = True
        Exit Function
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Public Sub CLEARVAR()
        m_strSalesTaxType = ""
        m_intCreditDays = ""
        m_strSpecialNotes = ""
        m_strPaymentTerms = ""
        m_strPricesAre = ""
        m_strPkgAndFwd = ""
        m_strFreight = ""
        m_strTransitInsurance = ""
        m_strOctroi = ""
        m_strModeOfDespatch = ""
        m_strDeliverySchedule = ""

    End Sub
    Public Function Find_Value(ByRef strField As String) As String
        On Error GoTo ErrHandler
        Dim Rs As ADODB.Recordset
        Rs = New ADODB.Recordset
        Rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
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
    Public Function addExciseDuty(ByRef pstrRow As Integer) As Boolean
        Dim varItem_Code As Object
        Dim strExDuty As String
        Dim rsTariffCode As ClsResultSetDB
        'change account Plug in
        varItem_Code = Nothing
        If gstrUNITID = "STH" Then
            Exit Function
        Else

            Call ssPOEntry.GetText(4, pstrRow, varItem_Code)
            rsTariffCode = New ClsResultSetDB

            rsTariffCode.GetResult("Select a.Excise_Duty from Tax_tariff_Mst a,ITem_Mst b where A.UNIT_CODE = B.UNIT_CODE AND b.Tariff_Code = a.Tariff_subhead and b.Item_code ='" & varItem_Code & "' AND A.UNIT_CODE='" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If rsTariffCode.GetNoRows > 0 Then
                strExDuty = rsTariffCode.GetValue("Excise_duty")
                If gblnGSTUnit = False Then
                    Call ssPOEntry.SetText(10, pstrRow, strExDuty)
                Else
                    Call ssPOEntry.SetText(10, pstrRow, "")
                End If

                addExciseDuty = True
            Else
                MsgBox("Tariff - Item Relationship Not Defined in Tariff Master.", MsgBoxStyle.Information, "empower")
                addExciseDuty = False
            End If
            rsTariffCode.ResultSetClose()
        End If
    End Function
    Public Sub UpdateHdrActiveFlag()
        Dim rsCustOrdHdr As ClsResultSetDB
        Dim rsCustOrdDtl As ClsResultSetDB
        Dim intMaxLoop As Short
        Dim intLoopCounter As Short
        Dim intmaxitems As Short
        Dim intMaxOverItem As Short
        Dim AmendmentNo As String
        rsCustOrdHdr = New ClsResultSetDB
        m_strSql = "select distinct(AmendMEnt_No)from cust_ord_hdr where UNIT_CODE='" & gstrUNITID & "' AND  Cust_ref='" & txtReferenceNo.Text & " 'and account_Code='" & txtCustomerCode.Text & "' and Consignee_Code='" & Trim(txtConsCode.Text) & "'  and active_Flag='A'"
        rsCustOrdHdr.GetResult(m_strSql)
        intMaxLoop = rsCustOrdHdr.GetNoRows
        rsCustOrdHdr.MoveFirst()
        For intLoopCounter = 1 To intMaxLoop
            AmendmentNo = Trim(rsCustOrdHdr.GetValue("Amendment_No"))
            m_strSql = "select Account_Code,Cust_Ref,Amendment_No,Item_Code,Rate,Order_Qty,Despatch_Qty,Active_Flag,Cust_Mtrl,Cust_DrgNo,Packing,Others,Excise_Duty,Cust_Drg_Desc,Tool_Cost,Authorized_flag,OpenSO,SalesTax_Type,PerValue,InternalSONo,RevisionNo,Remarks,MRP,Abantment_Code,AccessibleRateforMRP,Packing_Type,TOOL_AMOR_FLAG,ShowInAuth,ADD_Excise_Duty,HSNSACCODE,  ISHSNORSAC, CGSTTXRT_TYPE,SGSTTXRT_TYPE, UTGSTTXRT_TYPE,IGSTTXRT_TYPE,COMPENSATION_CESS  from cust_ord_dtl where UNIT_CODE='" & gstrUNITID & "' AND Cust_ref='" & txtReferenceNo.Text & " 'and account_Code='" & txtCustomerCode.Text & "' and amendment_no ='" & AmendmentNo & "' and active_Flag='O'"
            rsCustOrdDtl = New ClsResultSetDB
            rsCustOrdDtl.GetResult(m_strSql)
            intMaxOverItem = rsCustOrdDtl.GetNoRows
            m_strSql = "select Account_Code,Cust_Ref,Amendment_No,Item_Code,Rate,Order_Qty,Despatch_Qty,Active_Flag,Cust_Mtrl,Cust_DrgNo,Packing,Others,Excise_Duty,Cust_Drg_Desc,Tool_Cost,Authorized_flag,OpenSO,SalesTax_Type,PerValue,InternalSONo,RevisionNo,Remarks,MRP,Abantment_Code,AccessibleRateforMRP,Packing_Type,TOOL_AMOR_FLAG,ShowInAuth,ADD_Excise_Duty,HSNSACCODE,  ISHSNORSAC, CGSTTXRT_TYPE,SGSTTXRT_TYPE, UTGSTTXRT_TYPE,IGSTTXRT_TYPE,COMPENSATION_CESS from cust_ord_dtl where UNIT_CODE='" & gstrUNITID & "' AND Cust_ref='" & txtReferenceNo.Text & " 'and account_Code='" & txtCustomerCode.Text & "' and amendment_no ='" & AmendmentNo & "'"
            rsCustOrdDtl.GetResult(m_strSql)
            intmaxitems = rsCustOrdDtl.GetNoRows
            If intmaxitems = intMaxOverItem Then
                m_strSql = "Update cust_ord_hdr set active_Flag='O' where UNIT_CODE='" & gstrUNITID & "' AND Cust_ref='" & txtReferenceNo.Text & " 'and account_Code='" & txtCustomerCode.Text & "' and Consignee_Code='" & Trim(txtConsCode.Text) & "' and active_Flag='A' and amendment_no ='" & AmendmentNo & "'"
                mP_Connection.Execute(m_strSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End If
            rsCustOrdDtl.ResultSetClose()
            rsCustOrdHdr.MoveNext()
        Next
        rsCustOrdHdr.ResultSetClose()
    End Sub
    Public Sub PrintToReport()
        '*********************************************'
        'Author:                Ananya Nath
        'Arguments:             None
        'Return Value   :       None
        'Description    :       Used to print currently selected/entered sales Order.
        '*********************************************'
        On Error GoTo ErrHandler
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.AppStarting)
        Dim frmRpt As New eMProCrystalReportViewer
        Dim CR As ReportDocument
        CR = frmRpt.GetReportDocument()
        With frmRpt
            strSQL = ""
            If Len(Trim(Me.txtAmendmentNo.Text)) = 0 Then
                strSQL = " {cust_ord_hdr.unit_code}='" & gstrUNITID & "' AND {cust_ord_hdr.account_code} = '" & txtCustomerCode.Text & "' and {cust_ord_hdr.Consignee_Code} ='" & Trim(txtConsCode.Text) & "'  and {cust_ord_hdr.cust_ref} = '" & txtReferenceNo.Text & "' and  {cust_ord_hdr.amendment_no} = ''" 'Initialising Sql Query.
            Else
                strSQL = " {cust_ord_hdr.unit_code}='" & gstrUNITID & "' AND {cust_ord_hdr.account_code} = '" & txtCustomerCode.Text & "' and {cust_ord_hdr.Consignee_Code} ='" & Trim(txtConsCode.Text) & "'  and {cust_ord_hdr.cust_ref} = '" & txtReferenceNo.Text & "' and  {cust_ord_hdr.amendment_no} = '" & txtAmendmentNo.Text & "'" 'Initialising Sql Query.
            End If
            '************************
            CR.Load(My.Application.Info.DirectoryPath & "\Reports\rptSOPrinting.rpt")
            'Me.CrystalReport1.WindowState = Crystal.WindowStateConstants.crptMaximized
            'Me.CrystalReport1.set_Formulas(1, "Comp_name = '" & gstrCOMPANY & "'") ' company name will be printed
            'Me.CrystalReport1.set_Formulas(2, "Comp_address = '" & gstr_WRK_ADDRESS1 & " ' + '" & gstr_WRK_ADDRESS2 & " ' ") ' address will be printed
            CR.DataDefinition.FormulaFields("Comp_name").Text = "'" + gstrCOMPANY + "'"
            CR.DataDefinition.FormulaFields("Comp_address").Text = "'" + gstr_WRK_ADDRESS1 + gstr_WRK_ADDRESS2 + "'"
            .ShowExportButton = True
            'Me.CrystalReport1.WindowShowExportBtn = False
            'Me.CrystalReport1.WindowMaxButton = False
            'Me.CrystalReport1.WindowMinButton = False
            '12/12/2002 Added by nisha issue log no 1399
            'Me.CrystalReport1.WindowShowPrintSetupBtn = False
            'Me.CrystalReport1.WindowShowSearchBtn = True
            .ShowTextSearchButton = True
            'Me.CrystalReport1.SelectionFormula = strSQL
            CR.RecordSelectionFormula = strSQL
            'Me.CrystalReport1.WindowTitle = "Customer Purchase Order"
            .ReportHeader = "Customer Purchase Order"
            'Me.CrystalReport1.ReportFileName = My.Application.Info.DirectoryPath & "\Reports\rptSOPrinting.rpt"
            'Me.CrystalReport1.Action = 1
            .Show()
            'Me.CrystalReport1.PageZoom((120))
            .Zoom = 120
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
            '**************************
        End With
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
    End Sub
    Private Sub txtSTax_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSTax.TextChanged
        Call FillLabel("STAX")
    End Sub
    Private Sub txtSChSTax_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSChSTax.TextChanged
        Call FillLabel("SSCHTAX")
    End Sub
    Private Sub txtSTax_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtSTax.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case 112
                Call cmdHelp_Click(cmdHelp.Item(4), New System.EventArgs())
        End Select
    End Sub
    Private Sub txtSChSTax_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtSChSTax.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        Select Case KeyCode
            Case 112
                Call cmdHelp_Click(cmdHelp.Item(6), New System.EventArgs())
        End Select
    End Sub
    Private Sub txtSTax_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtSTax.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                txtSChSTax.Focus()
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtSChSTax_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtSChSTax.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                txtCreditTerms.Focus()
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtSTax_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSTax.Leave
        If Len(Trim(txtSTax.Text)) <> 0 Then
            m_strSql = "select TxRt_Rate_No,TxRt_RateDesc from Gen_TaxRate where unit_code='" & gstrUNITID & "' and Txrt_Rate_No = '" & txtSTax.Text & "'"
            rsdb = New ClsResultSetDB
            rsdb.GetResult(m_strSql)
            If rsdb.GetNoRows > 0 Then
                Call FillLabel("STAX")
            Else
                MsgBox("Entered S.Tax Code does not exist", MsgBoxStyle.Information, "empower")
                txtSTax.Text = ""
                txtSTax.Focus()
            End If
            rsdb.ResultSetClose()
        End If
    End Sub
    Private Sub txtSChSTax_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSChSTax.Leave
        If Len(Trim(txtSChSTax.Text)) <> 0 Then
            m_strSql = "select TxRt_Rate_No,TxRt_RateDesc from Gen_TaxRate where unit_code='" & gstrUNITID & "' and Txrt_Rate_No = '" & txtSChSTax.Text & "'"
            rsdb = New ClsResultSetDB
            rsdb.GetResult(m_strSql)
            If rsdb.GetNoRows > 0 Then
                Call FillLabel("SSCHTAX")
            Else
                MsgBox("Entered S.Tax Surcharge Code does not exist", MsgBoxStyle.Information, "empower")
                txtSChSTax.Text = ""
                txtSChSTax.Focus()
            End If
            rsdb.ResultSetClose()
        End If
    End Sub
    Public Function checkforitemRate(ByRef pintRow As Short) As Boolean
        Dim varDrgNo As Object
        Dim strItemRate As String
        Dim rsItemRate As ClsResultSetDB
        With ssPOEntry
            varDrgNo = Nothing
            Call .GetText(2, pintRow, varDrgNo)
            strItemRate = "select Edit_flg from ITemRate_Mst where unit_code='" & gstrUNITID & "' and Serial_No = (select max(serial_no) from "
            strItemRate = strItemRate & " itemrate_mst where unit_code='" & gstrUNITID & "' and Party_code = '" & txtCustomerCode.Text
            strItemRate = strItemRate & "' and item_code = '" & varDrgNo
            strItemRate = strItemRate & "' and datediff(mm,convert(varchar(10),'" & getDateForDB(DTDate.Value) & "',103),convert(varchar(10),DateFrom,103))<=0 "
            strItemRate = strItemRate & " and custVend_Flg ='C')"
            rsItemRate = New ClsResultSetDB
            rsItemRate.GetResult(strItemRate)
            checkforitemRate = rsItemRate.GetValue("Edit_Flg")
            rsItemRate.ResultSetClose()
        End With
    End Function
    Public Sub SetCellTypeCombo(ByVal intRow As Short)
        Dim strcustdtl As String
        Dim StrItemCode As Object
        Dim strDrgNo As Object
        Dim FinalstrItemCode As String
        Dim rsitem As ClsResultSetDB
        If IsNothing(rsitem) = False Then rsitem.ResultSetClose()
        rsitem = New ClsResultSetDB
        strDrgNo = Nothing
        Call ssPOEntry.GetText(2, intRow, strDrgNo)

        strcustdtl = "SElect isnull(item_code,'') as item_code from custITem_Mst where unit_code='" & gstrUNITID & "' and cust_drgNo ='" & strDrgNo & "'and Account_code ='" & txtCustomerCode.Text & "' and active = 1" '10532789 active = 1 added
        rsitem.GetResult(strcustdtl)
        With ssPOEntry
            .Col = 4
            .Row = intRow
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeComboBox
            .TypeComboBoxClear(4, .Row)
            rsitem.MoveFirst()
            FinalstrItemCode = ""
            While Not rsitem.EOFRecord
                StrItemCode = IIf((rsitem.GetValue("ITem_code") = "Unknown"), "", rsitem.GetValue("ITem_code"))
                FinalstrItemCode = FinalstrItemCode & StrItemCode & Chr(9) '& "[V]alue":
                rsitem.MoveNext()
            End While
            FinalstrItemCode = VB.Left(FinalstrItemCode, Len(FinalstrItemCode) - 1)
            .TypeComboBoxList = FinalstrItemCode
        End With
        rsitem.ResultSetClose()
    End Sub
    Public Sub SetCellStatic(ByRef intRow As Integer)
        With ssPOEntry
            .Col = 4
            .Row = intRow
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
        End With
    End Sub
    Public Function chkmultipleitem() As Boolean
        Dim StrItemCode As Object
        Dim strDrgNo As Object
        Dim rsitem As ClsResultSetDB
        rsitem = New ClsResultSetDB
        strDrgNo = Nothing
        Call ssPOEntry.GetText(2, ssPOEntry.ActiveRow, strDrgNo)

        StrItemCode = "Select top 1 1 from custITem_Mst where unit_code='" & gstrUNITID & "' and cust_drgNo ='" & strDrgNo & "'and Account_code ='" & txtCustomerCode.Text & "' and active = 1" '10532789 active = 1 added
        rsitem.GetResult(StrItemCode)
        If rsitem.GetNoRows > 1 Then
            chkmultipleitem = True
        Else
            chkmultipleitem = False
        End If
        rsitem.ResultSetClose()
    End Function
    Public Sub RowDetailsfromKeyBoard(ByRef pstrItemCode As Object, ByRef pstrDrgno As Object)
        Dim rsitem As ClsResultSetDB
        If cmdButtons.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            If Len(Trim(pstrItemCode)) > 0 Then
                m_strSql = "Select Account_Code,Cust_Ref,Amendment_No,Item_Code,Rate,Order_Qty,Despatch_Qty,Active_Flag,Cust_Mtrl,Cust_DrgNo,Packing,Others,Excise_Duty,Cust_Drg_Desc,Tool_Cost,Authorized_flag,OpenSO,SalesTax_Type,PerValue,InternalSONo,RevisionNo,Remarks,MRP,Abantment_Code,AccessibleRateforMRP,Packing_Type,TOOL_AMOR_FLAG,ShowInAuth,ADD_Excise_Duty,HSNSACCODE,  ISHSNORSAC, CGSTTXRT_TYPE,SGSTTXRT_TYPE, UTGSTTXRT_TYPE,IGSTTXRT_TYPE,COMPENSATION_CESS from Cust_ord_dtl where unit_code='" & gstrUNITID & "' and Item_code ='" & pstrItemCode & "' and cust_drgNo ='" & pstrDrgno & "' and cust_ref ='" & txtReferenceNo.Text & "' and amendment_no='" & Trim(txtAmendmentNo.Text) & "' and Active_flag ='A' "
                If IsNothing(rsitem) = False Then rsitem.ResultSetClose()

                rsitem = New ClsResultSetDB
                rsitem.GetResult(m_strSql)
                If rsitem.GetNoRows > 0 Then
                    Call ssPOEntry.SetText(2, ssPOEntry.MaxRows, rsitem.GetValue("Cust_DrgNo"))
                    Call ssPOEntry.SetText(4, ssPOEntry.MaxRows, rsitem.GetValue("Item_Code "))
                    lblCustPartDesc.Text = rsitem.GetValue("Cust_drg_desc")
                    Call ssPOEntry.SetText(5, ssPOEntry.MaxRows, rsitem.GetValue("Order_Qty"))
                    Call ssPOEntry.SetText(13, ssPOEntry.MaxRows, rsitem.GetValue("Rate"))
                    Call ssPOEntry.SetText(6, ssPOEntry.MaxRows, rsitem.GetValue("Rate") * CDbl(ctlPerValue.Text))
                    Call ssPOEntry.SetText(14, ssPOEntry.MaxRows, rsitem.GetValue("Cust_Mtrl"))
                    Call ssPOEntry.SetText(7, ssPOEntry.MaxRows, rsitem.GetValue("Cust_Mtrl") * CDbl(ctlPerValue.Text))
                    Call ssPOEntry.SetText(15, ssPOEntry.MaxRows, rsitem.GetValue("Tool_Cost"))
                    Call ssPOEntry.SetText(8, ssPOEntry.MaxRows, rsitem.GetValue("Tool_Cost") * CDbl(ctlPerValue.Text))
                    Call ssPOEntry.SetText(9, ssPOEntry.MaxRows, rsitem.GetValue("Packing"))
                    Call ssPOEntry.SetText(10, ssPOEntry.MaxRows, rsitem.GetValue("Excise_Duty"))
                    Call ssPOEntry.SetText(16, ssPOEntry.MaxRows, rsitem.GetValue("Others"))
                    Call ssPOEntry.SetText(11, ssPOEntry.MaxRows, rsitem.GetValue("Others") * CDbl(ctlPerValue.Text))
                    rsitem.ResultSetClose()
                Else
                    If txtAmendmentNo.Enabled = False Then
                        If Len(Trim(pstrItemCode)) > 0 Then
                            m_strSql = " select datefrom,dateto,Party_code,Item_code,Custvend_flg,Rate,serial_no,Discount_Flag,Discount_Amount,Cust_Supplied_Material,Tool_Cost,Packaging_Flag,Packaging_Amount,Others,currency_code,edit_flg from ITemRate_Mst where unit_code='" & gstrUNITID & "' and Serial_No = (select max(serial_no) from itemrate_mst where unit_code='" & gstrUNITID & "' and Party_code = '" & txtCustomerCode.Text & "' and item_code = '" & pstrDrgno & "' and datediff(mm,convert(varchar(10),'" & getDateForDB(DTDate.Value) & "',103),convert(varchar(10),DateFrom,103))<=0 and custVend_Flg ='C')"
                            If IsNothing(rsitem) = False Then rsitem.ResultSetClose()
                            rsitem = New ClsResultSetDB
                            rsitem.GetResult(m_strSql)
                            If rsitem.GetNoRows > 0 Then
                                If chkOpenSo.CheckState = System.Windows.Forms.CheckState.Unchecked Then
                                    ssPOEntry.Row = ssPOEntry.MaxRows
                                    ssPOEntry.Col = 1
                                    ssPOEntry.Col2 = 1
                                    ssPOEntry.Value = System.Windows.Forms.CheckState.Unchecked
                                    ssPOEntry.Row = ssPOEntry.MaxRows
                                    ssPOEntry.Row2 = ssPOEntry.MaxRows
                                    ssPOEntry.Col = 5
                                    ssPOEntry.Col2 = 5
                                    ssPOEntry.BlockMode = True
                                    ssPOEntry.Lock = False
                                    ssPOEntry.BlockMode = False
                                Else
                                    ssPOEntry.Row = ssPOEntry.MaxRows
                                    ssPOEntry.Col = 1
                                    ssPOEntry.Col2 = 1
                                    ssPOEntry.Value = System.Windows.Forms.CheckState.Checked
                                    ssPOEntry.Row = ssPOEntry.MaxRows
                                    ssPOEntry.Row2 = ssPOEntry.MaxRows
                                    ssPOEntry.Col = 5
                                    ssPOEntry.Col2 = 5
                                    ssPOEntry.BlockMode = True
                                    ssPOEntry.Lock = True
                                    ssPOEntry.BlockMode = False
                                End If
                                Call ssPOEntry.SetText(2, ssPOEntry.MaxRows, IIf((rsitem.GetValue("ITem_code") = "Unknown"), "", rsitem.GetValue("ITem_code")))
                                Call ssPOEntry.SetText(4, ssPOEntry.MaxRows, pstrItemCode)
                                Call ssPOEntry.SetText(13, ssPOEntry.MaxRows, rsitem.GetValue("Rate"))
                                Call ssPOEntry.SetText(6, ssPOEntry.MaxRows, rsitem.GetValue("Rate") * CDbl(ctlPerValue.Text))
                                Call ssPOEntry.SetText(14, ssPOEntry.MaxRows, rsitem.GetValue("Cust_supplied_Material"))
                                Call ssPOEntry.SetText(7, ssPOEntry.MaxRows, rsitem.GetValue("Cust_supplied_Material") * CDbl(ctlPerValue.Text))
                                Call ssPOEntry.SetText(15, ssPOEntry.MaxRows, rsitem.GetValue("Tool_Cost"))
                                Call ssPOEntry.SetText(8, ssPOEntry.MaxRows, rsitem.GetValue("Tool_Cost") * CDbl(ctlPerValue.Text))
                                If rsitem.GetValue("Packaging_Flag") = False Then
                                    Call ssPOEntry.SetText(9, ssPOEntry.MaxRows, (rsitem.GetValue("Packaging_Amount") * 100) / rsitem.GetValue("Rate"))
                                Else
                                    Call ssPOEntry.SetText(9, ssPOEntry.MaxRows, rsitem.GetValue("Packaging_Amount"))
                                End If
                                Call ssPOEntry.SetText(16, ssPOEntry.MaxRows, rsitem.GetValue("Others"))
                                Call ssPOEntry.SetText(11, ssPOEntry.MaxRows, rsitem.GetValue("Others") * CDbl(ctlPerValue.Text))
                                If rsitem.GetValue("Edit_flg") = False Then
                                    ssPOEntry.Col = 6
                                    ssPOEntry.Col2 = ssPOEntry.MaxCols
                                    ssPOEntry.Row = ssPOEntry.MaxRows
                                    ssPOEntry.Row2 = ssPOEntry.MaxRows
                                    ssPOEntry.BlockMode = True
                                    ssPOEntry.Lock = True
                                    ssPOEntry.BlockMode = False
                                Else
                                    ssPOEntry.Col = 6
                                    ssPOEntry.Col2 = ssPOEntry.MaxCols
                                    ssPOEntry.Row = ssPOEntry.MaxRows
                                    ssPOEntry.Row2 = ssPOEntry.MaxRows
                                    ssPOEntry.BlockMode = True
                                    ssPOEntry.Lock = False
                                    ssPOEntry.BlockMode = False
                                End If
                                rsitem.ResultSetClose()
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Call ssSetFocus(ssPOEntry.MaxRows, 3)
    End Sub
    Public Function GenerateDocumentNumber(ByVal pstrTableName As String, ByVal pstrDocNofield As String, ByRef pstrDateFieldName As String, ByVal pstrWantedDate As String) As String
        On Error GoTo ErrHandler
        Dim strCheckDOcNo As String 'Gets the Doc Number from Back End
        Dim strTempSeries As String 'Find the Numeric series in Doc No
        Dim NewTempSeries As String 'Generate a NEW Series
        Dim rsDocumentNoSO As ClsResultSetDB
        rsDocumentNoSO = New ClsResultSetDB
        If Len(Trim(pstrWantedDate)) > 0 Then 'For Post Dated Docs
            'No need to check for Previously made documents for After Dates
            mP_Connection.Execute("Set dateformat 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            rsDocumentNoSO.GetResult("Select DocNo = Max(convert(int,substring(" & pstrDocNofield & ",9,7))) from " & pstrTableName & " Where unit_code='" & gstrUNITID & "' and datePart(mm,ent_dt) = datePart(mm,'" & pstrWantedDate & "') and datePart(yyyy,ent_dt) = datePart(yyyy,'" & pstrWantedDate & "')")
            strCheckDOcNo = rsDocumentNoSO.GetValue("DocNo")
        End If
        rsDocumentNoSO.ResultSetClose()
        If Len(Trim(strCheckDOcNo)) > 0 Then 'That is the Document is Made for that Period
            'Add 1 to it
            strTempSeries = CStr(CDbl(strCheckDOcNo) + 1)
            If Val(strTempSeries) < 9999 Then
                strTempSeries = New String("0", 4 - Len(strTempSeries)) & strTempSeries 'Concatenate Zeroes before the Number
            End If
            strCheckDOcNo = CStr(DatePart(Microsoft.VisualBasic.DateInterval.Year, CDate(pstrWantedDate))) & "-"
            strCheckDOcNo = strCheckDOcNo & New String("0", 2 - Len(CStr(DatePart(Microsoft.VisualBasic.DateInterval.Month, CDate(pstrWantedDate))))) & CStr(DatePart(Microsoft.VisualBasic.DateInterval.Month, CDate(pstrWantedDate))) & "-"
            strCheckDOcNo = strCheckDOcNo & strTempSeries
            GenerateDocumentNumber = strCheckDOcNo
        Else 'The Document has not been made for that Period
            NewTempSeries = NewTempSeries & CStr(DatePart(Microsoft.VisualBasic.DateInterval.Year, CDate(pstrWantedDate))) & "-"
            NewTempSeries = NewTempSeries & New String("0", 2 - Len(CStr(DatePart(Microsoft.VisualBasic.DateInterval.Month, CDate(pstrWantedDate))))) & CStr(DatePart(Microsoft.VisualBasic.DateInterval.Month, CDate(pstrWantedDate))) & "-"
            NewTempSeries = NewTempSeries & "0001"
            GenerateDocumentNumber = NewTempSeries 'The Number Is Generated
        End If
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Exit Function
    End Function
    Public Function GenerateRevisionNo() As Short
        Dim rsRevisionNo As ClsResultSetDB
        rsRevisionNo = New ClsResultSetDB
        rsRevisionNo.GetResult("Select Revision = Max(RevisionNo) from Cust_ord_hdr where unit_code='" & gstrUNITID & "' and cust_ref = '" & txtReferenceNo.Text & "'")
        GenerateRevisionNo = IIf(IsDBNull(rsRevisionNo.GetValue("Revision")), 1, Val(rsRevisionNo.GetValue("Revision")) + 1)
    End Function
    Public Sub InsertPreviousSODetails(ByRef pstrAccountCode As String, ByRef pstrRef As String, ByRef pstrAmendment As String, ByRef pstrInternalSONo As String, ByRef pintRevisionNo As Short)
        '*********************************************'
        'Author:                Nisha Rai
        'Arguments:             Account_code , CustRef,Amendment_no , IntSONo,RevisionNo
        'Return Value   :       None
        'Description    :       To Insert active item details from base SO & its amendment which are not there in Grid.
        '*********************************************'
        Dim strSQL As String
        Dim strDrgItem As String
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        Dim varItemCode As Object
        Dim varDrgNo As Object
        Dim VarDelete As Object
        Dim rsCustOrdDtl As ClsResultSetDB
        On Error GoTo ErrHandler
        strSQL = "insert into cust_ord_dtl (unit_code,Account_Code,Cust_Ref, Amendment_No,Item_Code,Rate,Order_Qty,Despatch_Qty,Active_Flag,Cust_Mtrl,"
        strSQL = strSQL & " Cust_DrgNo,Packing,Others,Excise_Duty,Cust_Drg_Desc,Tool_Cost,Authorized_flag,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,"
        strSQL = strSQL & " OpenSO,SalesTax_Type,PerValue,InternalSONo,RevisionNo,EXTERNAL_SALESORDER_NO,ADD_Excise_Duty,HSNSACCODE,  ISHSNORSAC, CGSTTXRT_TYPE,SGSTTXRT_TYPE, UTGSTTXRT_TYPE,IGSTTXRT_TYPE,COMPENSATION_CESS  )"
        strSQL = strSQL & " (Select unit_code,Account_Code,Cust_Ref, Amendment_No = '" & pstrAmendment & "',Item_Code,Rate,Order_Qty,Despatch_Qty = 0 ,"
        strSQL = strSQL & " Active_Flag ,Cust_Mtrl,Cust_DrgNo,Packing,Others, Excise_Duty,Cust_Drg_Desc,Tool_Cost,Authorized_flag = 0 "
        strSQL = strSQL & " ,Ent_dt,Ent_UserId,Upd_dt,Upd_UserId,OpenSO,SalesTax_Type,PerValue,InternalSONo = '" & pstrInternalSONo & "',"
        strSQL = strSQL & " RevisionNo = " & pintRevisionNo & ",EXTERNAL_SALESORDER_NO , ADD_Excise_Duty,HSNSACCODE,  ISHSNORSAC, CGSTTXRT_TYPE,SGSTTXRT_TYPE, UTGSTTXRT_TYPE,IGSTTXRT_TYPE,COMPENSATION_CESS from Cust_ord_dtl where unit_code='" & gstrUNITID & "' and Account_code = '" & pstrAccountCode & "' "
        strSQL = strSQL & " and cust_ref = '" & pstrRef & "' and Active_flag = 'A' and authorized_flag = 1 "
        strSQL = strSQL & " and amendment_no <> '" & pstrAmendment & "'"
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
                    Call .GetText(4, intLoopCounter, varItemCode)
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
                strSQL = strSQL & strDrgItem & ")"
            End If
        End If
        mP_Connection.BeginTrans()
        mP_Connection.Execute(strSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        mP_Connection.CommitTrans()
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Exit Sub
    End Sub
    Public Function ToSetcustdrgDesc(ByRef plngRow As Integer) As String
        Dim varItemCode As Object
        Dim rsCustDrgDesc As ClsResultSetDB
        Dim varCPartCode As Object
        rsCustDrgDesc = New ClsResultSetDB
        With ssPOEntry
            varCPartCode = Nothing
            Call .GetText(2, plngRow, varCPartCode)
            varItemCode = Nothing
            Call .GetText(4, plngRow, varItemCode)
            If Len(Trim(varCPartCode)) > 0 Then
                If Len(Trim(varItemCode)) Then
                    rsCustDrgDesc.GetResult("Select Drg_desc from CustItem_Mst" & _
                        " where unit_code='" & gstrUNITID & "'" & _
                        " and account_code =  '" & Trim(txtCustomerCode.Text) & "'" & _
                        " and ITem_code = '" & Trim(varItemCode) & "'" & _
                        " and Cust_drgNo = '" & Trim(varCPartCode) & "'" & _
                        " and active = 1 ") '10532789 active = 1 added
                    If rsCustDrgDesc.GetNoRows > 0 Then
                        rsCustDrgDesc.MoveFirst()
                        lblCustPartDesc.Text = rsCustDrgDesc.GetValue("Drg_desc")
                        ToSetcustdrgDesc = rsCustDrgDesc.GetValue("Drg_desc")
                    Else
                        lblCustPartDesc.Text = ""
                        ToSetcustdrgDesc = ""
                    End If
                    rsCustDrgDesc.ResultSetClose()
                End If
            End If
        End With
    End Function
    Private Sub txtECSSTaxType_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtECSSTaxType.TextChanged
        If Len(txtECSSTaxType.Text) = 0 Then
            'lblECSStax_Per.Caption = "0.00"
        End If
    End Sub
    Private Sub txtECSSTaxType_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtECSSTaxType.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.cmdButtons.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
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
            If CheckExistanceOfFieldData((txtECSSTaxType.Text), "TxRt_Rate_No", "Gen_TaxRate", " (Tx_TaxeID='ECS') AND unit_code='" & gstrUNITID & "'") Then
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
        Dim strHelp As String
        On Error GoTo ErrHandler
        Select Case Me.cmdButtons.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                If Len(txtECSSTaxType.Text) = 0 Then 'To check if There is No Text Then Show All Help
                    strHelp = ShowList(1, (txtECSSTaxType.MaxLength), "", "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND (Tx_TaxeID='ECS')")
                    If strHelp = "-1" Then 'If No Record Exists In The Table
                        Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                        Exit Sub
                    Else
                        txtECSSTaxType.Text = strHelp
                    End If
                Else
                    strHelp = ShowList(1, (txtECSSTaxType.MaxLength), txtECSSTaxType.Text, "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND (Tx_TaxeID='ECS')")
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
    Private Sub cmdButtons_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.UCbtndgrp.ButtonClickEventArgs) Handles cmdButtons.ButtonClick
        Dim varExtraExciseDuty, varSalestax, varToolCost, varordqty, varItemCode, varRate, varCustSuppMaterial, varPkg, varExiseDuty, varSurchargeSalesTax As Object
        Dim varDespatchQty, varCustItemCode, varCustItemDesc, varOthers As Object
        Dim varDeleteFlag As Object
        Dim intLoop As Short
        Dim intMaxLoop As Short
        Dim rsSalesParameter As ClsResultSetDB
        On Error GoTo ErrHandler
        Dim strSQL As String
        Dim strErrMsg As String
        Dim blnInvalidData As Boolean
        Dim intRow As Short
        Dim enmValue As eMPowerFunctions.ConfirmWindowReturnEnum
        Dim varItem As Object
        Dim vardrawing As Object
        blnInvalidData = False
        Dim varCustMat As Object
        Dim varEx As Object
        Dim counter As Short
        Dim varCustSuppMat As Object
        Dim valAddExcise As Object
        Dim varexternalsalesorder As Object

        Dim oDR As SqlDataReader

        Select Case e.Button
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD 'Add Record
                '                Call EnableControls(False, Me, True)
                txtCustomerCode.Enabled = True
                txtCustomerCode.BackColor = System.Drawing.Color.White
                cmdHelp(0).Enabled = True
                ssPOEntry.MaxRows = 0
                m_strSalesTaxType = ""
                txtCustomerCode.Focus()
                cmbPOType.Text = ObsoleteManagement.GetItemString(cmbPOType, 1)
                cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE) = True
                chkShipAddress.Enabled = True
                GrpSODoc.Visible = True
                GrpSODoc.Enabled = True
                lblFilePath.Text = ""
                lblFilePath.Tag = ""
                BtnUpload.Enabled = True
                btnViewDoc.Enabled = False
                btnRemoveDoc.Enabled = False
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT 'Edit Record
                ssPOEntry.Enabled = True
                If gblnGSTUnit = False Then
                    txtSTax.Enabled = True : txtSTax.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : cmdHelp(4).Enabled = True
                    cmdServiceTax.Enabled = True
                    txtService.Enabled = True
                    txtService.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    cmdSBCtax.Enabled = True
                    txtSBC.Enabled = True
                    txtSBC.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    cmdKKCtax.Enabled = True
                    txtKKC.Enabled = True
                    txtKKC.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    '1.Surcharge on S.Tax
                    txtSChSTax.Enabled = True : txtSChSTax.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : cmdHelp(6).Enabled = True
                    '****
                    txtECSSTaxType.Enabled = True : txtECSSTaxType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : CmdECSSTaxType.Enabled = True
                End If
                'Addition ends here
                If blnNoneditableCreditTerms_onSO Then
                    txtCreditTerms.Enabled = False : txtCreditTerms.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : cmdHelp(5).Enabled = False
                Else
                    txtCreditTerms.Enabled = True : txtCreditTerms.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : cmdHelp(5).Enabled = True
                End If
                chkOpenSo.Enabled = True
                txtCustomerCode.Enabled = False
                txtCustomerCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                cmdHelp(0).Enabled = False
                txtReferenceNo.Enabled = False
                txtReferenceNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                cmdHelp(1).Enabled = False
                txtAmendmentNo.Enabled = False
                txtAmendmentNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                cmdHelp(3).Enabled = False
                txtCurrencyType.Enabled = False
                cmdHelp(2).Enabled = False
                cmbPOType.Enabled = False
                cmbPOType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                txtAmendReason.Enabled = True
                cmdchangetype.Enabled = True
                If Len(Trim(txtAmendmentNo.Text)) = 0 Then
                    DTAmendmentDate.Enabled = False
                    txtAmendReason.Enabled = False
                    txtAmendReason.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    cmbPOType.Enabled = True
                    cmbPOType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    txtCurrencyType.Enabled = True
                    txtCurrencyType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    cmdHelp(2).Enabled = True
                    'ctlPerValue.Enabled = True
                    If blnNoneditableCreditTerms_onSO Then
                        txtCreditTerms.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : txtCreditTerms.Enabled = False : cmdHelp(5).Enabled = False
                    Else
                        txtCreditTerms.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : txtCreditTerms.Enabled = True : cmdHelp(5).Enabled = True
                    End If
                    ctlPerValue.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    'change account Plug in
                    With ssPOEntry
                        .Col = 1
                        .Col2 = .MaxCols
                        .BlockMode = True
                        .Lock = False
                        .BlockMode = False
                        intMaxLoop = .MaxRows
                        For intLoop = 1 To intMaxLoop
                            .Col = 2
                            .Col2 = 4
                            .Row = intLoop
                            .Row2 = intLoop
                            .BlockMode = True
                            .Lock = True
                            .BlockMode = False
                            rsSalesParameter = New ClsResultSetDB
                            rsSalesParameter.GetResult("select ItemRateLink from Sales_Parameter WHERE unit_code='" & gstrUNITID & "' ")
                            If rsSalesParameter.GetValue("ItemRateLink") = True Then
                                rsSalesParameter.ResultSetClose()
                                If checkforitemRate(intLoop) = False Then
                                    .Col = 6
                                    .Col2 = .MaxCols
                                    .Row = intLoop
                                    .Row2 = intLoop
                                    .BlockMode = True
                                    .Lock = True
                                    .BlockMode = False
                                Else
                                    .Col = 6
                                    .Col2 = .MaxCols
                                    .Row = intLoop
                                    .Row2 = intLoop
                                    .BlockMode = True
                                    .Lock = False
                                    .BlockMode = False
                                End If
                                .Col = 6
                                .Col2 = .MaxCols
                                .Row = intLoop
                                .Row2 = intLoop
                                .BlockMode = True
                                .Lock = False
                                .BlockMode = False
                            End If
                        Next
                    End With
                Else
                    DTAmendmentDate.Enabled = True
                    txtAmendReason.Enabled = True
                    txtAmendReason.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    cmbPOType.Enabled = False
                    cmbPOType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    txtCurrencyType.Enabled = False
                    txtCurrencyType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    cmdHelp(2).Enabled = False
                    'change account Plug in
                    With ssPOEntry
                        .Col = 1
                        .Col2 = .MaxCols
                        .BlockMode = True
                        .Lock = False
                        .BlockMode = False
                        .Col = 1
                        .Col2 = 1
                        .BlockMode = True
                        .Lock = False
                        .BlockMode = False
                    End With
                End If
                'DTDate.Enabled = True
                If blnSO_EDITABLE = True Then  '' Added by priti on 25 Jul 2025 
                    DTDate.Enabled = True
                Else
                    DTDate.Enabled = False
                End If
                DTEffectiveDate.Enabled = True
                DTValidDate.Enabled = True
                GrpSODoc.Visible = True
                BtnUpload.Enabled = True
                btnViewDoc.Enabled = True
                btnRemoveDoc.Enabled = True

                'If txtAmendmentNo.Text.Trim <> "" Then
                '    strSQL = "SELECT DOC_NAME,MSMEDOC FROM Cust_Ord_DOC_Dtl WHERE UNIT_CODE ='" & gstrUNITID & "' AND Account_Code ='" & txtCustomerCode.Text.Trim & "' AND Cust_Ref='" & txtReferenceNo.Text.Trim & "' AND ISACTIVE=1 AND Amendment_No='" & txtAmendmentNo.Text.Trim & "'"

                'Else
                '    strSQL = "SELECT DOC_NAME,MSMEDOC FROM Cust_Ord_DOC_Dtl WHERE UNIT_CODE ='" & gstrUNITID & "' AND Account_Code ='" & txtCustomerCode.Text.Trim & "' AND Cust_Ref='" & txtReferenceNo.Text.Trim & "' AND ISACTIVE=1"

                'End If

                'oDR = SqlConnectionclass.ExecuteReader(strSQL)
                'If oDR.HasRows = True Then
                '    oDR.Read()
                '    GrpSODoc.Visible = True
                '    BtnUpload.Enabled = True
                '    btnViewDoc.Enabled = True
                '    btnRemoveDoc.Enabled = True
                '    lblFilePath.Tag = oDR("DOC_NAME")
                'Else
                '    GrpSODoc.Visible = True
                '    BtnUpload.Enabled = True
                '    btnViewDoc.Enabled = False
                '    btnRemoveDoc.Enabled = False
                '    lblFilePath.Tag = ""
                '    lblFilePath.Text = ""
                'End If
                oDR = Nothing
                With ssPOEntry
                    .Row = 1
                    .Col = 2
                    .Action = 0
                End With
                LockGrid()

                '10736222-changes done for CT2 ARE 3 functionality
                ChkCT2Reqd.Enabled = False
                If ChkCT2Reqd.Checked = True Then
                    txtECSSTaxType.Text = ""
                    txtECSSTaxType.Enabled = False
                    CmdECSSTaxType.Enabled = False
                End If
                chkShipAddress.Enabled = True

                Exit Sub
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE 'Delete Record
                enmValue = ConfirmWindow(10054, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION)
                If enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                    ' deleting the record from cust_ord_hdr table
                    strSQL = "delete cust_ord_hdr where unit_code='" & gstrUNITID & "' and Account_Code='"
                    strSQL = strSQL & txtCustomerCode.Text & "' and  Consignee_Code='" & Trim(txtConsCode.Text) & "'  and Cust_Ref='"
                    strSQL = strSQL & txtReferenceNo.Text & "' and Amendment_No='"
                    strSQL = strSQL & txtAmendmentNo.Text & "'"
                    ' deleting the record from cust_ord_dtl table
                    Call DeleteRow()
                    Call DeleteDocumentPO()
                    mP_Connection.Execute(strSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    cmdButtons.Revert()
                    cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD) = True
                    cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                    cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                    cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                    Call EnableControls(False, Me, True)
                    txtCustomerCode.Enabled = True
                    txtCustomerCode.BackColor = System.Drawing.Color.White
                    cmdHelp(0).Enabled = True
                    ssPOEntry.MaxRows = 0
                    txtCustomerCode.Focus()
                Else
                    txtCustomerCode.Focus()
                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE 'Save Record
                If ValidRecord() = False Then Exit Sub
                If ChkCT2Reqd.Checked = False Then
                    If Len(txtECSSTaxType.Text) = 0 And gblnGSTUnit = False And gstrUNITID <> "STH" Then
                        If MsgBox("You have not entered Cess on ED.Do you want to continue", MsgBoxStyle.YesNo + MsgBoxStyle.Information, "Empower") = MsgBoxResult.No Then
                            txtECSSTaxType.Focus()
                            Exit Sub
                        End If
                    End If
                End If
                counter = ssPOEntry.MaxRows

                If counter = 0 Then
                    MessageBox.Show("Item Details Not Entered.", ResolveResString(100), MessageBoxButtons.OK)
                    Exit Sub
                End If

                For intRow = 1 To counter 'Checking if all details have been entered correctly
                    varCustSuppMat = Nothing
                    Call ssPOEntry.GetText(7, intRow, varCustSuppMat)
                    valAddExcise = Nothing
                    Call ssPOEntry.GetText(11, intRow, valAddExcise)
                    If Val(varCustSuppMat) > 0 And Len((Trim(valAddExcise))) = 0 And (gblnGSTUnit = False And gstrUNITID <> "STH") Then
                        MsgBox("Please Enter Additional Excise Duty.", MsgBoxStyle.Information, ResolveResString(100))
                        ssPOEntry.Col = 11
                        ssPOEntry.Row = ssPOEntry.ActiveRow
                        ssPOEntry.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                        gblnCancelUnload = True : gblnFormAddEdit = True
                        Exit Sub
                    End If
                    varCustMat = Nothing

                    Call ssPOEntry.GetText(7, intRow, varCustMat)
                    varEx = Nothing
                    Call ssPOEntry.GetText(11, intRow, varEx)
                    If varCustMat > 0 And IsNothing(Trim(varEx)) Then
                        MsgBox("Please Enter Additional Excise Duty,")
                        ssPOEntry.Col = 11
                        ssPOEntry.Row = intRow
                        ssPOEntry.Focus()
                        ssPOEntry.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                        Exit Sub
                    End If

                    If DataExist("SELECT SOUPLD_LINE_LEVEL_SALESORDER FROM CUSTOMER_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & Trim(txtCustomerCode.Text) & "'and SOUPLD_LINE_LEVEL_SALESORDER =1") = True Then
                        varexternalsalesorder = Nothing
                        Call ssPOEntry.GetText(19, intRow, varexternalsalesorder)
                        If Len(Trim(varexternalsalesorder)) <= 0 Then
                            MsgBox("Please Enter External Sales order ")
                            ssPOEntry.Col = 19
                            ssPOEntry.Row = intRow
                            ssPOEntry.Focus()
                            ssPOEntry.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                            Exit Sub
                        End If

                    End If
                    If ssPOEntry.MaxRows = 1 Or intRow <> ssPOEntry.MaxRows Then
                        If Not ValidRowData(intRow, 0) Then
                            gblnCancelUnload = True : gblnFormAddEdit = True
                            Exit Sub
                        End If
                    Else
                        'change account Plug in
                        varItem = Nothing
                        Call ssPOEntry.GetText(2, ssPOEntry.MaxRows, varItem)
                        vardrawing = Nothing
                        Call ssPOEntry.GetText(2, ssPOEntry.MaxRows, vardrawing)
                        If ((Len(Trim(varItem)) <= 0) And (Len(Trim(vardrawing)) <= 0)) And ssPOEntry.MaxRows > 1 Then
                            ssPOEntry.MaxRows = ssPOEntry.MaxRows - 1
                        Else
                            If Not ValidRowData(intRow, 0) Then
                                gblnCancelUnload = True : gblnFormAddEdit = True
                                Exit Sub
                            Else
                                With ssPOEntry
                                    .Row = .MaxRows
                                    .Col = 8
                                    If Val(.Text) = 0 Then
                                        rsSalesParameter = New ClsResultSetDB
                                        rsSalesParameter.GetResult("Select ToolCostMsg from Sales_Parameter WHERE unit_code='" & gstrUNITID & "'")
                                        If rsSalesParameter.GetValue("ToolCostMsg") = True Then
                                            rsSalesParameter.ResultSetClose()
                                            If (MsgBox("You have Entered 0 Tool Cost,Save Data ?", MsgBoxStyle.YesNo, "empower")) = MsgBoxResult.No Then
                                                .Row = .MaxRows
                                                .Col = 8
                                                .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                                .Focus() : Exit Sub
                                            End If
                                        End If
                                    End If
                                End With
                            End If
                        End If
                    End If
                Next
                Select Case cmdButtons.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD 'Check for mode when save button was clicked
                        ReDim ArrDispatchQty(ssPOEntry.MaxRows - 1)
                        If ssPOEntry.MaxRows = 0 Then
                            MsgBox("Please enter Item Details", MsgBoxStyle.Information, ResolveResString(100))
                            Exit Sub
                        End If
                        With mP_Connection
                            .BeginTrans()
                            Call InsertRowCustOrdHdr()
                            Call InsertRow()
                            .CommitTrans()
                            Call UPLoadPODocument()
                            Call SendMail()

                            rsSalesParameter = New ClsResultSetDB
                            rsSalesParameter.GetResult("Select AppendSOItem from sales_parameter WHERE unit_code='" & gstrUNITID & "' ")
                            If rsSalesParameter.GetValue("AppendSOItem") = True Then
                                'ISSUE ID : 10763705
                                If mblnappendsoitem_customer = True Then
                                    'ISSUE ID : 10763705
                                    Call InsertPreviousSODetails(Trim(txtCustomerCode.Text), Trim(txtReferenceNo.Text), Trim(txtAmendmentNo.Text), Trim(lblIntSONoDes.Text), CShort(Trim(lblRevisionNo.Text)))
                                End If
                            End If

                            GrpSODoc.Visible = True
                            GrpSODoc.Enabled = False
                            lblFilePath.Text = ""
                            lblFilePath.Tag = ""
                            BtnUpload.Enabled = False
                            btnViewDoc.Enabled = False
                            btnRemoveDoc.Enabled = False
                            rsSalesParameter.ResultSetClose()
                        End With
                        cmdButtons.Revert()
                        cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD) = True
                        cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                        cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                        cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                        Call EnableControls(False, Me, False)
                        txtCustomerCode.Enabled = True
                        txtCustomerCode.BackColor = System.Drawing.Color.White
                        cmdHelp(0).Enabled = True
                        ssPOEntry.MaxRows = 0
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT ' in case of edit, update the record in the cust_ord_hdr table
                        With mP_Connection
                            'To Confirm the deletion of marked rows
                            If ConfirmDeletion() = False Then Exit Sub
                            .BeginTrans()
                            Call UpdateRow()
                            'delete the record in the cust_ord_dtl table
                            Call DeleteRow()
                            ' Add all the records in the grid to the table cust_ord_dtl which are not marked for deletion
                            Call InsertRow()
                            .CommitTrans()
                            Call UPLoadPODocument()
                            Call SendMail()
                        End With
                        cmdButtons.Revert()
                        cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD) = True
                        cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                        cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                        cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                        Call EnableControls(False, Me, False)
                        txtCustomerCode.Enabled = True
                        txtCustomerCode.BackColor = System.Drawing.Color.White
                        cmdHelp(0).Enabled = True
                        ssPOEntry.MaxRows = 0
                        GrpSODoc.Visible = True
                        GrpSODoc.Enabled = False
                        lblFilePath.Text = ""
                        lblFilePath.Tag = ""
                        btnRemoveDoc.Enabled = False
                        BtnUpload.Enabled = False
                        btnViewDoc.Enabled = False
                End Select
                If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                    Call ConfirmWindow(10049, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                Else
                    If Len(Trim(txtAmendmentNo.Text)) = 0 Then
                        MsgBox("SO Successfully update with Internal SO No " & lblIntSONoDes.Text, MsgBoxStyle.Information, "empower")
                    Else
                        Call ConfirmWindow(10049, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                    End If
                End If
                gblnCancelUnload = False : gblnFormAddEdit = False
                If lblConsDesc.Text <> "" Then lblConsDesc.Text = ""
                Call EnableControls(False, Me, True)
                txtCustomerCode.Enabled = True
                txtCustomerCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                cmdHelp(0).Enabled = True
                txtCustomerCode.Focus()
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL
                If ConfirmWindow(10053, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION, 60095) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                    Call EnableControls(False, Me, True)
                    txtCustomerCode.Enabled = True
                    txtCustomerCode.BackColor = System.Drawing.Color.White
                    cmdHelp(0).Enabled = True
                    ssPOEntry.MaxRows = 0
                    txtCustomerCode.Focus()
                    gblnCancelUnload = False : gblnFormAddEdit = False
                    cmdButtons.Revert()
                    cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD) = True
                    cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                    cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                    cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                    Call RefreshForm()
                    If cmdButtons.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                    End If
                Else
                    Exit Sub
                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT
                Call PrintToReport()
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                Me.Close()
        End Select
        Call CLEARVAR()
        If Not oDR Is Nothing AndAlso oDR.IsClosed = False Then
            oDR.Close()
        End If
        oDR = Nothing
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub cmdButtons_MouseDown(ByVal Sender As Object, ByVal e As UCActXCtl.UCbtndgrp.MouseDownEventArgs) Handles cmdButtons.MouseDown
        On Error GoTo ErrHandler
        Select Case e.Index
            Case 0
                m_blnCloseFlag = True
            Case 4
                m_blnCloseFlag = True
            Case 6
                m_blnCloseFlag = True
        End Select
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub DTDate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DTDate.KeyDown
        On Error GoTo ErrHandler
        If e.KeyCode = System.Windows.Forms.Keys.Return Then
            If DTAmendmentDate.Enabled = True Then
                DTAmendmentDate.Focus()
            Else
                If DTEffectiveDate.Enabled Then
                    DTEffectiveDate.Focus()
                Else
                    cmdButtons.Focus()
                End If
            End If
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub DTDate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DTDate.KeyPress
        On Error GoTo ErrHandler
        Dim KeyAscii As Short = Asc(e.KeyChar)
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub DTDate_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles DTDate.Validating
        On Error GoTo ErrHandler
        ''INC1411849
        'm_strSql = "select Company_Code,Phone,Fax,Email,Financial_StartDate,Financial_EndDate,Reg_No,ECC_No,PLA_No,PF_No,IEC_No,CST_No,ESI_No,LST_No,Bank_Name1,Bank_Name2,Bank_Name3,Excise_1,Excise_2,Excise_3,Range_1,Range_2,Division,Commissionerate,mrpmonth,Excise_Ac,Add_Excise_Ac,Salestax_Ac,Add_Salestax_Ac,TDS_Ac,Insurance_Ac,Surcharge_Ac,OtherCharges_Ac,Purchasest_Ac,Rec_lock,Shifts,Exporter_code,EOU_flag,Invoice_Rule,SampleExp_Ac,Logo,Base_Currency,Tin_No,activity_code,EnvRemarks,RegisteredOfficeAddress,EnvRemarks2,PANNO from company_mst WHERE unit_code='" & gstrUNITID & "' "
        'rsdb = New ClsResultSetDB
        'rsdb.GetResult(m_strSql)
        'rsdb.ResultSetClose()
        ''INC1411849
        '******
        If DTDate.Value > DateValue(CStr(FinancialYearDates(eMPowerFunctions.FinancialYearDatesEnum.DATE_END))) Then
            Call ConfirmWindow(10074, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
            e.Cancel = True
            Exit Sub
        End If
        If DTDate.Value > GetServerDate() Then
            MsgBox("Date Can not be greater then Current Date")
            e.Cancel = True
            Exit Sub
        End If
        dtSODate = DTDate.Value
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub DTDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DTDate.ValueChanged
        Dim rsSalesParameter As ClsResultSetDB
        rsSalesParameter = New ClsResultSetDB
        rsSalesParameter.GetResult("Select ItemRateLink from Sales_Parameter WHERE unit_code='" & gstrUNITID & "'")
        If rsSalesParameter.GetValue("ItemRateLink") = True Then
            If DateDiff(Microsoft.VisualBasic.DateInterval.Day, dtSODate, DTDate.Value) <> 0 Then
                If MsgBox("Change in SO Date will remove the all ITem Details from Grid.", MsgBoxStyle.YesNo, "Empower") = MsgBoxResult.Yes Then
                    ssPOEntry.MaxRows = 0
                    Call ADDRow()
                    DTDate.Focus()
                Else
                    DTDate.Value = GetServerDate() : DTDate.Focus()
                End If
            End If
        End If
    End Sub
    Private Sub DTValidDate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DTValidDate.KeyDown
        On Error GoTo ErrHandler
        If e.KeyCode = System.Windows.Forms.Keys.Return Then
            If txtCurrencyType.Enabled = True Then
                txtCurrencyType.Focus()
            Else
                txtAmendReason.Focus()
            End If
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub DTValidDate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DTValidDate.KeyPress
        On Error GoTo ErrHandler
        Dim KeyAscii As Short = Asc(e.KeyChar)
        Select Case KeyAscii
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub DTValidDate_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles DTValidDate.LostFocus
        Call DTValidDate_Validating(DTValidDate, New System.ComponentModel.CancelEventArgs(False))
    End Sub
    Private Sub DTValidDate_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles DTValidDate.Validating
        Dim Cancel As Boolean = e.Cancel
        On Error GoTo ErrHandler
        Dim rsGetDate As New ClsResultSetDB
        m_strSql = "Select * from company_mst WHERE unit_code='" & gstrUNITID & "' "
        rsGetDate.GetResult(m_strSql)
        rsGetDate.ResultSetClose()
        If DTValidDate.Value < DateValue(CStr(FinancialYearDates(eMPowerFunctions.FinancialYearDatesEnum.DATE_START))) Then
            Call ConfirmWindow(10073, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
            Cancel = True
            DTValidDate.Value = FinancialYearDates(eMPowerFunctions.FinancialYearDatesEnum.DATE_START)
            DTValidDate.Focus()
            GoTo EventExitSub
        End If
        If DTValidDate.Value > DateValue(CStr(FinancialYearDates(eMPowerFunctions.FinancialYearDatesEnum.DATE_END))) Then
            Call ConfirmWindow(10074, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
            Cancel = True
            DTValidDate.Value = FinancialYearDates(eMPowerFunctions.FinancialYearDatesEnum.DATE_END)
            DTValidDate.Focus()
            GoTo EventExitSub
        End If
        If DTValidDate.Value < GetServerDate() Then
            MsgBox("Valid Date Cannot be Less then Current Date.", MsgBoxStyle.OkOnly, "empower")
            Cancel = True
            DTValidDate.Value = GetServerDate()
            DTValidDate.Focus()
            GoTo EventExitSub
        End If
        GoTo EventExitSub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        GoTo EventExitSub
EventExitSub:
        e.Cancel = Cancel
    End Sub
    Private Sub DTAmendmentDate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DTAmendmentDate.KeyDown
        On Error GoTo ErrHandler
        If e.KeyCode = System.Windows.Forms.Keys.Return Then
            Call DTAmendmentDate_Validating(DTAmendmentDate, New System.ComponentModel.CancelEventArgs(False))
            If blnValidAmendDate = True Then
                DTEffectiveDate.Focus()
            End If
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub DTAmendmentDate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DTAmendmentDate.KeyPress
        On Error GoTo ErrHandler
        Dim KeyAscii As Short = Asc(e.KeyChar)
        Select Case KeyAscii
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub DTAmendmentDate_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles DTAmendmentDate.Validating
        Dim Cancel As Boolean = e.Cancel
        On Error GoTo ErrHandler
        System.Windows.Forms.Application.DoEvents()
        If DTAmendmentDate.Value > DTValidDate.Value Then
            Call ConfirmWindow(10146, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
            blnValidAmendDate = False
            Cancel = True
            DTAmendmentDate.Focus()
            GoTo EventExitSub
        Else
            blnValidAmendDate = True
        End If
        If DTAmendmentDate.Value < DTDate.Value Then
            Call ConfirmWindow(10147, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
            blnValidAmendDate = False
            Cancel = True
            DTAmendmentDate.Focus()
            GoTo EventExitSub
        Else
            blnValidAmendDate = True
        End If
        If DTAmendmentDate.Value > GetServerDate() Then
            MsgBox("Date Can not be Greater Then Current Date", MsgBoxStyle.Information, "empower")
            blnValidAmendDate = False
            Cancel = True
            DTAmendmentDate.Focus()
            GoTo EventExitSub
        Else
            blnValidAmendDate = True
        End If
        GoTo EventExitSub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        GoTo EventExitSub
EventExitSub:
        e.Cancel = Cancel
    End Sub
    Private Sub DTEffectiveDate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DTEffectiveDate.KeyDown
        On Error GoTo ErrHandler
        If e.KeyCode = System.Windows.Forms.Keys.Return Then
            If DTValidDate.Enabled = True Then
                DTValidDate.Focus()
            Else
                txtAmendReason.Focus()
            End If
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub DTEffectiveDate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DTEffectiveDate.KeyPress
        On Error GoTo ErrHandler
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If KeyAscii = 39 Or KeyAscii = 34 Or KeyAscii = 96 Then KeyAscii = 0
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub DTEffectiveDate_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles DTEffectiveDate.LostFocus
        On Error GoTo ErrHandler
        Dim rsGetDate As New ClsResultSetDB
        m_strSql = "Select * from company_mst WHERE unit_code='" & gstrUNITID & "' "
        rsGetDate.GetResult(m_strSql)
        rsGetDate.ResultSetClose()
        If DTEffectiveDate.Value < DTDate.Value Then
            MsgBox("Effective Date Cannot Be Less than SO Date", MsgBoxStyle.Information, "empower")
            DTEffectiveDate.Value = DTDate.Value
            DTEffectiveDate.Focus()
            Exit Sub
        End If
        If DTEffectiveDate.Value > DateValue(CStr(FinancialYearDates(eMPowerFunctions.FinancialYearDatesEnum.DATE_END))) Then
            MsgBox("Effective Date Cannot Be Greater than Financial End Date", MsgBoxStyle.Information, "empower")
            DTEffectiveDate.Value = DTDate.Value
            DTEffectiveDate.Focus()
            Exit Sub
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub ctlPerValue_Change(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlPerValue.Change
        Dim intLoopCounter As Short
        Dim rsSalesParameter As ClsResultSetDB
        Dim varItemCode As Object
        Dim varDrgNo As Object
        Dim varRate As Object
        Dim varCustSupp As Object
        Dim varToolCost As Object
        Dim varOthers As Object
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
                .Col = 7 : .Text = "Cust Supp Mat. (Per Unit)"
                .Row = 0
                .Col = 8 : .Text = "Tool Cost (Per Unit)"
                .Row = 0
                .Col = 11 : .Text = "Addition Excise Duty"
                .Row = 0
                .Col = 12 : .Text = "Others (Per Unit)"
            End If
            rsSalesParameter = New ClsResultSetDB
            rsSalesParameter.GetResult("Select ItemRateLink from Sales_Parameter WHERE unit_code='" & gstrUNITID & "' ")
            If rsSalesParameter.GetValue("ItemRateLink") = True Then
                If (Len(Trim(txtAmendmentNo.Text)) = 0) And (txtAmendmentNo.Enabled = False) Then
                    For intLoopCounter = 1 To ssPOEntry.MaxRows
                        varDrgNo = Nothing
                        Call .GetText(2, intLoopCounter, varDrgNo)
                        varItemCode = Nothing
                        Call .GetText(4, intLoopCounter, varItemCode)
                        If (Len(Trim(CStr(varDrgNo))) > 0) And (Len(Trim(CStr(varItemCode))) > 0) Then
                            varRate = Nothing
                            Call .GetText(13, intLoopCounter, varRate)
                            varCustSupp = Nothing
                            Call .GetText(14, intLoopCounter, varCustSupp)
                            varToolCost = Nothing
                            Call .GetText(15, intLoopCounter, varToolCost)
                            varOthers = Nothing
                            Call .GetText(16, intLoopCounter, varOthers)
                            Call .SetText(6, intLoopCounter, varRate * CDbl(ctlPerValue.Text))
                            Call .SetText(7, intLoopCounter, varCustSupp * CDbl(ctlPerValue.Text))
                            Call .SetText(8, intLoopCounter, varToolCost * CDbl(ctlPerValue.Text))
                            Call .SetText(11, intLoopCounter, varOthers * CDbl(ctlPerValue.Text))
                        End If
                    Next
                End If
            End If
            rsSalesParameter.ResultSetClose()
        End With
    End Sub
    Private Sub ctlPerValue_KeyPress(ByVal Sender As Object, ByVal e As UCActXCtl.UCctlFloat.KeyPressEventArgs) Handles ctlPerValue.KeyPress
        Select Case e.KeyAscii
            Case System.Windows.Forms.Keys.Return
                If cmbPOType.SelectedIndex = 4 Then
                    cmdButtons.Focus()
                Else
                    cmdchangetype.Focus()
                End If
        End Select
    End Sub
    Private Sub ssPOEntry_Advance(ByVal sender As System.Object, ByVal e As AxFPSpreadADO._DSpreadEvents_AdvanceEvent) Handles ssPOEntry.Advance
        On Error GoTo ErrHandler
        If cmdButtons.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            Call ADDRow()
        End If
        With ssPOEntry
            .Col = 1
            .Row = .MaxRows : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : If .Enabled Then .Focus()
        End With
        Exit Sub            'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub ssPOEntry_ButtonClicked(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ButtonClickedEvent) Handles ssPOEntry.ButtonClicked
        On Error GoTo ErrHandler
        Dim rsDesc As New ClsResultSetDB
        Dim rsSalesParametere As ClsResultSetDB
        Dim varHelpItem As Object
        Dim strSOEntry() As String
        Dim strtest As String
        Dim rsitem As ClsResultSetDB
        Dim strReturnCustRef As String
        If Len(Trim(txtCurrencyType.Text)) = 0 Then Exit Sub 'for currency type check
        rsSalesParametere = New ClsResultSetDB
        Dim strCT2Condition As String = ""
        Dim strDockCodeCondition As String = ""
        Dim strdockcode As String = ""
        Dim strCustitemcode As String = ""
        Dim stritemcode As String = ""
        Dim StrServiceCond As String = ""
        'GST 
        Dim strGSTtaxdetails() As String
        Dim objSQLConn As SqlConnection
        Dim objReader As SqlDataReader
        Dim objCommand As SqlCommand
        Dim STRSQL As String
        'GST 

        '10797956 
        '10736222-changes done for CT2 ARE 3 functionality - to enable CT2 Reqd flag
        If ChkCT2Reqd.Checked = True Then
            strCT2Condition = " And Exists(select * From CT2_Cust_Item_Linkage Z where A.UNIT_CODE=z.Unit_Code and A.Account_Code=z.Customer_Code and A.Item_code =z.Item_Code and A.Cust_Drgno=z.Cust_drgno And z.Active=1 and z.isAuthorized=1)"
            With ssPOEntry
                .Col = 10
                .Text = "EX0"
                .Lock = True
                .Col = 11
                .Text = "AED0"
                .Lock = True
            End With
        End If
        If Mid(cmbPOType.Text, 1, 1) = "V" Then
            StrServiceCond = " AND ITEM_MAIN_GRP ='M' "
        End If
        rsSalesParametere.GetResult("Select ItemRateLink from Sales_parameter WHERE unit_code='" & gstrUNITID & "' ")
        'change account Plug in
        If e.col = 3 Then
            If txtAmendmentNo.Enabled = False Then
                If rsSalesParametere.GetValue("ItemRateLink") = True Then
                    '10532789 active = 1 added
                    'begin
                    '10869290
                    If Mid(cmbPOType.Text, 1, 1) = "V" Then
                        strSOEntry = ctlEMPHelpSOEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select a.Cust_drgNo,a.ITem_code,a.drg_Desc" &
                            " from CustItem_Mst a,Itemrate_Mst b,Item_mst C where A.UNIT_CODE=B.UNIT_CODE AND" &
                            " a.unit_code='" & gstrUNITID & "' and a.Account_Code='" & txtCustomerCode.Text & "'" &
                            " and a.Account_code = b.Party_Code and a.cust_DrgNo = b.Item_code" &
                            " and a.active = 1 and a.Unit_Code=C.unit_Code and a.Item_Code=C.item_Code and C.Status='A' and C.Hold_Flag=0 and C.Item_Main_Grp='M'" &
                            " and b.serial_no=(select max(serial_no) from itemrate_mst i1" &
                            " Where i1.unit_code=b.unit_code and i1.party_code = b.party_code" &
                            " and i1.item_code=b.item_code)" &
                            " and datediff(mm,'" & getDateForDB(DTDate.Value) & "',b.DateFrom)<=0" &
                            " and CustVend_Flg = 'C'" & IIf(StrServiceCond.Trim.Length = 0, "", StrServiceCond) & IIf(strCT2Condition.Trim.Length = 0, "", strCT2Condition))
                    Else
                        strtest = "Select a.Cust_drgNo,a.ITem_code,a.drg_Desc from CustItem_Mst a,Itemrate_Mst b" &
                        " where A.UNIT_CODE=B.UNIT_CODE AND a.Account_Code='" & txtCustomerCode.Text & "'" &
                        " and a.Account_code = b.Party_Code and a.cust_DrgNo = b.Item_code" &
                        " and a.active = 1 and b.serial_no=(select max(serial_no) from itemrate_mst i1" &
                        " Where i1.unit_code=b.unit_code and i1.party_code = b.party_code" &
                        " and i1.item_code=b.item_code)" &
                        " and datediff(mm,convert(datetime,'" & getDateForDB(DTDate.Value) & "'),convert(datetime,b.DateFrom))<=0" &
                        " and CustVend_Flg = 'C' and a.unit_code='" & gstrUNITID & "' "
                        'end

                        '10532789 active = 1 added
                        'begin
                        strSOEntry = ctlEMPHelpSOEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select a.Cust_drgNo,a.ITem_code,a.drg_Desc" &
                            " from CustItem_Mst a,Itemrate_Mst b where A.UNIT_CODE=B.UNIT_CODE AND" &
                            " a.unit_code='" & gstrUNITID & "' and a.Account_Code='" & txtCustomerCode.Text & "'" &
                            " and a.Account_code = b.Party_Code and a.cust_DrgNo = b.Item_code" &
                            " and a.active = 1 and b.serial_no=(select max(serial_no) from itemrate_mst i1" &
                            " Where i1.unit_code=b.unit_code and i1.party_code = b.party_code" &
                            " and i1.item_code=b.item_code)" &
                            " and datediff(mm,'" & getDateForDB(DTDate.Value) & "',b.DateFrom)<=0" &
                            " and CustVend_Flg = 'C'" & IIf(StrServiceCond.Trim.Length = 0, "", StrServiceCond) & IIf(strCT2Condition.Trim.Length = 0, "", strCT2Condition))
                        'end
                    End If
                Else
                    If Mid(cmbPOType.Text, 1, 1) = "V" Then
                        STRSQL = "select dbo.UDF_ISDOCKCODE_ASN( '" & gstrUNITID & "','" & txtCustomerCode.Text.Trim & "')"
                        If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(STRSQL)) = True Then

                            If ssPOEntry.MaxRows = 1 Then 'FUNCTIONALITY IS ON 
                                strSOEntry = ctlEMPHelpSOEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select Cust_drgNo,a.ITem_code,drg_Desc from CustItem_Mst a,Item_mst b" &
                                            " where a.Unit_Code=b.Unit_Code and a.item_Code=b.item_Code and b.status='A' and b.Hold_Flag=0 and b.Item_Main_Grp='M' and a.unit_code='" & gstrUNITID & "'" &
                                            " and Account_Code='" & txtCustomerCode.Text & "'" &
                                            " and a.active = 1" & IIf(strCT2Condition.Trim.Length = 0, "", strCT2Condition))
                            Else
                                With ssPOEntry
                                    .Row = 1
                                    .Col = 2
                                    strCustitemcode = .Text
                                    .Col = 4
                                    stritemcode = .Text

                                End With
                                STRSQL = "SELECT DOCKCODE FROM CUSTITEM_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND ACCOUNT_CODE='" & txtCustomerCode.Text & "'" &
                                         " AND ITEM_CODE='" & stritemcode & "' AND CUST_DRGNO='" & strCustitemcode & "' AND ACTIVE=1 "
                                strdockcode = Convert.ToString(SqlConnectionclass.ExecuteScalar(STRSQL))
                                strDockCodeCondition = "and Dockcode in ( '" & strdockcode & "' )"
                                strSOEntry = ctlEMPHelpSOEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select Cust_drgNo,a.ITem_code,drg_Desc from CustItem_Mst a,Item_mst b" &
                                            " where a.Unit_Code=b.Unit_Code and a.item_Code=b.item_Code and b.status='A' and b.Hold_Flag=0 and b.Item_Main_Grp='M' and a.unit_code='" & gstrUNITID & "'" &
                                            " and Account_Code='" & txtCustomerCode.Text & "'" &
                                            " and a.active = 1" & IIf(strCT2Condition.Trim.Length = 0, "", strCT2Condition) &
                                                    IIf(strDockCodeCondition.Trim.Length = 0, "", strDockCodeCondition))
                            End If
                        Else 'FUNCTIONALITY OF DOCKCODE IS OFF 
                            strSOEntry = ctlEMPHelpSOEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select Cust_drgNo,a.ITem_code,drg_Desc from CustItem_Mst a,Item_mst b" &
                                            " where a.Unit_Code=b.Unit_Code and a.item_Code=b.item_Code and b.status='A' and b.Hold_Flag=0 and b.Item_Main_Grp='M' and a.unit_code='" & gstrUNITID & "'" &
                                            " and Account_Code='" & txtCustomerCode.Text & "'" &
                                            " and a.active = 1" & IIf(strCT2Condition.Trim.Length = 0, "", strCT2Condition)) '10532789 active = 1 added
                        End If
                    Else
                        '10808160--Ends
                        STRSQL = "select dbo.UDF_ISDOCKCODE_ASN( '" & gstrUNITID & "','" & txtCustomerCode.Text.Trim & "')"
                        If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(STRSQL)) = True Then

                            If ssPOEntry.MaxRows = 1 Then 'FUNCTIONALITY IS ON 
                                strSOEntry = ctlEMPHelpSOEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select Cust_drgNo,ITem_code,drg_Desc from CustItem_Mst a" &
                                            " where unit_code='" & gstrUNITID & "'" &
                                            " and Account_Code='" & txtCustomerCode.Text & "'" &
                                            " and active = 1" & IIf(strCT2Condition.Trim.Length = 0, "", strCT2Condition))
                            Else
                                With ssPOEntry
                                    .Row = 1
                                    .Col = 2
                                    strCustitemcode = .Text
                                    .Col = 4
                                    stritemcode = .Text

                                End With
                                STRSQL = "SELECT DOCKCODE FROM CUSTITEM_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND ACCOUNT_CODE='" & txtCustomerCode.Text & "'" &
                                         " AND ITEM_CODE='" & stritemcode & "' AND CUST_DRGNO='" & strCustitemcode & "' AND ACTIVE=1 "
                                strdockcode = Convert.ToString(SqlConnectionclass.ExecuteScalar(STRSQL))
                                strDockCodeCondition = "and Dockcode in ( '" & strdockcode & "' )"
                                strSOEntry = ctlEMPHelpSOEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select Cust_drgNo,ITem_code,drg_Desc from CustItem_Mst a" &
                                            " where unit_code='" & gstrUNITID & "'" &
                                            " and Account_Code='" & txtCustomerCode.Text & "'" &
                                            " and active = 1" & IIf(strCT2Condition.Trim.Length = 0, "", strCT2Condition) &
                                                    IIf(strDockCodeCondition.Trim.Length = 0, "", strDockCodeCondition))
                            End If
                        Else 'FUNCTIONALITY OF DOCKCODE IS OFF 
                            strSOEntry = ctlEMPHelpSOEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select Cust_drgNo,ITem_code,drg_Desc from CustItem_Mst a" &
                                            " where unit_code='" & gstrUNITID & "'" &
                                            " and Account_Code='" & txtCustomerCode.Text & "'" &
                                            " and active = 1" & IIf(strCT2Condition.Trim.Length = 0, "", strCT2Condition)) '10532789 active = 1 added
                        End If
                    End If
                End If
            End If
            '10808160--Ends


            If UBound(strSOEntry) < 0 Then
                rsSalesParametere.ResultSetClose()
                Exit Sub
            End If
            If strSOEntry(0) = "0" Then
                Call ConfirmWindow(10013, ConfirmWindowButtonsEnum.BUTTON_OK, ConfirmWindowImagesEnum.IMG_INFO)
            Else
                strReturnCustRef = CheckForMultipleOpenSO(txtCustomerCode.Text, txtConsCode.Text, txtReferenceNo.Text, strSOEntry(0), strSOEntry(1))
                If Len(strReturnCustRef) > 0 Then
                    MsgBox("More than one Sales Order cannot be active for Customer Item Combination" & strReturnCustRef, vbInformation, ResolveResString(100))
                    Exit Sub
                End If
                Call ssPOEntry.SetText(2, ssPOEntry.ActiveRow, strSOEntry(0))
                Call ssPOEntry.SetText(4, ssPOEntry.ActiveRow, strSOEntry(1))
                lblCustPartDesc.Text = strSOEntry(2)
                'GST CHANGE
                If gblnGSTUnit = True Then
                    ssPOEntry.Row = 1 : ssPOEntry.Row2 = ssPOEntry.MaxRows : ssPOEntry.Col = 16 : ssPOEntry.Col2 = 25 : ssPOEntry.BlockMode = True : ssPOEntry.Lock = True : ssPOEntry.BlockMode = False
                End If
                STRSQL = "set dateformat 'dmy' SELECT * FROM DBO.UFN_GST_ITEMWISETAXES('" & gstrUNITID & "','" & txtCustomerCode.Text.Trim & "','" & strSOEntry(1) & "','" & DTEffectiveDate.Value & "','" & DTValidDate.Value & "')"
                objSQLConn = SqlConnectionclass.GetConnection()
                objCommand = New SqlCommand(STRSQL, objSQLConn)
                objReader = objCommand.ExecuteReader()
                If objReader.HasRows = True Then
                    objReader.Read()
                    Call ssPOEntry.SetText(20, ssPOEntry.ActiveRow, objReader.GetValue(1))
                    Call ssPOEntry.SetText(21, ssPOEntry.ActiveRow, objReader.GetValue(2))
                    Call ssPOEntry.SetText(22, ssPOEntry.ActiveRow, objReader.GetValue(3))
                    Call ssPOEntry.SetText(23, ssPOEntry.ActiveRow, objReader.GetValue(4))
                    Call ssPOEntry.SetText(24, ssPOEntry.ActiveRow, objReader.GetValue(5))
                    Call ssPOEntry.SetText(25, ssPOEntry.ActiveRow, objReader.GetValue(6))
                    Call ssPOEntry.SetText(26, ssPOEntry.ActiveRow, objReader.GetValue(0))
                    'GST CHANGE
                End If

                objReader = Nothing
                objSQLConn.Close()
                objSQLConn = Nothing
                'GST CHANGE
                m_strSql = "SELECT SUM(RATE) AS RATE FROM CUSTSUPPLIEDITEM_MST " &
                            " WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE = '" & txtCustomerCode.Text & "'" &
                            " AND FINISH_ITEM_CODE = '" & strSOEntry(1) & "'" &
                            " AND ACTIVE_FLAG = 1" &
                            " AND GETDATE() BETWEEN VALID_FROM AND VALID_TO"
                rsitem = New ClsResultSetDB
                rsitem.GetResult(m_strSql)
                If rsitem.GetNoRows > 0 Then
                    Call ssPOEntry.SetText(7, ssPOEntry.ActiveRow, IIf(rsitem.GetValue("RATE") = "", 0, rsitem.GetValue("RATE")))
                End If
                rsitem.ResultSetClose()
                rsitem = Nothing
            End If
            If DataExist("SELECT SOUPLD_LINE_LEVEL_SALESORDER FROM CUSTOMER_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & Trim(txtCustomerCode.Text) & "' and SOUPLD_LINE_LEVEL_SALESORDER =1 ") = True Then
                m_strSql = "SELECT CD.EXTERNAL_SALESORDER_NO  from CUST_ORD_HDR CH, Cust_Ord_Dtl CD where " &
                            "CH.UNIT_CODE = CD.UNIT_CODE AND CH.ACCOUNT_CODE =CD.ACCOUNT_CODE And CH.Cust_Ref = CD.Cust_Ref AND CH.Amendment_No =CD.Amendment_No " &
                            " AND CH.Active_Flag ='A' AND CH.Authorized_Flag =1 " &
                            " AND CH.UNIT_CODE='" & gstrUNITID & "' AND CH.ACCOUNT_CODE='" & Me.txtCustomerCode.Text.Trim & "' AND " &
                            " CH.CUST_REF='" & txtReferenceNo.Text & "'" &
                            " AND ITEM_CODE = '" & strSOEntry(1) & "'" &
                            " AND CUST_DRGNO= '" & strSOEntry(0) & "'"

                rsitem = New ClsResultSetDB
                rsitem.GetResult(m_strSql)
                If rsitem.GetNoRows > 0 Then
                    Call ssPOEntry.SetText(19, ssPOEntry.ActiveRow, IIf(rsitem.GetValue("EXTERNAL_SALESORDER_NO") = "", "", rsitem.GetValue("EXTERNAL_SALESORDER_NO")))
                End If
                rsitem.ResultSetClose()
                rsitem = Nothing

            End If

            If cmdButtons.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                If IsNothing(rsitem) = False Then rsitem.ResultSetClose()
                rsitem = New ClsResultSetDB
                If Len(Trim(m_Item_Code)) > 0 Then
                    m_strSql = "SElect Account_Code,Cust_Ref,Amendment_No,Item_Code,Rate,Order_Qty,Despatch_Qty,Active_Flag,Cust_Mtrl,Cust_DrgNo,Packing,Others,Excise_Duty,Cust_Drg_Desc,Tool_Cost,Authorized_flag,OpenSO,SalesTax_Type,PerValue,InternalSONo,RevisionNo,Remarks,MRP,Abantment_Code,AccessibleRateforMRP,Packing_Type,TOOL_AMOR_FLAG,ShowInAuth,ADD_Excise_Duty from Cust_ord_dtl where unit_code='" & gstrUNITID & "' and  Item_code ='" & m_Item_Code & "' and cust_drgNo ='" & varHelpItem & "' and cust_ref ='" & txtReferenceNo.Text & "' and Active_flag ='A' and Authorized_Flag =1"
                    rsitem = New ClsResultSetDB
                    rsitem.GetResult(m_strSql)
                    If rsitem.GetNoRows > 0 Then
                        If rsitem.GetValue("OpenSO") = False Then
                            ssPOEntry.Row = ssPOEntry.MaxRows : ssPOEntry.Col = 1 : ssPOEntry.Col2 = 1 : ssPOEntry.Value = CheckState.Unchecked
                            ssPOEntry.Row = ssPOEntry.MaxRows : ssPOEntry.Row2 = ssPOEntry.MaxRows : ssPOEntry.Col = 5 : ssPOEntry.Col2 = 5 : ssPOEntry.BlockMode = True : ssPOEntry.Lock = False : ssPOEntry.BlockMode = False
                        Else
                            ssPOEntry.Row = ssPOEntry.MaxRows : ssPOEntry.Col = 1 : ssPOEntry.Col2 = 1 : ssPOEntry.Value = CheckState.Checked
                            ssPOEntry.Row = ssPOEntry.MaxRows : ssPOEntry.Row2 = ssPOEntry.MaxRows : ssPOEntry.Col = 5 : ssPOEntry.Col2 = 5 : ssPOEntry.BlockMode = True : ssPOEntry.Lock = True : ssPOEntry.BlockMode = False
                        End If
                        Call ssPOEntry.SetText(2, ssPOEntry.MaxRows, rsitem.GetValue("Cust_DrgNo"))
                        Call ssPOEntry.SetText(4, ssPOEntry.MaxRows, rsitem.GetValue("Item_Code "))
                        Call ssPOEntry.SetText(5, ssPOEntry.MaxRows, rsitem.GetValue("Order_Qty"))
                        Call ssPOEntry.SetText(13, ssPOEntry.MaxRows, rsitem.GetValue("Rate"))
                        Call ssPOEntry.SetText(6, ssPOEntry.MaxRows, (rsitem.GetValue("Rate") * ctlPerValue.Text))
                        Call ssPOEntry.SetText(14, ssPOEntry.MaxRows, rsitem.GetValue("Cust_Mtrl"))
                        Call ssPOEntry.SetText(7, ssPOEntry.MaxRows, (rsitem.GetValue("Cust_Mtrl") * ctlPerValue.Text))
                        Call ssPOEntry.SetText(15, ssPOEntry.MaxRows, rsitem.GetValue("Tool_Cost"))
                        Call ssPOEntry.SetText(8, ssPOEntry.MaxRows, (rsitem.GetValue("Tool_Cost") * ctlPerValue.Text))
                        Call ssPOEntry.SetText(9, ssPOEntry.MaxRows, rsitem.GetValue("Packing"))
                        Call ssPOEntry.SetText(10, ssPOEntry.MaxRows, rsitem.GetValue("Excise_Duty"))
                        If gblnGSTUnit = False Then
                            addExciseDuty(e.row)
                        End If

                        Call ssPOEntry.SetText(16, ssPOEntry.MaxRows, rsitem.GetValue("Others"))
                        Call ssPOEntry.SetText(11, ssPOEntry.MaxRows, (rsitem.GetValue("Others") * ctlPerValue.Text))
                        Call ssPOEntry.SetText(17, ssPOEntry.MaxRows, rsitem.GetValue("Despatch_Qty"))
                    Else
                        If txtAmendmentNo.Enabled = False Then
                            If Len(Trim(m_Item_Code)) > 0 Then
                                If rsSalesParametere.GetValue("ItemRateLink") = True Then
                                    m_strSql = " select datefrom,dateto,Party_code,Item_code,Custvend_flg,Rate,serial_no,Discount_Flag,Discount_Amount,Cust_Supplied_Material,Tool_Cost,Packaging_Flag,Packaging_Amount,Others,currency_code,edit_flg from ITemRate_Mst where unit_code='" & gstrUNITID & "' and  Serial_No = (select max(serial_no) from itemrate_mst where unit_code='" & gstrUNITID & "' and  Party_code = '" & txtCustomerCode.Text & "' and item_code = '" & varHelpItem & "' and datediff(mm,convert(varchar(10),'" & getDateForDB(DTDate.Value) & "',103),convert(varchar(10),DateFrom,103))<=0 and custVend_Flg ='C')"
                                    If IsNothing(rsitem) = False Then rsitem.ResultSetClose()
                                    rsitem = New ClsResultSetDB
                                    rsitem.GetResult(m_strSql)
                                    If rsitem.GetNoRows > 0 Then
                                        If chkOpenSo.CheckState = CheckState.Unchecked Then
                                            ssPOEntry.Row = ssPOEntry.MaxRows : ssPOEntry.Col = 1 : ssPOEntry.Col2 = 1 : ssPOEntry.Value = CheckState.Unchecked
                                        Else
                                            ssPOEntry.Row = ssPOEntry.MaxRows : ssPOEntry.Col = 1 : ssPOEntry.Col2 = 1 : ssPOEntry.Value = CheckState.Checked
                                        End If
                                        Call ssPOEntry.SetText(2, ssPOEntry.MaxRows, IIf((rsitem.GetValue("ITem_code") = "Unknown"), "", rsitem.GetValue("ITem_code")))
                                        Call ssPOEntry.SetText(4, ssPOEntry.MaxRows, m_Item_Code)
                                        Call ssPOEntry.SetText(6, ssPOEntry.MaxRows, (rsitem.GetValue("Rate") * ctlPerValue.Text))
                                        Call ssPOEntry.SetText(13, ssPOEntry.MaxRows, rsitem.GetValue("Rate"))
                                        '***
                                        Call ssPOEntry.SetText(14, ssPOEntry.MaxRows, rsitem.GetValue("Cust_supplied_Material"))
                                        Call ssPOEntry.SetText(7, ssPOEntry.MaxRows, (rsitem.GetValue("Cust_supplied_Material") * ctlPerValue.Text))
                                        Call ssPOEntry.SetText(15, ssPOEntry.MaxRows, rsitem.GetValue("Tool_Cost"))
                                        Call ssPOEntry.SetText(8, ssPOEntry.MaxRows, (rsitem.GetValue("Tool_Cost") * ctlPerValue.Text))
                                        If rsitem.GetValue("Packaging_Flag") = False Then
                                            Call ssPOEntry.SetText(9, ssPOEntry.MaxRows, ((rsitem.GetValue("Packaging_Amount") * 100) / rsitem.GetValue("Rate")))
                                        Else
                                            Call ssPOEntry.SetText(9, ssPOEntry.MaxRows, rsitem.GetValue("Packaging_Amount"))
                                        End If
                                        If gblnGSTUnit = False AndAlso addExciseDuty(e.row) = False Then
                                            ssPOEntry.MaxRows = ssPOEntry.MaxRows - 1
                                            If ssPOEntry.MaxRows < 1 Then
                                                Call ADDRow()
                                            End If
                                            ssPOEntry.Row = ssPOEntry.MaxRows : ssPOEntry.Col = 2 : ssPOEntry.Focus()
                                        End If
                                        Call ssPOEntry.SetText(16, ssPOEntry.MaxRows, rsitem.GetValue("Others"))
                                        Call ssPOEntry.SetText(11, ssPOEntry.MaxRows, (rsitem.GetValue("Others") * ctlPerValue.Text))
                                        If rsitem.GetValue("Edit_flg") = False Then
                                            ssPOEntry.Col = 6 : ssPOEntry.Col2 = ssPOEntry.MaxCols : ssPOEntry.Row = ssPOEntry.MaxRows : ssPOEntry.Row2 = ssPOEntry.MaxRows : ssPOEntry.BlockMode = True : ssPOEntry.Lock = True : ssPOEntry.BlockMode = False
                                        Else
                                            ssPOEntry.Col = 6 : ssPOEntry.Col2 = ssPOEntry.MaxCols : ssPOEntry.Row = ssPOEntry.MaxRows : ssPOEntry.Row2 = 25 : ssPOEntry.BlockMode = True : ssPOEntry.Lock = False : ssPOEntry.BlockMode = False
                                        End If
                                    End If
                                Else
                                    If chkOpenSo.CheckState = CheckState.Unchecked Then
                                        ssPOEntry.Row = ssPOEntry.MaxRows : ssPOEntry.Col = 1 : ssPOEntry.Col2 = 1 : ssPOEntry.Value = CheckState.Unchecked
                                    Else
                                        ssPOEntry.Row = ssPOEntry.MaxRows : ssPOEntry.Col = 1 : ssPOEntry.Col2 = 1 : ssPOEntry.Value = CheckState.Checked
                                    End If
                                    Call ssPOEntry.SetText(2, ssPOEntry.MaxRows, varHelpItem)
                                    Call ssPOEntry.SetText(4, ssPOEntry.MaxRows, m_Item_Code)
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        With ssPOEntry
            If e.col <> 1 Then
                Call ssSetFocus(e.row, 2)
            End If
        End With
        rsSalesParametere.ResultSetClose()
        Exit Sub    'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub ssPOEntry_ClickEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ClickEvent) Handles ssPOEntry.ClickEvent
        On Error GoTo ErrHandler
        Dim vartext As Object
        Dim varItemCode As Object
        Dim varDrgNo As Object
        Dim rsSalesParameter As ClsResultSetDB
        rsSalesParameter = New ClsResultSetDB
        If chkmultipleitem() = True And ssPOEntry.ActiveCol = 4 Then
            Call SetCellTypeCombo(ssPOEntry.ActiveRow)
        End If
        If e.col = 1 And e.row <> 0 Then
            Call ssSetFocus(e.row, 1)
            Exit Sub
        End If
        If e.col = 0 And e.row <> 0 Then
            vartext = Nothing
            Call ssPOEntry.GetText(0, e.row, vartext)
            If vartext = "*" Then
                Call ssPOEntry.SetText(0, e.row, "")
                With ssPOEntry
                    .BlockMode = True
                    .Row = e.row
                    .Row2 = e.row
                    .Col = 1
                    .Col2 = .MaxCols
                    If cmdButtons.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                        .Lock = False
                    End If
                    .ForeColor = Color.Black
                    .BlockMode = False
                    .Row = e.row : .Row2 = e.row : .Col = 2 : .Col2 = 4 : .BlockMode = True : .Lock = True : .BlockMode = False
                    .Row = e.row : .Col = 1
                    If .Value = True Then
                        .Row = e.row : .Row2 = e.row : .Col = 5 : .Col2 = 5 : .BlockMode = True : .Lock = True : .BlockMode = False
                    End If
                    varDrgNo = Nothing
                    Call .GetText(2, e.row, varDrgNo)
                    varItemCode = Nothing
                    Call .GetText(4, e.row, varItemCode)
                    If (Len(Trim(varDrgNo)) > 0) And Len(Trim(varItemCode)) > 0 Then
                        rsSalesParameter.GetResult("Select ItemRateLink from Sales_Parameter where unit_code='" & gstrUNITID & "'")
                        If rsSalesParameter.GetValue("ItemRateLink") = True Then
                            If checkforitemRate(CInt(e.row)) = False Then
                                .Row = e.row : .Row2 = e.row : .Col = 6 : .Col2 = .MaxCols : .BlockMode = True : .Lock = True : .BlockMode = False
                            Else
                                .Row = e.row : .Row2 = e.row : .Col = 6 : .Col2 = .MaxCols : .BlockMode = True : .Lock = False : .BlockMode = False
                            End If
                        Else
                            .Row = e.row : .Row2 = e.row : .Col = 6 : .Col2 = .MaxCols : .BlockMode = True : .Lock = False : .BlockMode = False
                        End If
                    Else
                        If cmdButtons.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                            .Row = 1 : .Row2 = .MaxRows : .Col = 1 : .Col2 = .MaxCols : .BlockMode = True : .Lock = False : .BlockMode = False
                        End If
                    End If
                End With
            Else
                Call ssPOEntry.SetText(0, e.row, "*")
                With ssPOEntry
                    .BlockMode = True
                    .Row = e.row
                    .Row2 = e.row
                    .Col = 1
                    .Col2 = .MaxCols
                    .ForeColor = Color.Red
                    .Lock = True
                    .BlockMode = False
                End With
            End If
            '        End If
        End If
        Exit Sub    'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub ssPOEntry_EditChange(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_EditChangeEvent) Handles ssPOEntry.EditChange
        Dim varCustSupp
        If e.col = 7 Then
            varCustSupp = Nothing
            ssPOEntry.GetText(e.col, e.row, varCustSupp)
            If varCustSupp <= 0 Then
                ssPOEntry.SetText(11, e.row, "")
                ssPOEntry.Col = 11
                ssPOEntry.Row = ssPOEntry.ActiveRow
                ssPOEntry.CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
            Else
                ssPOEntry.Col = 11
                ssPOEntry.Row = ssPOEntry.ActiveRow
                ssPOEntry.CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
            End If
        End If
    End Sub
    Private Sub ssPOEntry_KeyDownEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyDownEvent) Handles ssPOEntry.KeyDownEvent
        On Error GoTo ErrHandler
        Dim varHelpItem As Object
        Dim rsDesc As New ClsResultSetDB
        Dim rsSalesParameter As New ClsResultSetDB
        Dim inti As Integer
        Dim strReturnCustRef As String
        Dim strSOEntry() As String
        Dim StrCT2Condition As String = ""
        Dim strDockCodeCondition As String = ""
        Dim strdockcode As String = ""
        Dim stritemcode As String
        Dim strcustitemcode As String
        Dim objSQLConn As SqlConnection
        Dim objReader As SqlDataReader
        Dim objCommand As SqlCommand
        Dim STRSQL As String

        '10797956 
        If ChkCT2Reqd.Checked = True Then
            StrCT2Condition = " And Exists(select * From CT2_Cust_Item_Linkage Z where A.UNIT_CODE=z.Unit_Code and A.Account_Code=z.Customer_Code and A.Item_code =z.Item_Code and A.Cust_Drgno=z.Cust_drgno And z.Active=1 and z.isAuthorized=1)"
            With ssPOEntry
                .Col = 10
                .Text = "EX0"
                .Lock = True
                .Col = 11
                .Text = "AED0"
                .Lock = True
            End With
        End If

        'If user has pressed Ctrl + N, then add a new row
        rsSalesParameter.GetResult("Select ItemRateLink from Sales_Parameter  where unit_code='" & gstrUNITID & "'")
        If (e.shift = 2 And e.keyCode = Keys.N) Then
            'Add a blank row
            If ValidRowData(ssPOEntry.ActiveRow, 0) Then Call ADDRow()
            'Setting the focus
            'change account Plug in
            ssPOEntry.Row = ssPOEntry.MaxRows : ssPOEntry.Col = 2 : ssPOEntry.Action = FPSpreadADO.ActionConstants.ActionActiveCell
        End If
        'change account Plug in
        If ssPOEntry.ActiveCol = 2 Or ssPOEntry.ActiveCol = 4 Then
            If e.keyCode = 112 Then
                If Len(Trim(txtCurrencyType.Text)) = 0 Then Exit Sub 'for Currency Type entered Validation
                If txtAmendmentNo.Enabled = False Then
                    If rsSalesParameter.GetValue("ItemRateLink") = True Then
                        '10532789 active = 1 added
                        'begin
                        If Mid(cmbPOType.Text, 1, 1) = "V" Then
                            strSOEntry = ctlEMPHelpSOEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select a.Cust_drgNo,a.ITem_code,a.drg_Desc from CustItem_Mst a,Itemrate_Mst b,Item_mst c where a.unit_code=b.unit_code and  a.unit_code='" & gstrUNITID & "' and a.Account_Code='" & txtCustomerCode.Text & "' and a.Account_code = b.Party_Code and a.cust_DrgNo = b.Item_code and a.Unit_Code=c.Unit_Code and a.Item_Code=c.Item_code and c.status='A' and c.Hold_Flag=0 and c.ITEM_MAIN_GRP='M' and  a.active = 1 and b.serial_no=(select max(serial_no) from itemrate_mst i1 Where i1.unit_code=b.unit_code and i1.party_code = b.party_code and i1.item_code=b.item_code) and datediff(mm,'" & getDateForDB(DTDate.Value) & "',b.DateFrom)<=0 and CustVend_Flg = 'C'" & IIf(StrCT2Condition.Trim.Length = 0, "", StrCT2Condition))
                        Else
                            strSOEntry = ctlEMPHelpSOEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select a.Cust_drgNo,a.ITem_code,a.drg_Desc from CustItem_Mst a,Itemrate_Mst b where a.unit_code=b.unit_code and  a.unit_code='" & gstrUNITID & "' and a.Account_Code='" & txtCustomerCode.Text & "' and a.Account_code = b.Party_Code and a.cust_DrgNo = b.Item_code and a.active = 1 and b.serial_no=(select max(serial_no) from itemrate_mst i1 Where i1.unit_code=b.unit_code and i1.party_code = b.party_code and i1.item_code=b.item_code) and datediff(mm,'" & getDateForDB(DTDate.Value) & "',b.DateFrom)<=0 and CustVend_Flg = 'C'" & IIf(StrCT2Condition.Trim.Length = 0, "", StrCT2Condition))
                        End If
                        'end
                    Else
                        '10869290
                        If Mid(cmbPOType.Text, 1, 1) = "V" Then
                            '10856126
                            STRSQL = "select dbo.UDF_ISDOCKCODE_ASN( '" & gstrUNITID & "','" & txtCustomerCode.Text.Trim & "')"
                            If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(STRSQL)) = True Then
                                If ssPOEntry.MaxRows = 1 Then 'FUNCTIONALITY IS ON 
                                    strSOEntry = ctlEMPHelpSOEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select Cust_drgNo,a.ITem_code,drg_Desc from CustItem_Mst A, Item_mst B where A.Item_Code=B.Item_Code and A.Unit_Code=B.Unit_Code and B.status ='A' and B.Hold_Flag=0 and B.Item_Main_Grp='M' and a.unit_code='" & gstrUNITID & "' and a.Account_Code='" & txtCustomerCode.Text & "' and a.active = 1" & IIf(StrCT2Condition.Trim.Length = 0, "", StrCT2Condition)) '10532789 active = 1 added
                                Else
                                    With ssPOEntry
                                        .Row = 1
                                        .Col = 2
                                        strcustitemcode = .Text
                                        .Col = 4
                                        stritemcode = .Text

                                    End With
                                    STRSQL = "SELECT DOCKCODE FROM CUSTITEM_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND ACCOUNT_CODE='" & txtCustomerCode.Text & "'" &
                                             " AND ITEM_CODE='" & stritemcode & "' AND CUST_DRGNO='" & strcustitemcode & "' AND ACTIVE=1 "
                                    strdockcode = Convert.ToString(SqlConnectionclass.ExecuteScalar(STRSQL))
                                    strDockCodeCondition = "and Dockcode in ( '" & strdockcode & "' )"
                                    strSOEntry = ctlEMPHelpSOEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select Cust_drgNo,a.ITem_code,drg_Desc from CustItem_Mst a, Item_mst b" &
                                                " where a.Unit_Code=b.Unit_Code and b.Status='A' and b.Item_Main_Grp = 'M' and a.Item_Code=b.Item_Code and a.unit_code='" & gstrUNITID & "'" &
                                                " and a.Account_Code='" & txtCustomerCode.Text & "'" &
                                                " and a.active = 1" & IIf(StrCT2Condition.Trim.Length = 0, "", StrCT2Condition) &
                                                        IIf(strDockCodeCondition.Trim.Length = 0, "", strDockCodeCondition))
                                    'strSOEntry = ctlEMPHelpSOEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select Cust_drgNo,ITem_code,drg_Desc from CustItem_Mst A where  unit_code='" & gstrUNITID & "' and Account_Code='" & txtCustomerCode.Text & "' and active = 1" & IIf(StrCT2Condition.Trim.Length = 0, "", StrCT2Condition) & IIf(strDockCodeCondition.Trim.Length = 0, "", strDockCodeCondition)) '10532789 active = 1 added
                                End If
                            Else
                                strSOEntry = ctlEMPHelpSOEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select Cust_drgNo,a.ITem_code,drg_Desc from CustItem_Mst A, Item_mst B where a.Unit_Code=b.Unit_Code and a.Item_Code=b.Item_Code and b.Status='A' and b.Hold_Flag=0 and b.Item_Main_Grp='M' and a.unit_code='" & gstrUNITID & "' and a.Account_Code='" & txtCustomerCode.Text & "' and a.active = 1" & IIf(StrCT2Condition.Trim.Length = 0, "", StrCT2Condition)) '10532789 active = 1 added
                            End If
                        Else
                            '10856126
                            STRSQL = "select dbo.UDF_ISDOCKCODE_ASN( '" & gstrUNITID & "','" & txtCustomerCode.Text.Trim & "')"
                            If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(STRSQL)) = True Then
                                If ssPOEntry.MaxRows = 1 Then 'FUNCTIONALITY IS ON 
                                    strSOEntry = ctlEMPHelpSOEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select Cust_drgNo,ITem_code,drg_Desc from CustItem_Mst A where  unit_code='" & gstrUNITID & "' and Account_Code='" & txtCustomerCode.Text & "' and active = 1" & IIf(StrCT2Condition.Trim.Length = 0, "", StrCT2Condition)) '10532789 active = 1 added
                                Else
                                    With ssPOEntry
                                        .Row = 1
                                        .Col = 2
                                        strcustitemcode = .Text
                                        .Col = 4
                                        stritemcode = .Text

                                    End With
                                    STRSQL = "SELECT DOCKCODE FROM CUSTITEM_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND ACCOUNT_CODE='" & txtCustomerCode.Text & "'" &
                                             " AND ITEM_CODE='" & stritemcode & "' AND CUST_DRGNO='" & strcustitemcode & "' AND ACTIVE=1 "
                                    strdockcode = Convert.ToString(SqlConnectionclass.ExecuteScalar(STRSQL))
                                    strDockCodeCondition = "and Dockcode in ( '" & strdockcode & "' )"
                                    strSOEntry = ctlEMPHelpSOEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select Cust_drgNo,ITem_code,drg_Desc from CustItem_Mst a" &
                                                " where unit_code='" & gstrUNITID & "'" &
                                                " and Account_Code='" & txtCustomerCode.Text & "'" &
                                                " and active = 1" & IIf(StrCT2Condition.Trim.Length = 0, "", StrCT2Condition) &
                                                        IIf(strDockCodeCondition.Trim.Length = 0, "", strDockCodeCondition))
                                    'strSOEntry = ctlEMPHelpSOEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select Cust_drgNo,ITem_code,drg_Desc from CustItem_Mst A where  unit_code='" & gstrUNITID & "' and Account_Code='" & txtCustomerCode.Text & "' and active = 1" & IIf(StrCT2Condition.Trim.Length = 0, "", StrCT2Condition) & IIf(strDockCodeCondition.Trim.Length = 0, "", strDockCodeCondition)) '10532789 active = 1 added
                                End If
                            Else
                                strSOEntry = ctlEMPHelpSOEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select Cust_drgNo,ITem_code,drg_Desc from CustItem_Mst A where  unit_code='" & gstrUNITID & "' and Account_Code='" & txtCustomerCode.Text & "' and active = 1" & IIf(StrCT2Condition.Trim.Length = 0, "", StrCT2Condition)) '10532789 active = 1 added
                            End If
                        End If
                    End If
                Else
                    If Mid(cmbPOType.Text, 1, 1) = "V" Then

                        STRSQL = "select dbo.UDF_ISDOCKCODE_ASN( '" & gstrUNITID & "','" & txtCustomerCode.Text.Trim & "')"
                        If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(STRSQL)) = True Then
                            With ssPOEntry
                                .Row = 1
                                .Col = 2
                                strcustitemcode = .Text
                                .Col = 4
                                stritemcode = .Text
                            End With
                            STRSQL = "SELECT DOCKCODE FROM CUSTITEM_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND ACCOUNT_CODE='" & txtCustomerCode.Text & "'" &
                                     " AND ITEM_CODE='" & stritemcode & "' AND CUST_DRGNO='" & strcustitemcode & "' AND ACTIVE=1 "
                            strdockcode = Convert.ToString(SqlConnectionclass.ExecuteScalar(STRSQL))
                            strDockCodeCondition = "and Dockcode in ( '" & strdockcode & "' )"

                            strSOEntry = ctlEMPHelpSOEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select Cust_drgNo,a.ITem_code,drg_Desc from CustItem_Mst A, Item_mst B where A.Unit_Code=B.Unit_Code and A.Item_Code=B.Item_Code and B.Status='A' and B.Hold_Flag=0 and B.Item_Main_Grp='M' and A.unit_code='" & gstrUNITID & "' and A.Account_Code='" & txtCustomerCode.Text & "' and A.active = 1" & IIf(StrCT2Condition.Trim.Length = 0, "", StrCT2Condition) & IIf(strDockCodeCondition.Trim.Length = 0, "", strDockCodeCondition)) '10532789 active = 1 added
                        Else
                            strSOEntry = ctlEMPHelpSOEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select Cust_drgNo,a.ITem_code,drg_Desc from CustItem_Mst A, Item_mst B where A.Unit_Code=B.Unit_Code and A.Item_Code=B.Item_Code and B.Status='A' and B.Hold_Flag=0 and B.Item_Main_Grp='M' and A.unit_code='" & gstrUNITID & "' and A.Account_Code='" & txtCustomerCode.Text & "' and A.active = 1" & IIf(StrCT2Condition.Trim.Length = 0, "", StrCT2Condition)) '10532789 active = 1 added
                        End If

                    Else
                        ''''10856126
                        STRSQL = "select dbo.UDF_ISDOCKCODE_ASN( '" & gstrUNITID & "','" & txtCustomerCode.Text.Trim & "')"
                        If Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(STRSQL)) = True Then
                            With ssPOEntry
                                .Row = 1
                                .Col = 2
                                strcustitemcode = .Text
                                .Col = 4
                                stritemcode = .Text
                            End With
                            STRSQL = "SELECT DOCKCODE FROM CUSTITEM_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND ACCOUNT_CODE='" & txtCustomerCode.Text & "'" &
                                     " AND ITEM_CODE='" & stritemcode & "' AND CUST_DRGNO='" & strcustitemcode & "' AND ACTIVE=1 "
                            strdockcode = Convert.ToString(SqlConnectionclass.ExecuteScalar(STRSQL))
                            strDockCodeCondition = "and Dockcode in ( '" & strdockcode & "' )"

                            strSOEntry = ctlEMPHelpSOEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select Cust_drgNo,ITem_code,drg_Desc from CustItem_Mst A where  unit_code='" & gstrUNITID & "' and Account_Code='" & txtCustomerCode.Text & "' and active = 1" & IIf(StrCT2Condition.Trim.Length = 0, "", StrCT2Condition) & IIf(strDockCodeCondition.Trim.Length = 0, "", strDockCodeCondition)) '10532789 active = 1 added
                        Else
                            strSOEntry = ctlEMPHelpSOEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select Cust_drgNo,ITem_code,drg_Desc from CustItem_Mst A where  unit_code='" & gstrUNITID & "' and Account_Code='" & txtCustomerCode.Text & "' and active = 1" & IIf(StrCT2Condition.Trim.Length = 0, "", StrCT2Condition)) '10532789 active = 1 added
                        End If
                        '''
                    End If


                End If
                If UBound(strSOEntry) < 0 Then Exit Sub
                If strSOEntry(0) = "0" Then
                    Call ConfirmWindow(10013, ConfirmWindowButtonsEnum.BUTTON_OK, ConfirmWindowImagesEnum.IMG_INFO)
                Else
                    strReturnCustRef = CheckForMultipleOpenSO(txtCustomerCode.Text, txtConsCode.Text, txtReferenceNo.Text, strSOEntry(0), strSOEntry(1))
                    If Len(strReturnCustRef) > 0 Then
                        MsgBox("More than one Sales Order cannot be active for Customer Item Combination" & strReturnCustRef, vbInformation, ResolveResString(100))
                        Exit Sub
                    End If
                    Call ssPOEntry.SetText(2, ssPOEntry.ActiveRow, strSOEntry(0))
                    Call ssPOEntry.SetText(4, ssPOEntry.ActiveRow, strSOEntry(1))
                    lblCustPartDesc.Text = strSOEntry(2)
                End If
                'GST DETAILS
                If gblnGSTUnit = True Then
                    STRSQL = "set dateformat 'dmy' SELECT * FROM DBO.UFN_GST_ITEMWISETAXES('" & gstrUNITID & "','" & txtCustomerCode.Text.Trim & "','" & strSOEntry(1) & "','" & DTEffectiveDate.Value & "','" & DTValidDate.Value & "')"
                    objSQLConn = SqlConnectionclass.GetConnection()
                    objCommand = New SqlCommand(STRSQL, objSQLConn)
                    objReader = objCommand.ExecuteReader()
                    If objReader.HasRows = True Then
                        objReader.Read()
                        Call ssPOEntry.SetText(20, ssPOEntry.ActiveRow, objReader.GetValue(1))
                        Call ssPOEntry.SetText(21, ssPOEntry.ActiveRow, objReader.GetValue(2))
                        Call ssPOEntry.SetText(22, ssPOEntry.ActiveRow, objReader.GetValue(3))
                        Call ssPOEntry.SetText(23, ssPOEntry.ActiveRow, objReader.GetValue(4))
                        Call ssPOEntry.SetText(24, ssPOEntry.ActiveRow, objReader.GetValue(5))
                        Call ssPOEntry.SetText(25, ssPOEntry.ActiveRow, objReader.GetValue(6))
                        Call ssPOEntry.SetText(26, ssPOEntry.ActiveRow, objReader.GetValue(0))
                        'GST CHANGE
                    End If

                    objReader = Nothing
                    objSQLConn.Close()
                    objSQLConn = Nothing

                End If
                'GST DETAILS

                If cmdButtons.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                    Dim rsitem As ClsResultSetDB
                    If Len(Trim(m_Item_Code)) > 0 Then
                        m_strSql = "SElect Account_Code,Cust_Ref,Amendment_No,Order_Date,Amendment_Date,Active_Flag,Currency_Code,Valid_Date,Effect_Date,Term_Payment,Special_Remarks,Pay_Remarks,Price_Remarks,Packing_Remarks,Frieght_Remarks,Transport_Remarks,Octorai_Remarks,Mode_Despatch,Delivery,First_Authorized,Second_Authorized,Third_Authorized,Authorized_Flag,Reason,PO_Type,SalesTax_Type,OpenSO,AddCustSupp,PerValue,InternalSONo,RevisionNo,Surcharge_code,Future_SO,ECESS_Code,Consignee_Code,ADDVAT_TYPE,HSNSACCODE,  ISHSNORSAC, CGSTTXRT_TYPE,SGSTTXRT_TYPE, UTGSTTXRT_TYPE,IGSTTXRT_TYPE,COMPENSATION_CESS from Cust_ord_dtl where  unit_code='" & gstrUNITID & "' and Item_code ='" & m_Item_Code & "' and cust_drgNo ='" & varHelpItem & "' and cust_ref ='" & txtReferenceNo.Text & "' and Active_flag ='A' and Authorized_Flag =1"
                        If IsNothing(rsitem) = False Then rsitem.ResultSetClose()
                        rsitem = New ClsResultSetDB
                        rsitem.GetResult(m_strSql)
                        If rsitem.GetNoRows > 0 Then
                            'change account Plug in
                            Call ssPOEntry.SetText(2, ssPOEntry.MaxRows, rsitem.GetValue("Cust_DrgNo"))
                            Call ssPOEntry.SetText(4, ssPOEntry.MaxRows, rsitem.GetValue("Item_Code "))
                            Call ssPOEntry.SetText(5, ssPOEntry.MaxRows, rsitem.GetValue("Order_Qty"))
                            Call ssPOEntry.SetText(13, ssPOEntry.MaxRows, rsitem.GetValue("Rate"))
                            Call ssPOEntry.SetText(6, ssPOEntry.MaxRows, (rsitem.GetValue("Rate") * ctlPerValue.Text))
                            Call ssPOEntry.SetText(14, ssPOEntry.MaxRows, rsitem.GetValue("Cust_Mtrl"))
                            Call ssPOEntry.SetText(7, ssPOEntry.MaxRows, (rsitem.GetValue("Cust_Mtrl") * ctlPerValue.Text))
                            Call ssPOEntry.SetText(15, ssPOEntry.MaxRows, rsitem.GetValue("Tool_Cost"))
                            Call ssPOEntry.SetText(8, ssPOEntry.MaxRows, (rsitem.GetValue("Tool_Cost") * ctlPerValue.Text))
                            Call ssPOEntry.SetText(9, ssPOEntry.MaxRows, rsitem.GetValue("Packing"))
                            Call ssPOEntry.SetText(10, ssPOEntry.MaxRows, rsitem.GetValue("Excise_Duty"))
                            Call ssPOEntry.SetText(16, ssPOEntry.MaxRows, rsitem.GetValue("Others"))
                            Call ssPOEntry.SetText(11, ssPOEntry.MaxRows, (rsitem.GetValue("Others") * ctlPerValue.Text))
                            rsitem.ResultSetClose()
                        Else
                            If txtAmendmentNo.Enabled = False Then
                                If Len(Trim(m_Item_Code)) > 0 Then
                                    If rsSalesParameter.GetValue("ItemRateLink") = True Then
                                        m_strSql = " select datefrom,dateto,Party_code,Item_code,Custvend_flg,Rate,serial_no,Discount_Flag,Discount_Amount,Cust_Supplied_Material,Tool_Cost,Packaging_Flag,Packaging_Amount,Others,currency_code,edit_flg from ITemRate_Mst where unit_code='" & gstrUNITID & "' and Serial_No = (select max(serial_no) from itemrate_mst where unit_code='" & gstrUNITID & "' and Party_code = '" & txtCustomerCode.Text & "' and item_code = '" & varHelpItem & "' and datediff(mm,convert(varchar(10),'" & getDateForDB(DTDate.Value) & "',103),convert(varchar(10),DateFrom,103))<=0 and custVend_Flg ='C')"
                                        If IsNothing(rsitem) = False Then rsitem.ResultSetClose()
                                        rsitem = New ClsResultSetDB
                                        rsitem.GetResult(m_strSql)
                                        If rsitem.GetNoRows > 0 Then
                                            If chkOpenSo.CheckState = CheckState.Unchecked Then
                                                ssPOEntry.Row = ssPOEntry.MaxRows : ssPOEntry.Col = 1 : ssPOEntry.Col2 = 1 : ssPOEntry.Value = CheckState.Unchecked
                                                ssPOEntry.Row = ssPOEntry.MaxRows : ssPOEntry.Row2 = ssPOEntry.MaxRows : ssPOEntry.Col = 5 : ssPOEntry.Col2 = 5 : ssPOEntry.BlockMode = True : ssPOEntry.Lock = False : ssPOEntry.BlockMode = False
                                            Else
                                                ssPOEntry.Row = ssPOEntry.MaxRows : ssPOEntry.Col = 1 : ssPOEntry.Col2 = 1 : ssPOEntry.Value = CheckState.Checked
                                                ssPOEntry.Row = ssPOEntry.MaxRows : ssPOEntry.Row2 = ssPOEntry.MaxRows : ssPOEntry.Col = 5 : ssPOEntry.Col2 = 5 : ssPOEntry.BlockMode = True : ssPOEntry.Lock = True : ssPOEntry.BlockMode = False
                                            End If
                                            Call ssPOEntry.SetText(2, ssPOEntry.MaxRows, IIf((rsitem.GetValue("ITem_code") = "Unknown"), "", rsitem.GetValue("ITem_code")))
                                            Call ssPOEntry.SetText(4, ssPOEntry.MaxRows, m_Item_Code)
                                            Call ssPOEntry.SetText(13, ssPOEntry.MaxRows, rsitem.GetValue("Rate"))
                                            Call ssPOEntry.SetText(6, ssPOEntry.MaxRows, (rsitem.GetValue("Rate") * ctlPerValue.Text))
                                            Call ssPOEntry.SetText(14, ssPOEntry.MaxRows, rsitem.GetValue("Cust_supplied_Material"))
                                            Call ssPOEntry.SetText(7, ssPOEntry.MaxRows, (rsitem.GetValue("Cust_supplied_Material") * ctlPerValue.Text))
                                            Call ssPOEntry.SetText(15, ssPOEntry.MaxRows, rsitem.GetValue("Tool_Cost"))
                                            Call ssPOEntry.SetText(8, ssPOEntry.MaxRows, (rsitem.GetValue("Tool_Cost") * ctlPerValue.Text))
                                            If rsitem.GetValue("Packaging_Flag") = False Then
                                                Call ssPOEntry.SetText(9, ssPOEntry.MaxRows, ((rsitem.GetValue("Packaging_Amount") * 100) / rsitem.GetValue("Rate")))
                                            Else
                                                Call ssPOEntry.SetText(9, ssPOEntry.MaxRows, rsitem.GetValue("Packaging_Amount"))
                                            End If
                                            Call ssPOEntry.SetText(16, ssPOEntry.MaxRows, rsitem.GetValue("Others"))
                                            Call ssPOEntry.SetText(11, ssPOEntry.MaxRows, (rsitem.GetValue("Others") * ctlPerValue.Text))
                                            If rsitem.GetValue("Edit_flg") = False Then
                                                ssPOEntry.Col = 6 : ssPOEntry.Col2 = ssPOEntry.MaxCols : ssPOEntry.Row = ssPOEntry.MaxRows : ssPOEntry.Row2 = ssPOEntry.MaxRows : ssPOEntry.BlockMode = True : ssPOEntry.Lock = True : ssPOEntry.BlockMode = False
                                            Else
                                                ssPOEntry.Col = 6 : ssPOEntry.Col2 = ssPOEntry.MaxCols : ssPOEntry.Row = ssPOEntry.MaxRows : ssPOEntry.Row2 = ssPOEntry.MaxRows : ssPOEntry.BlockMode = True : ssPOEntry.Lock = False : ssPOEntry.BlockMode = False
                                            End If
                                        Else
                                            'If itemratelink in salesparameter is false then
                                            If chkOpenSo.CheckState = CheckState.Unchecked Then
                                                ssPOEntry.Row = ssPOEntry.MaxRows : ssPOEntry.Col = 1 : ssPOEntry.Col2 = 1 : ssPOEntry.Value = CheckState.Unchecked
                                                ssPOEntry.Row = ssPOEntry.MaxRows : ssPOEntry.Row2 = ssPOEntry.MaxRows : ssPOEntry.Col = 5 : ssPOEntry.Col2 = 5 : ssPOEntry.BlockMode = True : ssPOEntry.Lock = False : ssPOEntry.BlockMode = False
                                            Else
                                                ssPOEntry.Row = ssPOEntry.MaxRows : ssPOEntry.Col = 1 : ssPOEntry.Col2 = 1 : ssPOEntry.Value = CheckState.Checked
                                                ssPOEntry.Row = ssPOEntry.MaxRows : ssPOEntry.Row2 = ssPOEntry.MaxRows : ssPOEntry.Col = 5 : ssPOEntry.Col2 = 5 : ssPOEntry.BlockMode = True : ssPOEntry.Lock = True : ssPOEntry.BlockMode = False
                                            End If
                                            Call ssPOEntry.SetText(2, ssPOEntry.MaxRows, varHelpItem)
                                            Call ssPOEntry.SetText(4, ssPOEntry.MaxRows, m_Item_Code)
                                        End If
                                        rsitem.ResultSetClose()
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        ElseIf ssPOEntry.ActiveCol = 10 Then
            If e.keyCode = 112 Then
                If Len(Trim(txtCurrencyType.Text)) = 0 Then Exit Sub 'for Currency Type entered Validation
                If ChkCT2Reqd.Checked = False Then
                    If gblnGSTUnit = False Then
                        With ssPOEntry
                            .Row = .ActiveRow : .Col = .ActiveCol
                            varHelpItem = ShowList(1, 6, Trim(.Text), "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND Tx_TaxeID='EXC'")
                            If varHelpItem = "-1" Then
                                MsgBox("No Excise Available to Display", vbInformation, "empower")
                                .Text = ""
                            Else
                                Call ssPOEntry.SetText(10, ssPOEntry.ActiveRow, varHelpItem)
                            End If
                        End With
                    End If

                End If
            End If
        ElseIf ssPOEntry.ActiveCol = 11 Then
            If e.keyCode = 112 Then
                If Len(Trim(txtCurrencyType.Text)) = 0 Then Exit Sub 'for Currency Type entered Validation
                If ChkCT2Reqd.Checked = False Then
                    If gblnGSTUnit = False Then
                        With ssPOEntry
                            .Row = .ActiveRow : .Col = .ActiveCol
                            varHelpItem = ShowList(1, 6, Trim(.Text), "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND Tx_TaxeID='AED'")
                            If varHelpItem = "-1" Then
                                MsgBox("No Excise Available to Display", vbInformation, "empower")
                                .Text = ""
                            Else
                                Call ssPOEntry.SetText(11, ssPOEntry.ActiveRow, varHelpItem)
                            End If
                        End With
                    End If
                End If
            End If
        End If
        rsSalesParameter.ResultSetClose()
        Exit Sub    'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub ssPOEntry_KeyPressEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyPressEvent) Handles ssPOEntry.KeyPressEvent
        On Error GoTo ErrHandler
        Select Case e.keyAscii
            Case 39, 34, 96
                e.keyAscii = 0
        End Select
        Exit Sub            'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub ssPOEntry_LeaveCell(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_LeaveCellEvent) Handles ssPOEntry.LeaveCell
        On Error GoTo ErrHandler
        'Exit Sub
        Dim rsSalesParameter As ClsResultSetDB
        rsSalesParameter = New ClsResultSetDB
        If e.newRow < 1 Then Exit Sub
        If ValidRowData(e.row, e.col) = True Then
            'change account Plug in
            If (e.col = 2) Or (e.col = 4) Then
                With ssPOEntry
                    .Col = 2 : .Row = e.row
                    If Len(Trim(.Text)) > 0 Then
                        .Col = 4 : .Row = e.row
                        If Len(Trim(.Text)) > 0 Then
                            If ChkCT2Reqd.Checked = False Then
                                If gblnGSTUnit = False AndAlso addExciseDuty(e.row) = False And gstrUNITID <> "STH" Then
                                    'If addExciseDuty(e.row) = False Then
                                    ssPOEntry.MaxRows = ssPOEntry.MaxRows - 1
                                    If ssPOEntry.MaxRows < 1 Then
                                        Call ADDRow()
                                    End If
                                    ssPOEntry.Row = ssPOEntry.MaxRows : ssPOEntry.Col = 2 : ssPOEntry.Focus()
                                End If
                            End If
                        End If
                    End If
                End With
            End If
        Else
            With ssPOEntry
                'change account Plug in
                .Row = intRow : .Row2 = intRow : .Col = 11 : .Col2 = 11 : .BlockMode = True : .Lock = False : .BlockMode = False
            End With
        End If
        If (e.col = 1) Then
            With ssPOEntry
                .Row = e.row : .Col = 1
                If .Value = 1 Then
                    .Row = e.row : .Row2 = e.row : .Col = 5 : .Col2 = 5 : .BlockMode = True : .Text = 0 : .Lock = True : .BlockMode = False
                Else
                    .Row = e.row : .Row2 = e.row : .Col = 5 : .Col2 = 5 : .BlockMode = True : .Lock = False : .BlockMode = False
                End If
            End With
        End If
        If (e.col = 8) Then
            With ssPOEntry
                .Row = e.row : .Col = 8
                If Val(.Text) = 0 Then
                    rsSalesParameter.GetResult("Select ToolCostMsg from Sales_Parameter where unit_code='" & gstrUNITID & "'")
                    If rsSalesParameter.GetValue("ToolCostMsg") = True Then
                        MsgBox("You have Entered 0 Tool Cost", vbInformation, "empower")
                        .Row = e.newRow : .Col = e.newCol : .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                        .Focus()
                    End If
                End If
            End With
        End If
        If (e.col = 9) Then
            'With ssPOEntry
            '    .Row = e.row : .Col = 9
            '    If Val(.Text) > 100 Then
            '        MsgBox("Packing  % can not be greater than 100")
            '        .Row = e.row : .Col = 9
            '        .Action = FPSpreadADO.ActionConstants.ActionActiveCell
            '        .Focus()
            '    End If
            'End With
        End If
        If (e.col = 10) Then
            With ssPOEntry
                .Row = e.row : .Col = 10
                If Len(Trim(.Text)) > 0 Then
                    If ChkCT2Reqd.Checked = False Then
                        If gblnGSTUnit = False Then
                            rsSalesParameter.GetResult("Select Txrt_Rate_no,TxRt_Percentage from Gen_TaxRate where unit_code='" & gstrUNITID & "' and Tx_TaxeID = 'EXC' and Txrt_Rate_no = '" & .Text & "'")
                            If rsSalesParameter.GetNoRows = 0 Then
                                MsgBox("Invalid Excise Code.", vbInformation, "empower")
                                .Row = e.row : .Col = 10
                                .Text = ""
                                .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                .Focus()
                            End If
                        Else
                            MsgBox("Excise not allowed.", vbInformation, "empower")
                            .Row = e.row : .Col = 10
                            .Text = ""

                        End If
                    End If
                End If
            End With
        End If
        Dim varCustSupp
        If (e.col = 11) Then
            With ssPOEntry
                .Row = e.row : .Col = 11
                If Len(Trim(.Text)) > 0 Then
                    If ChkCT2Reqd.Checked = False Then
                        If gblnGSTUnit = False Then
                            rsSalesParameter.GetResult("Select Txrt_Rate_no,TxRt_Percentage from Gen_TaxRate where unit_code='" & gstrUNITID & "' and Tx_TaxeID = 'AED'")
                            If rsSalesParameter.GetNoRows = 0 Then
                                MsgBox("Invalid Additional Excise Duty Code.", vbInformation, "empower")
                                .Row = e.row : .Col = 11
                                .Text = ""
                                .Focus()
                                .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                            End If
                        Else
                            MsgBox("Additional Excise not allowed.", vbInformation, "empower")
                            .Row = e.row : .Col = 11
                            .Text = ""

                        End If
                    End If
                End If
            End With
        End If
        If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Or cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
            'change account Plug in
            If (e.col = 2) Or (e.col = 4) Then
                Dim strDrgNo As Object
                Dim GetDetails As Boolean
                Dim rsitem As ClsResultSetDB
                Dim strcustdtl As String
                Dim StrItemCode As Object
                If IsNothing(rsitem) = False Then rsitem.ResultSetClose()
                rsitem = New ClsResultSetDB
                Dim strCT2Condition As String = ""
                'change account Plug in
                If e.col = 2 Then
                    strDrgNo = Nothing
                    Call ssPOEntry.GetText(e.col, ssPOEntry.MaxRows, strDrgNo)

                    If ChkCT2Reqd.Checked = True Then
                        strcustdtl = "Select ITem_code,Drg_desc from custITem_Mst where unit_code='" & gstrUNITID & "' and cust_drgNo ='" & strDrgNo & "'and Account_code ='" & txtCustomerCode.Text & "' and active = 1" '10532789 active = 1 added
                        rsitem.GetResult(strcustdtl)
                        If rsitem.GetNoRows > 1 Then
                            GetDetails = False
                            MsgBox("This Part code has more then two Items linked, Please select one from Item ListBox", vbInformation, "empower")
                            SetCellTypeCombo(e.row)
                            Call ssSetFocus(e.row, 4)
                            Exit Sub
                        End If
                    End If
                    '10736222-changes done for CT2 ARE 3 functionality - to enable CT2 Reqd flag
                    If ChkCT2Reqd.Checked = True Then
                        strCT2Condition = " And Exists(select * From CT2_Cust_Item_Linkage Z where A.UNIT_CODE=z.Unit_Code and A.Account_Code=z.Customer_Code and A.Item_code =z.Item_Code and A.Cust_Drgno=z.Cust_drgno And z.Active=1 and z.isAuthorized=1)"
                    End If

                    '10736222-changes done for CT2 ARE 3 functionality - to enable CT2 Reqd flag
                    strcustdtl = "Select ITem_code,Drg_desc from custITem_Mst A where ACTIVE = 1 and unit_code='" & gstrUNITID & "' and cust_drgNo ='" & strDrgNo & "' and Account_code ='" & txtCustomerCode.Text & "' " & strCT2Condition
                    rsitem.GetResult(strcustdtl)
                    If rsitem.GetNoRows > 1 Then
                        GetDetails = False
                        MsgBox("This Part code has more then two Items linked, Please select one from Item ListBox", vbInformation, ResolveResString(100))
                        SetCellTypeCombo(e.row)
                        Call ssSetFocus(e.row, 4)
                        Exit Sub
                    Else
                        GetDetails = True
                        StrItemCode = IIf((rsitem.GetValue("ITem_code") = "Unknown"), "", rsitem.GetValue("ITem_code"))
                        lblCustPartDesc.Text = rsitem.GetValue("Drg_desc")
                        Call ssPOEntry.SetText(4, ssPOEntry.MaxRows, StrItemCode)
                        rsSalesParameter.GetResult("Select ItemRateLink from Sales_Parameter where unit_code='" & gstrUNITID & "'")
                        If rsSalesParameter.GetValue("ItemRateLink") = True Then
                            Call RowDetailsfromKeyBoard(StrItemCode, strDrgNo)
                            '******
                        End If
                    End If
                End If
                If e.col = 4 Then
                    Call SetCellStatic(e.row)
                End If
                If GetDetails = True Then
                    m_strSql = "Select Account_Code,Cust_Ref,Amendment_No,Item_Code,Rate,Order_Qty,Despatch_Qty,Active_Flag,Cust_Mtrl,Cust_DrgNo,Packing,Others,Excise_Duty,Cust_Drg_Desc,Tool_Cost,Authorized_flag,OpenSO,SalesTax_Type,PerValue,InternalSONo,RevisionNo,Remarks,MRP,Abantment_Code,AccessibleRateforMRP,Packing_Type,TOOL_AMOR_FLAG,ShowInAuth,ADD_Excise_Duty,HSNSACCODE,  ISHSNORSAC, CGSTTXRT_TYPE,SGSTTXRT_TYPE, UTGSTTXRT_TYPE,IGSTTXRT_TYPE,COMPENSATION_CESS   from Cust_ord_dtl where unit_code='" & gstrUNITID & "' and Item_code ='" & StrItemCode & "' and cust_drgNo ='" & strDrgNo & "' and cust_ref ='" & txtReferenceNo.Text & "' and amendment_no='" & Trim(txtAmendmentNo.Text) & "' and Active_flag ='A' and account_code='" & txtCustomerCode.Text & "'"
                    'Ends here
                    rsitem.GetResult(m_strSql)
                    If rsitem.GetNoRows > 0 Then
                        'change account Plug in
                        If rsitem.GetValue("OpenSO") = False Then
                            ssPOEntry.Row = ssPOEntry.MaxRows : ssPOEntry.Col = 1 : ssPOEntry.Col2 = 1 : ssPOEntry.Value = CheckState.Unchecked
                        Else
                            ssPOEntry.Row = ssPOEntry.MaxRows : ssPOEntry.Col = 1 : ssPOEntry.Col2 = 1 : ssPOEntry.Value = CheckState.Checked
                        End If
                        Call ssPOEntry.SetText(2, ssPOEntry.MaxRows, rsitem.GetValue("Cust_DrgNo"))
                        lblCustPartDesc.Text = rsitem.GetValue("Cust_Drg_desc")
                        Call ssPOEntry.SetText(4, ssPOEntry.MaxRows, rsitem.GetValue("Item_Code "))
                        Call ssPOEntry.SetText(5, ssPOEntry.MaxRows, rsitem.GetValue("Order_Qty"))
                        Call ssPOEntry.SetText(13, ssPOEntry.MaxRows, rsitem.GetValue("Rate"))
                        Call ssPOEntry.SetText(6, ssPOEntry.MaxRows, (rsitem.GetValue("Rate") * ctlPerValue.Text))
                        Call ssPOEntry.SetText(14, ssPOEntry.MaxRows, rsitem.GetValue("Cust_Mtrl"))
                        Call ssPOEntry.SetText(7, ssPOEntry.MaxRows, (rsitem.GetValue("Cust_Mtrl") * ctlPerValue.Text))
                        Call ssPOEntry.SetText(15, ssPOEntry.MaxRows, rsitem.GetValue("Tool_Cost"))
                        Call ssPOEntry.SetText(8, ssPOEntry.MaxRows, (rsitem.GetValue("Tool_Cost") * ctlPerValue.Text))
                        Call ssPOEntry.SetText(9, ssPOEntry.MaxRows, rsitem.GetValue("Packing"))
                        Call ssPOEntry.SetText(10, ssPOEntry.MaxRows, rsitem.GetValue("Excise_Duty"))
                        Call ssPOEntry.SetText(16, ssPOEntry.MaxRows, rsitem.GetValue("Others"))
                        Call ssPOEntry.SetText(11, ssPOEntry.MaxRows, (rsitem.GetValue("Others") * ctlPerValue.Text))
                        'added by nisha on 23/12/2002
                        Call ssPOEntry.SetText(17, ssPOEntry.MaxRows, rsitem.GetValue("Despatch_qty"))
                        '***
                        'GST DETAILS
                        Call ssPOEntry.SetText(21, ssPOEntry.MaxRows, rsitem.GetValue("CGSTTXRT_TYPE"))
                        Call ssPOEntry.SetText(22, ssPOEntry.MaxRows, rsitem.GetValue("SGSTTXRT_TYPE"))
                        Call ssPOEntry.SetText(23, ssPOEntry.MaxRows, rsitem.GetValue("UTGSTTXRT_TYPE"))
                        Call ssPOEntry.SetText(24, ssPOEntry.MaxRows, rsitem.GetValue("IGSTTXRT_TYPE"))
                        Call ssPOEntry.SetText(25, ssPOEntry.MaxRows, rsitem.GetValue("COMPENSATION_CESS"))
                        'GST DETAILS
                        Call ssSetFocus(ssPOEntry.MaxRows, 3)
                    End If
                End If

            End If
        End If
        lblCustPartDesc.Text = ""
        Call ToSetcustdrgDesc(e.newRow)
        Exit Sub            'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub ctlHeader_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlHeader.Click
        On Error GoTo ErrHandler
        Call ShowHelp("HLP" & Mid(Me.Name, 4, Len(Me.Name)) & ".htm") '("pddrep_listofitems.htm")
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub LockGrid()
        On Error GoTo ErrHandler
        Dim intGridCnt As Integer = 0
        With ssPOEntry
            For intGridCnt = 1 To ssPOEntry.MaxRows
                .Row = intGridCnt
                .Row2 = intGridCnt
                .Col = 1
                .Col2 = 1
                .BlockMode = True
                If .Value = System.Windows.Forms.CheckState.Checked Then
                    .BlockMode = False
                    .Row = intGridCnt
                    .Row2 = intGridCnt
                    .Col = 5
                    .Col2 = 5
                    .Lock = True
                    .BlockMode = True
                End If
                .BlockMode = False
            Next
            .Col = 7
            .Col2 = 7
            .Row = 1
            .Row2 = .MaxRows
            .BlockMode = True
            .Lock = False
            .BlockMode = False
        End With
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Function CheckForMultipleOpenSO(ByVal Account_Code As String, ByVal Cons_Code As String, ByVal Cust_ref As String, ByVal Cust_DrgNo As String, ByVal Item_code As String) As String
        Dim rstHelpDb As ClsResultSetDB
        Dim blnopenclosedso As Boolean

        Try
            rstHelpDb = New ClsResultSetDB
            If chkOpenSo.Checked Then
                blnopenclosedso = 1
            Else
                blnopenclosedso = 0
            End If
            Call rstHelpDb.GetResult("Select dbo.UDF_CHECK_ACTIVE_SO_ITEM('" & gstrUNITID & "' ,'" & Account_Code & "','" & Cons_Code & "','" & Cust_ref & "','" & Cust_DrgNo & "','" & Item_code & "','" & blnopenclosedso & "') as ActiveSalesOrder", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If rstHelpDb.GetNoRows >= 1 Then
                If Len(rstHelpDb.GetValue("ActiveSalesOrder")) > 0 Then
                    CheckForMultipleOpenSO = rstHelpDb.GetValue("ActiveSalesOrder")
                Else
                    CheckForMultipleOpenSO = ""
                End If
            End If
            rstHelpDb.ResultSetClose()
            rstHelpDb = Nothing
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Function

    '10736222-changes done for CT2 ARE 3 functionality - to enable CT2 Reqd flag
    Private Sub ChkCT2Reqd_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkCT2Reqd.CheckedChanged
        Dim intcounter As Short
        Try
            If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                If ChkCT2Reqd.Checked = True Then
                    txtECSSTaxType.Text = ""
                    txtECSSTaxType.Enabled = False
                    CmdECSSTaxType.Enabled = False
                    With ssPOEntry
                        If .MaxRows = 0 Then Exit Sub
                        If .MaxRows = 1 Then
                            .Col = 2
                            .Row = 1
                            If .Text.Trim.Length = 0 Then
                                Exit Sub
                            End If
                        End If
                    End With

                    If MessageBox.Show("Are you sure you want to enable CT2 Flag." & vbCr & "Enabling it will remove all the items in grid." & vbCr & "Please confirm ???", ResolveResString(100), MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                        ssPOEntry.MaxRows = 0
                        ADDRow()
                    Else
                        ChkCT2Reqd.Checked = False
                        txtECSSTaxType.Text = ""
                        txtECSSTaxType.Enabled = True
                        CmdECSSTaxType.Enabled = True
                    End If
                Else '10797956 
                    txtECSSTaxType.Text = ""
                    txtECSSTaxType.Enabled = True
                    CmdECSSTaxType.Enabled = True
                    With ssPOEntry
                        For intcounter = 1 To .MaxRows
                            .Row = intcounter : .Row2 = intcounter : .Col = 10 : .Col2 = 11 : .BlockMode = True : .Lock = False : .BlockMode = False
                        Next
                    End With
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    '10869290
    Private Sub cmdServiceTax_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdServiceTax.Click
        Dim StrSql As String = String.Empty
        Dim StrSrvCHelp() As String

        Try
            Select Case Me.cmdButtons.Mode
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                    StrSql = " SELECT TXRT_RATE_NO,TXRT_RATEDESC FROM GEN_TAXRATE WHERE UNIT_CODE='" & gstrUNITID & "' AND (TX_TAXEID='SRT')  AND ((ISNULL(DEACTIVE_FLAG,0) <> 1) OR (CAST(GETDATE() AS DATE) <= DEACTIVE_DATE))"
                    StrSrvCHelp = Me.ctlEMPHelpSOEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, StrSql, "Service Tax Code Help")
                    If UBound(StrSrvCHelp) <= 0 Then Exit Sub
                    If StrSrvCHelp(0) = "0" Then
                        Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK) : txtService.Text = "" : txtService.Focus() : Exit Sub
                    Else
                        txtService.Text = StrSrvCHelp(0)
                        lblservicedesc.Text = StrSrvCHelp(1)
                    End If
            End Select

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtService_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtService.TextChanged
        Try
            If Len(txtService.Text) = 0 Then
                lblservicedesc.Text = ""
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtService_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtService.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)

        Try
            Select Case KeyAscii
                Case System.Windows.Forms.Keys.Return
                    Select Case Me.cmdButtons.Mode
                        Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                            If Len(txtService.Text) > 0 Then
                                Call txtService_Validating(txtService, New System.ComponentModel.CancelEventArgs(False))
                            Else
                                chkOpenSo.Focus()
                            End If
                    End Select
                Case 39, 34, 96
                    KeyAscii = 0
            End Select
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtService_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtService.Validating

        Try
            If Len(txtService.Text) > 0 Then
                If CheckExistanceOfFieldData((txtService.Text), "TxRt_Rate_No", "Gen_TaxRate", " (TX_TAXEID='SRT') and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))") Then
                    Call FillLabel("SRT")
                    chkOpenSo.Focus()
                Else
                    Call ConfirmWindow(10248, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                    txtService.Text = ""
                    If txtService.Enabled Then txtService.Focus()
                End If
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtService_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtService.KeyUp
        Dim KeyCode As Short = e.KeyCode

        Try
            If KeyCode = 112 Then
                If cmdServiceTax.Enabled Then Call cmdServiceTax_Click(cmdServiceTax, New System.EventArgs())
            End If
            Exit Sub
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub cmbPOType_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles cmbPOType.Validating

        Try
            If cmbPOType.Text = "V-SERVICE" And (gblnGSTUnit = False Or gstrUNITID = "STH") Then
                cmdServiceTax.Enabled = True
                txtService.Enabled = True
                txtService.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                cmdSBCtax.Enabled = True
                txtSBC.Enabled = True
                txtSBC.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                cmdKKCtax.Enabled = True
                txtKKC.Enabled = True
                txtKKC.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            Else
                cmdServiceTax.Enabled = False
                txtService.Enabled = False
                txtService.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                cmdSBCtax.Enabled = False
                txtSBC.Enabled = False
                txtSBC.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                cmdKKCtax.Enabled = False
                txtKKC.Enabled = False
                txtKKC.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub cmdSBCtax_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSBCtax.Click
        Dim StrSql As String = String.Empty
        Dim StrSrvCHelp() As String

        Try
            Select Case Me.cmdButtons.Mode
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                    StrSql = " SELECT TXRT_RATE_NO,TXRT_RATEDESC FROM GEN_TAXRATE WHERE UNIT_CODE='" & gstrUNITID & "' AND (TX_TAXEID='SBC')  AND ((ISNULL(DEACTIVE_FLAG,0) <> 1) OR (CAST(GETDATE() AS DATE) <= DEACTIVE_DATE))"
                    StrSrvCHelp = ctlEMPHelpSOEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, StrSql, "SBC Tax Code Help")
                    If UBound(StrSrvCHelp) <= 0 Then Exit Sub
                    If StrSrvCHelp(0) = "0" Then
                        Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK) : txtSBC.Text = "" : txtSBC.Focus() : Exit Sub
                    Else
                        txtSBC.Text = StrSrvCHelp(0)
                        'lblservicedesc.Text = StrSrvCHelp(1)
                    End If
            End Select

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtSBC_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSBC.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)

        Try
            Select Case KeyAscii
                Case System.Windows.Forms.Keys.Return
                    Select Case Me.cmdButtons.Mode
                        Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                            If Len(txtSBC.Text) > 0 Then
                                Call txtSBC_Validating(txtSBC, New System.ComponentModel.CancelEventArgs(False))
                            Else
                                chkOpenSo.Focus()
                            End If
                    End Select
                Case 39, 34, 96
                    KeyAscii = 0
            End Select
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtSBC_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtSBC.Validating
        Try
            If Len(txtSBC.Text) > 0 Then
                If CheckExistanceOfFieldData((txtSBC.Text), "TxRt_Rate_No", "Gen_TaxRate", " (TX_TAXEID='SBC') and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))") Then
                    Call FillLabel("SBC")
                    chkOpenSo.Focus()
                Else
                    Call ConfirmWindow(10248, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                    txtSBC.Text = ""
                    If txtSBC.Enabled Then txtSBC.Focus()
                End If
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try

    End Sub

    Private Sub cmdKKCtax_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdKKCtax.Click
        Dim StrSql As String = String.Empty
        Dim StrSrvCHelp() As String

        Try
            Select Case Me.cmdButtons.Mode
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                    StrSql = " SELECT TXRT_RATE_NO,TXRT_RATEDESC FROM GEN_TAXRATE WHERE UNIT_CODE='" & gstrUNITID & "' AND (TX_TAXEID='KKC')  AND ((ISNULL(DEACTIVE_FLAG,0) <> 1) OR (CAST(GETDATE() AS DATE) <= DEACTIVE_DATE))"
                    StrSrvCHelp = ctlEMPHelpSOEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, StrSql, "KKC Tax Code Help")
                    If UBound(StrSrvCHelp) <= 0 Then Exit Sub
                    If StrSrvCHelp(0) = "0" Then
                        Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK) : txtKKC.Text = "" : txtKKC.Focus() : Exit Sub
                    Else
                        txtKKC.Text = StrSrvCHelp(0)
                        'lblservicedesc.Text = StrSrvCHelp(1)
                    End If
            End Select

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtKKC_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtKKC.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)

        Try
            Select Case KeyAscii
                Case System.Windows.Forms.Keys.Return
                    Select Case Me.cmdButtons.Mode
                        Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                            If Len(txtKKC.Text) > 0 Then
                                Call txtKKC_Validating(txtSBC, New System.ComponentModel.CancelEventArgs(False))
                            Else
                                chkOpenSo.Focus()
                            End If
                    End Select
                Case 39, 34, 96
                    KeyAscii = 0
            End Select
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub txtKKC_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtKKC.Validating
        Try
            If Len(txtKKC.Text) > 0 Then
                If CheckExistanceOfFieldData((txtKKC.Text), "TxRt_Rate_No", "Gen_TaxRate", " (TX_TAXEID='KKC') and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))") Then
                    Call FillLabel("KKC")
                    chkOpenSo.Focus()
                Else
                    Call ConfirmWindow(10248, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                    txtKKC.Text = ""
                    If txtKKC.Enabled Then txtKKC.Focus()
                End If
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try

    End Sub


    Private Sub txtAmendmentNo_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtAmendmentNo.Validating

    End Sub

    Private Sub chkShipAddress_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkShipAddress.CheckedChanged
        Dim StrSql As String = String.Empty
        Dim StrSrvCHelp() As String

        Try
            If chkShipAddress.Checked = True Then
                Select Case Me.cmdButtons.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        StrSql = "select Distinct Shipping_Code,Shipping_Desc,Ship_Address1,Ship_Address2,Ship_State,GSTIN_ID from Customer_Shipping_Dtl where unit_code='" & gstrUNITID & "' and InActive_Flag=0 and customer_code='" & Trim(txtCustomerCode.Text) & "'"
                        StrSrvCHelp = Me.ctlEMPHelpSOEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, StrSql, "Ship Address Code Help")
                        If UBound(StrSrvCHelp) <= 0 Then
                            chkShipAddress.Checked = False
                            Exit Sub
                        End If

                        If StrSrvCHelp(0) = "0" Then
                            Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK) : txtShipAddress.Text = "" : txtShipAddress.Focus() : Exit Sub
                        Else
                            txtShipAddress.Text = StrSrvCHelp(0)
                            '                            lblservicedesc.Text = StrSrvCHelp(1)
                            lblShipAddress_Details.Text = StrSrvCHelp(1)
                        End If
                End Select
            Else
                txtShipAddress.Text = ""
                lblShipAddress_Details.Text = ""
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub

    Private Sub BtnUpload_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnUpload.Click
        Try
            If txtReferenceNo.Text.Trim = "" Then
                Exit Sub
            End If
            Dim oFDialog As New OpenFileDialog()
            oFDialog.Filter = "pdf files(*.pdf)|*.pdf"
            oFDialog.FilterIndex = 3
            oFDialog.RestoreDirectory = True
            oFDialog.Title = "Select a file to upload"
            If oFDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                lblFilePath.Tag = ""
                lblFilePath.Text = ""
                lblFilePath.Tag = oFDialog.SafeFileName
                lblFilePath.Text = oFDialog.FileName
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub


    Private Sub btnViewDoc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnViewDoc.Click
        Dim strSQL As String = String.Empty
        Dim sqlRdr As SqlDataReader = Nothing
        Dim fileBytes() As Byte
        Dim strFileName As String = String.Empty
        Try
            If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Or UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
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

                    lblFilePath.Tag = sqlRdr("DOC_NAME")
                    strFileName = Path.GetTempPath() & lblFilePath.Tag.ToString()
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
            End If
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

    Private Sub btnRemoveDoc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemoveDoc.Click
        If txtReferenceNo.Text.Trim = "" Then
            Exit Sub
        End If
        Dim strSQL As String = String.Empty

        Try
            If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Or UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                lblFilePath.Tag = ""
                lblFilePath.Text = ""
                If txtAmendmentNo.Text.Trim <> "" Then
                    strSQL = "Delete FROM Cust_Ord_DOC_Dtl WHERE UNIT_CODE ='" & gstrUNITID & "' AND Account_Code ='" & txtCustomerCode.Text.Trim & "' AND Cust_Ref='" & txtReferenceNo.Text.Trim & "' AND ISACTIVE=1 AND Amendment_No='" & txtAmendmentNo.Text.Trim & "'"

                Else
                    strSQL = "Delete FROM Cust_Ord_DOC_Dtl WHERE UNIT_CODE ='" & gstrUNITID & "' AND Account_Code ='" & txtCustomerCode.Text.Trim & "' AND Cust_Ref='" & txtReferenceNo.Text.Trim & "' AND ISACTIVE=1"

                End If
                mP_Connection.Execute(strSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

                lblFilePath.Tag = ""
                lblFilePath.Text = ""
                btnViewDoc.Enabled = False
                MsgBox("Document Remove Successfully.", MsgBoxStyle.Information, ResolveResString(100))

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, ResolveResString(100), MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Sub

   
End Class