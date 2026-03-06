Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports CrystalDecisions.CrystalReports.Engine
Imports System.IO
Imports System.Data.SqlClient
Imports System.Text
Imports Microsoft.VisualBasic.Compatibility


Friend Class frmMKTTRN0001
    Inherits System.Windows.Forms.Form
    '---------------------------------------------------------------------------
    '(C) 2001 MIND, All rights reserved
    '
    'File Name          :   frmMKTTRN0001.frm.
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
    'chkOpenSo.value = 1 is replaced with chkOpenSo.value = vbChecked
    '---------------------------------------------------------------------------
    '---------------------------------------------------------------------------
    '18/12/2004
    'changes done by nisha for Future SO functionility changes
    '21/12/2004
    'changes done by nisha for to Allow Editable if It is Authorized.
    '28/12/2004
    'changes done by Brij for Addition of remarks Column
    '05/02/2005
    'changes done by Arshad to validate Tariff other than Job work case.
    '25/02/2005
    'changes done by nisha for so help bug correction
    '05/03/2005
    'Changes by Ravjeet, The query fetching SO No generating error coz no column name was specified for column three ie. InternalSONo
    '18/03/2005
    'Changes by Arshad, Search option for Customer Drawing No and Item Code and Forms selection
    'Revised By         :   Prashant Dhingra
    'Revision Date      :   14/06/2005
    'Issue Id           :   15024
    'Revised History    :   Control Button not visible
    '''******************************************************
    'Revision History   :    Changes done by Ashutosh on 10-10-2005 , Issue Id:15876, Bug fix of SO entry form , If Item is saved as closed one but it still saved as Open.
    '''******************************************************
    'Revised By         :   Jogender
    'Revision Date      :   31/05/2006
    'Issue Id           :   17975
    'Revised History    :   MRP & Abatment addition on the Form
    ''''''****************************************************
    'Revised By         :   Jogender
    'Revision Date      :   06/06/2006
    'Issue Id           :   18021
    'Revised History    :   Per unit Excise Price on the Form
    '---------------------------------------------------------------------------
    'Revised By         :   Manoj Kr. Vaish
    'Revision Date      :   20-Nov-2007
    'Issue Id           :   21551
    'Revised History    :   Add New Tax VAT with Sale Tax help
    '---------------------------------------------------------------------------
    'Revised By        -    Manoj Vaish
    'Revision Date     -    06 Aug 2008
    'Issue ID          -    eMpro-20080805-20745
    'Revision History  -    Rectifcation of .Net Conversion Issues
    '-----------------------------------------------------------------------------
    'Revised By        -    Manoj Vaish
    'Revision Date     -    27 Feb 2009
    'Issue ID          -    eMpro-20090227-27987
    'Revision History  -    Consignee Changes for commercial invoice at Mate Units
    '-----------------------------------------------------------------------------
    'Revised By        -    Manoj Vaish
    'Revision Date     -    20 Jul 2009
    'Issue ID          -    eMpro-20090720-33879
    'Revision History  -    Addition of Additional VAT 
    '-----------------------------------------------------------------------------
    'Revised By        -    Vinod Singh
    'Revision Date     -    25/04/2011
    'Revision History  -    Changes for Multi Unit
    '-----------------------------------------------------------------------------
    'Revised By        -    Prashant Rajpal
    'Revision Date     -    13 july 2011
    'Issue ID          -    10114335
    'Revision History  -    Sales order should be increased by 34 characters for HILEX only
    '-----------------------------------------------------------------------------
    ' Change By Deepak on 11-Oct-2011 for Support Change Management-----------------------
    '-----------------------------------------------------------------------------
    'Revised By        :    Prashant Rajpal
    'Issue ID          :    10117810
    'Revision Date     :    22 JULY 2011
    'History           :    Customer item description Wrong Saved. 
    '=======================================================================================
    'Modified by Roshan Singh on 09 nov 2011 for Multi Unit Change.
    '=======================================================================================
    'Revised By        :    Prashant Rajpal
    'Issue ID          :    10177787 
    'Revision Date     :    04 jan 2012
    'History           :    Add Vat Field not given correct data
    '=======================================================================================
    'Modified by Roshan Singh on 31 JAN 2012 for Multi Unit Change.
    '=======================================================================================
    'Revised By        :    Shubhra Verma
    'Issue ID          :    10193979 
    'Revision Date     :    13 Feb 2012
    'History           :    Inactive Part Numbers are appearing in Sales Order Entry Form.
    '=======================================================================================
    '=======================================================================================
    'Modified by Virendra Gupta on 14 Feb 2012 for Multi Unit Change.
    '=======================================================================================
    'Revised By        :    Prashant Rajpal
    'Issue ID          :    10239863 
    'Revision Date     :    21 june 2012
    'History           :    changes for Trading Invoice 
    '=======================================================================================
    'REVISED BY         :   SHUBHRA VERMA
    'REVISED ON         :   12 MAR 2012
    'ISSUE ID           :   10354980  
    'DESCRIPTION        :   VALIDATION ADDED, SO THAT "SO" WILL NOT GET SAVED WITHOUT ITEM DETAILS
    '***************************************************************************************
    'REVISED BY         :   PRASHANT RAJPAL
    'REVISED ON         :   21 MAY 2013
    'ISSUE ID           :   10390572 
    'DESCRIPTION        :   CHANGES FOR SHINDENGEN RELATED TO SALES TAX HELP..
    '***************************************************************************************
    'REVISED BY         :   PRASHANT RAJPAL
    'ISSUE ID           :   10439883
    'DESCRIPTION        :   CHANGES ON DISCOUNT INVOICE FUNCTIONLITY 
    'REVISION DATE      :    25 -AUG 2013-03 SEP 2013
    '***************************************************************************************
    'REVISED BY         :   PRASHANT RAJPAL
    'ISSUE ID           :   10515727
    'DESCRIPTION        :   ISSUE IN OPEN SALES ORDER ,QUANTITY COLUMN WRONGLY APPEARED IN GRID 
    'REVISION DATE      :   14 JAN 2014
    '***************************************************************************************
    'REVISED BY         :   PRASHANT RAJPAL
    'ISSUE ID           :   10561117  
    'DESCRIPTION        :   MAE MULTI UNIT CHANGES 
    'REVISION DATE      :   19 MAR 2014
    '***************************************************************************************
    'REVISED BY         :   PRASHANT RAJPAL
    'ISSUE ID           :   10610274   
    'DESCRIPTION        :   HILEX- SALESORDERPREVIEW:SOMETIMES BLANK REPORT
    'REVISION DATE      :   10 JUNE2014
    '***************************************************************************************
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
    'REVISED BY         :   Abhinav Kumar
    'ISSUE ID           :   10869290 
    'DESCRIPTION        :   eMPro- Service Invoice Functionality
    'REVISION DATE      :   21 Sep 2015
    '***************************************************************************************
    'REVISED BY         :   Prashant Rajpal
    'ISSUE ID           :   10940008
    'DESCRIPTION        :   Service Invoice Issue 
    'REVISION DATE      :   02-Dec-2015
    '***************************************************************************************
    'REVISED BY         :   Parveen Kumar
    'ISSUE ID           :   10844039
    'DESCRIPTION        :   Change in SO Authorization. 
    'REVISION DATE      :   18-May-2016
    '***************************************************************************************
    'REVISED BY         :   PRASHANT RAJPAL
    'ISSUE ID           :   
    'DESCRIPTION        :   GST ISSUE
    'REVISION DATE      :   18-May-2016

    'REVISED BY         :   ASHISH SHARMA
    'ISSUE ID           :   101188073
    'DESCRIPTION        :   GST For EOU 
    'REVISION DATE      :   11 JUL 2017

    'REVISED BY         :   ABHIJIT KUMAR SINGH
    'ISSUE ID           :   101342142
    'DESCRIPTION        :   SHIP ADDRESS FIELD ADDTION 
    'REVISION DATE      :   21 AUG 2017

    'REVISED BY         :   AMIT KUMAR RANA
    'ISSUE ID           :   0
    'DESCRIPTION        :   OPTIMIZATION DONE AS WHILE EDIT UPDATED IT WAS TAKING 45 MINUTS FOR 700+ ITEMS 
    'REVISION DATE      :   02 APR 2022
    '***************************************************************************************

    Dim datatable_MasterData As DataTable
    Dim datatable_MasterData_GEN_TAXRATE As DataTable
    Dim datatable_ExistingSoDetails As DataTable

    Dim m_blnHelpFlag, m_blnCloseFlag As Boolean
    Dim rsdb As New ClsResultSetDB
    Dim mintFormIndex, intRow As Short
    Dim mstrCode As String
    Dim m_Item_Code As String
    Dim m_blnChangeFormFlg As Boolean
    Dim mvalid As Boolean
    Dim m_strSql, strSql As String
    Dim m_blnGetAmendmentDetails As Boolean
    Dim rsRefNo As New ClsResultSetDB
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
    Public mstrFormDetails As String
    Dim mblnDiscountFunctionality As Boolean
    Dim mblnDiscountMandatory As Boolean
    '10561117  
    Dim mblnMarkupFunctionality As Boolean
    Dim mblnpackingdefined As Boolean = False
    '10561117  
    'ISSUE ID : 10763705
    Dim mblnappendsoitem_customer As Boolean = False
    '10940008
    Dim mblnServiceInvoicemate As Boolean
    '10940008
    Dim MBLNALLOW_MULTIPLEHSNITEMS_CUSTOMER As Boolean = False
    Dim _blnEOUFlag As Boolean = False
    Dim strexportType As String
    Dim mblnIsExecuting = False
    Dim blnNoneditableCreditTerms_onSO As Boolean = False
    Dim mblnCADM_SALES_ORDER_CREATION As Boolean
    Dim blnOneCustDrgNo_MutipleItem As Boolean = False
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
            'strSql = "SELECT top 1 Party_Code FROM DBO.FIN_GET_RELATED_PARTIES ('" & gstrUNITID & "') WHERE CUST_VEND='C' and Party_Code='" & Trim(txtCustomerCode.Text) & "'"
            'If IsRecordExists(strSql) = True Then
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
            'ISSUE ID 10239863
            If UCase(Trim(cmbPOType.Text)) = "Q-TRADING" Then
                MsgBox("Open Sales order Not possible for Trading Type SO", MsgBoxStyle.OkOnly, ResolveResString(100))
                chkOpenSo.Checked = False
                Exit Sub
            End If
            'ISSUE ID 10239863
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
    Private Sub cmbPOType_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles cmbPOType.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        If cmbPOType.Text = "MRP-SPARES" Then
            ssPOEntry.BlockMode = True
            ssPOEntry.Col = 19
            ssPOEntry.Col2 = 19
            ssPOEntry.ColHidden = False
            ssPOEntry.TypeFloatMin = "0.0000"
            ssPOEntry.Col = 20
            ssPOEntry.Col2 = 20
            ssPOEntry.ColHidden = False
            ssPOEntry.BlockMode = False
            ssPOEntry.Col = 21
            ssPOEntry.Col2 = 21
            ssPOEntry.ColHidden = False
            ssPOEntry.TypeFloatMin = "0.0000"
            'ISSUE ID 10239863
        ElseIf cmbPOType.Text = "Q-TRADING" Then
            ssPOEntry.BlockMode = True
            ssPOEntry.Col = 10
            ssPOEntry.Col2 = 10
            ssPOEntry.Row = 1
            ssPOEntry.Row2 = ssPOEntry.MaxRows
            ssPOEntry.Text = ""
            ssPOEntry.BlockMode = False
            'ISSUE ID 10239863
        ElseIf cmbPOType.Text = "V-SERVICE" And gblnGSTUnit = False Then
            '10869290
            cmdServiceTax.Enabled = True
            txtService.Enabled = True
            txtService.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            cmdSBCtax.Enabled = True
            txtSBC.Enabled = True
            txtSBC.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            cmdKKCtax.Enabled = True
            txtKKC.Enabled = True
            txtKKC.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            ssPOEntry.BlockMode = True
            ssPOEntry.Col = 19
            ssPOEntry.Col2 = 19
            ssPOEntry.ColHidden = True
            ssPOEntry.Col = 20
            ssPOEntry.Col2 = 20
            ssPOEntry.ColHidden = True
            ssPOEntry.Col = 21
            ssPOEntry.Col2 = 21
            ssPOEntry.ColHidden = True
            ssPOEntry.BlockMode = False

        Else
            cmdServiceTax.Enabled = False
            txtService.Enabled = False
            txtService.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            cmdSBCtax.Enabled = False
            txtSBC.Enabled = False
            txtSBC.Text = ""
            txtSBC.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            cmdKKCtax.Enabled = False
            txtKKC.Enabled = False
            txtKKC.Text = ""
            txtKKC.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            ssPOEntry.BlockMode = True
            ssPOEntry.Col = 19
            ssPOEntry.Col2 = 19
            ssPOEntry.ColHidden = True
            ssPOEntry.Col = 20
            ssPOEntry.Col2 = 20
            ssPOEntry.ColHidden = True
            ssPOEntry.Col = 21
            ssPOEntry.Col2 = 21
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
    Private Sub cmdchangetype_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdchangetype.Click
        On Error GoTo ErrHandler
        Dim strsalesTerms As String
        Dim rssalesTerms As ClsResultSetDB
        Dim strSql As String
        If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Or cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            strSql = "select a.*,b.* from cust_ord_hdr a,cust_ord_dtl b where a.unit_code=b.unit_code and a.Amendment_No=b.Amendment_No and a.cust_Ref=b.Cust_Ref and a.Account_Code=b.Account_Code and a.unit_code='" & gstrUNITID & "' and a.Cust_ref='" & txtReferenceNo.Text & "' and a.amendment_No='" & txtAmendmentNo.Text & "'"
        ElseIf cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            strSql = "select * from cust_ord_hdr  where unit_code='" & gstrUNITID & "' and Cust_ref='" & txtReferenceNo.Text & "' and account_code='" & txtCustomerCode.Text & "'"
        End If
        m_pstrSql = strSql
        rssalesTerms = New ClsResultSetDB
        strsalesTerms = "Select Description From SaleTerms_Mst Where SaleTerms_Type ='PY' and unit_code='" & gstrUNITID & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date)) "
        rssalesTerms.GetResult(strsalesTerms, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        If rssalesTerms.GetNoRows > 0 Then
            strsalesTerms = "Select Description From SaleTerms_Mst Where SaleTerms_Type ='PR' and unit_code='" & gstrUNITID & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
            rssalesTerms.GetResult(strsalesTerms, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            If rssalesTerms.GetNoRows > 0 Then
                strsalesTerms = "Select Description From SaleTerms_Mst Where SaleTerms_Type ='PK' and unit_code='" & gstrUNITID & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date)) "
                rssalesTerms.GetResult(strsalesTerms, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                If rssalesTerms.GetNoRows > 0 Then
                    strsalesTerms = "Select Description From SaleTerms_Mst Where SaleTerms_Type ='FR' and unit_code='" & gstrUNITID & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date)) "
                    rssalesTerms.GetResult(strsalesTerms, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                    If rssalesTerms.GetNoRows > 0 Then
                        strsalesTerms = "Select Description From SaleTerms_Mst Where SaleTerms_Type ='TR' and unit_code='" & gstrUNITID & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date)) "
                        rssalesTerms.GetResult(strsalesTerms, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                        If rssalesTerms.GetNoRows > 0 Then
                            strsalesTerms = "Select Description From SaleTerms_Mst Where SaleTerms_Type ='OC' and unit_code='" & gstrUNITID & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
                            rssalesTerms.GetResult(strsalesTerms, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                            If rssalesTerms.GetNoRows > 0 Then
                                strsalesTerms = "Select Description From SaleTerms_Mst Where SaleTerms_Type ='MO' and unit_code='" & gstrUNITID & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
                                rssalesTerms.GetResult(strsalesTerms, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                                If rssalesTerms.GetNoRows > 0 Then
                                    strsalesTerms = "Select Description From SaleTerms_Mst Where SaleTerms_Type ='DL' and unit_code='" & gstrUNITID & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
                                    rssalesTerms.GetResult(strsalesTerms, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
                                    If rssalesTerms.GetNoRows > 0 Then
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
                                        Exit Sub
                                    End If
                                Else
                                    Call ConfirmWindow(10480, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                                    cmdchangetype.Focus()
                                    Exit Sub
                                End If
                            Else
                                Call ConfirmWindow(10480, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                                cmdchangetype.Focus()
                                Exit Sub
                            End If
                        Else
                            Call ConfirmWindow(10480, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            cmdchangetype.Focus()
                            Exit Sub
                        End If
                    Else
                        Call ConfirmWindow(10480, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                        cmdchangetype.Focus()
                        Exit Sub
                    End If
                Else
                    Call ConfirmWindow(10480, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                    cmdchangetype.Focus()
                    Exit Sub
                End If
            Else
                Call ConfirmWindow(10480, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                cmdchangetype.Focus()
                Exit Sub
            End If
        Else
            Call ConfirmWindow(10480, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
            cmdchangetype.Focus()
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
    Private Sub cmdForms_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdForms.Click
        frmMKTTRN0041.mstrFormDetails = mstrFormDetails
        frmMKTTRN0041.Cust_Code = txtCustomerCode.Text
        frmMKTTRN0041.ParentForm_Renamed = "FRMMKTTRN0001"
        frmMKTTRN0041.Left = cmdForms.Left
        frmMKTTRN0041.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(cmdForms.Top) + VB6.PixelsToTwipsY(cmdForms.Height) + 1000)
        frmMKTTRN0041.ShowDialog()
        If Len(Trim(mstrFormDetails)) = 0 Then
            SetFormButtonStyle(True)
        Else
            SetFormButtonStyle()
        End If
    End Sub
    Private Sub cmdHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHelp.Click
        Dim Index As Short = cmdHelp.GetIndex(eventSender)
        On Error GoTo ErrHandler
        Dim varRetVal As Object
        Dim strhelp1() As String
        Dim strmessage As String
        Dim RSCHECK As ClsResultSetDB
        On Error GoTo ErrHandler
        Dim strAmend, strString As String
        varRetVal = Nothing
        Select Case Index
            '*****Customer Code Help
            Case 0
                Select Case cmdButtons.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        With Me.txtCustomerCode
                            If Len(.Text) = 0 Then
                                varRetVal = ShowList(1, .MaxLength, "", "Customer_Code", "Cust_Name", "Customer_Mst", "and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))")
                                If varRetVal = "-1" Then
                                    Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                                    .Text = ""
                                Else
                                    .Text = varRetVal
                                    .Focus()
                                    '' added by priti on 11 May 2020 for OPEN SO Issue
                                    'strSql = "SELECT top 1 Party_Code FROM DBO.FIN_GET_RELATED_PARTIES ('" & gstrUNITID & "') WHERE CUST_VEND='C' and Party_Code='" & Trim(txtCustomerCode.Text) & "'"
                                    'If IsRecordExists(strSql) = True Then
                                    '    MsgBox("THIS CUSTOMER IS RELATED PARTIES.OPEN SO IS NOT ALLOWED FOR RELATED PARTIES. ", MsgBoxStyle.Information, ResolveResString(100))
                                    'End If
                                    '' code ends by priti on 11 May 2020 for OPEN SO Issue
                                End If
                            Else
                                varRetVal = ShowList(1, .MaxLength, .Text, "Customer_code", "Cust_Name", "Customer_Mst", "and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))")
                                If varRetVal = "-1" Then
                                    Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                                    .Text = ""
                                Else
                                    .Text = varRetVal
                                    '' added by priti on 11 May 2020 for OPEN SO Issue
                                    'strSql = "SELECT top 1 Party_Code FROM DBO.FIN_GET_RELATED_PARTIES ('" & gstrUNITID & "') WHERE CUST_VEND='C' and Party_Code='" & Trim(txtCustomerCode.Text) & "'"
                                    'If IsRecordExists(strSql) = True Then
                                    '    MsgBox("THIS CUSTOMER IS RELATED PARTIES.OPEN SO IS NOT ALLOWED FOR RELATED PARTIES. ", MsgBoxStyle.Information, ResolveResString(100))
                                    'End If
                                    '' code ends by priti on 11 May 2020 for OPEN SO Issue
                                End If
                            End If
                        End With
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT, UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                        With Me.txtCustomerCode
                            If Len(.Text) = 0 Then
                                varRetVal = ShowList(1, .MaxLength, "", "a.Account_Code", "b.Cust_Name", "Cust_ord_hdr a,Customer_mst b", "and a.unit_code=b.unit_code and a.Account_code = b.Customer_code and ((isnull(b.deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= b.deactive_date))", , , , , , "a.unit_code")
                                If varRetVal = "-1" Then
                                    Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                                    .Text = ""
                                Else
                                    .Text = varRetVal
                                    '' added by priti on 11 May 2020 for OPEN SO Issue
                                    ' strSql = "SELECT top 1 Party_Code FROM DBO.FIN_GET_RELATED_PARTIES ('" & gstrUNITID & "') WHERE CUST_VEND='C' and Party_Code='" & Trim(txtCustomerCode.Text) & "'"
                                    'If IsRecordExists(strSql) = True Then
                                    '   MsgBox("THIS CUSTOMER IS RELATED PARTIES.OPEN SO IS NOT ALLOWED FOR RELATED PARTIES. ", MsgBoxStyle.Information, ResolveResString(100))
                                    'End If
                                    '' code ends by priti on 11 May 2020 for OPEN SO Issue
                                End If
                            Else
                                varRetVal = ShowList(1, .MaxLength, .Text, "a.Account_Code", "b.Cust_Name", "Cust_ord_hdr a,Customer_mst b", "and a.unit_code=b.unit_code and a.Account_code =b.Customer_code and ((isnull(b.deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= b.deactive_date))", , , , , , "a.unit_code")
                                If varRetVal = "-1" Then
                                    Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                                    .Text = ""
                                Else
                                    .Text = varRetVal
                                    '' added by priti on 11 May 2020 for OPEN SO Issue
                                    'strSql = "SELECT top 1 Party_Code FROM DBO.FIN_GET_RELATED_PARTIES ('" & gstrUNITID & "') WHERE CUST_VEND='C' and Party_Code='" & Trim(txtCustomerCode.Text) & "'"
                                    'If IsRecordExists(strSql) = True Then
                                    '    MsgBox("THIS CUSTOMER IS RELATED PARTIES.OPEN SO IS NOT ALLOWED FOR RELATED PARTIES. ", MsgBoxStyle.Information, ResolveResString(100))
                                    'End If
                                    ' '' code ends by priti on 11 May 2020 for OPEN SO Issue
                                End If
                            End If
                            .Focus()
                        End With
                End Select
                '*****Ref. No. help
            Case 1
                Select Case cmdButtons.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        With Me.txtReferenceNo
                            RSCHECK = New ClsResultSetDB
                            Call RSCHECK.GetResult("SELECT TOP 1 1 FROM CUSTOMER_MST WHERE unit_code='" & gstrUNITID & "' and CUSTOMER_CODE = '" & txtCustomerCode.Text.ToString.Trim & "' and INVOICEAGAINSTAGREEMENTMST = '1' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                            If RSCHECK.GetNoRows > 0 Then
                                strhelp1 = ctlEMPHelp1.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "SELECT Doc_no,Customer_code,InternalAssyPart,CustAssyPart FROM AGREEMENT_HDR H WHERE UNIT_CODE='" & gstrUNITID & "' AND NOT EXISTS (SELECT TOP 1 1 FROM CUST_ORD_HDR I WHERE H.UNIT_CODE=I.UNIT_CODE AND  H.CUSTOMER_CODE = I.ACCOUNT_CODE AND CONVERT(VARCHAR(50),H.DOC_NO) = I.CUST_REF AND I.unit_code='" & gstrUNITID & "' ) AND ACTIVE = '1' ")
                            Else
                                If Len(.Text) = 0 Then
                                    strhelp1 = ctlEMPHelp1.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select distinct Cust_Ref,'Order_date' = " & DateColumnNameInShowList("Order_date") & ",isnull(InternalSONo,'') as InternalSoNo from cust_ord_hdr where unit_code='" & gstrUNITID & "' and Account_Code='" & Me.txtCustomerCode.Text & "' and Active_Flag='A'")
                                Else
                                    strhelp1 = ctlEMPHelp1.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select distinct Cust_Ref,'Order_date' = " & DateColumnNameInShowList("Order_date") & ",isnull(InternalSONo,'') as InternalSoNo from cust_ord_hdr where unit_code='" & gstrUNITID & "' and Account_Code='" & Me.txtCustomerCode.Text & "' and Active_Flag='A' and cust_ref= '" & Me.txtReferenceNo.Text & "'")
                                End If
                            End If
                            If Not (UBound(strhelp1) = -1) Then
                                If (Len(strhelp1(0)) >= 1) And strhelp1(0) = "0" Then
                                    If (Len(LTrim(RTrim(Me.txtReferenceNo.Text))) > 0) Then
                                        strmessage = "No job Order  are Defined with Prefix [" & Me.txtReferenceNo.Text & "]"
                                        strmessage = strmessage & vbCrLf & "To View the list, Clear the Text and try Again."
                                        MsgBox(strmessage, MsgBoxStyle.Information, ResolveResString(100))
                                    Else
                                        MsgBox("No job order are Defined", MsgBoxStyle.Information, ResolveResString(100))
                                    End If
                                    Me.txtReferenceNo.Focus()
                                    Exit Sub
                                Else
                                    Me.txtReferenceNo.Text = strhelp1(0)
                                End If
                            End If
                            .Focus()
                        End With
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT, UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                        With Me.txtReferenceNo
                            If Len(.Text) = 0 Then
                                strhelp1 = ctlEMPHelp1.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select Cust_Ref,'Order_date' = " & DateColumnNameInShowList("Order_date") & ",InternalSONo from cust_ord_hdr where unit_code='" & gstrUNITID & "' and Account_Code='" & Me.txtCustomerCode.Text & "'  and Amendment_No ='' ")
                            Else
                                strhelp1 = ctlEMPHelp1.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select Cust_Ref,'Order_date' = " & DateColumnNameInShowList("Order_date") & ",InternalSONo from cust_ord_hdr where unit_code='" & gstrUNITID & "' and Account_Code='" & Me.txtCustomerCode.Text & "' and Amendment_No ='' and cust_ref= '" & Me.txtReferenceNo.Text & "'")
                            End If
                            If Not (UBound(strhelp1) = -1) Then
                                If (Len(strhelp1(0)) >= 1) And strhelp1(0) = "0" Then
                                    If (Len(LTrim(RTrim(Me.txtReferenceNo.Text))) > 0) Then
                                        strmessage = "No Reference No. are Defined with Prefix [" & Me.txtReferenceNo.Text & "]"
                                        strmessage = strmessage & vbCrLf & "To View the list, Clear the Text and try Again."
                                        MsgBox(strmessage, MsgBoxStyle.Information, ResolveResString(100))
                                    Else
                                        MsgBox("No Reference No. are Defined", MsgBoxStyle.Information, ResolveResString(100))
                                    End If
                                    Me.txtReferenceNo.Focus()
                                    Exit Sub
                                Else
                                    Me.txtReferenceNo.Text = strhelp1(0)
                                End If
                            End If
                            .Focus()
                        End With
                End Select
                '******Currency Code Help
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
                        If txtAmendmentNo.Text <> "" Then strAmend = "where unit_code='" & gstrUNITID & "' and cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "' and Account_Code='" & txtCustomerCode.Text & "' and Amendment_No like '" & strString & "' " Else strAmend = " where cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'and Account_Code='" & txtCustomerCode.Text & "'"
                        strhelp1 = ctlEMPHelpSOEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "SELECT Amendment_No,'Amendment_date'=" & DateColumnNameInShowList("Amendment_date") & ", Cust_Ref  FROM cust_ord_hdr " & strAmend & " ", "List of All Amendments", CStr(1))
                        If Not (UBound(strhelp1) = -1) Then
                            If (Len(strhelp1(0)) >= 1) And strhelp1(0) = "0" Then
                                Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                                .Text = ""
                            Else
                                .Text = strhelp1(0)
                            End If
                        End If
                        .Focus()
                    End With
                End If
                '*****STax Help
            Case 4
                With txtSTax
                    If Len(.Text) = 0 Then
                        If UCase(Trim(cmbPOType.Text)) = "Q-TRADING" Then
                            varRetVal = ShowList(1, .MaxLength, "", "txrt_Rate_No", "txrt_rateDesc", "Gen_TaxRate", "and tx_TaxeID in ('LSTT','CSTT','VATT') and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date)) ")
                        Else
                            varRetVal = ShowList(1, .MaxLength, "", "txrt_Rate_No", "txrt_rateDesc", "Gen_TaxRate", "and tx_TaxeID in ('LST','CST','VAT') and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date)) ")
                        End If
                        If varRetVal = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            .Text = ""
                        Else
                            .Text = varRetVal
                        End If
                    Else
                        If UCase(Trim(cmbPOType.Text)) = "Q-TRADING" Then
                            varRetVal = ShowList(1, .MaxLength, .Text, "txrt_Rate_No", "txrt_rateDesc", "Gen_TaxRate", "and tx_TaxeID in ('LSTT','CSTT','VATT') and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date)) ")
                        Else
                            varRetVal = ShowList(1, .MaxLength, .Text, "txrt_Rate_No", "txrt_rateDesc", "Gen_TaxRate", "and tx_TaxeID in ('LST','CST','VAT') and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date)) ")
                        End If
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
                        varRetVal = ShowList(1, .MaxLength, "", "crtrm_termID", "crTrm_desc", "Gen_CreditTrmMaster", "and crtrm_status =1")
                        If varRetVal = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            .Text = ""
                        Else
                            .Text = varRetVal
                        End If
                    Else
                        varRetVal = ShowList(1, .MaxLength, .Text, "crtrm_termID", "crTrm_desc", "Gen_CreditTrmMaster", "and crtrm_status =1")
                        If varRetVal = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            .Text = ""
                        Else
                            .Text = varRetVal
                        End If
                    End If
                    .Focus()
                End With
                '***
                '****added by Ajay on 18/07/2003
                '1.Surcharge on S.Tax
            Case 6
                With txtSChSTax
                    If Len(.Text) = 0 Then
                        varRetVal = ShowList(1, .MaxLength, "", "txrt_Rate_No", "txrt_rateDesc", "Gen_TaxRate", "and tx_TaxeID ='SST' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))")
                        If varRetVal = "-1" Then
                            Call ConfirmWindow(10010, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            .Text = ""
                        Else
                            .Text = varRetVal
                        End If
                    Else
                        varRetVal = ShowList(1, .MaxLength, .Text, "TxRt_Rate_no", "TxRt_RateDesc", "Gen_TaxRate", "and tx_TaxeID ='SST' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))")
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
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub cmdHelp_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles cmdHelp.MouseDown
        Dim Button As Short = eventArgs.Button \ &H100000
        Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
        Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
        Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
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
    Private Sub ctlHeader_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ctlHeader.Click
        On Error GoTo ErrHandler
        Call ShowHelp("HLP" & Mid(Me.Name, 4, Len(Me.Name)) & ".htm") '("pddrep_listofitems.htm")
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub ctlPerValue_Change(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        Dim intLoopCounter As Short
        Dim rsSalesParameter As New ClsResultSetDB
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
                .Col = 11 : .Text = "Others (Per Unit)"
            End If
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
        End With
    End Sub
    Private Sub DTAmendmentDate_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles DTAmendmentDate.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        System.Windows.Forms.Application.DoEvents()
        If DTAmendmentDate.Value > DTValidDate.Value Then
            Call ConfirmWindow(10146, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
            blnValidAmendDate = False
            Cancel = False
            Me.DTEffectiveDate.Focus()
            GoTo EventExitSub
        Else
            blnValidAmendDate = True
        End If
        If DTAmendmentDate.Value < DTDate.Value Then
            Call ConfirmWindow(10147, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
            blnValidAmendDate = False
            Cancel = True
            Me.DTEffectiveDate.Focus()
            GoTo EventExitSub
        Else
            blnValidAmendDate = True
        End If
        If DTAmendmentDate.Value > GetServerDate() Then
            MsgBox("Amendment Date Can not be Greater Than Current Date", MsgBoxStyle.Information, ResolveResString(100))
            blnValidAmendDate = False
            Cancel = True
            Me.DTEffectiveDate.Focus()
            GoTo EventExitSub
        Else
            blnValidAmendDate = True
        End If

        If GetPlantName() = "HILEX" And DTAmendmentDate.Value < GetServerDate() Then
            MsgBox("Amendment Date Can not be less Than Current Date", MsgBoxStyle.Information, ResolveResString(100))
            blnValidAmendDate = False
            Cancel = True
            Me.DTEffectiveDate.Focus()
            GoTo EventExitSub
        Else
            blnValidAmendDate = True
        End If


        GoTo EventExitSub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        GoTo EventExitSub
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub DTEffectiveDate_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles DTEffectiveDate.Leave
        On Error GoTo ErrHandler
        If DTEffectiveDate.Value < DTDate.Value Then
            MsgBox("Effective Date Cannot Be Less than SO Date", MsgBoxStyle.Information, ResolveResString(100))
            DTEffectiveDate.Value = DTDate.Value
            DTEffectiveDate.Focus()
            Exit Sub
        End If
        If DTEffectiveDate.Value > DateValue(CStr(FinancialYearDates(eMPowerFunctions.FinancialYearDatesEnum.DATE_END))) Then
            MsgBox("Effective Date Cannot Be Greater than Financial End Date", MsgBoxStyle.Information, ResolveResString(100))
            DTEffectiveDate.Value = DTDate.Value
            DTEffectiveDate.Focus()
            Exit Sub
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub DTValidDate_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles DTValidDate.Leave
        Call DTValidDate_Validating(DTValidDate, New System.ComponentModel.CancelEventArgs(False))
    End Sub
    Private Sub frmMKTTRN0001_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        On Error GoTo ErrHandler
        'Make the node normal font
        frmModules.NodeFontBold(Tag) = False
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0001_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Or cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
            If KeyAscii = System.Windows.Forms.Keys.Escape Then
                Call cmdButtons_ButtonClick(cmdButtons, New UCActXCtl.UCbtndgrp.ButtonClickEventArgs(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL))
            End If
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
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
            If mvalid = False Then
                If txtAmendReason.Enabled Then
                    txtAmendReason.Focus()
                Else
                    cmdButtons.Focus()
                End If
            End If
        End If
        mvalid = False
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
    Sub SetFormButtonStyle(Optional ByRef blnBold As Boolean = False)
        On Error GoTo ErrHandler
        Dim gobjdb As New ClsResultSetDB
        If blnBold = True Then
            cmdForms.Font = VB6.FontChangeBold(cmdForms.Font, False)
            Exit Sub
        End If
        If Len(Trim(mstrFormDetails)) <> 0 Then
            cmdForms.Font = VB6.FontChangeBold(cmdForms.Font, True)
        Else
            gobjdb.GetResult("Select Form_type from forms_DTL where unit_code='" & gstrUNITID & "' and Doc_Type='9998' and PO_NO='" & Trim(txtReferenceNo.Text) & "' and Amendment_No='" & Trim(txtAmendmentNo.Text) & "' and Account_Code='" & Trim(txtCustomerCode.Text) & "'")
            If Not gobjdb.EOFRecord Then
                cmdForms.Font = VB6.FontChangeBold(cmdForms.Font, True)
            Else
                cmdForms.Font = VB6.FontChangeBold(cmdForms.Font, False)
            End If
        End If
        gobjdb.ResultSetClose()
        gobjdb = Nothing
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
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
                If txtAddVAT.Enabled = True Then txtAddVAT.Focus()
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
            rsdb.GetResult(m_strSql)
            If rsdb.GetNoRows > 0 Then
                Call FillLabel("CREDIT")
            Else
                cmdButtons.Focus()
                MsgBox("Entered Credit Term does not exist", MsgBoxStyle.Information, ResolveResString(100))
                txtCreditTerms.Text = ""
                txtCreditTerms.Focus()
            End If
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
        Dim rsdb As New ClsResultSetDB
        If Trim(txtCurrencyType.Text) = "" Or Len(Trim(txtCurrencyType.Text)) = 0 Then
            GoTo EventExitSub
        End If
        m_strSql = "Select TOP 1 1 from Currency_mst where unit_code='" & gstrUNITID & "' and Currency_code = '" & txtCurrencyType.Text & "'"
        Call rsdb.GetResult(m_strSql)
        If rsdb.GetNoRows = 0 Then
            Call ConfirmWindow(10144, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
            blnValidCurrency = False
            Cancel = True
            GoTo EventExitSub
        Else
            Call SSMaxLength()
        End If
        blnValidCurrency = True
        rsdb.ResultSetClose()
        GoTo EventExitSub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        GoTo EventExitSub
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtCustomerCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCustomerCode.TextChanged
        On Error GoTo ErrHandler
        Call FillLabel("CUSTOMER")
        Call FillLabel("CURRENCY")
        If Len(Trim(txtCustomerCode.Text)) <> 0 Then
            m_strSql = "Select TOP 1 1 from cust_Ord_hdr where unit_code='" & gstrUNITID & "' and account_code ='" & txtCustomerCode.Text & "'"
            rsdb.GetResult(m_strSql)
            If rsdb.GetNoRows > 0 Then
                txtReferenceNo.Enabled = True
                txtReferenceNo.BackColor = System.Drawing.Color.White
                cmdHelp(1).Enabled = True
            Else
                txtReferenceNo.Enabled = False
                txtReferenceNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                cmdHelp(1).Enabled = False
            End If
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
            Call txtCustomerCode_Leave(txtCustomerCode, New System.EventArgs())
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
    Private Sub TxtCustomerCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCustomerCode.KeyPress
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
    Private Sub txtCustomerCode_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCustomerCode.Leave
        On Error GoTo ErrHandler
        mvalid = False
        If Len(Trim(txtCustomerCode.Text)) <> 0 Then
            Select Case cmdButtons.Mode
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT, UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                    m_strSql = "Select TOP 1 1 from cust_Ord_hdr where unit_code='" & gstrUNITID & "' and Account_code ='" & Trim(txtCustomerCode.Text) & "'"
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                    m_strSql = "Select TOP 1 1 from Customer_Mst where unit_code='" & gstrUNITID & "' and Customer_code ='" & Trim(txtCustomerCode.Text) & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
            End Select
            rsdb.GetResult(m_strSql)
            If rsdb.GetNoRows > 0 Then
                txtReferenceNo.Enabled = True
                txtReferenceNo.BackColor = System.Drawing.Color.White
                cmdHelp(1).Enabled = True
                Call FillLabel("CUSTOMER")
                ''ADDED BY SUMIT KUMAR ON 17 JULY 2019 FOR CHECK EXTERNAL SO FLAG
                Call HideExternalSo()
                blnOneCustDrgNo_MutipleItem = SqlConnectionclass.ExecuteScalar("SELECT OneCustDrgNo_MutipleItem FROM customer_mst WHERE CUSTOMER_CODE='" & txtCustomerCode.Text & "' and UNIT_CODE='" & gstrUNITID & "'")

            Else
                cmdButtons.Focus()
                MsgBox("Customer Code does not exist", MsgBoxStyle.Information, ResolveResString(100))
                txtCustomerCode.Text = ""
                txtCustomerCode.Focus()
                mvalid = True
            End If
        End If
        m_blnCloseFlag = False
        m_blnHelpFlag = False
        If StrComp(Trim(txtCustomerCode.Text), mstrPrevAccountCode, CompareMethod.Text) <> 0 Then
            Call CLEARVAR()
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub TxtCustomerCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtCustomerCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        Dim rsCD As New ClsResultSetDB
        With ssPOEntry
            .Col = 19
            .Col2 = 19
            .ColHidden = True
            .Col = 20
            .Col2 = 20
            .ColHidden = True
        End With
        If m_blnCloseFlag = True Then
            m_blnCloseFlag = False
            GoTo EventExitSub
        End If
        If Len(Trim(txtCustomerCode.Text)) = 0 Then
            GoTo EventExitSub
        Else
            m_strSql = "Select TOP 1 1 from Customer_mst where unit_code='" & gstrUNITID & "' and customer_Code='" & Trim(txtCustomerCode.Text) & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
            rsCD.GetResult(m_strSql)
            If rsCD.GetNoRows = 0 Then
                Call ConfirmWindow(10145, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                txtCustomerCode.Text = ""
                txtReferenceNo.Text = ""
                Cancel = True
                txtCustomerCode.Focus()
                GoTo EventExitSub
            Else
                '10736222-changes done for CT2 ARE 3 functionality - to enable CT2 Reqd flag
                If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                    strSql = "Select dbo.UDF_IsCT2Customer('" & gstrUNITID & "','" & txtCustomerCode.Text.Trim & "')"
                    ChkCT2Reqd.Enabled = Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSql))
                    FillDataTables()
                End If
                'ISSUE ID : 10763705 
                strSql = "Select appendsoitem from customer_mst where unit_code='" & gstrUNITID & "' and customer_code='" & txtCustomerCode.Text.Trim & "'"
                mblnappendsoitem_customer = Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSql))
                'ISSUE ID : 10763705 
                'GST CHANGES
                strSql = "Select ALLOW_MULTIPLE_HSN_ITEMS from customer_mst where unit_code='" & gstrUNITID & "' and customer_code='" & txtCustomerCode.Text.Trim & "'"
                MBLNALLOW_MULTIPLEHSNITEMS_CUSTOMER = Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSql))
                'GST CHANGES


            End If
        End If
        GoTo EventExitSub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        GoTo EventExitSub
EventExitSub:
        eventArgs.Cancel = Cancel
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
    End Sub
    Private Sub TxtReferenceNo_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtReferenceNo.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = 112 Then Call cmdHelp_Click(cmdHelp.Item(1), New System.EventArgs())
        If KeyCode = System.Windows.Forms.Keys.Return Then
            If Me.txtAmendmentNo.Enabled = True Then
                Me.txtAmendmentNo.Focus()
            Else
            End If
        ElseIf (KeyCode = 39) Or (KeyCode = 34) Or (KeyCode = 96) Then
            KeyCode = 0
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub DTValidDate_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles DTValidDate.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
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
            MsgBox("Valid Date Cannot be Less then Current Date.", MsgBoxStyle.OkOnly, ResolveResString(100))
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
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub frmMKTTRN0001_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo ErrHandler
        'Checking the form name in the Windows list
        mdifrmMain.CheckFormName = mintFormIndex
        frmModules.NodeFontBold(Tag) = True
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0001_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Or cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
            If KeyCode = System.Windows.Forms.Keys.Escape Then
                Call cmdButtons_ButtonClick(cmdButtons, New UCActXCtl.UCbtndgrp.ButtonClickEventArgs(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL))
            End If
        End If
        If KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0 Then
            Call ctlHeader_ClickEvent(ctlHeader, New System.EventArgs())
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0001_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo ErrHandler
        Dim rsGetDate As New ClsResultSetDB
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
        m_strSql = "Select Financial_EndDate from company_mst WHERE unit_code='" & gstrUNITID & "' "
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
            .Col = 12
            .Col2 = 12
            .ColHidden = True
            .Col = 13
            .Col2 = 13
            .ColHidden = True
            .Col = 14
            .Col2 = 14
            .ColHidden = True
            .Col = 15
            .Col2 = 15
            .ColHidden = True
            .Col = 16
            .Col2 = 16
            .ColHidden = True
            .Col = 17
            .Col2 = 17
            .ColHidden = True
            .Col = 18
            .Col2 = 18
            .ColHidden = False
            .Col = 19
            .Col2 = 19
            .ColHidden = True
            .Col = 20
            .Col2 = 20
            .ColHidden = True
            .Col = 21
            .Col2 = 21
            .ColHidden = True
            .Col = 33
            .Col2 = 33
            .ColHidden = True
        End With
        Call InitializeSpreed()
        DTDate.Value = GetServerDate()
        DTEffectiveDate.Value = GetServerDate()
        DTAmendmentDate.Value = GetServerDate()
        DTValidDate.Value = rsGetDate.GetValue("Financial_EndDate")
        rsGetDate.ResultSetClose()
        rsGetDate = Nothing
        Call SSMaxLength()
        CopySSPoEntry.Visible = False
        m_blnHelpFlag = False
        m_blnCloseFlag = False
        Call AddPOType()
        Call addExportSotype()
        Call AdddiscountSelection()
        Call AddmarkUpSelection()
        CmbDiscounttype.Text = ObsoleteManagement.GetItemString(CmbDiscounttype, 0)
        cmbPOType.Text = ObsoleteManagement.GetItemString(cmbPOType, 0)
        CmbMarkuptype.Text = ObsoleteManagement.GetItemString(CmbMarkuptype, 0)
        m_strSalesTaxType = ""
        frmSearch.Enabled = True
        optPartNo.Enabled = True
        optPartNo.Checked = True
        optPartNo.Checked = True
        optItem.Enabled = True
        txtsearch.Enabled = True
        txtsearch.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        ' Change By Deepak on 11-Oct-2011 for Support Change Management-----------------------
        'issue id :10114335
        If UCase(Trim(GetPlantName)) = "HILEX" Then
            txtReferenceNo.MaxLength = 34
        Else
            txtReferenceNo.MaxLength = 25
        End If
        mblnDiscountFunctionality = CBool(Find_Value("SELECT DISCOUNT_ON_INVOICE  FROM SALES_PARAMETER WHERE UNIT_CODE='" & gstrUNITID & "'"))

        If mblnDiscountFunctionality = False Then
            With ssPOEntry
                .Col = 22
                .Col2 = 22
                .ColHidden = True
                .Col = 23
                .Col2 = 23
                .ColHidden = True
            End With
            Me.FrmDiscountValue.Visible = False
        Else
            With ssPOEntry
                .Col = 22
                .Col2 = 22
                .ColHidden = False
                .Col = 23
                .Col2 = 23
                .ColHidden = False
            End With
            Me.FrmDiscountValue.Visible = True


        End If
        'Issue id :10114335  end 
        '10561117  
        mblnMarkupFunctionality = CBool(Find_Value("SELECT MARKUP_FUNCTIONALITY  FROM SALES_PARAMETER WHERE UNIT_CODE='" & gstrUNITID & "'"))
        If mblnMarkupFunctionality = False Then
            With ssPOEntry
                .Col = 24
                .Col2 = 24
                .ColHidden = True
                .Col = 25
                .Col2 = 25
                .ColHidden = True
            End With
            Me.FrmMarkup.Visible = False
        Else
            With ssPOEntry
                .Col = 24
                .Col2 = 24
                .ColHidden = False
                .Col = 25
                .Col2 = 25
                .ColHidden = False
            End With
            Me.FrmMarkup.Visible = True
        End If
        '10561117  end 
        '10940008
        mblnServiceInvoicemate = CBool(Find_Value("SELECT ServiceInvoice_MATE FROM SALES_PARAMETER WHERE UNIT_CODE='" + gstrUNITID + "'"))
        '10940008
        If gblnGSTUnit = True Then
            cmdAddVAT.Enabled = False
            txtAddVAT.Enabled = False
            txtAddVAT.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            _blnEOUFlag = CBool(Find_Value("SELECT ISNULL(EOU_FLAG,0) EOU_FLAG FROM SALES_PARAMETER WHERE UNIT_CODE='" & gstrUnitId & "'"))
        End If
        '' Added by priti for CAD functionality on 11 Nov 2024
        mblnCADM_SALES_ORDER_CREATION = CBool(Find_Value("SELECT ISNULL(CADM_SALES_ORDER_CREATION,0)  FROM SALES_PARAMETER WHERE UNIT_CODE='" & gstrUNITID & "'"))
        blnSO_EDITABLE = Convert.ToBoolean(SqlConnectionclass.ExecuteScalar("Select isnull(SO_EDITABLE,0) from Sales_parameter where  unit_code='" & gstrUNITID & "'"))
        If blnSO_EDITABLE = True Then  '' Added by priti on 25 Jul 2025 
            DTDate.Enabled = True
        Else
            DTDate.Enabled = False
        End If
        '-----------------------------------------
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0001_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
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
        If gblnCancelUnload = True Then
            eventArgs.Cancel = True
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub frmMKTTRN0001_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
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
        Dim TMPDr As DataRow() 'AMIT02APR2022
        Dim rsParameter As ClsResultSetDB
        rsParameter = New ClsResultSetDB
        Dim rsitem As New ClsResultSetDB
        Dim varQty, varDrawPartNo, varItemCode, varOpenSO As Object
        Dim varCustSuppMat, varRate, varToolCost As Object
        Dim varflg, varPF, varStax, varExd, varSSt, varOthers, varDespatch As Object
        Dim varAbatment, varMRP, varAccessibleRateforMRP As Object
        Dim dummyVarItem As Object
        Dim varDelFlag As Object
        Dim intRow As Short ' to get the values in the grid
        Dim strMeasure As String
        Dim rsMeasure As ClsResultSetDB
        Dim emptyvar As Object
        Dim strtest As String
        Dim vardiscounttype As Object
        Dim vardiscountvalue As Object
        Dim varmarkuptype As Object
        Dim varmarkupvalue As Object
        'GST CHANGE
        Dim varHSNCode As Object
        Dim varfirstItemCode As Object
        Dim varHSNSACCode As Object
        Dim VARCGSTTAX As Object
        Dim VARSGSTTAX As Object
        Dim VARUTGSTTAX As Object
        Dim VARIGSTTAX As Object

        'mblnDiscountMandatory = CBool(Find_Value("SELECT isnull(SOUPLD_Discounteditable,0)   FROM customer_mst (Nolock) WHERE  customer_code='" & txtCustomerCode.Text.Trim() & "' and UNIT_CODE='" + gstrUNITID + "'"))

        TMPDr = datatable_MasterData.Select("ACCOUNT_CODE = '" & txtCustomerCode.Text.Trim() & "'")
        If TMPDr.Length > 0 Then
            mblnDiscountMandatory = CBool(TMPDr(0)("SOUPLD_Discounteditable").ToString)
        Else
            mblnDiscountMandatory = ""
        End If

        'GST CHANGE
        varDelFlag = Nothing
        Call ssPOEntry.GetText(0, Row, varDelFlag)
        ssPOEntry.Row = Row
        ssPOEntry.Col = 1
        varOpenSO = ssPOEntry.Value
        varOpenSO = Nothing
        Call ssPOEntry.GetText(1, Row, varOpenSO)
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
        varOthers = Nothing
        Call ssPOEntry.GetText(11, Row, varOthers)
        varMRP = Nothing
        Call ssPOEntry.GetText(19, Row, varMRP)
        varAbatment = Nothing
        Call ssPOEntry.GetText(20, Row, varAbatment)
        varAccessibleRateforMRP = Nothing
        Call ssPOEntry.GetText(21, Row, varAccessibleRateforMRP)
        vardiscounttype = Nothing
        Call ssPOEntry.GetText(22, Row, vardiscounttype)
        vardiscountvalue = Nothing
        Call ssPOEntry.GetText(23, Row, vardiscountvalue)
        varmarkuptype = Nothing
        Call ssPOEntry.GetText(24, Row, varmarkuptype)
        varmarkupvalue = Nothing
        Call ssPOEntry.GetText(25, Row, varmarkupvalue)
        varHSNSACCode = Nothing
        Call ssPOEntry.GetText(26, Row, varHSNSACCode)
        VARCGSTTAX = Nothing
        Call ssPOEntry.GetText(27, Row, VARCGSTTAX)
        VARSGSTTAX = Nothing
        Call ssPOEntry.GetText(28, Row, VARSGSTTAX)
        VARUTGSTTAX = Nothing
        Call ssPOEntry.GetText(29, Row, VARUTGSTTAX)
        VARIGSTTAX = Nothing
        Call ssPOEntry.GetText(30, Row, VARIGSTTAX)
        'GST CHANGES
        If gblnGSTUnit = True Then
            varfirstItemCode = Nothing
            Call ssPOEntry.GetText(4, 1, varfirstItemCode)
            '  varHSNCode = FindScalar_Value("SELECT HSN_SAC_CODE FROM ITEM_MST (Nolock) WHERE UNIT_CODE='" & gstrUNITID & "' AND ITEM_CODE='" & varfirstItemCode & "'")
            TMPDr = datatable_MasterData.Select("ITEM_CODE = '" & Trim(varfirstItemCode) & "'")
            If TMPDr.Length > 0 Then
                varHSNCode = TMPDr(0)("HSN_SAC_CODE").ToString
            Else
                varHSNCode = ""
            End If

        End If

        If gblnGSTUnit = True And ssPOEntry.MaxRows >= 1 And UCase(Trim(cmbPOType.Text)) <> "EXPORT" And varDrawPartNo <> "" Then
            If (varHSNSACCode = "") Then
                ValidRowData = False
                MsgBox("HSN/SAC CODE SHOULD NOT BE BLANK " + varItemCode, MsgBoxStyle.OkOnly, ResolveResString(100))
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

        If varDelFlag = "*" Then
            ValidRowData = True
            Exit Function
        End If
        'ISSUE ID 10239863
        If Val(varOpenSO) = 1 And UCase(Trim(cmbPOType.Text)) = "Q-TRADING" Then
            MsgBox("Open Sales order is not Allowed for Trading Sales order", MsgBoxStyle.OkOnly, ResolveResString(100))
            With ssPOEntry
                .Row = .ActiveRow : Col = 1 : .Value = 0
                .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                .Focus()
            End With
            ValidRowData = False
            Call ssSetFocus(Row)
            ssPOEntry.Focus()

            Exit Function

        End If
        'ISSUE ID 10239863
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
            '10940008
            If Mid(cmbPOType.Text, 1, 1) = "V" And mblnServiceInvoicemate = True Then
                'm_strSql = "Select TOP 1 1 from custitem_mst as A (Nolock), Item_MST as B (Nolock) where A.UNIT_CODE=B.UNIT_CODE AND ACTIVE = 1 and B.Status = 'A' and B.hold_flag = 0 and B.ITEM_MAIN_GRP = 'M' AND a.Item_code=b.item_code and a.ITem_code = '" & varItemCode & "' and Cust_Drgno = '" & Trim(varDrawPartNo) & "' and account_code='" & Trim(txtCustomerCode.Text) & "' AND A.unit_code='" & gstrUNITID & "'"

                TMPDr = datatable_MasterData.Select("ACTIVE = 1 and Status = 'A' and hold_flag = 0 and ITEM_MAIN_GRP = 'M' AND ITem_code = '" & varItemCode & "' and Cust_Drgno = '" & Trim(varDrawPartNo) & "'")


            Else
                'm_strSql = "Select TOP 1 1 from custitem_mst as A (Nolock), Item_MST as B (Nolock) where A.UNIT_CODE=B.UNIT_CODE AND ACTIVE = 1 AND a.Item_code=b.item_code and B.Status = 'A' and B.hold_flag = 0 and a.ITem_code = '" & varItemCode & "' and Cust_Drgno = '" & Trim(varDrawPartNo) & "' and account_code='" & Trim(txtCustomerCode.Text) & "' AND A.unit_code='" & gstrUNITID & "'"

                TMPDr = datatable_MasterData.Select("ACTIVE = 1 AND Status = 'A' and hold_flag = 0 and ITem_code = '" & varItemCode & "' and Cust_Drgno = '" & Trim(varDrawPartNo) & "'")

            End If

            'rsitem.GetResult(m_strSql)
            'If rsitem.GetNoRows <= 0 Then
            If TMPDr.Length <= 0 Then
                'MsgBox("Please check refrence of this  item in  " & vbCrLf & "Item Master or Customer Item Master or in Item Rate Master")
                MsgBox("Please check reference of this  Cust. Part: " + " '" & Trim(varDrawPartNo) & "' " + "  " & vbCrLf & "Item Code:" + " '" & Trim(varItemCode) & "' " + " " & vbCrLf & "in Item Master or Customer Item Master or in Item Rate Master")
                'Call ConfirmWindow(10154, BUTTON_OK, IMG_INFO)
                ValidRowData = False
                Call ssPOEntry.SetText(2, Row, "")
                Call ssPOEntry.SetText(4, Row, "")
                Call ssSetFocus(Row)
                ssPOEntry.Focus()
                Exit Function
            End If

            Dim dummyVarItemHSNCODE As Object
            dummyVarItemHSNCODE = Nothing

            'm_strSql = "Select TOP 1 1 from cust_ord_dtl where unit_code='" & gstrUNITID & "' and Cust_Drgno = '" & Trim(varDrawPartNo) & "'and account_code='" & txtCustomerCode.Text & "' and cust_ref='" & txtReferenceNo.Text & "'and active_Flag='A' and ITem_code = '" & varItemCode & "' and amendment_no = '" & txtAmendmentNo.Text & "'"
            'rsitem.GetResult(m_strSql)

            TMPDr = datatable_ExistingSoDetails.Select("unit_code='" & gstrUNITID & "' and Cust_Drgno = '" & Trim(varDrawPartNo) & "'and account_code='" & txtCustomerCode.Text & "' and cust_ref='" & txtReferenceNo.Text & "'and active_Flag='A' and ITem_code = '" & varItemCode & "' and amendment_no = '" & txtAmendmentNo.Text & "'")
            'If rsitem.GetNoRows > 1 Then
            If TMPDr.Length > 1 Then
                'Call ConfirmWindow(10069, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                MsgBox("Item Code:" + " '" & Trim(varItemCode) & "' " + " Already Exists please enter another Item Code")
                ValidRowData = False
                Call ssPOEntry.SetText(2, Row, "")
                Call ssPOEntry.SetText(4, Row, "")
                Call ssSetFocus(Row)
                ssPOEntry.Focus()
                Exit Function
            End If
            'GST CHANGES


            'strSql = "Select ALLOW_MULTIPLE_HSN_ITEMS from customer_mst (Nolock) where unit_code='" & gstrUNITID & "' and customer_code='" & txtCustomerCode.Text.Trim & "'"
            'MBLNALLOW_MULTIPLEHSNITEMS_CUSTOMER = Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSql))

            TMPDr = datatable_MasterData.Select("unit_code='" & gstrUNITID & "' and Account_code='" & txtCustomerCode.Text.Trim & "'")
            If TMPDr.Length > 0 Then
                MBLNALLOW_MULTIPLEHSNITEMS_CUSTOMER = TMPDr(0)("ALLOW_MULTIPLE_HSN_ITEMS").ToString
            Else
                MBLNALLOW_MULTIPLEHSNITEMS_CUSTOMER = "0"
            End If

            If gblnGSTUnit = True And MBLNALLOW_MULTIPLEHSNITEMS_CUSTOMER = False And ssPOEntry.MaxRows > 1 Then
                For intRow = 1 To ssPOEntry.MaxRows
                    dummyVarItemHSNCODE = Nothing
                    Call ssPOEntry.GetText(4, intRow, dummyVarItemHSNCODE)

                    ' dummyVarItemHSNCODE = FindScalar_Value("SELECT HSN_SAC_CODE FROM ITEM_MST (Nolock) WHERE UNIT_CODE='" & gstrUNITID & "' AND ITEM_CODE='" & dummyVarItemHSNCODE & "'")
                    TMPDr = datatable_MasterData.Select("UNIT_CODE='" & gstrUNITID & "' AND ITEM_CODE='" & dummyVarItemHSNCODE & "'")
                    If TMPDr.Length > 0 Then
                        dummyVarItemHSNCODE = TMPDr(0)("HSN_SAC_CODE").ToString
                    Else
                        dummyVarItemHSNCODE = ""
                    End If


                    If dummyVarItemHSNCODE <> "" And dummyVarItemHSNCODE <> varHSNCode Then
                        MsgBox("MULTPLE HSN NOT ALLOWED , PREVIOUS ITEM LINKED WITH HSN CODE : " + varHSNCode)
                        ValidRowData = False
                        Call ssSetFocus(Row)
                        ssPOEntry.Focus()
                        emptyvar = ""
                        Call ssPOEntry.SetText(2, ssPOEntry.ActiveRow, emptyvar)
                        Exit Function
                    End If

                    dummyVarItem = Nothing
                    Call ssPOEntry.GetText(2, intRow, dummyVarItem)
                    If blnOneCustDrgNo_MutipleItem = False Then '' This condition added by priti Hilex change Multiple item mapped with one drawing no on 31 Jan 2025
                        If dummyVarItem = varDrawPartNo And intRow <> Row Then
                            'Call ConfirmWindow(10156, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            MsgBox("Customer Item Code:" + " '" & Trim(dummyVarItem) & "' " + "This Item Already Exists in the List ")
                            ValidRowData = False
                            Call ssSetFocus(Row)
                            ssPOEntry.Focus()
                            emptyvar = ""
                            If intRow > Row Then
                                Call ssPOEntry.SetText(2, intRow, emptyvar)
                                Call ssPOEntry.SetText(4, intRow, emptyvar)
                            Else
                                Call ssPOEntry.SetText(2, Row, emptyvar)
                                Call ssPOEntry.SetText(4, Row, emptyvar)
                            End If
                            Call ssSetFocus(Row, 2)
                            Exit Function
                        End If
                    End If
                Next
            Else
                'GST CHANGES
                If blnOneCustDrgNo_MutipleItem = False Then '' This condition added by priti Hilex change Multiple item mapped with one drawing no on 31 Jan 2025

                    For intRow = 1 To ssPOEntry.MaxRows
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
                End If
            End If
        End If



        If Col = 0 Or Col = 4 Then ' if col is 3rd
            If Len(Trim(varItemCode)) = 0 Then
                ValidRowData = False
                Call ssPOEntry.SetText(2, Row, "")
                Call ssSetFocus(Row)
                ssPOEntry.Focus()
                Exit Function
            End If


            'm_strSql = "Select TOP 1 1 from Custitem_mst as A (Nolock), Item_MST As B (Nolock) where A.UNIT_CODE=B.UNIT_CODE AND A.Item_code=b.Item_code and Status='A' and A.ACTIVE = 1 and Hold_Flag=0 and Cust_Drgno = '" & Trim(varDrawPartNo) & "' and Account_Code='" & Trim(txtCustomerCode.Text) & "' and A.Item_code ='" & Trim(varItemCode) & "' AND A.unit_code='" & gstrUNITID & "' "
            'rsitem.GetResult(m_strSql)

            TMPDr = datatable_MasterData.Select("Status='A' and ACTIVE = 1 and Hold_Flag=0 and Cust_Drgno='" & Trim(varDrawPartNo) & "' and Account_Code='" & Trim(txtCustomerCode.Text) & "' and Item_code ='" & Trim(varItemCode) & "' AND unit_code='" & gstrUNITID & "'")


            ' If rsitem.GetNoRows <= 0 Then
            If TMPDr.Length <= 0 Then
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
        If (Col = 0 Or Col = 5) Then ' if col is 4
            If chkOpenSo.CheckState = 0 Then

                If Val(varOpenSO) = 0 Then
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
                            If MsgBox("Despatch Qty for item: " + " '" & Trim(dummyVarItem) & "' " + " is [ " & varDespatch & " ] would you like to add this Quantity You have entered.", MsgBoxStyle.YesNo, ResolveResString(100)) = MsgBoxResult.Yes Then
                                Call ssPOEntry.SetText(5, Row, varQty + varDespatch)
                                varQty = Nothing
                                Call ssPOEntry.GetText(5, Row, varQty)
                            End If
                        End If
                    End If
                    'CHECK FOR MEASURMENT UNIT

                    'strMeasure = "select a.Decimal_allowed_flag from Measure_Mst a (Nolock),Item_Mst b (Nolock)"
                    'strMeasure = strMeasure & " where A.UNIT_CODE=B.UNIT_CODE AND b.cons_Measure_Code=a.Measure_Code and b.Item_Code = '" & varItemCode & "' AND A.unit_code='" & gstrUNITID & "'"
                    'rsMeasure = New ClsResultSetDB
                    'rsMeasure.GetResult(strMeasure, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)

                    TMPDr = datatable_MasterData.Select("Item_Code = '" & varItemCode & "' AND unit_code='" & gstrUNITID & "'")

                    'If rsMeasure.GetValue("Decimal_allowed_flag") = False Then
                    If Convert.ToBoolean(TMPDr(0)("Decimal_allowed_flag").ToString()) = False Then
                        If System.Math.Round(varQty, 3) - Val(varQty) <> 0 Then
                            'Call ConfirmWindow(10455, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            MsgBox("The Quantity of Item:" + " '" & Trim(varItemCode) & "' " + "is not defined in Decimals/Fractions.")
                            ValidRowData = False
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
                ElseIf Val(varOpenSO) = 1 Then
                    If varQty > 0 Then
                        MsgBox("This Item (" & varDrawPartNo & ") is Open , Quantity should not be greater than 0.", MsgBoxStyle.OkOnly, "eMPro")
                        ValidRowData = False
                        Call ssSetFocus(Row, 5)
                        ssPOEntry.Focus()
                        Exit Function
                    End If
                ElseIf varOpenSO = "" Then
                    If varQty <= 0 Or Val(Trim(varQty)) <= 0 Then
                        If varDrawPartNo <> "" Then
                            'Call ConfirmWindow(10224, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            MsgBox("For Customer Item Code: " + " '" & Trim(varDrawPartNo) & "' " + "Quantity is Zero which can't be Zero")
                            ValidRowData = False
                            ssPOEntry.Col = 5
                            Call ssSetFocus(Row, 5)
                            ssPOEntry.Focus()
                            Exit Function
                        End If
                    End If
                End If
            End If 'Flag Check
        End If
        If Col = 0 Or Col = 6 Then ' if col is 5
            If varRate = 0 Or Len(Trim(varRate)) = 0 Then
                MsgBox("Enter Rate Greater than 0", MsgBoxStyle.OkOnly, "empower")
                ValidRowData = False
                Call ssSetFocus(Row, 6)
                ssPOEntry.Focus()
                Exit Function
            End If
            If varQty < 0 Then
                MsgBox("Enter Rate Greater than 0", MsgBoxStyle.OkOnly, "empower")
                ValidRowData = False
                ssPOEntry.Col = 6
                Call ssSetFocus(Row, 6)
                ssPOEntry.Focus()
                Exit Function
            End If
        End If



        If mblnDiscountFunctionality = True Then
            If Col = 23 Then ' if col is 5

                If vardiscounttype = "NONE" And vardiscountvalue > 0 Then
                    MsgBox("Discount Rate can't be Defined for NONE case ", MsgBoxStyle.OkOnly, ResolveResString(100))
                    ValidRowData = False
                    Call ssSetFocus(Row, 23)
                    Call ssPOEntry.SetText(23, Row, "0.00")
                    ssPOEntry.Focus()
                    Exit Function
                End If
                If varRate > 0 And vardiscounttype = "[V]alue" And vardiscountvalue >= varRate And vardiscountvalue > 0 Then
                    MsgBox("Discount Rate can't be Equal to or greater than item Rate", MsgBoxStyle.OkOnly, ResolveResString(100))
                    ValidRowData = False
                    Call ssSetFocus(Row, 23)
                    Call ssPOEntry.SetText(23, Row, "0.00")
                    ssPOEntry.Focus()
                    Exit Function
                End If
                If vardiscounttype = "[P]ercentage" And vardiscountvalue > 100 Then
                    MsgBox(" % cannot be greater than 100", MsgBoxStyle.OkOnly, ResolveResString(100))
                    ValidRowData = False
                    Call ssSetFocus(Row, 23)
                    Call ssPOEntry.SetText(23, Row, "0.00")
                    ssPOEntry.Focus()
                    Exit Function
                End If
            End If
        End If
        '10561117  
        If mblnMarkupFunctionality = True Then
            If Col = 25 Then
                If varmarkuptype = "[P]ercentage" And varmarkupvalue > 100 Then
                    MsgBox(" % cannot be greater than 100", MsgBoxStyle.OkOnly, ResolveResString(100))
                    ValidRowData = False
                    Call ssSetFocus(Row, 25)
                    Call ssPOEntry.SetText(25, Row, "")
                    ssPOEntry.Focus()
                    Exit Function
                End If
            End If
        End If
        '10561117  end 

        If Col = 0 Or Col = 10 Then ' if col is 5
            If UCase(Trim(cmbPOType.Text)) = "JOB WORK" Then
                If Len(Trim(varExd)) >= 1 Then
                    With ssPOEntry


                        TMPDr = datatable_MasterData_GEN_TAXRATE.Select("Tx_TaxeID = 'EXC' and Txrt_Rate_no = '" & varExd & "'")

                        'rsParameter.GetResult("Select Txrt_Rate_no,TxRt_Percentage from Gen_TaxRate (Nolock) where unit_code='" & gstrUNITID & "' and Tx_TaxeID = 'EXC' and Txrt_Rate_no = '" & varExd & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))")
                        'If rsParameter.GetNoRows = 0 Then
                        If TMPDr.Length <= 0 Then
                            MsgBox("Invalid Excise Code.", MsgBoxStyle.Information, ResolveResString(100))
                            .Row = .ActiveRow
                            .Col = 10 : .Text = ""
                            .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                            .Focus()
                            ValidRowData = False
                            Exit Function
                        End If
                    End With
                End If
            ElseIf UCase(Trim(cmbPOType.Text)) = "Q-TRADING" Then
            Else
                'gst changes
                If Len(Trim(varExd)) = 0 Then
                    If gblnGSTUnit = False Or (gblnGSTUnit And _blnEOUFlag) Then
                        'gst changes
                        MsgBox("Excise Duty Cannot be blank", MsgBoxStyle.Information, ResolveResString(100))
                        ValidRowData = False
                        ssPOEntry.Col = 10
                        Call ssSetFocus(Row, 10)
                        ssPOEntry.Focus()
                        Exit Function
                    End If

                ElseIf Len(Trim(varExd)) >= 1 Then
                    With ssPOEntry

                        TMPDr = datatable_MasterData_GEN_TAXRATE.Select("Tx_TaxeID = 'EXC' and Txrt_Rate_no = '" & varExd & "'")
                        ' rsParameter.GetResult("Select Txrt_Rate_no,TxRt_Percentage from Gen_TaxRate (Nolock) where unit_code='" & gstrUNITID & "' and Tx_TaxeID = 'EXC' and Txrt_Rate_no = '" & varExd & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))")

                        'If rsParameter.GetNoRows = 0 Then
                        If TMPDr.Length = 0 Then
                            MsgBox("Invalid Excise Code.", MsgBoxStyle.Information, ResolveResString(100))
                            .Row = .ActiveRow
                            .Col = 10 : .Text = ""
                            .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                            .Focus()
                            ValidRowData = False
                            Exit Function
                        End If
                    End With
                End If
            End If
        End If

        If Col = 0 Or Col = 9 Then
            If Len(Trim(varPF)) = 0 And Len(Trim(varItemCode)) > 0 Then
                MsgBox("Packing Tax Cannot be left blank.", MsgBoxStyle.Information, ResolveResString(100))
                mblnpackingdefined = True
                ValidRowData = False
                ssPOEntry.Col = 9
                Call ssSetFocus(Row, 9)
                ssPOEntry.Focus()
                Exit Function
            ElseIf Len(Trim(varPF)) >= 1 Then
                With ssPOEntry
                    TMPDr = datatable_MasterData_GEN_TAXRATE.Select("Tx_TaxeID = 'PKT' and Txrt_Rate_no = '" & Trim(varPF) & "'")
                    ' rsParameter.GetResult("Select Txrt_Rate_no,TxRt_Percentage from Gen_TaxRate (Nolock) where unit_code='" & gstrUNITID & "' and Tx_TaxeID = 'PKT' and Txrt_Rate_no = '" & Trim(varPF) & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))")
                    If TMPDr.Length = 0 Then
                        MsgBox("Invalid Packing Code.", MsgBoxStyle.Information, ResolveResString(100))
                        .Row = .ActiveRow
                        .Col = 9 : .Text = ""
                        .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                        .Focus()
                        ValidRowData = False
                        Exit Function
                    End If
                End With
            End If
        End If
        'GST CHANGES

        'GST CHANGES
        If UCase(Trim(cmbPOType.Text)) = "MRP-SPARES" Then
            With ssPOEntry
                .Col = 19
                .Col2 = 19
                .ColHidden = False
                .Col = 20
                .Col2 = 20
                .ColHidden = False
                .Col = 21
                .Col2 = 21
                .ColHidden = False
            End With
            If Col = 0 Or Col = 19 Then ' if col is 19
            End If
            If Col = 0 Or Col = 21 Then ' if col is 21
                If (varAccessibleRateforMRP = 0 Or Len(Trim(varAccessibleRateforMRP)) = 0) And varMRP > 0 Then
                    MsgBox("Enter Accessible Rate more than 0", MsgBoxStyle.OkOnly, ResolveResString(100))
                    ValidRowData = False
                    Call ssSetFocus(Row, 21)
                    ssPOEntry.Focus()
                    Exit Function
                End If
            End If
            If Col = 0 Or Col = 20 Then ' if col is 19
                If Len(Trim(varAbatment)) >= 1 Then
                    With ssPOEntry

                        TMPDr = datatable_MasterData_GEN_TAXRATE.Select("Tx_TaxeID = 'ABNT' and Txrt_Rate_no = '" & varAbatment & "'")
                        'rsParameter.GetResult("Select Txrt_Rate_no,TxRt_Percentage from Gen_TaxRate (Nolock) where unit_code='" & gstrUNITID & "' and Tx_TaxeID = 'ABNT' and Txrt_Rate_no = '" & varAbatment & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))")
                        ' If rsParameter.GetNoRows = 0 Then
                        If TMPDr.Length = 0 Then
                            MsgBox("Invalid Abatment Code.", MsgBoxStyle.Information, ResolveResString(100))
                            .Row = .ActiveRow
                            .Col = 20 : .Text = ""
                            .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                            .Focus()
                            ValidRowData = False
                            Exit Function
                        End If
                    End With
                ElseIf varMRP > 0 And Len(Trim(varAbatment)) = 0 And gblnGSTUnit = True Then
                    MsgBox("Amendment Code Cannot be blank", MsgBoxStyle.Information, ResolveResString(100))
                    ValidRowData = False
                    ssPOEntry.Col = 20
                    Call ssSetFocus(Row, 20)
                    ssPOEntry.Focus()
                    Exit Function
                End If
            End If
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
        Dim rsCust As New ClsResultSetDB
        Select Case UCase(pstrCode)
            Case "CUSTOMER"
                m_strSql = "select Cust_Name,Credit_Days from Customer_mst where unit_code='" & gstrUNITID & "' and Customer_code='" & txtCustomerCode.Text & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
                rsCust.GetResult(m_strSql)
                lblCustDesc.ForeColor = System.Drawing.Color.White
                lblCustDesc.Text = IIf(UCase(rsCust.GetValue("Cust_Name")) = "UNKNOWN", "", rsCust.GetValue("Cust_Name"))
                txtCreditTerms.Text = IIf(UCase(rsCust.GetValue("Credit_Days")) = "UNKNOWN", "", rsCust.GetValue("Credit_Days"))
                blnNoneditableCreditTerms_onSO = Convert.ToBoolean(SqlConnectionclass.ExecuteScalar("Select isnull(NoneditableCreditTerms_onSO,0) from customer_mst where  unit_code='" & gstrUNITID & "' and Customer_code='" & txtCustomerCode.Text & "'"))

            Case "STAX"
                m_strSql = "select TxRt_Rate_No,TxRt_RateDesc from Gen_TaxRate where unit_code='" & gstrUNITID & "' and Txrt_Rate_No = '" & txtSTax.Text & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
                rsCust.GetResult(m_strSql)
                lblSTaxDesc.ForeColor = System.Drawing.Color.White
                lblSTaxDesc.Text = IIf(UCase(rsCust.GetValue("TxRt_RateDesc")) = "UNKNOWN", "", rsCust.GetValue("TxRt_RateDesc"))
                '1.Surcharge on S.Tax
            Case "SSCHTAX"
                m_strSql = "SELECT TxRt_Rate_no, TxRt_RateDesc FROM gen_taxrate where unit_code='" & gstrUNITID & "' and Txrt_Rate_No= '" & txtSChSTax.Text & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
                rsCust.GetResult(m_strSql)
                lblSChSTaxDesc.ForeColor = System.Drawing.Color.White
                lblSChSTaxDesc.Text = IIf(UCase(rsCust.GetValue("TxRt_RateDesc")) = "UNKNOWN", "", rsCust.GetValue("TxRt_RateDesc"))
                '------------------->>
            Case "CREDIT"
                m_strSql = "select crtrm_TermID,crtrm_Desc from Gen_CreditTrmMaster where unit_code='" & gstrUNITID & "' and crtrm_TermID = '" & txtCreditTerms.Text & "'"
                rsCust.GetResult(m_strSql)
                lblCreditTermDesc.ForeColor = System.Drawing.Color.White
                lblCreditTermDesc.Text = IIf(UCase(rsCust.GetValue("crtrm_Desc")) = "UNKNOWN", "", rsCust.GetValue("crtrm_Desc"))
            Case "CURRENCY"
                m_strSql = "select Cust_Name,currency_code from Customer_mst where unit_code='" & gstrUNITID & "' and Customer_code='" & txtCustomerCode.Text & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
                rsCust.GetResult(m_strSql)
                lblCustDesc.ForeColor = System.Drawing.Color.White
                txtCurrencyType.Text = IIf(UCase(rsCust.GetValue("currency_code")) = "UNKNOWN", "", rsCust.GetValue("currency_code"))
            Case "ADDVAT"
                m_strSql = "select TxRt_Rate_No,TxRt_RateDesc from Gen_TaxRate where unit_code='" & gstrUNITID & "' and Txrt_Rate_No = '" & txtAddVAT.Text & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
                rsCust.GetResult(m_strSql)
                lbladdvatdesc.ForeColor = System.Drawing.Color.White
                lbladdvatdesc.Text = IIf(UCase(rsCust.GetValue("TxRt_RateDesc")) = "UNKNOWN", "", rsCust.GetValue("TxRt_RateDesc"))
                '10869290
            Case "SRT"
                m_strSql = "select TxRt_Rate_No,TxRt_RateDesc from Gen_TaxRate where unit_code='" & gstrUNITID & "' and Txrt_Rate_No = '" & txtService.Text & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
                rsCust.GetResult(m_strSql)
                lblservicedesc.ForeColor = System.Drawing.Color.White
                lblservicedesc.Text = IIf(UCase(rsCust.GetValue("TxRt_RateDesc")) = "UNKNOWN", "", rsCust.GetValue("TxRt_RateDesc"))
        End Select
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    ''ADDED BY SUMIT KUMAR ON 17 JULY 2019 FOR CHECK EXTERNAL SO FLAG
    Public Sub HideExternalSo()
        On Error GoTo ErrHandler
        If DataExist("SELECT TOP 1 ENABLE_EXTERNAL_SALESNO FROM CUSTOMER_MST(NOLOCK) WHERE  UNIT_CODE='" & gstrUNITID & "' AND  CUSTOMER_CODE='" & Trim(txtCustomerCode.Text) & "' AND ENABLE_EXTERNAL_SALESNO =1 AND ((ISNULL(DEACTIVE_FLAG,0) <> 1) OR (CONVERT(VARCHAR(12),GETDATE(),106)<= CONVERT(VARCHAR(12),DEACTIVE_DATE,106)))") = True Then
            With ssPOEntry
                .Col = 33
                .Col2 = 33
                .ColHidden = False

            End With
        Else
            With ssPOEntry
                .Col = 33
                .Col2 = 33
                .ColHidden = True

            End With
        End If
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
            'Cancel = True
            cmbPOType.Enabled = True
            cmbPOType.BackColor = System.Drawing.Color.White
            cmbPOType.Focus()
            Exit Sub
        End If
        '10869290
        'If Not (UCase(Trim(cmbPOType.Text)) = "OEM" Or UCase(Trim(cmbPOType.Text)) = "JOB WORK" Or UCase(Trim(cmbPOType.Text)) = "SPARES" Or UCase(Trim(cmbPOType.Text)) = "MRP-SPARES" Or UCase(Trim(cmbPOType.Text)) = "EXPORT") Then
        If Not (UCase(Trim(cmbPOType.Text)) = "OEM" Or UCase(Trim(cmbPOType.Text)) = "JOB WORK" Or UCase(Trim(cmbPOType.Text)) = "SPARES" Or UCase(Trim(cmbPOType.Text)) = "MRP-SPARES" Or UCase(Trim(cmbPOType.Text)) = "EXPORT" Or UCase(Trim(cmbPOType.Text)) = "Q-TRADING" Or UCase(Trim(cmbPOType.Text)) = "A-SUB ASSEMBLY" Or UCase(Trim(cmbPOType.Text)) = "V-SERVICE") Then
            'MsgBox("Please Enter valid P.O. Type (OEM,J,S,E,M)", MsgBoxStyle.Information, ResolveResString(100))
            MsgBox("Please Enter valid P.O. Type (OEM,J,S,E,M,T,A,V)", MsgBoxStyle.Information, ResolveResString(100))
            cmbPOType.Enabled = True
            cmbPOType.BackColor = System.Drawing.Color.White
            cmbPOType.Focus()
            Exit Sub
        End If
        If cmbPOType.SelectedIndex = 4 And GetPlantName() <> "HILEX" Then
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
    Public Sub GetAmendmentDetailsNewUnderDevelopment()
        '----------------------------------------------------------------------------
        'Argument       :   Nil
        'Return Value   :   Nil
        'Function       :   Get the details if there is an amendment
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        mblnIsExecuting = True
        m_blnGetAmendmentDetails = True
        'Dim rsAD As New ClsResultSetDB
        Dim strAuthFlg As String
        Dim dtHeader As DataTable
        Dim dtDetail As DataTable
        m_strSql = "select * from cust_ord_dtl where unit_code='" & gstrUNITID & "' and Cust_ref='" & txtReferenceNo.Text & " 'and account_Code='" & txtCustomerCode.Text & "'and amendment_No='" & txtAmendmentNo.Text & "' order by Cust_Drgno"
        'rsdb.GetResult(m_strSql)
        dtDetail = SqlConnectionclass.GetDataTable(m_strSql)
        m_strSql = "select * from cust_ord_hdr where unit_code='" & gstrUNITID & "' and Cust_ref='" & txtReferenceNo.Text & " 'and account_Code='" & txtCustomerCode.Text & "' and amendment_No='" & txtAmendmentNo.Text & "'"
        dtHeader = SqlConnectionclass.GetDataTable(m_strSql)
        'rsAD.GetResult(m_strSql)
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        If dtHeader.Rows.Count > 0 Then
            'rsAD.MoveFirst()

            If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                DTDate.Value = CDate(dtHeader.Rows(0)("Order_Date"))
            End If

            '10736222-changes done for CT2 ARE 3 functionality
            ChkCT2Reqd.Checked = dtHeader.Rows(0)("CT2_Reqd_In_SO")

            lblIntSONoDes.Text = dtHeader.Rows(0)("InternalSONo")
            lblRevisionNo.Text = dtHeader.Rows(0)("RevisionNo")
            DTAmendmentDate.Value = dtHeader.Rows(0)("Amendment_Date")
            DTEffectiveDate.Value = dtHeader.Rows(0)("Effect_Date")
            DTValidDate.Value = dtHeader.Rows(0)("Valid_Date")
            txtCurrencyType.Text = dtHeader.Rows(0)("Currency_Code")
            '10177787
            txtAddVAT.Text = dtHeader.Rows(0)("addvat_type")
            '10177787
            ctlPerValue.Text = dtHeader.Rows(0)("PerValue")
            txtAmendReason.Text = dtHeader.Rows(0)("Reason")
            strpotype = dtHeader.Rows(0)("PO_Type")
            strexportType = dtHeader.Rows(0)("ExportSoType")
            If strexportType = "WP" Then
                cmbExporttype.SelectedIndex = 1
            ElseIf strexportType = "WOP" Then
                cmbExporttype.SelectedIndex = 2
            Else
                cmbExporttype.SelectedIndex = 0
            End If
            With ssPOEntry
                .Col = 19
                .Col2 = 19
                .ColHidden = True
                .Col = 20
                .Col2 = 20
                .ColHidden = True
                .Col = 21
                .Col2 = 21
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
                Case "T"
                    Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 6)
                Case "A"
                    Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 7)
                Case "V"
                    Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 8)

                    With ssPOEntry
                        .Col = 19
                        .Col2 = 19
                        .ColHidden = False
                        .Col = 20
                        .Col2 = 20
                        .ColHidden = False
                        .Col = 21
                        .Col2 = 21
                        .ColHidden = False
                    End With
            End Select
            txtSTax.Text = dtHeader.Rows(0)("salestax_Type")
            txtSChSTax.Text = dtHeader.Rows(0)("Surcharge_code")
            If dtHeader.Rows(0)("OpenSO") = False Then
                chkOpenSo.CheckState = System.Windows.Forms.CheckState.Unchecked
            Else
                chkOpenSo.CheckState = System.Windows.Forms.CheckState.Checked
            End If
            txtCreditTerms.Text = dtHeader.Rows(0)("Term_Payment")
            'rsdb.MoveFirst()
            ssPOEntry.MaxRows = 0

            prgItemDetails.Minimum = 0 : prgItemDetails.Value = 0 : prgItemDetails.Maximum = dtDetail.Rows.Count
            prgItemDetails.Visible = True
            CopySSPoEntry.Visible = True : ssPOEntry.Visible = False
            ssPOEntry.MaxRows = dtDetail.Rows.Count
            Call SSMaxLength()
            For intLoopCounter = 0 To dtDetail.Rows.Count - 1
                If dtDetail.Rows(intLoopCounter)("OpenSO") = False Then
                    ssPOEntry.Row = intLoopCounter
                    ssPOEntry.Col = 1
                    ssPOEntry.Value = 0
                Else
                    ssPOEntry.Row = intLoopCounter
                    ssPOEntry.Col = 1
                    ssPOEntry.Value = 1
                End If
                Call ssPOEntry.SetText(2, intLoopCounter, dtDetail.Rows(intLoopCounter)("Cust_DrgNo"))
                Call ssPOEntry.SetText(4, intLoopCounter, dtDetail.Rows(intLoopCounter)("Item_Code"))
                Call ssPOEntry.SetText(5, intLoopCounter, dtDetail.Rows(intLoopCounter)("Order_Qty"))
                Call ssPOEntry.SetText(13, intLoopCounter, dtDetail.Rows(intLoopCounter)("Rate"))
                Call ssPOEntry.SetText(6, intLoopCounter, dtDetail.Rows(intLoopCounter)("Rate") * CDbl(ctlPerValue.Text))
                Call ssPOEntry.SetText(14, intLoopCounter, dtDetail.Rows(intLoopCounter)("Cust_Mtrl"))
                Call ssPOEntry.SetText(7, intLoopCounter, dtDetail.Rows(intLoopCounter)("Cust_Mtrl") * CDbl(ctlPerValue.Text))
                Call ssPOEntry.SetText(15, intLoopCounter, dtDetail.Rows(intLoopCounter)("Tool_Cost"))
                Call ssPOEntry.SetText(8, intLoopCounter, dtDetail.Rows(intLoopCounter)("Tool_Cost") * CDbl(ctlPerValue.Text))
                Call ssPOEntry.SetText(9, intLoopCounter, dtDetail.Rows(intLoopCounter)("Packing_Type"))
                Call ssPOEntry.SetText(10, intLoopCounter, dtDetail.Rows(intLoopCounter)("Excise_Duty"))
                Call ssPOEntry.SetText(16, intLoopCounter, dtDetail.Rows(intLoopCounter)("Others"))
                Call ssPOEntry.SetText(11, intLoopCounter, dtDetail.Rows(intLoopCounter)("Others") * CDbl(ctlPerValue.Text))
                Call ssPOEntry.SetText(19, intLoopCounter, dtDetail.Rows(intLoopCounter)("MRP"))
                Call ssPOEntry.SetText(20, intLoopCounter, dtDetail.Rows(intLoopCounter)("Abantment_code"))
                Call ssPOEntry.SetText(21, intLoopCounter, dtDetail.Rows(intLoopCounter)("AccessibleRateforMRP"))
                ssPOEntry.Col = 22
                ssPOEntry.Row = intLoopCounter
                ssPOEntry.CellType = FPSpreadADO.CellTypeConstants.CellTypeComboBox
                Call ssPOEntry.SetText(22, intLoopCounter, dtDetail.Rows(intLoopCounter)("DISCOUNT_TYPE"))
                ssPOEntry.TypeComboBoxList = "NONE" & Chr(9) & "[P]ercentage" & Chr(9) & "[V]alue"
                If dtDetail.Rows(intLoopCounter)("DISCOUNT_TYPE") = "[P]ercentage" Then
                    ssPOEntry.TypeComboBoxCurSel = 1
                ElseIf dtDetail.Rows(intLoopCounter)("DISCOUNT_TYPE") = "[V]alue" Then
                    ssPOEntry.TypeComboBoxCurSel = 2
                Else
                    ssPOEntry.TypeComboBoxCurSel = 0
                End If


                ssPOEntry.Col = 23
                ssPOEntry.Row = intLoopCounter
                ssPOEntry.CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                Call ssPOEntry.SetText(23, intLoopCounter, dtDetail.Rows(intLoopCounter)("DISCOUNT_VALUE"))
                Call ssPOEntry.SetText(26, intLoopCounter, dtDetail.Rows(intLoopCounter)("HSNSACCODE"))
                Call ssPOEntry.SetText(27, intLoopCounter, dtDetail.Rows(intLoopCounter)("CGSTTXRT_TYPE"))
                Call ssPOEntry.SetText(28, intLoopCounter, dtDetail.Rows(intLoopCounter)("SGSTTXRT_TYPE"))
                Call ssPOEntry.SetText(29, intLoopCounter, dtDetail.Rows(intLoopCounter)("UTGSTTXRT_TYPE"))
                Call ssPOEntry.SetText(30, intLoopCounter, dtDetail.Rows(intLoopCounter)("IGSTTXRT_TYPE"))
                Call ssPOEntry.SetText(32, intLoopCounter, dtDetail.Rows(intLoopCounter)("ISHSNORSAC"))
                Call ssPOEntry.SetText(33, intLoopCounter, dtDetail.Rows(intLoopCounter)("external_salesorder_no"))

                prgItemDetails.Value = prgItemDetails.Value + 1
            Next
            prgItemDetails.Visible = False
            CopySSPoEntry.Visible = False : ssPOEntry.Visible = True


        Else
            Call ConfirmWindow(10128, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
            ssPOEntry.MaxRows = 0
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

        cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = True
        mblnIsExecuting = False
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        mblnIsExecuting = False
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
        Try

            mblnIsExecuting = True
            m_blnGetAmendmentDetails = True
            Dim rsAD As New ClsResultSetDB
            Dim strAuthFlg As String
            FillDataTables()

            'm_strSql = "select * from cust_ord_dtl where unit_code='" & gstrUNITID & "' and Cust_ref='" & txtReferenceNo.Text & " 'and account_Code='" & txtCustomerCode.Text & "'and amendment_No='" & txtAmendmentNo.Text & "' order by Cust_Drgno"
            m_strSql = "select * from cust_ord_dtl where unit_code='" & gstrUNITID & "' and Cust_ref='" & txtReferenceNo.Text & " 'and account_Code='" & txtCustomerCode.Text & "'and amendment_No='" & txtAmendmentNo.Text & "'"
            m_strSql += " and exists (select top 1 1 from custitem_mst cm(nolock) inner join item_mst im (nolock) "
            m_strSql += " on cm.unit_code = cust_ord_dtl.unit_code and cm.account_code=cust_ord_dtl.account_code and cm.item_code=cust_ord_dtl.item_code "
            m_strSql += "and cm.cust_drgno =cust_ord_dtl.Cust_DrgNo and im.unit_code=cust_ord_dtl.unit_code and im.item_code=cust_ord_dtl.item_code "
            m_strSql += "and cm.active=1 and im.status='A' AND IM.HOLD_FLAG=0 ) order by Cust_Drgno"

            rsdb.GetResult(m_strSql)
            m_strSql = "select * from cust_ord_hdr where unit_code='" & gstrUNITID & "' and Cust_ref='" & txtReferenceNo.Text & " 'and account_Code='" & txtCustomerCode.Text & "' and amendment_No='" & txtAmendmentNo.Text & "'"
            rsAD.GetResult(m_strSql)
            Dim intLoopCounter As Short
            Dim intMaxLoop As Short
            If rsAD.GetNoRows > 0 Then
                rsAD.MoveFirst()
                If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                    DTDate.Value = CDate(rsAD.GetValue("Order_Date"))
                End If

                '10736222-changes done for CT2 ARE 3 functionality
                ChkCT2Reqd.Checked = rsAD.GetValue("CT2_Reqd_In_SO")

                lblIntSONoDes.Text = rsAD.GetValue("InternalSONo")
                lblRevisionNo.Text = rsAD.GetValue("RevisionNo")
                DTAmendmentDate.Value = rsAD.GetValue("Amendment_Date")
                DTEffectiveDate.Value = rsAD.GetValue("Effect_Date")
                DTValidDate.Value = rsAD.GetValue("Valid_Date")
                txtCurrencyType.Text = rsAD.GetValue("Currency_Code")
                '10177787
                txtAddVAT.Text = rsAD.GetValue("addvat_type")
                '10177787
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
                With ssPOEntry
                    .Col = 19
                    .Col2 = 19
                    .ColHidden = True
                    .Col = 20
                    .Col2 = 20
                    .ColHidden = True
                    .Col = 21
                    .Col2 = 21
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
                    Case "T"
                        Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 6)
                    Case "A"
                        Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 7)
                    Case "V"
                        Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 8)

                        With ssPOEntry
                            .Col = 19
                            .Col2 = 19
                            .ColHidden = False
                            .Col = 20
                            .Col2 = 20
                            .ColHidden = False
                            .Col = 21
                            .Col2 = 21
                            .ColHidden = False
                        End With
                End Select
                txtSTax.Text = rsAD.GetValue("salestax_Type")
                txtSChSTax.Text = rsAD.GetValue("Surcharge_code")
                If rsAD.GetValue("OpenSO") = False Then
                    chkOpenSo.CheckState = System.Windows.Forms.CheckState.Unchecked
                Else
                    chkOpenSo.CheckState = System.Windows.Forms.CheckState.Checked
                End If
                txtCreditTerms.Text = rsAD.GetValue("Term_Payment")
                rsdb.MoveFirst()
                ssPOEntry.MaxRows = 0
                intMaxLoop = rsdb.RowCount : rsdb.MoveFirst()
                prgItemDetails.Minimum = 0 : prgItemDetails.Value = 0 : prgItemDetails.Maximum = intMaxLoop
                prgItemDetails.Visible = True
                CopySSPoEntry.Visible = True : ssPOEntry.Visible = False
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
                    Call ssPOEntry.SetText(9, intLoopCounter, rsdb.GetValue("Packing_Type"))
                    Call ssPOEntry.SetText(10, intLoopCounter, rsdb.GetValue("Excise_Duty"))
                    Call ssPOEntry.SetText(16, intLoopCounter, rsdb.GetValue("Others"))
                    Call ssPOEntry.SetText(11, intLoopCounter, rsdb.GetValue("Others") * CDbl(ctlPerValue.Text))
                    Call ssPOEntry.SetText(19, intLoopCounter, rsdb.GetValue("MRP"))
                    Call ssPOEntry.SetText(20, intLoopCounter, rsdb.GetValue("Abantment_code"))
                    Call ssPOEntry.SetText(21, intLoopCounter, rsdb.GetValue("AccessibleRateforMRP"))
                    ssPOEntry.Col = 22
                    ssPOEntry.Row = intLoopCounter
                    ssPOEntry.CellType = FPSpreadADO.CellTypeConstants.CellTypeComboBox
                    Call ssPOEntry.SetText(22, intLoopCounter, rsdb.GetValue("DISCOUNT_TYPE"))
                    ssPOEntry.TypeComboBoxList = "NONE" & Chr(9) & "[P]ercentage" & Chr(9) & "[V]alue"
                    If rsdb.GetValue("DISCOUNT_TYPE") = "[P]ercentage" Then
                        ssPOEntry.TypeComboBoxCurSel = 1
                    ElseIf rsdb.GetValue("DISCOUNT_TYPE") = "[V]alue" Then
                        ssPOEntry.TypeComboBoxCurSel = 2
                    Else
                        ssPOEntry.TypeComboBoxCurSel = 0
                    End If


                    ssPOEntry.Col = 23
                    ssPOEntry.Row = intLoopCounter
                    ssPOEntry.CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                    Call ssPOEntry.SetText(23, intLoopCounter, rsdb.GetValue("DISCOUNT_VALUE"))
                    Call ssPOEntry.SetText(26, intLoopCounter, rsdb.GetValue("HSNSACCODE"))
                    Call ssPOEntry.SetText(27, intLoopCounter, rsdb.GetValue("CGSTTXRT_TYPE"))
                    Call ssPOEntry.SetText(28, intLoopCounter, rsdb.GetValue("SGSTTXRT_TYPE"))
                    Call ssPOEntry.SetText(29, intLoopCounter, rsdb.GetValue("UTGSTTXRT_TYPE"))
                    Call ssPOEntry.SetText(30, intLoopCounter, rsdb.GetValue("IGSTTXRT_TYPE"))
                    Call ssPOEntry.SetText(32, intLoopCounter, rsdb.GetValue("ISHSNORSAC"))
                    Call ssPOEntry.SetText(33, intLoopCounter, rsdb.GetValue("external_salesorder_no"))
                    rsdb.MoveNext()
                    prgItemDetails.Value = prgItemDetails.Value + 1
                Next
                prgItemDetails.Visible = False
                CopySSPoEntry.Visible = False : ssPOEntry.Visible = True
            Else
                Call ConfirmWindow(10128, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                ssPOEntry.MaxRows = 0
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
            cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = True
            Exit Sub 'This is to avoid the execution of the error handler

        Catch ex As Exception

            MsgBox(ex.Message)
        Finally
            mblnIsExecuting = False
ErrHandler:
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End Try

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
            .Col = 2
            .Focus()
        End With
        With ssPOEntry
            .BlockMode = True
            .Col = 3
            .Col2 = 3
            .Row = 1
            .Row2 = .MaxRows
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
            .TypeButtonPicture = My.Resources.resEmpower.ico111.ToBitmap()
            .BlockMode = False
        End With
        For inti = 5 To 8
            Call ssPOEntry.SetText(inti, ssPOEntry.MaxRows, 0)
        Next
        For inti = 11 To 12
            Call ssPOEntry.SetText(inti, ssPOEntry.MaxRows, 0)
        Next
        Call ssPOEntry.SetText(19, ssPOEntry.MaxRows, 0)
        Call SSMaxLength()
        'With ssPOEntry
        '    .Col = 1
        '    If .MaxRows > 1 Then
        '        .Row = .MaxRows
        '        .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
        '    End If
        'End With
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
        Dim rstAD As New ClsResultSetDB
        Dim strSalesTaxType As String
        If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Or cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            m_strSql = "select a.SalesTax_Type from cust_ord_hdr a,cust_ord_dtl b where A.UNIT_CODE=B.UNIT_CODE AND a.Amendment_No=b.Amendment_No and a.cust_Ref=b.Cust_Ref and a.Account_Code=b.Account_Code and a.account_code='" & txtCustomerCode.Text.Trim & "' and a.Cust_ref='" & txtReferenceNo.Text & "' and a.amendment_No='" & txtAmendmentNo.Text & "' AND A.unit_code='" & gstrUNITID & "' "
            rstAD.GetResult(m_strSql)
            If rstAD.GetNoRows > 0 Then
                strSalesTaxType = rstAD.GetValue("SalesTax_Type")
            End If
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
    '    Public Sub GetReferenceDetails()
    '        '----------------------------------------------------------------------------
    '        'Argument       :   Nil
    '        'Return Value   :   Nil
    '        'Function       :   Get the additional details
    '        'Comments       :   Nil
    '        '----------------------------------------------------------------------------
    '        On Error GoTo ErrHandler
    '        FillDataTables()
    '        Dim rsAD As New ClsResultSetDB
    '        Dim rscurrency As ClsResultSetDB
    '        Dim intLoopCounter As Short
    '        Dim intMaxCounter As Short
    '        Dim intDecimal As Short
    '        Dim strMax As String
    '        Dim strMin As String
    '        m_strSql = "select * from cust_ord_dtl where unit_code='" & gstrUNITID & "' and Cust_ref='" & txtReferenceNo.Text & " 'and account_Code='" & txtCustomerCode.Text & "' and amendment_No = '" & txtAmendmentNo.Text & "'" ' and active_Flag IN('A','L') order by Cust_drgNo"
    '        rsAD.GetResult(m_strSql)
    '        m_strSql = "select * from cust_ord_hdr where unit_code='" & gstrUNITID & "' and Cust_ref='" & txtReferenceNo.Text & " 'and account_Code='" & txtCustomerCode.Text & "'and active_Flag in ('A','L')"
    '        rsRefNo.GetResult(m_strSql)
    '        If rsAD.GetNoRows > 0 Then
    '            rsAD.MoveFirst()
    '            txtAmendmentNo.Text = " "
    '            lblIntSONoDes.Text = rsRefNo.GetValue("InternalSONo")
    '            lblRevisionNo.Text = rsRefNo.GetValue("RevisionNo")
    '            If Len(Trim(txtAmendmentNo.Text)) > 0 Then
    '                DTAmendmentDate.Value = rsRefNo.GetValue("Amendment_Date")
    '            Else
    '                DTAmendmentDate.Value = GetServerDate()
    '            End If

    '            '10736222-changes done for CT2 ARE 3 functionality - to enable CT2 Reqd flag
    '            ChkCT2Reqd.Checked = rsRefNo.GetValue("CT2_Reqd_In_SO")
    '            ChkCT2Reqd.Enabled = False
    '            txtCreditTerms.Text = rsRefNo.GetValue("Term_Payment")
    '            DTEffectiveDate.Value = rsRefNo.GetValue("Effect_Date")
    '            DTValidDate.Value = rsRefNo.GetValue("Valid_Date")
    '            txtCurrencyType.Text = rsRefNo.GetValue("Currency_Code")
    '            ctlPerValue.Text = rsRefNo.GetValue("PerValue")
    '            txtAmendReason.Text = rsRefNo.GetValue("Reason")
    '            DTDate.Value = rsRefNo.GetValue("Order_date")
    '            strpotype = rsRefNo.GetValue("PO_Type")
    '            strSOType = rsRefNo.GetValue("salestax_Type")
    '            strexportType = rsRefNo.GetValue("exportsotype")
    '            If strexportType = "WP" Then
    '                cmbExporttype.SelectedIndex = 1
    '            ElseIf strexportType = "WOP" Then
    '                cmbExporttype.SelectedIndex = 2
    '            Else
    '                cmbExporttype.SelectedIndex = 0
    '            End If

    '            If strpotype = "M" Then
    '                ssPOEntry.Col = 19
    '                ssPOEntry.Col2 = 19
    '                ssPOEntry.ColHidden = False
    '                ssPOEntry.Col = 20
    '                ssPOEntry.Col2 = 20
    '                ssPOEntry.ColHidden = False
    '                ssPOEntry.Col = 21
    '                ssPOEntry.Col2 = 21
    '                ssPOEntry.ColHidden = False
    '            Else
    '                ssPOEntry.Col = 19
    '                ssPOEntry.Col2 = 19
    '                ssPOEntry.ColHidden = True
    '                ssPOEntry.Col = 20
    '                ssPOEntry.Col2 = 20
    '                ssPOEntry.ColHidden = True
    '                ssPOEntry.Col = 21
    '                ssPOEntry.Col2 = 21
    '                ssPOEntry.ColHidden = True
    '            End If

    '            If strexportType = "WP" Then
    '                cmbExporttype.SelectedIndex = 1
    '            ElseIf strexportType = "WOP" Then
    '                cmbExporttype.SelectedIndex = 2
    '            Else
    '                cmbExporttype.SelectedIndex = 0
    '            End If

    '            '10869290
    '            Select Case UCase(strpotype)
    '                Case "O"
    '                    Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 1)
    '                Case "S"
    '                    Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 2)
    '                Case "J"
    '                    Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 3)
    '                Case "E"
    '                    Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 4)
    '                Case "M"
    '                    Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 5)
    '                Case "T"
    '                    Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 6)
    '                Case "A"
    '                    Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 7)
    '                Case "V"
    '                    Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 8)

    '                    ssPOEntry.Col = 19
    '                    ssPOEntry.Col2 = 19
    '                    ssPOEntry.ColHidden = False
    '                    ssPOEntry.Col = 20
    '                    ssPOEntry.Col2 = 20
    '                    ssPOEntry.ColHidden = False
    '                    ssPOEntry.Col = 21
    '                    ssPOEntry.Col2 = 21
    '                    ssPOEntry.ColHidden = False
    '            End Select
    '            txtSTax.Text = strSOType
    '            txtSChSTax.Text = rsRefNo.GetValue("Surcharge_code")
    '            txtAddVAT.Text = IIf(IsDBNull(rsRefNo.GetValue("AddVAT_Type")), "", rsRefNo.GetValue("AddVAT_Type"))
    '            If txtAddVAT.Text.Trim.Length > 1 Then
    '                Call txtAddVAT_Validating(txtAddVAT, New System.ComponentModel.CancelEventArgs(False))
    '            End If
    '            '10869290
    '            txtService.Text = IIf(IsDBNull(rsRefNo.GetValue("SERVICETAX_TYPE")), "", rsRefNo.GetValue("SERVICETAX_TYPE"))
    '            If txtService.Text.Trim.Length > 1 Then
    '                Call txtService_Validating(txtService, New System.ComponentModel.CancelEventArgs(False))
    '            End If

    '            txtSBC.Text = IIf(IsDBNull(rsRefNo.GetValue("SBCTAX_TYPE")), "", rsRefNo.GetValue("SBCTAX_TYPE"))
    '            If txtSBC.Text.Trim.Length > 1 Then
    '                Call txtSBC_Validating(txtService, New System.ComponentModel.CancelEventArgs(False))
    '            End If

    '            txtKKC.Text = IIf(IsDBNull(rsRefNo.GetValue("KKCTAX_TYPE")), "", rsRefNo.GetValue("KKCTAX_TYPE"))
    '            If txtKKC.Text.Trim.Length > 1 Then
    '                Call txtKKC_Validating(txtService, New System.ComponentModel.CancelEventArgs(False))
    '            End If

    '            If rsRefNo.GetValue("OpenSO") = False Then
    '                chkOpenSo.CheckState = System.Windows.Forms.CheckState.Unchecked
    '            Else
    '                chkOpenSo.CheckState = System.Windows.Forms.CheckState.Checked
    '            End If
    '            rsAD.MoveFirst()
    '            ssPOEntry.MaxRows = 0
    '            If cmdButtons.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
    '                If Len(Trim(txtCurrencyType.Text)) Then
    '                    rscurrency = New ClsResultSetDB
    '                    rscurrency.GetResult("Select decimal_Place from Currency_Mst where unit_code='" & gstrUNITID & "' and currency_code ='" & Trim(txtCurrencyType.Text) & "'")
    '                    intDecimal = rscurrency.GetValue("Decimal_Place")
    '                End If
    '                If intDecimal <= 0 Then
    '                    intDecimal = 2
    '                End If
    '                strMin = "0." : strMax = "99999999."
    '                For intLoopCounter = 1 To intDecimal
    '                    strMin = strMin & "0"
    '                    strMax = strMax & "9"
    '                Next
    '                intMaxCounter = rsAD.GetNoRows
    '                prgItemDetails.Value = 0 : prgItemDetails.Minimum = 0 : prgItemDetails.Maximum = intMaxCounter
    '                prgItemDetails.Visible = True
    '                rsAD.MoveFirst()
    '                With ssPOEntry
    '                    For intLoopCounter = 1 To intMaxCounter
    '                        ssPOEntry.MaxRows = ssPOEntry.MaxRows + 1
    '                        m_custItemDesc = rsAD.GetValue("Cust_Drg_Desc")
    '                        If rsAD.GetValue("OpenSO") = False Then
    '                            .Col = 1
    '                            .Row = intLoopCounter
    '                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
    '                            ssPOEntry.Value = 0
    '                        Else
    '                            .Col = 1
    '                            .Row = intLoopCounter
    '                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
    '                            ssPOEntry.Value = 1
    '                        End If
    '                        .Col = 2
    '                        .Row = intLoopCounter
    '                        .TypeMaxEditLen = 30
    '                        Call ssPOEntry.SetText(2, ssPOEntry.MaxRows, rsAD.GetValue("Cust_DrgNo"))
    '                        .Col = 4
    '                        .Row = intLoopCounter
    '                        .TypeMaxEditLen = 16
    '                        Call ssPOEntry.SetText(4, ssPOEntry.MaxRows, rsAD.GetValue("Item_Code "))
    '                        .Col = 5
    '                        .Row = intLoopCounter
    '                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
    '                        .TypeFloatDecimalPlaces = 2
    '                        .TypeFloatMin = "0.00"
    '                        .TypeFloatMax = "9999999.99"
    '                        Call ssPOEntry.SetText(5, ssPOEntry.MaxRows, rsAD.GetValue("Order_Qty"))
    '                        .Col = 13
    '                        .Row = intLoopCounter
    '                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
    '                        .TypeFloatDecimalPlaces = intDecimal
    '                        .TypeFloatMin = strMin
    '                        .TypeFloatMax = strMax
    '                        Call ssPOEntry.SetText(13, ssPOEntry.MaxRows, rsAD.GetValue("Rate"))
    '                        .Col = 6
    '                        .Row = intLoopCounter
    '                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
    '                        .TypeFloatDecimalPlaces = intDecimal
    '                        .TypeFloatMin = strMin
    '                        .TypeFloatMax = strMax
    '                        Call ssPOEntry.SetText(6, ssPOEntry.MaxRows, rsAD.GetValue("Rate") * CDbl(ctlPerValue.Text))
    '                        .Col = 14
    '                        .Row = intLoopCounter
    '                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
    '                        .TypeFloatDecimalPlaces = intDecimal
    '                        .TypeFloatMin = strMin
    '                        .TypeFloatMax = strMax
    '                        Call ssPOEntry.SetText(14, ssPOEntry.MaxRows, rsAD.GetValue("Cust_Mtrl"))
    '                        .Col = 7
    '                        .Row = intLoopCounter
    '                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
    '                        .TypeFloatDecimalPlaces = intDecimal
    '                        .TypeFloatMin = strMin
    '                        .TypeFloatMax = strMax
    '                        Call ssPOEntry.SetText(7, ssPOEntry.MaxRows, rsAD.GetValue("Cust_Mtrl") * CDbl(ctlPerValue.Text))
    '                        .Col = 15
    '                        .Row = intLoopCounter
    '                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
    '                        .TypeFloatDecimalPlaces = 4
    '                        .TypeFloatMin = strMin
    '                        .TypeFloatMax = strMax
    '                        Call ssPOEntry.SetText(15, ssPOEntry.MaxRows, rsAD.GetValue("Tool_Cost"))
    '                        .Col = 8
    '                        .Row = intLoopCounter
    '                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
    '                        .TypeFloatDecimalPlaces = 4
    '                        .TypeFloatMin = strMin
    '                        .TypeFloatMax = "99999999.9999"
    '                        Call ssPOEntry.SetText(8, ssPOEntry.MaxRows, rsAD.GetValue("Tool_Cost") * CDbl(ctlPerValue.Text))
    '                        .Col = 9
    '                        .Row = intLoopCounter
    '                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
    '                        Call ssPOEntry.SetText(9, ssPOEntry.MaxRows, rsAD.GetValue("Packing_Type"))
    '                        .Col = 10
    '                        .Row = intLoopCounter
    '                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
    '                        Call ssPOEntry.SetText(10, ssPOEntry.MaxRows, rsAD.GetValue("Excise_Duty"))
    '                        .Col = 11
    '                        .Row = intLoopCounter
    '                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
    '                        .TypeFloatDecimalPlaces = 2
    '                        .TypeFloatMin = "0.00"
    '                        .TypeFloatMax = "99999999.99"
    '                        Call ssPOEntry.SetText(16, ssPOEntry.MaxRows, rsAD.GetValue("Others") * CDbl(ctlPerValue.Text))
    '                        .Col = 11
    '                        .Row = intLoopCounter
    '                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
    '                        .TypeFloatDecimalPlaces = 2
    '                        .TypeFloatMin = "0.00"
    '                        .TypeFloatMax = "99999999.99"
    '                        Call ssPOEntry.SetText(11, ssPOEntry.MaxRows, rsAD.GetValue("Others"))
    '                        .Col = 18
    '                        .Row = intLoopCounter
    '                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
    '                        Call ssPOEntry.SetText(18, ssPOEntry.MaxRows, rsAD.GetValue("Remarks"))
    '                        .Col = 19
    '                        .Row = intLoopCounter
    '                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
    '                        .TypeFloatDecimalPlaces = intDecimal
    '                        .TypeFloatMin = strMin
    '                        .TypeFloatMax = strMax
    '                        Call ssPOEntry.SetText(19, ssPOEntry.MaxRows, rsAD.GetValue("MRP"))
    '                        .Col = 20
    '                        .Row = intLoopCounter
    '                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
    '                        Call ssPOEntry.SetText(20, ssPOEntry.MaxRows, rsAD.GetValue("abantment_code"))
    '                        .Col = 21
    '                        .Row = intLoopCounter
    '                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
    '                        .TypeFloatDecimalPlaces = intDecimal
    '                        .TypeFloatMin = strMin
    '                        .TypeFloatMax = strMax
    '                        Call ssPOEntry.SetText(21, ssPOEntry.MaxRows, rsAD.GetValue("AccessibleRateforMRP"))

    '                        .Col = 22
    '                        .Row = intLoopCounter
    '                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeComboBox
    '                        Call ssPOEntry.SetText(22, ssPOEntry.MaxRows, rsAD.GetValue("DISCOUNT_TYPE"))
    '                        .TypeComboBoxList = "NONE" & Chr(9) & "[P]ercentage" & Chr(9) & "[V]alue"
    '                        If rsAD.GetValue("DISCOUNT_TYPE") = "[P]ercentage" Then
    '                            .TypeComboBoxCurSel = 1
    '                        ElseIf rsAD.GetValue("DISCOUNT_TYPE") = "[V]alue" Then
    '                            .TypeComboBoxCurSel = 2
    '                        Else
    '                            .TypeComboBoxCurSel = 0
    '                        End If
    '                        .Col = 23
    '                        .Row = intLoopCounter
    '                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
    '                        Call ssPOEntry.SetText(23, ssPOEntry.MaxRows, rsAD.GetValue("DISCOUNT_VALUE"))


    '                        .Col = 24
    '                        .Row = intLoopCounter
    '                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeComboBox
    '                        Call ssPOEntry.SetText(24, ssPOEntry.MaxRows, rsAD.GetValue("MARKUP_TYPE"))
    '                        .TypeComboBoxList = "NONE" & Chr(9) & "[P]ercentage" & Chr(9) & "[V]alue"
    '                        If rsAD.GetValue("MARKUP_TYPE") = "[P]ercentage" Then
    '                            .TypeComboBoxCurSel = 1
    '                        ElseIf rsAD.GetValue("MARKUP_TYPE") = "[V]alue" Then
    '                            .TypeComboBoxCurSel = 2
    '                        Else
    '                            .TypeComboBoxCurSel = 0
    '                        End If
    '                        .Col = 25
    '                        .Row = intLoopCounter
    '                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
    '                        Call ssPOEntry.SetText(25, ssPOEntry.MaxRows, rsAD.GetValue("MARKUP_VALUE"))
    '                        'GST CHANGE
    '                        .Col = 26
    '                        .Row = intLoopCounter
    '                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
    '                        Call ssPOEntry.SetText(26, ssPOEntry.MaxRows, rsAD.GetValue("HSNSACCODE"))
    '                        .Col = 27
    '                        .Row = intLoopCounter
    '                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
    '                        Call ssPOEntry.SetText(27, ssPOEntry.MaxRows, rsAD.GetValue("CGSTTXRT_TYPE"))
    '                        .Col = 28
    '                        .Row = intLoopCounter
    '                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
    '                        Call ssPOEntry.SetText(28, ssPOEntry.MaxRows, rsAD.GetValue("SGSTTXRT_TYPE"))
    '                        .Col = 29
    '                        .Row = intLoopCounter
    '                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
    '                        Call ssPOEntry.SetText(29, ssPOEntry.MaxRows, rsAD.GetValue("UTGSTTXRT_TYPE"))
    '                        .Col = 30
    '                        .Row = intLoopCounter
    '                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
    '                        Call ssPOEntry.SetText(30, ssPOEntry.MaxRows, rsAD.GetValue("IGSTTXRT_TYPE"))
    '                        .Col = 31
    '                        .Row = intLoopCounter
    '                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
    '                        Call ssPOEntry.SetText(31, ssPOEntry.MaxRows, rsAD.GetValue("COMPENSATION_CESS"))
    '                        .Col = 32
    '                        .Row = intLoopCounter
    '                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
    '                        Call ssPOEntry.SetText(32, ssPOEntry.MaxRows, rsAD.GetValue("ISHSNORSAC"))
    '                        Call ssPOEntry.SetText(33, ssPOEntry.MaxRows, rsAD.GetValue("EXTERNAL_SALESORDER_NO"))
    '                        'GST CHANGE
    '                        rsAD.MoveNext()
    '                        prgItemDetails.Value = prgItemDetails.Value + 1
    '                    Next
    '                    prgItemDetails.Visible = False
    '                End With
    '            End If
    '        Else
    '            Call ConfirmWindow(10129, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
    '            Exit Sub
    '        End If
    '        With ssPOEntry
    '            .Enabled = True
    '            .BlockMode = True
    '            .Col = 3
    '            .Col2 = 3
    '            .Row = 1
    '            .Row2 = .MaxRows
    '            .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
    '            .TypeButtonPicture = My.Resources.resEmpower.ico111.ToBitmap()
    '            .BlockMode = False
    '            If cmdButtons.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
    '                .Row = 1
    '                .Row2 = .MaxRows
    '                .Col = 1
    '                .Col2 = .MaxCols
    '                .BlockMode = True
    '                .Lock = True
    '                .BlockMode = False
    '            End If
    '        End With
    '        cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = True
    '        Exit Sub 'This is to avoid the execution of the error handler
    'ErrHandler:
    '        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    '        Exit Sub
    '    End Sub
    Public Sub GetReferenceDetails()
        '----------------------------------------------------------------------------
        'Argument       :   Nil
        'Return Value   :   Nil
        'Function       :   Get the additional details
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        Try

            mblnIsExecuting = True

            FillDataTables()
            Dim rsAD As New ClsResultSetDB
            Dim rscurrency As ClsResultSetDB
            Dim intLoopCounter As Short
            Dim intMaxCounter As Short
            Dim intDecimal As Short
            Dim strMax As String
            Dim strMin As String
            m_strSql = "select * from cust_ord_dtl where unit_code='" & gstrUNITID & "' and Cust_ref='" & txtReferenceNo.Text & " 'and account_Code='" & txtCustomerCode.Text & "' and amendment_No = '" & txtAmendmentNo.Text & "'" ' and active_Flag IN('A','L') order by Cust_drgNo"
            rsAD.GetResult(m_strSql)
            m_strSql = "select * from cust_ord_hdr where unit_code='" & gstrUNITID & "' and Cust_ref='" & txtReferenceNo.Text & " 'and account_Code='" & txtCustomerCode.Text & "'and active_Flag in ('A','L')"
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

                If strpotype = "M" Then
                    ssPOEntry.Col = 19
                    ssPOEntry.Col2 = 19
                    ssPOEntry.ColHidden = False
                    ssPOEntry.Col = 20
                    ssPOEntry.Col2 = 20
                    ssPOEntry.ColHidden = False
                    ssPOEntry.Col = 21
                    ssPOEntry.Col2 = 21
                    ssPOEntry.ColHidden = False
                Else
                    ssPOEntry.Col = 19
                    ssPOEntry.Col2 = 19
                    ssPOEntry.ColHidden = True
                    ssPOEntry.Col = 20
                    ssPOEntry.Col2 = 20
                    ssPOEntry.ColHidden = True
                    ssPOEntry.Col = 21
                    ssPOEntry.Col2 = 21
                    ssPOEntry.ColHidden = True
                End If

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
                    Case "M"
                        Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 5)
                    Case "T"
                        Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 6)
                    Case "A"
                        Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 7)
                    Case "V"
                        Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 8)

                        ssPOEntry.Col = 19
                        ssPOEntry.Col2 = 19
                        ssPOEntry.ColHidden = False
                        ssPOEntry.Col = 20
                        ssPOEntry.Col2 = 20
                        ssPOEntry.ColHidden = False
                        ssPOEntry.Col = 21
                        ssPOEntry.Col2 = 21
                        ssPOEntry.ColHidden = False
                End Select
                txtSTax.Text = strSOType
                txtSChSTax.Text = rsRefNo.GetValue("Surcharge_code")
                txtAddVAT.Text = IIf(IsDBNull(rsRefNo.GetValue("AddVAT_Type")), "", rsRefNo.GetValue("AddVAT_Type"))
                If txtAddVAT.Text.Trim.Length > 1 Then
                    Call txtAddVAT_Validating(txtAddVAT, New System.ComponentModel.CancelEventArgs(False))
                End If
                '10869290
                txtService.Text = IIf(IsDBNull(rsRefNo.GetValue("SERVICETAX_TYPE")), "", rsRefNo.GetValue("SERVICETAX_TYPE"))
                If txtService.Text.Trim.Length > 1 Then
                    Call txtService_Validating(txtService, New System.ComponentModel.CancelEventArgs(False))
                End If

                txtSBC.Text = IIf(IsDBNull(rsRefNo.GetValue("SBCTAX_TYPE")), "", rsRefNo.GetValue("SBCTAX_TYPE"))
                If txtSBC.Text.Trim.Length > 1 Then
                    Call txtSBC_Validating(txtService, New System.ComponentModel.CancelEventArgs(False))
                End If

                txtKKC.Text = IIf(IsDBNull(rsRefNo.GetValue("KKCTAX_TYPE")), "", rsRefNo.GetValue("KKCTAX_TYPE"))
                If txtKKC.Text.Trim.Length > 1 Then
                    Call txtKKC_Validating(txtService, New System.ComponentModel.CancelEventArgs(False))
                End If

                If rsRefNo.GetValue("OpenSO") = False Then
                    chkOpenSo.CheckState = System.Windows.Forms.CheckState.Unchecked
                Else
                    chkOpenSo.CheckState = System.Windows.Forms.CheckState.Checked
                End If
                If mblnCADM_SALES_ORDER_CREATION = True Then
                    txtCADRefNo.Text = SqlConnectionclass.ExecuteScalar("Select TOP 1 CADMORDERID from CADM_SALES_ORDER_CREATION where CADMORDERID ='" & lblIntSONoDes.Text & "' and Unit_code='" & gstrUNITID & "'")
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
                            Call ssPOEntry.SetText(4, ssPOEntry.MaxRows, rsAD.GetValue("Item_Code "))
                            .Col = 5
                            .Row = intLoopCounter
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                            .TypeFloatDecimalPlaces = 2
                            .TypeFloatMin = "0.00"
                            .TypeFloatMax = "9999999.99"
                            Call ssPOEntry.SetText(5, ssPOEntry.MaxRows, rsAD.GetValue("Order_Qty"))
                            .Col = 13
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
                            .Col = 14
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
                            .Col = 15
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
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                            Call ssPOEntry.SetText(9, ssPOEntry.MaxRows, rsAD.GetValue("Packing_Type"))
                            .Col = 10
                            .Row = intLoopCounter
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                            Call ssPOEntry.SetText(10, ssPOEntry.MaxRows, rsAD.GetValue("Excise_Duty"))
                            .Col = 11
                            .Row = intLoopCounter
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                            .TypeFloatDecimalPlaces = 2
                            .TypeFloatMin = "0.00"
                            .TypeFloatMax = "99999999.99"
                            Call ssPOEntry.SetText(16, ssPOEntry.MaxRows, rsAD.GetValue("Others") * CDbl(ctlPerValue.Text))
                            .Col = 11
                            .Row = intLoopCounter
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                            .TypeFloatDecimalPlaces = 2
                            .TypeFloatMin = "0.00"
                            .TypeFloatMax = "99999999.99"
                            Call ssPOEntry.SetText(11, ssPOEntry.MaxRows, rsAD.GetValue("Others"))
                            .Col = 18
                            .Row = intLoopCounter
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                            Call ssPOEntry.SetText(18, ssPOEntry.MaxRows, rsAD.GetValue("Remarks"))
                            .Col = 19
                            .Row = intLoopCounter
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                            .TypeFloatDecimalPlaces = intDecimal
                            .TypeFloatMin = strMin
                            .TypeFloatMax = strMax
                            Call ssPOEntry.SetText(19, ssPOEntry.MaxRows, rsAD.GetValue("MRP"))
                            .Col = 20
                            .Row = intLoopCounter
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                            Call ssPOEntry.SetText(20, ssPOEntry.MaxRows, rsAD.GetValue("abantment_code"))
                            .Col = 21
                            .Row = intLoopCounter
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                            .TypeFloatDecimalPlaces = intDecimal
                            .TypeFloatMin = strMin
                            .TypeFloatMax = strMax
                            Call ssPOEntry.SetText(21, ssPOEntry.MaxRows, rsAD.GetValue("AccessibleRateforMRP"))

                            .Col = 22
                            .Row = intLoopCounter
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeComboBox
                            Call ssPOEntry.SetText(22, ssPOEntry.MaxRows, rsAD.GetValue("DISCOUNT_TYPE"))
                            .TypeComboBoxList = "NONE" & Chr(9) & "[P]ercentage" & Chr(9) & "[V]alue"
                            If rsAD.GetValue("DISCOUNT_TYPE") = "[P]ercentage" Then
                                .TypeComboBoxCurSel = 1
                            ElseIf rsAD.GetValue("DISCOUNT_TYPE") = "[V]alue" Then
                                .TypeComboBoxCurSel = 2
                            Else
                                .TypeComboBoxCurSel = 0
                            End If
                            .Col = 23
                            .Row = intLoopCounter
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                            Call ssPOEntry.SetText(23, ssPOEntry.MaxRows, rsAD.GetValue("DISCOUNT_VALUE"))


                            .Col = 24
                            .Row = intLoopCounter
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeComboBox
                            Call ssPOEntry.SetText(24, ssPOEntry.MaxRows, rsAD.GetValue("MARKUP_TYPE"))
                            .TypeComboBoxList = "NONE" & Chr(9) & "[P]ercentage" & Chr(9) & "[V]alue"
                            If rsAD.GetValue("MARKUP_TYPE") = "[P]ercentage" Then
                                .TypeComboBoxCurSel = 1
                            ElseIf rsAD.GetValue("MARKUP_TYPE") = "[V]alue" Then
                                .TypeComboBoxCurSel = 2
                            Else
                                .TypeComboBoxCurSel = 0
                            End If
                            .Col = 25
                            .Row = intLoopCounter
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                            Call ssPOEntry.SetText(25, ssPOEntry.MaxRows, rsAD.GetValue("MARKUP_VALUE"))
                            'GST CHANGE
                            .Col = 26
                            .Row = intLoopCounter
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                            Call ssPOEntry.SetText(26, ssPOEntry.MaxRows, rsAD.GetValue("HSNSACCODE"))
                            .Col = 27
                            .Row = intLoopCounter
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                            Call ssPOEntry.SetText(27, ssPOEntry.MaxRows, rsAD.GetValue("CGSTTXRT_TYPE"))
                            .Col = 28
                            .Row = intLoopCounter
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                            Call ssPOEntry.SetText(28, ssPOEntry.MaxRows, rsAD.GetValue("SGSTTXRT_TYPE"))
                            .Col = 29
                            .Row = intLoopCounter
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                            Call ssPOEntry.SetText(29, ssPOEntry.MaxRows, rsAD.GetValue("UTGSTTXRT_TYPE"))
                            .Col = 30
                            .Row = intLoopCounter
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                            Call ssPOEntry.SetText(30, ssPOEntry.MaxRows, rsAD.GetValue("IGSTTXRT_TYPE"))
                            .Col = 31
                            .Row = intLoopCounter
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                            Call ssPOEntry.SetText(31, ssPOEntry.MaxRows, rsAD.GetValue("COMPENSATION_CESS"))
                            .Col = 32
                            .Row = intLoopCounter
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                            Call ssPOEntry.SetText(32, ssPOEntry.MaxRows, rsAD.GetValue("ISHSNORSAC"))
                            Call ssPOEntry.SetText(33, ssPOEntry.MaxRows, rsAD.GetValue("EXTERNAL_SALESORDER_NO"))
                            'GST CHANGE
                            rsAD.MoveNext()
                            prgItemDetails.Value = prgItemDetails.Value + 1
                        Next
                        prgItemDetails.Visible = False
                    End With
                End If
            Else
                Call ConfirmWindow(10129, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
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
            cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = True
            Exit Sub 'This is to avoid the execution of the error handler

        Catch ex As Exception
ErrHandler:
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally

            mblnIsExecuting = False
        End Try

        Exit Sub
    End Sub

    Public Sub DeleteRow()
        '----------------------------------------------------------------------------
        'Argument       :   Nil
        'Return Value   :   Nil
        'Function       :   Deletes a row
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim DataTab1 As DataTable
        Dim TMPDr As DataRow()
        Dim intRow As Short
        Dim rsDespatchQuantity As ClsResultSetDB
        Dim varItemCode As Object
        Dim varDrawCode As Object
        Dim strItemCodeList As String = String.Empty
        Dim strCustDrgNoList As String = String.Empty

        DataTab1 = SqlConnectionclass.GetDataTable("Select unit_code,Account_Code,Cust_Ref,Amendment_No,Authorized_flag,Item_Code,Cust_drgNo,Despatch_Qty from cust_ord_dtl where unit_code='" & gstrUNITID & "' and Account_Code='" & txtCustomerCode.Text & "'and Cust_Ref='" & txtReferenceNo.Text & "' and Amendment_No= '" & txtAmendmentNo.Text & "'  and Authorized_flag = 0")

        ReDim ArrDispatchQty(ssPOEntry.MaxRows - 1)
        rsDespatchQuantity = New ClsResultSetDB
        For intRow = 1 To ssPOEntry.MaxRows
            varItemCode = Nothing
            Call ssPOEntry.GetText(4, intRow, varItemCode)

            If strItemCodeList = String.Empty Then
                strItemCodeList = "'" + varItemCode.ToString() + "'"
            Else
                strItemCodeList = strItemCodeList + "," + "'" + varItemCode.ToString() + "'"
            End If


            varDrawCode = Nothing
            Call ssPOEntry.GetText(2, intRow, varDrawCode)

            If strCustDrgNoList = String.Empty Then
                strCustDrgNoList = "'" + varDrawCode.ToString() + "'"
            Else
                strCustDrgNoList = strCustDrgNoList + "," + "'" + varDrawCode.ToString() + "'"
            End If


            'rsDespatchQuantity.GetResult("Select Despatch_Qty from cust_ord_dtl where unit_code='" & gstrUNITID & "' and Account_Code='" & txtCustomerCode.Text & "'and Cust_Ref='" & txtReferenceNo.Text & "' and Amendment_No= '" & txtAmendmentNo.Text & "' and Item_Code= '" & varItemCode & "' and Cust_drgNo ='" & varDrawCode & "' and Authorized_flag = 0")
            'If rsDespatchQuantity.GetNoRows > 0 Then
            '    ArrDispatchQty(intRow - 1) = rsDespatchQuantity.GetValue("Despatch_Qty")
            'Else
            '    ArrDispatchQty(intRow - 1) = 0
            'End If


            TMPDr = DataTab1.Select("unit_code='" & gstrUNITID & "' and Account_Code='" & txtCustomerCode.Text & "' and Cust_Ref='" & txtReferenceNo.Text & "' and Amendment_No= '" & txtAmendmentNo.Text & "' and Item_Code= '" & varItemCode & "' and Cust_drgNo ='" & varDrawCode & "' and Authorized_flag = 0")
            If TMPDr.Length > 0 Then
                ArrDispatchQty(intRow - 1) = TMPDr(0)("Despatch_Qty").ToString()
            Else
                ArrDispatchQty(intRow - 1) = 0
            End If

            'strSql = "delete cust_ord_dtl where unit_code='" & gstrUNITID & "' and Account_Code='"
            'strSql = strSql & txtCustomerCode.Text & "'and Cust_Ref='"
            'strSql = strSql & txtReferenceNo.Text & "' and Amendment_No= '"
            'strSql = strSql & txtAmendmentNo.Text & "' and Item_Code= '"
            'strSql = strSql & varItemCode & "' and Cust_drgNo ='" & varDrawCode & "' and Authorized_flag = 0"
            'mP_Connection.Execute(strSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

        Next

        If strItemCodeList <> String.Empty And strCustDrgNoList <> String.Empty Then
            strSql = "delete cust_ord_dtl where unit_code='" & gstrUNITID & "' and Account_Code='"
            strSql = strSql & txtCustomerCode.Text & "'and Cust_Ref='"
            strSql = strSql & txtReferenceNo.Text & "' and Amendment_No= '"
            strSql = strSql & txtAmendmentNo.Text & "'"
            mP_Connection.Execute(strSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
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
        'Revision History   :    Changes done by Ashutosh on 10-10-2005 , Issue Id:15876, Bug fix of SO entry form , If Item is saved as closed one but it still saved as Open.
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim varDeleteFlag As Object
        Dim clsDrgDes As ClsResultSetDB
        Dim strSql As String
        Dim varExtraExciseDuty, varSalestax, varToolCost, varordqty, varItemCode, varRate, varCustSuppMaterial, varPkg, varExiseDuty, varSurchargeSalesTax As Object
        Dim varAbatment_code, varMRP, varAccessibleRateforMRP, varexternalsalesorder As Object
        Dim varOthers, varCustItemDesc, varCustItemCode, varDespatchQty, varOpenSO As Object
        Dim varRemarks As Object
        Dim strSql1 As String
        Dim dblPackingPer As Double
        Dim varDiscountvalue As Object
        Dim varDiscounttype As Object
        Dim varMarkupvalue As Object
        Dim varMarkuptype As Object
        'GST CHANGES
        Dim VARHSNSACCODE As Object
        Dim VARISHSNORSAC As Object
        Dim VARCGSTTXRT_HEAD As Object
        Dim VARSGSTTXRT_HEAD As Object
        Dim VARIGSTTXRT_HEAD As Object
        Dim VARUGSTTXRT_HEAD As Object
        Dim VARCOMPENSATIONCESS_HEAD As Object
        Dim mblnWithPayExpInvoice As Boolean = False
        Dim dtCustDrg As DataTable
        Dim dtTaxPercent As DataTable
        'GST CHANGES

        strSql = "Select isnull(WithPayExpInvoice,0) from sales_parameter (nolock) where unit_code='" & gstrUNITID & "'"
        mblnWithPayExpInvoice = Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSql))

        strSql = "SELECT distinct drg_desc,account_code,Cust_drgno,Item_code FROM  custitem_mst (nolock) WHERE unit_code = '" & gstrUNITID & "' and account_code='" & txtCustomerCode.Text.Trim & "' and active=1 "
        dtCustDrg = SqlConnectionclass.GetDataTable(strSql)

        strSql = "SELECT  TxRt_Rate_No ,TxRt_Percentage FROM Gen_TaxRate (nolock) WHERE  unit_code = '" & gstrUNITID & "'"
        dtTaxPercent = SqlConnectionclass.GetDataTable(strSql)

        For intRow = 1 To ssPOEntry.MaxRows
            varDeleteFlag = Nothing
            Call ssPOEntry.GetText(0, intRow, varDeleteFlag)
            If varDeleteFlag <> "*" Then 'to get the values from the grid
                varCustItemCode = Nothing
                Call ssPOEntry.GetText(2, intRow, varCustItemCode)
                'Getting the Drawing No. Description

                'issue id 10117810
                varItemCode = Nothing
                Call ssPOEntry.GetText(4, intRow, varItemCode)
                'strSql = "SELECT drg_desc FROM  custitem_mst WHERE unit_code='" & gstrUNITID & "' and Cust_drgno = '" & Trim(varCustItemCode) & "'"
                'strSql = "SELECT drg_desc FROM  custitem_mst WHERE unit_code = '" & gstrUNITID & "' and account_code='" & txtCustomerCode.Text.Trim & "' and Cust_drgno = '" & Trim(varCustItemCode) & "' and item_code= '" & Trim(varItemCode) & "' and active=1 "
                ''issue id 10117810

                'clsDrgDes = New ClsResultSetDB
                'If clsDrgDes.GetResult(strSql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly) And clsDrgDes.GetNoRows > 0 Then
                '    clsDrgDes.MoveFirst()
                '    m_custItemDesc = Trim(clsDrgDes.GetValue("drg_desc"))
                '    clsDrgDes.ResultSetClose()
                '    clsDrgDes = Nothing
                'End If
                Dim strCustDesc As DataRow() = dtCustDrg.Select("Cust_drgno = '" & Trim(varCustItemCode) & "' and item_code='" & Trim(varItemCode) & "'")
                If strCustDesc.Length > 0 Then
                    m_custItemDesc = strCustDesc(0)("drg_desc").ToString
                Else
                    m_custItemDesc = ""
                End If


                varOpenSO = Nothing
                Call ssPOEntry.GetText(1, intRow, varOpenSO)
                'issue id 10117810 comment starts
                'varItemCode = Nothing
                'Call ssPOEntry.GetText(4, intRow, varItemCode)
                'issue id 10117810 comment end 
                varordqty = Nothing
                Call ssPOEntry.GetText(5, intRow, varordqty)
                varRate = Nothing
                Call ssPOEntry.GetText(6, intRow, varRate)
                If CDbl(ctlPerValue.Text) >= 1 Then
                    varRate = varRate / CDbl(ctlPerValue.Text)
                End If
                varCustSuppMaterial = Nothing
                Call ssPOEntry.GetText(7, intRow, varCustSuppMaterial)
                varToolCost = Nothing
                Call ssPOEntry.GetText(8, intRow, varToolCost)
                varPkg = Nothing
                Call ssPOEntry.GetText(9, intRow, varPkg)
                'clsDrgDes = New ClsResultSetDB
                'strSql1 = "SELECT TxRt_Rate_No ,TxRt_Percentage FROM Gen_TaxRate WHERE unit_code='" & gstrUNITID & "' and TxRt_Rate_No ='" & Trim(varPkg) & "'"
                'If clsDrgDes.GetResult(strSql1, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly) Then
                '    If clsDrgDes.GetNoRows > 0 Then
                '        clsDrgDes.MoveFirst()
                '        dblPackingPer = clsDrgDes.GetValue("TxRt_Percentage")
                '    Else
                '        dblPackingPer = 0
                '    End If
                '    clsDrgDes.ResultSetClose()
                '    clsDrgDes = Nothing
                'End If

                Dim strTaxPercent As DataRow() = dtTaxPercent.Select("TxRt_Rate_No = '" & Trim(varPkg) & "'")
                If strTaxPercent.Length > 0 Then
                    dblPackingPer = strTaxPercent(0)("TxRt_Percentage").ToString
                Else
                    dblPackingPer = 0
                End If
                varExiseDuty = Nothing
                Call ssPOEntry.GetText(10, intRow, varExiseDuty)
                varOthers = Nothing
                Call ssPOEntry.GetText(11, intRow, varOthers)
                varRemarks = Nothing
                Call ssPOEntry.GetText(18, intRow, varRemarks)
                varMRP = Nothing
                Call ssPOEntry.GetText(19, intRow, varMRP)
                varAbatment_code = Nothing
                Call ssPOEntry.GetText(20, intRow, varAbatment_code)
                varAccessibleRateforMRP = Nothing
                Call ssPOEntry.GetText(21, intRow, varAccessibleRateforMRP)
                If Val(varAccessibleRateforMRP) = 0 Then varAccessibleRateforMRP = 0

                varDiscounttype = Nothing
                Call ssPOEntry.GetText(22, intRow, varDiscounttype)
                varDiscountvalue = Nothing
                Call ssPOEntry.GetText(23, intRow, varDiscountvalue)
                varMarkuptype = Nothing
                Call ssPOEntry.GetText(24, intRow, varMarkuptype)
                varMarkupvalue = Nothing
                Call ssPOEntry.GetText(25, intRow, varMarkupvalue)
                ''ADDED BY SUMIT KUMAR ON 9 JULY 2019
                varexternalsalesorder = Nothing
                Call ssPOEntry.GetText(33, intRow, varexternalsalesorder)

                If (varDiscounttype = "NONE") Then
                    varDiscountvalue = "0"
                End If

                If mblnDiscountFunctionality = False Then
                    varDiscounttype = "-"
                    varDiscountvalue = 0
                End If

                If (varMarkuptype = "NONE") Then
                    varMarkupvalue = "0"
                End If
                '10561117  
                If mblnMarkupFunctionality = False Then
                    varMarkuptype = "-"
                    varMarkupvalue = 0
                End If
                '10561117  end 
                'gst changes
                If gblnGSTUnit = True Then
                    VARHSNSACCODE = Nothing
                    Call ssPOEntry.GetText(26, intRow, VARHSNSACCODE)
                    VARCGSTTXRT_HEAD = Nothing
                    Call ssPOEntry.GetText(27, intRow, VARCGSTTXRT_HEAD)
                    VARSGSTTXRT_HEAD = Nothing
                    Call ssPOEntry.GetText(28, intRow, VARSGSTTXRT_HEAD)
                    VARUGSTTXRT_HEAD = Nothing
                    Call ssPOEntry.GetText(29, intRow, VARUGSTTXRT_HEAD)
                    VARIGSTTXRT_HEAD = Nothing
                    Call ssPOEntry.GetText(30, intRow, VARIGSTTXRT_HEAD)
                    VARCOMPENSATIONCESS_HEAD = Nothing
                    Call ssPOEntry.GetText(31, intRow, VARCOMPENSATIONCESS_HEAD)
                    VARISHSNORSAC = Nothing
                    Call ssPOEntry.GetText(32, intRow, VARISHSNORSAC)
                End If
                'gst changes

                If mblnWithPayExpInvoice = False Then
                    'VARHSNSACCODE = Nothing
                    'VARCGSTTXRT_HEAD = Nothing
                    'VARSGSTTXRT_HEAD = Nothing
                    'VARUGSTTXRT_HEAD = Nothing
                    'VARIGSTTXRT_HEAD = Nothing
                    'VARCOMPENSATIONCESS_HEAD = Nothing
                Else
                    If mblnWithPayExpInvoice = True Then
                        If (UCase(cmbExporttype.Text) = "WITHOUT PAY") Then
                            'VARHSNSACCODE = Nothing
                            VARCGSTTXRT_HEAD = Nothing
                            VARSGSTTXRT_HEAD = Nothing
                            VARUGSTTXRT_HEAD = Nothing
                            VARIGSTTXRT_HEAD = Nothing
                            VARCOMPENSATIONCESS_HEAD = Nothing
                        ElseIf (UCase(cmbExporttype.Text) = "None") Then
                            'VARHSNSACCODE = Nothing
                            VARCGSTTXRT_HEAD = Nothing
                            VARSGSTTXRT_HEAD = Nothing
                            VARUGSTTXRT_HEAD = Nothing
                            VARIGSTTXRT_HEAD = Nothing
                            VARCOMPENSATIONCESS_HEAD = Nothing
                        End If
                    End If
                End If

                If cmbPOType.Text <> "MRP-SPARES" Then
                    varMRP = "0"
                    varAbatment_code = ""
                End If
                strSql = "Insert into Cust_Ord_Dtl (UNIT_CODE,Account_Code,Packing_Type, Cust_Ref, Amendment_No,InternalSONo,RevisionNo, "
                strSql = strSql & "Item_Code , Rate, Order_Qty, Despatch_Qty, "
                strSql = strSql & "Active_Flag, Cust_Mtrl, Cust_DrgNo, Packing, Others,"
                strSql = strSql & "Excise_Duty,"
                strSql = strSql & "Remarks,MRP,abantment_code,AccessibleRateforMRP,"
                strSql = strSql & "Cust_Drg_Desc,"
                strSql = strSql & "Tool_Cost, Authorized_flag, openSO,Ent_dt, Ent_UserId, Upd_dt, Upd_UserId,PerValue,Discount_type,Discount_value,Markup_type,Markup_value,ISHSNORSAC,HSNSACCODE,CGSTTXRT_TYPE,SGSTTXRT_TYPE,UTGSTTXRT_TYPE,IGSTTXRT_TYPE,COMPENSATION_CESS,external_salesorder_no)"
                strSql = strSql & " values('" & gstrUNITID & "','" & Trim(txtCustomerCode.Text) & "','"
                strSql = strSql & Trim(varPkg) & "','"
                strSql = strSql & Trim(txtReferenceNo.Text) & "','"
                strSql = strSql & Trim(txtAmendmentNo.Text) & "','"
                strSql = strSql & Trim(lblIntSONoDes.Text) & "'," & Trim(lblRevisionNo.Text) & ",'"
                strSql = strSql & Trim(varItemCode) & "',"
                strSql = strSql & Trim(varRate) & ","
                strSql = strSql & IIf(IsNothing(varordqty), 0, varordqty) & "," & ArrDispatchQty(intRow - 1) & ",'A',"
                strSql = strSql & IIf(IsNothing(varCustSuppMaterial), 0, varCustSuppMaterial) & ",'"
                strSql = strSql & varCustItemCode & "'," & dblPackingPer & ","
                strSql = strSql & Val(varOthers) & ",'"
                strSql = strSql & IIf(IsNothing(varExiseDuty), 0, varExiseDuty) & "','"
                strSql = strSql & Trim(varRemarks) & "',"
                strSql = strSql & IIf(Trim(varMRP) = "", 0, Trim(varMRP)) & ",'"
                strSql = strSql & Trim(varAbatment_code) & "',"
                If gblnGSTUnit = True Then
                    strSql = strSql & (Trim(varRate)) & ",'"
                Else
                    strSql = strSql & (Trim(varAccessibleRateforMRP)) & ",'"

                End If
                strSql = strSql & Trim(m_custItemDesc) & "',"
                strSql = strSql & IIf(IsNothing(varToolCost), 0, varToolCost) & ",0,"
                If chkOpenSo.CheckState = System.Windows.Forms.CheckState.Checked Then
                    strSql = strSql & "1,"
                Else
                    If Val(varOpenSO) = 0 Or varOpenSO = "" Then
                        strSql = strSql & "0,"
                    Else
                        strSql = strSql & "1,"
                    End If
                End If
                strSql = strSql & " getdate()" & ",'" & mP_User & "'," & "getdate() "
                strSql = strSql & ",'" & mP_User & "'"

                If CDbl(ctlPerValue.Text) >= 1 Then
                    strSql = strSql & "," & ctlPerValue.Text & ",'" & varDiscounttype & "'," & varDiscountvalue & ",'" & varMarkuptype & "'," & varMarkupvalue
                Else
                    strSql = strSql & ", 1 ,'" & varDiscounttype & "'," & varDiscountvalue & ",'" & varMarkuptype & "'," & varMarkupvalue
                End If
                strSql = strSql & ",'" & VARISHSNORSAC & "','" & VARHSNSACCODE & "','" & VARCGSTTXRT_HEAD & "','" & VARSGSTTXRT_HEAD & "','" & VARUGSTTXRT_HEAD & "','" & VARIGSTTXRT_HEAD & "','" & VARCOMPENSATIONCESS_HEAD & "' ,'" & varexternalsalesorder & "')"
                mP_Connection.Execute(strSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End If
        Next
        clsDrgDes = Nothing
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
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
        strSql = "update cust_ord_hdr set Order_Date='"
        strSql = strSql & getDateForDB(DTDate.Value) & "',"
        If Len(Trim(txtAmendmentNo.Text)) > 0 Then
            strSql = strSql & "Amendment_Date='"
            strSql = strSql & getDateForDB(DTAmendmentDate.Value) & "',"
        End If
        strSql = strSql & "Currency_Code='"
        strSql = strSql & txtCurrencyType.Text & "',Valid_Date='"
        strSql = strSql & getDateForDB(DTValidDate.Value) & "',Effect_Date='"
        strSql = strSql & getDateForDB(DTEffectiveDate.Value) & "',Term_Payment='"
        strSql = strSql & Trim(txtCreditTerms.Text) & "',Special_Remarks='" & m_strSpecialNotes & "',Pay_Remarks='"
        strSql = strSql & m_strPaymentTerms & "',Price_Remarks='" & m_strPricesAre & "',Packing_Remarks='"
        strSql = strSql & m_strPkgAndFwd & "',Frieght_Remarks='" & m_strFreight & "',Transport_Remarks='"
        strSql = strSql & m_strTransitInsurance & "',Octorai_Remarks='" & m_strOctroi & "',Mode_Despatch='"
        strSql = strSql & m_strModeOfDespatch & "',Delivery='" & m_strDeliverySchedule & "',"
        strSql = strSql & "Reason='" & txtAmendReason.Text & "',PO_Type='"
        strSql = strSql & Mid(cmbPOType.Text, 1, 1) & "',"
        If chkOpenSo.CheckState = System.Windows.Forms.CheckState.Checked Then
            strSql = strSql & " OpenSO = 1,"
        Else
            strSql = strSql & " OpenSO = 0,"
        End If

        '10736222-changes done for CT2 ARE 3 functionality - to enable CT2 Reqd flag
        If ChkCT2Reqd.CheckState = System.Windows.Forms.CheckState.Checked Then
            strSql = strSql & " CT2_Reqd_In_SO = 1,"
        Else
            strSql = strSql & " CT2_Reqd_In_SO = 0,"
        End If

        strSql = strSql & " SalesTax_type = '" & Trim(txtSTax.Text) & "', future_so = 0,"
        strSql = strSql & " Ent_dt="
        strSql = strSql & " getdate() " & ",Ent_UserId='" & mP_User & "',Upd_dt="
        strSql = strSql & " getdate() " & ",Upd_UserId='" & mP_User & "'"
        strSql = strSql & " , Surcharge_Code = '" & Trim(txtSChSTax.Text) & "'"
        strSql = strSql & " , AddVAT_Type = '" & Trim(txtAddVAT.Text) & "'"
        strSql = strSql & " , SERVICETAX_TYPE = '" & Trim(txtService.Text) & "'"
        strSql = strSql & " , SBCTAX_TYPE = '" & Trim(txtSBC.Text) & "'"
        strSql = strSql & " , KKCTAX_TYPE = '" & Trim(txtKKC.Text) & "'"
        If CDbl(ctlPerValue.Text) >= 1 Then
            strSql = strSql & ", PerValue = " & ctlPerValue.Text & " where unit_code='" & gstrUNITID & "' and Account_Code='"
        Else
            strSql = strSql & ", PerValue = 1 where unit_code='" & gstrUNITID & "' and Account_Code='"
        End If
        strSql = strSql & txtCustomerCode.Text & "'and Cust_Ref='"
        strSql = strSql & txtReferenceNo.Text & "'and Amendment_No='" & txtAmendmentNo.Text & "'"
        mP_Connection.Execute("Set dateformat 'mdy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        mP_Connection.Execute(strSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
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
        strSql = "Insert into Cust_Ord_Hdr (unit_code,Account_Code,Consignee_code ,Cust_Ref, Amendment_No,InternalSONo,RevisionNo, Order_Date, "
        strSql = strSql & "Amendment_Date, Active_Flag, "
        strSql = strSql & " Currency_Code, Valid_Date,"
        strSql = strSql & "Effect_Date, Term_Payment, Special_Remarks, Pay_Remarks, "
        strSql = strSql & "Price_Remarks, Packing_Remarks, Frieght_Remarks, Transport_Remarks,"
        strSql = strSql & "Octorai_Remarks, Mode_Despatch, Delivery, First_Authorized,"
        strSql = strSql & "Second_Authorized, Third_Authorized, Authorized_Flag, Reason, "
        '10736222-changes done for CT2 ARE 3 functionality - to enable CT2 Reqd flag
        'If cmbPOType.Text.ToUpper = "EXPORT" Then
        strSql = strSql & "PO_Type, SalesTax_Type,AddVAT_Type,OpenSO,CT2_Reqd_In_SO,Ent_dt, Ent_UserId, Upd_dt, Upd_UserId,PerValue, Surcharge_Code,SERVICETAX_TYPE,SBCTAX_TYPE,KKCTAX_TYPE,ShipAddress_Code,ShipAddress_Desc,ExportSotype)"
        'Else
        'strSql = strSql & "PO_Type, SalesTax_Type,AddVAT_Type,OpenSO,CT2_Reqd_In_SO,Ent_dt, Ent_UserId, Upd_dt, Upd_UserId,PerValue, Surcharge_Code,SERVICETAX_TYPE,SBCTAX_TYPE,KKCTAX_TYPE,ShipAddress_Code,ShipAddress_Desc)"
        'End If

        strSql = strSql & " Values('" & gstrUNITID & "','" & Trim(txtCustomerCode.Text) & "','" & Trim(txtCustomerCode.Text) & "','" & Trim(txtReferenceNo.Text) & "','" & Trim(txtAmendmentNo.Text) & "',"
        strSql = strSql & "'" & lblIntSONoDes.Text & "',"
        If Len(Trim(txtAmendmentNo.Text)) > 0 Then
            lblRevisionNo.Text = CStr(GenerateRevisionNo())
            strSql = strSql & lblRevisionNo.Text & ",'"
        Else
            lblRevisionNo.Text = "0"
            strSql = strSql & "0,'"
        End If
        strSql = strSql & getDateForDB(DTDate.Value) & "','" & IIf(Len(Me.txtAmendmentNo.Text) = 0, System.DBNull.Value, getDateForDB(DTAmendmentDate.Value)) & "','A' ,'"
        strSql = strSql & Trim(txtCurrencyType.Text) & "','" & getDateForDB(DTValidDate.Value) & "','" & getDateForDB(DTEffectiveDate.Value) & "','"
        strSql = strSql & IIf(Len(Trim(txtCreditTerms.Text)) = 0, 0, Trim(txtCreditTerms.Text)) & "','"
        strSql = strSql & Trim(m_strSpecialNotes) & "','"
        strSql = strSql & Trim(m_strPaymentTerms) & "','" & Trim(m_strPricesAre) & "','" & Trim(m_strPkgAndFwd) & "','" & Trim(m_strFreight) & "','"
        strSql = strSql & Trim(m_strTransitInsurance) & "','" & Trim(m_strOctroi) & "','" & Trim(m_strModeOfDespatch) & "','" & Trim(m_strDeliverySchedule) & "','',"
        strSql = strSql & "'','','','" & Trim(txtAmendReason.Text) & "','" & Mid(cmbPOType.Text, 1, 1) & "','" & Trim(txtSTax.Text) & "','" & Trim(txtAddVAT.Text) & "',"
        If chkOpenSo.CheckState = System.Windows.Forms.CheckState.Checked Then
            strSql = strSql & "1,"
        Else
            strSql = strSql & "0,"
        End If

        '10736222-changes done for CT2 ARE 3 functionality - to enable CT2 Reqd flag
        If ChkCT2Reqd.CheckState = System.Windows.Forms.CheckState.Checked Then
            strSql = strSql & "1,"
        Else
            strSql = strSql & "0,"
        End If

        strSql = strSql & " getdate() " & ",'" & mP_User & "'," & " getdate() " & ",'" & mP_User & "'"
        If CDbl(Val(ctlPerValue.Text)) >= 1 Then
            strSql = strSql & "," & ctlPerValue.Text & ",'" & Trim(txtSChSTax.Text) & "','"
        Else
            strSql = strSql & ", 1,'" & Trim(txtSChSTax.Text) & "','"
        End If

        'strSql = strSql & Trim(txtService.Text) & "','" & Trim(txtSBC.Text) & "','" & Trim(txtKKC.Text) & "')"
        If UCase(cmbExporttype.Text) = "WITHOUT PAY" Then
            strexportsotext = "WOP"
        ElseIf UCase(cmbExporttype.Text) = "WITH PAY" Then
            strexportsotext = "WP"
        Else
            strexportsotext = ""
        End If

        'If cmbPOType.Text.ToUpper = "EXPORT" Then
        strSql = strSql & Trim(txtService.Text) & "','" & Trim(txtSBC.Text) & "','" & Trim(txtKKC.Text) & "','" & ship_address_code & "','" & ship_address_desc & "','" & strexportsotext & "')"
        '        Else
        '       strSql = strSql & Trim(txtService.Text) & "','" & Trim(txtSBC.Text) & "','" & Trim(txtKKC.Text) & "','" & ship_address_code & "','" & ship_address_desc & "')"
        '      End If

        mP_Connection.Execute("Set dateformat 'mdy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        mP_Connection.Execute(strSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)

        '' Added by priti for CAD functionality on 11 Nov 2024
        If Len(txtCADRefNo.Text) > 0 Then
            Dim strSQl As String = "Update CADM_SALES_ORDER_CREATION SET CADMORDERID='" & lblIntSONoDes.Text & "' Where unit_code='" & gstrUNITID & "' and CUST_REF='" & txtReferenceNo.Text & "' "
            SqlConnectionclass.ExecuteNonQuery(strSQl)
        End If

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
        Dim rstAD As New ClsResultSetDB
        '***Changed by nisha on 06/09/2002 to remove S.Tax Type and Credit Days from Sub Form
        If Len(Trim(m_strPaymentTerms)) = 0 Then
            '***
            If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Or cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                m_strSql = "select Special_Remarks,Pay_Remarks,Price_Remarks,Packing_Remarks,Frieght_Remarks,Transport_Remarks,Octorai_Remarks,Mode_Despatch,Delivery from cust_ord_hdr a,cust_ord_dtl b where a.unit_code=b.unit_code and a.Amendment_No=b.Amendment_No and a.cust_Ref=b.Cust_Ref and a.Account_Code=b.Account_Code and a.unit_code='" & gstrUNITID & "' and a.Cust_ref='" & txtReferenceNo.Text & "' and a.amendment_No='" & txtAmendmentNo.Text & "' and a.account_code='" & txtCustomerCode.Text & "'"
            ElseIf cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                m_strSql = "select Special_Remarks,Pay_Remarks,Price_Remarks,Packing_Remarks,Frieght_Remarks,Transport_Remarks,Octorai_Remarks,Mode_Despatch,Delivery from cust_ord_hdr where unit_code='" & gstrUNITID & "' and Cust_ref='" & txtReferenceNo.Text & "' and account_code='" & txtCustomerCode.Text & "'"
            End If
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
            Else
                CheckFormDetails = False
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
        Dim rsAD As New ClsResultSetDB
        Dim strAuthFlg As String
        m_strSql = "select Account_Code,Cust_Ref,Amendment_No,Item_Code,Rate,Order_Qty,Despatch_Qty,Active_Flag,Cust_Mtrl,Cust_DrgNo,Packing,Others,Excise_Duty,Cust_Drg_Desc,Tool_Cost,Authorized_flag,OpenSO,SalesTax_Type,PerValue,InternalSONo,RevisionNo,Remarks,MRP,Abantment_Code,AccessibleRateforMRP,Packing_Type,TOOL_AMOR_FLAG,ShowInAuth,ADD_Excise_Duty,HSNSACCODE,  ISHSNORSAC, CGSTTXRT_TYPE,SGSTTXRT_TYPE, UTGSTTXRT_TYPE,IGSTTXRT_TYPE,COMPENSATION_CESS   from cust_ord_dtl where unit_code='" & gstrUNITID & "' and Cust_ref='" & txtReferenceNo.Text & "' and account_Code='" & txtCustomerCode.Text & "' and Authorized_Flag=1 order by Cust_drgNo"
        rsdb.GetResult(m_strSql)
        m_strSql = "select Account_Code,Cust_Ref,Amendment_No,Order_Date,Amendment_Date,Active_Flag,Currency_Code,Valid_Date,Effect_Date,Term_Payment,Special_Remarks,Pay_Remarks,Price_Remarks,Packing_Remarks,Frieght_Remarks,Transport_Remarks,Octorai_Remarks,Mode_Despatch,Delivery,First_Authorized,Second_Authorized,Third_Authorized,Authorized_Flag,Reason,PO_Type,SalesTax_Type,OpenSO,AddCustSupp,PerValue,InternalSONo,RevisionNo,Surcharge_code,Future_SO,ECESS_Code,Consignee_Code,ADDVAT_TYPE from cust_ord_hdr where unit_code='" & gstrUNITID & "' and Cust_ref='" & txtReferenceNo.Text & "' and account_Code='" & txtCustomerCode.Text & "' and authorized_Flag=1"
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
            strexportType = rsAD.GetValue("ExportSotype")
            If strexportType = "WP" Then
                cmbExporttype.SelectedIndex = 1
            ElseIf strexportType = "WOP" Then
                cmbExporttype.SelectedIndex = 2
            Else
                cmbExporttype.SelectedIndex = 0
            End If
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
                Case "T"
                    Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 6)
                Case "A"
                    Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 7)
                Case "V"
                    Me.cmbPOType.Text = ObsoleteManagement.GetItemString(Me.cmbPOType, 8)

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
                Call ssPOEntry.SetText(9, ssPOEntry.MaxRows, rsdb.GetValue("Packing_type"))
                Call ssPOEntry.SetText(10, ssPOEntry.MaxRows, rsdb.GetValue("Excise_Duty"))
                Call ssPOEntry.SetText(16, ssPOEntry.MaxRows, rsdb.GetValue("Others"))
                Call ssPOEntry.SetText(11, ssPOEntry.MaxRows, rsdb.GetValue("Others") * CDbl(ctlPerValue.Text))
                rsdb.MoveNext()
            Loop
        Else
            Call ConfirmWindow(10130, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
            ssPOEntry.MaxRows = 0
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
        Dim rsAD As New ClsResultSetDB
        Dim varItemCode As Object
        Dim varDrgNo As Object
        Dim AmendmentNo As String
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        Dim intmaxitems As Short
        Dim rsCustOrdHdr As ClsResultSetDB
        m_strSql = "select top 1 1 from cust_ord_dtl where unit_code='" & gstrUNITID & "' and Cust_ref='" & txtReferenceNo.Text & " 'and account_Code='" & txtCustomerCode.Text & "' and active_Flag='A'"
        rsAD.GetResult(m_strSql)
        intmaxitems = rsAD.GetNoRows
        intMaxLoop = ssPOEntry.MaxRows
        ReDim ArrDispatchQty(intMaxLoop - 1)
        For intLoopCounter = 1 To intMaxLoop
            varItemCode = Nothing
            Call ssPOEntry.GetText(4, intLoopCounter, varItemCode)
            varDrgNo = Nothing
            Call ssPOEntry.GetText(2, intLoopCounter, varDrgNo)
            m_strSql = "select Despatch_qty from cust_ord_dtl where unit_code='" & gstrUNITID & "' and Cust_ref='" & txtReferenceNo.Text & "' and account_Code='" & txtCustomerCode.Text & "' and active_Flag='A' and Item_Code ='" & varItemCode & "' and Cust_drgNo ='" & varDrgNo & "'"
            rsAD.GetResult(m_strSql)
            If rsAD.GetNoRows >= 1 Then
                ArrDispatchQty(intLoopCounter - 1) = rsAD.GetValue("Despatch_qty")
                m_strSql = "update cust_ord_dtl set Active_Flag='O' where Cust_ref='" & txtReferenceNo.Text & "' and account_Code='" & txtCustomerCode.Text & "' and active_Flag='A' and Item_Code='" & varItemCode & "' and Cust_drgNo ='" & varDrgNo & "'"
                mP_Connection.Execute(m_strSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            Else
                ArrDispatchQty(intLoopCounter - 1) = 0
            End If
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
            If Len(Trim(txtCurrencyType.Text)) Then
                rscurrency = New ClsResultSetDB
                rscurrency.GetResult("Select decimal_Place from Currency_Mst (Nolock) where unit_code='" & gstrUNITID & "' and currency_code ='" & Trim(txtCurrencyType.Text) & "'")
                intDecimal = rscurrency.GetValue("Decimal_Place")
            End If
            For intRow = 1 To .MaxRows
                .Row = intRow
                .Col = 1
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
                .Col = 2
                .TypeMaxEditLen = 50
                .Focus()
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
                'If Len(Trim(txtCurrencyType.Text)) Then
                '    rscurrency = New ClsResultSetDB
                '    rscurrency.GetResult("Select decimal_Place from Currency_Mst where unit_code='" & gstrUNITID & "' and currency_code ='" & Trim(txtCurrencyType.Text) & "'")
                '    intDecimal = rscurrency.GetValue("Decimal_Place")
                'End If
                If intDecimal <= 0 Then
                    intDecimal = 2
                End If
                strMin = "0." : strMax = "99999999."
                For intLoopCounter = 1 To intDecimal
                    strMin = strMin & "0"
                    strMax = strMax & "9"
                Next
                '****************
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
                .Row = intRow
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                .Col = 10
                .Row = intRow
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                .Col = 18
                .Row = intRow
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                .Col = 19
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                .TypeFloatDecimalPlaces = intDecimal
                .TypeFloatMin = strMin
                .TypeFloatMax = strMax
                .Col = 21
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                .TypeFloatDecimalPlaces = intDecimal
                .TypeFloatMin = strMin
                .TypeFloatMax = strMax
                If mblnDiscountFunctionality = True Then
                    CmbDiscounttype.Enabled = True
                    CmbDiscounttype.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    txtdiscountvalue.Enabled = True
                    txtdiscountvalue.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)

                    .Col = 22
                    .Row = .MaxRows
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeComboBox
                    .TypeComboBoxList = "NONE" & Chr(9) & "[P]ercentage" & Chr(9) & "[V]alue"
                    If Me.CmbDiscounttype.Text = "[P]ercentage" Then
                        .TypeComboBoxCurSel = 1
                    ElseIf CmbDiscounttype.Text = "[V]alue" Then
                        .TypeComboBoxCurSel = 2
                    Else
                        .TypeComboBoxCurSel = 0
                    End If

                    .Col = 23
                    .Row = .MaxRows
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                    .TypeFloatDecimalPlaces = intDecimal
                    .TypeFloatMin = strMin
                    .TypeFloatMax = strMax
                    If .TypeComboBoxCurSel = 0 And txtdiscountvalue.Text.Trim.Length <= 0 Then
                        Call .SetText(23, .MaxRows, "0")
                    Else
                        Call .SetText(23, .MaxRows, txtdiscountvalue.Text.Trim)
                    End If
                End If
                '10561117  
                If mblnMarkupFunctionality = True Then

                    CmbMarkuptype.Enabled = True
                    CmbMarkuptype.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    txtmarkupvalue.Enabled = True
                    txtmarkupvalue.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)

                    .Col = 24
                    .Row = intRow
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeComboBox
                    .TypeComboBoxList = "NONE" & Chr(9) & "[P]ercentage" & Chr(9) & "[V]alue"
                    If Me.CmbMarkuptype.Text = "[P]ercentage" Then
                        .TypeComboBoxCurSel = 1
                    ElseIf CmbMarkuptype.Text = "[V]alue" Then
                        .TypeComboBoxCurSel = 2
                    Else
                        .TypeComboBoxCurSel = 0
                    End If

                    .Col = 25
                    .Row = intRow
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                    .TypeFloatDecimalPlaces = intDecimal
                    .TypeFloatMin = strMin
                    .TypeFloatMax = strMax
                    If .TypeComboBoxCurSel = 0 And txtmarkupvalue.Text.Trim.Length <= 0 Then
                        Call .SetText(25, .MaxRows, "0")
                    Else
                        Call .SetText(25, .MaxRows, txtmarkupvalue.Text.Trim)
                    End If

                End If
                '10561117  end 
                'GST CHANGES
                If gblnGSTUnit = True Then
                    .Col = 26
                    .Row = .MaxRows
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                End If
                'GST CHANGES
            Next intRow
        End With
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Public Sub RefreshForm()
        Dim rsGetDate As ClsResultSetDB
        On Error GoTo Err_Handler
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
        rsGetDate.GetResult("select Financial_EndDate from Company_Mst where unit_code='" & gstrUNITID & "' ")
        DTValidDate.Value = rsGetDate.GetValue("Financial_EndDate")
        ssPOEntry.MaxRows = 0
        m_strSalesTaxType = ""
        lblCustPartDesc.Text = ""
        '10736222-changes done for CT2 ARE 3 functionality
        ChkCT2Reqd.Enabled = False
        ChkCT2Reqd.Checked = False
        cmbExporttype.SelectedIndex = 0
        FillDataTables(True)
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Sub AddPOType()
        '10869290
        On Error GoTo Err_Handler
        cmbPOType.Items.Insert(0, "None")
        cmbPOType.Items.Insert(1, "OEM")
        cmbPOType.Items.Insert(2, "SPARES")
        cmbPOType.Items.Insert(3, "JOB WORK")
        cmbPOType.Items.Insert(4, "Export")
        cmbPOType.Items.Insert(5, "MRP-SPARES")
        cmbPOType.Items.Insert(6, "Q-TRADING")
        cmbPOType.Items.Insert(7, "A-SUB ASSEMBLY")
        cmbPOType.Items.Insert(8, "V-SERVICE")
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
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

    Public Function ValidRecord() As Boolean

        Dim strErrMsg As String
        Dim ctlBlank As System.Windows.Forms.Control
        Dim lNo As Integer
        Dim varDrawPartNo1, varItemCode1, varOpenSO As Object
        Dim strSql As String
        Dim strCustPartNo As String = String.Empty
        Dim blnIsEopRequired As Boolean = False

        On Error GoTo Err_Handler

        blnInvalidData = False
        ValidRecord = False
        lNo = 1
        strErrMsg = ResolveResString(10059) & vbCrLf & vbCrLf
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
        ' '' code ends here
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
        If cmbPOType.SelectedIndex <> 4 And GetPlantName() <> "HILEX" Then
            If CheckFormDetails() = False Then
                blnInvalidData = True
                strErrMsg = strErrMsg & vbCrLf & lNo & "." & "Sales Terms "
                lNo = lNo + 1
                If ctlBlank Is Nothing Then ctlBlank = cmdchangetype
            End If
        End If

        If cmbPOType.Text.Trim = "Q-TRADING" Then
            If chkOpenSo.Checked = True Then
                blnInvalidData = True
                strErrMsg = strErrMsg & vbCrLf & lNo & "." & "Open Sales order Not Possible"
                lNo = lNo + 1
                If ctlBlank Is Nothing Then ctlBlank = cmbPOType
            End If
        End If
        '10869290
        If UCase(Trim(cmbPOType.Text)) = "V-SERVICE".ToUpper And gblnGSTUnit = False Then
            If Len(Trim(txtService.Text)) = 0 Then
                blnInvalidData = True
                strErrMsg = strErrMsg & vbCrLf & lNo & "." & "Service Tax "
                lNo = lNo + 1
                If ctlBlank Is Nothing Then ctlBlank = txtService
            End If
        End If

        If UCase(Trim(cmbPOType.Text)) = "V-SERVICE".ToUpper And gblnGSTUnit = False Then
            If Len(Trim(txtSBC.Text)) = 0 Then
                blnInvalidData = True
                strErrMsg = strErrMsg & vbCrLf & lNo & "." & "SBC Tax Code "
                lNo = lNo + 1
                If ctlBlank Is Nothing Then ctlBlank = txtSBC
            End If

            If Len(Trim(txtKKC.Text)) = 0 Then
                blnInvalidData = True
                strErrMsg = strErrMsg & vbCrLf & lNo & "." & "KKC Tax Code "
                lNo = lNo + 1
                If ctlBlank Is Nothing Then ctlBlank = txtKKC
            End If
        End If

        strErrMsg = VB.Left(strErrMsg, Len(strErrMsg) - 1)
        strErrMsg = strErrMsg & " ."
        lNo = lNo + 1
        blnValidAmendDate = True


        If mblnDiscountMandatory = True Then
            Dim intdiscountvalue, intLoopCount As Integer
            Dim strdiscountype As String
            For intLoopCount = 1 To ssPOEntry.MaxRows
                With Me.ssPOEntry
                    .Col = 22
                    .Row = intLoopCount
                    strdiscountype = .Text
                    .Col = 23
                    .Row = intLoopCount
                    intdiscountvalue = Val(.Text)
                    If (strdiscountype = "NONE" Or strdiscountype = "") Then
                        MsgBox("Please select Disocunt type. ", MsgBoxStyle.OkOnly, ResolveResString(100))
                        blnValidAmendDate = False
                        Exit Function
                    End If
                End With
            Next
        End If

        If Len(Trim(txtAmendmentNo.Text)) >= 1 Then
            If DTAmendmentDate.Value > DTValidDate.Value Then
                Call ConfirmWindow(10146, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                blnValidAmendDate = False
                Me.DTValidDate.Focus()
                Exit Function
            Else
                blnValidAmendDate = True
            End If
            If DTAmendmentDate.Value < DTDate.Value Then
                Call ConfirmWindow(10147, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                blnValidAmendDate = False
                Me.DTValidDate.Focus()
                Exit Function
            Else
                blnValidAmendDate = True
            End If
            'on requirement of MATE for Back Date SO Entry
            If DTAmendmentDate.Value > GetServerDate() Then
                MsgBox("Amendment Date Can not be Greater Than Current Date", MsgBoxStyle.Information, ResolveResString(100))
                blnValidAmendDate = False
                Me.DTValidDate.Focus()
                Exit Function
            Else
                blnValidAmendDate = True
            End If
        End If

        If blnInvalidData = True And blnValidAmendDate = True Then
            gblnCancelUnload = True
            Call MsgBox(strErrMsg, MsgBoxStyle.Information, "Error")
            ctlBlank.Focus()
            Exit Function
        End If

        If mblnDiscountFunctionality = True Then
            Dim intdiscountvalue, intLoopCount As Integer
            Dim strdiscountype As String
            For intLoopCount = 1 To ssPOEntry.MaxRows
                With Me.ssPOEntry
                    .Col = 22
                    .Row = intLoopCount
                    strdiscountype = .Text
                    .Col = 23
                    .Row = intLoopCount
                    intdiscountvalue = Val(.Text)
                    If strdiscountype = "[P]ercentage" And intdiscountvalue > 100 Then
                        MsgBox("% Can not be Greater Than 100 ", MsgBoxStyle.Information, ResolveResString(100))
                        blnValidAmendDate = False
                        Exit Function
                    End If
                End With
            Next
        End If

        If mblnMarkupFunctionality = True Then
            Dim intmarkupvalue, intLoopCount As Integer
            Dim strmarkuptype As String

            For intLoopCount = 1 To ssPOEntry.MaxRows
                With Me.ssPOEntry
                    .Col = 24
                    .Row = intLoopCount
                    strmarkuptype = .Text
                    .Col = 25
                    .Row = intLoopCount
                    intmarkupvalue = Val(.Text)
                    If strmarkuptype = "[P]ercentage" And intmarkupvalue > 100 Then
                        MsgBox("% Can not be Greater Than 100 ", MsgBoxStyle.Information, ResolveResString(100))
                        blnValidAmendDate = False
                        Exit Function
                    End If
                End With
            Next
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

                        strSql = "SELECT TOP 1 CUST_DRGNO FROM BUDGETITEM_MST WHERE UNIT_CODE= '" & gstrUNITID & "' AND ACCOUNT_CODE = '" & Trim(txtCustomerCode.Text) & "' " &
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

        'Anupam Kumar 26112025 -Start   'Check already exists cust_ref with amendmend No before save
        If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD AndAlso Not String.IsNullOrWhiteSpace(txtReferenceNo.Text) Then
            Dim query As String = "SELECT TOP 1 1 FROM CUST_ORD_HDR WHERE UNIT_CODE='" & gstrUNITID & "' AND ACCOUNT_CODE='" & Trim(txtCustomerCode.Text) & "' AND CUST_REF='" & Trim(txtReferenceNo.Text) & "' AND AMENDMENT_NO='" & Trim(txtAmendmentNo.Text) & "'"

            Dim result As Object = SqlConnectionclass.ExecuteScalar(query)
            If result IsNot Nothing Then
                If (String.IsNullOrWhiteSpace(txtAmendmentNo.Text)) Then
                    MessageBox.Show("Given Refrence No is already Saved.", "Empro", MessageBoxButtons.OK)
                Else
                    MessageBox.Show("Given Refrence No and Amendment No are already Saved.", "Empro", MessageBoxButtons.OK)
                End If
                Exit Function
            End If
        End If
        'Anupam Kumar 26112025 -End

        '10808160--Ends
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
    Public Function addExciseDuty(ByRef pstrRow As Integer) As Boolean
        Dim varItem_Code As Object
        Dim strExDuty As String
        Dim rsTariffCode As ClsResultSetDB
        varItem_Code = Nothing

        Call ssPOEntry.GetText(4, pstrRow, varItem_Code)
        rsTariffCode = New ClsResultSetDB
        rsTariffCode.GetResult("Select a.Excise_Duty from Tax_tariff_Mst a,ITem_Mst b where a.unit_code=b.unit_code and b.Tariff_Code = a.Tariff_subhead and a.unit_code='" & gstrUNITID & "' and b.Item_code ='" & varItem_Code & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsTariffCode.GetNoRows > 0 Then
            'GST CHANGES
            strExDuty = rsTariffCode.GetValue("Excise_duty")
            If gblnGSTUnit = False Or (gblnGSTUnit And _blnEOUFlag) Then
                Call ssPOEntry.SetText(10, pstrRow, strExDuty)
            Else
                Call ssPOEntry.SetText(10, pstrRow, "")
            End If
            addExciseDuty = True
            'GST CHANGES

        Else
            MsgBox("Tariff - Item Relationship Not Defined in Tariff Master.", MsgBoxStyle.Information, ResolveResString(100))
            addExciseDuty = False

        End If
    End Function
    Public Sub PrintToReport()
        '*********************************************'
        'Author:                Ananya Nath
        'Arguments:             None
        'Return Value   :       None
        'Description    :       Used to print currently selected/entered sales Order.
        '*********************************************'
        Dim strReportName As String
        On Error GoTo ErrHandler
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.AppStarting)
        Dim frmRpt As New eMProCrystalReportViewer
        Dim CR As ReportDocument
        CR = frmRpt.GetReportDocument()
        With frmRpt
            '**********************Making a string for record selection.
            strSql = ""
            ' 10610274 STARTS :USE TRIM FUNCTION IN ALL TEXT BUTTONS 
            If Len(Trim(Me.txtAmendmentNo.Text)) = 0 Then
                strSql = " {cust_ord_hdr.unit_code}='" & gstrUNITID & "' and {cust_ord_hdr.account_code} = '" & txtCustomerCode.Text.Trim & "' and {cust_ord_hdr.cust_ref} = '" & txtReferenceNo.Text.Trim & "' and  {cust_ord_hdr.amendment_no} = ''" 'Initialising Sql Query.
            Else
                strSql = " {cust_ord_hdr.unit_code}='" & gstrUNITID & "' and {cust_ord_hdr.account_code} = '" & txtCustomerCode.Text.Trim & "' and {cust_ord_hdr.cust_ref} = '" & txtReferenceNo.Text.Trim & "' and  {cust_ord_hdr.amendment_no} = '" & txtAmendmentNo.Text.Trim & "'" 'Initialising Sql Query.
            End If
            '************************' 10610274 
            strReportName = GetPlantName()
            strReportName = "\Reports\rptSOPrinting_" & strReportName & ".rpt"
            If Not CheckFile(strReportName) Then
                strReportName = "\Reports\rptSOPrinting.rpt"
            End If
            'CrystalReport1.ReportFileName = My.Application.Info.DirectoryPath & strReportName
            CR.Load(My.Application.Info.DirectoryPath & strReportName)
            .ReportHeader = "Customer Purchase Order"
            'Me.CrystalReport1.set_Formulas(1, "Comp_name = '" & gstrCOMPANY & "'") ' company name will be printed
            CR.DataDefinition.FormulaFields("Comp_name").Text = "'" + gstrCOMPANY + "'"
            'Me.CrystalReport1.set_Formulas(2, "Comp_address = '" & gstr_WRK_ADDRESS1 & " ' + '" & gstr_WRK_ADDRESS2 & " ' ") ' address will be printed
            CR.DataDefinition.FormulaFields("Comp_address").Text = "'" + gstr_WRK_ADDRESS1 + gstr_WRK_ADDRESS2 + "'"
            'Me.CrystalReport1.WindowShowExportBtn = True
            .ShowExportButton = True
            'Me.CrystalReport1.WindowMaxButton = False
            'Me.CrystalReport1.WindowMinButton = False
            '12/12/2002 Added by nisha issue log no 1399
            'Me.CrystalReport1.WindowShowPrintSetupBtn = True
            'Me.CrystalReport1.WindowShowSearchBtn = True
            .ShowTextSearchButton = True
            'Me.CrystalReport1.SelectionFormula = strSql
            CR.RecordSelectionFormula = strSql
            'Me.CrystalReport1.WindowTitle = "Customer Purchase Order"
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
    Private Sub TxtSearch_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtsearch.TextChanged
        Call SearchItem()
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
            m_strSql = "select TxRt_Rate_No,TxRt_RateDesc from Gen_TaxRate where unit_code='" & gstrUNITID & "' and Txrt_Rate_No = '" & txtSTax.Text & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
            rsdb.GetResult(m_strSql)
            If rsdb.GetNoRows > 0 Then
                Call FillLabel("STAX")
            Else
                cmdButtons.Focus()
                MsgBox("Entered S.Tax Code does not exist", MsgBoxStyle.Information, ResolveResString(100))
                txtSTax.Text = ""
                txtSTax.Focus()
            End If
        End If
    End Sub
    '****added by Ajay on 18/07/2003
    Private Sub txtSChSTax_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSChSTax.Leave
        If Len(Trim(txtSChSTax.Text)) <> 0 Then
            m_strSql = "select TxRt_Rate_No,TxRt_RateDesc from Gen_TaxRate where unit_code='" & gstrUNITID & "' and Txrt_Rate_No = '" & txtSChSTax.Text & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
            rsdb.GetResult(m_strSql)
            If rsdb.GetNoRows > 0 Then
                Call FillLabel("SSCHTAX")
            Else
                cmdButtons.Focus()
                MsgBox("Entered S.Tax Surcharge Code does not exist", MsgBoxStyle.Information, ResolveResString(100))
                txtSChSTax.Text = ""
                txtSChSTax.Focus()
            End If
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
        End With
    End Function
    Public Sub SetCellTypeCombo(ByVal intRow As Short)
        Dim strcustdtl As String
        Dim strItemCode As Object
        Dim strDrgNo As Object
        Dim FinalstrItemCode As String
        Dim rsitem As ClsResultSetDB
        rsitem = New ClsResultSetDB
        strDrgNo = Nothing
        Call ssPOEntry.GetText(2, intRow, strDrgNo)
        strcustdtl = "SElect ITem_code from custITem_Mst where  unit_code='" & gstrUNITID & "' and cust_drgNo ='" & strDrgNo & "'and Account_code ='" & txtCustomerCode.Text & "'"
        rsitem.GetResult(strcustdtl)
        With ssPOEntry
            .Col = 4
            .Row = intRow
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeComboBox
            .TypeComboBoxClear(4, .Row)
            rsitem.MoveFirst()
            FinalstrItemCode = ""
            While Not rsitem.EOFRecord
                strItemCode = IIf((rsitem.GetValue("ITem_code") = "Unknown"), "", rsitem.GetValue("ITem_code"))
                FinalstrItemCode = FinalstrItemCode & strItemCode & Chr(9) '& "[V]alue":
                rsitem.MoveNext()
            End While
            FinalstrItemCode = VB.Left(FinalstrItemCode, Len(FinalstrItemCode) - 1)
            .TypeComboBoxList = FinalstrItemCode
        End With
    End Sub
    Public Sub SetCellStatic(ByRef intRow As Integer)
        With ssPOEntry
            .Col = 4
            .Row = intRow
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
        End With
    End Sub
    Public Function chkmultipleitem() As Boolean
        Dim strItemCode As Object
        Dim strDrgNo As Object
        Dim rsitem As ClsResultSetDB
        rsitem = New ClsResultSetDB
        strDrgNo = Nothing
        Call ssPOEntry.GetText(2, ssPOEntry.ActiveRow, strDrgNo)
        strItemCode = "Select top 1 1 from custITem_Mst where  unit_code='" & gstrUNITID & "' and cust_drgNo ='" & strDrgNo & "' and Account_code ='" & txtCustomerCode.Text & "'"
        rsitem.GetResult(strItemCode)
        If rsitem.GetNoRows > 1 Then
            chkmultipleitem = True
        Else
            chkmultipleitem = False
        End If
    End Function
    Public Sub RowDetailsfromKeyBoard(ByRef pstrItemCode As Object, ByRef pstrDrgno As Object)
        Dim rsitem As ClsResultSetDB
        If cmdButtons.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            If Len(Trim(pstrItemCode)) > 0 Then
                If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                    'ISSUE ID 10515727
                    m_strSql = "Select * from Cust_ord_dtl where unit_code='" & gstrUNITID & "' and Item_code ='" & pstrItemCode & "' and cust_drgNo ='" & pstrDrgno & "' and cust_ref ='" & txtReferenceNo.Text & "' and amendment_no='" & Trim(txtAmendmentNo.Text) & "' and Active_flag ='A' and account_code='" & txtCustomerCode.Text & "'"
                Else
                    m_strSql = "Select * from Cust_ord_dtl where unit_code='" & gstrUNITID & "' and Item_code ='" & pstrItemCode & "' and cust_drgNo ='" & pstrDrgno & "' and cust_ref ='" & txtReferenceNo.Text & "' and Active_flag ='A' and Authorized_Flag =1 and account_code='" & txtCustomerCode.Text & "'"
                End If
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
                    Call ssPOEntry.SetText(9, ssPOEntry.MaxRows, rsitem.GetValue("Packing_Type"))
                    Call ssPOEntry.SetText(10, ssPOEntry.MaxRows, rsitem.GetValue("Excise_Duty"))
                    Call ssPOEntry.SetText(16, ssPOEntry.MaxRows, rsitem.GetValue("Others"))
                    Call ssPOEntry.SetText(11, ssPOEntry.MaxRows, rsitem.GetValue("Others") * CDbl(ctlPerValue.Text))
                Else
                    If txtAmendmentNo.Enabled = False Then
                        If Len(Trim(pstrItemCode)) > 0 Then
                            m_strSql = " select * from ITemRate_Mst where unit_code='" & gstrUNITID & "' and Serial_No = (select max(serial_no) from itemrate_mst where unit_code='" & gstrUNITID & "' and Party_code = '" & txtCustomerCode.Text & "' and item_code = '" & pstrDrgno & "' and datediff(mm,convert(varchar(10),'" & getDateForDB(DTDate.Value) & "',103),convert(varchar(10),DateFrom,103))<=0 and custVend_Flg ='C')"
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
        pstrWantedDate = VB6.Format(pstrWantedDate, gstrDateFormat)
        If Len(Trim(pstrWantedDate)) > 0 Then 'For Post Dated Docs
            'No need to check for Previously made documents for After Dates
            mP_Connection.Execute("Set dateformat 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            rsDocumentNoSO.GetResult("Select DocNo = Max(convert(int,substring(" & pstrDocNofield & ",9,7))) from " & pstrTableName & " Where unit_code='" & gstrUNITID & "' and datePart(mm, ent_dt) = datePart(mm,'" & getDateForDB(pstrWantedDate) & "') and datePart(yyyy,ent_dt) = datePart(yyyy,'" & getDateForDB(pstrWantedDate) & "')")
            strCheckDOcNo = rsDocumentNoSO.GetValue("DocNo")
        End If
        If Len(Trim(strCheckDOcNo)) > 0 Then 'That is the Document is Made for that Period
            'Add 1 to it
            strTempSeries = CStr(CDbl(strCheckDOcNo) + 1)
            If Val(strTempSeries) < 9999 Then
                strTempSeries = New String("0", 4 - Len(strTempSeries)) & strTempSeries 'Concatenate Zeroes before the Number
            End If
            strCheckDOcNo = CStr(DatePart(Microsoft.VisualBasic.DateInterval.Year, ConvertToDate(pstrWantedDate))) & "-"
            strCheckDOcNo = strCheckDOcNo & New String("0", 2 - Len(CStr(DatePart(Microsoft.VisualBasic.DateInterval.Month, ConvertToDate(pstrWantedDate))))) & CStr(DatePart(Microsoft.VisualBasic.DateInterval.Month, ConvertToDate(pstrWantedDate))) & "-"
            strCheckDOcNo = strCheckDOcNo & strTempSeries
            GenerateDocumentNumber = strCheckDOcNo
        Else 'The Document has not been made for that Period
            NewTempSeries = NewTempSeries & CStr(DatePart(Microsoft.VisualBasic.DateInterval.Year, ConvertToDate(pstrWantedDate))) & "-"
            NewTempSeries = NewTempSeries & New String("0", 2 - Len(CStr(DatePart(Microsoft.VisualBasic.DateInterval.Month, ConvertToDate(pstrWantedDate))))) & CStr(DatePart(Microsoft.VisualBasic.DateInterval.Month, ConvertToDate(pstrWantedDate))) & "-"
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
        rsRevisionNo.ResultSetClose()
        rsRevisionNo = Nothing
    End Function
    Public Sub InsertPreviousSODetails(ByRef pstrAccountCode As String, ByRef pstrRef As String, ByRef pstrAmendment As String, ByRef pstrInternalSONo As String, ByRef pintRevisionNo As Short)
        '*********************************************'
        'Author:                Nisha Rai
        'Arguments:             Account_code , CustRef,Amendment_no , IntSONo,RevisionNo
        'Return Value   :       None
        'Description    :       To Insert active item details from base SO & its amendment which are not there in Grid.
        '*********************************************'
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
        strSql = strSql & " OpenSO,SalesTax_Type,PerValue,InternalSONo,RevisionNo,EXTERNAL_SALESORDER_NO,Packing_Type,HSNSACCODE,ISHSNORSAC,TXRT_HEAD,TXRT_TYPE,TXRT_PER)"
        strSql = strSql & " (Select DISTINCT rate,'" & gstrUNITID & "',Account_Code,Cust_Ref, Amendment_No = '" & pstrAmendment & "',Item_Code,Order_Qty,Despatch_Qty = 0 ,"
        strSql = strSql & " Active_Flag ,Cust_Mtrl,Cust_DrgNo,Packing,Others, Excise_Duty,Cust_Drg_Desc,Tool_Cost,Authorized_flag = 0 "
        strSql = strSql & " ,getdate(),'" & mP_User & "',getdate(),'" & mP_User & "',OpenSO,SalesTax_Type,PerValue,InternalSONo = '" & pstrInternalSONo & "',"
        strSql = strSql & " RevisionNo = " & pintRevisionNo & ",EXTERNAL_SALESORDER_NO,Packing_Type,HSNSACCODE,ISHSNORSAC,TXRT_HEAD,TXRT_TYPE,TXRT_PER from Cust_ord_dtl where  unit_code='" & gstrUNITID & "' and Account_code = '" & pstrAccountCode & "' "
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
                strSql = strSql & strDrgItem & ")"
            End If
        End If
        'STRUPDATERATESQL = " UPDATE A SET A.RATE= B.RATE FROM " & _
        '                   " Cust_Ord_Dtl  A INNER JOIN (" & _
        '               " SELECT CH.UNIT_CODE, CH.Cust_Ref ,CH.Account_Code,Item_Code, " & _
        '              " Cust_DrgNo,RATE,MAX(Amendment_Date ) Amendment_Date  FROM CUST_ORD_DTL CD INNER JOIN Cust_Ord_Hdr CH " & _
        '         " ON CH.UNIT_CODE =CD.UNIT_CODE AND CH.Account_Code =CD.Account_Code AND CH.Cust_Ref =CD.Cust_Ref" & _
        '    " AND CH.Amendment_No =CD.Amendment_No  WHERE CH.UNIT_CODE='" & gstrUNITID & "' AND CH.ACCOUNT_CODE = '" & pstrAccountCode & "' AND CH.CUST_REF = '" & pstrRef & "' AND CH.ACTIVE_FLAG = 'A' AND CH.AUTHORIZED_FLAG = 1 " & _
        '   " AND CD.Active_Flag ='A' GROUP BY CH.UNIT_CODE,CH.Cust_Ref ,CH.Account_Code,Item_Code,Cust_DrgNo,RATE  ) B" & _
        ' " ON A.UNIT_CODE =B.UNIT_CODE AND A.Account_Code =B.Account_Code AND A.Cust_Ref =B.Cust_Ref AND A.Item_Code =B.Item_Code AND A.Cust_DrgNo =B.Cust_DrgNo " & _
        '" AND A.UNIT_CODE='" & gstrUNITID & "' AND A.ACCOUNT_CODE = '" & pstrAccountCode & "' AND A.CUST_REF = '" & pstrRef & "' AND A.ACTIVE_FLAG = 'A' AND AUTHORIZED_FLAG = 0   AND A.AMENDMENT_NO = '" & pstrAmendment & "'"
        '
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
                    rsCustDrgDesc.GetResult("Select Drg_desc from CustItem_Mst where  unit_code='" & gstrUNITID & "' and account_code =  '" & Trim(txtCustomerCode.Text) & "' and ITem_code = '" & Trim(varItemCode) & "' and Cust_drgNo = '" & Trim(varCPartCode) & "'")
                    If rsCustDrgDesc.GetNoRows > 0 Then
                        rsCustDrgDesc.MoveFirst()
                        lblCustPartDesc.Text = rsCustDrgDesc.GetValue("Drg_desc")
                        ToSetcustdrgDesc = rsCustDrgDesc.GetValue("Drg_desc")
                    Else
                        lblCustPartDesc.Text = ""
                        ToSetcustdrgDesc = ""
                    End If
                End If
            End If
        End With
    End Function
    Private Sub SearchItem()
        '---------------------------------------------------------------------
        'Created By     -   Arshad Ali
        '---------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim intCount As Short
        With ssPOEntry
            .Row = -1
            .Col = -1
            .Font = VB6.FontChangeBold(.Font, False)
            If optPartNo.Checked Then
                .Col = 2
            End If
            If optItem.Checked Then
                .Col = 4
            End If
            For intCount = 1 To .MaxRows
                .Row = intCount
                If UCase(Mid(.Text, 1, Len(txtsearch.Text))) = UCase(txtsearch.Text) Then
                    .TopRow = .Row
                    .Col = -1
                    .Font = VB6.FontChangeBold(.Font, True)
                    Exit Sub
                End If
            Next
        End With
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub InitializeSpreed()
        On Error GoTo ErrHandler
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        With Me.ssPOEntry
            .Row = 0
            .Col = 19 : .Text = "MRP"
            .Row = 0
            .Col = 20 : .Text = "Abatment"
            .Row = 0
            .Col = 21 : .Text = "Accessible Rate"
            .CursorStyle = FPSpreadADO.CursorStyleConstants.CursorStyleArrow
        End With
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        Exit Sub
ErrHandler:
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub cmdButtons_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.UCbtndgrp.ButtonClickEventArgs) Handles cmdButtons.ButtonClick
        Dim varExtraExciseDuty, varSalestax, varToolCost, varordqty, varItemCode, varRate, varCustSuppMaterial, varPkg, varExiseDuty, varSurchargeSalesTax As Object
        Dim varDespatchQty, varCustItemCode, varCustItemDesc, varOthers As Object
        Dim varDeleteFlag As Object
        Dim intLoop As Short
        Dim intMaxLoop As Short
        Dim arrMain() As String
        Dim arrDet() As String
        Dim strFormSQL As String
        Dim intOuterCount As Short
        Dim rsSalesParameter As ClsResultSetDB
        rsSalesParameter = New ClsResultSetDB
        On Error GoTo ErrHandler
        Dim strSql As String
        Dim strErrMsg As String
        Dim blnInvalidData As Boolean
        Dim intRow As Short
        Dim enmValue As eMPowerFunctions.ConfirmWindowReturnEnum
        Dim varItem As Object
        Dim vardrawing As Object
        blnInvalidData = False
        Dim counter As Short
        Dim vardiscounttype As Object
        Dim vardiscountvalue As Object
        Dim varexternalsalesorder As Object
        Select Case e.Button
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD 'Add Record
                Call EnableControls(False, Me, True)
                txtCustomerCode.Enabled = True
                txtCustomerCode.BackColor = System.Drawing.Color.White
                Me.lblIntSONoDes.Text = ""
                cmdHelp(0).Enabled = True
                Call RefreshForm()
                ssPOEntry.MaxRows = 0
                With ssPOEntry
                    .Col = 19
                    .Col2 = 19
                    .ColHidden = True
                    .Col = 20
                    .Col2 = 20
                    .ColHidden = True
                End With
                m_strSalesTaxType = ""
                txtCustomerCode.Focus()
                cmbPOType.Text = ObsoleteManagement.GetItemString(cmbPOType, 1)
                CmbDiscounttype.Text = ObsoleteManagement.GetItemString(CmbDiscounttype, 0)
                txtdiscountvalue.Text = "0"
                CmbMarkuptype.Text = ObsoleteManagement.GetItemString(CmbMarkuptype, 0)
                txtmarkupvalue.Text = "0"
                cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE) = True
                frmSearch.Enabled = True
                optPartNo.Enabled = True
                optPartNo.Checked = True
                optItem.Enabled = True
                txtsearch.Enabled = True
                txtsearch.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                mstrFormDetails = ""
                cmdForms.Font = VB6.FontChangeBold(cmdForms.Font, False)
                chkShipAddress.Enabled = True
                '' Added by priti for CAD functionality on 11 Nov 2024
                If mblnCADM_SALES_ORDER_CREATION = True Then
                    txtCADRefNo.Enabled = False
                    cmdCADMOrder.Enabled = True
                Else
                    txtCADRefNo.Visible = False
                    cmdCADMOrder.Visible = False
                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT 'Edit Record
                strSql = "Select Future_so From Cust_ord_hdr Where  unit_code='" & gstrUNITID & "' and Account_code = '" & Trim(txtCustomerCode.Text) & "'"
                strSql = strSql & " and cust_ref = '" & Trim(txtReferenceNo.Text) & "' and amendment_no = '"
                strSql = strSql & Trim(txtAmendmentNo.Text) & "'"
                rsSalesParameter.GetResult(strSql)
                If rsSalesParameter.GetValue("Future_so") = True Then
                    If MsgBox("This is Future SO [Authorised], Changes in this SO with Update the Deatls of This SO and Make it UnAuthorised. Would you like to Proceed...", MsgBoxStyle.YesNo, ResolveResString(100)) = MsgBoxResult.No Then
                        cmdButtons.Revert()
                        Exit Sub
                    End If
                End If
                ssPOEntry.Enabled = True
                If gblnGSTUnit = False Then
                    txtSTax.Enabled = True : txtSTax.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : cmdHelp(4).Enabled = True
                    txtAddVAT.Enabled = True
                    txtAddVAT.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    cmdAddVAT.Enabled = True
                End If
                cmdForms.Enabled = True
                strSql = "Select PO_Type from cust_ord_hdr where unit_code='" & gstrUNITID & "' and cust_ref ='" & txtReferenceNo.Text & "' and Account_Code='" & txtCustomerCode.Text & "' and Active_Flag ='A' and PO_Type='V'"
                If IsNothing(strSql) = False And gblnGSTUnit = False Then
                    cmdServiceTax.Enabled = True
                    txtService.Enabled = True
                    txtService.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    cmdSBCtax.Enabled = True
                    txtSBC.Enabled = True
                    txtSBC.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    cmdKKCtax.Enabled = True
                    txtKKC.Enabled = True
                    txtKKC.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                End If
                SetFormButtonStyle()
                '1.Surcharge on S.Tax
                If gblnGSTUnit = False Then
                    txtSChSTax.Enabled = True : txtSChSTax.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : cmdHelp(6).Enabled = True
                End If

                '****
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
                CmbDiscounttype.Enabled = False
                CmbDiscounttype.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                txtdiscountvalue.Enabled = False
                txtdiscountvalue.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                txtmarkupvalue.Enabled = False
                txtmarkupvalue.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)

                txtAmendReason.Enabled = True
                cmdchangetype.Enabled = True
                If Len(Trim(txtAmendmentNo.Text)) = 0 Then
                    DTAmendmentDate.Enabled = False
                    txtAmendReason.Enabled = False
                    txtAmendReason.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    cmbPOType.Enabled = True
                    cmbPOType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    CmbDiscounttype.Enabled = True
                    CmbDiscounttype.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    txtdiscountvalue.Enabled = True
                    txtdiscountvalue.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    txtmarkupvalue.Enabled = True
                    txtmarkupvalue.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    txtCurrencyType.Enabled = True
                    txtCurrencyType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    cmdHelp(2).Enabled = True
                    ctlPerValue.Enabled = True
                    ctlPerValue.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
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
                            rsSalesParameter.GetResult("select ItemRateLink from Sales_Parameter where unit_code='" & gstrUNITID & "'")
                            If rsSalesParameter.GetValue("ItemRateLink") = True Then
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
                    CmbDiscounttype.Enabled = False
                    CmbDiscounttype.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    txtdiscountvalue.Enabled = False
                    txtdiscountvalue.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    txtmarkupvalue.Enabled = False
                    txtmarkupvalue.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    txtCurrencyType.Enabled = False
                    txtCurrencyType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    cmdHelp(2).Enabled = False
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
                With ssPOEntry
                    .Row = 1
                    .Col = 2
                    .Action = 0
                End With
                LockGrid()

                '10736222-changes done for CT2 ARE 3 functionality
                ChkCT2Reqd.Enabled = False
                chkShipAddress.Enabled = True
                Exit Sub
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE 'Delete Record
                enmValue = ConfirmWindow(10054, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION)
                If enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                    ' deleting the record from cust_ord_hdr table
                    strSql = "delete cust_ord_hdr where unit_code='" & gstrUNITID & "' and Account_Code='"
                    strSql = strSql & txtCustomerCode.Text & "' and Cust_Ref='"
                    strSql = strSql & txtReferenceNo.Text & "' and Amendment_No='"
                    strSql = strSql & txtAmendmentNo.Text & "'"
                    ' deleting the record from cust_ord_dtl table
                    Call DeleteRow()
                    mP_Connection.Execute("DELETE FROM Forms_dtl WHERE unit_code='" & gstrUNITID & "' and DOC_TYPE=9998 AND PO_NO='" & Trim(txtReferenceNo.Text) & "' AND AMENDMENT_NO='" & Trim(txtAmendmentNo.Text) & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    mP_Connection.Execute(strSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
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
                Dim IsENABLE_EXTERNAL_SALESNO As Boolean = False
                If DataExist("SELECT TOP 1 ENABLE_EXTERNAL_SALESNO FROM CUSTOMER_MST(NOLOCK) WHERE  UNIT_CODE='" & gstrUNITID & "' AND  CUSTOMER_CODE='" & Trim(txtCustomerCode.Text) & "' AND ENABLE_EXTERNAL_SALESNO =1 AND ((ISNULL(DEACTIVE_FLAG,0) <> 1) OR (CONVERT(VARCHAR(12),GETDATE(),106)<= CONVERT(VARCHAR(12),DEACTIVE_DATE,106)))") = True Then
                    IsENABLE_EXTERNAL_SALESNO = True
                End If
                With ssPOEntry

                    For counter = 1 To ssPOEntry.MaxRows
                        vardiscounttype = Nothing
                        Call ssPOEntry.GetText(22, counter, vardiscounttype)
                        vardiscountvalue = Nothing
                        Call ssPOEntry.GetText(23, counter, vardiscountvalue)
                        If (vardiscounttype = "[P]ercentage" Or vardiscounttype = "[V]alue") Then
                            If Val(vardiscountvalue) <= 0 Or vardiscountvalue Is Nothing Then
                                MsgBox("Discount Value can't be zero or blank", MsgBoxStyle.OkOnly, ResolveResString(100))
                                .Row = counter
                                .Col = 23
                                ssPOEntry.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                ssPOEntry.Focus()
                                Exit Sub
                            End If
                        End If

                        ''ADDED BY SUMIT KUMAR TO MAKE THE FALG MANDATORY ON 09 JULY 2019

                        If IsENABLE_EXTERNAL_SALESNO = True Then
                            varexternalsalesorder = Nothing
                            Call ssPOEntry.GetText(33, counter, varexternalsalesorder)
                            If Len(Trim(varexternalsalesorder)) <= 0 Then
                                MsgBox("Please Enter External Sales order ")
                                ssPOEntry.Col = 33
                                ssPOEntry.Row = counter
                                ssPOEntry.Focus()
                                ssPOEntry.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                Exit Sub
                            End If

                        End If
                    Next
                End With

                counter = ssPOEntry.MaxRows

                If counter = 0 Then
                    MessageBox.Show("Item Details Not Entered.", ResolveResString(100), MessageBoxButtons.OK)
                    Exit Sub
                End If


                For intRow = 1 To counter 'Checking if all details have been entered correctly
                    If ssPOEntry.MaxRows = 1 Or intRow <> ssPOEntry.MaxRows Then
                        If Not ValidRowData(intRow, 0) Then
                            gblnCancelUnload = True : gblnFormAddEdit = True
                            Exit Sub
                        End If
                    Else
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
                                        rsSalesParameter.GetResult("Select ToolCostMsg from Sales_Parameter where unit_code='" & gstrUNITID & "' ")
                                        If rsSalesParameter.GetValue("ToolCostMsg") = True Then
                                            If (MsgBox("You have Entered 0 Tool Cost,Save Data ?", MsgBoxStyle.YesNo, ResolveResString(100))) = MsgBoxResult.No Then
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
                        arrMain = Split(mstrFormDetails, "^")
                        strFormSQL = ""
                        For intOuterCount = 0 To UBound(arrMain) - 1
                            arrDet = Split(arrMain(intOuterCount), "|")
                            strFormSQL = strFormSQL & "INSERT INTO Forms_dtl(UNIT_CODE,DOC_TYPE, PO_NO, AMENDMENT_NO, SERIAL_NO, FORM_TYPE, FORM_NO, Account_code)"
                            strFormSQL = strFormSQL & " VALUES('" & gstrUNITID & "',9998,'" & txtReferenceNo.Text & "','" & Trim(txtAmendmentNo.Text) & "','" & intOuterCount & "','" & arrDet(0) & "','" & arrDet(1) & "', '" & Trim(txtCustomerCode.Text) & "')" & vbCrLf
                        Next
                        With mP_Connection
                            .BeginTrans()
                            Call InsertRowCustOrdHdr()
                            Call InsertRow()
                            If Len(strFormSQL) > 0 Then .Execute(strFormSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            .CommitTrans()
                            '10844039--Starts
                            'rsSalesParameter.GetResult("Select AppendSOItem from sales_parameter where unit_code='" & gstrUNITID & "'")
                            'If rsSalesParameter.GetValue("AppendSOItem") = True Then
                            '    'ISSUE ID : 10763705
                            '    If mblnappendsoitem_customer = False Then
                            '        'ISSUE ID : 10763705
                            '        Call InsertPreviousSODetails(Trim(txtCustomerCode.Text), Trim(txtReferenceNo.Text), Trim(txtAmendmentNo.Text), Trim(lblIntSONoDes.Text), CShort(Trim(lblRevisionNo.Text)))
                            '    End If
                            'End If
                            '10844039--Ends
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
                        If mblnDiscountFunctionality = True AndAlso ValidDiscountvalue() = True Then
                            MsgBox(" % cannot be greater than 100", MsgBoxStyle.OkOnly, ResolveResString(100))
                            Exit Sub
                        End If
                        If mblnMarkupFunctionality = True Then
                            If ValidMarkupvalue() = True Then
                                MsgBox(" % cannot be greater than 100", MsgBoxStyle.OkOnly, ResolveResString(100))
                                Exit Sub
                            End If
                        End If

                        strFormSQL = "DELETE FROM Forms_dtl WHERE unit_code='" & gstrUNITID & "' and DOC_TYPE=9998 AND Account_code='" & Trim(txtCustomerCode.Text) & "' and PO_NO='" & Trim(txtReferenceNo.Text) & "' AND AMENDMENT_NO='" & Trim(txtAmendmentNo.Text) & "'" & vbCrLf
                        If Len(mstrFormDetails) > 0 Then
                            arrMain = Split(mstrFormDetails, "^")
                            For intOuterCount = 0 To UBound(arrMain) - 1
                                arrDet = Split(arrMain(intOuterCount), "|")
                                strFormSQL = strFormSQL & "INSERT INTO Forms_dtl(UNIT_CODE,DOC_TYPE, PO_NO, AMENDMENT_NO, SERIAL_NO, FORM_TYPE, FORM_NO, Account_code)"
                                strFormSQL = strFormSQL & " VALUES('" & gstrUNITID & "',9998,'" & txtReferenceNo.Text & "','" & txtAmendmentNo.Text & "','" & intOuterCount & "','" & arrDet(0) & "','" & arrDet(1) & "', '" & Trim(txtCustomerCode.Text) & "')" & vbCrLf
                            Next
                        End If
                        With mP_Connection
                            'To Confirm the deletion of marked rows
                            If ConfirmDeletion() = False Then Exit Sub
                            .BeginTrans()
                            Call UpdateRow()
                            'delete the record in the cust_ord_dtl table
                            Call DeleteRow()
                            ' Add all the records in the grid to the table cust_ord_dtl which are not marked for deletion
                            Call InsertRow()
                            If Len(strFormSQL) > 0 Then .Execute(strFormSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                            .CommitTrans()
                            '.RollbackTrans()
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
                End Select
                If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                    Call ConfirmWindow(10049, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                Else
                    If Len(Trim(txtAmendmentNo.Text)) = 0 Then
                        MsgBox("SO Successfully update with Internal SO No " & lblIntSONoDes.Text, MsgBoxStyle.Information, ResolveResString(100))
                    Else
                        Call ConfirmWindow(10049, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                    End If
                End If
                gblnCancelUnload = False : gblnFormAddEdit = False
                Call EnableControls(False, Me, True)
                txtCustomerCode.Enabled = True
                txtCustomerCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                cmdHelp(0).Enabled = True
                frmSearch.Enabled = True
                optPartNo.Enabled = True
                optPartNo.Checked = True
                optItem.Enabled = True
                txtsearch.Enabled = True
                txtsearch.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                mstrFormDetails = ""
                txtCustomerCode.Focus()
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL
                If ConfirmWindow(10053, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION, 60095) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                    Call EnableControls(False, Me, True)
                    frmSearch.Enabled = True
                    optPartNo.Enabled = True
                    optPartNo.Checked = True
                    optItem.Enabled = True
                    txtsearch.Enabled = True
                    txtsearch.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
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

        If KeyAscii = 39 Or KeyAscii = 34 Or KeyAscii = 96 Then KeyAscii = 0
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
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
        If KeyAscii = 39 Or KeyAscii = 34 Or KeyAscii = 96 Then KeyAscii = 0
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
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
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub DTAmendmentDate_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles DTAmendmentDate.KeyPress
        On Error GoTo ErrHandler
        Dim KeyAscii As Short = Asc(e.KeyChar)
        If KeyAscii = 39 Or KeyAscii = 34 Or KeyAscii = 96 Then KeyAscii = 0
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub cmdButtons_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdButtons.Load
    End Sub
    Private Sub ssPOEntry_Advance(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_AdvanceEvent) Handles ssPOEntry.Advance
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
        If mblnIsExecuting = True Then
            Exit Sub
        End If

        Dim rsDesc As New ClsResultSetDB
        Dim rsSalesParametere As New ClsResultSetDB
        Dim varHelpItem As Object
        Dim strSOEntry() As String
        Dim strtest As String
        Dim strReturnCustRef As String
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        intMaxLoop = ssPOEntry.MaxRows
        Dim strCT2Condition As String = ""
        Dim StrServiceCond As String = ""
        'GST 
        Dim rsitem As ClsResultSetDB
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
            End With
        End If
        '10940008
        If Mid(cmbPOType.Text, 1, 1) = "V" And mblnServiceInvoicemate = True Then
            StrServiceCond = " AND ITEM_MAIN_GRP ='M' "
        End If
        rsSalesParametere.GetResult("Select ItemRateLink from Sales_parameter (Nolock) where unit_code='" & gstrUNITID & "'")
        If Len(Trim(txtCurrencyType.Text)) = 0 Then Exit Sub 'for currency type check
        '10869290
        If e.col = 3 Then
            If txtAmendmentNo.Enabled = False Then
                If rsSalesParametere.GetValue("ItemRateLink") = True Then
                    strtest = "Select a.Cust_drgNo,a.ITem_code,a.drg_Desc from CustItem_Mst a,Itemrate_Mst b where a.unit_code=b.unit_code and a.unit_code='" & gstrUNITID & "' and a.Account_Code='" & txtCustomerCode.Text & "' AND A.ACTIVE = 1 and a.Account_code = b.Party_Code and a.cust_DrgNo = b.Item_code and b.serial_no=(select max(serial_no) from itemrate_mst i1 Where i1.party_code = b.party_code and i1.item_code=b.item_code and i1.unit_code=b.unit_code and i1.unit_code='" & gstrUNITID & "') and datediff(mm,convert(datetime,'" & Format(DTDate.Value, "dd mmm yyyy") & "'),convert(datetime,b.DateFrom))<=0 and CustVend_Flg = 'C' "
                    strSOEntry = ctlEMPHelpSOEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select a.Cust_drgNo,a.Item_code,a.drg_Desc from CustItem_Mst a,Itemrate_Mst b, Item_MST as C where A.ACTIVE = 1 and A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE=C.UNIT_CODE  AND A.Item_code=C.Item_Code and C.Status='A' and C.Hold_Flag=0 AND A.UNIT_CODE='" & gstrUNITID & "' AND a.Account_Code='" & txtCustomerCode.Text & "' and a.Account_code = b.Party_Code and a.cust_DrgNo = b.Item_code and b.serial_no=(select max(serial_no) from itemrate_mst i1 Where i1.party_code = b.party_code and i1.item_code=b.item_code AND i1.unit_code=b.unit_code and i1.unit_code='" & gstrUNITID & "') and datediff(mm,'" & getDateForDB(DTDate.Value) & "',b.DateFrom)<=0 and CustVend_Flg = 'C' " & IIf(StrServiceCond.Trim.Length = 0, "", StrServiceCond) & IIf(strCT2Condition.Trim.Length = 0, "", strCT2Condition))
                Else
                    If Mid(cmbPOType.Text, 1, 1) = "A" Then ' FOR SUB ASSEMBLY 
                        strSOEntry = ctlEMPHelpSOEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select A.Cust_drgNo, A.ITem_code, A.drg_Desc from CustItem_Mst as A, Item_MST as B where A.ACTIVE = 1 and A.unit_code=B.unit_Code and A.Item_code=B.Item_Code and Status='A' and Hold_Flag=0 and Account_Code='" & txtCustomerCode.Text & "' and a.unit_code='" & gstrUNITID & "'AND ITEM_MAIN_GRP ='S' " & IIf(strCT2Condition.Trim.Length = 0, "", strCT2Condition))
                    Else
                        strSOEntry = ctlEMPHelpSOEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select A.Cust_drgNo, A.ITem_code, A.drg_Desc from CustItem_Mst as A, Item_MST as B where A.ACTIVE = 1 and A.unit_code=B.unit_Code and A.Item_code=B.Item_Code and Status='A' and Hold_Flag=0 and Account_Code='" & txtCustomerCode.Text & "' and a.unit_code='" & gstrUNITID & "' " & IIf(StrServiceCond.Trim.Length = 0, "", StrServiceCond) & IIf(strCT2Condition.Trim.Length = 0, "", strCT2Condition))
                    End If
                End If
            Else
                If Mid(cmbPOType.Text, 1, 1) = "A" Then ' FOR SUB ASSEMBLY 
                    strSOEntry = ctlEMPHelpSOEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select A.Cust_drgNo, A.ITem_code, A.drg_Desc from CustItem_Mst as A, Item_MST as B where A.ACTIVE = 1 and A.unit_code=B.unit_Code and A.Item_code=B.Item_Code and Status='A' and Hold_Flag=0 and Account_Code='" & txtCustomerCode.Text & "' and a.unit_code='" & gstrUNITID & "' AND ITEM_MAIN_GRP ='S' " & IIf(strCT2Condition.Trim.Length = 0, "", strCT2Condition))
                Else
                    strSOEntry = ctlEMPHelpSOEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select A.Cust_drgNo, A.ITem_code, A.drg_Desc from CustItem_Mst as A, Item_MST as B where A.ACTIVE = 1 and A.unit_code=B.unit_Code and A.Item_code=B.Item_Code and Status='A' and Hold_Flag=0 and Account_Code='" & txtCustomerCode.Text & "' and a.unit_code='" & gstrUNITID & "' " & IIf(StrServiceCond.Trim.Length = 0, "", StrServiceCond) & IIf(strCT2Condition.Trim.Length = 0, "", strCT2Condition))
                End If
            End If

            If UBound(strSOEntry) <= 0 Then Exit Sub
            If strSOEntry(0) = "0" Then
                Call ConfirmWindow(10013, ConfirmWindowButtonsEnum.BUTTON_OK, ConfirmWindowImagesEnum.IMG_INFO)
            Else
                strReturnCustRef = CheckForMultipleOpenSO(txtCustomerCode.Text, txtReferenceNo.Text, strSOEntry(0), strSOEntry(1))
                If Len(strReturnCustRef) > 0 Then
                    MsgBox("More than one Sales Order cannot be active for Customer Item Combination   " & strReturnCustRef, vbInformation, ResolveResString(100))
                    Exit Sub
                End If
                Call ssPOEntry.SetText(2, ssPOEntry.ActiveRow, strSOEntry(0))
                Call ssPOEntry.SetText(4, ssPOEntry.ActiveRow, strSOEntry(1))
                lblCustPartDesc.Text = strSOEntry(2)
                'GST CHANGE
                'GST CHANGE
                If gblnGSTUnit = True Then
                    ssPOEntry.Row = 1 : ssPOEntry.Row2 = ssPOEntry.MaxRows : ssPOEntry.Col = 15 : ssPOEntry.Col2 = 32 : ssPOEntry.BlockMode = True : ssPOEntry.Lock = True : ssPOEntry.BlockMode = False
                End If

                STRSQL = "set dateformat 'dmy' SELECT * FROM DBO.UFN_GST_ITEMWISETAXES('" & gstrUNITID & "','" & txtCustomerCode.Text.Trim & "','" & strSOEntry(1) & "','" & DTEffectiveDate.Value & "','" & DTValidDate.Value & "')"
                objSQLConn = SqlConnectionclass.GetConnection()
                objCommand = New SqlCommand(STRSQL, objSQLConn)
                objReader = objCommand.ExecuteReader()
                If objReader.HasRows = True Then
                    objReader.Read()
                    Call ssPOEntry.SetText(26, ssPOEntry.ActiveRow, objReader.GetValue(1))
                    Call ssPOEntry.SetText(27, ssPOEntry.ActiveRow, objReader.GetValue(2))
                    Call ssPOEntry.SetText(28, ssPOEntry.ActiveRow, objReader.GetValue(3))
                    Call ssPOEntry.SetText(29, ssPOEntry.ActiveRow, objReader.GetValue(4))
                    Call ssPOEntry.SetText(30, ssPOEntry.ActiveRow, objReader.GetValue(5))
                    Call ssPOEntry.SetText(31, ssPOEntry.ActiveRow, objReader.GetValue(6))
                    Call ssPOEntry.SetText(32, ssPOEntry.ActiveRow, objReader.GetValue(0))
                    'ssPOEntry.Row = ssPOEntry.MaxRows : ssPOEntry.Row2 = ssPOEntry.MaxRows : ssPOEntry.Col = 8 : ssPOEntry.Col2 = 30 : ssPOEntry.BlockMode = True : ssPOEntry.Lock = True : ssPOEntry.BlockMode = False
                    'ssPOEntry.Row = ssPOEntry.MaxRows : ssPOEntry.Row2 = ssPOEntry.MaxRows : ssPOEntry.Col = 10 : ssPOEntry.Col2 = 10 : ssPOEntry.BlockMode = True : ssPOEntry.Lock = True : ssPOEntry.BlockMode = False
                    'GST CHANGE
                End If
                objReader = Nothing
                objSQLConn.Close()
                objSQLConn = Nothing
            End If
            ''ADDED BY SUMIT KUMAR ON 09 JULY 2019
            If DataExist("SELECT TOP 1 ENABLE_EXTERNAL_SALESNO FROM CUSTOMER_MST(NOLOCK) WHERE  UNIT_CODE='" & gstrUNITID & "' AND  CUSTOMER_CODE='" & Trim(txtCustomerCode.Text) & "' AND ENABLE_EXTERNAL_SALESNO =1 AND ((ISNULL(DEACTIVE_FLAG,0) <> 1) OR (CONVERT(VARCHAR(12),GETDATE(),106)<= CONVERT(VARCHAR(12),DEACTIVE_DATE,106)))") = True Then
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
                    Call ssPOEntry.SetText(33, ssPOEntry.ActiveRow, IIf(rsitem.GetValue("EXTERNAL_SALESORDER_NO") = "", "", rsitem.GetValue("EXTERNAL_SALESORDER_NO")))
                End If
                rsitem.ResultSetClose()
                rsitem = Nothing

            End If
            ''ENED SUMIT 
            'GST CHANGE
            If cmdButtons.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then

                If Len(Trim(m_Item_Code)) > 0 Then
                    'ISSUE ID : 10515727
                    'm_strSql = " Select * from Cust_ord_dtl where unit_code='" & gstrUNITID & "' and Item_code ='" & m_Item_Code & "' and cust_drgNo ='" & varHelpItem & "' and cust_ref ='" & txtReferenceNo.Text & "' and Active_flag ='A' and Authorized_Flag =1"
                    m_strSql = " Select * from Cust_ord_dtl where unit_code='" & gstrUNITID & "' and account_code='" & txtCustomerCode.Text & "' and Item_code ='" & m_Item_Code & "' and cust_drgNo ='" & varHelpItem & "' and cust_ref ='" & txtReferenceNo.Text & "' and Active_flag ='A' and Authorized_Flag =1"
                    rsitem = New ClsResultSetDB
                    rsitem.GetResult(m_strSql)
                    If rsitem.GetNoRows > 0 Then
                        If rsitem.GetValue("OpenSO") = False Then
                            ssPOEntry.Row = ssPOEntry.MaxRows : ssPOEntry.Col = 1 : ssPOEntry.Col2 = 1 : ssPOEntry.CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox : ssPOEntry.Value = CheckState.Unchecked
                            ssPOEntry.Row = ssPOEntry.MaxRows : ssPOEntry.Row2 = ssPOEntry.MaxRows : ssPOEntry.Col = 5 : ssPOEntry.Col2 = 5 : ssPOEntry.BlockMode = True : ssPOEntry.Lock = False : ssPOEntry.BlockMode = False
                        Else
                            ssPOEntry.Row = ssPOEntry.MaxRows : ssPOEntry.Col = 1 : ssPOEntry.Col2 = 1 : ssPOEntry.Value = CheckState.Unchecked
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
                        Call ssPOEntry.SetText(9, ssPOEntry.MaxRows, rsitem.GetValue("Packing_Type"))
                        Call ssPOEntry.SetText(10, ssPOEntry.MaxRows, rsitem.GetValue("Excise_Duty"))
                        If UCase(Trim(cmbPOType.Text)) <> "JOB WORK" And (gblnGSTUnit = False Or (gblnGSTUnit And _blnEOUFlag)) Then
                            addExciseDuty(e.row)
                        End If
                        Call ssPOEntry.SetText(16, ssPOEntry.MaxRows, rsitem.GetValue("Others"))
                        Call ssPOEntry.SetText(11, ssPOEntry.MaxRows, (rsitem.GetValue("Others") * ctlPerValue.Text))
                        Call ssPOEntry.SetText(17, ssPOEntry.MaxRows, rsitem.GetValue("Despatch_Qty"))
                    Else
                        If txtAmendmentNo.Enabled = False Then
                            If Len(Trim(m_Item_Code)) > 0 Then
                                If rsSalesParametere.GetValue("ItemRateLink") = True Then
                                    m_strSql = " select * from ITemRate_Mst where unit_code='" & gstrUNITID & "' and Serial_No = (select max(serial_no) from itemrate_mst where Party_code = '" & txtCustomerCode.Text & "' and item_code = '" & varHelpItem & "' and datediff(mm,convert(varchar(10),'" & getDateForDB(DTDate.Value) & "',103),convert(varchar(10),DateFrom,103))<=0 and custVend_Flg ='C' and unit_code='" & gstrUNITID & "' )"
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
                                        Call ssPOEntry.SetText(14, ssPOEntry.MaxRows, rsitem.GetValue("Cust_supplied_Material"))
                                        Call ssPOEntry.SetText(7, ssPOEntry.MaxRows, (rsitem.GetValue("Cust_supplied_Material") * ctlPerValue.Text))
                                        Call ssPOEntry.SetText(15, ssPOEntry.MaxRows, rsitem.GetValue("Tool_Cost"))
                                        Call ssPOEntry.SetText(8, ssPOEntry.MaxRows, (rsitem.GetValue("Tool_Cost") * ctlPerValue.Text))
                                        If rsitem.GetValue("Packaging_Flag") = False Then
                                            Call ssPOEntry.SetText(9, ssPOEntry.MaxRows, ((rsitem.GetValue("Packaging_Amount") * 100) / rsitem.GetValue("Rate")))
                                        Else
                                            Call ssPOEntry.SetText(9, ssPOEntry.MaxRows, rsitem.GetValue("Packaging_Amount"))
                                        End If
                                        If UCase(Trim(cmbPOType.Text)) <> "JOB WORK" And (gblnGSTUnit = False Or (gblnGSTUnit And _blnEOUFlag)) Then
                                            If addExciseDuty(e.row) = False Then
                                                ssPOEntry.MaxRows = ssPOEntry.MaxRows - 1
                                                If ssPOEntry.MaxRows < 1 Then
                                                    Call ADDRow()
                                                End If
                                                ssPOEntry.Row = ssPOEntry.MaxRows : ssPOEntry.Col = 2 : ssPOEntry.Focus()
                                            End If
                                        End If
                                        Call ssPOEntry.SetText(16, ssPOEntry.MaxRows, rsitem.GetValue("Others"))
                                        Call ssPOEntry.SetText(11, ssPOEntry.MaxRows, (rsitem.GetValue("Others") * ctlPerValue.Text))
                                        If rsitem.GetValue("Edit_flg") = False Then
                                            ssPOEntry.Col = 6 : ssPOEntry.Col2 = ssPOEntry.MaxCols : ssPOEntry.Row = ssPOEntry.MaxRows : ssPOEntry.Row2 = ssPOEntry.MaxRows : ssPOEntry.BlockMode = True : ssPOEntry.Lock = True : ssPOEntry.BlockMode = False
                                        Else
                                            ssPOEntry.Col = 6 : ssPOEntry.Col2 = ssPOEntry.MaxCols : ssPOEntry.Row = ssPOEntry.MaxRows : ssPOEntry.Row2 = ssPOEntry.MaxRows : ssPOEntry.BlockMode = True : ssPOEntry.Lock = False : ssPOEntry.BlockMode = False
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
                    Call .GetText(2, e.row, varDrgNo)
                    Call .GetText(4, e.row, varItemCode)
                    If (Len(Trim(varDrgNo)) > 0) And Len(Trim(varItemCode)) > 0 Then
                        rsSalesParameter.GetResult("Select ItemRateLink from Sales_Parameter where unit_code='" & gstrUNITID & "' ")
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
        If cmbPOType.Text.Trim = "Q-TRADING" Then
            With ssPOEntry
                .BlockMode = True
                .Row = e.row
                .Row2 = .MaxRows
                .Col = 10
                .Col2 = 10
                .Lock = True
                .BlockMode = False
            End With
        End If
        Exit Sub    'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub ssPOEntry_KeyDownEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyDownEvent) Handles ssPOEntry.KeyDownEvent
        On Error GoTo ErrHandler
        Dim varHelpItem As Object
        Dim rsDesc As New ClsResultSetDB
        Dim rsSalesParameter As New ClsResultSetDB
        Dim rsAbatmentRate As New ClsResultSetDB
        Dim inti As Integer
        Dim strReturnCustRef As String
        Dim varMRP, varAbatment, varAccessibleRateforMRP As Object
        Dim strSOEntry() As String
        Dim strCT2Condition As String = ""
        Dim StrServiceCond As String = ""
        '10797956 
        '10736222-changes done for CT2 ARE 3 functionality - to enable CT2 Reqd flag
        'AMIT RANA 28 06 2017 
        If gblnGSTUnit = True And ssPOEntry.ActiveCol = 2 And e.keyCode = 112 Then 'f1 help
            ssPOEntry_ButtonClicked(sender, New AxFPSpreadADO._DSpreadEvents_ButtonClickedEvent(3, ssPOEntry.ActiveRow, 1))
            Exit Sub
        End If
        'AMIT RANA 28 06 2017


        If ChkCT2Reqd.Checked = True Then
            strCT2Condition = " And Exists(select * From CT2_Cust_Item_Linkage Z where A.UNIT_CODE=z.Unit_Code and A.Account_Code=z.Customer_Code and A.Item_code =z.Item_Code and A.Cust_Drgno=z.Cust_drgno And z.Active=1 and z.isAuthorized=1)"
            With ssPOEntry
                .Col = 10
                .Text = "EX0"
                .Lock = True
            End With
        End If
        '10869290
        '10940008
        If Mid(cmbPOType.Text, 1, 1) = "V" And mblnServiceInvoicemate = True Then
            StrServiceCond = " AND ITEM_MAIN_GRP ='M' "
        End If

        'If user has pressed Ctrl + N, then add a new row
        rsSalesParameter.GetResult("Select ItemRateLink from Sales_Parameter where unit_code='" & gstrUNITID & "'")

        'Added By ekta uniyal on 15 Apr 2014 
        'Issue Id - 10528368
        If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then Exit Sub
        'End Here

        If (e.shift = 2 And e.keyCode = Keys.N) Then
            'Add a blank row
            If ValidRowData(ssPOEntry.ActiveRow, 0) Then Call ADDRow()
            'Setting the focus
            ssPOEntry.Row = ssPOEntry.MaxRows
            If mblnpackingdefined = False Then
                ssPOEntry.Col = 2 : ssPOEntry.Action = FPSpreadADO.ActionConstants.ActionActiveCell
            Else
                ssPOEntry.Col = 9 : ssPOEntry.Action = FPSpreadADO.ActionConstants.ActionActiveCell
            End If

        End If
        If ssPOEntry.ActiveCol = 2 Or ssPOEntry.ActiveCol = 4 Then
            If e.keyCode = 112 Then
                If Len(Trim(txtCurrencyType.Text)) = 0 Then Exit Sub 'for Currency Type entered Validation
                If txtAmendmentNo.Enabled = False Then
                    If rsSalesParameter.GetValue("ItemRateLink") = True Then
                        strSOEntry = ctlEMPHelpSOEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select a.Cust_drgNo,a.ITem_code,a.drg_Desc from CustItem_Mst a, Itemrate_Mst b, Item_MST as C where A.ACTIVE = 1 and A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE=C.UNIT_CODE and A.Item_code=C.Item_code and Status='A' and Hold_Flag=0 and a.unit_code='" & gstrUNITID & "' and a.Account_Code='" & txtCustomerCode.Text & "' and a.Account_code = b.Party_Code and a.cust_DrgNo = b.Item_code and b.serial_no=(select max(serial_no) from itemrate_mst i1 Where i1.party_code = b.party_code and i1.item_code=b.item_code and i1.unit_code=b.unit_code and i1.unit_code='" & gstrUNITID & "') and datediff(mm,'" & getDateForDB(DTDate.Value) & "',b.DateFrom)<=0 and CustVend_Flg = 'C'" & IIf(StrServiceCond.Trim.Length = 0, "", StrServiceCond) & IIf(strCT2Condition.Trim.Length = 0, "", strCT2Condition))
                    Else
                        strSOEntry = ctlEMPHelpSOEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select A.Cust_drgNo, A.ITem_code, A.drg_Desc from CustItem_Mst as A, Item_MST as B where A.ACTIVE = 1 and a.unit_code=b.unit_code and A.Item_code=B.Item_code and B.Status='A' and B.hold_flag=0 and a.unit_code='" & gstrUNITID & "' and Account_Code='" & txtCustomerCode.Text & "'" & IIf(StrServiceCond.Trim.Length = 0, "", StrServiceCond) & IIf(strCT2Condition.Trim.Length = 0, "", strCT2Condition))
                    End If
                Else
                    If Mid(cmbPOType.Text, 1, 1) = "V" Then
                        strSOEntry = ctlEMPHelpSOEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select Cust_drgNo,a.ITem_code,drg_Desc from CustItem_Mst a, item_mst b where a.ACTIVE = 1 and a.unit_Code = b.Unit_Code and a.item_code = b.item_code and b.status = 'A' and b.Hold_flag = 0 and a.unit_code='" & gstrUNITID & "' and a.Account_Code='" & txtCustomerCode.Text & "'" & IIf(StrServiceCond.Trim.Length = 0, "", StrServiceCond) & IIf(strCT2Condition.Trim.Length = 0, "", strCT2Condition))
                    Else
                        strSOEntry = ctlEMPHelpSOEntry.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "Select Cust_drgNo,ITem_code,drg_Desc from CustItem_Mst where ACTIVE = 1 and unit_code='" & gstrUNITID & "' and Account_Code='" & txtCustomerCode.Text & "'" & IIf(strCT2Condition.Trim.Length = 0, "", strCT2Condition))
                    End If

                End If
                If UBound(strSOEntry) <= 0 Then Exit Sub
                If strSOEntry(0) = "0" Then
                    Call ConfirmWindow(10013, ConfirmWindowButtonsEnum.BUTTON_OK, ConfirmWindowImagesEnum.IMG_INFO)
                Else
                    strReturnCustRef = CheckForMultipleOpenSO(txtCustomerCode.Text, txtReferenceNo.Text, strSOEntry(0), strSOEntry(1))
                    If Len(strReturnCustRef) > 0 Then
                        MsgBox("More than one Sales Order cannot be active for Customer Item Combination    " & strReturnCustRef, vbInformation, ResolveResString(100))
                        Exit Sub
                    End If
                    Call ssPOEntry.SetText(2, ssPOEntry.ActiveRow, strSOEntry(0))
                    Call ssPOEntry.SetText(4, ssPOEntry.ActiveRow, strSOEntry(1))
                    lblCustPartDesc.Text = strSOEntry(2)
                End If
                If cmdButtons.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                    Dim rsitem As ClsResultSetDB
                    If Len(Trim(m_Item_Code)) > 0 Then
                        m_strSql = "SElect ACCOUNT_CODE,CUST_REF,AMENDMENT_NO,ITEM_CODE,RATE,ORDER_QTY,DESPATCH_QTY,ACTIVE_FLAG,CUST_MTRL,CUST_DRGNO,PACKING,OTHERS,EXCISE_DUTY,CUST_DRG_DESC,TOOL_COST,AUTHORIZED_FLAG,OPENSO,SALESTAX_TYPE,PERVALUE,INTERNALSONO,REVISIONNO,REMARKS,MRP,ABANTMENT_CODE,ACCESSIBLERATEFORMRP,PACKING_TYPE,TOOL_AMOR_FLAG,SHOWINAUTH,ADD_EXCISE_DUTY from Cust_ord_dtl where unit_code='" & gstrUNITID & "' and Item_code ='" & m_Item_Code & "' and cust_drgNo ='" & varHelpItem & "' and cust_ref ='" & txtReferenceNo.Text & "' and Active_flag ='A' and Authorized_Flag =1"
                        rsitem = New ClsResultSetDB
                        rsitem.GetResult(m_strSql)
                        If rsitem.GetNoRows > 0 Then
                            Call ssPOEntry.SetText(2, ssPOEntry.MaxRows, rsitem.GetValue("Cust_DrgNo"))
                            Call ssPOEntry.SetText(4, ssPOEntry.MaxRows, rsitem.GetValue("Item_Code "))
                            Call ssPOEntry.SetText(5, ssPOEntry.MaxRows, rsitem.GetValue("Order_Qty"))
                            Call ssPOEntry.SetText(13, ssPOEntry.MaxRows, rsitem.GetValue("Rate"))
                            Call ssPOEntry.SetText(6, ssPOEntry.MaxRows, (rsitem.GetValue("Rate") * ctlPerValue.Text))
                            Call ssPOEntry.SetText(14, ssPOEntry.MaxRows, rsitem.GetValue("Cust_Mtrl"))
                            Call ssPOEntry.SetText(7, ssPOEntry.MaxRows, (rsitem.GetValue("Cust_Mtrl") * ctlPerValue.Text))
                            Call ssPOEntry.SetText(15, ssPOEntry.MaxRows, rsitem.GetValue("Tool_Cost"))
                            Call ssPOEntry.SetText(8, ssPOEntry.MaxRows, (rsitem.GetValue("Tool_Cost") * ctlPerValue.Text))
                            Call ssPOEntry.SetText(9, ssPOEntry.MaxRows, rsitem.GetValue("Packing_Type"))
                            Call ssPOEntry.SetText(10, ssPOEntry.MaxRows, rsitem.GetValue("Excise_Duty"))
                            Call ssPOEntry.SetText(16, ssPOEntry.MaxRows, rsitem.GetValue("Others"))
                            Call ssPOEntry.SetText(11, ssPOEntry.MaxRows, (rsitem.GetValue("Others") * ctlPerValue.Text))
                            Call ssPOEntry.SetText(19, ssPOEntry.MaxRows, rsitem.GetValue("MRP"))
                            Call ssPOEntry.SetText(20, ssPOEntry.MaxRows, rsitem.GetValue("abantment_code"))
                            Call ssPOEntry.SetText(21, ssPOEntry.MaxRows, rsitem.GetValue("AccessibleRateforMRP"))
                        Else
                            If txtAmendmentNo.Enabled = False Then
                                If Len(Trim(m_Item_Code)) > 0 Then
                                    If rsSalesParameter.GetValue("ItemRateLink") = True Then
                                        m_strSql = " select DATEFROM,DATETO,PARTY_CODE,ITEM_CODE,CUSTVEND_FLG,RATE,SERIAL_NO,DISCOUNT_FLAG,DISCOUNT_AMOUNT,CUST_SUPPLIED_MATERIAL,TOOL_COST,PACKAGING_FLAG,PACKAGING_AMOUNT,OTHERS,CURRENCY_CODE,EDIT_FLG from ITemRate_Mst where unit_code='" & gstrUNITID & "' and Serial_No = (select max(serial_no) from itemrate_mst where unit_code='" & gstrUNITID & "' and Party_code = '" & txtCustomerCode.Text & "' and item_code = '" & varHelpItem & "' and datediff(mm,convert(varchar(10),'" & getDateForDB(DTDate.Value) & "',103),convert(varchar(10),DateFrom,103))<=0 and custVend_Flg ='C')"
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
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        ElseIf ssPOEntry.ActiveCol = 9 Then
            If e.keyCode = 112 Then
                With ssPOEntry
                    .Row = .ActiveRow : .Col = .ActiveCol
                    varHelpItem = ShowList(1, 6, Trim(.Text), "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND Tx_TaxeID='PKT' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))")
                    If varHelpItem = "-1" Then
                        MsgBox("Packing Code Does Not Exist", vbInformation, ResolveResString(100))
                        .Text = ""
                    Else
                        Call ssPOEntry.SetText(9, ssPOEntry.ActiveRow, Trim(varHelpItem))
                        .Focus()
                    End If
                End With
            End If
        ElseIf ssPOEntry.ActiveCol = 10 Then
            If e.keyCode = 112 Then
                If Len(Trim(txtCurrencyType.Text)) = 0 Then Exit Sub 'for Currency Type entered Validation
                If cmbPOType.Text.Trim = "Q-TRADING" Then Exit Sub 'for Currency Type entered Validation
                If gblnGSTUnit = False Or (gblnGSTUnit And _blnEOUFlag) Then
                    If ChkCT2Reqd.Checked = False Then
                        With ssPOEntry
                            .Row = .ActiveRow : .Col = .ActiveCol
                            varHelpItem = ShowList(1, 6, Trim(.Text), "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND Tx_TaxeID='EXC' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))")
                            If varHelpItem = "-1" Then
                                MsgBox("Excise Code Does Not Exist", vbInformation, ResolveResString(100))
                                .Text = ""
                            Else
                                Call ssPOEntry.SetText(10, ssPOEntry.ActiveRow, varHelpItem)
                            End If
                        End With
                    End If
                End If
            End If
        ElseIf ssPOEntry.ActiveCol = 20 Then
            If e.keyCode = 112 Then
                If Len(Trim(txtCurrencyType.Text)) = 0 Then Exit Sub 'for Currency Type entered Validation
                With ssPOEntry
                    .Row = .ActiveRow : .Col = .ActiveCol
                    varHelpItem = ShowList(1, 6, Trim(.Text), "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND Tx_TaxeID='ABNT' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))")
                    If varHelpItem = "-1" Then
                        MsgBox("Abatment Code Does Not Exist", vbInformation, ResolveResString(100))
                    Else
                        Call ssPOEntry.SetText(20, ssPOEntry.ActiveRow, varHelpItem)
                        Call ssPOEntry.GetText(19, ssPOEntry.ActiveRow, varMRP)
                        Call ssPOEntry.GetText(20, ssPOEntry.ActiveRow, varAbatment)
                        m_strSql = "select txrt_percentage from Gen_TaxRate where unit_code='" & gstrUNITID & "' and Tx_TaxeID='ABNT' and txrt_rate_no='" & varAbatment & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
                        rsAbatmentRate = New ClsResultSetDB
                        rsAbatmentRate.GetResult(m_strSql)
                        If rsAbatmentRate.GetNoRows > 0 Then
                            varAbatment = Val(rsAbatmentRate.GetValue("txrt_percentage"))
                            varAccessibleRateforMRP = varMRP - (Trim(varMRP) * Trim(varAbatment)) / 100
                            Call ssPOEntry.SetText(21, ssPOEntry.ActiveRow, varAccessibleRateforMRP)
                        End If
                    End If
                End With
                rsAbatmentRate.ResultSetClose()
                rsAbatmentRate = Nothing
            End If
            If e.keyCode = 13 Then
                If Len(Trim(txtCurrencyType.Text)) = 0 Then Exit Sub 'for Currency Type entered Validation
                With ssPOEntry
                    Call ssPOEntry.GetText(19, ssPOEntry.ActiveRow, varMRP)
                    Call ssPOEntry.GetText(20, ssPOEntry.ActiveRow, varAbatment)
                    m_strSql = "select txrt_percentage from Gen_TaxRate where unit_code='" & gstrUNITID & "' and Tx_TaxeID='ABNT' and txrt_rate_no='" & varAbatment & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
                    rsAbatmentRate = New ClsResultSetDB
                    rsAbatmentRate.GetResult(m_strSql)
                    If rsAbatmentRate.GetNoRows > 0 Then
                        varAbatment = Val(rsAbatmentRate.GetValue("txrt_percentage"))
                        varAccessibleRateforMRP = varMRP - (Trim(varMRP) * Trim(varAbatment)) / 100
                        Call ssPOEntry.SetText(21, ssPOEntry.ActiveRow, varAccessibleRateforMRP)
                    Else
                        If varMRP > 0 Then
                            MsgBox("Abatment Code Does Not Exist", vbInformation, ResolveResString(100))
                            Call ssPOEntry.SetText(21, ssPOEntry.ActiveRow, "")
                            Call ssSetFocus(ssPOEntry.ActiveRow, 20)
                            ssPOEntry.Focus()
                            Exit Sub
                        End If
                    End If
                End With
                rsAbatmentRate.ResultSetClose()
                rsAbatmentRate = Nothing
            End If
        ElseIf ssPOEntry.ActiveCol = 27 And e.keyCode = Keys.F1 Then
            If Not gblnGSTUnit Then Exit Sub
            If gblnGSTUnit And Not _blnEOUFlag Then Exit Sub
            With ssPOEntry
                .Row = .ActiveRow : .Col = .ActiveCol
                varHelpItem = ShowList(1, 6, Trim(.Text), "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND Tx_TaxeID='CGST' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))")
                If varHelpItem = "-1" Then
                    MsgBox("CGST Type Does Not Exist", vbInformation, ResolveResString(100))
                    .Text = ""
                Else
                    Call ssPOEntry.SetText(27, ssPOEntry.ActiveRow, varHelpItem)
                End If
            End With
        ElseIf ssPOEntry.ActiveCol = 28 And e.keyCode = Keys.F1 Then
            If Not gblnGSTUnit Then Exit Sub
            If gblnGSTUnit And Not _blnEOUFlag Then Exit Sub
            With ssPOEntry
                .Row = .ActiveRow : .Col = .ActiveCol
                varHelpItem = ShowList(1, 6, Trim(.Text), "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND Tx_TaxeID='SGST' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))")
                If varHelpItem = "-1" Then
                    MsgBox("SGST Type Does Not Exist", vbInformation, ResolveResString(100))
                    .Text = ""
                Else
                    Call ssPOEntry.SetText(28, ssPOEntry.ActiveRow, varHelpItem)
                End If
            End With
        ElseIf ssPOEntry.ActiveCol = 29 And e.keyCode = Keys.F1 Then
            If Not gblnGSTUnit Then Exit Sub
            If gblnGSTUnit And Not _blnEOUFlag Then Exit Sub
            With ssPOEntry
                .Row = .ActiveRow : .Col = .ActiveCol
                varHelpItem = ShowList(1, 6, Trim(.Text), "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND Tx_TaxeID='UTGST' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))")
                If varHelpItem = "-1" Then
                    MsgBox("UTGST Type Does Not Exist", vbInformation, ResolveResString(100))
                    .Text = ""
                Else
                    Call ssPOEntry.SetText(29, ssPOEntry.ActiveRow, varHelpItem)
                End If
            End With
        ElseIf ssPOEntry.ActiveCol = 30 And e.keyCode = Keys.F1 Then
            If Not gblnGSTUnit Then Exit Sub
            If gblnGSTUnit And Not _blnEOUFlag Then Exit Sub
            With ssPOEntry
                .Row = .ActiveRow : .Col = .ActiveCol
                varHelpItem = ShowList(1, 6, Trim(.Text), "TxRt_Rate_No", "TxRt_Percentage", "Gen_TaxRate", "AND Tx_TaxeID='IGST' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))")
                If varHelpItem = "-1" Then
                    MsgBox("IGST Type Does Not Exist", vbInformation, ResolveResString(100))
                    .Text = ""
                Else
                    Call ssPOEntry.SetText(30, ssPOEntry.ActiveRow, varHelpItem)
                End If
            End With
        End If
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
        Dim rsSalesParameter As ClsResultSetDB
        rsSalesParameter = New ClsResultSetDB
        Dim rsAbatmentRate As New ClsResultSetDB
        Dim varMRP, varAbatment, varAccessibleRateforMRP As Object
        If e.newRow < 1 Then Exit Sub
        If ValidRowData(e.row, e.col) = True Then
            If (e.col = 2) Or (e.col = 4) Then
                With ssPOEntry
                    .Col = 2 : .Row = e.row
                    If Len(Trim(.Text)) > 0 Then
                        .Col = 4 : .Row = e.row
                        If Len(Trim(.Text)) > 0 Then
                            'If UCase(Trim(cmbPOType.Text)) <> "JOB WORK" Then
                            If Not (UCase(Trim(cmbPOType.Text)) = "Q-TRADING" Or UCase(Trim(cmbPOType.Text)) = "JOB WORK") And (gblnGSTUnit = False Or (gblnGSTUnit And _blnEOUFlag)) Then
                                If ChkCT2Reqd.Checked = False Then
                                    If addExciseDuty(e.row) = False Then
                                        ssPOEntry.MaxRows = ssPOEntry.MaxRows - 1
                                    End If
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
            If e.col = 19 Then
                Call ssPOEntry.GetText(19, e.row, varMRP)
                Call ssPOEntry.GetText(20, e.row, varAbatment)
                If varAbatment <> "" Then
                    m_strSql = "select txrt_percentage from Gen_TaxRate where unit_code='" & gstrUNITID & "' and Tx_TaxeID='ABNT' and txrt_rate_no='" & varAbatment & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
                    rsAbatmentRate = New ClsResultSetDB
                    rsAbatmentRate.GetResult(m_strSql)
                    If rsAbatmentRate.GetNoRows > 0 Then
                        varAbatment = CLng(rsAbatmentRate.GetValue("txrt_percentage"))
                        varAccessibleRateforMRP = varMRP - (Trim(varMRP) * Trim(varAbatment)) / 100
                        Call ssPOEntry.SetText(21, e.row, varAccessibleRateforMRP)
                    End If
                End If
            End If
        Else
            With ssPOEntry
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
                        MsgBox("You have Entered 0 Tool Cost", vbInformation, ResolveResString(100))
                        .Row = e.newRow : .Col = e.newCol : .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                        .Focus()
                    End If
                End If
            End With
        End If
        If (e.col = 9) Then
            With ssPOEntry
                .Row = e.row : .Col = 9
                If Len(Trim(.Text)) > 0 Then
                    rsSalesParameter.GetResult("Select Txrt_Rate_no,TxRt_Percentage from Gen_TaxRate where unit_code='" & gstrUNITID & "' and Tx_TaxeID = 'PKT' and Txrt_Rate_no = '" & .Text & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))")
                    If rsSalesParameter.GetNoRows = 0 Then
                        MsgBox("Invalid Packing Code.", vbInformation, ResolveResString(100))
                        .Row = e.row : .Col = 9
                        .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                        .Focus()
                    End If
                End If
            End With
        End If
        If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Or cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
            If (e.col = 2) Or (e.col = 4) Then
                Dim strDrgNo As Object
                Dim GetDetails As Boolean
                Dim rsitem As ClsResultSetDB
                Dim strcustdtl As String
                Dim StrItemCode As Object
                rsitem = New ClsResultSetDB
                If e.col = 2 Then
                    strDrgNo = Nothing
                    Call ssPOEntry.GetText(e.col, ssPOEntry.MaxRows, strDrgNo)

                    If ChkCT2Reqd.Checked = True Then
                        strcustdtl = "Select ITem_code,Drg_desc from custITem_Mst where ACTIVE = 1 and unit_code='" & gstrUNITID & "' and cust_drgNo ='" & strDrgNo & "' and Account_code ='" & txtCustomerCode.Text & "'"
                        rsitem.GetResult(strcustdtl)
                        If rsitem.GetNoRows > 1 Then
                            GetDetails = False
                            MsgBox("This Part code has more then two Items linked, Please select one from Item ListBox", vbInformation, ResolveResString(100))
                            SetCellTypeCombo(e.row)
                            Call ssSetFocus(e.row, 4)
                            Exit Sub
                        End If
                    End If

                    Dim strCT2Condition As String = ""

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
                        End If
                    End If
                End If
                If e.col = 4 Then
                    Call SetCellStatic(e.row)
                End If
                If GetDetails = True Then
                    If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                        'ISSUE ID : 10515727
                        'm_strSql = "Select Account_Code,Cust_Ref,Amendment_No,Item_Code,Rate,Order_Qty,Despatch_Qty,Active_Flag,Cust_Mtrl,Cust_DrgNo,Packing,Others,Excise_Duty,Cust_Drg_Desc,Tool_Cost,Authorized_flag,OpenSO,SalesTax_Type,PerValue,InternalSONo,RevisionNo,Remarks,MRP,Abantment_Code,AccessibleRateforMRP,Packing_Type,TOOL_AMOR_FLAG,ShowInAuth,ADD_Excise_Duty,discount_type,discount_value from Cust_ord_dtl where unit_code='" & gstrUNITID & "' and Item_code ='" & StrItemCode & "' and cust_drgNo ='" & strDrgNo & "' and cust_ref ='" & Trim(txtReferenceNo.Text) & "' and Active_flag ='A' and Authorized_Flag =1"
                        m_strSql = "Select Account_Code,Cust_Ref,Amendment_No,Item_Code,Rate,Order_Qty,Despatch_Qty,Active_Flag,Cust_Mtrl,Cust_DrgNo,Packing,Others,Excise_Duty,Cust_Drg_Desc,Tool_Cost,Authorized_flag,OpenSO,SalesTax_Type,PerValue,InternalSONo,RevisionNo,Remarks,MRP,Abantment_Code,AccessibleRateforMRP,Packing_Type,TOOL_AMOR_FLAG,ShowInAuth,ADD_Excise_Duty,discount_type,discount_value from Cust_ord_dtl where unit_code='" & gstrUNITID & "' and account_code='" & txtCustomerCode.Text & "' and Item_code ='" & StrItemCode & "' and cust_drgNo ='" & strDrgNo & "' and cust_ref ='" & Trim(txtReferenceNo.Text) & "' and Active_flag ='A' and Authorized_Flag =1"
                    Else
                        'm_strSql = "Select Account_Code,Cust_Ref,Amendment_No,Item_Code,Rate,Order_Qty,Despatch_Qty,Active_Flag,Cust_Mtrl,Cust_DrgNo,Packing,Others,Excise_Duty,Cust_Drg_Desc,Tool_Cost,Authorized_flag,OpenSO,SalesTax_Type,PerValue,InternalSONo,RevisionNo,Remarks,MRP,Abantment_Code,AccessibleRateforMRP,Packing_Type,TOOL_AMOR_FLAG,ShowInAuth,ADD_Excise_Duty , discount_type,discount_value  from Cust_ord_dtl where unit_code='" & gstrUNITID & "' and Item_code ='" & StrItemCode & "' and cust_drgNo ='" & strDrgNo & "' and cust_ref ='" & Trim(txtReferenceNo.Text) & "' and amendment_no='" & Trim(txtAmendmentNo.Text) & "' and Active_flag ='A'"
                        m_strSql = "Select Account_Code,Cust_Ref,Amendment_No,Item_Code,Rate,Order_Qty,Despatch_Qty,Active_Flag,Cust_Mtrl,Cust_DrgNo,Packing,Others,Excise_Duty,Cust_Drg_Desc,Tool_Cost,Authorized_flag,OpenSO,SalesTax_Type,PerValue,InternalSONo,RevisionNo,Remarks,MRP,Abantment_Code,AccessibleRateforMRP,Packing_Type,TOOL_AMOR_FLAG,ShowInAuth,ADD_Excise_Duty , discount_type,discount_value from Cust_ord_dtl where unit_code='" & gstrUNITID & "' and account_code='" & txtCustomerCode.Text & "' and Item_code ='" & StrItemCode & "' and cust_drgNo ='" & strDrgNo & "' and cust_ref ='" & Trim(txtReferenceNo.Text) & "' and amendment_no='" & Trim(txtAmendmentNo.Text) & "' and Active_flag ='A'"
                    End If
                    'ISSUE ID : 10515727
                    rsitem.GetResult(m_strSql)
                    If rsitem.GetNoRows > 0 Then
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
                        Call ssPOEntry.SetText(6, ssPOEntry.MaxRows, (rsitem.GetValue("Rate") * Val(ctlPerValue.Text)))
                        Call ssPOEntry.SetText(14, ssPOEntry.MaxRows, rsitem.GetValue("Cust_Mtrl"))
                        Call ssPOEntry.SetText(7, ssPOEntry.MaxRows, (rsitem.GetValue("Cust_Mtrl") * Val(ctlPerValue.Text)))
                        Call ssPOEntry.SetText(15, ssPOEntry.MaxRows, rsitem.GetValue("Tool_Cost"))
                        Call ssPOEntry.SetText(8, ssPOEntry.MaxRows, (rsitem.GetValue("Tool_Cost") * Val(ctlPerValue.Text)))
                        Call ssPOEntry.SetText(9, ssPOEntry.MaxRows, rsitem.GetValue("Packing_Type"))
                        Call ssPOEntry.SetText(10, ssPOEntry.MaxRows, rsitem.GetValue("Excise_Duty"))
                        Call ssPOEntry.SetText(16, ssPOEntry.MaxRows, rsitem.GetValue("Others"))
                        Call ssPOEntry.SetText(11, ssPOEntry.MaxRows, (rsitem.GetValue("Others") * Val(ctlPerValue.Text)))
                        Call ssPOEntry.SetText(17, ssPOEntry.MaxRows, rsitem.GetValue("Despatch_qty"))
                        Call ssPOEntry.SetText(19, ssPOEntry.MaxRows, rsitem.GetValue("MRP"))
                        Call ssPOEntry.SetText(20, ssPOEntry.MaxRows, rsitem.GetValue("Abantment_code"))
                        Call ssPOEntry.SetText(21, ssPOEntry.MaxRows, rsitem.GetValue("AccessibleRateforMRP"))
                        ssPOEntry.Col = 22
                        ssPOEntry.Row = ssPOEntry.MaxRows
                        ssPOEntry.CellType = FPSpreadADO.CellTypeConstants.CellTypeComboBox
                        Call ssPOEntry.SetText(22, ssPOEntry.MaxRows, rsitem.GetValue("DISCOUNT_TYPE"))
                        ssPOEntry.TypeComboBoxList = "NONE" & Chr(9) & "[P]ercentage" & Chr(9) & "[V]alue"
                        If rsitem.GetValue("DISCOUNT_TYPE") = "[P]ercentage" Then
                            ssPOEntry.TypeComboBoxCurSel = 1
                        ElseIf rsitem.GetValue("DISCOUNT_TYPE") = "[V]alue" Then
                            ssPOEntry.TypeComboBoxCurSel = 2
                        Else
                            ssPOEntry.TypeComboBoxCurSel = 0
                        End If


                        ssPOEntry.Col = 23
                        ssPOEntry.Row = ssPOEntry.MaxRows
                        ssPOEntry.CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                        Call ssPOEntry.SetText(23, ssPOEntry.MaxRows, rsitem.GetValue("DISCOUNT_VALUE"))

                        Call ssSetFocus(ssPOEntry.MaxRows, 3)
                    Else
                        Call ssPOEntry.SetText(5, ssPOEntry.MaxRows, "0")
                        Call ssPOEntry.SetText(6, ssPOEntry.MaxRows, "0")
                        Call ssPOEntry.SetText(8, ssPOEntry.MaxRows, "0")
                        Call ssPOEntry.SetText(11, ssPOEntry.MaxRows, "0")
                        Call ssPOEntry.SetText(12, ssPOEntry.MaxRows, "0")
                        ssPOEntry.Col = 22
                        ssPOEntry.Row = ssPOEntry.MaxRows
                        ssPOEntry.CellType = FPSpreadADO.CellTypeConstants.CellTypeComboBox
                        Call ssPOEntry.SetText(22, ssPOEntry.MaxRows, rsitem.GetValue("DISCOUNT_TYPE"))
                        ssPOEntry.TypeComboBoxList = "NONE" & Chr(9) & "[P]ercentage" & Chr(9) & "[V]alue"
                        If Me.CmbDiscounttype.Text = "[P]ercentage" Then
                            ssPOEntry.TypeComboBoxCurSel = 1
                        ElseIf CmbDiscounttype.Text = "[V]alue" Then
                            ssPOEntry.TypeComboBoxCurSel = 2
                        Else
                        End If

                    End If
                End If
            End If
        End If
        Call ToSetcustdrgDesc(e.newRow)
        Exit Sub            'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub txtReferenceNo_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtReferenceNo.Validating
        On Error GoTo ErrHandler
        Dim rsRefNo As New ClsResultSetDB
        Dim rsSalesParameter As New ClsResultSetDB
        Dim strAns As MsgBoxResult
        Dim varPOFlg As Object
        Dim intAnswer As Short
        Dim intbase As Short
        Dim Cancel As Boolean = e.Cancel
        mvalid = False
        If m_blnCloseFlag = True Then 'incase close button is clicked then exit
            m_blnCloseFlag = False
            Exit Sub
        End If
        If m_blnHelpFlag = True Then 'incase help button is clicked then exit
            m_blnHelpFlag = False
            'Exit Sub
        End If
        If Len(Trim(txtReferenceNo.Text)) > 0 Then 'if reference no is not blank
            m_strSql = " Select cust_ref from cust_ord_hdr where unit_code='" & gstrUNITID & "' and cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'and Account_Code='" & Trim(Me.txtCustomerCode.Text) & "'"
            Call rsRefNo.GetResult(m_strSql)
            If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                If rsRefNo.GetNoRows = 1 Then
                    intbase = 1
                End If
                If rsRefNo.GetNoRows > 1 Then
                    intAnswer = MsgBox("Would You Like to View Base SO", MsgBoxStyle.YesNo, ResolveResString(100))
                    If intAnswer = 6 Then
                        txtAmendmentNo.Enabled = False : txtAmendmentNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                        intbase = 1
                    Else
                        txtAmendmentNo.Enabled = True : txtAmendmentNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        intbase = rsRefNo.GetNoRows
                    End If
                End If
            Else
                intbase = rsRefNo.GetNoRows
            End If
            If intbase = 1 Then 'if only one record for the reference no is existing
                If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                    m_strSql = " Select cust_ref,exportsotype from cust_ord_hdr where unit_code='" & gstrUNITID & "' and cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'and Account_Code='" & Trim(Me.txtCustomerCode.Text) & "' and Active_Flag='A' and (Authorized_flag=1 or Future_SO =1)"
                    Call rsRefNo.GetResult(m_strSql)
                    If rsRefNo.GetNoRows > 0 Then
                        'strans = MsgBox("This Reference No Already Exists and is authorized. Do You Wish to Enter An Amendment?", vbYesNo)
                        strAns = ConfirmWindow(10131, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                        If strAns = MsgBoxResult.Yes Then
                            txtAmendmentNo.Enabled = True
                            txtAmendmentNo.BackColor = System.Drawing.Color.White
                            ''
                            If rsRefNo.GetValue("exportsotype") = "With Pay" Or rsRefNo.GetValue("exportsotype") = "Without Pay" Then
                                cmbExporttype.Enabled = True
                                cmbExporttype.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                            End If
                            ''
                            cmdHelp(3).Enabled = False
                            DTAmendmentDate.Value = GetServerDate()
                            If mblnDiscountFunctionality = True Then
                                CmbDiscounttype.Enabled = True
                                CmbDiscounttype.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                                txtdiscountvalue.Enabled = True
                                txtdiscountvalue.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                            End If
                            If mblnMarkupFunctionality = True Then
                                CmbMarkuptype.Enabled = True
                                CmbMarkuptype.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                                txtmarkupvalue.Enabled = True
                                txtmarkupvalue.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                            End If
                            'txtmarkupvalue.Enabled = True
                            'txtmarkupvalue.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                            Call GetReferenceDetails()
                            ssPOEntry.MaxRows = 1
                            Call SSMaxLength()
                            cmdchangetype.Enabled = True
                            If txtAmendmentNo.Enabled Then txtAmendmentNo.Focus()
                            mvalid = False
                            Cancel = True
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
                        cmdForms.Enabled = False
                        Exit Sub
                    End If
                ElseIf cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                    m_strSql = " Select top 1 1 from cust_ord_hdr where unit_code='" & gstrUNITID & "' and cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'and Account_Code='" & Trim(Me.txtCustomerCode.Text) & "' and Active_Flag='A' and amendment_No =''"
                    Call rsRefNo.GetResult(m_strSql)
                    If rsRefNo.GetNoRows > 0 Then
                        m_strSql = " Select cust_ref from cust_ord_hdr where unit_code='" & gstrUNITID & "' and cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'and Account_Code='" & Trim(Me.txtCustomerCode.Text) & "' and Active_Flag ='A' and Authorized_flag=0 and amendment_No =''"
                        Call rsRefNo.GetResult(m_strSql)
                        If rsRefNo.GetNoRows > 0 Then
                            m_strSql = " Select cust_ref from cust_ord_hdr where unit_code='" & gstrUNITID & "' and cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'and Account_Code='" & Trim(Me.txtCustomerCode.Text) & "' and Active_Flag ='A' and future_SO=1 and amendment_No =''"
                            Call rsRefNo.GetResult(m_strSql)
                            If rsRefNo.GetNoRows > 0 Then
                                MsgBox("This Is Future SO(AUTHORISED)", MsgBoxStyle.Information, ResolveResString(100))
                            End If
                            Call GetReferenceDetails()
                            cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD) = True
                            cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = True
                            cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = True
                            cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE) = True
                            cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL) = False
                            cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = True
                        Else
                            Call ConfirmWindow(10133, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            Call GetReferenceDetails()
                            mvalid = True
                            Exit Sub
                        End If
                    Else
                        Call ConfirmWindow(10134, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                        Call GetReferenceDetails()
                        mvalid = True
                        Exit Sub
                    End If
                End If
                ' If No of records for the reference no is more than 1 that means amendment no exists
            ElseIf intbase > 1 Then
                If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                    rsSalesParameter.GetResult("Select AppendSOItem from Sales_parameter where unit_code='" & gstrUNITID & "'")
                    'ISSUE ID : 10763705
                    If rsSalesParameter.GetValue("AppendSOItem") = True And mblnappendsoitem_customer = True Then
                        'ISSUE ID : 10763705
                        'incase an amendment already exists and is not authorized
                        m_strSql = " Select top 1 1 from cust_ord_hdr where unit_code='" & gstrUNITID & "' and cust_ref ='" & Trim(Me.txtReferenceNo.Text) & "'and Account_Code='" & Trim(Me.txtCustomerCode.Text) & "' and Active_Flag ='A' and (Authorized_flag=0 and future_SO = 0)"
                        Call rsRefNo.GetResult(m_strSql)
                        If rsRefNo.GetNoRows > 0 Then 'incase a not authorized amendment exists
                            cmdButtons.Focus()
                            Call ConfirmWindow(10135, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            txtReferenceNo.Text = ""
                            txtReferenceNo.Focus()
                            mvalid = True
                            Exit Sub
                        Else
                            Me.txtAmendmentNo.Enabled = True
                            txtAmendmentNo.BackColor = System.Drawing.Color.White
                            Call GetReferenceDetails()
                            ssPOEntry.MaxRows = 1
                            Call SSMaxLength()
                            Exit Sub
                        End If
                    Else
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
                    With ssPOEntry
                        .Col = 19
                        .Col2 = 19
                        .ColHidden = True
                        .Col = 20
                        .Col2 = 20
                        .ColHidden = True
                        .Col = 21
                        .Col2 = 21
                        .ColHidden = True
                    End With
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
                    CmbDiscounttype.Enabled = True
                    CmbDiscounttype.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    txtdiscountvalue.Enabled = True
                    txtdiscountvalue.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    txtmarkupvalue.Enabled = True
                    txtmarkupvalue.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    CmbMarkuptype.Enabled = True
                    CmbMarkuptype.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    cmbExporttype.Enabled = False
                    cmbExporttype.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    ssPOEntry.Enabled = True
                    cmdchangetype.Enabled = True
                    txtAmendmentNo.Enabled = False
                    txtAmendmentNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                    cmdHelp(3).Enabled = False
                    'for account Plug in
                    '1.S.Tax at Header Lavel
                    '2.Credit Terms at Main Screen
                    If gblnGSTUnit = False Then
                        txtSTax.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : txtSTax.Enabled = True : cmdHelp(4).Enabled = True
                    End If

                    If blnNoneditableCreditTerms_onSO Then
                        txtCreditTerms.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : txtCreditTerms.Enabled = False : cmdHelp(5).Enabled = False
                    Else
                        txtCreditTerms.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : txtCreditTerms.Enabled = True : cmdHelp(5).Enabled = True
                    End If

                    cmdForms.Enabled = True
                    SetFormButtonStyle()
                    If gblnGSTUnit = False Then
                        txtSChSTax.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : txtSChSTax.Enabled = True : cmdHelp(6).Enabled = True
                    End If

                    ctlPerValue.Enabled = True
                    ctlPerValue.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                    ctlPerValue.Text = 1
                    chkOpenSo.Enabled = True
                    If gblnGSTUnit = False Then
                        txtAddVAT.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        txtAddVAT.Enabled = True
                        cmdAddVAT.Enabled = True
                    End If
                    '10869290
                    If cmbPOType.Text = "V-SERVICE" Then
                        txtService.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        txtService.Enabled = True
                        cmdServiceTax.Enabled = True
                        txtSBC.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        txtSBC.Enabled = True
                        cmdSBCtax.Enabled = True
                        txtKKC.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        txtKKC.Enabled = True
                        cmdKKCtax.Enabled = True
                    End If
                    Call ADDRow()
                    Call GetDetailsFromAgreementMaster(txtCustomerCode.Text, txtReferenceNo.Text)
                    DTDate.Focus()
                    Cancel = True
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
        End With
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub DTDate_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles DTDate.Validating
        On Error GoTo ErrHandler
        If DTDate.Value > DateValue(CStr(FinancialYearDates(eMPowerFunctions.FinancialYearDatesEnum.DATE_END))) Then
            Call ConfirmWindow(10074, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
            e.Cancel = True
            Exit Sub
        End If
        If DTDate.Value > GetServerDate() Then
            MsgBox("Date Can not be greater than Current Date", MsgBoxStyle.OkOnly, ResolveResString(100))
            e.Cancel = True
            Exit Sub
        End If
        dtSODate = DTDate.Value
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub cmdAddVAT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddVAT.Click
        Dim strHelp As String
        Dim strSTaxHelp() As String
        On Error GoTo ErrHandler
        Select Case Me.cmdButtons.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                strHelp = "Select TxRt_Rate_No,TxRt_RateDesc from Gen_taxRate where unit_code='" & gstrUNITID & "' and Tx_TaxeID in('ADVAT','ADCST') and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
                strSTaxHelp = Me.ctlEMPHelp1.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, strHelp, "Add. VAT/CST Tax Help")
                If UBound(strSTaxHelp) <= 0 Then Exit Sub
                If strSTaxHelp(0) = "0" Then
                    Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK) : txtAddVAT.Text = "" : txtAddVAT.Focus() : Exit Sub
                Else
                    txtAddVAT.Text = strSTaxHelp(0)
                    lbladdvatdesc.Text = strSTaxHelp(1)
                End If
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub txtAddVAT_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtAddVAT.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.cmdButtons.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        If Len(txtAddVAT.Text) > 0 Then
                            Call txtAddVAT_Validating(txtAddVAT, New System.ComponentModel.CancelEventArgs(False))
                        Else
                            chkOpenSo.Focus()
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
            lbladdvatdesc.Text = ""
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
            If CheckExistanceOfFieldData((txtAddVAT.Text), "TxRt_Rate_No", "Gen_TaxRate", " (Tx_TaxeID='ADVAT' OR Tx_TaxeID='ADCST') and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))") Then
                Call FillLabel("ADDVAT")
                chkOpenSo.Focus()
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
            strTableSql = "select " & Trim(pstrColumnName) & " from " & Trim(pstrTableName) & " where unit_code='" & gstrUNITID & "' and " & Trim(pstrColumnName) & "='" & Trim(pstrFieldText) & "' and " & pstrCondition
        Else
            strTableSql = "select " & Trim(pstrColumnName) & " from " & Trim(pstrTableName) & " where unit_code='" & gstrUNITID & "' and " & Trim(pstrColumnName) & "='" & Trim(pstrFieldText) & "'"
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
    Private Function CheckForMultipleOpenSO(ByVal Account_Code As String, ByVal Cust_ref As String, ByVal Cust_DrgNo As String, ByVal Item_code As String) As String
        Dim rstHelpDb As ClsResultSetDB
        Dim blnopenclosedso As Boolean

        Try
            rstHelpDb = New ClsResultSetDB
            If chkOpenSo.Checked Then
                blnopenclosedso = 1
            Else
                blnopenclosedso = 0
            End If

            Call rstHelpDb.GetResult("Select dbo.UDF_CHECK_ACTIVE_SO_ITEM('" & gstrUNITID & "' ,'" & Account_Code & "','" & Account_Code & "','" & Cust_ref & "','" & Cust_DrgNo & "','" & Item_code & "','" & blnopenclosedso & "') as ActiveSalesOrder", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
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
    Private Sub txtAmendmentNo_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtAmendmentNo.Validating
        On Error GoTo ErrHandler
        Dim inti As Short
        Dim rsAmend As New ClsResultSetDB
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
            m_strSql = "Select top 1 1 from cust_ord_hdr where unit_code='" & gstrUNITID & "' and Amendment_No='" & Trim(txtAmendmentNo.Text) & "' and Cust_Ref ='" & Trim(txtReferenceNo.Text) & "'and Account_Code='" & Trim(txtCustomerCode.Text) & "'  "
            rsAmend.GetResult(m_strSql)
            If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                If rsAmend.GetNoRows > 0 Then
                    cmdButtons.Focus()
                    Call ConfirmWindow(10141, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                    txtAmendmentNo.Text = ""
                    mvalid = True
                    txtAmendmentNo.Focus()
                    Exit Sub
                Else
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
                    If gblnGSTUnit = False Then
                        txtAddVAT.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : txtAddVAT.Enabled = True : cmdAddVAT.Enabled = True
                    End If

                    cmdForms.Enabled = True
                    SetFormButtonStyle()
                    If blnNoneditableCreditTerms_onSO Then
                        txtCreditTerms.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : txtCreditTerms.Enabled = False : cmdHelp(5).Enabled = False
                    Else
                        txtCreditTerms.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : txtCreditTerms.Enabled = True : cmdHelp(5).Enabled = True
                    End If
                    chkOpenSo.Enabled = True
                    '1.Surcharge on S.Tax
                    If gblnGSTUnit = False Then
                        txtSChSTax.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : txtSChSTax.Enabled = True : cmdHelp(6).Enabled = True
                    End If

                    With Me.ssPOEntry
                        .Enabled = True
                        .Row = 1
                        .Row2 = .MaxRows
                        .Col = 5
                        .Col2 = .MaxCols
                        .BlockMode = True
                        .Lock = False
                        .BlockMode = False
                        '***********
                    End With
                End If
            ElseIf cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                m_strSql = "Select top 1 1 from cust_ord_hdr where unit_code='" & gstrUNITID & "' and Amendment_No='" & Trim(txtAmendmentNo.Text) & "' and Cust_Ref ='" & Trim(txtReferenceNo.Text) & "'and Account_Code='" & Trim(txtCustomerCode.Text) & "'"
                rsAmend.GetResult(m_strSql)
                If rsAmend.GetNoRows > 0 Then
                    m_strSql = "Select top 1 1 from cust_ord_hdr where unit_code='" & gstrUNITID & "' and Amendment_No='" & Trim(txtAmendmentNo.Text) & "' and Cust_Ref ='" & Trim(txtReferenceNo.Text) & "'and Account_Code='" & Trim(txtCustomerCode.Text) & "' and Active_Flag='A' "
                    rsAmend.GetResult(m_strSql)
                    If rsAmend.GetNoRows > 0 Then
                        m_strSql = "Select top 1 1 from cust_ord_hdr where unit_code='" & gstrUNITID & "' and Amendment_No='" & Trim(txtAmendmentNo.Text) & "' and Cust_Ref ='" & Trim(txtReferenceNo.Text) & "'and Account_Code='" & Trim(txtCustomerCode.Text) & "' and Active_Flag='A' and (authorized_Flag=1) "
                        rsAmend.GetResult(m_strSql)
                        If rsAmend.GetNoRows > 0 Then
                            Call ConfirmWindow(10142, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            cmdButtons.Focus()
                            Call GetAmendmentDetails()
                            mvalid = True
                            cmdButtons.Focus()

                            Exit Sub
                        Else
                            m_strSql = "Select top 1 1 from cust_ord_hdr where unit_code='" & gstrUNITID & "' and Amendment_No='" & Trim(txtAmendmentNo.Text) & "' and Cust_Ref ='" & Trim(txtReferenceNo.Text) & "'and Account_Code='" & Trim(txtCustomerCode.Text) & "' and Active_Flag='A' and (future_so =1) "
                            rsAmend.GetResult(m_strSql)
                            If rsAmend.GetNoRows > 0 Then
                                MsgBox("This is Future SO(AUTHORISED).", MsgBoxStyle.Information, ResolveResString(100))
                            End If
                            Call GetAmendmentDetails()
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
                        Call ConfirmWindow(10143, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                        Call GetAmendmentDetails()
                        mvalid = True
                        cmdButtons.Focus()
                        Exit Sub
                    End If
                Else
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
        m_strSql = "Select top 1  1 from cust_ord_hdr where unit_code='" & gstrUNITID & "' and Cust_Ref ='" & Trim(txtReferenceNo.Text) & "'and Account_Code='" & Trim(txtCustomerCode.Text) & "'and po_type='Q'"
        rsAmend.GetResult(m_strSql)
        If rsAmend.GetNoRows > 0 Then
            cmbPOType.Text = "Q-TRADING"
        End If
        mvalid = False
        cmdButtons.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = True
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Function GetDetailsFromAgreementMaster(ByVal Account_Code As String, ByVal Cust_ref As String)
        Dim rstHelpDb As ClsResultSetDB
        Dim rsGetRate As ClsResultSetDB
        Dim RSCHECK As ClsResultSetDB
        Dim rsGetTaxRate As ClsResultSetDB
        Dim strmessage As String
        strmessage = "Customer Reference No. is Invalid. "
        strmessage = strmessage & vbCrLf & " 1.Customer Reference No. Should be same as agreement No."
        strmessage = strmessage & vbCrLf & " 2.No Sales Order Should Exist for the Customer Reference No. "
        Try
            RSCHECK = New ClsResultSetDB
            Call RSCHECK.GetResult("SELECT TOP 1 1 FROM CUSTOMER_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE = '" & Account_Code.ToString.Trim & "' and INVOICEAGAINSTAGREEMENTMST = '1' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If RSCHECK.GetNoRows > 0 Then
                rstHelpDb = New ClsResultSetDB
                Call rstHelpDb.GetResult("SELECT InternalAssyPart,CustAssyPart FROM AGREEMENT_HDR H WHERE UNIT_CODE='" & gstrUNITID & "'  AND NOT EXISTS ( SELECT TOP 1 1 FROM CUST_ORD_HDR I WHERE h.unit_code=i.unit_code and H.CUSTOMER_CODE = I.ACCOUNT_CODE AND CONVERT(VARCHAR(50),H.DOC_NO) = I.CUST_REF   AND I.UNIT_CODE='" & gstrUNITID & "') AND CONVERT(VARCHAR(50),H.DOC_NO) = '" & Cust_ref.ToString.Trim & "' AND CUSTOMER_CODE = '" & Account_Code.ToString.Trim & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                If rstHelpDb.GetNoRows = 0 Then
                    MsgBox(strmessage, MsgBoxStyle.Information, ResolveResString(100))
                Else
                    With ssPOEntry
                        .Row = .MaxRows : .Col = 4 : .Text = rstHelpDb.GetValue("InternalAssyPart")
                        .Row = .MaxRows : .Col = 2 : .Text = rstHelpDb.GetValue("CustAssyPart")
                        rsGetRate = New ClsResultSetDB
                        Call rsGetRate.GetResult("Select dbo.UDF_GETRATEOFITEMFROMAGREEMENTMASTER('" & gstrUNITID & "','" & Account_Code & "','" & Cust_ref & "','" & rstHelpDb.GetValue("CustAssyPart").ToString.Trim & "','" & rstHelpDb.GetValue("InternalAssyPart").ToString.Trim & "') as Basic_Rate", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        If rsGetRate.GetNoRows >= 1 Then
                            .Row = .MaxRows : .Col = 6 : .Text = rsGetRate.GetValue("Basic_Rate")
                        End If
                        rsGetRate.ResultSetClose()
                        rsGetRate = Nothing
                        rsGetTaxRate = New ClsResultSetDB
                        Call rsGetTaxRate.GetResult("Select Top 1 TaxValue from AgreementTaxDtl where UNIT_CODE='" & gstrUNITID & "'  AND TaxId = 'PKT' and CONVERT(VARCHAR(50),DOC_NO) = '" & Cust_ref.ToString.Trim & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        If rsGetTaxRate.GetNoRows >= 1 Then
                            .Row = .MaxRows : .Col = 9 : .Text = rsGetTaxRate.GetValue("TaxValue")
                        End If
                        rsGetTaxRate.ResultSetClose()
                        rsGetTaxRate = Nothing
                        rsGetTaxRate = New ClsResultSetDB
                        Call rsGetTaxRate.GetResult("Select Top 1 TaxValue from AgreementTaxDtl where UNIT_CODE='" & gstrUNITID & "'  AND TaxId = 'EXC' and CONVERT(VARCHAR(50),DOC_NO) = '" & Cust_ref.ToString.Trim & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        If rsGetTaxRate.GetNoRows >= 1 Then
                            .Row = .MaxRows : .Col = 10 : .Text = rsGetTaxRate.GetValue("TaxValue")
                        End If
                        rsGetTaxRate.ResultSetClose()
                        rsGetTaxRate = Nothing
                        rsGetTaxRate = New ClsResultSetDB
                        Call rsGetTaxRate.GetResult("Select Top 1 TaxValue from AgreementTaxDtl where UNIT_CODE='" & gstrUNITID & "'  AND TaxID in ('LST','CST','VAT') and CONVERT(VARCHAR(50),DOC_NO) = '" & Cust_ref.ToString.Trim & "' ", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        If rsGetTaxRate.GetNoRows >= 1 Then
                            txtSTax.Text = rsGetTaxRate.GetValue("TaxValue")
                        End If
                        rsGetTaxRate.ResultSetClose()
                        rsGetTaxRate = Nothing
                        rsGetTaxRate = New ClsResultSetDB
                        Call rsGetTaxRate.GetResult("Select Top 1 TaxValue from AgreementTaxDtl where UNIT_CODE='" & gstrUNITID & "'  AND TaxId = 'SST' and  CONVERT(VARCHAR(50),DOC_NO) = '" & Cust_ref.ToString.Trim & "' ", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        If rsGetTaxRate.GetNoRows >= 1 Then
                            txtSChSTax.Text = rsGetTaxRate.GetValue("TaxValue")
                        End If
                        rsGetTaxRate.ResultSetClose()
                        rsGetTaxRate = Nothing
                        rsGetTaxRate = New ClsResultSetDB
                        Call rsGetTaxRate.GetResult("Select Top 1 TaxValue from AgreementTaxDtl where UNIT_CODE='" & gstrUNITID & "'  AND TaxID in('ADVAT','ADCST') and CONVERT(VARCHAR(50),DOC_NO) = '" & Cust_ref.ToString.Trim & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        If rsGetTaxRate.GetNoRows >= 1 Then
                            txtAddVAT.Text = rsGetTaxRate.GetValue("TaxValue")
                        End If
                        rsGetTaxRate.ResultSetClose()
                        rsGetTaxRate = Nothing
                    End With
                End If
                rstHelpDb.ResultSetClose()
                rstHelpDb = Nothing
            End If
            RSCHECK.ResultSetClose()
            RSCHECK = Nothing
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Finally
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        End Try
    End Function
    Public Sub AdddiscountSelection()
        On Error GoTo Err_Handler
        CmbDiscounttype.Items.Insert(0, "None")
        CmbDiscounttype.Items.Insert(1, "[P]ercentage")
        CmbDiscounttype.Items.Insert(2, "[V]alue")
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Sub AddmarkUpSelection()
        On Error GoTo Err_Handler
        CmbMarkuptype.Items.Insert(0, "None")
        CmbMarkuptype.Items.Insert(1, "[P]ercentage")
        CmbMarkuptype.Items.Insert(2, "[V]alue")
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtdiscountvalue_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
        If txtdiscountvalue.Text.Trim.Length > 0 And IsNumeric(Me.txtdiscountvalue.Text) = False Then
            MsgBox("Enter numeric values !", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100))
            Me.txtdiscountvalue.Text = ""
            Me.txtdiscountvalue.Focus()
            Exit Sub
        End If
        If CmbDiscounttype.Text = "[P]ercentage" Then
            If Val(txtdiscountvalue.Text) < 0 Or Val(txtdiscountvalue.Text) > 100 Then
                MsgBox("% Can't be Greater than 100 ", MsgBoxStyle.Information, ResolveResString(100))
                txtdiscountvalue.Text = ""
                Exit Sub
            End If
        End If

        If ssPOEntry.MaxRows >= 1 Then
            Call SSMaxLength()
        End If
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
    Public Function FindScalar_Value(ByVal strField As String) As String
        On Error GoTo ErrHandler
        Dim Find_Value As String
        Find_Value = SqlConnectionclass.ExecuteScalar(strField)
        If Trim(Find_Value + "") = "" Then
            Find_Value = ""
        End If
        FindScalar_Value = Find_Value
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Public Function ValidDiscountvalue() As Boolean
        Dim strErrMsg As String
        Dim varDiscounttype As Object
        Dim varDiscountvalue As Object
        Dim intLoopCounter As Integer
        On Error GoTo Err_Handler

        ValidDiscountvalue = False

        With ssPOEntry
            For intLoopCounter = 1 To .MaxRows

                varDiscounttype = Nothing
                Call ssPOEntry.GetText(22, intLoopCounter, varDiscounttype)
                varDiscountvalue = Nothing
                Call ssPOEntry.GetText(23, intLoopCounter, varDiscountvalue)
                If varDiscounttype = "[P]ercentage" And varDiscountvalue > 100 Then
                    ValidDiscountvalue = True
                    Exit Function
                End If

            Next

        End With
        ValidDiscountvalue = False
        Exit Function
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

    Private Sub txtdiscountvalue_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtdiscountvalue.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        On Error GoTo ErrHandler
        If KeyAscii <> 46 Then
            If KeyAscii < 48 Or KeyAscii > 57 Then
                e.KeyChar = ""
            End If
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub

    Private Sub txtdiscountvalue_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtdiscountvalue.TextChanged
        Dim strMin As String
        Dim strMax As String

        With Me.ssPOEntry
            If mblnDiscountFunctionality = True And .MaxRows > 0 Then
                .Col = 22
                .Row = .MaxRows
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeComboBox
                .TypeComboBoxList = "NONE" & Chr(9) & "[P]ercentage" & Chr(9) & "[V]alue"
                If Me.CmbDiscounttype.Text = "[P]ercentage" Then
                    .TypeComboBoxCurSel = 1
                ElseIf CmbDiscounttype.Text = "[V]alue" Then
                    .TypeComboBoxCurSel = 2
                Else
                    .TypeComboBoxCurSel = 0
                End If

                .Col = 23
                .Row = .MaxRows
                strMin = "0." : strMax = "99999999."
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                .TypeFloatDecimalPlaces = 4

                .TypeFloatMin = strMin
                .TypeFloatMax = strMax
                'Call .SetText(22, .MaxRows, CmbDiscounttype.Text.Trim)
                Call .SetText(23, .MaxRows, txtdiscountvalue.Text.Trim)
            End If
        End With
    End Sub

    Private Sub CmbDiscounttype_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles CmbDiscounttype.Validating
        Dim strMin As String
        Dim strMax As String

        With Me.ssPOEntry
            If mblnDiscountFunctionality = True And .MaxRows > 0 Then
                .Col = 22
                .Row = .MaxRows
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeComboBox
                .TypeComboBoxList = "NONE" & Chr(9) & "[P]ercentage" & Chr(9) & "[V]alue"
                If Me.CmbDiscounttype.Text = "[P]ercentage" Then
                    .TypeComboBoxCurSel = 1
                ElseIf CmbDiscounttype.Text = "[V]alue" Then
                    .TypeComboBoxCurSel = 2
                Else
                    .TypeComboBoxCurSel = 0
                End If

                .Col = 23
                .Row = .MaxRows
                strMin = "0." : strMax = "99999999."
                .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                .TypeFloatDecimalPlaces = 4

                .TypeFloatMin = strMin
                .TypeFloatMax = strMax
                'Call .SetText(22, .MaxRows, CmbDiscounttype.Text.Trim)
                Call .SetText(23, .MaxRows, txtdiscountvalue.Text.Trim)
            End If
        End With
    End Sub
    Public Function ValidMarkupvalue() As Boolean
        Dim strErrMsg As String
        Dim varmarkuptype As Object
        Dim varmarkupvalue As Object
        Dim intLoopCounter As Integer
        On Error GoTo Err_Handler

        ValidMarkupvalue = False

        With ssPOEntry
            For intLoopCounter = 1 To .MaxRows

                varmarkuptype = Nothing
                Call ssPOEntry.GetText(24, intLoopCounter, varmarkuptype)
                varmarkupvalue = Nothing
                Call ssPOEntry.GetText(25, intLoopCounter, varmarkupvalue)
                If varmarkuptype = "[P]ercentage" And varmarkupvalue > 100 Then
                    ValidMarkupvalue = True

                    Exit Function
                End If
            Next

        End With
        ValidMarkupvalue = False
        Exit Function
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

    Private Sub txtmarkupvalue_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtmarkupvalue.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        On Error GoTo ErrHandler
        If KeyAscii <> 46 Then
            If (KeyAscii < 48 Or KeyAscii > 57) Then
                e.KeyChar = ""
            End If
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub

    Private Sub txtmarkupvalue_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtmarkupvalue.Validating
        If txtmarkupvalue.Text.Trim.Length > 0 And IsNumeric(Me.txtmarkupvalue.Text) = False Then
            MsgBox("Enter numeric values !", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, ResolveResString(100))
            Me.txtmarkupvalue.Text = ""
            Me.txtmarkupvalue.Focus()
            Exit Sub
        End If
        If CmbMarkuptype.Text = "[P]ercentage" Then
            If Val(txtmarkupvalue.Text) < 0 Or Val(txtmarkupvalue.Text) > 100 Then
                MsgBox("% Can't be Greater than 100 ", MsgBoxStyle.Information, ResolveResString(100))
                txtmarkupvalue.Text = ""
                Exit Sub
            End If
        End If
        If ssPOEntry.MaxRows >= 1 Then
            Call SSMaxLength()
        End If
    End Sub

    '10736222-changes done for CT2 ARE 3 functionality - to enable CT2 Reqd flag
    Private Sub ChkCT2Reqd_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChkCT2Reqd.CheckedChanged
        Dim intCounter As Short
        Try
            If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                If ChkCT2Reqd.Checked = True Then
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
                    End If
                Else '10797956 
                    With ssPOEntry
                        '.Row = intRow : .Row2 = intRow : .Col = 10 : .Col2 = 10 : .BlockMode = True : .Lock = False : .BlockMode = False
                        For intCounter = 1 To .MaxRows
                            .Row = intCounter
                            .Col = 10
                            .Lock = False
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
                    StrSrvCHelp = Me.ctlEMPHelp1.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, StrSql, "Service Tax Code Help")
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

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSBCtax.Click
        Dim StrSql As String = String.Empty
        Dim StrSrvCHelp() As String

        Try
            Select Case Me.cmdButtons.Mode
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                    StrSql = " SELECT TXRT_RATE_NO,TXRT_RATEDESC FROM GEN_TAXRATE WHERE UNIT_CODE='" & gstrUNITID & "' AND (TX_TAXEID='SBC')  AND ((ISNULL(DEACTIVE_FLAG,0) <> 1) OR (CAST(GETDATE() AS DATE) <= DEACTIVE_DATE))"
                    StrSrvCHelp = Me.ctlEMPHelp1.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, StrSql, "SBC Tax Code Help")
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

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdKKCtax.Click
        Dim StrSql As String = String.Empty
        Dim StrSrvCHelp() As String

        Try
            Select Case Me.cmdButtons.Mode
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                    StrSql = " SELECT TXRT_RATE_NO,TXRT_RATEDESC FROM GEN_TAXRATE WHERE UNIT_CODE='" & gstrUNITID & "' AND (TX_TAXEID='KKC')  AND ((ISNULL(DEACTIVE_FLAG,0) <> 1) OR (CAST(GETDATE() AS DATE) <= DEACTIVE_DATE))"
                    StrSrvCHelp = Me.ctlEMPHelp1.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, StrSql, "KKC Tax Code Help")
                    If UBound(StrSrvCHelp) <= 0 Then Exit Sub
                    If StrSrvCHelp(0) = "0" Then
                        Call ConfirmWindow(10187, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK) : txtSBC.Text = "" : txtSBC.Focus() : Exit Sub
                    Else
                        txtKKC.Text = StrSrvCHelp(0)
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

    Private Sub txtKKC_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtKKC.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)

        Try
            Select Case KeyAscii
                Case System.Windows.Forms.Keys.Return
                    Select Case Me.cmdButtons.Mode
                        Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                            If Len(txtKKC.Text) > 0 Then
                                Call txtKKC_Validating(txtKKC, New System.ComponentModel.CancelEventArgs(False))
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

    Private Sub chkShipAddress_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkShipAddress.CheckedChanged
        Dim StrSql As String = String.Empty
        Dim StrSrvCHelp() As String

        Try
            If chkShipAddress.Checked = True Then
                Select Case Me.cmdButtons.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        StrSql = "select Distinct Shipping_Code,Shipping_Desc,Ship_State,GSTIN_ID from Customer_Shipping_Dtl where unit_code='" & gstrUNITID & "' and InActive_Flag=0 and customer_code='" & Trim(txtCustomerCode.Text) & "'"
                        StrSrvCHelp = Me.ctlEMPHelp1.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, StrSql, "Ship Address Code Help")
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

    Private Sub FillDataTables(Optional ByVal ClearDataTables As Boolean = False)

        'If ClearDataTables = False Then
        Dim strDtSql As String = String.Empty
        strDtSql = strDtSql + " SELECT CUSTITEM_MST.UNIT_CODE, CUSTITEM_MST.ACCOUNT_CODE,CUSTITEM_MST.CUST_DRGNO,CUSTITEM_MST.ITEM_CODE,ITEM_MST.HSN_SAC_CODE,"
        strDtSql = strDtSql + " ISNULL(CUSTOMER_MST.SOUPLD_DISCOUNTEDITABLE,0) SOUPLD_DISCOUNTEDITABLE,"
        strDtSql = strDtSql + " ITEM_MST.STATUS,ITEM_MST.HOLD_FLAG,ITEM_MST.ITEM_MAIN_GRP, CUSTITEM_MST.ACTIVE,CUSTOMER_MST.ALLOW_MULTIPLE_HSN_ITEMS,"
        strDtSql = strDtSql + " MEASURE_MST.DECIMAL_ALLOWED_FLAG"
        strDtSql = strDtSql + " FROM CUSTITEM_MST (NOLOCK) INNER JOIN ITEM_MST (NOLOCK) ON CUSTITEM_MST.ITEM_CODE=ITEM_MST.ITEM_CODE "
        strDtSql = strDtSql + " INNER JOIN CUSTOMER_MST (NOLOCK) ON CUSTITEM_MST.ACCOUNT_CODE=CUSTOMER_MST.CUSTOMER_CODE AND CUSTITEM_MST.UNIT_CODE=CUSTOMER_MST.UNIT_CODE"
        strDtSql = strDtSql + " AND CUSTITEM_MST.UNIT_CODE=ITEM_MST.UNIT_CODE "
        strDtSql = strDtSql + " INNER JOIN MEASURE_MST ON MEASURE_MST.UNIT_CODE=ITEM_MST.UNIT_CODE AND MEASURE_MST.MEASURE_CODE=ITEM_MST.CONS_MEASURE_CODE"
        strDtSql = strDtSql + " WHERE CUSTITEM_MST.ACCOUNT_CODE='" + txtCustomerCode.Text.Trim() + "' AND CUSTITEM_MST.UNIT_CODE='" + gstrUNITID + "'"

        datatable_MasterData = SqlConnectionclass.GetDataTable(strDtSql) 'AMIT02APR2022
        datatable_MasterData_GEN_TAXRATE = SqlConnectionclass.GetDataTable("SELECT UNIT_CODE,TX_TAXEID,TXRT_RATE_NO,TXRT_PERCENTAGE FROM GEN_TAXRATE (NOLOCK) WHERE UNIT_CODE='" + gstrUNITID + "' AND ((ISNULL(DEACTIVE_FLAG,0) <> 1) OR (CAST(GETDATE() AS DATE) <= DEACTIVE_DATE))") 'AMIT02APR2022
        datatable_ExistingSoDetails = SqlConnectionclass.GetDataTable("Select UNIT_CODE,CUST_DRGNO,ACCOUNT_CODE,CUST_REF,ACTIVE_FLAG,ITEM_CODE,AMENDMENT_NO from cust_ord_dtl where unit_code='" & gstrUNITID & "' and account_code='" & txtCustomerCode.Text & "' and cust_ref='" & txtReferenceNo.Text & "'and active_Flag='A'  and amendment_no = '" & txtAmendmentNo.Text & "'")

        '   'Else
        '  If Not IsNothing(datatable_MasterData) Then datatable_MasterData.Clear()
        ' If Not IsNothing(datatable_MasterData_GEN_TAXRATE) Then datatable_MasterData_GEN_TAXRATE.Clear()
        'If Not IsNothing(datatable_ExistingSoDetails) Then datatable_ExistingSoDetails.Clear()

        'End If

    End Sub


    Private Sub CtlHeader_Load(sender As Object, e As EventArgs) Handles ctlHeader.Load

    End Sub

    '' CADM START 

    Private Sub cmdCADMOrder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCADMOrder.Click
        Dim strCADMOrderHelp() As String
        Try
            Dim StrSql As String = "select [ACCOUNT_CODE],[CUST_REF],[ORDER_DATE],[PO_TYPE],[VALID_DATE],INTERNALSONO as [CADMREFER NO],[CUST_DRGNO],[ITEM_CODE],[ORDER_QTY],[RATE], " &
            "[TOOL_COST],[PACKING],[EXCISE_DUTY],[OTHERS],[HSNSACCODE],[CGSTTXRT_TYPE],[SGSTTXRT_TYPE],[UTGSTTXRT_TYPE],[IGSTTXRT_TYPE], " &
            "[COMPENSATION_CESS] from CADM_SALES_ORDER_CREATION  WHERE UNIT_CODE='" & gstrUNITID & "' and ISNULL(CADMORDERID,'')='' ORDER BY ORDER_DATE "
            strCADMOrderHelp = ctlEMPHelp1.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, StrSql, "CADM Sales Order Details")
            If UBound(strCADMOrderHelp) < 0 Then
                Call MsgBox("No Sale Orders Defined.", MsgBoxStyle.Information, ResolveResString(100))
            Else
                txtCustomerCode.Text = Trim(strCADMOrderHelp(0))
                Me.txtReferenceNo.Text = Trim(strCADMOrderHelp(1))
                DTDate.Value = Trim(strCADMOrderHelp(2))
                cmbPOType.Text = "OEM"
                DTValidDate.Value = Trim(strCADMOrderHelp(4))
                Me.txtCADRefNo.Text = Trim(strCADMOrderHelp(5))
                CADMCustomerValidating()
                CADMReferenceNoValidation()
                If Len(txtCustomerCode.Text) > 0 Then
                    ssPOEntry.MaxRows = 0
                    StrSql = "select [CUST_DRGNO],[ITEM_CODE],[ORDER_QTY],[RATE], " &
                    "[TOOL_COST],[PACKING],[EXCISE_DUTY],[OTHERS],[HSNSACCODE],[CGSTTXRT_TYPE],[SGSTTXRT_TYPE],[UTGSTTXRT_TYPE],[IGSTTXRT_TYPE], " &
                    "[COMPENSATION_CESS] from CADM_SALES_ORDER_CREATION  WHERE UNIT_CODE='" & gstrUNITID & "' AND CUST_REF='" & txtReferenceNo.Text & "'"
                    Dim dt As DataTable = SqlConnectionclass.GetDataTable(StrSql)
                    Dim intRow As Integer = 0
                    If dt.Rows.Count > 0 Then
                        For Each row As DataRow In dt.Rows
                            With ssPOEntry
                                ADDRowCADM()
                                intRow = intRow + 1
                                .Row = intRow
                                .Col = 2 : .Text = Convert.ToString(row("CUST_DRGNO"))
                                .Col = 4 : .Text = Convert.ToString(row("ITEM_CODE"))
                                .Col = 5 : .Text = Convert.ToString(row("ORDER_QTY"))
                                .Col = 6 : .Text = Convert.ToString(row("RATE"))
                                .Col = 8 : .Text = Convert.ToString(row("TOOL_COST"))
                                .Col = 9 : .Text = Convert.ToString(row("PACKING"))
                                .Col = 10 : .Text = Convert.ToString(row("EXCISE_DUTY"))
                                .Col = 11 : .Text = Convert.ToString(row("OTHERS"))
                                Dim strItemCode = Convert.ToString(row("ITEM_CODE"))
                                StrSql = "set dateformat 'dmy' SELECT * FROM DBO.UFN_GST_ITEMWISETAXES('" & gstrUNITID & "','" & txtCustomerCode.Text.Trim & "','" & Convert.ToString(row("ITEM_CODE")) & "','" & DTEffectiveDate.Value & "','" & DTValidDate.Value & "')"
                                Dim dtTax As DataTable = SqlConnectionclass.GetDataTable(StrSql)
                                If dtTax.Rows.Count > 0 Then
                                    .Col = 26 : .Text = Convert.ToString(dtTax.Rows(0)(1))
                                    .Col = 27 : .Text = Convert.ToString(dtTax.Rows(0)(2))
                                    .Col = 28 : .Text = Convert.ToString(dtTax.Rows(0)(3))
                                    .Col = 29 : .Text = Convert.ToString(dtTax.Rows(0)(4))
                                    .Col = 30 : .Text = Convert.ToString(dtTax.Rows(0)(5))
                                    .Col = 31 : .Text = Convert.ToString(dtTax.Rows(0)(6))
                                    .Col = 32 : .Text = Convert.ToString(dtTax.Rows(0)(0))
                                End If
                            End With
                        Next
                    End If
                End If
            End If
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub CADMCustomerValidating()
        Try
            Dim rsCD As New ClsResultSetDB
            With ssPOEntry
                .Col = 19
                .Col2 = 19
                .ColHidden = True
                .Col = 20
                .Col2 = 20
                .ColHidden = True
            End With
            If m_blnCloseFlag = True Then
                m_blnCloseFlag = False
                Exit Sub
            End If
            If Len(Trim(txtCustomerCode.Text)) = 0 Then
                Exit Sub
            Else
                m_strSql = "Select TOP 1 1 from Customer_mst where unit_code='" & gstrUNITID & "' and customer_Code='" & Trim(txtCustomerCode.Text) & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
                rsCD.GetResult(m_strSql)
                If rsCD.GetNoRows = 0 Then
                    Call ConfirmWindow(10145, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                    txtCustomerCode.Text = ""
                    txtReferenceNo.Text = ""
                    txtCADRefNo.Text = ""
                    txtCustomerCode.Focus()
                    Exit Sub
                Else
                    '10736222-changes done for CT2 ARE 3 functionality - to enable CT2 Reqd flag
                    If cmdButtons.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                        txtCustomerCode.Enabled = False
                        _cmdHelp_0.Enabled = False
                        txtReferenceNo.Enabled = False
                        strSql = "Select dbo.UDF_IsCT2Customer('" & gstrUNITID & "','" & txtCustomerCode.Text.Trim & "')"
                        ChkCT2Reqd.Enabled = Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSql))
                        FillDataTables()
                    End If
                    'ISSUE ID : 10763705 
                    strSql = "Select appendsoitem from customer_mst where unit_code='" & gstrUNITID & "' and customer_code='" & txtCustomerCode.Text.Trim & "'"
                    mblnappendsoitem_customer = Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSql))
                    'ISSUE ID : 10763705 
                    'GST CHANGES
                    strSql = "Select ALLOW_MULTIPLE_HSN_ITEMS from customer_mst where unit_code='" & gstrUNITID & "' and customer_code='" & txtCustomerCode.Text.Trim & "'"
                    MBLNALLOW_MULTIPLEHSNITEMS_CUSTOMER = Convert.ToBoolean(SqlConnectionclass.ExecuteScalar(strSql))
                    'GST CHANGES


                End If
            End If

        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Private Sub CADMReferenceNoValidation()
        Try
            With ssPOEntry
                .Col = 19
                .Col2 = 19
                .ColHidden = True
                .Col = 20
                .Col2 = 20
                .ColHidden = True
                .Col = 21
                .Col2 = 21
                .ColHidden = True
            End With
            'DTDate.Enabled = True
            'DTValidDate.Enabled = True
            'DTEffectiveDate.Enabled = True
            'txtCurrencyType.Enabled = True
            'txtCurrencyType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            'cmdHelp(2).Enabled = True
            'cmbPOType.Enabled = True
            'cmbPOType.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            CmbDiscounttype.Enabled = True
            CmbDiscounttype.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            txtdiscountvalue.Enabled = True
            txtdiscountvalue.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            txtmarkupvalue.Enabled = True
            txtmarkupvalue.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            CmbMarkuptype.Enabled = True
            CmbMarkuptype.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            cmbExporttype.Enabled = False
            cmbExporttype.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            ssPOEntry.Enabled = True
            cmdchangetype.Enabled = True
            txtAmendmentNo.Enabled = False
            txtAmendmentNo.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            cmdHelp(3).Enabled = False
            If blnNoneditableCreditTerms_onSO Then
                txtCreditTerms.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : txtCreditTerms.Enabled = False : cmdHelp(5).Enabled = False
            Else
                txtCreditTerms.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : txtCreditTerms.Enabled = True : cmdHelp(5).Enabled = True
            End If

            cmdForms.Enabled = True
            ctlPerValue.Enabled = True
            ctlPerValue.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            ctlPerValue.Text = 1
            chkOpenSo.Enabled = True
            'Call ADDRow()
            Call GetDetailsFromAgreementMaster(txtCustomerCode.Text, txtReferenceNo.Text)
            DTDate.Focus()

            Exit Sub
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    Public Sub ADDRowCADM()
        '----------------------------------------------------------------------------
        'Argument       :   Nil
        'Return Value   :   Nil
        'Function       :   Adds a new row in the grid
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim inti As Short

        ssPOEntry.MaxRows = ssPOEntry.MaxRows + 1
        With ssPOEntry
            .Col = 2
            .Focus()
        End With
        With ssPOEntry
            .BlockMode = True
            .Col = 3
            .Col2 = 3
            .Row = 1
            .Row2 = .MaxRows
            .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
            .TypeButtonPicture = My.Resources.resEmpower.ico111.ToBitmap()
            .BlockMode = False
        End With
        With ssPOEntry
            .BlockMode = True
            .Col = 2
            .Col2 = 5
            .Row = 1
            .Lock = True
            .Row2 = .MaxRows
            .BlockMode = False
        End With
        With ssPOEntry
            .BlockMode = True
            .Col = 7
            .Col2 = 8
            .Row = 1
            .Lock = True
            .Row2 = .MaxRows
            .BlockMode = False
        End With
        For inti = 5 To 8
            Call ssPOEntry.SetText(inti, ssPOEntry.MaxRows, 0)
        Next
        For inti = 11 To 12
            Call ssPOEntry.SetText(inti, ssPOEntry.MaxRows, 0)
        Next
        Call ssPOEntry.SetText(19, ssPOEntry.MaxRows, 0)
        Call SSMaxLength()
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
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub

    Private Sub frmMKTTRN0001_Closed(sender As Object, e As EventArgs) Handles Me.Closed

    End Sub
End Class