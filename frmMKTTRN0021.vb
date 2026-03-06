Option Strict Off
Option Explicit On
Friend Class frmMKTTRN0021
	Inherits System.Windows.Forms.Form
	'===================================================================================
	' (c) MothersonSumi INfotech & Design Ltd. All rights reserverd.
	' File Name         :   FRMMKTTRN0021.frm
	' Function          :   Used to select items
	' Created By        :   Nisha
	' Created On        :   15 May, 2001
	' Revision History  :   Nisha Rai
	'21/09/2001 MARKED CHECKED BY BCs changed on version 3
	'03/10/2001 MARKED CHECKED BY BCs  for jobwork invoice changed on version 6
	'09/10/2001 jobwork invoice changed on version 7 for set status to one in case of Schedule of Daily/Monthly
	'07/03/2002 Change done in case of export form "EXPTRN00010.frm"  & Error reported from MATE
	'22/03/2002 CHANGED SIZE OF THE FORM FOR MSSLED
	'26/03/2002 CHANGED TO INCLUDE SEMIFINISHED GOOD IN RAW MATERIAL TYPE.
	'19/04/2002 changed for tariff code
	'08/05/2002 Changes for Scrap invoiceing
	'29/05/02 TO REMOVE THE CHECK FOR ITEM BALANCE qTY IN CASE OF EXPORT INVOICE
	'11/06/2002 Message change No ITem found for Selected invoice type
	'23/07/2002 changed to add Grin Linking in Rejection Invoice
	'CHANGES DONE BY NISHA ON 13/03/2003
	'1.FOR FINAL MERGING & FOR FROM BOX & TO bOX UPDATION WHILE EDITING INVOICE
	'2.For Grin Cancellation flag
	'3.SAMPLE INVOICE TOOL COST COLUMN
	'4.CUNSUMABLES & MISC. SALE IN CASE OF NORMAL RAW MATERIAL INVOICE
	'changed by nisha on 21/03/2003 for financial rollover
	'17/04/2003 by nisha daily schedueles to show in form
	'16/05/2003  for summit issues
	'changes done by nisha on 26/11/2003 for mapl grnRejection case
	' changes made by Pooja on 30/01/04
	' Trading and Fini8shed goods added to Transfer Inputs
	'Changes Done by nisha to synchronise SUNVAC on 16/10/2004
	'revision Date ---14-jan-2005
	'Changed made to select Items in list of Finished goods By Brij for SunVac as item group changed from finished to Semi finished
	'===================================================================================
	'Change by Sandeep On 31-March-2005
	'REJECTION INVOICE TRACKING
	'SHOW THE LIST OF ITEMS AS SELECTED DOCUMENTS IN INV ENTRY
	'===================================================================================
	'Revised By    : By ashutosh on 16-11-2005
	'History       : Issue Id: 16222,Allow items for different tariffs in invoice AND
	'              : Do not allow any item with zero or blank tariff code.
	'===================================================================================
	'Revised By    : By ashutosh on 29-11-2005
	'History       : Issue Id:16338, Allow items w/o tariff code in case of transfer invoice.
	'===================================================================================
	'Revised By    : By ashutosh on 25-01-2006
	'History       : Issue Id:16964, Provision for selecting more than 7 items on Export invoice(Only for MATE Noida)
	'===================================================================================
	'Revised By    : Ashutosh ,issue Id:17355
	'History       : On 13-04-2006 , Tarriff code validation.
	'==================================================================================================
	'Revised By    : Ashutosh ,issue Id:17575
	'History       : On 18-04-2006 , Provision for selecting more than 7 items on
	'              : Invoice according to parameter is sales_parameter.
	'==================================================================================================
	'Revised By         : Ashutosh Verma
	'Revision On        : 06-10-2006
	'Issue ID           : 18702
	'History            : Provision for sales order in Transfer Invoice.(Parametrise)
    '==================================================================================================
    'Revised By        -    Vinod Singh
    'Revision Date     -    09/05/2011
    'Revision History  -    Changes for Multi Unit
    '==================================================================================================
    '=======================================================================================
    'Revised By         : Prashant Rajpal
    'Revision On        : 20-07-2011
    'Issue ID           : 10116365
    'History            : Reorder Call changes
    '=======================================================================================
    'Modified by Deepak on 11-Oct-2011 for MultiUnit Change Management
    '=======================================================================================
    'REVISED BY     :  VINOD SINGH
    'REVISED DATE   :  27 AUG 2013
    'ISSUE ID       :  10378778
    'PURPOSE        :  GLOBAL TOOL CHAGES
    '=======================================================================================
    ' REVISION DATE     : 24 FEB 2014
    ' REVISED BY        : PRASHANT RAJPAL
    ' ISSUE ID          : 10533478
    ' REVISION HISTORY  : WIP FG FUNCTIONALITY CHANGES
    '=======================================================================================
    ' REVISION DATE     : 31 MAR 2015
    ' REVISED BY        : Abhinav Kumar
    ' ISSUE ID          : 10776605
    ' REVISION HISTORY  : eMPro- Service Invoice issue at SMIIEL
    '**********************************************************************************************************************
    'REVISED BY     -  PRASHANT RAJPAL
    'REVISED ON     -  18 SEP 2015
    'PURPOSE        -  10869290 -SERVICE INVOICE 
    '-------------------------------------------------------------------------------------------------------------------------------------------------------
    'REVISED BY     -  PRASHANT RAJPAL
    'REVISED ON     -  06 oct-07-oct-2015 2015
    'PURPOSE        -  10912868  -DS -MAE INVOICE ISSUE
    '-------------------------------------------------------------------------------------------------------------------------------------------------------
    'REVISED BY     -  PRASHANT RAJPAL
    'REVISED ON     -  14 oct 2016
    'PURPOSE        -  10869291 INTER DIVISION INVOICE

    'REVISED BY     -  ASHISH SHARMA
    'REVISED ON     -  06 SEP 2017
    'PURPOSE        -  101254587 - Global Tool Master and Tool Master Enhancement Phase-II
    '-------------------------------------------------------------------------------------------------------------------------------------------------------


    Dim mCtlHdrItemCode As System.Windows.Forms.ColumnHeader
	Dim mCtlHdrDrawingNo As System.Windows.Forms.ColumnHeader
	Dim mCtlHdrDescription As System.Windows.Forms.ColumnHeader
	Dim intCheckCounter As Short
	Dim mListItemUserId As System.Windows.Forms.ListViewItem
	Dim mstrInvType As String
	Dim mstrInvSubType As String
	Dim mstrItemText As String
	Dim blnExpinv As Boolean
    Dim intIteminSp As Short
    Public IsTradingInVoice As Boolean = False 'AMIT RANA
    Dim mstraddmode As Boolean = False
    Dim mblnSTILLAGE As Boolean = False
    Dim mblnSTILLAGEFunctionality As Boolean = False
    Dim mblnServiceInvoicemate As Boolean = False

    Public Property ISSTILLAGE() As Boolean
        Get
            ISSTILLAGE = mblnSTILLAGE
        End Get
        Set(ByVal Value As Boolean)
            mblnSTILLAGE = Value
        End Set
    End Property
    Public Property STILLAGEfunctinality() As Boolean
        Get
            STILLAGEfunctinality = mblnSTILLAGEFunctionality
        End Get
        Set(ByVal Value As Boolean)
            mblnSTILLAGEFunctionality = Value
        End Set
    End Property
    Public Property serviceinvoicemate() As Boolean
        '10869290
        Get
            serviceinvoicemate = mblnServiceInvoicemate
        End Get
        Set(ByVal Value As Boolean)
            mblnServiceInvoicemate = Value
        End Set
    End Property

    Private Sub CmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdCancel.Click
        On Error GoTo ErrHandler
        Me.Close()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
		'Code Modified By   -   Nitin Sood
		'No of Items Selected in Challan can be Till 7
		'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
		'===================================================================================
		'Revised By    : By ashutosh on 25-01-2006
		'History       : Issue Id:16964, Provision for selecting more than 7 items on Export invoice(Only for MATE Noida)
		'===================================================================================
		'Revised By   : Ashutosh ,issue Id:17575
		'History      : On 18-04-2006 , Provision for selecting more than 7 items on
		'             : Invoice according to parameter is sales_parameter.
		'====================================================================================
		
        On Error GoTo ErrHandler

        'AMIT RANA
        If IsTradingInVoice = True Then
            If Me.lvwItemCode.CheckedItems.Count > 1 Then
                MsgBox("Only One Item Is Allowed To Be Selected In The List.", MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            End If
        End If
        'AMIT RANA

		mstrItemText = "" : intCheckCounter = intIteminSp
        Dim intSubItem As Short
        Dim gobjDB As ClsResultSetDB
        ' Change By Deepak on 11-Oct-2011 for support Change Management---------
        Dim blnMoreThan7ItemInInvoice As Boolean
        gobjDB = New ClsResultSetDB
        blnMoreThan7ItemInInvoice = False

        gobjDB.GetResult("Select MoreThan7ItemInInvoice from sales_parameter where unit_code='" & gstrUNITID & "'")
        If gobjDB.GetValue("MoreThan7ItemInInvoice") = True Then
            blnMoreThan7ItemInInvoice = True
        Else
            blnMoreThan7ItemInInvoice = False
        End If
        '--------------------------------------------------------------------------
        For intSubItem = 0 To lvwItemCode.Items.Count - 1
            If Me.lvwItemCode.Items.Item(intSubItem).Checked = True Then
                intCheckCounter = intCheckCounter + 1
                If blnExpinv = False Then
                    ' Change By Deepak on 11-Oct-2011 for support Change Management---------
                    If blnMoreThan7ItemInInvoice = True Then
                        gobjDB = New ClsResultSetDB
                        gobjDB.GetResult("select MaximumItemsInInvoices from sales_parameter where unit_code='" & gstrUNITID & "'")
                        'If intCheckCounter > 7 Then
                        If intCheckCounter > gobjDB.GetValue("MaximumItemsInInvoices") Then
                            'Call ConfirmWindow(10415, BUTTON_OK) 
                            ' MsgBox("No. Of Items Selected Should be Less than 7", MsgBoxStyle.Information, "empower")
                            MsgBox("No. Of Items Selected Should not be greater than " & gobjDB.GetValue("MaximumItemsInInvoices") & "", MsgBoxStyle.Information, "empower")
                            mstrItemText = ""
                            Exit Sub
                        End If
                        End If
                    '---------------------------------------------------------------
                Else
                    gobjDB = New ClsResultSetDB
                    gobjDB.GetResult("Select EOU_Flag,company_code from Company_Mst where unit_code='" & gstrUNITID & "'")
                    If gobjDB.GetValue("EOU_Flag") = False Then
                        gobjDB.ResultSetClose()
                        gobjDB = New ClsResultSetDB
                        gobjDB.GetResult("Select MoreThan7ItemInInvoice from sales_parameter where unit_code='" & gstrUNITID & "'")
                        If gobjDB.GetValue("MoreThan7ItemInInvoice") = False Then
                            gobjDB.ResultSetClose()
                            gobjDB = Nothing
                            If intCheckCounter > 7 Then
                                MsgBox("No. Of Items Selected Should be Less than 7", MsgBoxStyle.Information, "empower")
                                mstrItemText = ""
                                Exit Sub
                            End If
                        End If
                    End If
                End If
                mstrItemText = mstrItemText & "'" & Trim(Me.lvwItemCode.Items.Item(intSubItem).SubItems(1).Text) & "',"
            End If
        Next intSubItem
        If mstraddmode = False Then
            If lvwItemCode.Items.Count <> intCheckCounter Then
                MsgBox("Kindly Select All Items ", MsgBoxStyle.Information, ResolveResString(100))
                mstrItemText = ""
                Exit Sub
            End If
        End If

        If Len(mstrItemText) = 0 Then
            Call ConfirmWindow(10418, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
            Me.lvwItemCode.Focus()
            Exit Sub
        End If
        Me.Close()
            Exit Sub

ErrHandler:     'The Error Handling Code Starts here
                Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub frmMKTTRN0021_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo ErrHandler
        Call AddColumnsInListView()
        'Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(mdifrmMain.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2)
        'Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(mdifrmMain.Width) - VB6.PixelsToTwipsX(frmModules.Width)) / 2.3)
        optPartNo.Checked = True
        lvwItemCode.FullRowSelect = True
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub AddColumnsInListView()
        '***********************************
        'To add Columns Headers in the ListView in the form load
        ''Changes done by Ashutosh on 01 May 2007 ,issue Id:19911
        '***********************************
        On Error GoTo ErrHandler
        With Me.lvwItemCode
            mCtlHdrItemCode = .Columns.Add("")
            If (UCase(mstrInvType) = ("TRANSFER INVOICE") Or UCase(mstrInvType) = "INTER-DIVISION") And UCase(mstrInvSubType) = ("FINISHED GOODS") Then
                mCtlHdrItemCode.Text = "Drawing No."
            Else
                mCtlHdrItemCode.Text = "Item Code"
            End If
            If InvoiceForMTL() = False Then
                mCtlHdrItemCode.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(lvwItemCode.Width) / 4)
            Else
                mCtlHdrItemCode.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(lvwItemCode.Width) / 3)
            End If
            mCtlHdrDrawingNo = .Columns.Add("")
            If (UCase(mstrInvType) = ("TRANSFER INVOICE") Or UCase(mstrInvType) = "INTER-DIVISION") And UCase(mstrInvSubType) = ("FINISHED GOODS") Then
                mCtlHdrDrawingNo.Text = "Item Code"
            Else
                mCtlHdrDrawingNo.Text = "Drawing No."
            End If
            If InvoiceForMTL() = False Then
                mCtlHdrDrawingNo.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(lvwItemCode.Width) / 4)
            Else
                mCtlHdrDrawingNo.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(lvwItemCode.Width) / 3)
            End If
            mCtlHdrDescription = .Columns.Add("")
            mCtlHdrDescription.Text = "Description"
            If InvoiceForMTL() = False Then
                mCtlHdrDescription.Width = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(lvwItemCode.Width) / 4))
            Else
                mCtlHdrDescription.Width = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(lvwItemCode.Width) / 3))
            End If
            mCtlHdrDescription = .Columns.Add("")
            If gblnGSTUnit = True Then
                mCtlHdrDescription.Text = "HSN/SAC CODE "
            Else
                mCtlHdrDescription.Text = "Tariff Code"
            End If

            If InvoiceForMTL() = False Then
                mCtlHdrDescription.Width = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(lvwItemCode.Width) / 4) - 100)
            Else
                mCtlHdrDescription.Width = 0
            End If
        End With
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Public Function SelectDataFromCustOrd_Dtl(ByRef pstrCustno As String, ByRef pstrRefNo As String, ByRef pstrAmmNo As String, ByRef pstrSubType As String, ByRef pstrInvType As String, ByRef pstrstockLocation As String, Optional ByRef pstrCondition As String = "", Optional ByRef intAlreadyItem As Short = 0) As String
        '***********************************
        'To Get Data From Cust_Ord_Dtl
        '***********************************
        On Error GoTo ErrHandler
        Dim strSelectSql As String 'Declared To Make Select Query
        Dim Validyrmon As String
        Dim effectyrmon As String
        Dim validMon As String
        Dim effectMon As String
        Dim rsCustOrdDtl As ClsResultSetDB
        Dim rsCustOrdHdr As ClsResultSetDB
        Dim intRecordCount As Short 'To Hold Record Count
        Dim intCount As Short
        Dim strDate As String
        'for item selection more then one 4 in case of Export invoice
        intIteminSp = intAlreadyItem
        If pstrInvType = "EXPORT INVOICE" Then
            blnExpinv = True
        Else
            blnExpinv = False
        End If
        strDate = VB6.Format(GetServerDate(), gstrDateFormat)
        Me.lvwItemCode.Items.Clear() 'initially clear all items in the listview
        strSelectSql = "Select effectMon=convert(char(2),month(effect_date)),effectYr=convert(char(4),Year(effect_date)),"
        strSelectSql = strSelectSql & " validMon=convert(char(2),month(Valid_date)),validYr=convert(char(4),year(Valid_date))"
        strSelectSql = strSelectSql & " from Cust_Ord_hdr where unit_code='" & gstrUNITID & "' and"
        strSelectSql = strSelectSql & " Account_Code='" & Trim(pstrCustno) & "' and Cust_Ref='" & Trim(pstrRefNo) & "'"
        strSelectSql = strSelectSql & " and Amendment_No='" & Trim(pstrAmmNo) & "' and Active_Flag = 'A'"
        rsCustOrdHdr = New ClsResultSetDB
        rsCustOrdHdr.GetResult(strSelectSql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        If rsCustOrdHdr.GetNoRows > 0 Then
            validMon = CStr(Month(GetServerDate))
            If CDbl(validMon) < 10 Then
                validMon = "0" & validMon
            End If
            Validyrmon = Year(GetServerDate) & validMon
            effectMon = rsCustOrdHdr.GetValue("EffectMon")
            If CDbl(effectMon) < 10 Then
                effectMon = "0" & effectMon
            End If
            effectyrmon = rsCustOrdHdr.GetValue("effectYr") & effectMon
            mstrInvType = pstrInvType : mstrInvSubType = pstrSubType
            mstraddmode = True
            Select Case UCase(pstrInvType)
                Case "NORMAL INVOICE", "EXPORT INVOICE", "SERVICE INVOICE"
                    Select Case UCase(pstrSubType)
                        Case "FINISHED GOODS"
                            strSelectSql = makeSelectSql(pstrCustno, pstrRefNo, pstrAmmNo, effectyrmon, Validyrmon, pstrstockLocation, strDate, "'F','S'", pstrCondition)
                        Case "COMPONENTS"
                            strSelectSql = MakeSelectSubQuery(pstrCustno, pstrRefNo, pstrAmmNo, pstrstockLocation, "'C'", pstrCondition)
                        Case "RAW MATERIAL"
                            strSelectSql = MakeSelectSubQuery(pstrCustno, pstrRefNo, pstrAmmNo, pstrstockLocation, "'R','S','B','M'", pstrCondition)
                        Case "ASSETS"
                            'CHANGED BY VINOD
                            'strSelectSql = MakeSelectSubQuery(pstrCustno, pstrRefNo, pstrAmmNo, pstrstockLocation, "'P'", pstrCondition)
                            strSelectSql = MakeSelectSubQuery_AssetInvoice(pstrCustno, pstrRefNo, pstrAmmNo, pstrstockLocation, "'P'", pstrCondition)
                        Case "TRADING GOODS"
                            strSelectSql = makeSelectSql(pstrCustno, pstrRefNo, pstrAmmNo, effectyrmon, Validyrmon, pstrstockLocation, strDate, "'F','S'", pstrCondition)
                        Case "TOOLS & DIES"
                            'CHANGED BY VINOD 
                            'strSelectSql = MakeSelectSubQuery(pstrCustno, pstrRefNo, pstrAmmNo, pstrstockLocation, "'P','A'", pstrCondition)
                            strSelectSql = MakeSelectSubQuery_ToolInvoice(pstrCustno, pstrRefNo, pstrAmmNo, pstrstockLocation, "'P','A'", pstrCondition)
                        Case "EXPORTS"
                            strSelectSql = makeSelectSql(pstrCustno, pstrRefNo, pstrAmmNo, effectyrmon, Validyrmon, pstrstockLocation, strDate, "'F','S'", pstrCondition)
                            'Changes done against issue ID 10776605 by Abhinav
                        Case "SAMPLE"
                            strSelectSql = makeSelectSql(pstrCustno, pstrRefNo, pstrAmmNo, effectyrmon, Validyrmon, pstrstockLocation, strDate, "'R','C','A','M'", pstrCondition)
                        Case "SERVICE"
                            '10869290
                            If serviceinvoicemate = True Then
                                strSelectSql = MakeSelectSubQuery(pstrCustno, pstrRefNo, pstrAmmNo, pstrstockLocation, "'M'", pstrCondition)
                            Else
                                strSelectSql = MakeSelectSubQuery(pstrCustno, pstrRefNo, pstrAmmNo, pstrstockLocation, "'F','S','P'", pstrCondition)
                            End If
                    End Select
                Case "JOBWORK INVOICE"
                    strSelectSql = makeSelectSql(pstrCustno, pstrRefNo, pstrAmmNo, effectyrmon, Validyrmon, pstrstockLocation, strDate, "'F'", pstrCondition)
                Case "TRANSFER INVOICE", "INTER-DIVISION"
                    Select Case UCase(pstrSubType)
                        Case "INPUTS"
                            strSelectSql = makeSelectSql(pstrCustno, pstrRefNo, pstrAmmNo, effectyrmon, Validyrmon, pstrstockLocation, strDate, "'R','C','M','N','S','B','A','F','T'", pstrCondition)
                        Case "FINISHED GOODS"
                            strSelectSql = makeSelectSql(pstrCustno, pstrRefNo, pstrAmmNo, effectyrmon, Validyrmon, pstrstockLocation, strDate, "'F','S'", pstrCondition)
                        Case "ASSETS"
                            'CHANGED BY VINOD
                            'strSelectSql = makeSelectSql(pstrCustno, pstrRefNo, pstrAmmNo, effectyrmon, Validyrmon, pstrstockLocation, strDate, "'P'", pstrCondition)
                            strSelectSql = makeSelectSql_AssetInvoice(pstrCustno, pstrRefNo, pstrAmmNo, effectyrmon, Validyrmon, pstrstockLocation, strDate, "'P'", pstrCondition)
                    End Select
            End Select
        Else
            rsCustOrdHdr.ResultSetClose()
            rsCustOrdHdr = Nothing
            strSelectSql = "Select effect_date,"
            strSelectSql = strSelectSql & " Valid_date "
            strSelectSql = strSelectSql & " from Cust_Ord_hdr where unit_code='" & gstrUNITID & "' and "
            strSelectSql = strSelectSql & " Account_Code='" & Trim(pstrCustno) & "' and Cust_Ref='" & Trim(pstrRefNo) & "'"
            strSelectSql = strSelectSql & " and Amendment_No='" & Trim(pstrAmmNo) & "' and Active_flag ='A'"
            rsCustOrdHdr = New ClsResultSetDB
            rsCustOrdHdr.GetResult(strSelectSql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            If rsCustOrdHdr.GetNoRows > 0 Then
                Validyrmon = rsCustOrdHdr.GetValue("valid_date")
                effectyrmon = rsCustOrdHdr.GetValue("Effect_date")
            End If
            rsCustOrdHdr.ResultSetClose()
            rsCustOrdHdr = Nothing
            Select Case pstrSubType
                Case "COMPONENTS"
                    strSelectSql = makeSelectSql(pstrCustno, pstrRefNo, pstrAmmNo, effectyrmon, Validyrmon, pstrstockLocation, strDate, "'C'", pstrCondition)
                Case "TRADING GOODS"
                    strSelectSql = makeSelectSql(pstrCustno, pstrRefNo, pstrAmmNo, effectyrmon, Validyrmon, pstrstockLocation, strDate, "'T'", pstrCondition)
                Case "ASSETS"
                    'strSelectSql = makeSelectSql(pstrCustno, pstrRefNo, pstrAmmNo, effectyrmon, Validyrmon, pstrstockLocation, strDate, "'P'", pstrCondition)
                    strSelectSql = makeSelectSql_AssetInvoice(pstrCustno, pstrRefNo, pstrAmmNo, effectyrmon, Validyrmon, pstrstockLocation, strDate, "'P'", pstrCondition)
                Case "TOOLS & DIES"
                    strSelectSql = makeSelectSql(pstrCustno, pstrRefNo, pstrAmmNo, effectyrmon, Validyrmon, pstrstockLocation, strDate, "'A','P'", pstrCondition)
                Case "RAW MATERIAL"
                    strSelectSql = makeSelectSql(pstrCustno, pstrRefNo, pstrAmmNo, effectyrmon, Validyrmon, pstrstockLocation, strDate, "'R','S','B','M'", pstrCondition)
                Case "SCRAP"
                    strSelectSql = makeSelectSql(pstrCustno, pstrRefNo, pstrAmmNo, effectyrmon, Validyrmon, pstrstockLocation, strDate, "'R','C'", pstrCondition)
            End Select
        End If
        rsCustOrdDtl = New ClsResultSetDB
        If strSelectSql = "" Then Exit Function
        rsCustOrdDtl.GetResult(strSelectSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        intRecordCount = rsCustOrdDtl.GetNoRows 'assign record count to integer variable
        If intRecordCount > 0 Then '          'if record found
            rsCustOrdDtl.MoveFirst() 'move to first record
            For intCount = 0 To intRecordCount - 1
                mListItemUserId = Me.lvwItemCode.Items.Add(rsCustOrdDtl.GetValue("Item_code"))
                If mListItemUserId.SubItems.Count > 1 Then
                    mListItemUserId.SubItems(1).Text = rsCustOrdDtl.GetValue("Cust_Drgno")
                Else
                    mListItemUserId.SubItems.Insert(1, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsCustOrdDtl.GetValue("Cust_Drgno")))
                End If
                If mListItemUserId.SubItems.Count > 2 Then
                    mListItemUserId.SubItems(2).Text = rsCustOrdDtl.GetValue("Cust_Drg_Desc")
                Else
                    mListItemUserId.SubItems.Insert(2, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsCustOrdDtl.GetValue("Cust_Drg_Desc")))
                End If
                If mListItemUserId.SubItems.Count > 3 Then
                    If gblnGSTUnit = True Then
                        mListItemUserId.SubItems(3).Text = rsCustOrdDtl.GetValue("HSN_SAC_CODE")
                    Else
                        mListItemUserId.SubItems(3).Text = rsCustOrdDtl.GetValue("Tariff_Code")
                    End If

                Else
                    If gblnGSTUnit = True Then
                        mListItemUserId.SubItems.Insert(3, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsCustOrdDtl.GetValue("HSN_SAC_CODE")))
                    Else
                        mListItemUserId.SubItems.Insert(3, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsCustOrdDtl.GetValue("Tariff_Code")))
                    End If

                End If
                rsCustOrdDtl.MoveNext() 'move to next record
            Next intCount
            rsCustOrdDtl.ResultSetClose()
            rsCustOrdDtl = Nothing
        Else
            MsgBox("No Items for selected Invoice in Sales Order.Please Check Following :" & vbCrLf & "1. Item in SO are Active and Not on Hold." & vbCrLf & "2. Check Balance of Items for location " & pstrstockLocation & "." & vbCrLf & "3. Check Marketing Schedule in Case of Finished\Trading Goods in SO.", MsgBoxStyle.Information, "empower")
            Exit Function
        End If
        Me.ShowDialog()
        SelectDataFromCustOrd_Dtl = mstrItemText
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Sub frmMKTTRN0021_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
    End Sub
    Private Sub lvwItemCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles lvwItemCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                CmdOk.Focus()
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
    Public Function SelectDatafromItem_Mst(ByRef pstrInvType As String, ByRef pstrInvSubtype As String, ByRef pstrstockLocation As String, Optional ByRef pstrAccountCode As String = "", Optional ByRef pstrItemNotin As String = "", Optional ByRef intAlreadyItem As Short = 0) As Object
        On Error GoTo ErrHandler
        Dim strItembal As String
        Dim rsItembal As ClsResultSetDB
        Dim intRecordCount As Short 'To Hold Record Count
        Dim intCount As Short
        'for item selection more then one 4 in case of Export invoice
        intIteminSp = intAlreadyItem
        mstraddmode = True
        If pstrInvType = "EXPORT INVOICE" Then
            blnExpinv = True
        Else
            blnExpinv = False
        End If
        mstrInvType = pstrInvType : mstrInvSubType = pstrInvSubtype
        Select Case pstrInvType
            Case "NORMAL INVOICE"
                Select Case pstrInvSubtype
                    Case "TRADING GOODS"
                        If gblnGSTUnit = True Then
                            strItembal = "SELECT Distinct(a.Item_Code),a.description,a.HSN_SAC_CODE FROM Item_Mst a,Itembal_Mst b"
                        Else
                            strItembal = "SELECT Distinct(a.Item_Code),a.description,a.Tariff_code FROM Item_Mst a,Itembal_Mst b"
                        End If

                        strItembal = strItembal & " where  a.unit_code=b.unit_code and a.Item_Code=b.Item_Code and a.Item_Main_Grp ='T'"
                        strItembal = strItembal & " and cur_bal >0 and a.Status ='A' and a.Hold_Flag <> 1 and a.unit_code='" & gstrUNITID & "'"
                        If Len(Trim(pstrItemNotin)) > 0 Then
                            strItembal = strItembal & " and a.Item_code not in (" & pstrItemNotin & ")"
                        End If
                    Case "ASSETS"
                        'QUERY FILER REVISED BY VINOD FOR GLOBAL TOOL CHANGES 
                        If gblnGSTUnit = True Then
                            strItembal = "SELECT Distinct(a.Item_Code),a.description,a.HSN_SAC_CODE FROM Item_Mst a,Itembal_Mst b"
                        Else
                            strItembal = "SELECT Distinct(a.Item_Code),a.description,a.Tariff_code FROM Item_Mst a,Itembal_Mst b"
                        End If

                        strItembal = strItembal & " where a.unit_code=b.unit_code and a.Item_Code=b.Item_Code and a.Item_Main_Grp ='P'"
                        strItembal = strItembal & " and cur_bal >0 and a.Status ='A' and a.Hold_Flag <> 1 and a.unit_code='" & gstrUNITID & "'"
                        If Len(Trim(pstrItemNotin)) > 0 Then
                            strItembal = strItembal & " and a.Item_code not in (" & pstrItemNotin & ")"
                        End If
                        strItembal = strItembal & " AND a.ITEM_CODE NOT IN "
                        strItembal = strItembal & " ( "
                        strItembal = strItembal & " SELECT I.ITEM_CODE FROM ITEM_MST I INNER JOIN GLOBAL_TOOL_MST G "
                        strItembal = strItembal & " ON I.ITEM_CODE = G.GLOBAL_TOOL_CODE_FOR_ITEM_MASTER WHERE I.UNIT_CODE = '" & gstrUNITID & "'"
                        strItembal = strItembal & " AND ITEM_MAIN_GRP ='P' "
                        strItembal = strItembal & " AND G.CATEGORY ='MOULDING' AND I.STATUS='A'"
                        strItembal = strItembal & " ) "
                        strItembal = strItembal & " UNION ALL "
                        If gblnGSTUnit = True Then
                            strItembal = strItembal & " SELECT Distinct(a.Item_Code),a.description,a.HSN_SAC_CODE FROM Item_Mst a,Itembal_Mst b"
                        Else
                            strItembal = strItembal & " SELECT Distinct(a.Item_Code),a.description,a.Tariff_code FROM Item_Mst a,Itembal_Mst b"
                        End If

                        strItembal = strItembal & " where a.unit_code=b.unit_code and a.Item_Code=b.Item_Code and a.Item_Main_Grp ='P'"
                        strItembal = strItembal & " and cur_bal >0 and a.Status ='A' and a.Hold_Flag <> 1 and a.unit_code='" & gstrUNITID & "'"
                        If Len(Trim(pstrItemNotin)) > 0 Then
                            strItembal = strItembal & " and a.Item_code not in (" & pstrItemNotin & ")"
                        End If
                        strItembal = strItembal & " AND A.ITEM_CODE IN "
                        strItembal = strItembal & " ( "
                        strItembal = strItembal & " SELECT I.ITEM_CODE FROM ITEM_MST I INNER JOIN GLOBAL_TOOL_MST G "
                        strItembal = strItembal & " ON I.ITEM_CODE = G.GLOBAL_TOOL_CODE_FOR_ITEM_MASTER "
                        strItembal = strItembal & " WHERE I.UNIT_CODE = '" & gstrUNITID & "' AND ITEM_MAIN_GRP ='P' "
                        strItembal = strItembal & " AND G.CATEGORY ='MOULDING' AND I.STATUS='A' AND G.VENDOR_CODE = '" & pstrAccountCode & "'"
                        strItembal = strItembal & " AND G.TRANSFER_COMPLETED = 0 AND G.TRANSFER_TYPE in ('SALE','CUSTOMER_TOOL_SALE') AND ISNULL(G.INV_NO,0) = 0 "
                        strItembal = strItembal & " ) "

                    Case "TOOLS & DIES"
                        'QUERY FILER REVISED BY VINOD FOR GLOBAL TOOL CHANGES 
                        If gblnGSTUnit = True Then
                            strItembal = "SELECT Distinct(a.Item_Code),a.description,a.HSN_SAC_CODE FROM Item_Mst a,Itembal_Mst b"
                        Else
                            strItembal = "SELECT Distinct(a.Item_Code),a.description,a.Tariff_code FROM Item_Mst a,Itembal_Mst b"
                        End If

                        strItembal = strItembal & " where a.unit_code=b.unit_code and a.Item_Code=b.Item_Code and a.Item_Main_Grp in('P','A')"
                        strItembal = strItembal & " and cur_bal >0 and a.Status ='A' and a.Hold_Flag <> 1 and a.unit_code='" & gstrUNITID & "'"
                        If Len(Trim(pstrItemNotin)) > 0 Then
                            strItembal = strItembal & " and a.Item_code not in (" & pstrItemNotin & ")"
                        End If
                        strItembal = strItembal & " AND a.ITEM_CODE NOT IN "
                        strItembal = strItembal & " ( "
                        strItembal = strItembal & " SELECT I.ITEM_CODE FROM ITEM_MST I INNER JOIN GLOBAL_TOOL_MST G "
                        strItembal = strItembal & " ON I.ITEM_CODE = G.GLOBAL_TOOL_CODE_FOR_ITEM_MASTER WHERE I.UNIT_CODE = '" & gstrUNITID & "'"
                        strItembal = strItembal & " AND ITEM_MAIN_GRP ='P' "
                        strItembal = strItembal & " AND G.CATEGORY ='MOULDING' AND I.STATUS='A'"
                        strItembal = strItembal & " ) "
                    Case "RAW MATERIAL"
                        If gblnGSTUnit = True Then
                            strItembal = "SELECT Distinct(a.Item_Code),a.description,a.HSN_SAC_CODE FROM Item_Mst a,Itembal_Mst b"
                        Else
                            strItembal = "SELECT Distinct(a.Item_Code),a.description,a.Tariff_code FROM Item_Mst a,Itembal_Mst b"
                        End If

                        strItembal = strItembal & " where a.unit_code=b.unit_code and a.Item_Code=b.Item_Code and a.Item_Main_Grp IN('C','R','B','M')"
                        strItembal = strItembal & " and cur_bal >0 and a.Status ='A' and a.Hold_Flag <> 1 and a.unit_code='" & gstrUNITID & "'"
                        If Len(Trim(pstrItemNotin)) > 0 Then
                            strItembal = strItembal & " and a.Item_code not in (" & pstrItemNotin & ")"
                        End If
                        '101254587
                        strItembal = strItembal & CreateSubQueryForGlobalToolItemCheck("'C','R','B','M'", pstrAccountCode)
                        'Changes for scrap invoiceing
                    Case "SCRAP"
                        If gblnGSTUnit = True Then
                            strItembal = "SELECT Distinct(a.Item_Code),a.description,a.HSN_SAC_CODE FROM Item_Mst a,Itembal_Mst b"
                        Else
                            strItembal = "SELECT Distinct(a.Item_Code),a.description,a.Tariff_code FROM Item_Mst a,Itembal_Mst b"
                        End If

                        strItembal = strItembal & " where a.unit_code=b.unit_code and a.Item_Code=b.Item_Code and a.Item_Code in (Select Item_Code  from ItemBal_Mst Where unit_code='" & gstrUNITID & "' and Location_Code ='" & pstrstockLocation & "' and cur_Bal > 0)"
                        strItembal = strItembal & " and a.Status ='A' and a.Hold_Flag <> 1 and a.unit_code='" & gstrUNITID & "'"
                        If Len(Trim(pstrItemNotin)) > 0 Then
                            strItembal = strItembal & " and a.Item_code not in (" & pstrItemNotin & ")"
                        End If
                    Case "COMPONENTS"
                        If gblnGSTUnit = True Then
                            strItembal = "SELECT Distinct(a.Item_Code),a.description,a.HSN_SAC_CODE FROM Item_Mst a,Itembal_Mst b"
                        Else
                            strItembal = "SELECT Distinct(a.Item_Code),a.description,a.Tariff_code FROM Item_Mst a,Itembal_Mst b"
                        End If

                        strItembal = strItembal & " where a.unit_code=b.unit_code and a.Item_Code=b.Item_Code and a.Item_Main_Grp ='C'"
                        strItembal = strItembal & " and cur_bal >0 and a.Status ='A' and a.Hold_Flag <> 1 and a.unit_code='" & gstrUNITID & "'"
                        If Len(Trim(pstrItemNotin)) > 0 Then
                            strItembal = strItembal & " and a.Item_code not in (" & pstrItemNotin & ")"
                        End If
                End Select
            Case "SAMPLE INVOICE"
                Select Case pstrInvSubtype
                    Case "FINISHED GOODS"
                        If gblnGSTUnit = True Then
                            strItembal = "SELECT Distinct(a.Item_Code),a.description,a.HSN_SAC_CODE FROM Item_Mst a,Itembal_Mst b"
                        Else
                            strItembal = "SELECT Distinct(a.Item_Code),a.description,a.Tariff_code FROM Item_Mst a,Itembal_Mst b"
                        End If

                        strItembal = strItembal & " where a.unit_code=b.unit_code and  a.Item_Code=b.Item_Code and a.Item_Main_Grp in ('F','S')"
                        strItembal = strItembal & " and cur_bal >0 and a.Status ='A' and a.Hold_Flag <> 1"
                        strItembal = strItembal & " and b.Location_Code = '" & pstrstockLocation & "' and a.unit_code='" & gstrUNITID & "'"
                        If Len(Trim(pstrItemNotin)) > 0 Then
                            strItembal = strItembal & " and a.Item_code not in (" & pstrItemNotin & ")"
                        End If
                    Case "RAW MATERIAL"
                        If gblnGSTUnit = True Then
                            strItembal = "SELECT Distinct(a.Item_Code),a.description,a.HSN_SAC_CODE FROM Item_Mst a,Itembal_Mst b"
                        Else
                            strItembal = "SELECT Distinct(a.Item_Code),a.description,a.Tariff_code FROM Item_Mst a,Itembal_Mst b"
                        End If

                        strItembal = strItembal & " where a.unit_code = b.unit_code and a.Item_Code=b.Item_Code and a.Item_Main_Grp in ('R','S')"
                        strItembal = strItembal & " and cur_bal >0 and a.Status ='A' and a.Hold_Flag <> 1 and a.unit_code='" & gstrUNITID & "'"
                        strItembal = strItembal & " and b.Location_Code = '" & pstrstockLocation & "'"
                        If Len(Trim(pstrItemNotin)) > 0 Then
                            strItembal = strItembal & " and a.Item_code not in (" & pstrItemNotin & ")"
                        End If
                    Case "RAW MATERIAL & FINISHED GOODS"
                        If gblnGSTUnit = True Then
                            strItembal = "SELECT Distinct(a.Item_Code),a.description,a.HSN_SAC_CODE FROM Item_Mst a,Itembal_Mst b"
                        Else
                            strItembal = "SELECT Distinct(a.Item_Code),a.description,a.Tariff_code FROM Item_Mst a,Itembal_Mst b"
                        End If
                        strItembal = strItembal & " where a.unit_code=b.unit_code and a.Item_Code=b.Item_Code and a.Item_Main_Grp in ('F','S','R')"
                        strItembal = strItembal & " and cur_bal >0 and a.Status ='A' and a.Hold_Flag <> 1 and a.unit_code='" & gstrUNITID & "'"
                        strItembal = strItembal & " and b.Location_Code = '" & pstrstockLocation & "'"
                        If Len(Trim(pstrItemNotin)) > 0 Then
                            strItembal = strItembal & " and a.Item_code not in (" & pstrItemNotin & ")"
                        End If
                    Case "COMPONENTS"
                        If gblnGSTUnit = True Then
                            strItembal = "SELECT Distinct(a.Item_Code),a.description,a.HSN_SAC_CODE FROM Item_Mst a,Itembal_Mst b"
                        Else
                            strItembal = "SELECT Distinct(a.Item_Code),a.description,a.Tariff_code FROM Item_Mst a,Itembal_Mst b"
                        End If

                        strItembal = strItembal & " where a.unit_code=b.unit_code and a.Item_Code=b.Item_Code and a.Item_Main_Grp in ('C','S')"
                        strItembal = strItembal & " and cur_bal >0 and a.Status ='A' and a.Hold_Flag <> 1"
                        strItembal = strItembal & " and b.Location_Code = '" & pstrstockLocation & "' and a.unit_code='" & gstrUNITID & "'"
                        If Len(Trim(pstrItemNotin)) > 0 Then
                            strItembal = strItembal & " and a.Item_code not in (" & pstrItemNotin & ")"
                        End If
                End Select
            Case "TRANSFER INVOICE", "INTER-DIVISION"
                Select Case pstrInvSubtype
                    Case "ASSETS"
                        'QUERY FILER REVISED BY VINOD FOR GLOBAL TOOL CHANGES 
                        If gblnGSTUnit = True Then
                            strItembal = "SELECT Distinct(a.Item_Code),a.description,a.HSN_SAC_CODE FROM Item_Mst a,Itembal_Mst b"
                        Else
                            strItembal = "SELECT Distinct(a.Item_Code),a.description,a.Tariff_code FROM Item_Mst a,Itembal_Mst b"
                        End If

                        strItembal = strItembal & " where a.unit_code=b.unit_code and a.Item_Code=b.Item_Code and a.Item_Main_Grp = 'P' "
                        strItembal = strItembal & " and cur_bal >0 and a.Status ='A' and a.Hold_Flag <> 1"
                        strItembal = strItembal & " and b.Location_Code = '" & pstrstockLocation & "' and a.unit_code='" & gstrUNITID & "'"
                        If Len(Trim(pstrItemNotin)) > 0 Then
                            strItembal = strItembal & " and a.Item_code not in (" & pstrItemNotin & ")"
                        End If
                        strItembal = strItembal & " AND a.ITEM_CODE NOT IN "
                        strItembal = strItembal & " ( "
                        strItembal = strItembal & " SELECT I.ITEM_CODE FROM ITEM_MST I INNER JOIN GLOBAL_TOOL_MST G "
                        strItembal = strItembal & " ON I.ITEM_CODE = G.GLOBAL_TOOL_CODE_FOR_ITEM_MASTER WHERE I.UNIT_CODE = '" & gstrUNITID & "'"
                        strItembal = strItembal & " AND ITEM_MAIN_GRP ='P' "
                        strItembal = strItembal & " AND G.CATEGORY ='MOULDING' AND I.STATUS='A'"
                        strItembal = strItembal & " ) "
                        strItembal = strItembal & " UNION ALL "
                        If gblnGSTUnit = True Then
                            strItembal = strItembal & " SELECT Distinct(a.Item_Code),a.description,a.HSN_SAC_CODE FROM Item_Mst a,Itembal_Mst b"
                        Else
                            strItembal = strItembal & " SELECT Distinct(a.Item_Code),a.description,a.Tariff_code FROM Item_Mst a,Itembal_Mst b"
                        End If

                        strItembal = strItembal & " where a.unit_code=b.unit_code and a.Item_Code=b.Item_Code and a.Item_Main_Grp = 'P' "
                        strItembal = strItembal & " and cur_bal >0 and a.Status ='A' and a.Hold_Flag <> 1"
                        strItembal = strItembal & " and b.Location_Code = '" & pstrstockLocation & "' and a.unit_code='" & gstrUNITID & "'"
                        If Len(Trim(pstrItemNotin)) > 0 Then
                            strItembal = strItembal & " and a.Item_code not in (" & pstrItemNotin & ")"
                        End If
                        strItembal = strItembal & " AND A.ITEM_CODE IN "
                        strItembal = strItembal & " ( "
                        strItembal = strItembal & " SELECT I.ITEM_CODE FROM ITEM_MST I INNER JOIN GLOBAL_TOOL_MST G "
                        strItembal = strItembal & " ON I.ITEM_CODE = G.GLOBAL_TOOL_CODE_FOR_ITEM_MASTER "
                        strItembal = strItembal & " WHERE I.UNIT_CODE = '" & gstrUNITID & "' AND ITEM_MAIN_GRP ='P' "
                        strItembal = strItembal & " AND G.CATEGORY ='MOULDING' AND I.STATUS='A' AND G.VENDOR_CODE = '" & pstrAccountCode & "'"
                        strItembal = strItembal & " AND G.TRANSFER_COMPLETED = 0 AND G.TRANSFER_TYPE in ('SALE','CUSTOMER_TOOL_SALE') AND ISNULL(G.INV_NO,0) = 0 "
                        strItembal = strItembal & " ) "
                    Case "FINISHED GOODS"
                        If gblnGSTUnit = True Then
                            strItembal = "SELECT Distinct a.Item_Code,c.Cust_drgNo,c.Drg_Desc, a.hsn_sac_code FROM Item_Mst a,Itembal_Mst b,CustItem_Mst c "
                        Else
                            strItembal = "SELECT Distinct a.Item_Code,c.Cust_drgNo,c.Drg_Desc, a.Tariff_code FROM Item_Mst a,Itembal_Mst b,CustItem_Mst c "
                        End If
                        strItembal = strItembal & " where A.UNIT_CODE = B.UNIT_CODE AND B.UNIT_CODE = C.UNIT_CODE and a.Item_Code=b.Item_Code and a.Item_Main_Grp in('F','S') and a.Item_Code = c.ITem_Code"
                        strItembal = strItembal & " and cur_bal >0 and a.Status ='A' and a.Hold_Flag <> 1 and c.Account_code ='" & pstrAccountCode & "'"
                        strItembal = strItembal & " and b.Location_Code = '" & pstrstockLocation & "' and a.unit_code='" & gstrUNITID & "'"
                        If Len(Trim(pstrItemNotin)) > 0 Then
                            strItembal = strItembal & " and a.Item_code not in (" & pstrItemNotin & ")"
                        End If
                    Case "INPUTS"
                        If gblnGSTUnit = True Then
                            strItembal = "SELECT Distinct(a.Item_Code),a.description,a.hsn_sac_code FROM Item_Mst a,Itembal_Mst b"
                        Else
                            strItembal = "SELECT Distinct(a.Item_Code),a.description,a.Tariff_code FROM Item_Mst a,Itembal_Mst b"
                        End If

                        strItembal = strItembal & " where a.unit_code = b.unit_code and a.Item_Code=b.Item_Code and a.Item_Main_Grp in('R','C','M','N','S','B','A','F','T')"
                        strItembal = strItembal & " and cur_bal >0 and a.Status ='A' and a.Hold_Flag <> 1"
                        strItembal = strItembal & " and b.Location_Code = '" & pstrstockLocation & "' and a.unit_code='" & gstrUNITID & "'"
                        If Len(Trim(pstrItemNotin)) > 0 Then
                            strItembal = strItembal & " and a.Item_code not in (" & pstrItemNotin & ")"
                        End If
                        '101254587
                        strItembal = strItembal & CreateSubQueryForGlobalToolItemCheck("'R','C','M','N','S','B','A','F','T'", pstrAccountCode)
                End Select
            Case "REJECTION"
                If gblnGSTUnit = True Then
                    strItembal = "SELECT Distinct(a.Item_Code),a.description,c.hsn_sac_code FROM vend_item a ,Itembal_Mst b,Item_Mst c"
                Else
                    strItembal = "SELECT Distinct(a.Item_Code),a.description,c.Tariff_code FROM vend_item a ,Itembal_Mst b,Item_Mst c"
                End If

                strItembal = strItembal & " where a.unit_code=b.unit_code and b.unit_code=c.unit_code and a.Item_Code=b.Item_Code and a.Item_code = c.Item_code and a.Account_code ='" & pstrAccountCode & "' "
                strItembal = strItembal & " and cur_bal >0 "
                strItembal = strItembal & " and b.Location_Code = '" & pstrstockLocation & "' and a.unit_code='" & gstrUNITID & "'"
                If Len(Trim(pstrItemNotin)) > 0 Then
                    strItembal = strItembal & " and a.Item_code not in (" & pstrItemNotin & ")"
                End If
            Case "SERVICE INVOICE"
                If gblnGSTUnit = True Then
                    strItembal = "SELECT Distinct(a.Item_Code), a.description, a.HSN_SAC_CODE FROM Item_mst a where a.unit_code='" & gstrUnitId & "' and a.Item_Main_Grp='M' and a.Status = 'A' and a.Hold_flag <> 1"
                Else
                    strItembal = "SELECT Distinct(a.Item_Code), a.description, a.Tariff_code FROM Item_mst a where a.unit_code='" & gstrUnitId & "' and a.Item_Main_Grp='M' and a.Status = 'A' and a.Hold_flag <> 1"
                End If

                If Len(Trim(pstrItemNotin)) > 0 Then
                    strItembal = strItembal & " and a.Item_code not in (" & pstrItemNotin & ")"
                End If
                '101254587
                strItembal = strItembal & CreateSubQueryForGlobalToolItemCheck("'M'", pstrAccountCode)
        End Select
        rsItembal = New ClsResultSetDB
        rsItembal.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        intRecordCount = rsItembal.GetNoRows 'assign record count to integer variable
        If intRecordCount > 0 Then '          'if record found
            rsItembal.MoveFirst() 'move to first record
            If ((UCase(pstrInvType) = "TRANSFER INVOICE") Or (UCase(pstrInvType) = "INTER-DIVISION")) And UCase(pstrInvSubtype) = "FINISHED GOODS" Then
                For intCount = 0 To intRecordCount - 1
                    mListItemUserId = Me.lvwItemCode.Items.Add(rsItembal.GetValue("Cust_drgNo"))
                    If mListItemUserId.SubItems.Count > 1 Then
                        mListItemUserId.SubItems(1).Text = rsItembal.GetValue("Item_code")
                    Else
                        mListItemUserId.SubItems.Insert(1, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsItembal.GetValue("Item_code")))
                    End If
                    If mListItemUserId.SubItems.Count > 2 Then
                        mListItemUserId.SubItems(2).Text = rsItembal.GetValue("Drg_Desc")
                    Else
                        mListItemUserId.SubItems.Insert(2, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsItembal.GetValue("Drg_Desc")))
                    End If
                    If mListItemUserId.SubItems.Count > 3 Then
                        If gblnGSTUnit = True Then
                            mListItemUserId.SubItems(3).Text = rsItembal.GetValue("HSN_SAC_CODE")
                        Else
                            mListItemUserId.SubItems(3).Text = rsItembal.GetValue("Tariff_Code")
                        End If
                    Else
                        If gblnGSTUnit = True Then
                            mListItemUserId.SubItems.Insert(3, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsItembal.GetValue("hsn_sac_code")))
                        Else
                            mListItemUserId.SubItems.Insert(3, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsItembal.GetValue("Tariff_Code")))
                        End If

                    End If
                    rsItembal.MoveNext() 'move to next record
                Next intCount
            Else
                For intCount = 0 To intRecordCount - 1
                    mListItemUserId = Me.lvwItemCode.Items.Add(rsItembal.GetValue("Item_code"))
                    If mListItemUserId.SubItems.Count > 1 Then
                        mListItemUserId.SubItems(1).Text = rsItembal.GetValue("Item_code")
                    Else
                        mListItemUserId.SubItems.Insert(1, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsItembal.GetValue("Item_code")))
                    End If
                    If mListItemUserId.SubItems.Count > 2 Then
                        mListItemUserId.SubItems(2).Text = rsItembal.GetValue("Description")
                    Else
                        mListItemUserId.SubItems.Insert(2, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsItembal.GetValue("Description")))
                    End If
                    If mListItemUserId.SubItems.Count > 3 Then
                        If gblnGSTUnit = True Then
                            mListItemUserId.SubItems(3).Text = rsItembal.GetValue("HSN_SAC_CODE")
                        Else
                            mListItemUserId.SubItems(3).Text = rsItembal.GetValue("Tariff_Code")
                        End If

                    Else
                        If gblnGSTUnit = True Then
                            mListItemUserId.SubItems.Insert(3, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsItembal.GetValue("HSN_SAC_CODE")))
                        Else
                            mListItemUserId.SubItems.Insert(3, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsItembal.GetValue("Tariff_Code")))
                        End If

                    End If
                    rsItembal.MoveNext() 'move to next record
                Next intCount
            End If
            rsItembal.ResultSetClose()
            rsItembal = Nothing
        Else
            If ((UCase(pstrInvType) = "TRANSFER INVOICE") Or UCase(pstrInvType) = "INTER-DIVISION") And UCase(pstrInvSubtype) = "FINISHED GOODS" Then
                MsgBox("No items details defined  for above Invoice combination,Please Check Following :" & vbCrLf & "1. Item should be Active and Not on Hold." & vbCrLf & "2. Check Balance of Items for location " & pstrstockLocation & "." & vbCrLf & "3.Item is not defined in Customer ITem Master.", MsgBoxStyle.Information, "empower")
            Else
                MsgBox("No items details defined  for above Invoice combination,Please Check Following :" & vbCrLf & "1. Item should be Active and Not on Hold." & vbCrLf & "2. Check Balance of Items for location " & pstrstockLocation & ".", MsgBoxStyle.Information, "empower")
            End If
            Exit Function
        End If
        Me.ShowDialog()
        SelectDatafromItem_Mst = mstrItemText
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Public Function SelectDatafromsaleDtl(ByRef pstrchallanNo As Object) As Object
        On Error GoTo ErrHandler
        Dim strsaledtl As String
        Dim strInvType As String
        Dim rssaledtl As ClsResultSetDB
        Dim rsInvType As ClsResultSetDB
        Dim intRecordCount As Short 'To Hold Record Count
        Dim intCount As Short
        strInvType = "select a.description,a.Sub_type_Description from saleconf a,saleschallan_dtl b where a.unit_code=b.unit_code and a.Invoice_type =b.Invoice_Type and b.Doc_no = " & Val(pstrchallanNo) & "  and a.unit_code='" & gstrUNITID & "' and datediff(dd,b.Invoice_Date,a.fin_start_date)<=0  and datediff(dd,a.fin_end_date,b.Invoice_Date)<=0"
        rsInvType = New ClsResultSetDB
        rsInvType.GetResult(strInvType, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        mstrInvType = UCase(rsInvType.GetValue("Description"))
        mstrInvSubType = UCase(rsInvType.GetValue("sub_type_Description"))
        If UCase(rsInvType.GetValue("Description")) = "EXPORT INVOICE" Then
            blnExpinv = True
        Else
            blnExpinv = False
        End If
        rsInvType.ResultSetClose()
        rsInvType = Nothing
        strsaledtl = ""
        If gblnGSTUnit = True Then
            strsaledtl = "Select a.Item_Code,a.Cust_ITem_Code,a.Cust_Item_Desc,b.hsn_sac_code from Sales_Dtl a,Item_Mst b where a.unit_code=b.unit_code and a.ITem_code = b.ITem_code and a.unit_code='" & gstrUNITID & "' and Doc_No ="
        Else
            strsaledtl = "Select a.Item_Code,a.Cust_ITem_Code,a.Cust_Item_Desc,b.Tariff_Code from Sales_Dtl a,Item_Mst b where a.unit_code=b.unit_code and a.ITem_code = b.ITem_code and a.unit_code='" & gstrUNITID & "' and Doc_No ="
        End If

        strsaledtl = strsaledtl & pstrchallanNo
        rssaledtl = New ClsResultSetDB
        rssaledtl.GetResult(strsaledtl, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        intRecordCount = rssaledtl.GetNoRows 'assign record count to integer variable
        If intRecordCount > 0 Then '          'if record found
            rssaledtl.MoveFirst() 'move to first record
            mstraddmode = False
            For intCount = 0 To intRecordCount - 1
                mListItemUserId = Me.lvwItemCode.Items.Add(rssaledtl.GetValue("Item_code"))
                If mListItemUserId.SubItems.Count > 1 Then
                    mListItemUserId.SubItems(1).Text = rssaledtl.GetValue("Cust_Item_code")
                Else
                    mListItemUserId.SubItems.Insert(1, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rssaledtl.GetValue("Cust_Item_code")))
                End If
                If mListItemUserId.SubItems.Count > 2 Then
                    mListItemUserId.SubItems(2).Text = rssaledtl.GetValue("Cust_Item_Desc")
                Else
                    mListItemUserId.SubItems.Insert(2, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rssaledtl.GetValue("Cust_Item_Desc")))
                End If
                If mListItemUserId.SubItems.Count > 3 Then
                    If gblnGSTUnit = True Then
                        mListItemUserId.SubItems(3).Text = rssaledtl.GetValue("hsn_sac_code")
                    Else
                        mListItemUserId.SubItems(3).Text = rssaledtl.GetValue("Tariff_code")
                    End If

                Else
                    If gblnGSTUnit = True Then
                        mListItemUserId.SubItems.Insert(3, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rssaledtl.GetValue("hsn_sac_code")))
                    Else
                        mListItemUserId.SubItems.Insert(3, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rssaledtl.GetValue("Tariff_code")))
                    End If

                End If
                rssaledtl.MoveNext() 'move to next record
            Next intCount
            rssaledtl.ResultSetClose()
            rssaledtl = Nothing
        End If
        Me.ShowDialog()
        SelectDatafromsaleDtl = mstrItemText
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Public Function makeSelectSql(ByRef pstrCustno As String, ByRef pstrRefNo As String, ByRef pstrAmmNo As String, ByRef effectyrmon As String, ByRef Validyrmon As String, ByRef pstrstockLocation As String, ByRef strDate As String, ByRef pstrItemin As String, Optional ByRef pstrCondition As String = "") As String
        Dim strSelectSql As String
        Dim strNextWorkDay As String
        Dim RsobjSchedules As New ADODB.Recordset
        Dim blnCalendarDateTrac As Boolean
        If RsobjSchedules.State = ADODB.ObjectStateEnum.adStateOpen Then RsobjSchedules.Close()
        RsobjSchedules.Open("SELECT DSWiseTracking,CalendarDateTrac FROM sales_parameter where unit_code='" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If Not RsobjSchedules.EOF Then
            If IIf(RsobjSchedules.Fields(1).Value, 1, 0) = 1 Then
                blnCalendarDateTrac = True
            Else
                blnCalendarDateTrac = False
            End If
        End If
        RsobjSchedules.Close()
        If blnCalendarDateTrac Then
            strNextWorkDay = GetNextWorkingDay(strDate)
            If strNextWorkDay = "-1" Then
                makeSelectSql = ""
                Exit Function
            End If
        End If
        If gblnGSTUnit = True Then
            strSelectSql = "Select b.Item_Code,c.Cust_DrgNo,c.Cust_Drg_Desc,d.hsn_sac_code from Cust_Ord_hdr a,MonthlyMktSchedule b,Cust_ord_dtl c,Item_Mst d where "
        Else
            strSelectSql = "Select b.Item_Code,c.Cust_DrgNo,c.Cust_Drg_Desc,d.Tariff_Code from Cust_Ord_hdr a,MonthlyMktSchedule b,Cust_ord_dtl c,Item_Mst d where "
        End If

        strSelectSql = strSelectSql & " a.unit_code = b.unit_code and b.unit_code = c.unit_code and c.unit_code = d.unit_code "
        strSelectSql = strSelectSql & " and a.Cust_ref = c.Cust_ref and a.amendment_No = c.amendment_No and a.Account_code=c.account_code And c.Active_Flag ='A' "
        strSelectSql = strSelectSql & " and a.account_code=b.Account_code and c.Cust_drgNo=b.Cust_drgNo and b.ITem_code = d.Item_code and a.Account_Code='" & Trim(pstrCustno) & "' and a.unit_code='" & gstrUNITID & "'"
        strSelectSql = strSelectSql & " and a.Cust_Ref='" & Trim(pstrRefNo) & "' and a.Amendment_No='" & Trim(pstrAmmNo) & "' and b.status = 1 and b.Schedule_flag =1 and b.Year_Month =  " & Validyrmon
        strSelectSql = strSelectSql & " and b.Item_Code in(Select a.Item_code from Item_MSt a,Itembal_mst b where a.unit_code=b.unit_code and a.unit_code='" & gstrUnitId & "' and a.Item_Main_grp in (" & Trim(pstrItemin) & ") and a.Item_code = b.Item_code and b.Location_code ='" & pstrstockLocation & "' and b.Cur_bal >0 and a.hold_flag =0 and a.Status = 'A'"
        '101254587
        strSelectSql = strSelectSql & CreateSubQueryForGlobalToolItemCheck(pstrItemin, pstrCustno)
        If Len(Trim(pstrCondition)) > 0 Then
            strSelectSql = strSelectSql & " and a.Item_code not in(" & pstrCondition & "))"
        Else
            strSelectSql = strSelectSql & ")"
        End If
        strSelectSql = strSelectSql & " UNION "
        If gblnGSTUnit = True Then
            strSelectSql = strSelectSql & " Select b.Item_Code,c.Cust_DrgNo,c.Cust_Drg_Desc,d.hsn_sac_code from Cust_Ord_hdr a,DailyMktSchedule b,Cust_ord_dtl c,ITem_Mst d  where "
        Else
            strSelectSql = strSelectSql & " Select b.Item_Code,c.Cust_DrgNo,c.Cust_Drg_Desc,d.Tariff_Code from Cust_Ord_hdr a,DailyMktSchedule b,Cust_ord_dtl c,ITem_Mst d  where "
        End If

        strSelectSql = strSelectSql & " a.unit_code = b.unit_code and b.unit_code = c.unit_code and c.unit_code = d.unit_code "
        strSelectSql = strSelectSql & " and a.Cust_ref = c.Cust_ref and a.amendment_No = c.amendment_No and a.Account_code=c.account_code"
        strSelectSql = strSelectSql & " and a.account_code=b.Account_code and c.Cust_drgNo=b.Cust_drgNo and b.ITem_code =d.ITem_code and b.status = 1 And c.Active_Flag ='A' and a.Account_Code='" & Trim(pstrCustno) & "' and a.unit_code='" & gstrUNITID & "' "
        strSelectSql = strSelectSql & " and a.Cust_Ref='" & Trim(pstrRefNo) & "' and a.Amendment_No='" & Trim(pstrAmmNo) & "' "
        If blnCalendarDateTrac Then
            If Month(ConvertToDate(strDate)) = Month(ConvertToDate(strNextWorkDay)) Then
                '10912868
                'strSelectSql = strSelectSql & " and  datepart(mm,b.trans_date) = '" & Month(ConvertToDate(strDate)) & "' And  b.Trans_Date <= '" & VB6.Format(strNextWorkDay, "dd/mmm/yyyy") & "'  And DatePart(yyyy, b.Trans_Date) = '" & Year(ConvertToDate(strDate)) & "'"
                strSelectSql = strSelectSql & " And  b.Trans_Date <= '" & VB6.Format(strNextWorkDay, "dd/mmm/yyyy") & "'"
            ElseIf (Month(ConvertToDate(strDate)) + 1) = Month(ConvertToDate(strNextWorkDay)) Then
                'strSelectSql = strSelectSql & " and  datepart(mm,b.trans_date) in ('" & Month(ConvertToDate(strDate)) & "','" & Month(ConvertToDate(strNextWorkDay)) & "') And  b.Trans_Date <= '" & VB6.Format(strNextWorkDay, "dd/mmm/yyyy") & "'  And DatePart(yyyy, b.Trans_Date) = '" & Year(ConvertToDate(strDate)) & "'"
                strSelectSql = strSelectSql & " And  b.Trans_Date <= '" & VB6.Format(strNextWorkDay, "dd/mmm/yyyy") & "'  And DatePart(yyyy, b.Trans_Date) = '" & Year(ConvertToDate(strDate)) & "'"
            ElseIf (Year(ConvertToDate(strDate)) + 1) = Year(ConvertToDate(strNextWorkDay)) Then
                'strSelectSql = strSelectSql & " and  ((datepart(yyyy,b.Trans_Date) = '" & Year(ConvertToDate(strDate)) & "' and datepart(mm,b.Trans_Date)='" & Month(ConvertToDate(strDate)) & "') or ( datepart(yyyy,b.Trans_Date) = '" & Year(CDate(strNextWorkDay)) & "' and datepart(mm,b.Trans_Date)='" & Month(CDate(strNextWorkDay)) & "')) And  b.Trans_Date <= '" & VB6.Format(strNextWorkDay, "dd/mmm/yyyy") & "'"
                strSelectSql = strSelectSql & " and b.Trans_Date <= '" & VB6.Format(strNextWorkDay, "dd/mmm/yyyy") & "'"
            End If
        Else
            strSelectSql = strSelectSql & " and  datepart(mm,b.trans_date) = '" & Month(ConvertToDate(strDate)) & "' And  b.Trans_Date <= '" & getDateForDB(strDate) & "'  And DatePart(yyyy, b.Trans_Date) = '" & Year(ConvertToDate(strDate)) & "'"
        End If
        strSelectSql = strSelectSql & " and b.Item_Code in(Select a.Item_code from Item_MSt a,Itembal_mst b where a.unit_code = b.unit_code and a.Item_Main_grp in (" & Trim(pstrItemin) & ") and a.Item_code = b.Item_code and b.Location_code ='" & pstrstockLocation & "' and b.Cur_bal >0 and a.hold_flag =0 and a.Status = 'A' and a.unit_code='" & gstrUNITID & "'"
        'ISSUE ID 10533478
        If gBlnWIPFGProcess And WIP_FG_Customer(pstrCustno) Then
            strSelectSql = strSelectSql & " AND A.WIP_FG_ITEM = 0"
        End If
        'ISSUE ID 10533478

        'ISSUE ID 10713941 
        If mblnSTILLAGEFunctionality = True Then
            If mblnSTILLAGE = True Then
                strSelectSql = strSelectSql & " AND EXISTS ("
                strSelectSql = strSelectSql & " SELECT TOP 1 1 FROM CUSTITEM_MST CM INNER JOIN PKG_STYLE_MST PM "
                strSelectSql = strSelectSql & " ON CM.UNIT_CODE =PM.UNIT_CODE AND CM.PKG_STYLECODE=PM.PKG_STYLE_C "
                strSelectSql = strSelectSql & "AND CM.UNIT_CODE =C.UNIT_CODE AND CM.ACCOUNT_CODE =C.ACCOUNT_CODE "
                strSelectSql = strSelectSql & "AND CM.ITEM_CODE =C.ITEM_CODE AND CM.CUST_DRGNO =C.CUST_DRGNO "
                strSelectSql = strSelectSql & " AND CM.ACTIVE =1 AND PM.ISSTILLAGETYPE=1 ) "
            Else
                strSelectSql = strSelectSql & " AND EXISTS ("
                strSelectSql = strSelectSql & " SELECT TOP 1 1 FROM CUSTITEM_MST CM INNER JOIN PKG_STYLE_MST PM "
                strSelectSql = strSelectSql & " ON CM.UNIT_CODE =PM.UNIT_CODE AND CM.PKG_STYLECODE=PM.PKG_STYLE_C "
                strSelectSql = strSelectSql & "AND CM.UNIT_CODE =C.UNIT_CODE AND CM.ACCOUNT_CODE =C.ACCOUNT_CODE "
                strSelectSql = strSelectSql & "AND CM.ITEM_CODE =C.ITEM_CODE AND CM.CUST_DRGNO =C.CUST_DRGNO "
                strSelectSql = strSelectSql & " AND CM.ACTIVE =1 AND PM.ISSTILLAGETYPE=0 ) "
            End If
        End If
        '101254587
        strSelectSql = strSelectSql & CreateSubQueryForGlobalToolItemCheck(pstrItemin, pstrCustno)
        'ISSUE ID 10713941 
        If Len(Trim(pstrCondition)) > 0 Then
            strSelectSql = strSelectSql & " and a.Item_code not in( " & pstrCondition & "))"
        Else
            strSelectSql = strSelectSql & ")"
        End If

        makeSelectSql = strSelectSql
    End Function
    Public Function MakeSelectSubQuery(ByRef pstrCustno As String, ByRef pstrRefNo As String, ByRef pstrAmmNo As String, ByRef pstrstockLocation As String, ByRef pstrItemin As String, Optional ByRef pstrItemNotin As String = "") As String
        Dim strSelectSql As String
        If gblnGSTUnit = True Then
            strSelectSql = "Select c.Item_Code,c.Cust_DrgNo,c.Cust_Drg_Desc,d.HSN_SAC_CODE from Cust_Ord_hdr a,Cust_ord_dtl c,Item_Mst d where "
        Else
            strSelectSql = "Select c.Item_Code,c.Cust_DrgNo,c.Cust_Drg_Desc,d.Tariff_Code from Cust_Ord_hdr a,Cust_ord_dtl c,Item_Mst d where "
        End If

        strSelectSql = strSelectSql & " a.unit_code = c.unit_code and c.unit_code = d.unit_code "
        strSelectSql = strSelectSql & " and a.Cust_ref = c.Cust_ref and a.amendment_No = c.amendment_No and a.Account_code=c.account_code and a.unit_code='" & gstrUNITID & "' "
        strSelectSql = strSelectSql & " and  c.Item_code = d.Item_code and a.Account_Code='" & Trim(pstrCustno) & "' and a.Cust_Ref='" & Trim(pstrRefNo)
        strSelectSql = strSelectSql & "' and a.Amendment_No='" & Trim(pstrAmmNo) & "' And c.Active_Flag = 'A' "

        If mblnServiceInvoicemate = False Then
            strSelectSql = strSelectSql & " and c.Item_Code in (Select a.Item_code from Item_MSt a,Itembal_mst b where a.unit_code=b.unit_code and a.Item_Main_grp in (" & pstrItemin & ") and a.Item_code = b.Item_code and b.Location_code ='" & pstrstockLocation & "' and b.Cur_bal >0 and a.hold_flag =0 and a.Status = 'A' and a.unit_code='" & gstrUnitId & "'"
            '101254587
            strSelectSql = strSelectSql & CreateSubQueryForGlobalToolItemCheck(pstrItemin, pstrCustno)
            If Len(Trim(pstrItemNotin)) > 0 Then
                strSelectSql = strSelectSql & " and a.Item_code not in ( " & pstrItemNotin & "))"
            Else
                strSelectSql = strSelectSql & ")"
            End If
        Else
            strSelectSql = strSelectSql & " and c.Item_Code in (Select a.Item_code from Item_MSt a WHERE a.Item_Main_grp in (" & pstrItemin & ")  AND a.hold_flag =0 and a.Status = 'A' and a.unit_code='" & gstrUnitId & "'"
            '101254587
            strSelectSql = strSelectSql & CreateSubQueryForGlobalToolItemCheck(pstrItemin, pstrCustno)
            If Len(Trim(pstrItemNotin)) > 0 Then
                strSelectSql = strSelectSql & " and a.Item_code not in ( " & pstrItemNotin & "))"
            Else
                strSelectSql = strSelectSql & ")"
            End If
        End If

        MakeSelectSubQuery = strSelectSql
    End Function
    Private Sub SearchItem()
        '---------------------------------------------------------------------
        'Created By     -   Shruti Khanna\(Name Changed - Nitin Sood)
        '---------------------------------------------------------------------
        Dim itmFound As System.Windows.Forms.ListViewItem ' FoundItem variable.
        On Error GoTo ErrHandler
        If optDescription.Checked = True Then
            itmFound = SearchText((txtsearch.Text), optDescription, lvwItemCode, "2")
        Else
            itmFound = SearchText((txtsearch.Text), optPartNo, lvwItemCode)
        End If
        If itmFound Is Nothing Then ' If no match,
            Exit Sub
        Else
            itmFound.EnsureVisible() ' Scroll ListView to show found ListItem.
            itmFound.Selected = True ' Select the ListItem.
            ' Return focus to the control to see selection.
            lvwItemCode.Enabled = True
            If Len(txtsearch.Text) > 0 Then itmFound.Font = VB6.FontChangeBold(itmFound.Font, True)
        End If
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub optDescription_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optDescription.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo ErrHandler
            With lvwItemCode
                .Sort()
                ListViewColumnSorter.SortListView(lvwItemCode, 2, SortOrder.Ascending)
                .Sorting = System.Windows.Forms.SortOrder.Ascending
            End With
            Exit Sub
ErrHandler:
            gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub
    Private Sub OptItemCode_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optItemCode.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo ErrHandler
            With lvwItemCode
                .Sort()
                ListViewColumnSorter.SortListView(lvwItemCode, 0, SortOrder.Ascending)
                .Sorting = System.Windows.Forms.SortOrder.Ascending
            End With
            Exit Sub
ErrHandler:
            gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub
    Private Sub optPartNo_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optPartNo.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo ErrHandler
            With lvwItemCode
                .Sort()
                ListViewColumnSorter.SortListView(lvwItemCode, 1, SortOrder.Ascending)
                .Sorting = System.Windows.Forms.SortOrder.Ascending
            End With
            Exit Sub
ErrHandler:
            gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub
    Private Sub TxtSearch_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtsearch.TextChanged
        Call SearchItem()
    End Sub
    Public Function AddDataFromGrinDtl(ByRef pstrVend As String, ByRef dblGrnNo As Double, ByRef pstrstockLocation As String, Optional ByRef intAlreadyItem As Short = 0, Optional ByRef pstrCondition As String = "") As String
        Dim rsGrnDtl As ClsResultSetDB
        Dim strSql As String
        Dim StrItemCode As String
        Dim strItemNot As String
        Dim arrRejAcpt(,) As Object
        Dim intLoopCounter As Short
        Dim intArrLoopCount As Short
        Dim intMaxLoop As Short
        Dim intUbound As Short
        mstrInvType = "REJECTION" : mstrInvSubType = "REJECTION"
        On Error GoTo ErrHandler
        rsGrnDtl = New ClsResultSetDB
        strSql = "select a.Doc_No,a.Item_code,a.Rejected_Quantity,Despatch_Quantity = isnull(a.Despatch_Quantity,0),"
        strSql = strSql & " Inspected_Quantity = isnull(a.Inspected_Quantity,0), RGP_Quantity = isnull(a.RGP_Quantity,0)  from grn_Dtl a,"
        strSql = strSql & " grn_hdr b Where a.unit_code=b.unit_code and "
        strSql = strSql & " a.Doc_type = b.Doc_type And a.Doc_No = b.Doc_No and "
        strSql = strSql & " a.From_Location = b.From_Location and a.From_Location ='01R1'"
        strSql = strSql & " and a.Rejected_quantity > 0 and b.Vendor_code = '" & pstrVend
        strSql = strSql & "' and a.Doc_No = " & dblGrnNo & " AND ISNULL(GRN_Cancelled,0) = 0 and a.unit_code='" & gstrUNITID & "' "
        If Len(Trim(pstrCondition)) > 0 Then
            strSql = strSql & " and a.Item_code not in (" & pstrCondition & ")"
        End If
        rsGrnDtl.GetResult(strSql)
        If rsGrnDtl.GetNoRows > 0 Then
            intMaxLoop = rsGrnDtl.GetNoRows : rsGrnDtl.MoveFirst() : ReDim arrRejAcpt(2, intMaxLoop - 1) : intUbound = intMaxLoop - 1
            '****To Fatch all Doc_No and Rejected Quantity in Array
            intUbound = intMaxLoop - 1
            For intLoopCounter = 0 To intMaxLoop - 1
                arrRejAcpt(0, intLoopCounter) = rsGrnDtl.GetValue("Item_Code")
                arrRejAcpt(1, intLoopCounter) = rsGrnDtl.GetValue("Rejected_Quantity") - rsGrnDtl.GetValue("Despatch_Quantity") - rsGrnDtl.GetValue("Inspected_Quantity") - rsGrnDtl.GetValue("RGP_Quantity")
                rsGrnDtl.MoveNext()
            Next
            strItemNot = ""
            For intArrLoopCount = 0 To intUbound
                StrItemCode = "'" & arrRejAcpt(0, intArrLoopCount) & "'"
                If arrRejAcpt(1, intArrLoopCount) <= 0 Then
                    If Len(Trim(strItemNot)) > 0 Then
                        strItemNot = strItemNot & "," & StrItemCode
                    Else
                        strItemNot = StrItemCode
                    End If
                End If
            Next
            If Len(Trim(strItemNot)) > 0 Then
                If gblnGSTUnit = True Then
                    strSql = "select a.Doc_No,a.Item_code,a.Accepted_Quantity,c.HSN_SAC_CODE ,c.Description from grn_dtl a,grn_hdr b,Item_Mst c where "
                Else
                    strSql = "select a.Doc_No,a.Item_code,a.Accepted_Quantity,c.Tariff_code,c.Description from grn_dtl a,grn_hdr b,Item_Mst c where "
                End If

                strSql = strSql & "a.unit_code=b.unit_code and b.unit_code =c.unit_code "
                strSql = strSql & " and a.Doc_type = b.Doc_type and a.Doc_no = b.Doc_No "
                strSql = strSql & " and a.From_Location = b.From_Location "
                strSql = strSql & " and a.Item_Code = c.ITem_code and b.From_Location ='01R1'"
                strSql = strSql & " and a.Item_code Not in (" & strItemNot & ")"
                strSql = strSql & " and c.Status = 'A' and Hold_Flag =0"
                strSql = strSql & " and a.Rejected_quantity > 0 and b.Vendor_code = '" & pstrVend
                strSql = strSql & "' and a.Doc_No = " & dblGrnNo & " AND ISNULL(GRN_Cancelled,0) = 0 and a.unit_code='" & gstrUNITID & "'"
                strSql = strSql & " and a.Item_code in (Select Item_Code from ItemBal_Mst Where unit_code='" & gstrUNITID & "' and  Location_Code = '"
                strSql = strSql & pstrstockLocation & "' and Cur_bal > 0)"
                If Len(Trim(pstrCondition)) > 0 Then
                    strSql = strSql & " and a.Item_code not in (" & pstrCondition & ")"
                End If
            Else
                If gblnGSTUnit = True Then
                    strSql = "select a.Doc_No,a.Item_code,a.Accepted_Quantity,c.HSN_SAC_CODE,c.Description from grn_dtl a,grn_hdr b,Item_Mst c where "
                Else
                    strSql = "select a.Doc_No,a.Item_code,a.Accepted_Quantity,c.Tariff_code,c.Description from grn_dtl a,grn_hdr b,Item_Mst c where "
                End If

                strSql = strSql & " a.unit_code = b.unit_code and b.unit_code = c.unit_code "
                strSql = strSql & " and a.Doc_type = b.Doc_type and a.Doc_no = b.Doc_No "
                strSql = strSql & " and a.From_Location = b.From_Location "
                strSql = strSql & " and a.Item_Code = c.ITem_code and b.From_Location ='01R1'"
                strSql = strSql & " and c.Status = 'A' and Hold_Flag =0"
                strSql = strSql & " and a.Rejected_quantity > 0 and b.Vendor_code = '" & pstrVend
                strSql = strSql & "' and a.Doc_No = " & dblGrnNo & " AND ISNULL(GRN_Cancelled,0) = 0 "
                strSql = strSql & " and a.Item_code in (Select Item_Code from ItemBal_Mst Where unit_code='" & gstrUNITID & "' and Location_Code = '"
                strSql = strSql & pstrstockLocation & "' and Cur_bal > 0)"
                If Len(Trim(pstrCondition)) > 0 Then
                    strSql = strSql & " and a.Item_code not in (" & pstrCondition & ")"
                End If
            End If
            rsGrnDtl.ResultSetClose()
            rsGrnDtl = New ClsResultSetDB
            rsGrnDtl.GetResult(strSql)
            intMaxLoop = rsGrnDtl.GetNoRows 'assign record count to integer variable
            If intMaxLoop > 0 Then '          'if record found
                rsGrnDtl.MoveFirst() 'move to first record
                For intLoopCounter = 0 To intMaxLoop - 1
                    mListItemUserId = Me.lvwItemCode.Items.Add(rsGrnDtl.GetValue("Item_code"))
                    If mListItemUserId.SubItems.Count > 1 Then
                        mListItemUserId.SubItems(1).Text = rsGrnDtl.GetValue("Item_code")
                    Else
                        mListItemUserId.SubItems.Insert(1, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsGrnDtl.GetValue("Item_code")))
                    End If
                    If mListItemUserId.SubItems.Count > 2 Then
                        mListItemUserId.SubItems(2).Text = rsGrnDtl.GetValue("Description")
                    Else
                        mListItemUserId.SubItems.Insert(2, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsGrnDtl.GetValue("Description")))
                    End If
                    If mListItemUserId.SubItems.Count > 3 Then
                        If gblnGSTUnit = True Then
                            mListItemUserId.SubItems(3).Text = rsGrnDtl.GetValue("HSN_SAC_CODE")
                        Else
                            mListItemUserId.SubItems(3).Text = rsGrnDtl.GetValue("Tariff_Code")
                        End If

                    Else
                        If gblnGSTUnit = True Then
                            mListItemUserId.SubItems.Insert(3, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsGrnDtl.GetValue("HSN_SAC_CODE")))
                        Else
                            mListItemUserId.SubItems.Insert(3, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsGrnDtl.GetValue("Tariff_Code")))
                        End If

                    End If
                    rsGrnDtl.MoveNext() 'move to next record
                Next
            Else
                MsgBox("No Items for selected Invoice in Sales Order.Please Check Following :" & vbCrLf & "1. Item in Grin are Active and Not on Hold." & vbCrLf & "2. Check Balance of Items for location " & pstrstockLocation & "." & vbCrLf & "3. Check supplimentry Grin for items in Grin(Selected) ", MsgBoxStyle.Information, "empower")
            End If
        End If
        rsGrnDtl.ResultSetClose()
        rsGrnDtl = Nothing
        Me.ShowDialog()
        AddDataFromGrinDtl = mstrItemText
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Public Function AddDataFromGRNORLRN(ByRef pstrVend As String, ByRef strDocNo As String, ByRef pstrstockLocation As String, ByRef strRejType As String, Optional ByRef intAlreadyItem As Short = 0, Optional ByRef pstrCondition As String = "") As String
        Dim rsGrnDtl As ClsResultSetDB
        Dim strSql As String
        Dim StrItemCode As String
        Dim strItemNot As String
        Dim arrRejAcpt(,) As Object
        Dim intLoopCounter As Short
        Dim intArrLoopCount As Short
        Dim intMaxLoop As Short
        Dim intUbound As Short
        mstrInvType = "REJECTION" : mstrInvSubType = "REJECTION"
        On Error GoTo ErrHandler
        rsGrnDtl = New ClsResultSetDB
        If strRejType = "LRN" Then
            Dim CheckLRNVendor_onRejInvoice As Boolean = SqlConnectionclass.ExecuteScalar("select isnull(CheckLRNVendor_onRejInvoice,0) from sales_parameter (NOLOCK) where UNIT_CODE = '" & gstrUNITID & "'")
            If gblnGSTUnit = True Then
                If CheckLRNVendor_onRejInvoice = True Then
                    strSql = "Select B.Item_Code, I.Description, I.HSN_SAC_CODE from LRN_HDR as a Inner Join LRN_DTL as b on A.UNIT_CODE = B.UNIT_CODE and a.doc_No = b.doc_no and a.Doc_Type = b.doc_Type and a.from_Location = b.from_location Inner join Item_Mst as I On b.unit_code=i.unit_code and b.item_code=i.item_code where b.Item_code In ( Select Item_code from ItemBal_Mst where Cur_Bal>0 and Location_code ='" & pstrstockLocation & "' and unit_code='" & gstrUNITID & "') and Authorized_Code IS Not Null and a.Doc_No IN (" & strDocNo & ") and vendor_code='" & pstrVend & "' and a.unit_code='" & gstrUNITID & "'"
                Else
                    strSql = "Select B.Item_Code, I.Description, I.HSN_SAC_CODE from LRN_HDR as a Inner Join LRN_DTL as b on A.UNIT_CODE = B.UNIT_CODE and a.doc_No = b.doc_no and a.Doc_Type = b.doc_Type and a.from_Location = b.from_location Inner join Item_Mst as I On b.unit_code=i.unit_code and b.item_code=i.item_code where b.Item_code In ( Select Item_code from ItemBal_Mst where Cur_Bal>0 and Location_code ='" & pstrstockLocation & "' and unit_code='" & gstrUNITID & "') and Authorized_Code IS Not Null and a.Doc_No IN (" & strDocNo & ") and a.unit_code='" & gstrUNITID & "'"
                End If
            Else
                strSql = "Select B.Item_Code, I.Description, I.Tariff_code from LRN_HDR as a Inner Join LRN_DTL as b on A.UNIT_CODE = B.UNIT_CODE and a.doc_No = b.doc_no and a.Doc_Type = b.doc_Type and a.from_Location = b.from_location Inner join Item_Mst as I On b.unit_code=i.unit_code and b.item_code=i.item_code where b.Item_code In ( Select Item_code from ItemBal_Mst where Cur_Bal>0 and Location_code ='" & pstrstockLocation & "' and unit_code='" & gstrUNITID & "') and Authorized_Code IS Not Null and a.Doc_No IN (" & strDocNo & ") and a.unit_code='" & gstrUNITID & "' "
            End If

            If Len(Trim(pstrCondition)) > 0 Then
                strSql = strSql & " and B.Item_code not in (" & pstrCondition & ")"
            End If
            rsGrnDtl.ResultSetClose()
            rsGrnDtl = New ClsResultSetDB
            rsGrnDtl.GetResult(strSql)
            intMaxLoop = rsGrnDtl.GetNoRows 'assign record count to integer variable
            If intMaxLoop > 0 Then '          'if record found
                rsGrnDtl.MoveFirst() 'move to first record
                For intLoopCounter = 0 To intMaxLoop - 1
                    mListItemUserId = Me.lvwItemCode.Items.Add(rsGrnDtl.GetValue("Item_code"))
                    If mListItemUserId.SubItems.Count > 1 Then
                        mListItemUserId.SubItems(1).Text = rsGrnDtl.GetValue("Item_code")
                    Else
                        mListItemUserId.SubItems.Insert(1, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsGrnDtl.GetValue("Item_code")))
                    End If
                    If mListItemUserId.SubItems.Count > 2 Then
                        mListItemUserId.SubItems(2).Text = rsGrnDtl.GetValue("Description")
                    Else
                        mListItemUserId.SubItems.Insert(2, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsGrnDtl.GetValue("Description")))
                    End If
                    If mListItemUserId.SubItems.Count > 3 Then
                        If gblnGSTUnit = True Then
                            mListItemUserId.SubItems(3).Text = rsGrnDtl.GetValue("HSN_SAC_CODE")
                        Else
                            mListItemUserId.SubItems(3).Text = rsGrnDtl.GetValue("Tariff_Code")
                        End If

                    Else
                        If gblnGSTUnit = True Then
                            mListItemUserId.SubItems.Insert(3, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsGrnDtl.GetValue("HSN_SAC_CODE")))
                        Else
                            mListItemUserId.SubItems.Insert(3, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsGrnDtl.GetValue("Tariff_Code")))
                        End If

                    End If
                    rsGrnDtl.MoveNext() 'move to next record
                Next
            End If
        Else
            strSql = "select a.Doc_No,a.Item_code,a.Rejected_Quantity,Despatch_Quantity = isnull(a.Despatch_Quantity,0),"
            strSql = strSql & " Inspected_Quantity = isnull(a.Inspected_Quantity,0), RGP_Quantity = isnull(a.RGP_Quantity,0)  from grn_Dtl a,"
            strSql = strSql & " grn_hdr b Where a.unit_code = b.unit_code and "
            strSql = strSql & " a.Doc_type = b.Doc_type And a.Doc_No = b.Doc_No and "
            strSql = strSql & " a.From_Location = b.From_Location and a.From_Location ='01R1'"
            strSql = strSql & " and a.Rejected_quantity > 0 and b.Vendor_code = '" & pstrVend
            strSql = strSql & "' and a.Doc_No in (" & strDocNo & ") AND ISNULL(GRN_Cancelled,0) = 0 and a.unit_code='" & gstrUNITID & "'"
            If Len(Trim(pstrCondition)) > 0 Then
                strSql = strSql & " and a.Item_code not in (" & pstrCondition & ")"
            End If
            rsGrnDtl.ResultSetClose()
            rsGrnDtl = New ClsResultSetDB()
            rsGrnDtl.GetResult(strSql)
            If rsGrnDtl.GetNoRows > 0 Then
                intMaxLoop = rsGrnDtl.GetNoRows : rsGrnDtl.MoveFirst() : ReDim arrRejAcpt(2, intMaxLoop - 1) : intUbound = intMaxLoop - 1
                '****To Fatch all Doc_No and Rejected Quantity in Array
                intUbound = intMaxLoop - 1
                For intLoopCounter = 0 To intMaxLoop - 1
                    arrRejAcpt(0, intLoopCounter) = rsGrnDtl.GetValue("Item_Code")
                    arrRejAcpt(1, intLoopCounter) = rsGrnDtl.GetValue("Rejected_Quantity") - rsGrnDtl.GetValue("Despatch_Quantity") - rsGrnDtl.GetValue("Inspected_Quantity") - rsGrnDtl.GetValue("RGP_Quantity")
                    rsGrnDtl.MoveNext()
                Next
                strItemNot = ""
                For intArrLoopCount = 0 To intUbound
                    StrItemCode = "'" & arrRejAcpt(0, intArrLoopCount) & "'"
                    If arrRejAcpt(1, intArrLoopCount) <= 0 Then
                        If Len(Trim(strItemNot)) > 0 Then
                            strItemNot = strItemNot & "," & StrItemCode
                        Else
                            strItemNot = StrItemCode
                        End If
                    End If
                Next
                If Len(Trim(strDocNo)) = 0 Then
                    If gblnGSTUnit = True Then
                        strSql = "select Distinct a.Item_code, c.HSN_SAC_CODE,c.Description from grn_dtl a,grn_hdr b,Item_Mst c where "
                    Else
                        strSql = "select Distinct a.Item_code, c.Tariff_code,c.Description from grn_dtl a,grn_hdr b,Item_Mst c where "
                    End If

                    strSql = strSql & " a.unit_code = b.unit_code and b.unit_code = c.unit_code "
                    strSql = strSql & " and a.Doc_type = b.Doc_type and a.Doc_no = b.Doc_No "
                    strSql = strSql & " and a.From_Location = b.From_Location "
                    strSql = strSql & " and a.Item_Code = c.ITem_code and b.From_Location ='01R1'"
                    strSql = strSql & " and a.Item_code Not in (" & strItemNot & ")"
                    strSql = strSql & " and c.Status = 'A' and Hold_Flag =0"
                    strSql = strSql & " and a.Rejected_quantity > 0 and b.Vendor_code = '" & pstrVend
                    strSql = strSql & "' and a.Doc_No = " & strDocNo & " AND ISNULL(GRN_Cancelled,0) = 0  and a.unit_code='" & gstrUNITID & "' "
                    strSql = strSql & " and a.Item_code in (Select Item_Code from ItemBal_Mst Where unit_code='" & gstrUNITID & "' and Location_Code = '"
                    strSql = strSql & pstrstockLocation & "' and Cur_bal > 0)"
                    If Len(Trim(pstrCondition)) > 0 Then
                        strSql = strSql & " and a.Item_code not in (" & pstrCondition & ")"
                    End If
                Else
                    If gblnGSTUnit = True Then
                        strSql = "select Distinct a.Item_code, C.HSN_SAC_CODE,c.Description from grn_dtl a, grn_hdr b, Item_Mst c where "
                    Else
                        strSql = "select Distinct a.Item_code, c.Tariff_code,c.Description from grn_dtl a, grn_hdr b, Item_Mst c where "
                    End If

                    strSql = strSql & " a.unit_code = b.unit_code and b.unit_code = c.unit_code "
                    strSql = strSql & " and a.Doc_type = b.Doc_type and a.Doc_no = b.Doc_No "
                    strSql = strSql & " and a.From_Location = b.From_Location "
                    strSql = strSql & " and a.Item_Code = c.ITem_code and b.From_Location ='01R1'"
                    strSql = strSql & " and c.Status = 'A' and Hold_Flag =0"
                    strSql = strSql & " and a.Rejected_quantity > 0 and b.Vendor_code = '" & pstrVend
                    strSql = strSql & "' and a.Doc_No In ( " & strDocNo & ") AND ISNULL(GRN_Cancelled,0) = 0  and a.unit_code='" & gstrUNITID & "'"
                    strSql = strSql & " and a.Item_code in (Select Item_Code from ItemBal_Mst Where unit_code='" & gstrUNITID & "' and Location_Code = '"
                    strSql = strSql & pstrstockLocation & "' and Cur_bal > 0)"
                    If Len(Trim(pstrCondition)) > 0 Then
                        strSql = strSql & " and a.Item_code not in (" & pstrCondition & ")"
                    End If
                End If
                rsGrnDtl.ResultSetClose()
                rsGrnDtl = New ClsResultSetDB
                rsGrnDtl.GetResult(strSql)
                intMaxLoop = rsGrnDtl.GetNoRows 'assign record count to integer variable
                If intMaxLoop > 0 Then '          'if record found
                    rsGrnDtl.MoveFirst() 'move to first record
                    For intLoopCounter = 0 To intMaxLoop - 1
                        mListItemUserId = Me.lvwItemCode.Items.Add(rsGrnDtl.GetValue("Item_code"))
                        If mListItemUserId.SubItems.Count > 1 Then
                            mListItemUserId.SubItems(1).Text = rsGrnDtl.GetValue("Item_code")
                        Else
                            mListItemUserId.SubItems.Insert(1, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsGrnDtl.GetValue("Item_code")))
                        End If
                        If mListItemUserId.SubItems.Count > 2 Then
                            mListItemUserId.SubItems(2).Text = rsGrnDtl.GetValue("Description")
                        Else
                            mListItemUserId.SubItems.Insert(2, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsGrnDtl.GetValue("Description")))
                        End If
                        If mListItemUserId.SubItems.Count > 3 Then
                            If gblnGSTUnit = True Then
                                mListItemUserId.SubItems(3).Text = rsGrnDtl.GetValue("HSN_SAC_CODE")
                            Else
                                mListItemUserId.SubItems(3).Text = rsGrnDtl.GetValue("Tariff_Code")
                            End If

                        Else
                            If gblnGSTUnit = True Then
                                mListItemUserId.SubItems.Insert(3, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsGrnDtl.GetValue("HSN_SAC_CODE")))
                            Else
                                mListItemUserId.SubItems.Insert(3, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsGrnDtl.GetValue("Tariff_Code")))
                            End If

                        End If
                        rsGrnDtl.MoveNext() 'move to next record
                    Next
                Else
                    MsgBox("No Items for selected Invoice in Sales Order.Please Check Following :" & vbCrLf & "1. Item in Grin are Active and Not on Hold." & vbCrLf & "2. Check Balance of Items for location " & pstrstockLocation & "." & vbCrLf & "3. Check supplimentry Grin for items in Grin(Selected) ", MsgBoxStyle.Information, "empower")
                    Exit Function
                End If
            End If
        End If
        rsGrnDtl.ResultSetClose()
        rsGrnDtl = Nothing
        Me.ShowDialog()
        AddDataFromGRNORLRN = mstrItemText
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
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
        strQuery = "select dt from calendar_mst where unit_code='" & gstrUNITID & "' and dt > '" & VB6.Format(pstrDate, "dd/mmm/yyyy") & "' and work_flg<>1 order by dt"
        If rsCalendarDate.State = ADODB.ObjectStateEnum.adStateOpen Then rsCalendarDate.Close()
        rsCalendarDate.Open(strQuery, mP_Connection, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockBatchOptimistic)
        If rsCalendarDate.EOF Or rsCalendarDate.BOF Or IsDBNull(rsCalendarDate.Fields("dt").Value) Then
            MsgBox("Date in Calendar Master not defined !", MsgBoxStyle.Information, "eMPro")
            GetNextWorkingDay = CStr(-1)
            rsCalendarDate.Close()
            Exit Function
        Else
            rsCalendarDate.MoveFirst()
            'GetNextWorkingDay = VB6.Format(rsCalendarDate.Fields("dt").Value, "dd/mmm/yyyy")
            GetNextWorkingDay = VB6.Format(rsCalendarDate.Fields("dt").Value, gstrDateFormat)
        End If
        rsCalendarDate.Close()
        Exit Function
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function InvoiceForMTL() As Boolean
        '----------------------------------------------------------------------------
        'Author         :   Ashutosh Verma
        'Argument       :
        'Return Value   :   TRUE / FALSE
        'Function       :
        'Comments       :   Created on 01 May 2007 ,  Issue Id:19911
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim clsMTLInvoice As ClsResultSetDB
        clsMTLInvoice = New ClsResultSetDB
        clsMTLInvoice.GetResult("Select isnull(InvoiceForMTLSharjah,0) as InvoiceForMTLSharjah from sales_parameter where unit_code='" & gstrUNITID & "'")
        If clsMTLInvoice.GetNoRows > 0 Then
            InvoiceForMTL = clsMTLInvoice.GetValue("InvoiceForMTLSharjah")
        Else
            InvoiceForMTL = False
        End If
        clsMTLInvoice.ResultSetClose()
        clsMTLInvoice = Nothing
        Exit Function
ErrHandler:
        InvoiceForMTL = False
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

#Region "Global Tool Invoice Changes"
    'ADDED BY VINOD FOR GLOBAL TOOL CHANGES
    Public Function MakeSelectSubQuery_AssetInvoice(ByRef pstrCustno As String, ByRef pstrRefNo As String, ByRef pstrAmmNo As String, ByRef pstrstockLocation As String, ByRef pstrItemin As String, Optional ByRef pstrItemNotin As String = "") As String
        'ADDED  BY VINOD
        Dim strSelectSql As String
        If gblnGSTUnit = True Then
            strSelectSql = "Select c.Item_Code,c.Cust_DrgNo,c.Cust_Drg_Desc,d.hsn_sac_code from Cust_Ord_hdr a,Cust_ord_dtl c,Item_Mst d where "
        Else
            strSelectSql = "Select c.Item_Code,c.Cust_DrgNo,c.Cust_Drg_Desc,d.Tariff_Code from Cust_Ord_hdr a,Cust_ord_dtl c,Item_Mst d where "
        End If

        strSelectSql = strSelectSql & " a.unit_code = c.unit_code and c.unit_code = d.unit_code "
        strSelectSql = strSelectSql & " and a.Cust_ref = c.Cust_ref and a.amendment_No = c.amendment_No and a.Account_code=c.account_code and a.unit_code='" & gstrUNITID & "' "
        strSelectSql = strSelectSql & " and  c.Item_code = d.Item_code and a.Account_Code='" & Trim(pstrCustno) & "' and a.Cust_Ref='" & Trim(pstrRefNo)
        strSelectSql = strSelectSql & "' and a.Amendment_No='" & Trim(pstrAmmNo) & "' And c.Active_Flag = 'A' "
        strSelectSql = strSelectSql & " and c.Item_Code in (Select a.Item_code from Item_MSt a,Itembal_mst b where a.unit_code=b.unit_code and a.Item_Main_grp in (" & pstrItemin & ") and a.Item_code = b.Item_code and b.Location_code ='" & pstrstockLocation & "' and b.Cur_bal >0 and a.hold_flag =0 and a.Status = 'A' and a.unit_code='" & gstrUNITID & "'"
        If Len(Trim(pstrItemNotin)) > 0 Then
            strSelectSql = strSelectSql & " and a.Item_code not in ( " & pstrItemNotin & "))"
        Else
            strSelectSql = strSelectSql & ")"
        End If
        strSelectSql = strSelectSql & " AND C.ITEM_CODE NOT IN "
        strSelectSql = strSelectSql & " ( "
        strSelectSql = strSelectSql & " SELECT I.ITEM_CODE FROM ITEM_MST I INNER JOIN GLOBAL_TOOL_MST G "
        strSelectSql = strSelectSql & " ON I.ITEM_CODE = G.GLOBAL_TOOL_CODE_FOR_ITEM_MASTER WHERE I.UNIT_CODE = '" & gstrUNITID & "'"
        strSelectSql = strSelectSql & " AND ITEM_MAIN_GRP ='P' "
        strSelectSql = strSelectSql & " AND G.CATEGORY ='MOULDING' AND I.STATUS='A'"
        strSelectSql = strSelectSql & " ) "
        strSelectSql = strSelectSql & " UNION ALL "
        If gblnGSTUnit = True Then
            strSelectSql = strSelectSql & " Select c.Item_Code,c.Cust_DrgNo,c.Cust_Drg_Desc,d.hsn_sac_code from Cust_Ord_hdr a,Cust_ord_dtl c,Item_Mst d where "
        Else
            strSelectSql = strSelectSql & " Select c.Item_Code,c.Cust_DrgNo,c.Cust_Drg_Desc,d.Tariff_Code from Cust_Ord_hdr a,Cust_ord_dtl c,Item_Mst d where "
        End If

        strSelectSql = strSelectSql & " a.unit_code = c.unit_code and c.unit_code = d.unit_code "
        strSelectSql = strSelectSql & " and a.Cust_ref = c.Cust_ref and a.amendment_No = c.amendment_No and a.Account_code=c.account_code and a.unit_code='" & gstrUNITID & "' "
        strSelectSql = strSelectSql & " and  c.Item_code = d.Item_code and a.Account_Code='" & Trim(pstrCustno) & "' and a.Cust_Ref='" & Trim(pstrRefNo)
        strSelectSql = strSelectSql & "' and a.Amendment_No='" & Trim(pstrAmmNo) & "' And c.Active_Flag = 'A' "
        strSelectSql = strSelectSql & " and c.Item_Code in (Select a.Item_code from Item_MSt a,Itembal_mst b where a.unit_code=b.unit_code and a.Item_Main_grp in (" & pstrItemin & ") and a.Item_code = b.Item_code and b.Location_code ='" & pstrstockLocation & "' and b.Cur_bal >0 and a.hold_flag =0 and a.Status = 'A' and a.unit_code='" & gstrUNITID & "'"
        If Len(Trim(pstrItemNotin)) > 0 Then
            strSelectSql = strSelectSql & " and a.Item_code not in ( " & pstrItemNotin & "))"
        Else
            strSelectSql = strSelectSql & ")"
        End If
        strSelectSql = strSelectSql & " AND C.ITEM_CODE IN "
        strSelectSql = strSelectSql & " ( "
        strSelectSql = strSelectSql & " SELECT I.ITEM_CODE FROM ITEM_MST I INNER JOIN GLOBAL_TOOL_MST G "
        strSelectSql = strSelectSql & " ON I.ITEM_CODE = G.GLOBAL_TOOL_CODE_FOR_ITEM_MASTER "
        strSelectSql = strSelectSql & " WHERE I.UNIT_CODE = '" & gstrUNITID & "' AND ITEM_MAIN_GRP ='P' "
        strSelectSql = strSelectSql & " AND G.CATEGORY ='MOULDING' AND I.STATUS='A' AND G.VENDOR_CODE = '" & pstrCustno & "'"
        strSelectSql = strSelectSql & " AND G.TRANSFER_COMPLETED = 0 AND G.TRANSFER_TYPE  in ('SALE','CUSTOMER_TOOL_SALE') AND ISNULL(G.INV_NO,0) = 0 "
        strSelectSql = strSelectSql & " ) "
        MakeSelectSubQuery_AssetInvoice = strSelectSql

    End Function

    Public Function makeSelectSql_AssetInvoice(ByRef pstrCustno As String, ByRef pstrRefNo As String, ByRef pstrAmmNo As String, ByRef effectyrmon As String, ByRef Validyrmon As String, ByRef pstrstockLocation As String, ByRef strDate As String, ByRef pstrItemin As String, Optional ByRef pstrCondition As String = "") As String
        'ADDED BY VINOD
        Dim strSelectSql As String
        Dim strNextWorkDay As String
        Dim RsobjSchedules As New ADODB.Recordset
        Dim blnCalendarDateTrac As Boolean
        If RsobjSchedules.State = ADODB.ObjectStateEnum.adStateOpen Then RsobjSchedules.Close()
        RsobjSchedules.Open("SELECT DSWiseTracking,CalendarDateTrac FROM sales_parameter where unit_code='" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If Not RsobjSchedules.EOF Then
            If IIf(RsobjSchedules.Fields(1).Value, 1, 0) = 1 Then
                blnCalendarDateTrac = True
            Else
                blnCalendarDateTrac = False
            End If
        End If
        RsobjSchedules.Close()
        If blnCalendarDateTrac Then
            strNextWorkDay = GetNextWorkingDay(strDate)
            If strNextWorkDay = "-1" Then
                makeSelectSql_AssetInvoice = ""
                Exit Function
            End If
        End If
        If gblnGSTUnit = True Then
            strSelectSql = "Select b.Item_Code,c.Cust_DrgNo,c.Cust_Drg_Desc,d.hsn_sac_code from Cust_Ord_hdr a,MonthlyMktSchedule b,Cust_ord_dtl c,Item_Mst d where "
        Else
            strSelectSql = "Select b.Item_Code,c.Cust_DrgNo,c.Cust_Drg_Desc,d.Tariff_Code from Cust_Ord_hdr a,MonthlyMktSchedule b,Cust_ord_dtl c,Item_Mst d where "
        End If

        strSelectSql = strSelectSql & " a.unit_code = b.unit_code and b.unit_code = c.unit_code and c.unit_code = d.unit_code "
        strSelectSql = strSelectSql & " and a.Cust_ref = c.Cust_ref and a.amendment_No = c.amendment_No and a.Account_code=c.account_code And c.Active_Flag ='A' "
        strSelectSql = strSelectSql & " and a.account_code=b.Account_code and c.Cust_drgNo=b.Cust_drgNo and b.ITem_code = d.Item_code and a.Account_Code='" & Trim(pstrCustno) & "' and a.unit_code='" & gstrUNITID & "'"
        strSelectSql = strSelectSql & " and a.Cust_Ref='" & Trim(pstrRefNo) & "' and a.Amendment_No='" & Trim(pstrAmmNo) & "' and b.status = 1 and b.Schedule_flag =1 and b.Year_Month =  " & Validyrmon
        strSelectSql = strSelectSql & " and b.Item_Code in(Select a.Item_code from Item_MSt a,Itembal_mst b where a.unit_code=b.unit_code and a.unit_code='" & gstrUNITID & "' and a.Item_Main_grp in (" & Trim(pstrItemin) & ") and a.Item_code = b.Item_code and b.Location_code ='" & pstrstockLocation & "' and b.Cur_bal >0 and a.hold_flag =0 and a.Status = 'A'"
        If Len(Trim(pstrCondition)) > 0 Then
            strSelectSql = strSelectSql & " and a.Item_code not in(" & pstrCondition & "))"
        Else
            strSelectSql = strSelectSql & ")"
        End If
        strSelectSql = strSelectSql & " AND B.ITEM_CODE NOT IN "
        strSelectSql = strSelectSql & " ( "
        strSelectSql = strSelectSql & " SELECT I.ITEM_CODE FROM ITEM_MST I INNER JOIN GLOBAL_TOOL_MST G "
        strSelectSql = strSelectSql & " ON I.ITEM_CODE = G.GLOBAL_TOOL_CODE_FOR_ITEM_MASTER WHERE I.UNIT_CODE = '" & gstrUNITID & "'"
        strSelectSql = strSelectSql & " AND ITEM_MAIN_GRP ='P' "
        strSelectSql = strSelectSql & " AND G.CATEGORY IN ('MOULDING','PRESSTOOL') AND I.STATUS='A'"
        strSelectSql = strSelectSql & " ) "

        strSelectSql = strSelectSql & " UNION ALL "
        If gblnGSTUnit = True Then
            strSelectSql = "Select b.Item_Code,c.Cust_DrgNo,c.Cust_Drg_Desc,d.hsn_sac_code from Cust_Ord_hdr a,MonthlyMktSchedule b,Cust_ord_dtl c,Item_Mst d where "
        Else
            strSelectSql = "Select b.Item_Code,c.Cust_DrgNo,c.Cust_Drg_Desc,d.Tariff_Code from Cust_Ord_hdr a,MonthlyMktSchedule b,Cust_ord_dtl c,Item_Mst d where "
        End If

        strSelectSql = strSelectSql & " a.unit_code = b.unit_code and b.unit_code = c.unit_code and c.unit_code = d.unit_code "
        strSelectSql = strSelectSql & " and a.Cust_ref = c.Cust_ref and a.amendment_No = c.amendment_No and a.Account_code=c.account_code And c.Active_Flag ='A' "
        strSelectSql = strSelectSql & " and a.account_code=b.Account_code and c.Cust_drgNo=b.Cust_drgNo and b.ITem_code = d.Item_code and a.Account_Code='" & Trim(pstrCustno) & "' and a.unit_code='" & gstrUNITID & "'"
        strSelectSql = strSelectSql & " and a.Cust_Ref='" & Trim(pstrRefNo) & "' and a.Amendment_No='" & Trim(pstrAmmNo) & "' and b.status = 1 and b.Schedule_flag =1 and b.Year_Month =  " & Validyrmon
        strSelectSql = strSelectSql & " and b.Item_Code in(Select a.Item_code from Item_MSt a,Itembal_mst b where a.unit_code=b.unit_code and a.unit_code='" & gstrUNITID & "' and a.Item_Main_grp in (" & Trim(pstrItemin) & ") and a.Item_code = b.Item_code and b.Location_code ='" & pstrstockLocation & "' and b.Cur_bal >0 and a.hold_flag =0 and a.Status = 'A'"
        If Len(Trim(pstrCondition)) > 0 Then
            strSelectSql = strSelectSql & " and a.Item_code not in(" & pstrCondition & "))"
        Else
            strSelectSql = strSelectSql & ")"
        End If
        strSelectSql = strSelectSql & " AND B.ITEM_CODE IN "
        strSelectSql = strSelectSql & " ( "
        strSelectSql = strSelectSql & " SELECT I.ITEM_CODE FROM ITEM_MST I INNER JOIN GLOBAL_TOOL_MST G "
        strSelectSql = strSelectSql & " ON I.ITEM_CODE = G.GLOBAL_TOOL_CODE_FOR_ITEM_MASTER "
        strSelectSql = strSelectSql & " WHERE I.UNIT_CODE = '" & gstrUNITID & "' AND ITEM_MAIN_GRP ='P' "
        strSelectSql = strSelectSql & " AND G.CATEGORY IN ('MOULDING','PRESSTOOL') AND I.STATUS='A' AND G.VENDOR_CODE = '" & pstrCustno & "'"
        strSelectSql = strSelectSql & " AND G.TRANSFER_COMPLETED = 0 AND G.TRANSFER_TYPE  in ('SALE','CUSTOMER_TOOL_SALE') AND ISNULL(G.INV_NO,0) = 0 "
        strSelectSql = strSelectSql & " ) "

        strSelectSql = strSelectSql & " UNION "
        If gblnGSTUnit = True Then
            strSelectSql = strSelectSql & " Select b.Item_Code,c.Cust_DrgNo,c.Cust_Drg_Desc,d.hsn_sac_code from Cust_Ord_hdr a,DailyMktSchedule b,Cust_ord_dtl c,ITem_Mst d  where "
        Else
            strSelectSql = strSelectSql & " Select b.Item_Code,c.Cust_DrgNo,c.Cust_Drg_Desc,d.Tariff_Code from Cust_Ord_hdr a,DailyMktSchedule b,Cust_ord_dtl c,ITem_Mst d  where "
        End If

        strSelectSql = strSelectSql & " a.unit_code = b.unit_code and b.unit_code = c.unit_code and c.unit_code = d.unit_code "
        strSelectSql = strSelectSql & " and a.Cust_ref = c.Cust_ref and a.amendment_No = c.amendment_No and a.Account_code=c.account_code"
        strSelectSql = strSelectSql & " and a.account_code=b.Account_code and c.Cust_drgNo=b.Cust_drgNo and b.ITem_code =d.ITem_code and b.status = 1 And c.Active_Flag ='A' and a.Account_Code='" & Trim(pstrCustno) & "' and a.unit_code='" & gstrUNITID & "' "
        strSelectSql = strSelectSql & " and a.Cust_Ref='" & Trim(pstrRefNo) & "' and a.Amendment_No='" & Trim(pstrAmmNo) & "' "
        If blnCalendarDateTrac Then
            If Month(ConvertToDate(strDate)) = Month(ConvertToDate(strNextWorkDay)) Then
                strSelectSql = strSelectSql & " and  datepart(mm,b.trans_date) = '" & Month(ConvertToDate(strDate)) & "' And  b.Trans_Date <= '" & VB6.Format(strNextWorkDay, "dd/mmm/yyyy") & "'  And DatePart(yyyy, b.Trans_Date) = '" & Year(ConvertToDate(strDate)) & "'"
            ElseIf (Month(ConvertToDate(strDate)) + 1) = Month(ConvertToDate(strNextWorkDay)) Then
                strSelectSql = strSelectSql & " and  datepart(mm,b.trans_date) in ('" & Month(ConvertToDate(strDate)) & "','" & Month(ConvertToDate(strNextWorkDay)) & "') And  b.Trans_Date <= '" & VB6.Format(strNextWorkDay, "dd/mmm/yyyy") & "'  And DatePart(yyyy, b.Trans_Date) = '" & Year(ConvertToDate(strDate)) & "'"
            ElseIf (Year(ConvertToDate(strDate)) + 1) = Year(ConvertToDate(strNextWorkDay)) Then
                strSelectSql = strSelectSql & " and  ((datepart(yyyy,b.Trans_Date) = '" & Year(ConvertToDate(strDate)) & "' and datepart(mm,b.Trans_Date)='" & Month(ConvertToDate(strDate)) & "') or ( datepart(yyyy,b.Trans_Date) = '" & Year(CDate(strNextWorkDay)) & "' and datepart(mm,b.Trans_Date)='" & Month(CDate(strNextWorkDay)) & "')) And  b.Trans_Date <= '" & VB6.Format(strNextWorkDay, "dd/mmm/yyyy") & "'"
            End If
        Else
            strSelectSql = strSelectSql & " and  datepart(mm,b.trans_date) = '" & Month(ConvertToDate(strDate)) & "' And  b.Trans_Date <= '" & getDateForDB(strDate) & "'  And DatePart(yyyy, b.Trans_Date) = '" & Year(ConvertToDate(strDate)) & "'"
        End If
        strSelectSql = strSelectSql & " and b.Item_Code in(Select a.Item_code from Item_MSt a,Itembal_mst b where a.unit_code = b.unit_code and a.Item_Main_grp in (" & Trim(pstrItemin) & ") and a.Item_code = b.Item_code and b.Location_code ='" & pstrstockLocation & "' and b.Cur_bal >0 and a.hold_flag =0 and a.Status = 'A' and a.unit_code='" & gstrUNITID & "'"
        If Len(Trim(pstrCondition)) > 0 Then
            strSelectSql = strSelectSql & " and a.Item_code not in( " & pstrCondition & "))"
        Else
            strSelectSql = strSelectSql & ")"
        End If
        strSelectSql = strSelectSql & " AND B.ITEM_CODE NOT IN "
        strSelectSql = strSelectSql & " ( "
        strSelectSql = strSelectSql & " SELECT I.ITEM_CODE FROM ITEM_MST I INNER JOIN GLOBAL_TOOL_MST G "
        strSelectSql = strSelectSql & " ON I.ITEM_CODE = G.GLOBAL_TOOL_CODE_FOR_ITEM_MASTER WHERE I.UNIT_CODE = '" & gstrUNITID & "'"
        strSelectSql = strSelectSql & " AND ITEM_MAIN_GRP ='P' "
        strSelectSql = strSelectSql & " AND G.CATEGORY IN ('MOULDING','PRESSTOOL') AND I.STATUS='A'"
        strSelectSql = strSelectSql & " ) "

        strSelectSql = strSelectSql & " UNION ALL "
        If gblnGSTUnit = True Then
            strSelectSql = strSelectSql & " Select b.Item_Code,c.Cust_DrgNo,c.Cust_Drg_Desc,d.hsn_sac_code from Cust_Ord_hdr a,DailyMktSchedule b,Cust_ord_dtl c,ITem_Mst d  where "
        Else
            strSelectSql = strSelectSql & " Select b.Item_Code,c.Cust_DrgNo,c.Cust_Drg_Desc,d.Tariff_Code from Cust_Ord_hdr a,DailyMktSchedule b,Cust_ord_dtl c,ITem_Mst d  where "
        End If

        strSelectSql = strSelectSql & " a.unit_code = b.unit_code and b.unit_code = c.unit_code and c.unit_code = d.unit_code "
        strSelectSql = strSelectSql & " and a.Cust_ref = c.Cust_ref and a.amendment_No = c.amendment_No and a.Account_code=c.account_code"
        strSelectSql = strSelectSql & " and a.account_code=b.Account_code and c.Cust_drgNo=b.Cust_drgNo and b.ITem_code =d.ITem_code and b.status = 1 And c.Active_Flag ='A' and a.Account_Code='" & Trim(pstrCustno) & "' and a.unit_code='" & gstrUNITID & "' "
        strSelectSql = strSelectSql & " and a.Cust_Ref='" & Trim(pstrRefNo) & "' and a.Amendment_No='" & Trim(pstrAmmNo) & "' "
        If blnCalendarDateTrac Then
            If Month(ConvertToDate(strDate)) = Month(ConvertToDate(strNextWorkDay)) Then
                strSelectSql = strSelectSql & " and  datepart(mm,b.trans_date) = '" & Month(ConvertToDate(strDate)) & "' And  b.Trans_Date <= '" & VB6.Format(strNextWorkDay, "dd/mmm/yyyy") & "'  And DatePart(yyyy, b.Trans_Date) = '" & Year(ConvertToDate(strDate)) & "'"
            ElseIf (Month(ConvertToDate(strDate)) + 1) = Month(ConvertToDate(strNextWorkDay)) Then
                strSelectSql = strSelectSql & " and  datepart(mm,b.trans_date) in ('" & Month(ConvertToDate(strDate)) & "','" & Month(ConvertToDate(strNextWorkDay)) & "') And  b.Trans_Date <= '" & VB6.Format(strNextWorkDay, "dd/mmm/yyyy") & "'  And DatePart(yyyy, b.Trans_Date) = '" & Year(ConvertToDate(strDate)) & "'"
            ElseIf (Year(ConvertToDate(strDate)) + 1) = Year(ConvertToDate(strNextWorkDay)) Then
                strSelectSql = strSelectSql & " and  ((datepart(yyyy,b.Trans_Date) = '" & Year(ConvertToDate(strDate)) & "' and datepart(mm,b.Trans_Date)='" & Month(ConvertToDate(strDate)) & "') or ( datepart(yyyy,b.Trans_Date) = '" & Year(CDate(strNextWorkDay)) & "' and datepart(mm,b.Trans_Date)='" & Month(CDate(strNextWorkDay)) & "')) And  b.Trans_Date <= '" & VB6.Format(strNextWorkDay, "dd/mmm/yyyy") & "'"
            End If
        Else
            strSelectSql = strSelectSql & " and  datepart(mm,b.trans_date) = '" & Month(ConvertToDate(strDate)) & "' And  b.Trans_Date <= '" & getDateForDB(strDate) & "'  And DatePart(yyyy, b.Trans_Date) = '" & Year(ConvertToDate(strDate)) & "'"
        End If
        strSelectSql = strSelectSql & " and b.Item_Code in(Select a.Item_code from Item_MSt a,Itembal_mst b where a.unit_code = b.unit_code and a.Item_Main_grp in (" & Trim(pstrItemin) & ") and a.Item_code = b.Item_code and b.Location_code ='" & pstrstockLocation & "' and b.Cur_bal >0 and a.hold_flag =0 and a.Status = 'A' and a.unit_code='" & gstrUNITID & "'"
        If Len(Trim(pstrCondition)) > 0 Then
            strSelectSql = strSelectSql & " and a.Item_code not in( " & pstrCondition & "))"
        Else
            strSelectSql = strSelectSql & ")"
        End If
        strSelectSql = strSelectSql & " AND B.ITEM_CODE IN "
        strSelectSql = strSelectSql & " ( "
        strSelectSql = strSelectSql & " SELECT I.ITEM_CODE FROM ITEM_MST I INNER JOIN GLOBAL_TOOL_MST G "
        strSelectSql = strSelectSql & " ON I.ITEM_CODE = G.GLOBAL_TOOL_CODE_FOR_ITEM_MASTER "
        strSelectSql = strSelectSql & " WHERE I.UNIT_CODE = '" & gstrUNITID & "' AND ITEM_MAIN_GRP ='P' "
        strSelectSql = strSelectSql & " AND G.CATEGORY IN ('MOULDING','PRESSTOOL') AND I.STATUS='A' AND G.VENDOR_CODE = '" & pstrCustno & "'"
        strSelectSql = strSelectSql & " AND G.TRANSFER_COMPLETED = 0 AND G.TRANSFER_TYPE IN ('SALE','CUSTOMER_TOOL_SALE') AND ISNULL(G.INV_NO,0) = 0 "
        strSelectSql = strSelectSql & " ) "
        makeSelectSql_AssetInvoice = strSelectSql
    End Function

    Public Function MakeSelectSubQuery_ToolInvoice(ByRef pstrCustno As String, ByRef pstrRefNo As String, ByRef pstrAmmNo As String, ByRef pstrstockLocation As String, ByRef pstrItemin As String, Optional ByRef pstrItemNotin As String = "") As String
        'ADDED  BY VINOD
        Dim strSelectSql As String
        If gblnGSTUnit = True Then
            strSelectSql = "Select c.Item_Code,c.Cust_DrgNo,c.Cust_Drg_Desc,d.hsn_sac_code from Cust_Ord_hdr a,Cust_ord_dtl c,Item_Mst d where "
        Else
            strSelectSql = "Select c.Item_Code,c.Cust_DrgNo,c.Cust_Drg_Desc,d.Tariff_Code from Cust_Ord_hdr a,Cust_ord_dtl c,Item_Mst d where "
        End If
        strSelectSql = strSelectSql & " a.unit_code = c.unit_code and c.unit_code = d.unit_code "
        strSelectSql = strSelectSql & " and a.Cust_ref = c.Cust_ref and a.amendment_No = c.amendment_No and a.Account_code=c.account_code and a.unit_code='" & gstrUNITID & "' "
        strSelectSql = strSelectSql & " and  c.Item_code = d.Item_code and a.Account_Code='" & Trim(pstrCustno) & "' and a.Cust_Ref='" & Trim(pstrRefNo)
        strSelectSql = strSelectSql & "' and a.Amendment_No='" & Trim(pstrAmmNo) & "' And c.Active_Flag = 'A' "
        strSelectSql = strSelectSql & " and c.Item_Code in (Select a.Item_code from Item_MSt a,Itembal_mst b where a.unit_code=b.unit_code and a.Item_Main_grp in (" & pstrItemin & ") and a.Item_code = b.Item_code and b.Location_code ='" & pstrstockLocation & "' and b.Cur_bal >0 and a.hold_flag =0 and a.Status = 'A' and a.unit_code='" & gstrUNITID & "'"
        If Len(Trim(pstrItemNotin)) > 0 Then
            strSelectSql = strSelectSql & " and a.Item_code not in ( " & pstrItemNotin & "))"
        Else
            strSelectSql = strSelectSql & ")"
        End If
        strSelectSql = strSelectSql & " AND C.ITEM_CODE NOT IN "
        strSelectSql = strSelectSql & " ( "
        strSelectSql = strSelectSql & " SELECT I.ITEM_CODE FROM ITEM_MST I INNER JOIN GLOBAL_TOOL_MST G "
        strSelectSql = strSelectSql & " ON I.ITEM_CODE = G.GLOBAL_TOOL_CODE_FOR_ITEM_MASTER WHERE I.UNIT_CODE = '" & gstrUNITID & "'"
        strSelectSql = strSelectSql & " AND ITEM_MAIN_GRP ='P' "
        strSelectSql = strSelectSql & " AND G.CATEGORY IN ('MOULDING','PRESSTOOL') AND I.STATUS='A'"
        strSelectSql = strSelectSql & " ) "
        MakeSelectSubQuery_ToolInvoice = strSelectSql

    End Function

    Public Function makeSelectSql_ToolInvoice(ByRef pstrCustno As String, ByRef pstrRefNo As String, ByRef pstrAmmNo As String, ByRef effectyrmon As String, ByRef Validyrmon As String, ByRef pstrstockLocation As String, ByRef strDate As String, ByRef pstrItemin As String, Optional ByRef pstrCondition As String = "") As String
        'ADDED BY VINOD
        Dim strSelectSql As String
        Dim strNextWorkDay As String
        Dim RsobjSchedules As New ADODB.Recordset
        Dim blnCalendarDateTrac As Boolean
        If RsobjSchedules.State = ADODB.ObjectStateEnum.adStateOpen Then RsobjSchedules.Close()
        RsobjSchedules.Open("SELECT DSWiseTracking,CalendarDateTrac FROM sales_parameter where unit_code='" & gstrUNITID & "'", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If Not RsobjSchedules.EOF Then
            If IIf(RsobjSchedules.Fields(1).Value, 1, 0) = 1 Then
                blnCalendarDateTrac = True
            Else
                blnCalendarDateTrac = False
            End If
        End If
        RsobjSchedules.Close()
        If blnCalendarDateTrac Then
            strNextWorkDay = GetNextWorkingDay(strDate)
            If strNextWorkDay = "-1" Then
                makeSelectSql_ToolInvoice = ""
                Exit Function
            End If
        End If
        If gblnGSTUnit = True Then
            strSelectSql = "Select b.Item_Code,c.Cust_DrgNo,c.Cust_Drg_Desc,d.hsn_sac_code from Cust_Ord_hdr a,MonthlyMktSchedule b,Cust_ord_dtl c,Item_Mst d where "
        Else
            strSelectSql = "Select b.Item_Code,c.Cust_DrgNo,c.Cust_Drg_Desc,d.Tariff_Code from Cust_Ord_hdr a,MonthlyMktSchedule b,Cust_ord_dtl c,Item_Mst d where "
        End If

        strSelectSql = strSelectSql & " a.unit_code = b.unit_code and b.unit_code = c.unit_code and c.unit_code = d.unit_code "
        strSelectSql = strSelectSql & " and a.Cust_ref = c.Cust_ref and a.amendment_No = c.amendment_No and a.Account_code=c.account_code And c.Active_Flag ='A' "
        strSelectSql = strSelectSql & " and a.account_code=b.Account_code and c.Cust_drgNo=b.Cust_drgNo and b.ITem_code = d.Item_code and a.Account_Code='" & Trim(pstrCustno) & "' and a.unit_code='" & gstrUNITID & "'"
        strSelectSql = strSelectSql & " and a.Cust_Ref='" & Trim(pstrRefNo) & "' and a.Amendment_No='" & Trim(pstrAmmNo) & "' and b.status = 1 and b.Schedule_flag =1 and b.Year_Month =  " & Validyrmon
        strSelectSql = strSelectSql & " and b.Item_Code in(Select a.Item_code from Item_MSt a,Itembal_mst b where a.unit_code=b.unit_code and a.unit_code='" & gstrUNITID & "' and a.Item_Main_grp in (" & Trim(pstrItemin) & ") and a.Item_code = b.Item_code and b.Location_code ='" & pstrstockLocation & "' and b.Cur_bal >0 and a.hold_flag =0 and a.Status = 'A'"
        If Len(Trim(pstrCondition)) > 0 Then
            strSelectSql = strSelectSql & " and a.Item_code not in(" & pstrCondition & "))"
        Else
            strSelectSql = strSelectSql & ")"
        End If
        strSelectSql = strSelectSql & " AND B.ITEM_CODE NOT IN "
        strSelectSql = strSelectSql & " ( "
        strSelectSql = strSelectSql & " SELECT I.ITEM_CODE FROM ITEM_MST I INNER JOIN GLOBAL_TOOL_MST G "
        strSelectSql = strSelectSql & " ON I.ITEM_CODE = G.GLOBAL_TOOL_CODE_FOR_ITEM_MASTER WHERE I.UNIT_CODE = '" & gstrUNITID & "'"
        strSelectSql = strSelectSql & " AND ITEM_MAIN_GRP ='P' "
        strSelectSql = strSelectSql & " AND G.CATEGORY IN ('MOULDING','PRESSTOOL') AND I.STATUS='A'"
        strSelectSql = strSelectSql & " ) "
        strSelectSql = strSelectSql & " UNION "
        If gblnGSTUnit = True Then
            strSelectSql = strSelectSql & " Select b.Item_Code,c.Cust_DrgNo,c.Cust_Drg_Desc,d.hsn_sac_code from Cust_Ord_hdr a,DailyMktSchedule b,Cust_ord_dtl c,ITem_Mst d  where "
        Else
            strSelectSql = strSelectSql & " Select b.Item_Code,c.Cust_DrgNo,c.Cust_Drg_Desc,d.Tariff_Code from Cust_Ord_hdr a,DailyMktSchedule b,Cust_ord_dtl c,ITem_Mst d  where "
        End If
        strSelectSql = strSelectSql & " a.unit_code = b.unit_code and b.unit_code = c.unit_code and c.unit_code = d.unit_code "
        strSelectSql = strSelectSql & " and a.Cust_ref = c.Cust_ref and a.amendment_No = c.amendment_No and a.Account_code=c.account_code"
        strSelectSql = strSelectSql & " and a.account_code=b.Account_code and c.Cust_drgNo=b.Cust_drgNo and b.ITem_code =d.ITem_code and b.status = 1 And c.Active_Flag ='A' and a.Account_Code='" & Trim(pstrCustno) & "' and a.unit_code='" & gstrUNITID & "' "
        strSelectSql = strSelectSql & " and a.Cust_Ref='" & Trim(pstrRefNo) & "' and a.Amendment_No='" & Trim(pstrAmmNo) & "' "
        If blnCalendarDateTrac Then
            If Month(ConvertToDate(strDate)) = Month(ConvertToDate(strNextWorkDay)) Then
                strSelectSql = strSelectSql & " and  datepart(mm,b.trans_date) = '" & Month(ConvertToDate(strDate)) & "' And  b.Trans_Date <= '" & VB6.Format(strNextWorkDay, "dd/mmm/yyyy") & "'  And DatePart(yyyy, b.Trans_Date) = '" & Year(ConvertToDate(strDate)) & "'"
            ElseIf (Month(ConvertToDate(strDate)) + 1) = Month(ConvertToDate(strNextWorkDay)) Then
                strSelectSql = strSelectSql & " and  datepart(mm,b.trans_date) in ('" & Month(ConvertToDate(strDate)) & "','" & Month(ConvertToDate(strNextWorkDay)) & "') And  b.Trans_Date <= '" & VB6.Format(strNextWorkDay, "dd/mmm/yyyy") & "'  And DatePart(yyyy, b.Trans_Date) = '" & Year(ConvertToDate(strDate)) & "'"
            ElseIf (Year(ConvertToDate(strDate)) + 1) = Year(ConvertToDate(strNextWorkDay)) Then
                strSelectSql = strSelectSql & " and  ((datepart(yyyy,b.Trans_Date) = '" & Year(ConvertToDate(strDate)) & "' and datepart(mm,b.Trans_Date)='" & Month(ConvertToDate(strDate)) & "') or ( datepart(yyyy,b.Trans_Date) = '" & Year(CDate(strNextWorkDay)) & "' and datepart(mm,b.Trans_Date)='" & Month(CDate(strNextWorkDay)) & "')) And  b.Trans_Date <= '" & VB6.Format(strNextWorkDay, "dd/mmm/yyyy") & "'"
            End If
        Else
            strSelectSql = strSelectSql & " and  datepart(mm,b.trans_date) = '" & Month(ConvertToDate(strDate)) & "' And  b.Trans_Date <= '" & getDateForDB(strDate) & "'  And DatePart(yyyy, b.Trans_Date) = '" & Year(ConvertToDate(strDate)) & "'"
        End If
        strSelectSql = strSelectSql & " and b.Item_Code in(Select a.Item_code from Item_MSt a,Itembal_mst b where a.unit_code = b.unit_code and a.Item_Main_grp in (" & Trim(pstrItemin) & ") and a.Item_code = b.Item_code and b.Location_code ='" & pstrstockLocation & "' and b.Cur_bal >0 and a.hold_flag =0 and a.Status = 'A' and a.unit_code='" & gstrUNITID & "'"
        If Len(Trim(pstrCondition)) > 0 Then
            strSelectSql = strSelectSql & " and a.Item_code not in( " & pstrCondition & "))"
        Else
            strSelectSql = strSelectSql & ")"
        End If
        strSelectSql = strSelectSql & " AND B.ITEM_CODE NOT IN "
        strSelectSql = strSelectSql & " ( "
        strSelectSql = strSelectSql & " SELECT I.ITEM_CODE FROM ITEM_MST I INNER JOIN GLOBAL_TOOL_MST G "
        strSelectSql = strSelectSql & " ON I.ITEM_CODE = G.GLOBAL_TOOL_CODE_FOR_ITEM_MASTER WHERE I.UNIT_CODE = '" & gstrUNITID & "'"
        strSelectSql = strSelectSql & " AND ITEM_MAIN_GRP ='P' "
        strSelectSql = strSelectSql & " AND G.CATEGORY IN ('MOULDING','PRESSTOOL') AND I.STATUS='A'"
        strSelectSql = strSelectSql & " ) "
        makeSelectSql_ToolInvoice = strSelectSql
    End Function

#End Region
    '101254587
    Private Function CreateSubQueryForGlobalToolItemCheck(ByVal strItemGroup As String, ByVal strVendorCode As String) As String
        Dim strSql As String = String.Empty
        If Len(Trim(strItemGroup)) > 0 Then
            If strItemGroup = "'M'" Then
                strSql = " AND"
                strSql = strSql & " ("
                strSql = strSql & " (a.Item_Main_Grp='M' AND a.Item_grp<>'CSMT') OR"
                strSql = strSql & " (a.Item_Main_Grp='M' AND a.Item_grp='CSMT' AND EXISTS (SELECT 1 FROM GLOBAL_TOOL_MST G WHERE G.GLOBAL_TOOL_CODE_FOR_ITEM_MASTER=a.ITEM_CODE AND G.MOULD_BELONGING ='CUSTOMER FUNDED' AND G.CATEGORY IN ('MOULDING','PRESSTOOL') AND ISNULL(G.INV_NO,0) = 0 AND G.TRANSFER_COMPLETED = 0 AND G.TRANSFER_TYPE IN ('SALE','CUSTOMER_TOOL_SALE') AND G.ISACTIVE = 1" & CreateVendorCondition(strVendorCode) & "))"
                strSql = strSql & " ) "
            ElseIf strItemGroup.Contains("'M'") Then
                strSql = " AND"
                strSql = strSql & " ("
                strSql = strSql & " (a.Item_Main_Grp IN (" & strItemGroup.Replace(",'M'", "").Replace("'M',", "") & ")) OR"
                strSql = strSql & " (a.Item_Main_Grp='M' AND a.Item_grp<>'CSMT') OR"
                strSql = strSql & " (a.Item_Main_Grp='M' AND a.Item_grp='CSMT' AND EXISTS (SELECT 1 FROM GLOBAL_TOOL_MST G WHERE G.GLOBAL_TOOL_CODE_FOR_ITEM_MASTER=a.ITEM_CODE AND G.MOULD_BELONGING ='CUSTOMER FUNDED' AND G.CATEGORY IN ('MOULDING','PRESSTOOL') AND ISNULL(G.INV_NO,0) = 0 AND G.TRANSFER_COMPLETED = 0 AND G.TRANSFER_TYPE IN ('SALE','CUSTOMER_TOOL_SALE') AND G.ISACTIVE = 1" & CreateVendorCondition(strVendorCode) & "))"
                strSql = strSql & " ) "
            ElseIf strItemGroup = "'P'" Then
                strSql = " AND"
                strSql = strSql & " ("
                strSql = strSql & " (a.Item_Main_Grp='P' AND a.Item_grp<>'TOOL') OR"
                strSql = strSql & " (a.Item_Main_Grp='P' AND a.Item_grp='TOOL' AND EXISTS (SELECT 1 FROM GLOBAL_TOOL_MST G WHERE G.GLOBAL_TOOL_CODE_FOR_ITEM_MASTER=a.ITEM_CODE AND G.MOULD_BELONGING IN ('MATE FUNDED','AMMORTIZED') AND G.CATEGORY IN ('MOULDING','PRESSTOOL') AND ISNULL(G.INV_NO,0) = 0 AND G.TRANSFER_COMPLETED = 0 AND G.TRANSFER_TYPE IN ('SALE','CUSTOMER_TOOL_SALE') AND G.ISACTIVE = 1" & CreateVendorCondition(strVendorCode) & "))"
                strSql = strSql & " ) "
            ElseIf strItemGroup.Contains("'P'") Then
                strSql = " AND"
                strSql = strSql & " ("
                strSql = strSql & " (a.Item_Main_Grp IN (" & strItemGroup.Replace(",'P'", "").Replace("'P',", "") & ")) OR"
                strSql = strSql & " (a.Item_Main_Grp='P' AND a.Item_grp<>'TOOL') OR"
                strSql = strSql & " (a.Item_Main_Grp='P' AND a.Item_grp='TOOL' AND EXISTS (SELECT 1 FROM GLOBAL_TOOL_MST G WHERE G.GLOBAL_TOOL_CODE_FOR_ITEM_MASTER=a.ITEM_CODE AND G.MOULD_BELONGING IN ('MATE FUNDED','AMMORTIZED') AND G.CATEGORY IN ('MOULDING','PRESSTOOL') AND ISNULL(G.INV_NO,0) = 0 AND G.TRANSFER_COMPLETED = 0 AND G.TRANSFER_TYPE IN ('SALE','CUSTOMER_TOOL_SALE') AND G.ISACTIVE = 1" & CreateVendorCondition(strVendorCode) & "))"
                strSql = strSql & " ) "
            End If
        End If
        Return strSql
    End Function
    Private Function CreateVendorCondition(ByVal strVendorCode As String) As String
        Dim strVendorSql As String = String.Empty
        If Len(Trim(strVendorCode)) > 0 Then
            strVendorSql = " AND G.VENDOR_CODE='" & strVendorCode & "' "
        End If
        Return strVendorSql
    End Function
End Class