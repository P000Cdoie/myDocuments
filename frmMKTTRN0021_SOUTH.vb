Option Strict Off
Option Explicit On
Friend Class frmMKTTRN0021_SOUTH
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
	'===================================================================================
	'Revised By R.Arul mozhi varman
	'Date : 29-04-2004
	'Reason : To raise the rejection invoice for customer supplied material while reject the material in Line rejection
	'=====================================================================================
	'Code changed by Arul on 24-08-2004
	'Error coming in "AddDataFromGrinDtl" sub procedure
	'======================================================================================
	'Code Changed by Arul on 07-03-2005
	'Reason : Multiple Sales Order selection in One Invoice
	'=======================================================================================
	'Changed by Arul on 15-06-2005
	'Changed for "ODBC driver not supported the requested properties" Error
	'=======================================================================================
	'Revised By     : Arul mozhi varman
	'Revised On     : 22-11-2005
	'Revised Reason : To fetch the items for Sample Export invoice
	'=======================================================================================
	'Revised By     : Arul mozhi varman
	'Revised On     : 19-12-2005
	'Revised Reason : To avail Sales order functionality for Transfer invoice (Finished Parts)
	'=======================================================================================
	'Revised By     : Arul mozhi varman
	'Revised On     : 28-12-2005
	'Revised Reason : To avail Sales order functionality for Transfer invoice (Inputs)
	'=======================================================================================
	'Revised By     : Arul mozhi varman
	'Revised On     : 18-04-2006
	'Revised Reason : To check the item status flag in rejection invoice select items
	'=======================================================================================
	'Revised By     : Ashutosh Verma
	'Revised On     : 22-01-2007 ,Issue ID:19352
	'Revised Reason : Consider Current month schedule while Invoicing.
	'=======================================================================================
    'Revised By        -    Vinod Singh
    'Revision Date     -    19/05/2011
    'Revision History  -    Changes for Multi Unit
    '=======================================================================================
    'Revised By        -    Prashant Rajpal
    'Revision Date     -    26/11/2012
    'Issue id          -    10307080 
    'Revision History  -    Required the option for SF code items sample invoicing for south form
    '=======================================================================================
    'Revised By        -    Prashant Rajpal
    'Revision Date     -    20/08/2013
    'Issue id          -    10229989 
    'Revision History  -    Multiple sales order for hyundai
    '=======================================================================================
    'REVISED BY     :  VINOD SINGH
    'REVISED DATE   :  26 SEP 2013
    'ISSUE ID       :  10378778
    'PURPOSE        :  GLOBAL TOOL CHAGES
    '=======================================================================================
    'REVISED BY     -  PRASHANT RAJPAL
    'REVISED ON     -  18 SEP 2015
    'PURPOSE        -  10869290 -SERVICE INVOICE 

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
    Dim mCtlHdrSchdate As System.Windows.Forms.ColumnHeader
    Dim mCtlHdrCustomerSoNo As System.Windows.Forms.ColumnHeader
    Dim mCtlHdrCustomerAmendmentNo As System.Windows.Forms.ColumnHeader
    Dim mblnlinelevelcustomer As Boolean
    Dim mstrCust_code As String
    Dim mstrshop_code As String
    Dim mstrinvoicetype As String
    Dim mstrinvoicesubtype As String
    Dim mintTOTALALREADYITEMINGRID As String
    Dim mblnPCACUSTOMER As Boolean
    Dim mintnoofbackdaysschedule_PCA As Integer



	Private Sub CmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdCancel.Click
		On Error GoTo ErrHandler
		Me.Close()
		Exit Sub
ErrHandler: 'The Error Handling Code Starts here
		Call gobjError.RaiseError(Err.Number, ERR.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Public Property SHOP_CODE() As String
        Get
            SHOP_CODE = mstrshop_code
        End Get
        Set(ByVal Value As String)
            mstrshop_code = Value
        End Set
    End Property
    Public Property TOTALALREADYITEMINGRID() As Integer
        Get
            TOTALALREADYITEMINGRID = mintTOTALALREADYITEMINGRID
        End Get
        Set(ByVal value As Integer)
            mintTOTALALREADYITEMINGRID = value
        End Set
    End Property
    Public Property Cust_Code() As String
        Get
            Cust_Code = mstrCust_code
        End Get
        Set(ByVal Value As String)
            mstrCust_code = Value
        End Set
    End Property
    Public Property Invoice_type() As String
        Get
            Invoice_type = mstrinvoicetype
        End Get
        Set(ByVal Value As String)
            mstrinvoicetype = Value
        End Set
    End Property
    Public Property Invoice_Subtype() As String
        Get
            Invoice_Subtype = mstrinvoicesubtype
        End Get
        Set(ByVal Value As String)
            mstrinvoicesubtype = Value
        End Set
    End Property
    Public Property ISPCACUSTOMER() As String
        Get
            ISPCACUSTOMER = mblnPCACUSTOMER
        End Get
        Set(ByVal Value As String)
            mblnPCACUSTOMER = Value
        End Set
    End Property
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
		'Code Modified By   -   Nitin Sood
		'No of Items Selected in Challan can be Till 7
		'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
        On Error GoTo ErrHandler
        Dim strSqlct2qry As String
        If Not ((gstrUNITID = "MST") Or (gstrUNITID = "MAR") Or (gstrUNITID = "M03") Or (gstrUNITID = "M02") Or (gstrUNITID = "M3W") Or (gstrUNITID = "M3T") Or (gstrUNITID = "SMK")) Then
            If mblnlinelevelcustomer = True Then
                If mintTOTALALREADYITEMINGRID + Me.lvwItemCode.CheckedItems.Count > 4 Then
                    MsgBox("Only Four Item are allowed To Be Selected In The List.", MsgBoxStyle.Information, ResolveResString(100))
                    Exit Sub
                End If
            End If
        End If
        ''Added by priti on 25th August 2025 INC1320314
        Dim AllowMaximumItemInInvoice As Boolean = Convert.ToBoolean(SqlConnectionclass.ExecuteScalar("select AllowMaximumItemInInvoice from Customer_mst where customer_code='" & Trim(Cust_Code) & "' and unit_code='" & gstrUNITID & "'"))
        If AllowMaximumItemInInvoice = True Then
            Dim intMaxitem As Integer = SqlConnectionclass.ExecuteScalar("select MaximumItemInInvoice from Customer_mst where customer_code='" & Trim(Cust_Code) & "' and unit_code='" & gstrUNITID & "'")
            If intMaxitem > 0 Then
                If mblnlinelevelcustomer = True Then
                    If mintTOTALALREADYITEMINGRID + Me.lvwItemCode.CheckedItems.Count > intMaxitem Then
                        MsgBox("No. Of Items Selected Should not be greater than " & intMaxitem, MsgBoxStyle.Information, "empower")
                        Exit Sub
                    End If
                End If
            End If
        End If
        ''code ends by priti on 25th August 2025 INC1320314

        mstrItemText = "" : intCheckCounter = intIteminSp
        Dim intSubItem As Short
        Dim mObjDB As ClsResultSetDB

        If mblnPCACUSTOMER = True And UCase(mstrInvType) = ("NORMAL INVOICE") And UCase(mstrInvSubType) = ("FINISHED GOODS") Then
            strSqlct2qry = "DELETE FROM TMP_PCA_ITEMSELECTION_PACKAGECODE WHERE IP_ADDRESS='" & gstrIpaddressWinSck & "'"
            SqlConnectionclass.ExecuteNonQuery(strSqlct2qry)

            For intSubItem = 0 To lvwItemCode.Items.Count - 1

                If Me.lvwItemCode.Items.Item(intSubItem).Checked = True Then

                    strSqlct2qry = "INSERT INTO TMP_PCA_ITEMSELECTION_PACKAGECODE (UNIT_CODE,ITEM_CODE,PART_CODE,PACKAGECODE,MESSAGECALLOFFNO,QUANTITY,IP_ADDRESS,MAXPOSSQTY)"
                    strSqlct2qry += " VALUES ('" & gstrUNITID & "','" & lvwItemCode.Items.Item(intSubItem).SubItems(0).Text & "'"
                    strSqlct2qry += ",'" & lvwItemCode.Items.Item(intSubItem).SubItems(1).Text & "'" & ",'" & lvwItemCode.Items.Item(intSubItem).SubItems(5).Text & "',"
                    strSqlct2qry += " '" & lvwItemCode.Items.Item(intSubItem).SubItems(4).Text & "','" & lvwItemCode.Items.Item(intSubItem).SubItems(6).Text & "','" & gstrIpaddressWinSck & "','" & lvwItemCode.Items.Item(intSubItem).SubItems(6).Text & "')"
                End If

                If strSqlct2qry.Length > 0 Then
                    SqlConnectionclass.ExecuteNonQuery(strSqlct2qry)
                End If

                strSqlct2qry = ""
            Next intSubItem


        End if
        For intSubItem = 0 To lvwItemCode.Items.Count - 1
            If Me.lvwItemCode.Items.Item(intSubItem).Checked = True Then
                intCheckCounter = intCheckCounter + 1
                If blnExpinv = False Then
                    If intCheckCounter > 1000 Then
                        'Call ConfirmWindow(10415, BUTTON_OK)
                        MsgBox("No. Of Items Selected Should be Less than 1000", MsgBoxStyle.Information, "empower")
                        mstrItemText = ""
                        Exit Sub
                    End If
                    'added by Nisha On 15/07/2002
                Else
                    mObjDB = New ClsResultSetDB
                    mObjDB.GetResult("Select EOU_Flag from Company_Mst where unit_code='" & gstrUNITID & "'")
                    If mObjDB.GetValue("EOU_Flag") = False Then
                        If intCheckCounter > 1000 Then
                            'Call ConfirmWindow(10415, BUTTON_OK)
                            MsgBox("No. Of Items Selected Should be Less than 1000", MsgBoxStyle.Information, "empower")
                            mstrItemText = ""
                            Exit Sub
                        End If
                    End If
                    mObjDB.ResultSetClose()
                    mObjDB = Nothing
                End If
                mstrItemText = mstrItemText & "'" & Trim(Me.lvwItemCode.Items.Item(intSubItem).SubItems(1).Text) & "',"
            End If
        Next intSubItem
        If Len(mstrItemText) = 0 Then
            Call ConfirmWindow(10418, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
            Me.lvwItemCode.Focus()
            Exit Sub
        End If
        Me.Close()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub frmMKTTRN0021_SOUTH_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo ErrHandler

        ''''''''SetBackGroundColorNew(Me, True)
        '   Call AddColumnsInListView()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub AddColumnsInListView()

        '***********************************
        'To add Columns Headers in the ListView in the form load
        '***********************************
        On Error GoTo ErrHandler
        If Len(Trim(Invoice_type)) > 0 Then
            If UCase(Invoice_type) <> "REJECTION" Then
                mblnlinelevelcustomer = Find_Value("SELECT SOUPLD_LINE_LEVEL_SALESORDER FROM CUSTOMER_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & Trim(Cust_Code) & "'")
                mblnPCACUSTOMER = Find_Value("SELECT PCA_CUSTOMER FROM CUSTOMER_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & Trim(Cust_Code) & "'")
                mintnoofbackdaysschedule_PCA = Find_Value("SELECT isnull(NOOFBACKDAYS_PCA_SCHEDULE,0) FROM CUSTOMER_MST WHERE UNIT_CODE='" & gstrUNITID & "' AND CUSTOMER_CODE='" & Trim(Cust_Code) & "'")
            End If
        End If
        With Me.lvwItemCode
            mCtlHdrItemCode = .Columns.Add("")
            If (UCase(mstrInvType) = ("TRANSFER INVOICE") Or UCase(mstrInvType) = ("INTER-DIVISION")) And UCase(mstrInvSubType) = ("FINISHED GOODS") Then
                mCtlHdrItemCode.Text = "Drawing No."
            Else
                mCtlHdrItemCode.Text = "Item Code"
            End If
            mCtlHdrItemCode.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(lvwItemCode.Width) / 4)
            mCtlHdrDrawingNo = .Columns.Add("")
            If (UCase(mstrInvType) = ("TRANSFER INVOICE") Or UCase(mstrInvType) = ("INTER-DIVISION")) And UCase(mstrInvSubType) = ("FINISHED GOODS") Then
                mCtlHdrDrawingNo.Text = "Item Code"
            Else
                mCtlHdrDrawingNo.Text = "Drawing No."
            End If
            mCtlHdrDrawingNo.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(lvwItemCode.Width) / 4)
            mCtlHdrDescription = .Columns.Add("")
            mCtlHdrDescription.Text = "Description"
            mCtlHdrDescription.Width = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(lvwItemCode.Width) / 4))
            mCtlHdrDescription = .Columns.Add("")
            mCtlHdrDescription.Text = "Tariff Code"
            mCtlHdrDescription.Width = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(lvwItemCode.Width) / 4) - 100)



            If mblnlinelevelcustomer = True And UCase(Invoice_type) = ("NORMAL INVOICE") And UCase(Invoice_Subtype) = ("FINISHED GOODS") Then

                mCtlHdrCustomerSoNo = .Columns.Add("")
                mCtlHdrCustomerSoNo.Text = "CUSTOMER SALES ORDER NO"
                mCtlHdrCustomerSoNo.Width = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(lvwItemCode.Width) / 8))

                mCtlHdrCustomerAmendmentNo = .Columns.Add("")
                mCtlHdrCustomerAmendmentNo.Text = "AMEND. NO"
                mCtlHdrCustomerAmendmentNo.Width = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(lvwItemCode.Width) / 8))

                mCtlHdrSchdate = .Columns.Add("")
                mCtlHdrSchdate.Text = "STOCK "
                mCtlHdrSchdate.Width = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(lvwItemCode.Width) / 8))

                mCtlHdrSchdate = .Columns.Add("")
                mCtlHdrSchdate.Text = "SHOP CODE"
                mCtlHdrSchdate.Width = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(lvwItemCode.Width) / 8))

            End If

            If mblnPCACUSTOMER = True And UCase(Invoice_type) <> "REJECTION" Then
                mCtlHdrCustomerSoNo = .Columns.Add("")
                mCtlHdrCustomerSoNo.Text = "CALLOFF NO"
                mCtlHdrCustomerSoNo.Width = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(lvwItemCode.Width) / 8))

                mCtlHdrCustomerAmendmentNo = .Columns.Add("")
                mCtlHdrCustomerAmendmentNo.Text = "PACKAGE CODE"
                mCtlHdrCustomerAmendmentNo.Width = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(lvwItemCode.Width) / 8))

                mCtlHdrCustomerAmendmentNo = .Columns.Add("")
                mCtlHdrCustomerAmendmentNo.Text = "QUANTITY"
                mCtlHdrCustomerAmendmentNo.Width = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(lvwItemCode.Width) / 8))

            End If

        End With
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Public Function SelectDataFromCustOrd_Dtl(ByRef pstrCustno As String, ByRef pstrRefNo As String, ByRef pstrAmmNo As String, ByRef pstrSubType As String, ByRef pstrInvType As String, ByRef pstrstockLocation As String, Optional ByRef pstrCondition As String = "", Optional ByRef intAlreadyItem As Short = 0, Optional ByRef pstrConsCode As String = "") As String
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
        Call AddColumnsInListView()
        Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(mdifrmMain.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2)
        'Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(mdifrmMain.Width) - VB6.PixelsToTwipsX(frmModules.Width)) / 2.3)
        optPartNo.Checked = True
        lvwItemCode.FullRowSelect = True
        'for item selection more then one 4 in case of Export invoice
        intIteminSp = intAlreadyItem
        If pstrInvType = "EXPORT INVOICE" Then
            blnExpinv = True
        Else
            blnExpinv = False
        End If
        strDate = VB6.Format(GetServerDate, gstrDateFormat)
        Me.lvwItemCode.Items.Clear() 'initially clear all items in the listview
        strSelectSql = "Select effectMon=convert(char(2),month(effect_date)),effectYr=convert(char(4),Year(effect_date)),"
        strSelectSql = strSelectSql & " validMon=convert(char(2),month(Valid_date)),validYr=convert(char(4),year(Valid_date))"
        strSelectSql = strSelectSql & " from Cust_Ord_hdr where "
        strSelectSql = strSelectSql & " unit_code='" & gstrUNITID & "' and Account_Code='" & Trim(pstrCustno) & "' and Cust_Ref in('" & Trim(pstrRefNo) & "')"
        strSelectSql = strSelectSql & " and Active_Flag = 'A'"
        ''Changes done by Ashutosh on 16 Apr 2007, Issue Id:19731
        If Len(Trim(pstrConsCode)) > 0 Then
            strSelectSql = strSelectSql & " and consignee_code='" & Trim(pstrConsCode) & "' "
        End If
        ''Changes for Issue id:19731 end here.
        rsCustOrdHdr = New ClsResultSetDB
        rsCustOrdHdr.GetResult(strSelectSql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        strSelectSql = ""
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
            'effectMon = rsCustOrdhdr.GetValue("EffectMon")
            effectyrmon = rsCustOrdHdr.GetValue("effectYr") & effectMon
            mstrInvType = pstrInvType : mstrInvSubType = pstrSubType
            Select Case UCase(pstrInvType)
                'Code Changed by Arul on 19-12-2005 to add the Transfer invoice option
                Case "NORMAL INVOICE", "EXPORT INVOICE", "TRANSFER INVOICE", "INTER-DIVISION"
                    Select Case UCase(pstrSubType)
                        Case "FINISHED GOODS"
                            strSelectSql = makeSelectSql(pstrCustno, pstrRefNo, pstrAmmNo, effectyrmon, Validyrmon, pstrstockLocation, strDate, "'F','S'", pstrCondition)
                        Case "COMPONENTS"
                            strSelectSql = MakeSelectSubQuery(pstrCustno, pstrRefNo, pstrAmmNo, pstrstockLocation, "'C'", pstrCondition)
                        Case "RAW MATERIAL"
                            strSelectSql = MakeSelectSubQuery(pstrCustno, pstrRefNo, pstrAmmNo, pstrstockLocation, "'R','S','B','M'", pstrCondition)
                        Case "ASSETS"
                            'strSelectSql = MakeSelectSubQuery(pstrCustno, pstrRefNo, pstrAmmNo, pstrstockLocation, "'P'", pstrCondition)
                            strSelectSql = MakeSelectSubQuery_Asset(pstrCustno, pstrRefNo, pstrAmmNo, pstrstockLocation, "'P'", pstrCondition)
                            'strSelectSql = "SELECT ITEM_CODE,CUST_DRGNO,CUST_DRG_DESC,TARIFF_CODE FROM UFN_INVOICE_ITEM_HELP('" & gstrUNITID & "','ASSET','" & pstrCustno & "','" & pstrRefNo & "','" & pstrCondition & "','P','" & pstrstockLocation & "') "
                        Case "TRADING GOODS"
                            strSelectSql = makeSelectSql(pstrCustno, pstrRefNo, pstrAmmNo, effectyrmon, Validyrmon, pstrstockLocation, strDate, "'T'", pstrCondition)
                            '            Case "SCRAP"
                            '                strSelectSql = MakeSelectSubQuery(pstrCustno, pstrRefNo, pstrAmmNo, pstrstockLocation, "'R','C'")
                        Case "TOOLS & DIES"
                            'strSelectSql = MakeSelectSubQuery(pstrCustno, pstrRefNo, pstrAmmNo, pstrstockLocation, "'P','A'", pstrCondition)
                            strSelectSql = MakeSelectSubQuery_Tool(pstrCustno, pstrRefNo, pstrAmmNo, pstrstockLocation, "'P','A'", pstrCondition)
                            'strSelectSql = "SELECT ITEM_CODE,CUST_DRGNO,CUST_DRG_DESC,TARIFF_CODE FROM UFN_INVOICE_ITEM_HELP('" & gstrUNITID & "','TOOL','" & pstrCustno & "','" & pstrRefNo & "','" & pstrCondition & "','P,A','" & pstrstockLocation & "') "
                            'Code Added By Arul on 22-11-2005 for SaPle Export imvoice
                            'Case "EXPORTS"
                        Case "EXPORTS", "SAMPLE"
                            'Changes Ends here
                            strSelectSql = makeSelectSql(pstrCustno, pstrRefNo, pstrAmmNo, effectyrmon, Validyrmon, pstrstockLocation, strDate, "'F','S'", pstrCondition, pstrConsCode)
                            'Code added by Arul on 28-12-2005
                        Case "ALL"
                            'Changes Ends here
                            strSelectSql = makeSelectSql(pstrCustno, pstrRefNo, pstrAmmNo, effectyrmon, Validyrmon, pstrstockLocation, strDate, "'F','S','C','R','B','M','N','A','T','P'", pstrCondition, pstrConsCode)
                            'Code added by Arul on 28-12-2005
                        Case "INPUTS"
                            strSelectSql = MakeSelectSubQuery(pstrCustno, pstrRefNo, pstrAmmNo, pstrstockLocation, "'C','R','S','B','M','N'", pstrCondition)
                            'Addition ends here

                    End Select
                Case "JOBWORK INVOICE"
                    'mP_Connection.Execute "set dateformat 'dmy'"
                    strSelectSql = makeSelectSql(pstrCustno, pstrRefNo, pstrAmmNo, effectyrmon, Validyrmon, pstrstockLocation, strDate, "'F'", pstrCondition)
                Case "SERVICE INVOICE"
                    'mP_Connection.Execute "set dateformat 'dmy'"
                    strSelectSql = makeSelectSql_service(pstrCustno, pstrRefNo, pstrAmmNo, effectyrmon, Validyrmon, pstrstockLocation, strDate, "'M'", pstrCondition)
            End Select
        Else
            rsCustOrdHdr.ResultSetClose()
            rsCustOrdHdr = Nothing
            strSelectSql = "Select effect_date,"
            strSelectSql = strSelectSql & " Valid_date"
            strSelectSql = strSelectSql & " from Cust_Ord_hdr where "
            strSelectSql = strSelectSql & " unit_code='" & gstrUNITID & "' and Account_Code='" & Trim(pstrCustno) & "' and Cust_Ref='" & Trim(pstrRefNo) & "'"
            strSelectSql = strSelectSql & " and Amendment_No='" & Trim(pstrAmmNo) & "' and Active_flag ='A'"
            rsCustOrdHdr = New ClsResultSetDB
            rsCustOrdHdr.GetResult(strSelectSql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            If rsCustOrdHdr.GetNoRows > 0 Then
                Validyrmon = rsCustOrdHdr.GetValue("valid_date")
                effectyrmon = rsCustOrdHdr.GetValue("Effect_date")
            End If
            Select Case pstrSubType
                Case "COMPONENTS"
                    strSelectSql = makeSelectSql(pstrCustno, pstrRefNo, pstrAmmNo, effectyrmon, Validyrmon, pstrstockLocation, strDate, "'C'", pstrCondition)
                Case "TRADING GOODS"
                    strSelectSql = makeSelectSql(pstrCustno, pstrRefNo, pstrAmmNo, effectyrmon, Validyrmon, pstrstockLocation, strDate, "'T'", pstrCondition)
                Case "ASSETS"
                    'strSelectSql = makeSelectSql(pstrCustno, pstrRefNo, pstrAmmNo, effectyrmon, Validyrmon, pstrstockLocation, strDate, "'P'", pstrCondition)
                    strSelectSql = makeSelectSql_Asset(pstrCustno, pstrRefNo, pstrAmmNo, effectyrmon, Validyrmon, pstrstockLocation, strDate, "'P'", pstrCondition)
                    'strSelectSql = "select * from UFN_INVOICE_ITEM_HELP_SCHEDULE_WISE('" & gstrUNITID & "','ASSET','" & pstrCustno & "','" & pstrRefNo & "','" & pstrAmmNo & "','" & effectyrmon & "'," & Validyrmon & ",'" & pstrstockLocation & "','" & getDateForDB(strDate) & "','P','" & pstrCondition & "','')"
                Case "TOOLS & DIES"
                    'strSelectSql = makeSelectSql(pstrCustno, pstrRefNo, pstrAmmNo, effectyrmon, Validyrmon, pstrstockLocation, strDate, "'A','P'", pstrCondition)
                    strSelectSql = makeSelectSql_Tool(pstrCustno, pstrRefNo, pstrAmmNo, effectyrmon, Validyrmon, pstrstockLocation, strDate, "'A','P'", pstrCondition)
                    'strSelectSql = "select * from UFN_INVOICE_ITEM_HELP_SCHEDULE_WISE('" & gstrUNITID & "','TOOL','" & pstrCustno & "','" & pstrRefNo & "','" & pstrAmmNo & "','" & effectyrmon & "','" & Validyrmon & "','" & pstrstockLocation & "','" & getDateForDB(strDate) & "','A','P','" & pstrCondition & "','')"
                Case "RAW MATERIAL"
                    strSelectSql = makeSelectSql(pstrCustno, pstrRefNo, pstrAmmNo, effectyrmon, Validyrmon, pstrstockLocation, strDate, "'R','S','B','M'", pstrCondition)
                Case "SCRAP"
                    strSelectSql = makeSelectSql(pstrCustno, pstrRefNo, pstrAmmNo, effectyrmon, Validyrmon, pstrstockLocation, strDate, "'R','C'", pstrCondition)
            End Select
        End If
        'Change By Virendra Gupta (Incase strSelectSql is empty)
        If strSelectSql = "" Then
            Exit Function
        End If
        'End
        rsCustOrdDtl = New ClsResultSetDB
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
                        mListItemUserId.SubItems(3).Text = rsCustOrdDtl.GetValue("hsn_sac_code")
                    Else
                        mListItemUserId.SubItems(3).Text = rsCustOrdDtl.GetValue("Tariff_Code")
                    End If

                Else
                    If gblnGSTUnit = True Then
                        mListItemUserId.SubItems.Insert(3, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsCustOrdDtl.GetValue("hsn_sac_code")))
                    Else
                        mListItemUserId.SubItems.Insert(3, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsCustOrdDtl.GetValue("Tariff_Code")))
                    End If

                End If
                If mblnlinelevelcustomer = True And UCase(Invoice_type) = ("NORMAL INVOICE") And UCase(Invoice_Subtype) = ("FINISHED GOODS") Then
                    If mListItemUserId.SubItems.Count > 4 Then
                        mListItemUserId.SubItems(4).Text = rsCustOrdDtl.GetValue("EXTERNAL_SALESORDER_NO")
                    Else
                        mListItemUserId.SubItems.Insert(4, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsCustOrdDtl.GetValue("EXTERNAL_SALESORDER_NO")))
                    End If

                    If mListItemUserId.SubItems.Count > 5 Then
                        mListItemUserId.SubItems(5).Text = rsCustOrdDtl.GetValue("AMENDMENT_NO")
                    Else
                        mListItemUserId.SubItems.Insert(5, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsCustOrdDtl.GetValue("AMENDMENT_NO")))
                    End If

                    If mListItemUserId.SubItems.Count > 6 Then
                        mListItemUserId.SubItems(6).Text = rsCustOrdDtl.GetValue("CUR_BAL")
                    Else
                        mListItemUserId.SubItems.Insert(6, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsCustOrdDtl.GetValue("CUR_BAL")))
                    End If

                    If mListItemUserId.SubItems.Count > 7 Then
                        mListItemUserId.SubItems(7).Text = rsCustOrdDtl.GetValue("shop_code")
                    Else
                        mListItemUserId.SubItems.Insert(7, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsCustOrdDtl.GetValue("shop_code")))
                    End If
                End If

                
                If mblnPCACUSTOMER = True And UCase(Invoice_type) = ("NORMAL INVOICE") And UCase(Invoice_Subtype) = ("FINISHED GOODS") Then
                    If mListItemUserId.SubItems.Count > 4 Then
                        mListItemUserId.SubItems(4).Text = rsCustOrdDtl.GetValue("MESSAGECALLOFFNO")
                    Else
                        mListItemUserId.SubItems.Insert(4, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsCustOrdDtl.GetValue("MESSAGECALLOFFNO")))
                    End If
                    If mListItemUserId.SubItems.Count > 5 Then
                        mListItemUserId.SubItems(5).Text = rsCustOrdDtl.GetValue("PACKAGECODE")
                    Else
                        mListItemUserId.SubItems.Insert(5, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsCustOrdDtl.GetValue("PACKAGECODE")))
                    End If
                    If mListItemUserId.SubItems.Count > 6 Then
                        mListItemUserId.SubItems(6).Text = rsCustOrdDtl.GetValue("QUANTITY")
                    Else
                        mListItemUserId.SubItems.Insert(6, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsCustOrdDtl.GetValue("QUANTITY")))
                    End If

                End If


                rsCustOrdDtl.MoveNext() 'move to next record
            Next intCount
            rsCustOrdDtl.ResultSetClose()
            rsCustOrdDtl = Nothing
        Else
            '10869290
            If UCase(pstrInvType) = "SERVICE INVOICE" Then
                MsgBox("No Items for selected Invoice in Sales Order.Please Check Following :" & vbCrLf & "1. Item in SO are Active and Not on Hold." & vbCrLf & "2. Check Marketing Schedule in Case of Finished\Trading Goods in SO.", MsgBoxStyle.Information, "empower")
            Else
                MsgBox("No Items for selected Invoice in Sales Order.Please Check Following :" & vbCrLf & "1. Item in SO are Active and Not on Hold." & vbCrLf & "2. Check Balance of Items for location " & pstrstockLocation & "." & vbCrLf & "3. Check Marketing Schedule in Case of Finished\Trading Goods in SO.", MsgBoxStyle.Information, "empower")
            End If
            '10869290
            Exit Function
        End If

        Me.ShowDialog()
        SelectDataFromCustOrd_Dtl = mstrItemText
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    'Private Sub lvwItemCode_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Private Sub frmMKTTRN0021_SOUTH_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error GoTo ErrHandler
        Me.Dispose()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub lvwItemCode_ItemChecked(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckedEventArgs) Handles lvwItemCode.ItemChecked
        Dim Item As System.Windows.Forms.ListViewItem = lvwItemCode.Items(e.Item.Index)
        Dim intSubItem As Short
        Dim strChapterCode As String
        Dim strShopCode As String
        strChapterCode = ""
        strShopCode = ""
        If gblnGSTUnit = False Then
            For intSubItem = 0 To lvwItemCode.Items.Count - 1
                If Me.lvwItemCode.Items.Item(intSubItem).Checked = True Then
                    If Len(Trim(strChapterCode)) = False Then
                        strChapterCode = lvwItemCode.Items.Item(intSubItem).SubItems(3).Text
                    Else
                        If StrComp(strChapterCode, lvwItemCode.Items.Item(intSubItem).SubItems(3).Text, CompareMethod.Text) <> 0 Then
                            MsgBox("Select Items of Same Tariff Code", MsgBoxStyle.Information, "empower")
                            lvwItemCode.Items.Item(e.Item.Index).Checked = False
                            lvwItemCode.Items.Item(intSubItem).Selected = True
                            Me.CmdOk.Focus()
                            Exit Sub
                        End If
                    End If
                End If
            Next intSubItem
        End If

        
        If mblnlinelevelcustomer = True And UCase(mstrInvType) = ("NORMAL INVOICE") And UCase(mstrInvSubType) = ("FINISHED GOODS") Then
            For intSubItem = 0 To lvwItemCode.Items.Count - 1
                If Me.lvwItemCode.Items.Item(intSubItem).Checked = True Then
                    If Len(Trim(strShopCode)) = False Then
                        strShopCode = lvwItemCode.Items.Item(intSubItem).SubItems(7).Text
                    Else
                        If StrComp(strShopCode, lvwItemCode.Items.Item(intSubItem).SubItems(7).Text, CompareMethod.Text) <> 0 Then
                            MsgBox("Select Items of different Shop Code", MsgBoxStyle.Information, "empower")
                            lvwItemCode.Items.Item(e.Item.Index).Checked = False
                            lvwItemCode.Items.Item(intSubItem).Selected = True
                            Me.CmdOk.Focus()
                            Exit Sub
                        End If
                    End If
                End If
            Next intSubItem

            For intSubItem = 0 To lvwItemCode.Items.Count - 1
                If Me.lvwItemCode.Items.Item(intSubItem).Checked = True Then
                    If Len(mstrshop_code.Trim) > 0 Then
                        If StrComp(mstrshop_code, lvwItemCode.Items.Item(intSubItem).SubItems(7).Text, CompareMethod.Text) <> 0 Then
                            MsgBox("Select Items of different Shop Code", MsgBoxStyle.Information, "empower")
                            lvwItemCode.Items.Item(e.Item.Index).Checked = False
                            lvwItemCode.Items.Item(intSubItem).Selected = True
                            Me.CmdOk.Focus()
                            Exit Sub
                        End If
                    End If
                End If
            Next intSubItem


        End If

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
        Call AddColumnsInListView()
        Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(mdifrmMain.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2)
        'Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(mdifrmMain.Width) - VB6.PixelsToTwipsX(frmModules.Width)) / 2.3)
        optPartNo.Checked = True
        lvwItemCode.FullRowSelect = True
        'for item selection more then one 4 in case of Export invoice
        intIteminSp = intAlreadyItem
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
                        strItembal = strItembal & " where a.unit_code=b.unit_code and a.Item_Code=b.Item_Code and a.Item_Main_Grp ='T'"
                        strItembal = strItembal & " and cur_bal >0 and a.Status ='A' and a.Hold_Flag <> 1 and a.unit_code='" & gstrUNITID & "'"
                        If Len(Trim(pstrItemNotin)) > 0 Then
                            strItembal = strItembal & " and a.Item_code not in (" & pstrItemNotin & ")"
                        End If
                    Case "ASSETS"
                        If gblnGSTUnit = True Then
                            strItembal = "SELECT Distinct(a.Item_Code),a.description,a.HSN_SAC_CODE FROM Item_Mst a,Itembal_Mst b"
                        Else
                            strItembal = "SELECT Distinct(a.Item_Code),a.description,a.Tariff_code FROM Item_Mst a,Itembal_Mst b"
                        End If

                        strItembal = strItembal & " where a.unit_code = b.unit_code and a.Item_Code=b.Item_Code and a.Item_Main_Grp ='P'"
                        strItembal = strItembal & " and cur_bal >0 and a.Status ='A' and a.Hold_Flag <> 1 and a.unit_code='" & gstrUnitId & "'"
                        '101254587
                        strItembal = strItembal & CreateSubQueryForGlobalToolItemCheck("'P'", pstrAccountCode)
                        If Len(Trim(pstrItemNotin)) > 0 Then
                            strItembal = strItembal & " and a.Item_code not in (" & pstrItemNotin & ")"
                        End If
                    Case "TOOLS & DIES"
                        If gblnGSTUnit = True Then
                            strItembal = "SELECT Distinct(a.Item_Code),a.description,a.HSN_SAC_CODE FROM Item_Mst a,Itembal_Mst b"
                        Else
                            strItembal = "SELECT Distinct(a.Item_Code),a.description,a.Tariff_code FROM Item_Mst a,Itembal_Mst b"
                        End If
                        strItembal = strItembal & " where a.unit_code = b.unit_code and a.Item_Code=b.Item_Code and a.Item_Main_Grp in('P','A')"
                        strItembal = strItembal & " and cur_bal >0 and a.Status ='A' and a.Hold_Flag <> 1 and a.unit_code='" & gstrUNITID & "'"
                        If Len(Trim(pstrItemNotin)) > 0 Then
                            strItembal = strItembal & " and a.Item_code not in (" & pstrItemNotin & ")"
                        End If
                    Case "RAW MATERIAL"
                        If gblnGSTUnit = True Then
                            strItembal = "SELECT Distinct(a.Item_Code),a.description,a.HSN_SAC_CODE FROM Item_Mst a,Itembal_Mst b"
                        Else
                            strItembal = "SELECT Distinct(a.Item_Code),a.description,a.Tariff_code FROM Item_Mst a,Itembal_Mst b"
                        End If

                        'strItembal = "SELECT Distinct(a.Item_Code),a.description,a.Tariff_code FROM Item_Mst a,Itembal_Mst b"
                        strItembal = strItembal & " where a.unit_code=b.unit_code and a.Item_Code = b.Item_Code and a.Item_Main_Grp IN('C','R','B','M')"
                        strItembal = strItembal & " and cur_bal >0 and a.Status ='A' and a.Hold_Flag <> 1 and a.unit_code='" & gstrUNITID & "'"
                        If Len(Trim(pstrItemNotin)) > 0 Then
                            strItembal = strItembal & " and a.Item_code not in (" & pstrItemNotin & ")"
                        End If
                        '08/05/2002 changes for scrap invoiceing
                    Case "SCRAP"
                        If gblnGSTUnit = True Then
                            strItembal = "SELECT Distinct(a.Item_Code),a.description,a.HSN_SAC_CODE FROM Item_Mst a,Itembal_Mst b"
                        Else
                            strItembal = "SELECT Distinct(a.Item_Code),a.description,a.Tariff_code FROM Item_Mst a,Itembal_Mst b"
                        End If

                        'strItembal = "SELECT Distinct(a.Item_Code),a.description,a.Tariff_code FROM Item_Mst a,Itembal_Mst b"
                        strItembal = strItembal & " where a.unit_code=b.unit_code  and a.Item_Code=b.Item_Code and a.Item_Code in (Select Item_Code  from ItemBal_Mst Where unit_code='" & gstrUNITID & "' and Location_Code ='" & pstrstockLocation & "' and cur_Bal > 0)"
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

                        'strItembal = "SELECT Distinct(a.Item_Code),a.description,a.Tariff_code FROM Item_Mst a,Itembal_Mst b"
                        strItembal = strItembal & " where a.unit_code = b.unit_code and a.Item_Code = b.Item_Code and a.Item_Main_Grp ='C'"
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

                        'strItembal = "SELECT Distinct(a.Item_Code),a.description,a.Tariff_code FROM Item_Mst a,Itembal_Mst b"
                        '10307080 
                        'strItembal = strItembal & " where a.unit_code = b.unit_code and a.Item_Code=b.Item_Code and a.Item_Main_Grp = 'F'"
                        strItembal = strItembal & " where a.unit_code = b.unit_code and a.Item_Code=b.Item_Code and a.Item_Main_Grp in('F','S') "
                        '10307080 
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

                        'strItembal = "SELECT Distinct(a.Item_Code),a.description,a.Tariff_code FROM Item_Mst a,Itembal_Mst b"
                        strItembal = strItembal & " where a.unit_code = b.unit_code and a.Item_Code=b.Item_Code and a.Item_Main_Grp ='R'"
                        strItembal = strItembal & " and cur_bal >0 and a.Status ='A' and a.Hold_Flag <> 1"
                        strItembal = strItembal & " and b.Location_Code = '" & pstrstockLocation & "' and a.unit_code='" & gstrUNITID & "'"
                        If Len(Trim(pstrItemNotin)) > 0 Then
                            strItembal = strItembal & " and a.Item_code not in (" & pstrItemNotin & ")"
                        End If
                    Case "COMPONENTS"
                        If gblnGSTUnit = True Then
                            strItembal = "SELECT Distinct(a.Item_Code),a.description,a.HSN_SAC_CODE FROM Item_Mst a,Itembal_Mst b"
                        Else
                            strItembal = "SELECT Distinct(a.Item_Code),a.description,a.Tariff_code FROM Item_Mst a,Itembal_Mst b"
                        End If
                        'strItembal = "SELECT Distinct(a.Item_Code),a.description,a.Tariff_code FROM Item_Mst a,Itembal_Mst b"
                        strItembal = strItembal & " where a.unit_code = b.unit_code and a.Item_Code=b.Item_Code and a.Item_Main_Grp ='C'"
                        strItembal = strItembal & " and cur_bal >0 and a.Status ='A' and a.Hold_Flag <> 1"
                        strItembal = strItembal & " and b.Location_Code = '" & pstrstockLocation & "' and a.unit_code='" & gstrUNITID & "'"
                        If Len(Trim(pstrItemNotin)) > 0 Then
                            strItembal = strItembal & " and a.Item_code not in (" & pstrItemNotin & ")"
                        End If
                End Select
            Case "TRANSFER INVOICE", "INTER-DIVISION"
                Select Case pstrInvSubtype
                    Case "ASSETS"
                        If gblnGSTUnit = True Then
                            strItembal = "SELECT Distinct(a.Item_Code),a.description,a.HSN_SAC_CODE FROM Item_Mst a,Itembal_Mst b"
                        Else
                            strItembal = "SELECT Distinct(a.Item_Code),a.description,a.Tariff_code FROM Item_Mst a,Itembal_Mst b"
                        End If

                        'strItembal = "SELECT Distinct(a.Item_Code),a.description,a.Tariff_code FROM Item_Mst a,Itembal_Mst b"
                        strItembal = strItembal & " where a.unit_code = b.unit_code and a.Item_Code=b.Item_Code and a.Item_Main_Grp ='P'"
                        strItembal = strItembal & " and cur_bal >0 and a.Status ='A' and a.Hold_Flag <> 1"
                        strItembal = strItembal & " and b.Location_Code = '" & pstrstockLocation & "' and a.unit_code='" & gstrUnitId & "'"
                        '101254587
                        strItembal = strItembal & CreateSubQueryForGlobalToolItemCheck("'P'", pstrAccountCode)
                        If Len(Trim(pstrItemNotin)) > 0 Then
                            strItembal = strItembal & " and a.Item_code not in (" & pstrItemNotin & ")"
                        End If
                    Case "FINISHED GOODS"
                        If gblnGSTUnit = True Then
                            strItembal = "SELECT Distinct a.Item_Code,c.Cust_drgNo,c.Drg_Desc, a.HSN_SAC_CODE FROM Item_Mst a,Itembal_Mst b,CustItem_Mst c "
                        Else
                            strItembal = "SELECT Distinct a.Item_Code,c.Cust_drgNo,c.Drg_Desc, a.Tariff_code FROM Item_Mst a,Itembal_Mst b,CustItem_Mst c "
                        End If
                        'strItembal = "SELECT Distinct a.Item_Code,c.Cust_drgNo,c.Drg_Desc, a.Tariff_code FROM Item_Mst a,Itembal_Mst b,CustItem_Mst c "
                        strItembal = strItembal & " where a.unit_code = b.unit_code and a.unit_code = c.unit_code and a.Item_Code=b.Item_Code and a.Item_Main_Grp IN('F','S') and a.Item_Code = c.ITem_Code"
                        strItembal = strItembal & " and cur_bal >0 and a.Status ='A' and a.Hold_Flag <> 1 and c.Account_code ='" & pstrAccountCode & "'"
                        strItembal = strItembal & " and b.Location_Code = '" & pstrstockLocation & "' and a.unit_code='" & gstrUNITID & "'"
                        If Len(Trim(pstrItemNotin)) > 0 Then
                            strItembal = strItembal & " and a.Item_code not in (" & pstrItemNotin & ")"
                        End If
                    Case "INPUTS"
                        '    strItembal = "SELECT Distinct(a.Item_Code),a.description,a.Tariff_code FROM Item_Mst a,Itembal_Mst b"
                        If gblnGSTUnit = True Then
                            strItembal = "SELECT Distinct(a.Item_Code),a.description,a.HSN_SAC_CODE FROM Item_Mst a,Itembal_Mst b"
                        Else
                            strItembal = "SELECT Distinct(a.Item_Code),a.description,a.Tariff_code FROM Item_Mst a,Itembal_Mst b"
                        End If
                        strItembal = strItembal & " where a.unit_code = b.unit_code and a.Item_Code=b.Item_Code and a.Item_Main_Grp in('R','C','M','N','S','B','A')"
                        strItembal = strItembal & " and cur_bal >0 and a.Status ='A' and a.Hold_Flag <> 1"
                        strItembal = strItembal & " and b.Location_Code = '" & pstrstockLocation & "' and a.unit_code='" & gstrUNITID & "'"
                        If Len(Trim(pstrItemNotin)) > 0 Then
                            strItembal = strItembal & " and a.Item_code not in (" & pstrItemNotin & ")"
                        End If
                End Select
            Case "REJECTION"
                If gblnGSTUnit = True Then
                    strItembal = "SELECT Distinct(a.Item_Code),a.description, c.hsn_sac_code FROM vend_item a (NOLOCK),Itembal_Mst b (NOLOCK),Item_Mst c (NOLOCK)"
                Else
                    strItembal = "SELECT Distinct(a.Item_Code),a.description,c.Tariff_code FROM vend_item a (NOLOCK),Itembal_Mst b (NOLOCK),Item_Mst c (NOLOCK)"
                End If

                strItembal = strItembal & " where a.unit_code = b.unit_code and a.unit_code = c.unit_code and a.Item_Code=b.Item_Code and a.Item_code = c.Item_code and a.Account_code ='" & pstrAccountCode & "'"
                strItembal = strItembal & " and cur_bal >0 and a.unit_code='" & gstrUNITID & "' "
                strItembal = strItembal & " and b.Location_Code = '" & pstrstockLocation & "' "
                'Code Added By Arul on 18-04-2006 To check the active flag
                strItembal = strItembal & " and c.status = 'A' "
                'Addition ends here
                If Len(Trim(pstrItemNotin)) > 0 Then
                    strItembal = strItembal & " and a.Item_code not in (" & pstrItemNotin & ")"
                End If
                strItembal = strItembal & " Union SELECT DISTINCT B.ITEM_CODE,E.DESCRIPTION,ISNULL(E.TARIFF_CODE,0) TARIFF_CODE FROM LRN_HDR A (NOLOCK),LRN_DTL B (NOLOCK),ITEM_MST E (NOLOCK),ITEMBAL_MST F (NOLOCK),GRN_HDR C (NOLOCK),GRN_DTL D (NOLOCK)"
                strItembal = strItembal & " WHERE A.DOC_NO = B.DOC_NO AND A.FROM_LOCATION = B.FROM_LOCATION "
                strItembal = strItembal & " AND A.UNIT_CODE = B.UNIT_CODE AND B.UNIT_CODE = E.UNIT_CODE AND C.UNIT_CODE = D.UNIT_code AND D.UNIT_CODE = B.UNIT_CODE AND B.UNIT_CODE = F.UNIT_CODE "
                strItembal = strItembal & " AND B.ITEM_CODE = E.ITEM_CODE AND B.ITEM_CODE = F.ITEM_CODE AND A.AUTHORIZED_CODE IS NOT NULL "
                strItembal = strItembal & " AND B.REJECTED_QUANTITY > 0 AND F.LOCATION_CODE = '" & pstrstockLocation & "' AND F.CUR_BAL > 0 "
                strItembal = strItembal & " AND D.ITEM_CODE = B.ITEM_CODE AND C.DOC_NO = D.DOC_NO AND C.VENDOR_CODE = '" & pstrAccountCode & "' AND ISNULL(C.GRN_CANCELLED,0) = 0 AND C.DOC_CATEGORY = 'Z' AND A.UNIT_CODE='" & gstrUNITID & "'"
                strItembal = strItembal & " AND E.status = 'A' "
                If Len(Trim(pstrItemNotin)) > 0 Then
                    strItembal = strItembal & " and b.Item_code not in (" & pstrItemNotin & ")"
                End If
                'Addition Ends Here
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
                        If gblnGSTUnit = False Then
                            mListItemUserId.SubItems(3).Text = rsItembal.GetValue("Tariff_Code")
                        Else
                            mListItemUserId.SubItems(3).Text = rsItembal.GetValue("hsn_sac_code")
                        End If

                    Else
                        If gblnGSTUnit = False Then
                            mListItemUserId.SubItems.Insert(3, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsItembal.GetValue("Tariff_Code")))
                        Else
                            mListItemUserId.SubItems.Insert(3, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsItembal.GetValue("hsn_sac_code")))
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
                        If gblnGSTUnit = False Then
                            mListItemUserId.SubItems(3).Text = rsItembal.GetValue("Tariff_Code")
                        Else
                            mListItemUserId.SubItems(3).Text = rsItembal.GetValue("HSN_SAC_CODE")
                        End If
                    Else
                        If gblnGSTUnit = False Then
                            mListItemUserId.SubItems.Insert(3, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsItembal.GetValue("Tariff_Code")))
                        Else
                            mListItemUserId.SubItems.Insert(3, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsItembal.GetValue("HSN_SAC_CODE ")))
                        End If
                    End If
                    rsItembal.MoveNext() 'move to next record
                Next intCount
            End If
            rsItembal.ResultSetClose()
            rsItembal = Nothing
        Else
            '    Call ConfirmWindow(10438, BUTTON_OK)
            '*** 11/06/2002 Changed Message
            If ((UCase(pstrInvType) = "TRANSFER INVOICE") Or (UCase(pstrInvType) = "INTER-DIVISION")) And UCase(pstrInvSubtype) = "FINISHED GOODS" Then
                MsgBox("No items details defined  for above Invoice combination,Please Check Following :" & vbCrLf & "1. Item should be Active and Not on Hold." & vbCrLf & "2. Check Balance of Items for location " & pstrstockLocation & "." & vbCrLf & "3.Item is not defined in Customer ITem Master.", MsgBoxStyle.Information, "empower")
            Else
                MsgBox("No items details defined  for above Invoice combination,Please Check Following :" & vbCrLf & "1. Item should be Active and Not on Hold." & vbCrLf & "2. Check Balance of Items for location " & pstrstockLocation & ".", MsgBoxStyle.Information, "empower")
            End If
            '***
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
        Call AddColumnsInListView()
        Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(mdifrmMain.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2)
        'Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(mdifrmMain.Width) - VB6.PixelsToTwipsX(frmModules.Width)) / 2.3)
        optPartNo.Checked = True
        lvwItemCode.FullRowSelect = True
        'changed due to more then 4 item selection in case of Export in voice.
        strInvType = "select a.description,a.Sub_type_Description from saleconf a,saleschallan_dtl b where a.unit_code = b.unit_code and a.Invoice_type =b.Invoice_Type and b.Doc_no = " & Val(pstrchallanNo) & " and datediff(dd,b.Invoice_Date,a.fin_start_date)<=0  and datediff(dd,a.fin_end_date,b.Invoice_Date)<=0 and a.unit_code='" & gstrUNITID & "'"
        'changes ends here
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
        '****************************
        strsaledtl = ""
        strsaledtl = "Select a.Item_Code,a.Cust_ITem_Code,a.Cust_Item_Desc,b.Tariff_Code from Sales_Dtl a,Item_Mst b where a.unit_code = b.unit_code and a.ITem_code = b.ITem_code and a.unit_code='" & gstrUNITID & "' and Doc_No ="
        strsaledtl = strsaledtl & pstrchallanNo
        rssaledtl = New ClsResultSetDB
        rssaledtl.GetResult(strsaledtl, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        intRecordCount = rssaledtl.GetNoRows 'assign record count to integer variable
        If intRecordCount > 0 Then '          'if record found
            chkall.Visible = True
            rssaledtl.MoveFirst() 'move to first record
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
                    mListItemUserId.SubItems(3).Text = rssaledtl.GetValue("Tariff_code")
                Else
                    mListItemUserId.SubItems.Insert(3, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rssaledtl.GetValue("Tariff_code")))
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
    'Private Function GetServerDate() As Date
    '    Dim objServerDate As ClsResultSetDB 'Class Object
    '    Dim strsql As String 'Stores the SQL statement
    '    'Build the SQL statement
    '    strsql = "SELECT CONVERT(datetime,getdate(),103)"
    '    'Creating the instance
    '    objServerDate = New ClsResultSetDB
    '    With objServerDate
    '        'Open the recordset
    '        Call .GetResult(strsql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
    '        'If we have a record, then getting the financial year else exiting
    '        If .GetNoRows <= 0 Then Exit Function
    '        'Getting the date
    '        'ServerDate = DateValue(Format(.GetValueByNo(0), gstrDateFormat))
    '        GetServerDate = CDate(VB6.Format(DateValue(.GetValueByNo(0)), gstrDateFormat))
    '        'Closing the recordset
    '        .ResultSetClose()
    '    End With
    '    'Releasing the object
    '    objServerDate = Nothing
    'End Function
    Public Function makeSelectSql(ByRef pstrCustno As String, ByRef pstrRefNo As String, ByRef pstrAmmNo As String, ByRef effectyrmon As String, ByRef Validyrmon As String, ByRef pstrstockLocation As String, ByRef strDate As String, ByRef pstrItemin As String, Optional ByRef pstrCondition As String = "", Optional ByRef pstrConsCode As String = "") As String
        '=======================================================================================
        'Revised By     : Ashutosh Verma
        'Revised On     : 22-01-2007 ,Issue ID:19352
        'Revised Reason : Consider Current month schedule while Invoicing.
        '=======================================================================================
        Dim strSelectSql As String
        strDate = getDateForDB(strDate)
        If gblnGSTUnit = True Then
            strSelectSql = "Select distinct b.Item_Code,c.Cust_DrgNo,c.Cust_Drg_Desc,d.HSN_SAC_CODE, c.EXTERNAL_SALESORDER_NO , C.AMENDMENT_NO , SHOP_CODE "
        Else
            strSelectSql = "Select distinct b.Item_Code,c.Cust_DrgNo,c.Cust_Drg_Desc,d.Tariff_Code , c.EXTERNAL_SALESORDER_NO , C.AMENDMENT_NO , SHOP_CODE "
        End If
        If ISPCACUSTOMER = True Then
            strSelectSql = strSelectSql & " ,PCF.MESSAGECALLOFFNO, PCF.PACKAGECODE, PCF.QUANTITY "
        End If
        strSelectSql = strSelectSql & " , CUR_BAL= (	SELECT CUR_BAL FROM ITEMBAL_MST I WHERE B.UNIT_CODE = I.UNIT_CODE 	AND B.ITEM_CODE = I.ITEM_CODE AND I.UNIT_CODE='" & gstrUNITID & "' "
        strSelectSql = strSelectSql & " AND I.LOCATION_CODE='" & pstrstockLocation & "')"
        strSelectSql = strSelectSql & " from Cust_Ord_hdr a,MonthlyMktSchedule b,Cust_ord_dtl c,Item_Mst d , custitem_mst cm "
        If ISPCACUSTOMER = True Then
            strSelectSql = strSelectSql & " ,PCA_FILEUPLOADING  PCF"
        End If
        '("")
        strSelectSql = strSelectSql & " where a.unit_code = b.unit_code and b.unit_code = c.unit_code and B.unit_code = d.unit_code "
        strSelectSql = strSelectSql & " and a.Cust_ref = c.Cust_ref and a.amendment_No = c.amendment_No and a.Account_code=c.account_code And c.Active_Flag ='A' And c.Authorized_flag = 1"
        strSelectSql = strSelectSql & " and cm.unit_code=b.unit_code and cm.account_code =b.account_code and cm.item_Code=b.item_code and cm.cust_drgno=b.Cust_drgNo "
        If ISPCACUSTOMER = True Then
            strSelectSql = strSelectSql & " and pcf.unit_code=c.unit_code and pcf.buyerspartnumber=c.cust_drgno and c.account_code=pcf.customer_code and pcf.customer_code ='" & Trim(pstrCustno) & "' "
        End If
        strSelectSql = strSelectSql & " and a.account_code=b.Account_code and c.Cust_drgNo=b.Cust_drgNo and b.ITem_code = d.Item_code and a.Account_Code='" & Trim(pstrCustno) & "' and a.unit_code='" & gstrUNITID & "'"
        If Len(Trim(pstrConsCode)) > 0 Then
            strSelectSql = strSelectSql & " and a.consignee_code='" & Trim(pstrConsCode) & "' and a.Consignee_code=b.Consignee_code "
        End If
        strSelectSql = strSelectSql & " and b.ITem_code = c.Item_code and a.Cust_Ref in ('" & Trim(pstrRefNo) & "') and b.status = 1 and b.Schedule_flag =1 and b.Year_Month =  " & Validyrmon
        strSelectSql = strSelectSql & " and b.Item_Code in(Select a.Item_code from Item_MSt a,Itembal_mst b where a.unit_code = b.unit_code and a.unit_code='" & gstrUnitId & "' and a.Item_Main_grp in (" & Trim(pstrItemin) & ") and a.Item_code = b.Item_code and b.Location_code ='" & pstrstockLocation & "' and b.Cur_bal >0 and a.hold_flag =0 and a.Status = 'A'"
        '101254587
        strSelectSql = strSelectSql & CreateSubQueryForGlobalToolItemCheck(pstrItemin, pstrCustno)
        If Len(Trim(pstrCondition)) > 0 Then
            strSelectSql = strSelectSql & " and a.Item_code not in(" & pstrCondition & "))"
        Else
            strSelectSql = strSelectSql & ")"
        End If
        strSelectSql = strSelectSql & " UNION "
        If gblnGSTUnit = True Then
            strSelectSql = strSelectSql & "Select distinct b.Item_Code,c.Cust_DrgNo,c.Cust_Drg_Desc,d.HSN_SAC_CODE, c.EXTERNAL_SALESORDER_NO , C.AMENDMENT_NO , SHOP_CODE "
        Else
            strSelectSql = strSelectSql & "Select distinct b.Item_Code,c.Cust_DrgNo,c.Cust_Drg_Desc,d.Tariff_Code , c.EXTERNAL_SALESORDER_NO , C.AMENDMENT_NO , SHOP_CODE "
        End If
        If ISPCACUSTOMER = True Then
            strSelectSql = strSelectSql & " ,PCF.MESSAGECALLOFFNO, PCF.PACKAGECODE ,PCF.QUANTITY"
        End If
        strSelectSql = strSelectSql & " , CUR_BAL=(	SELECT CUR_BAL FROM ITEMBAL_MST I WHERE B.UNIT_CODE = I.UNIT_CODE 	AND B.ITEM_CODE = I.ITEM_CODE AND I.UNIT_CODE='" & gstrUNITID & "' "
        strSelectSql = strSelectSql & " AND I.LOCATION_CODE='" & pstrstockLocation & "')"
        strSelectSql = strSelectSql & " from Cust_Ord_hdr a,Dailymktschedule   b,Cust_ord_dtl c,Item_Mst d , custitem_mst cm  "
        If ISPCACUSTOMER = True Then
            strSelectSql = strSelectSql & " ,PCA_FILEUPLOADING  PCF"
        End If
        strSelectSql = strSelectSql & " where a.unit_code=b.unit_code and b.unit_code=c.unit_code and B.unit_code  = d.unit_code and "
        strSelectSql = strSelectSql & " a.Cust_ref = c.Cust_ref and a.amendment_No = c.amendment_No and a.Account_code=c.account_code"
        strSelectSql = strSelectSql & " and a.account_code=b.Account_code and c.Cust_drgNo=b.Cust_drgNo and b.ITem_code =d.ITem_code "
        If ISPCACUSTOMER = True Then
            strSelectSql = strSelectSql & " and pcf.unit_code=c.unit_code and pcf.buyerspartnumber=c.cust_drgno and c.account_code=pcf.customer_code and pcf.customer_code ='" & Trim(pstrCustno) & "' "
        End If

        strSelectSql = strSelectSql & " and cm.unit_code=b.unit_code and cm.account_code =b.account_code and cm.item_Code=b.item_code and cm.cust_drgno=B.Cust_drgNo and "
        strSelectSql = strSelectSql & " b.status = 1 and b.Schedule_Flag = 1 And c.Active_Flag ='A' And c.Authorized_flag = 1 and a.Account_Code='" & Trim(pstrCustno) & "' "
        strSelectSql = strSelectSql & " and a.unit_code='" & gstrUNITID & "'"
        ''Changes done by Ashutosh on 16 Apr 2007, issue id:19731
        If Len(Trim(pstrConsCode)) > 0 Then
            strSelectSql = strSelectSql & " and a.consignee_code='" & Trim(pstrConsCode) & "' and a.Consignee_code=b.Consignee_code "
        End If
        ''Changes for issue id:19731 end here.
        ''*** Changes done By ashutosh on 22-01-2007, issue Id: 19352, Consider Current month schedule.
        strSelectSql = strSelectSql & " and b.ITem_code = c.Item_code and a.Cust_Ref in('" & Trim(pstrRefNo) & "') and b.trans_date <= '" & strDate & "' And datepart(mm,b.trans_date) = '" & Month(CDate(strDate)) & "' And DatePart(yyyy, b.Trans_Date) = '" & Year(CDate(strDate)) & "'"
        ''*** Changes for Issue Id:19352 end here.
        strSelectSql = strSelectSql & " and b.Item_Code in(Select a.Item_code from Item_MSt a,Itembal_mst b where a.unit_code = b.unit_code and a.unit_code ='" & gstrUnitId & "' and a.Item_Main_grp in (" & Trim(pstrItemin) & ") and a.Item_code = b.Item_code and b.Location_code ='" & pstrstockLocation & "' and b.Cur_bal >0 and a.hold_flag =0 and a.Status = 'A'"
        '101254587
        strSelectSql = strSelectSql & CreateSubQueryForGlobalToolItemCheck(pstrItemin, pstrCustno)
        If Len(Trim(pstrCondition)) > 0 Then
            strSelectSql = strSelectSql & " and a.Item_code not in( " & pstrCondition & "))"
        Else
            strSelectSql = strSelectSql & ")"
        End If
        If ISPCACUSTOMER = True Then
            strSelectSql = strSelectSql & " and pcf.scheduledeliverydate >= convert(DATE,getdate()-" & mintnoofbackdaysschedule_PCA & ",103) "
            strSelectSql = strSelectSql & " and not exists ( select top 1 1 from PCA_SCHEDULE_INVOICE_KNOCKOFF PFS WHERE PFS.UNIT_CODE = PCF.UNIT_CODE "
            strSelectSql = strSelectSql & " AND PFS.PACKAGE_CODE =PCF.PACKAGECODE AND PFS.PART_CODE=PCF.BuyersPartNumber  And cancelled_flag=0 ) "
            strSelectSql = strSelectSql & " ORDER BY B.ITEM_CODE,PCF.PACKAGECODE "
        End If

        makeSelectSql = strSelectSql
    End Function
    Public Function makeSelectSql_service(ByRef pstrCustno As String, ByRef pstrRefNo As String, ByRef pstrAmmNo As String, ByRef effectyrmon As String, ByRef Validyrmon As String, ByRef pstrstockLocation As String, ByRef strDate As String, ByRef pstrItemin As String, Optional ByRef pstrCondition As String = "", Optional ByRef pstrConsCode As String = "") As String
        '=======================================================================================
        'Revised By     : Ashutosh Verma
        'Revised On     : 22-01-2007 ,Issue ID:19352
        'Revised Reason : Consider Current month schedule while Invoicing.
        '=======================================================================================
        Dim strSelectSql As String
        strDate = getDateForDB(strDate)
        If gblnGSTUnit = True Then
            strSelectSql = "Select distinct b.Item_Code,c.Cust_DrgNo,c.Cust_Drg_Desc,d.hsn_sac_code , c.EXTERNAL_SALESORDER_NO , C.AMENDMENT_NO , SHOP_CODE "
        Else
            strSelectSql = "Select distinct b.Item_Code,c.Cust_DrgNo,c.Cust_Drg_Desc,d.Tariff_Code , c.EXTERNAL_SALESORDER_NO , C.AMENDMENT_NO , SHOP_CODE "
        End If

        strSelectSql = strSelectSql & " from Cust_Ord_hdr a,MonthlyMktSchedule b,Cust_ord_dtl c,Item_Mst d , custitem_mst cm where "
        strSelectSql = strSelectSql & " a.unit_code = b.unit_code and b.unit_code = c.unit_code and B.unit_code = d.unit_code "
        strSelectSql = strSelectSql & " and a.Cust_ref = c.Cust_ref and a.amendment_No = c.amendment_No and a.Account_code=c.account_code And c.Active_Flag ='A' And c.Authorized_flag = 1"
        strSelectSql = strSelectSql & " and cm.unit_code=b.unit_code and cm.account_code =b.account_code and cm.item_Code=b.item_code and cm.cust_drgno=b.Cust_drgNo "
        strSelectSql = strSelectSql & " and a.account_code=b.Account_code and c.Cust_drgNo=b.Cust_drgNo and b.ITem_code = d.Item_code and a.Account_Code='" & Trim(pstrCustno) & "' and a.unit_code='" & gstrUNITID & "'"
        If Len(Trim(pstrConsCode)) > 0 Then
            strSelectSql = strSelectSql & " and a.consignee_code='" & Trim(pstrConsCode) & "' and a.Consignee_code=b.Consignee_code "
        End If
        strSelectSql = strSelectSql & " and b.ITem_code = c.Item_code and a.Cust_Ref in ('" & Trim(pstrRefNo) & "') and b.status = 1 and b.Schedule_flag =1 and b.Year_Month =  " & Validyrmon
        strSelectSql = strSelectSql & " and b.Item_Code in(Select a.Item_code from Item_MSt a where a.unit_code = '" & gstrUnitId & "' and a.Item_Main_grp in (" & Trim(pstrItemin) & ")  and a.hold_flag =0 and a.Status = 'A'"
        '101254587
        strSelectSql = strSelectSql & CreateSubQueryForGlobalToolItemCheck(pstrItemin, pstrCustno)
        If Len(Trim(pstrCondition)) > 0 Then
            strSelectSql = strSelectSql & " and a.Item_code not in(" & pstrCondition & "))"
        Else
            strSelectSql = strSelectSql & ")"
        End If
        strSelectSql = strSelectSql & " UNION "
        If gblnGSTUnit = True Then
            strSelectSql = strSelectSql & "Select distinct b.Item_Code,c.Cust_DrgNo,c.Cust_Drg_Desc,d.hsn_sac_code , c.EXTERNAL_SALESORDER_NO , C.AMENDMENT_NO , SHOP_CODE"
        Else
            strSelectSql = strSelectSql & "Select distinct b.Item_Code,c.Cust_DrgNo,c.Cust_Drg_Desc,d.Tariff_Code , c.EXTERNAL_SALESORDER_NO , C.AMENDMENT_NO , SHOP_CODE"
        End If

        strSelectSql = strSelectSql & " from Cust_Ord_hdr a,Dailymktschedule   b,Cust_ord_dtl c,Item_Mst d , custitem_mst cm  where "
        strSelectSql = strSelectSql & " a.unit_code=b.unit_code and b.unit_code=c.unit_code and B.unit_code  = d.unit_code and "
        strSelectSql = strSelectSql & " a.Cust_ref = c.Cust_ref and a.amendment_No = c.amendment_No and a.Account_code=c.account_code"
        strSelectSql = strSelectSql & " and a.account_code=b.Account_code and c.Cust_drgNo=b.Cust_drgNo and b.ITem_code =d.ITem_code "
        strSelectSql = strSelectSql & " and cm.unit_code=b.unit_code and cm.account_code =b.account_code and cm.item_Code=b.item_code and cm.cust_drgno=B.Cust_drgNo and "
        strSelectSql = strSelectSql & " b.status = 1 and b.Schedule_Flag = 1 And c.Active_Flag ='A' And c.Authorized_flag = 1 and a.Account_Code='" & Trim(pstrCustno) & "' "
        strSelectSql = strSelectSql & " and a.unit_code='" & gstrUNITID & "'"
        ''Changes done by Ashutosh on 16 Apr 2007, issue id:19731
        If Len(Trim(pstrConsCode)) > 0 Then
            strSelectSql = strSelectSql & " and a.consignee_code='" & Trim(pstrConsCode) & "' and a.Consignee_code=b.Consignee_code "
        End If
        ''Changes for issue id:19731 end here.
        ''*** Changes done By ashutosh on 22-01-2007, issue Id: 19352, Consider Current month schedule.
        strSelectSql = strSelectSql & " and b.ITem_code = c.Item_code and a.Cust_Ref in('" & Trim(pstrRefNo) & "') and b.trans_date <= '" & strDate & "' And datepart(mm,b.trans_date) = '" & Month(CDate(strDate)) & "' And DatePart(yyyy, b.Trans_Date) = '" & Year(CDate(strDate)) & "'"
        ''*** Changes for Issue Id:19352 end here.
        strSelectSql = strSelectSql & " and b.Item_Code in(Select a.Item_code from Item_MSt a where a.unit_code = '" & gstrUnitId & "' and a.Item_Main_grp in (" & Trim(pstrItemin) & ") and a.hold_flag =0 and a.Status = 'A'"
        '101254587
        strSelectSql = strSelectSql & CreateSubQueryForGlobalToolItemCheck(pstrItemin, pstrCustno)
        If Len(Trim(pstrCondition)) > 0 Then
            strSelectSql = strSelectSql & " and a.Item_code not in( " & pstrCondition & "))"
        Else
            strSelectSql = strSelectSql & ")"
        End If
        makeSelectSql_service = strSelectSql
    End Function
    Public Function MakeSelectSubQuery(ByRef pstrCustno As String, ByRef pstrRefNo As String, ByRef pstrAmmNo As String, ByRef pstrstockLocation As String, ByRef pstrItemin As String, Optional ByRef pstrItemNotin As String = "") As String
        Dim strSelectSql As String
        If gblnGSTUnit = True Then
            strSelectSql = "Select c.Item_Code,c.Cust_DrgNo,c.Cust_Drg_Desc,d.hsn_sac_code  from Cust_Ord_hdr a,Cust_ord_dtl c,Item_Mst d where "
        Else
            strSelectSql = "Select c.Item_Code,c.Cust_DrgNo,c.Cust_Drg_Desc,d.Tariff_Code from Cust_Ord_hdr a,Cust_ord_dtl c,Item_Mst d where "
        End If
        strSelectSql = strSelectSql & " a.unit_code = c.unit_code and c.unit_code = d.unit_code and "
        strSelectSql = strSelectSql & " a.Cust_ref = c.Cust_ref and a.amendment_No = c.amendment_No and a.Account_code=c.account_code"
        'Code Commented And Added By    -   Nitin Sood
        'strSelectSql = strSelectSql & " and  c.Active_Flag =a.Active_flag and c.Item_code = d.Item_code and a.Account_Code='" & Trim(pstrCustno) & "' and a.Cust_Ref='" & Trim(pstrRefNo)
        strSelectSql = strSelectSql & " and  c.Item_code = d.Item_code and a.unit_code='" & gstrUNITID & "' and a.Account_Code='" & Trim(pstrCustno) & "' and a.Cust_Ref in('" & Trim(pstrRefNo)
        'Code changed by Arul on 07-03-2005
        'Authorized_flag checking condition added by below line
        'strSelectSql = strSelectSql & "') and a.Amendment_No='" & Trim(pstrAmmNo) & "' And c.Active_Flag = 'A' And c.Authorized_flag = 1 "
        strSelectSql = strSelectSql & "') and c.Active_Flag = 'A' And c.Authorized_flag = 1 "
        strSelectSql = strSelectSql & " and c.Item_Code in(Select a.Item_code from Item_MSt a,Itembal_mst b where a.unit_code = b.unit_code and a.unit_code='" & gstrUnitId & "' and a.Item_Main_grp in (" & pstrItemin & ") and a.Item_code = b.Item_code and b.Location_code ='" & pstrstockLocation & "' and b.Cur_bal >0 and a.hold_flag =0 and a.Status = 'A'"
        '101254587
        strSelectSql = strSelectSql & CreateSubQueryForGlobalToolItemCheck(pstrItemin, pstrCustno)
        If Len(Trim(pstrItemNotin)) > 0 Then
            strSelectSql = strSelectSql & " and a.Item_code not in ( " & pstrItemNotin & "))"
        Else
            strSelectSql = strSelectSql & ")"
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
            'itmFound.Selected = True ' Select the ListItem.
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
            If ISPCACUSTOMER = False Then
                With lvwItemCode
                    .Sort()
                    '.SortKey = 2
                    ListViewColumnSorter.SortListView(lvwItemCode, 2, SortOrder.Ascending)
                    .Sorting = System.Windows.Forms.SortOrder.Ascending
                End With
            End If
            
            Exit Sub
ErrHandler:
            gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub
    Private Sub OptItemCode_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optItemCode.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo ErrHandler
            If ISPCACUSTOMER = False Then
                With lvwItemCode
                    .Sort()
                    '.SortKey = 0
                    ListViewColumnSorter.SortListView(lvwItemCode, 0, SortOrder.Ascending)
                    .Sorting = System.Windows.Forms.SortOrder.Ascending
                End With
            End If
            Exit Sub
ErrHandler:
            gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub
    Private Sub optPartNo_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optPartNo.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo ErrHandler
            If ISPCACUSTOMER = False Then
                With lvwItemCode
                    .Sort()
                    '.SortKey = 1
                    ListViewColumnSorter.SortListView(lvwItemCode, 1, SortOrder.Ascending)
                    .Sorting = System.Windows.Forms.SortOrder.Ascending
                End With
            End If

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
        Dim strsql As String
        Dim StrItemCode As String
        Dim strItemNot As String
        Dim arrRejAcpt(,) As Object
        Dim intLoopCounter As Short
        Dim intArrLoopCount As Short
        Dim intMaxLoop As Short
        Dim intUbound As Short
        Call AddColumnsInListView()
        Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(mdifrmMain.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2)
        'Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(mdifrmMain.Width) - VB6.PixelsToTwipsX(frmModules.Width)) / 2.3)
        optPartNo.Checked = True
        lvwItemCode.FullRowSelect = True
        mstrInvType = "REJECTION" : mstrInvSubType = "REJECTION"
        On Error GoTo ErrHandler
        rsGrnDtl = New ClsResultSetDB
        strsql = "select a.Doc_No,a.Item_code,a.Rejected_Quantity,Despatch_Quantity = isnull(a.Despatch_Quantity,0),"
        strsql = strsql & " Inspected_Quantity = isnull(a.Inspected_Quantity,0), RGP_Quantity = isnull(a.RGP_Quantity,0)  from grn_Dtl a,"
        strsql = strsql & " grn_hdr b Where a.unit_code = b.unit_code and "
        strsql = strsql & " a.Doc_type = b.Doc_type And a.Doc_No = b.Doc_No and "
        strsql = strsql & " a.From_Location = b.From_Location and a.From_Location ='01R1' "
        strsql = strsql & " and a.Rejected_quantity > 0 and b.Vendor_code = '" & pstrVend
        strsql = strsql & "' and a.Doc_No = " & dblGrnNo & " AND ISNULL(GRN_Cancelled,0) = 0 and a.unit_code='" & gstrUNITID & "'"
        If Len(Trim(pstrCondition)) > 0 Then
            strsql = strsql & " and a.Item_code not in (" & pstrCondition & ")"
        End If
        rsGrnDtl.GetResult(strsql)
        If rsGrnDtl.GetNoRows > 0 Then
            intMaxLoop = rsGrnDtl.GetNoRows : rsGrnDtl.MoveFirst() : ReDim arrRejAcpt(2, intMaxLoop - 1) : intUbound = intMaxLoop - 1
            '****To Fatch all Doc_No and Rejected Quantity in Array
            intUbound = intMaxLoop - 1
            For intLoopCounter = 1 To intMaxLoop
                'arrRejAcpt(0, intLoopCounter - 1) = rsGrnDtl.GetValue("Doc_No")
                arrRejAcpt(0, intLoopCounter - 1) = rsGrnDtl.GetValue("Item_Code")
                arrRejAcpt(1, intLoopCounter - 1) = rsGrnDtl.GetValue("Rejected_Quantity") - rsGrnDtl.GetValue("Despatch_Quantity") - rsGrnDtl.GetValue("Inspected_Quantity") - rsGrnDtl.GetValue("RGP_Quantity")
                rsGrnDtl.MoveNext()
            Next
            '****
            strItemNot = ""
            For intArrLoopCount = 0 To intUbound
                StrItemCode = arrRejAcpt(0, intArrLoopCount)
                If arrRejAcpt(1, intArrLoopCount) <= 0 Then
                    If Len(Trim(strItemNot)) > 0 Then
                        'Code Changed by Arul on 24-08-2004
                        'strItemNot = strItemCode & ",'" & strItemCode & "'"
                        strItemNot = strItemNot & ",'" & StrItemCode & "'"
                        'Changes Ends here
                    Else
                        strItemNot = "'" & StrItemCode & "'"
                    End If
                End If
            Next
            If Len(Trim(strItemNot)) > 0 Then
                strsql = "select a.Doc_No,a.Item_code,a.Accepted_Quantity,c.Tariff_code,c.Description from grn_dtl a,grn_hdr b,Item_Mst c where "
                strsql = strsql & " a.unit_code= b.unit_code and A.unit_code = c.unit_code "
                strsql = strsql & " and a.Doc_type = b.Doc_type and a.Doc_no = b.Doc_No "
                strsql = strsql & " and a.From_Location = b.From_Location "
                strsql = strsql & " and a.Item_Code = c.ITem_code and b.From_Location ='01R1'"
                strsql = strsql & " and a.Item_code Not in (" & strItemNot & ")"
                strsql = strsql & " and c.Status = 'A' and Hold_Flag =0"
                strsql = strsql & " and a.Rejected_quantity > 0 and b.Vendor_code = '" & pstrVend
                strsql = strsql & "' and a.Doc_No = " & dblGrnNo & " AND ISNULL(GRN_Cancelled,0) = 0 "
                strsql = strsql & " and a.unit_code= '" & gstrUNITID & "'"
                strsql = strsql & " and a.Item_code in (Select Item_Code from ItemBal_Mst Where  unit_code= '" & gstrUNITID & "' and Location_Code = '"
                strsql = strsql & pstrstockLocation & "' and Cur_bal > 0)"
                If Len(Trim(pstrCondition)) > 0 Then
                    strsql = strsql & " and a.Item_code not in (" & pstrCondition & ")"
                End If
            Else
                If gblnGSTUnit = True Then
                    strsql = "select a.Doc_No,a.Item_code,a.Accepted_Quantity,c.hsn_sac_code,c.Description from grn_dtl a,grn_hdr b,Item_Mst c where "
                Else
                    strsql = "select a.Doc_No,a.Item_code,a.Accepted_Quantity,c.Tariff_code,c.Description from grn_dtl a,grn_hdr b,Item_Mst c where "
                End If

                strsql = strsql & " a.unit_code= b.unit_code and A.unit_code = c.unit_code "
                strsql = strsql & " and a.Doc_type = b.Doc_type and a.Doc_no = b.Doc_No "
                strsql = strsql & " and a.From_Location = b.From_Location "
                strsql = strsql & " and a.Item_Code = c.ITem_code and b.From_Location ='01R1'"
                strsql = strsql & " and c.Status = 'A' and Hold_Flag =0"
                strsql = strsql & "and a.Rejected_quantity > 0 and b.Vendor_code = '" & pstrVend
                strsql = strsql & "' and a.Doc_No = " & dblGrnNo & " AND ISNULL(GRN_Cancelled,0) = 0 "
                strsql = strsql & " and a.unit_code= '" & gstrUNITID & "'"
                strsql = strsql & " and a.Item_code in (Select Item_Code from ItemBal_Mst Where unit_code='" & gstrUNITID & "' and Location_Code = '"
                strsql = strsql & pstrstockLocation & "' and Cur_bal > 0)"
                If Len(Trim(pstrCondition)) > 0 Then
                    strsql = strsql & " and a.Item_code not in (" & pstrCondition & ")"
                End If
            End If
            rsGrnDtl.ResultSetClose()
            rsGrnDtl = Nothing
            rsGrnDtl = New ClsResultSetDB
            rsGrnDtl.GetResult(strsql)
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
                '        Call ConfirmWindow(10440, BUTTON_OK, IMG_INFO)
                '***11/06/2002 Changed Message
                MsgBox("No Items for selected Invoice in Sales Order.Please Check Following :" & vbCrLf & "1. Item in Grin are Active and Not on Hold." & vbCrLf & "2. Check Balance of Items for location " & pstrstockLocation & "." & vbCrLf & "3. Check supplimentry Grin for items in Grin(Selected) ", MsgBoxStyle.Information, "empower")
                '***
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
        Call AddColumnsInListView()
        rsGrnDtl = New ClsResultSetDB
        If strRejType = "LRN" Then
            If gblnGSTUnit = True Then
                strSql = "Select B.Item_Code, I.Description, I.hsn_sac_code from LRN_HDR as a Inner Join LRN_DTL as b on A.UNIT_CODE = B.UNIT_CODE and a.doc_No = b.doc_no and a.Doc_Type = b.doc_Type and a.from_Location = b.from_location Inner join Item_Mst as I On b.unit_code=i.unit_code and b.item_code=i.item_code where b.Item_code In ( Select Item_code from ItemBal_Mst where Cur_Bal>0 and Location_code ='" & pstrstockLocation & "' and unit_code='" & gstrUNITID & "') and Authorized_Code IS Not Null and a.Doc_No IN (" & strDocNo & ") "
            Else
                strSql = "Select B.Item_Code, I.Description, I.Tariff_code from LRN_HDR as a Inner Join LRN_DTL as b on A.UNIT_CODE = B.UNIT_CODE and a.doc_No = b.doc_no and a.Doc_Type = b.doc_Type and a.from_Location = b.from_location Inner join Item_Mst as I On b.unit_code=i.unit_code and b.item_code=i.item_code where b.Item_code In ( Select Item_code from ItemBal_Mst where Cur_Bal>0 and Location_code ='" & pstrstockLocation & "' and unit_code='" & gstrUNITID & "') and Authorized_Code IS Not Null and a.Doc_No IN (" & strDocNo & ") "
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
                            mListItemUserId.SubItems(3).Text = rsGrnDtl.GetValue("hsn_sac_code")
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
                        strSql = "select Distinct a.Item_code, c.hsn_sac_code,c.Description from grn_dtl a,grn_hdr b,Item_Mst c where "
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
                        strSql = "select Distinct a.Item_code, c.hsn_sac_code ,c.Description from grn_dtl a, grn_hdr b, Item_Mst c where "
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
                                mListItemUserId.SubItems(3).Text = rsGrnDtl.GetValue("hsn_sac_code")
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
#Region "GLOBAL TOOL INVOICE CHANGES"
    'ADDED BY VINOD FOR GLOBAL TOOL CHANGES
    Public Function MakeSelectSubQuery_Asset(ByRef pstrCustno As String, ByRef pstrRefNo As String, ByRef pstrAmmNo As String, ByRef pstrstockLocation As String, ByRef pstrItemin As String, Optional ByRef pstrItemNotin As String = "") As String
        Dim strSelectSql As String
        strSelectSql = "Select c.Item_Code,c.Cust_DrgNo,c.Cust_Drg_Desc,d.Tariff_Code,d.HSN_SAC_CODE from Cust_Ord_hdr a,Cust_ord_dtl c,Item_Mst d where "
        strSelectSql = strSelectSql & " a.unit_code = c.unit_code and c.unit_code = d.unit_code and "
        strSelectSql = strSelectSql & " a.Cust_ref = c.Cust_ref and a.amendment_No = c.amendment_No and a.Account_code=c.account_code"
        strSelectSql = strSelectSql & " and  c.Item_code = d.Item_code and a.unit_code='" & gstrUNITID & "' and a.Account_Code='" & Trim(pstrCustno) & "' and a.Cust_Ref in('" & Trim(pstrRefNo)
        strSelectSql = strSelectSql & "') and c.Active_Flag = 'A' And c.Authorized_flag = 1 "
        strSelectSql = strSelectSql & " and c.Item_Code in(Select a.Item_code from Item_MSt a,Itembal_mst b where a.unit_code = b.unit_code and a.unit_code='" & gstrUNITID & "' and a.Item_Main_grp in (" & pstrItemin & ") and a.Item_code = b.Item_code and b.Location_code ='" & pstrstockLocation & "' and b.Cur_bal >0 and a.hold_flag =0 and a.Status = 'A'"
        If Len(Trim(pstrItemNotin)) > 0 Then
            strSelectSql = strSelectSql & " and a.Item_code not in ( " & pstrItemNotin & "))"
        Else
            strSelectSql = strSelectSql & ")"
        End If
        strSelectSql = strSelectSql & " AND C.ITEM_CODE NOT IN "
        strSelectSql = strSelectSql & " ( "
        strSelectSql = strSelectSql & " SELECT I.ITEM_CODE FROM ITEM_MST I INNER JOIN GLOBAL_TOOL_MST G "
        strSelectSql = strSelectSql & " ON I.ITEM_CODE = G.GLOBAL_TOOL_CODE_FOR_ITEM_MASTER WHERE I.UNIT_CODE = '" & gstrUNITID & "' "
        strSelectSql = strSelectSql & " AND ITEM_MAIN_GRP ='P' "
        strSelectSql = strSelectSql & " AND G.CATEGORY IN ('MOULDING','PRESSTOOL') AND I.STATUS='A'"
        strSelectSql = strSelectSql & " ) "
        strSelectSql = strSelectSql & " UNION "
        strSelectSql = strSelectSql & " Select c.Item_Code,c.Cust_DrgNo,c.Cust_Drg_Desc,d.Tariff_Code,d.HSN_SAC_CODE from Cust_Ord_hdr a,Cust_ord_dtl c,Item_Mst d where "
        strSelectSql = strSelectSql & " a.unit_code = c.unit_code and c.unit_code = d.unit_code and "
        strSelectSql = strSelectSql & " a.Cust_ref = c.Cust_ref and a.amendment_No = c.amendment_No and a.Account_code=c.account_code"
        strSelectSql = strSelectSql & " and  c.Item_code = d.Item_code and a.unit_code='" & gstrUNITID & "' and a.Account_Code='" & Trim(pstrCustno) & "' and a.Cust_Ref in('" & Trim(pstrRefNo)
        strSelectSql = strSelectSql & "') and c.Active_Flag = 'A' And c.Authorized_flag = 1 "
        strSelectSql = strSelectSql & " and c.Item_Code in(Select a.Item_code from Item_MSt a,Itembal_mst b where a.unit_code = b.unit_code and a.unit_code='" & gstrUNITID & "' and a.Item_Main_grp in (" & pstrItemin & ") and a.Item_code = b.Item_code and b.Location_code ='" & pstrstockLocation & "' and b.Cur_bal >0 and a.hold_flag =0 and a.Status = 'A'"
        If Len(Trim(pstrItemNotin)) > 0 Then
            strSelectSql = strSelectSql & " and a.Item_code not in ( " & pstrItemNotin & "))"
        Else
            strSelectSql = strSelectSql & ")"
        End If
        strSelectSql = strSelectSql & "AND C.ITEM_CODE IN "
        strSelectSql = strSelectSql & "( "
        strSelectSql = strSelectSql & "SELECT I.ITEM_CODE FROM ITEM_MST I INNER JOIN GLOBAL_TOOL_MST G "
        strSelectSql = strSelectSql & "ON I.ITEM_CODE = G.GLOBAL_TOOL_CODE_FOR_ITEM_MASTER "
        strSelectSql = strSelectSql & "WHERE I.UNIT_CODE = '" & gstrUNITID & "' AND ITEM_MAIN_GRP ='P' "
        strSelectSql = strSelectSql & "AND G.CATEGORY IN ('MOULDING','PRESSTOOL') AND I.STATUS='A' AND G.VENDOR_CODE = '" & pstrCustno & "'"
        strSelectSql = strSelectSql & "AND G.TRANSFER_COMPLETED = 0 AND G.TRANSFER_TYPE in ('SALE','CUSTOMER_TOOL_SALE') AND ISNULL(G.INV_NO,0) = 0"
        strSelectSql = strSelectSql & ") "
        MakeSelectSubQuery_Asset = strSelectSql
    End Function

    Public Function MakeSelectSubQuery_Tool(ByRef pstrCustno As String, ByRef pstrRefNo As String, ByRef pstrAmmNo As String, ByRef pstrstockLocation As String, ByRef pstrItemin As String, Optional ByRef pstrItemNotin As String = "") As String
        Dim strSelectSql As String = String.Empty
        If gblnGSTUnit = True Then
            strSelectSql = "Select c.Item_Code,c.Cust_DrgNo,c.Cust_Drg_Desc,c.hsnSaccode as hsn_sac_code from Cust_Ord_hdr a,Cust_ord_dtl c,Item_Mst d where "
        Else
            strSelectSql = "Select c.Item_Code,c.Cust_DrgNo,c.Cust_Drg_Desc,d.Tariff_Code from Cust_Ord_hdr a,Cust_ord_dtl c,Item_Mst d where "
        End If

        strSelectSql = strSelectSql & " a.unit_code = c.unit_code and c.unit_code = d.unit_code and "
        strSelectSql = strSelectSql & " a.Cust_ref = c.Cust_ref and a.amendment_No = c.amendment_No and a.Account_code=c.account_code"
        strSelectSql = strSelectSql & " and  c.Item_code = d.Item_code and a.unit_code='" & gstrUNITID & "' and a.Account_Code='" & Trim(pstrCustno) & "' and a.Cust_Ref in('" & Trim(pstrRefNo)
        strSelectSql = strSelectSql & "') and c.Active_Flag = 'A' And c.Authorized_flag = 1 "
        strSelectSql = strSelectSql & " and c.Item_Code in(Select a.Item_code from Item_MSt a,Itembal_mst b where a.unit_code = b.unit_code and a.unit_code='" & gstrUNITID & "' and a.Item_Main_grp in (" & pstrItemin & ") and a.Item_code = b.Item_code and b.Location_code ='" & pstrstockLocation & "' and b.Cur_bal >0 and a.hold_flag =0 and a.Status = 'A'"
        If Len(Trim(pstrItemNotin)) > 0 Then
            strSelectSql = strSelectSql & " and a.Item_code not in ( " & pstrItemNotin & "))"
        Else
            strSelectSql = strSelectSql & ")"
        End If
        strSelectSql = strSelectSql & " AND C.ITEM_CODE NOT IN "
        strSelectSql = strSelectSql & " ( "
        strSelectSql = strSelectSql & " SELECT I.ITEM_CODE FROM ITEM_MST I INNER JOIN GLOBAL_TOOL_MST G "
        strSelectSql = strSelectSql & " ON I.ITEM_CODE = G.GLOBAL_TOOL_CODE_FOR_ITEM_MASTER WHERE I.UNIT_CODE = '" & gstrUNITID & "' "
        strSelectSql = strSelectSql & " AND ITEM_MAIN_GRP ='P' "
        strSelectSql = strSelectSql & " AND G.CATEGORY IN ('MOULDING','PRESSTOOL') AND I.STATUS='A'"
        strSelectSql = strSelectSql & " ) "
        Return strSelectSql
    End Function

    Public Function makeSelectSql_Asset(ByRef pstrCustno As String, ByRef pstrRefNo As String, _
                                        ByRef pstrAmmNo As String, ByRef effectyrmon As String, _
                                        ByRef Validyrmon As String, ByRef pstrstockLocation As String, _
                                        ByRef strDate As String, ByRef pstrItemin As String, _
                                        Optional ByRef pstrCondition As String = "", _
                                        Optional ByRef pstrConsCode As String = "") As String

        Dim strSelectSql As String = String.Empty
        strDate = getDateForDB(strDate)
        strSelectSql = "Select distinct b.Item_Code,c.Cust_DrgNo,c.Cust_Drg_Desc,d.Tariff_Code , c.EXTERNAL_SALESORDER_NO , C.AMENDMENT_NO , SHOP_CODE, CUR_BAL="
        strSelectSql = strSelectSql & " (	SELECT CUR_BAL FROM ITEMBAL_MST I WHERE B.UNIT_CODE = I.UNIT_CODE 	AND B.ITEM_CODE = I.ITEM_CODE AND I.UNIT_CODE='" & gstrUNITID & "' "
        strSelectSql = strSelectSql & " AND I.LOCATION_CODE='" & pstrstockLocation & "')"
        strSelectSql = strSelectSql & " from Cust_Ord_hdr a,MonthlyMktSchedule b,Cust_ord_dtl c,Item_Mst d , custitem_mst cm where "
        strSelectSql = strSelectSql & " a.unit_code = b.unit_code and b.unit_code = c.unit_code and B.unit_code = d.unit_code "
        strSelectSql = strSelectSql & " and a.Cust_ref = c.Cust_ref and a.amendment_No = c.amendment_No and a.Account_code=c.account_code And c.Active_Flag ='A' And c.Authorized_flag = 1"
        strSelectSql = strSelectSql & " and cm.unit_code=b.unit_code and cm.account_code =b.account_code and cm.item_Code=b.item_code and cm.cust_drgno=b.Cust_drgNo "
        strSelectSql = strSelectSql & " and a.account_code=b.Account_code and c.Cust_drgNo=b.Cust_drgNo and b.ITem_code = d.Item_code and a.Account_Code='" & Trim(pstrCustno) & "' and a.unit_code='" & gstrUNITID & "'"
        If Len(Trim(pstrConsCode)) > 0 Then
            strSelectSql = strSelectSql & " and a.consignee_code='" & Trim(pstrConsCode) & "' and a.Consignee_code=b.Consignee_code "
        End If
        strSelectSql = strSelectSql & " and b.ITem_code = c.Item_code and a.Cust_Ref in ('" & Trim(pstrRefNo) & "') and b.status = 1 and b.Schedule_flag =1 and b.Year_Month =  " & Validyrmon
        strSelectSql = strSelectSql & " and b.Item_Code in(Select a.Item_code from Item_MSt a,Itembal_mst b where a.unit_code = b.unit_code and a.unit_code='" & gstrUNITID & "' and a.Item_Main_grp in (" & Trim(pstrItemin) & ") and a.Item_code = b.Item_code and b.Location_code ='" & pstrstockLocation & "' and b.Cur_bal >0 and a.hold_flag =0 and a.Status = 'A'"
        If Len(Trim(pstrCondition)) > 0 Then
            strSelectSql = strSelectSql & " and a.Item_code not in(" & pstrCondition & "))"
        Else
            strSelectSql = strSelectSql & ")"
        End If
        strSelectSql = strSelectSql & " AND b.ITEM_CODE NOT IN "
        strSelectSql = strSelectSql & " ( "
        strSelectSql = strSelectSql & " SELECT I.ITEM_CODE FROM ITEM_MST I INNER JOIN GLOBAL_TOOL_MST G "
        strSelectSql = strSelectSql & " ON I.ITEM_CODE = G.GLOBAL_TOOL_CODE_FOR_ITEM_MASTER WHERE I.UNIT_CODE = '" & gstrUNITID & "' "
        strSelectSql = strSelectSql & " AND ITEM_MAIN_GRP ='P' "
        strSelectSql = strSelectSql & " AND G.CATEGORY IN ('MOULDING','PRESSTOOL') AND I.STATUS='A'"
        strSelectSql = strSelectSql & " ) "

        strSelectSql = strSelectSql & " UNION "

        strSelectSql = strSelectSql & " Select distinct b.Item_Code,c.Cust_DrgNo,c.Cust_Drg_Desc,d.Tariff_Code , c.EXTERNAL_SALESORDER_NO , C.AMENDMENT_NO , SHOP_CODE, CUR_BAL="
        strSelectSql = strSelectSql & " (	SELECT CUR_BAL FROM ITEMBAL_MST I WHERE B.UNIT_CODE = I.UNIT_CODE 	AND B.ITEM_CODE = I.ITEM_CODE AND I.UNIT_CODE='" & gstrUNITID & "' "
        strSelectSql = strSelectSql & " AND I.LOCATION_CODE='" & pstrstockLocation & "')"
        strSelectSql = strSelectSql & " from Cust_Ord_hdr a,MonthlyMktSchedule b,Cust_ord_dtl c,Item_Mst d , custitem_mst cm where "
        strSelectSql = strSelectSql & " a.unit_code = b.unit_code and b.unit_code = c.unit_code and B.unit_code = d.unit_code "
        strSelectSql = strSelectSql & " and a.Cust_ref = c.Cust_ref and a.amendment_No = c.amendment_No and a.Account_code=c.account_code And c.Active_Flag ='A' And c.Authorized_flag = 1"
        strSelectSql = strSelectSql & " and cm.unit_code=b.unit_code and cm.account_code =b.account_code and cm.item_Code=b.item_code and cm.cust_drgno=b.Cust_drgNo "
        strSelectSql = strSelectSql & " and a.account_code=b.Account_code and c.Cust_drgNo=b.Cust_drgNo and b.ITem_code = d.Item_code and a.Account_Code='" & Trim(pstrCustno) & "' and a.unit_code='" & gstrUNITID & "'"
        If Len(Trim(pstrConsCode)) > 0 Then
            strSelectSql = strSelectSql & " and a.consignee_code='" & Trim(pstrConsCode) & "' and a.Consignee_code=b.Consignee_code "
        End If
        strSelectSql = strSelectSql & " and b.ITem_code = c.Item_code and a.Cust_Ref in ('" & Trim(pstrRefNo) & "') and b.status = 1 and b.Schedule_flag =1 and b.Year_Month =  " & Validyrmon
        strSelectSql = strSelectSql & " and b.Item_Code in(Select a.Item_code from Item_MSt a,Itembal_mst b where a.unit_code = b.unit_code and a.unit_code='" & gstrUNITID & "' and a.Item_Main_grp in (" & Trim(pstrItemin) & ") and a.Item_code = b.Item_code and b.Location_code ='" & pstrstockLocation & "' and b.Cur_bal >0 and a.hold_flag =0 and a.Status = 'A'"
        If Len(Trim(pstrCondition)) > 0 Then
            strSelectSql = strSelectSql & " and a.Item_code not in(" & pstrCondition & "))"
        Else
            strSelectSql = strSelectSql & ")"
        End If
        strSelectSql = strSelectSql & "AND b.ITEM_CODE IN "
        strSelectSql = strSelectSql & "( "
        strSelectSql = strSelectSql & "SELECT I.ITEM_CODE FROM ITEM_MST I INNER JOIN GLOBAL_TOOL_MST G "
        strSelectSql = strSelectSql & "ON I.ITEM_CODE = G.GLOBAL_TOOL_CODE_FOR_ITEM_MASTER "
        strSelectSql = strSelectSql & "WHERE I.UNIT_CODE = '" & gstrUNITID & "' AND ITEM_MAIN_GRP ='P' "
        strSelectSql = strSelectSql & "AND G.CATEGORY IN ('MOULDING','PRESSTOOL') AND I.STATUS='A' AND G.VENDOR_CODE = '" & pstrCustno & "'"
        strSelectSql = strSelectSql & "AND G.TRANSFER_COMPLETED = 0 AND G.TRANSFER_TYPE in ('SALE','CUSTOMER_TOOL_SALE') AND ISNULL(G.INV_NO,0) = 0"
        strSelectSql = strSelectSql & ") "

        strSelectSql = strSelectSql & " UNION "

        strSelectSql = strSelectSql & "Select distinct b.Item_Code,c.Cust_DrgNo,c.Cust_Drg_Desc,d.Tariff_Code , c.EXTERNAL_SALESORDER_NO , C.AMENDMENT_NO , SHOP_CODE, CUR_BAL="
        strSelectSql = strSelectSql & " (	SELECT CUR_BAL FROM ITEMBAL_MST I WHERE B.UNIT_CODE = I.UNIT_CODE 	AND B.ITEM_CODE = I.ITEM_CODE AND I.UNIT_CODE='" & gstrUNITID & "' "
        strSelectSql = strSelectSql & " AND I.LOCATION_CODE='" & pstrstockLocation & "')"
        strSelectSql = strSelectSql & " from Cust_Ord_hdr a,Dailymktschedule   b,Cust_ord_dtl c,Item_Mst d , custitem_mst cm  where "
        strSelectSql = strSelectSql & " a.unit_code=b.unit_code and b.unit_code=c.unit_code and B.unit_code  = d.unit_code and "
        strSelectSql = strSelectSql & " a.Cust_ref = c.Cust_ref and a.amendment_No = c.amendment_No and a.Account_code=c.account_code"
        strSelectSql = strSelectSql & " and a.account_code=b.Account_code and c.Cust_drgNo=b.Cust_drgNo and b.ITem_code =d.ITem_code "
        strSelectSql = strSelectSql & " and cm.unit_code=b.unit_code and cm.account_code =b.account_code and cm.item_Code=b.item_code and cm.cust_drgno=B.Cust_drgNo and "
        strSelectSql = strSelectSql & " b.status = 1 and b.Schedule_Flag = 1 And c.Active_Flag ='A' And c.Authorized_flag = 1 and a.Account_Code='" & Trim(pstrCustno) & "' "
        strSelectSql = strSelectSql & " and a.unit_code='" & gstrUNITID & "'"
        If Len(Trim(pstrConsCode)) > 0 Then
            strSelectSql = strSelectSql & " and a.consignee_code='" & Trim(pstrConsCode) & "' and a.Consignee_code=b.Consignee_code "
        End If
        strSelectSql = strSelectSql & " and b.ITem_code = c.Item_code and a.Cust_Ref in('" & Trim(pstrRefNo) & "') and b.trans_date <= '" & strDate & "' And datepart(mm,b.trans_date) = '" & Month(CDate(strDate)) & "' And DatePart(yyyy, b.Trans_Date) = '" & Year(CDate(strDate)) & "'"
        strSelectSql = strSelectSql & " and b.Item_Code in(Select a.Item_code from Item_MSt a,Itembal_mst b where a.unit_code = b.unit_code and a.unit_code ='" & gstrUNITID & "' and a.Item_Main_grp in (" & Trim(pstrItemin) & ") and a.Item_code = b.Item_code and b.Location_code ='" & pstrstockLocation & "' and b.Cur_bal >0 and a.hold_flag =0 and a.Status = 'A'"
        If Len(Trim(pstrCondition)) > 0 Then
            strSelectSql = strSelectSql & " and a.Item_code not in( " & pstrCondition & "))"
        Else
            strSelectSql = strSelectSql & ")"
        End If
        strSelectSql = strSelectSql & " AND b.ITEM_CODE NOT IN "
        strSelectSql = strSelectSql & " ( "
        strSelectSql = strSelectSql & " SELECT I.ITEM_CODE FROM ITEM_MST I INNER JOIN GLOBAL_TOOL_MST G "
        strSelectSql = strSelectSql & " ON I.ITEM_CODE = G.GLOBAL_TOOL_CODE_FOR_ITEM_MASTER WHERE I.UNIT_CODE = '" & gstrUNITID & "' "
        strSelectSql = strSelectSql & " AND ITEM_MAIN_GRP ='P' "
        strSelectSql = strSelectSql & " AND G.CATEGORY IN ('MOULDING','PRESSTOOL') AND I.STATUS='A'"
        strSelectSql = strSelectSql & " ) "

        strSelectSql = strSelectSql & " UNION "

        strSelectSql = strSelectSql & "Select distinct b.Item_Code,c.Cust_DrgNo,c.Cust_Drg_Desc,d.Tariff_Code , c.EXTERNAL_SALESORDER_NO , C.AMENDMENT_NO , SHOP_CODE, CUR_BAL="
        strSelectSql = strSelectSql & " (	SELECT CUR_BAL FROM ITEMBAL_MST I WHERE B.UNIT_CODE = I.UNIT_CODE 	AND B.ITEM_CODE = I.ITEM_CODE AND I.UNIT_CODE='" & gstrUNITID & "' "
        strSelectSql = strSelectSql & " AND I.LOCATION_CODE='" & pstrstockLocation & "')"
        strSelectSql = strSelectSql & " from Cust_Ord_hdr a,Dailymktschedule   b,Cust_ord_dtl c,Item_Mst d , custitem_mst cm  where "
        strSelectSql = strSelectSql & " a.unit_code=b.unit_code and b.unit_code=c.unit_code and B.unit_code  = d.unit_code and "
        strSelectSql = strSelectSql & " a.Cust_ref = c.Cust_ref and a.amendment_No = c.amendment_No and a.Account_code=c.account_code"
        strSelectSql = strSelectSql & " and a.account_code=b.Account_code and c.Cust_drgNo=b.Cust_drgNo and b.ITem_code =d.ITem_code "
        strSelectSql = strSelectSql & " and cm.unit_code=b.unit_code and cm.account_code =b.account_code and cm.item_Code=b.item_code and cm.cust_drgno=B.Cust_drgNo and "
        strSelectSql = strSelectSql & " b.status = 1 and b.Schedule_Flag = 1 And c.Active_Flag ='A' And c.Authorized_flag = 1 and a.Account_Code='" & Trim(pstrCustno) & "' "
        strSelectSql = strSelectSql & " and a.unit_code='" & gstrUNITID & "'"
        If Len(Trim(pstrConsCode)) > 0 Then
            strSelectSql = strSelectSql & " and a.consignee_code='" & Trim(pstrConsCode) & "' and a.Consignee_code=b.Consignee_code "
        End If
        strSelectSql = strSelectSql & " and b.ITem_code = c.Item_code and a.Cust_Ref in('" & Trim(pstrRefNo) & "') and b.trans_date <= '" & strDate & "' And datepart(mm,b.trans_date) = '" & Month(CDate(strDate)) & "' And DatePart(yyyy, b.Trans_Date) = '" & Year(CDate(strDate)) & "'"
        strSelectSql = strSelectSql & " and b.Item_Code in(Select a.Item_code from Item_MSt a,Itembal_mst b where a.unit_code = b.unit_code and a.unit_code ='" & gstrUNITID & "' and a.Item_Main_grp in (" & Trim(pstrItemin) & ") and a.Item_code = b.Item_code and b.Location_code ='" & pstrstockLocation & "' and b.Cur_bal >0 and a.hold_flag =0 and a.Status = 'A'"
        If Len(Trim(pstrCondition)) > 0 Then
            strSelectSql = strSelectSql & " and a.Item_code not in( " & pstrCondition & "))"
        Else
            strSelectSql = strSelectSql & ")"
        End If
        strSelectSql = strSelectSql & "AND b.ITEM_CODE IN "
        strSelectSql = strSelectSql & "( "
        strSelectSql = strSelectSql & "SELECT I.ITEM_CODE FROM ITEM_MST I INNER JOIN GLOBAL_TOOL_MST G "
        strSelectSql = strSelectSql & "ON I.ITEM_CODE = G.GLOBAL_TOOL_CODE_FOR_ITEM_MASTER "
        strSelectSql = strSelectSql & "WHERE I.UNIT_CODE = '" & gstrUNITID & "' AND ITEM_MAIN_GRP ='P' "
        strSelectSql = strSelectSql & "AND G.CATEGORY IN ('MOULDING','PRESSTOOL') AND I.STATUS='A' AND G.VENDOR_CODE = '" & pstrCustno & "'"
        strSelectSql = strSelectSql & "AND G.TRANSFER_COMPLETED = 0 AND G.TRANSFER_TYPE in ('SALE','CUSTOMER_TOOL_SALE') AND ISNULL(G.INV_NO,0) = 0"
        strSelectSql = strSelectSql & ") "
        Return strSelectSql
    End Function

    Public Function makeSelectSql_Tool(ByRef pstrCustno As String, ByRef pstrRefNo As String, _
                                      ByRef pstrAmmNo As String, ByRef effectyrmon As String, _
                                      ByRef Validyrmon As String, ByRef pstrstockLocation As String, _
                                      ByRef strDate As String, ByRef pstrItemin As String, _
                                      Optional ByRef pstrCondition As String = "", _
                                      Optional ByRef pstrConsCode As String = "") As String

        Dim strSelectSql As String = String.Empty
        strDate = getDateForDB(strDate)
        strSelectSql = "Select distinct b.Item_Code,c.Cust_DrgNo,c.Cust_Drg_Desc,d.Tariff_Code , c.EXTERNAL_SALESORDER_NO , C.AMENDMENT_NO , SHOP_CODE, CUR_BAL="
        strSelectSql = strSelectSql & " ( SELECT CUR_BAL FROM ITEMBAL_MST I WHERE B.UNIT_CODE = I.UNIT_CODE 	AND B.ITEM_CODE = I.ITEM_CODE AND I.UNIT_CODE='" & gstrUNITID & "' "
        strSelectSql = strSelectSql & " AND I.LOCATION_CODE='" & pstrstockLocation & "')"
        strSelectSql = strSelectSql & " from Cust_Ord_hdr a,MonthlyMktSchedule b,Cust_ord_dtl c,Item_Mst d , custitem_mst cm where "
        strSelectSql = strSelectSql & " a.unit_code = b.unit_code and b.unit_code = c.unit_code and B.unit_code = d.unit_code "
        strSelectSql = strSelectSql & " and a.Cust_ref = c.Cust_ref and a.amendment_No = c.amendment_No and a.Account_code=c.account_code And c.Active_Flag ='A' And c.Authorized_flag = 1"
        strSelectSql = strSelectSql & " and cm.unit_code=b.unit_code and cm.account_code =b.account_code and cm.item_Code=b.item_code and cm.cust_drgno=b.Cust_drgNo "
        strSelectSql = strSelectSql & " and a.account_code=b.Account_code and c.Cust_drgNo=b.Cust_drgNo and b.ITem_code = d.Item_code and a.Account_Code='" & Trim(pstrCustno) & "' and a.unit_code='" & gstrUNITID & "'"
        If Len(Trim(pstrConsCode)) > 0 Then
            strSelectSql = strSelectSql & " and a.consignee_code='" & Trim(pstrConsCode) & "' and a.Consignee_code=b.Consignee_code "
        End If
        strSelectSql = strSelectSql & " and b.ITem_code = c.Item_code and a.Cust_Ref in ('" & Trim(pstrRefNo) & "') and b.status = 1 and b.Schedule_flag =1 and b.Year_Month =  " & Validyrmon
        strSelectSql = strSelectSql & " and b.Item_Code in(Select a.Item_code from Item_MSt a,Itembal_mst b where a.unit_code = b.unit_code and a.unit_code='" & gstrUNITID & "' and a.Item_Main_grp in (" & Trim(pstrItemin) & ") and a.Item_code = b.Item_code and b.Location_code ='" & pstrstockLocation & "' and b.Cur_bal >0 and a.hold_flag =0 and a.Status = 'A'"
        If Len(Trim(pstrCondition)) > 0 Then
            strSelectSql = strSelectSql & " and a.Item_code not in(" & pstrCondition & "))"
        Else
            strSelectSql = strSelectSql & ")"
        End If
        strSelectSql = strSelectSql & " AND b.ITEM_CODE NOT IN "
        strSelectSql = strSelectSql & " ( "
        strSelectSql = strSelectSql & " SELECT I.ITEM_CODE FROM ITEM_MST I INNER JOIN GLOBAL_TOOL_MST G "
        strSelectSql = strSelectSql & " ON I.ITEM_CODE = G.GLOBAL_TOOL_CODE_FOR_ITEM_MASTER WHERE I.UNIT_CODE = '" & gstrUNITID & "' "
        strSelectSql = strSelectSql & " AND ITEM_MAIN_GRP ='P' "
        strSelectSql = strSelectSql & " AND G.CATEGORY IN ('MOULDING','PRESSTOOL') AND I.STATUS='A'"
        strSelectSql = strSelectSql & " ) "

        strSelectSql = strSelectSql & " UNION "

        strSelectSql = strSelectSql & "Select distinct b.Item_Code,c.Cust_DrgNo,c.Cust_Drg_Desc,d.Tariff_Code , c.EXTERNAL_SALESORDER_NO , C.AMENDMENT_NO , SHOP_CODE, CUR_BAL="
        strSelectSql = strSelectSql & " (	SELECT CUR_BAL FROM ITEMBAL_MST I WHERE B.UNIT_CODE = I.UNIT_CODE 	AND B.ITEM_CODE = I.ITEM_CODE AND I.UNIT_CODE='" & gstrUNITID & "' "
        strSelectSql = strSelectSql & " AND I.LOCATION_CODE='" & pstrstockLocation & "')"
        strSelectSql = strSelectSql & " from Cust_Ord_hdr a,Dailymktschedule   b,Cust_ord_dtl c,Item_Mst d , custitem_mst cm  where "
        strSelectSql = strSelectSql & " a.unit_code=b.unit_code and b.unit_code=c.unit_code and B.unit_code  = d.unit_code and "
        strSelectSql = strSelectSql & " a.Cust_ref = c.Cust_ref and a.amendment_No = c.amendment_No and a.Account_code=c.account_code"
        strSelectSql = strSelectSql & " and a.account_code=b.Account_code and c.Cust_drgNo=b.Cust_drgNo and b.ITem_code =d.ITem_code "
        strSelectSql = strSelectSql & " and cm.unit_code=b.unit_code and cm.account_code =b.account_code and cm.item_Code=b.item_code and cm.cust_drgno=B.Cust_drgNo and "
        strSelectSql = strSelectSql & " b.status = 1 and b.Schedule_Flag = 1 And c.Active_Flag ='A' And c.Authorized_flag = 1 and a.Account_Code='" & Trim(pstrCustno) & "' "
        strSelectSql = strSelectSql & " and a.unit_code='" & gstrUNITID & "'"
        If Len(Trim(pstrConsCode)) > 0 Then
            strSelectSql = strSelectSql & " and a.consignee_code='" & Trim(pstrConsCode) & "' and a.Consignee_code=b.Consignee_code "
        End If
        strSelectSql = strSelectSql & " and b.ITem_code = c.Item_code and a.Cust_Ref in('" & Trim(pstrRefNo) & "') and b.trans_date <= '" & strDate & "' And datepart(mm,b.trans_date) = '" & Month(CDate(strDate)) & "' And DatePart(yyyy, b.Trans_Date) = '" & Year(CDate(strDate)) & "'"
        strSelectSql = strSelectSql & " and b.Item_Code in(Select a.Item_code from Item_MSt a,Itembal_mst b where a.unit_code = b.unit_code and a.unit_code ='" & gstrUNITID & "' and a.Item_Main_grp in (" & Trim(pstrItemin) & ") and a.Item_code = b.Item_code and b.Location_code ='" & pstrstockLocation & "' and b.Cur_bal >0 and a.hold_flag =0 and a.Status = 'A'"
        If Len(Trim(pstrCondition)) > 0 Then
            strSelectSql = strSelectSql & " and a.Item_code not in( " & pstrCondition & "))"
        Else
            strSelectSql = strSelectSql & ")"
        End If
        strSelectSql = strSelectSql & " AND b.ITEM_CODE NOT IN "
        strSelectSql = strSelectSql & " ( "
        strSelectSql = strSelectSql & " SELECT I.ITEM_CODE FROM ITEM_MST I INNER JOIN GLOBAL_TOOL_MST G "
        strSelectSql = strSelectSql & " ON I.ITEM_CODE = G.GLOBAL_TOOL_CODE_FOR_ITEM_MASTER WHERE I.UNIT_CODE = '" & gstrUNITID & "' "
        strSelectSql = strSelectSql & " AND ITEM_MAIN_GRP ='P' "
        strSelectSql = strSelectSql & " AND G.CATEGORY IN ('MOULDING','PRESSTOOL') AND I.STATUS='A'"
        strSelectSql = strSelectSql & " ) "

        Return strSelectSql
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

    Private Sub chkall_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkall.CheckedChanged
        Dim intCount As Integer = 0
        If chkall.Checked Then
            For intCount = 0 To Me.lvwItemCode.Items.Count - 1
                lvwItemCode.Items.Item(intCount).Checked = True
            Next intCount
        Else
            For intCount = 0 To Me.lvwItemCode.Items.Count - 1
                lvwItemCode.Items.Item(intCount).Checked = False
            Next intCount
        End If
    End Sub
End Class