Option Strict Off
Option Explicit On
Friend Class frmMKTTRN0062
	Inherits System.Windows.Forms.Form
	'-----------------------------------------------------------------------
    ' Copyright (c)     :   MIND Ltd.
    ' Form Name         :   frmMKTTRN0062
    ' Function Name     :   CDP Analysis
    ' Created By        :   Shubhra Verma
    ' Created On        :   08 Sep 2008
    ' Issue ID          :   eMpro-20080911-21453
	'-----------------------------------------------------------------------
    'Revised By     :   Shubhra Verma
    'Revised On     :   06 Jun 2011
    'Description    :   Multi Unit Changes
    '-----------------------------------------------------------------------
    Dim mintIndex As Short
    Dim mstrDocNo As String
    Dim mstrConsigneeCode As String
    Dim mstrCustPart As String
    Dim mstrReleaseDate As String
    Dim mstrWHCode As String
    Dim MSTRSHIPMENTDATE As String
    Dim mSTRCUSTCODE As String
    Dim MBLNSHIPMENTTHRUWH As Boolean
    Dim MBLNDAILYPULLFLAG As Boolean
    Dim mblnChk As Boolean
    Dim mblncallItemChk As Boolean
    Private Sub chkClose_Click()
        '-----------------------------------------------------------------------------------------------
        ' Author        : Shubhra Verma
        ' Arguments     : Nil
        ' Return Value  : Nil
        ' Function      : To Close the Form
        ' Datetime      : 08-Sep-2008
        '---------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Me.Close()
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0062_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        '-----------------------------------------------------------------------------------------------
        ' Author        : Shubhra Verma
        ' Arguments     : Nil
        ' Return Value  : Nil
        ' Function      : Set Property CheckFormName as per Rule book
        ' Datetime      : 08-Sep-2008
        '---------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        mdifrmMain.CheckFormName = mintIndex
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0062_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        '---------------------------------------------------------------------------------------------------
        ' Author        : Shubhra Verma
        ' Arguments     : Nil
        ' Return Value  : Nil
        ' Function      : Set Property NODEFONTBOLD as per Rule book
        ' Datetime      : 08-Sep-2008
        '---------------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Me.Tag) = False
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0062_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0 Then
            Call ctlFormHeader_ClickEvent(ctlFormHeader, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0062_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        '---------------------------------------------------------------------------------------------------
        ' Author        : Shubhra Verma
        ' Arguments     : Nil
        ' Return Value  : Nil
        ' Function      : Set Property AddFormNameToWindowList &
        '                 Function FitToClient as per Rule book
        ' Datetime      : 08-Sep-2008
        '-----------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        mintIndex = mdifrmMain.AddFormNameToWindowList(Me.Tag)
        Call FitToClient(Me, fraMain, ctlFormHeader, fracmdgrp, 600)
        lblFormula5.Text = "If Desp Date >= Upload Date And Remaining WH Stock < 0, Then Only System Proposes to Despatch."
        lstDocNo.LabelEdit = False
        lstDocNo.CheckBoxes = True
        lstDocNo.View = System.Windows.Forms.View.Details
        lstDocNo.Columns.Insert(0, "", "Doc No", -2)
        lstDocNo.Columns.Insert(1, "", "Cust Code", -2)
        lstDocNo.Columns.Insert(2, "", "Cust Name", -2)
        lstDocNo.Columns.Item(0).Width = VB6.TwipsToPixelsX(1300)
        lstDocNo.Columns.Item(1).Width = VB6.TwipsToPixelsX(1300)
        lstDocNo.Columns.Item(2).Width = VB6.TwipsToPixelsX(2500)
        lstConsigneeCode.View = System.Windows.Forms.View.Details
        lstConsigneeCode.LabelEdit = False
        lstConsigneeCode.CheckBoxes = True
        lstConsigneeCode.Columns.Insert(0, "", "Consignee Code", -2)
        lstConsigneeCode.Columns.Item(0).Width = VB6.TwipsToPixelsX(1300)
        lstConsigneeCode.Columns.Insert(1, "", "Consignee Name", -2)
        lstConsigneeCode.Columns.Item(1).Width = VB6.TwipsToPixelsX(1300)
        lstCustPartNo.View = System.Windows.Forms.View.Details
        lstCustPartNo.GridLines = True
        lstCustPartNo.Sorting = SortOrder.None
        lstCustPartNo.Columns.Insert(0, "", "CustPart No", -2)
        lstCustPartNo.Columns.Item(0).Width = VB6.TwipsToPixelsX(1700)
        lstCustPartNo.Columns.Insert(1, "", "CustPart Desc", -2)
        lstCustPartNo.Columns.Item(1).Width = VB6.TwipsToPixelsX(1200)
        lstProposal.View = System.Windows.Forms.View.Details
        lstProposal.GridLines = True
        lstProposal.Columns.Insert(0, "", "Release Date", -2)
        lstProposal.Columns.Insert(1, "", "Release Qty", -2)
        lstProposal.Columns.Insert(2, "", "WH Code", -2)
        lstProposal.Columns.Insert(3, "", "WHStock", -2)
        lstProposal.Columns.Insert(4, "", "InTransit Qty", -2)
        lstProposal.Columns.Insert(5, "", "Safety Stk", -2)
        lstProposal.Columns.Insert(6, "", "Desp Date", -2)
        lstProposal.Columns.Insert(7, "", "Desp Qty", -2)
        lstProposal.Columns.Insert(8, "", "Stk AsOn", -2)
        lstProposal.Columns.Item(0).Width = VB6.TwipsToPixelsX(1500)
        lstProposal.Columns.Item(1).Width = VB6.TwipsToPixelsX(1200)
        lstProposal.Columns.Item(2).Width = VB6.TwipsToPixelsX(1000)
        lstProposal.Columns.Item(3).Width = VB6.TwipsToPixelsX(1000)
        lstProposal.Columns.Item(4).Width = VB6.TwipsToPixelsX(1200)
        lstProposal.Columns.Item(5).Width = VB6.TwipsToPixelsX(1000)
        lstProposal.Columns.Item(6).Width = VB6.TwipsToPixelsX(1200)
        lstProposal.Columns.Item(7).Width = VB6.TwipsToPixelsX(1200)
        lstProposal.Columns.Item(8).Width = VB6.TwipsToPixelsX(1200)
        lstInTransitDetails.View = System.Windows.Forms.View.Details
        lstInTransitDetails.GridLines = True
        lstInTransitDetails.Columns.Insert(0, "", "Invoice Date", -2)
        lstInTransitDetails.Columns.Insert(1, "", "Invoice No", -2)
        lstInTransitDetails.Columns.Insert(2, "", "Invoice Qty", -2)
        lstInTransitDetails.Columns.Item(0).Width = (lstInTransitDetails.Width - 20) / lstInTransitDetails.Columns.Count
        lstInTransitDetails.Columns.Item(1).Width = (lstInTransitDetails.Width - 20) / lstInTransitDetails.Columns.Count
        lstInTransitDetails.Columns.Item(2).Width = (lstInTransitDetails.Width - 20) / lstInTransitDetails.Columns.Count
        lstFixedSafetyStock.View = System.Windows.Forms.View.Details
        lstFixedSafetyStock.GridLines = True
        lstFixedSafetyStock.Columns.Insert(0, "", "Daily Pull Rate", -2)
        lstFixedSafetyStock.Columns.Insert(1, "", "Safety Days", -2)
        lstFixedSafetyStock.Columns.Insert(2, "", "Safety Stock", -2)
        lstFixedSafetyStock.Columns.Item(0).Width = (lstFixedSafetyStock.Width - 10) / lstFixedSafetyStock.Columns.Count
        lstFixedSafetyStock.Columns.Item(1).Width = (lstFixedSafetyStock.Width - 10) / lstFixedSafetyStock.Columns.Count
        lstFixedSafetyStock.Columns.Item(2).Width = (lstFixedSafetyStock.Width - 10) / lstFixedSafetyStock.Columns.Count
        lstDynamicSafetyStock.View = System.Windows.Forms.View.Details
        lstDynamicSafetyStock.GridLines = True
        lstDynamicSafetyStock.Columns.Insert(0, "", "Cal. On", -2)
        lstDynamicSafetyStock.Columns.Insert(1, "", "Cal. On CallOffs of", -2)
        lstDynamicSafetyStock.Columns.Insert(2, "", "W/A Days", -2)
        lstDynamicSafetyStock.Columns.Insert(3, "", "Sum of Release Qty", -2)
        lstDynamicSafetyStock.Columns.Insert(4, "", "Safety Stock Per Day", -2)
        lstDynamicSafetyStock.Columns.Insert(5, "", "Safety Days", -2)
        lstDynamicSafetyStock.Columns.Insert(6, "", "Safety Stock", -2)
        lstDynamicSafetyStock.Columns.Item(0).Width = (lstDynamicSafetyStock.Width - 10) / lstDynamicSafetyStock.Columns.Count
        lstDynamicSafetyStock.Columns.Item(1).Width = (lstDynamicSafetyStock.Width - 10) / lstDynamicSafetyStock.Columns.Count
        lstDynamicSafetyStock.Columns.Item(2).Width = (lstDynamicSafetyStock.Width - 10) / lstDynamicSafetyStock.Columns.Count
        lstDynamicSafetyStock.Columns.Item(3).Width = (lstDynamicSafetyStock.Width - 10) / lstDynamicSafetyStock.Columns.Count
        lstDynamicSafetyStock.Columns.Item(4).Width = (lstDynamicSafetyStock.Width) / lstDynamicSafetyStock.Columns.Count
        lstDynamicSafetyStock.Columns.Item(5).Width = (lstDynamicSafetyStock.Width - 10) / lstDynamicSafetyStock.Columns.Count
        lstDynamicSafetyStock.Columns.Item(6).Width = (lstDynamicSafetyStock.Width - 10) / lstDynamicSafetyStock.Columns.Count
        lstShipment.Columns.Insert(0, "", "Bag Qty", -2)
        lstShipment.Columns.Insert(1, "", "Rmng WHStk", -2)
        lstShipment.Columns.Insert(2, "", "Min Safety Stk", -2)
        lstShipment.Columns.Insert(3, "", "Max Safety Stk", -2)
        lstShipment.Columns.Insert(4, "", "Desp Qty", -2)
        lstShipment.Columns.Insert(5, "", "Act Desp Qty", -2)
        lstShipment.Columns.Insert(6, "", "NoOfBags", -2)
        lstShipment.Columns.Insert(7, "", "Transit Days", -2)
        lstShipment.Columns.Insert(8, "", "Buffer Days", -2)
        lstShipment.Columns.Insert(9, "", "Upload Date", -2)
        lstShipment.View = System.Windows.Forms.View.Details
        lstShipment.GridLines = True
        lstShipment.Columns.Item(0).Width = VB6.TwipsToPixelsX(800)
        lstShipment.Columns.Item(1).Width = VB6.TwipsToPixelsX(1300)
        lstShipment.Columns.Item(2).Width = VB6.TwipsToPixelsX(1300)
        lstShipment.Columns.Item(3).Width = VB6.TwipsToPixelsX(1300)
        lstShipment.Columns.Item(4).Width = VB6.TwipsToPixelsX(1000)
        lstShipment.Columns.Item(5).Width = VB6.TwipsToPixelsX(1200)
        lstShipment.Columns.Item(6).Width = VB6.TwipsToPixelsX(1000)
        lstShipment.Columns.Item(7).Width = VB6.TwipsToPixelsX(1200)
        lstShipment.Columns.Item(8).Width = VB6.TwipsToPixelsX(1200)
        lstShipment.Columns.Item(9).Width = VB6.TwipsToPixelsX(1200)
        Call FillDocNo()
        cmbSearch.Items.Add(("Doc No"))
        cmbSearch.Items.Add(("Customer Code"))
        cmbSearch.Items.Add(("Customer Name"))
        cmbSearch.Items.Add(("Consignee Code"))
        cmbSearch.Items.Add(("Consignee Name"))
        cmbSearch.Items.Add(("Cust Part No"))
        cmbSearch.Items.Add(("Cust Part Desc"))
        cmbSearch.SelectedText = "Doc No"
        ToolTip1.SetToolTip(lstProposal, "Double Click On Row To View Formula")
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0062_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '-------------------------------------------------------------------------------------------
        ' Author        : Shubhra Verma
        ' Arguments     : Cancel as Integer
        ' Return Value  : NIL
        ' Function      : RemoveFormNameFromWindowList
        '                 Release form Object Memory from Database.
        ' Datetime      : 08-Sep-2008
        '--------------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Me.Dispose() 'Assign form to nothing
        mdifrmMain.RemoveFormNameFromWindowList = mintIndex
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub ctlFormHeader_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        On Error GoTo Err_Renamed
        Call ShowHelp("underconstruction.htm")
        Exit Sub
Err_Renamed:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Function FillDocNo() As Boolean
        '---------------------------------------------------------------------------------------------------
        ' Author        : Shubhra Verma
        ' Arguments     : Nil
        ' Return Value  : Nil
        ' Function      : fill Doc No, Customer Code And Customer Name In lstDocNo ListView
        ' Datetime      : 10-Sep-2008
        '-----------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim strsql As String
        Dim rsDocNo As ADODB.Recordset
        Dim DocList As System.Windows.Forms.ListView
        Dim intROWCOUNT As Integer
        lstDocNo.Items.Clear()
        lstDocNo.View = System.Windows.Forms.View.Details
        lstDocNo.GridLines = True
        lstDocNo.Sorting = SortOrder.None
        intROWCOUNT = 0
        strsql = "SELECT DISTINCT S.DOC_NO, H.CUST_CODE,C.CUST_NAME" & _
           " FROM SCHEDULE_UPLOAD_DARWIN_HDR H, CUSTOMER_MST C," & _
           " SCHEDULEPROPOSALCALCULATIONS S" & _
           " Where H.UNIT_CODE = C.UNIT_CODE AND C.UNIT_CODE = S.UNIT_CODE" & _
           " AND H.UNIT_CODE = '" & gstrUNITID & "' AND S.Doc_No = H.Doc_No" & _
           " AND C.CUSTOMER_CODE = H.CUST_CODE" & _
           " Union" & _
           " SELECT DISTINCT S.DOC_NO, H.CUST_CODE,C.CUST_NAME" & _
           " FROM SCHEDULE_UPLOAD_DARWINEDIFACT_HDR H, CUSTOMER_MST C," & _
           " SCHEDULEPROPOSALCALCULATIONS S" & _
           " Where H.UNIT_CODE = C.UNIT_CODE AND C.UNIT_CODE = S.UNIT_CODE" & _
           " AND H.UNIT_CODE = '" & gstrUNITID & "' AND S.Doc_No = H.Doc_No" & _
           " AND C.CUSTOMER_CODE = H.CUST_CODE" & _
           " Union" & _
           " SELECT DISTINCT S.DOC_NO, H.CUST_CODE, C.CUST_NAME" & _
           " FROM SCHEDULE_UPLOAD_COVISINT_HDR H, CUSTOMER_MST C, SCHEDULEPROPOSALCALCULATIONS S" & _
           " Where H.UNIT_CODE = C.UNIT_CODE AND C.UNIT_CODE = S.UNIT_CODE" & _
           " AND H.UNIT_CODE = '" & gstrUNITID & "' AND S.Doc_No = H.Doc_No" & _
           " AND C.CUSTOMER_CODE = H.CUST_CODE" & " ORDER BY S.DOC_NO DESC"
        rsDocNo = New ADODB.Recordset
        rsDocNo.Open(strsql, mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
        If rsDocNo.RecordCount > 0 Then
            rsDocNo.MoveFirst()
            intROWCOUNT = 0
            While Not rsDocNo.EOF
                lstDocNo.Items.Add(rsDocNo.Fields("doc_no").Value)
                If lstDocNo.Items.Item(intROWCOUNT).SubItems.Count > 1 Then
                    lstDocNo.Items.Item(intROWCOUNT).SubItems(1).Text = rsDocNo.Fields("CUST_CODE").Value
                Else
                    lstDocNo.Items.Item(intROWCOUNT).SubItems.Insert(1, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsDocNo.Fields("CUST_CODE").Value))
                End If
                If lstDocNo.Items.Item(intROWCOUNT).SubItems.Count > 2 Then
                    lstDocNo.Items.Item(intROWCOUNT).SubItems(2).Text = rsDocNo.Fields("CUST_NAME").Value
                Else
                    lstDocNo.Items.Item(intROWCOUNT).SubItems.Insert(2, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsDocNo.Fields("CUST_NAME").Value))
                End If
                intROWCOUNT = intROWCOUNT + 1
                rsDocNo.MoveNext()
            End While
        End If
        If lstDocNo.Items.Count > 0 Then
            lstDocNo.Items.Item(0).Checked = True
            mstrDocNo = lstDocNo.Items.Item(0).SubItems(0).Text.ToString
            mSTRCUSTCODE = lstDocNo.Items.Item(0).SubItems(1).Text
            'FillConsignee(CInt(lstDocNo.Items.Item(0).Text))
        End If
        If rsDocNo.State Then rsDocNo.Close() : rsDocNo = Nothing
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub SingleSelection(ByVal ListName As System.Windows.Forms.ListView, ByVal intINDEX As Integer)
        '---------------------------------------------------------------------------------------------------
        ' Author        : Shubhra Verma
        ' Arguments     : ListName, intINDEX
        ' Return Value  : Nil
        ' Function      : to Allow Only Single Selection In The ListView
        ' Datetime      : 09-Sep-2008
        '-----------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim intRowCount As Integer
        Dim Item As System.Windows.Forms.ListViewItem = ListName.Items(intINDEX)
        If mblnChk = False Then
            For intRowCount = 0 To ListName.Items.Count - 1
                mblnChk = True
                ListName.Items.Item(intRowCount).Checked = False
            Next
            mblnChk = True
            Item.Checked = True
            mblnChk = False
            Exit Sub
        Else
            Exit Sub
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub lstConsigneeCode_ItemCheck(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.ItemCheckEventArgs) Handles lstConsigneeCode.ItemCheck
        Dim Item As System.Windows.Forms.ListViewItem = lstConsigneeCode.Items(eventArgs.Index)
        On Error GoTo ErrHandler
        If mblncallItemChk = True Then
            Call SingleSelection(lstConsigneeCode, Item.Index)
        End If
        mblncallItemChk = True
        lstCustPartNo.Items.Clear()
        lstProposal.Items.Clear()
        lstInTransitDetails.Items.Clear()
        lstFixedSafetyStock.Items.Clear()
        lstDynamicSafetyStock.Items.Clear()
        lstShipment.Items.Clear()
        Call FILLCUSTPARTNO(mstrDocNo, Item.Text)
        mstrConsigneeCode = Item.Text
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub lstCustPartNo_ItemCheck(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.ItemCheckEventArgs) Handles lstCustPartNo.ItemCheck
        Dim Item As System.Windows.Forms.ListViewItem = lstCustPartNo.Items(eventArgs.Index)
        On Error GoTo ErrHandler
        If mblncallItemChk = True Then
            Call SingleSelection(lstCustPartNo, Item.Index)
        End If
        mblncallItemChk = True
        lstProposal.Items.Clear()
        lstInTransitDetails.Items.Clear()
        lstFixedSafetyStock.Items.Clear()
        lstDynamicSafetyStock.Items.Clear()
        lstShipment.Items.Clear()
        mstrCustPart = Item.Text
        Call FILLPROPOSALDETAILS(mstrDocNo, mstrConsigneeCode, Item.Text)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub lstDocNo_ItemCheck(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.ItemCheckEventArgs) Handles lstDocNo.ItemCheck
        Dim Item As System.Windows.Forms.ListViewItem = lstDocNo.Items(eventArgs.Index)
        On Error GoTo ErrHandler
        Dim rsSHIPMENTTHRUWH As ADODB.Recordset
        Dim strsql As String
        Dim intRowCount As Integer
        Call SingleSelection(lstDocNo, eventArgs.Index)
        mstrDocNo = Item.Text
        mSTRCUSTCODE = lstDocNo.Items.Item(Item.Index).SubItems(1).Text
        lstConsigneeCode.Items.Clear()
        lstCustPartNo.Items.Clear()
        lstProposal.Items.Clear()
        lstInTransitDetails.Items.Clear()
        lstFixedSafetyStock.Items.Clear()
        lstDynamicSafetyStock.Items.Clear()
        lstShipment.Items.Clear()
        rsSHIPMENTTHRUWH = New ADODB.Recordset
        strsql = "SELECT SHIPMENTTHRUWH FROM CUSTOMER_MST" & _
            " WHERE UNIT_CODE = '" & gstrUNITID & "' AND CUSTOMER_CODE = '" & mSTRCUSTCODE & "'"
        rsSHIPMENTTHRUWH.Open(strsql, mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
        If rsSHIPMENTTHRUWH.RecordCount > 0 Then
            rsSHIPMENTTHRUWH.MoveFirst()
            MBLNSHIPMENTTHRUWH = rsSHIPMENTTHRUWH.Fields("SHIPMENTTHRUWH").Value
        End If
        If rsSHIPMENTTHRUWH.State Then rsSHIPMENTTHRUWH.Close() : rsSHIPMENTTHRUWH = Nothing
        FillConsignee(CInt(Item.Text))
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub FillConsignee(ByVal intDocNo As Integer)
        '---------------------------------------------------------------------------------------------------
        ' Author        : Shubhra Verma
        ' Arguments     : intDocNo
        ' Return Value  : Nil
        ' Function      : fill Consignee Code And Consignee Name In lstConsigneeCode ListView
        ' Datetime      : 10-Sep-2008
        '-----------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim strsql As String
        Dim RSconsignee As ADODB.Recordset
        Dim LNGROWCOUNT As Integer
        Dim blnReturnVal As Boolean
        Dim rsSHIPMENTTHRUWH As ADODB.Recordset
        LNGROWCOUNT = 0
        lblFormula1.Text = "" : lblFormula2.Text = "" : lblFormula3.Text = "" : lblFormula4.Text = ""
        lstConsigneeCode.Items.Clear()
        strsql = "SELECT DISTINCT S.CONSIGNEE_CODE,CUST_NAME " & _
       " FROM SCHEDULEPROPOSALCALCULATIONS S, CUSTOMER_MST C" & _
       " Where C.UNIT_CODE = S.UNIT_CODE AND C.CUSTOMER_CODE = S.CONSIGNEE_CODE" & _
       " AND S.DOC_NO = '" & intDocNo & "' AND S.UNIT_CODE = '" & gstrUNITID & "'"
        RSconsignee = New ADODB.Recordset
        RSconsignee.Open(strsql, mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
        If RSconsignee.RecordCount > 0 Then
            RSconsignee.MoveFirst()
            LNGROWCOUNT = 0
            While Not RSconsignee.EOF
                mblncallItemChk = False
                lstConsigneeCode.Items.Add(RSconsignee.Fields("CONSIGNEE_CODE").Value)
                If lstConsigneeCode.Items.Item(LNGROWCOUNT).SubItems.Count > 1 Then
                    lstConsigneeCode.Items.Item(LNGROWCOUNT).SubItems(1).Text = RSconsignee.Fields("CUST_NAME").Value
                Else
                    lstConsigneeCode.Items.Item(LNGROWCOUNT).SubItems.Insert(1, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, RSconsignee.Fields("CUST_NAME").Value))
                End If
                LNGROWCOUNT = LNGROWCOUNT + 1
                RSconsignee.MoveNext()
            End While
        End If
        If RSconsignee.State Then RSconsignee.Close() : RSconsignee = Nothing
        rsSHIPMENTTHRUWH = New ADODB.Recordset
        strsql = "SELECT SHIPMENTTHRUWH FROM CUSTOMER_MST" & _
        " WHERE UNIT_CODE = '" & gstrUNITID & "' AND CUSTOMER_CODE = '" & mSTRCUSTCODE & "'"
        rsSHIPMENTTHRUWH.Open(strsql, mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
        If rsSHIPMENTTHRUWH.RecordCount > 0 Then
            rsSHIPMENTTHRUWH.MoveFirst()
            MBLNSHIPMENTTHRUWH = rsSHIPMENTTHRUWH.Fields("SHIPMENTTHRUWH").Value
        End If
        If rsSHIPMENTTHRUWH.State Then rsSHIPMENTTHRUWH.Close() : rsSHIPMENTTHRUWH = Nothing
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub FILLCUSTPARTNO(ByVal intDocNo As Integer, ByVal strCONSIGNEE As String)
        '---------------------------------------------------------------------------------------------------
        ' Author        : Shubhra Verma
        ' Arguments     : intDocNo, strCONSIGNEE
        ' Return Value  : Nil
        ' Function      : fill Cust Part No And Cust Part Desc In lstCustPartNo ListView
        ' Datetime      : 10-Sep-2008
        '-----------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim strsql As String
        Dim RSCUSTPART As ADODB.Recordset
        Dim LNGROWCOUNT As Integer
        LNGROWCOUNT = 0
        lstCustPartNo.Items.Clear()
        strsql = "SELECT DISTINCT S.ITEM_CODE,C.DRG_DESC" & _
       " From SCHEDULEPROPOSALCALCULATIONS S, CUSTITEM_MST C" & _
       " WHERE S.UNIT_CODE = C.UNIT_CODE AND S.UNIT_CODE = '" & gstrUNITID & "' AND" & _
       " S.DOC_NO = '" & intDocNo & "' AND S.CONSIGNEE_CODE = '" & strCONSIGNEE & "'" & _
       " AND S.ITEM_CODE = C.CUST_DRGNO AND C.ACTIVE = 1 AND C.SCHUPLDREQD = 1" & _
       " AND S.CONSIGNEE_CODE = C.ACCOUNT_CODE ORDER BY S.ITEM_CODE"
        RSCUSTPART = New ADODB.Recordset
        RSCUSTPART.Open(strsql, mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
        If RSCUSTPART.RecordCount > 0 Then
            RSCUSTPART.MoveFirst()
            LNGROWCOUNT = 0
            While Not RSCUSTPART.EOF
                mblncallItemChk = False
                lstCustPartNo.Items.Add(RSCUSTPART.Fields("ITEM_CODE").Value)
                If lstCustPartNo.Items.Item(LNGROWCOUNT).SubItems.Count > 1 Then
                    lstCustPartNo.Items.Item(LNGROWCOUNT).SubItems(1).Text = RSCUSTPART.Fields("DRG_DESC").Value
                Else
                    lstCustPartNo.Items.Item(LNGROWCOUNT).SubItems.Insert(1, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, RSCUSTPART.Fields("DRG_DESC").Value))
                End If
                LNGROWCOUNT = LNGROWCOUNT + 1
                RSCUSTPART.MoveNext()
            End While
        End If
        RSCUSTPART.Close()
        RSCUSTPART = Nothing
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub FILLPROPOSALDETAILS(ByVal intDocNo As Integer, ByVal strConsigneeCode As String, ByVal strItem As String)
        '---------------------------------------------------------------------------------------------------
        ' Author        : Shubhra Verma
        ' Arguments     : intDocNo, strConsigneeCode, strItem
        ' Return Value  : Nil
        ' Function      : fill Proposal Details In lstProposalDetails ListView
        ' Datetime      : 11-Sep-2008
        '-----------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim strsql As String = String.Empty
        Dim RSPROPOSAL As ADODB.Recordset
        Dim intROWCOUNT As Integer
        Dim STRSAFETYSTOCK As String = String.Empty
        lstProposal.Items.Clear()
        intROWCOUNT = 0
        strsql = "SELECT DISTINCT DAYSFORSAFETYSTOCK,SAFETYDAYS,DAILYPULLFLAG,dailypullrate,SUMOFRELEASEQTY" & _
           " From SCHEDULEPROPOSALCALCULATIONS" & _
           " WHERE UNIT_CODE = '" & gstrUNITID & "' AND DOC_NO = " & mstrDocNo & "" & _
           " AND ITEM_CODE = '" & mstrCustPart & "'" & _
           " AND CONSIGNEE_CODE = '" & mstrConsigneeCode & "' and" & _
           " dailypullrate = (select max(dailypullrate) From SCHEDULEPROPOSALCALCULATIONS" & _
           " WHERE UNIT_CODE = '" & gstrUNITID & "' AND DOC_NO = " & mstrDocNo & " AND ITEM_CODE = '" & mstrCustPart & "'  AND CONSIGNEE_CODE = '" & mstrConsigneeCode & "' )"
        RSPROPOSAL = New ADODB.Recordset
        RSPROPOSAL.Open(strsql, mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
        If RSPROPOSAL.RecordCount > 0 Then
            If RSPROPOSAL.Fields("DAILYPULLFLAG").Value = True Then
                STRSAFETYSTOCK = CStr(CInt(RSPROPOSAL.Fields("dailypullrate").Value * RSPROPOSAL.Fields("SAFETYDAYS").Value))
            Else
                STRSAFETYSTOCK = CStr(CInt(RSPROPOSAL.Fields("SUMOFRELEASEQTY").Value / IIf(RSPROPOSAL.Fields("DAYSFORSAFETYSTOCK").Value = 0, 1, RSPROPOSAL.Fields("DAYSFORSAFETYSTOCK").Value)) * RSPROPOSAL.Fields("SAFETYDAYS").Value)
            End If
        End If
        RSPROPOSAL.Close()
        RSPROPOSAL = Nothing
        strsql = "SELECT DISTINCT RELEASE_DT, RELEASE_QTY,WH_CODE,WH_STOCK," & _
       " RECEIVED_QTY , SHIPMENT_DT, SHIPMENT_QTY,WH_DATE" & _
       " From SCHEDULEPROPOSALCALCULATIONS WHERE UNIT_CODE = '" & gstrUNITID & "' AND" & _
       " DOC_NO = '" & intDocNo & "' AND CONSIGNEE_CODE = '" & strConsigneeCode & "'" & _
       " AND ITEM_CODE = '" & strItem & "'"
        RSPROPOSAL = New ADODB.Recordset
        RSPROPOSAL.Open(strsql, mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
        If RSPROPOSAL.RecordCount > 0 Then
            RSPROPOSAL.MoveFirst()
            intROWCOUNT = 0
            While Not RSPROPOSAL.EOF
                mblncallItemChk = False
                lstProposal.Items.Add(Format(RSPROPOSAL.Fields("RELEASE_DT").Value, gstrDateFormat))
                If lstProposal.Items.Item(intROWCOUNT).SubItems.Count > 1 Then
                    lstProposal.Items.Item(intROWCOUNT).SubItems(1).Text = RSPROPOSAL.Fields("RELEASE_QTY").Value
                Else
                    lstProposal.Items.Item(intROWCOUNT).SubItems.Insert(1, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, RSPROPOSAL.Fields("RELEASE_QTY").Value))
                End If
                If MBLNSHIPMENTTHRUWH = True Then
                    If lstProposal.Items.Item(intROWCOUNT).SubItems.Count > 2 Then
                        lstProposal.Items.Item(intROWCOUNT).SubItems(2).Text = RSPROPOSAL.Fields("WH_CODE").Value
                    Else
                        lstProposal.Items.Item(intROWCOUNT).SubItems.Insert(2, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, RSPROPOSAL.Fields("WH_CODE").Value))
                    End If
                    If lstProposal.Items.Item(intROWCOUNT).SubItems.Count > 3 Then
                        lstProposal.Items.Item(intROWCOUNT).SubItems(3).Text = RSPROPOSAL.Fields("WH_STOCK").Value
                    Else
                        lstProposal.Items.Item(intROWCOUNT).SubItems.Insert(3, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, RSPROPOSAL.Fields("WH_STOCK").Value))
                    End If
                    If lstProposal.Items.Item(intROWCOUNT).SubItems.Count > 4 Then
                        lstProposal.Items.Item(intROWCOUNT).SubItems(4).Text = RSPROPOSAL.Fields("RECEIVED_QTY").Value
                    Else
                        lstProposal.Items.Item(intROWCOUNT).SubItems.Insert(4, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, RSPROPOSAL.Fields("RECEIVED_QTY").Value))
                    End If
                    If lstProposal.Items.Item(intROWCOUNT).SubItems.Count > 5 Then
                        lstProposal.Items.Item(intROWCOUNT).SubItems(5).Text = STRSAFETYSTOCK
                    Else
                        lstProposal.Items.Item(intROWCOUNT).SubItems.Insert(5, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, STRSAFETYSTOCK))
                    End If
                Else
                    If lstProposal.Items.Item(intROWCOUNT).SubItems.Count > 2 Then
                        lstProposal.Items.Item(intROWCOUNT).SubItems(2).Text = ""
                    Else
                        lstProposal.Items.Item(intROWCOUNT).SubItems.Insert(2, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, ""))
                    End If
                    If lstProposal.Items.Item(intROWCOUNT).SubItems.Count > 3 Then
                        lstProposal.Items.Item(intROWCOUNT).SubItems(3).Text = ""
                    Else
                        lstProposal.Items.Item(intROWCOUNT).SubItems.Insert(3, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, ""))
                    End If
                    If lstProposal.Items.Item(intROWCOUNT).SubItems.Count > 4 Then
                        lstProposal.Items.Item(intROWCOUNT).SubItems(4).Text = ""
                    Else
                        lstProposal.Items.Item(intROWCOUNT).SubItems.Insert(4, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, ""))
                    End If
                    If lstProposal.Items.Item(intROWCOUNT).SubItems.Count > 5 Then
                        lstProposal.Items.Item(intROWCOUNT).SubItems(5).Text = ""
                    Else
                        lstProposal.Items.Item(intROWCOUNT).SubItems.Insert(5, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, ""))
                    End If
                End If
                If lstProposal.Items.Item(intROWCOUNT).SubItems.Count > 6 Then
                    lstProposal.Items.Item(intROWCOUNT).SubItems(6).Text = Format(RSPROPOSAL.Fields("SHIPMENT_DT").Value, gstrDateFormat)

                Else
                    lstProposal.Items.Item(intROWCOUNT).SubItems.Insert(6, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, Format(RSPROPOSAL.Fields("SHIPMENT_DT").Value, gstrDateFormat)))
                End If
                If lstProposal.Items.Item(intROWCOUNT).SubItems.Count > 7 Then
                    lstProposal.Items.Item(intROWCOUNT).SubItems(7).Text = RSPROPOSAL.Fields("SHIPMENT_QTY").Value
                Else
                    lstProposal.Items.Item(intROWCOUNT).SubItems.Insert(7, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, RSPROPOSAL.Fields("SHIPMENT_QTY").Value))
                End If
                Dim strWH_Date As String
                If IsDBNull(RSPROPOSAL.Fields("WH_DATE").Value) Then
                    strWH_Date = ""
                Else
                    strWH_Date = RSPROPOSAL.Fields("WH_DATE").Value
                End If

                If lstProposal.Items.Item(intROWCOUNT).SubItems.Count > 8 Then
                    lstProposal.Items.Item(intROWCOUNT).SubItems(8).Text = IIf(strWH_Date = "", "", Format(strWH_Date, gstrDateFormat))
                Else
                    lstProposal.Items.Item(intROWCOUNT).SubItems.Insert(8, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, IIf(strWH_Date = "", "", Format(strWH_Date, gstrDateFormat))))
                End If
                intROWCOUNT = intROWCOUNT + 1
                RSPROPOSAL.MoveNext()
            End While
        End If
        RSPROPOSAL.Close()
        RSPROPOSAL = Nothing
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub lstDynamicSafetyStock_ItemClick(ByVal Item As System.Windows.Forms.ListViewItem)
        On Error GoTo ErrHandler
        Call FN_DISPLAYFORMULA(lstDynamicSafetyStock, (Item.Index))
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub lstFixedSafetyStock_ItemClick(ByVal Item As System.Windows.Forms.ListViewItem)
        On Error GoTo ErrHandler
        Call FN_DISPLAYFORMULA(lstFixedSafetyStock, (Item.Index))
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub lstProposal_ItemCheck(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.ItemCheckEventArgs) Handles lstProposal.ItemCheck
        Dim Item As System.Windows.Forms.ListViewItem = lstProposal.Items(eventArgs.Index)
        On Error GoTo ErrHandler
        If mblncallItemChk = True Then
            Call SingleSelection(lstProposal, Item.Index)
        End If
        mblncallItemChk = True
        mstrReleaseDate = getDateForDB(Item.Text)
        mstrWHCode = Item.SubItems.Item(2).Text
        MSTRSHIPMENTDATE = Item.SubItems.Item(6).Text
        lstInTransitDetails.Items.Clear()
        lstFixedSafetyStock.Items.Clear()
        lstDynamicSafetyStock.Items.Clear()
        lstShipment.Items.Clear()
        Call FILLSAFETYSTOCKDETAILS()
        Call FILLINTRANSITDETAILS()
        Call FILLSHIPMENTDETAILS((Item.Index))
        Call FN_DISPLAYFORMULA(lstProposal, Item.Index)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub FILLINTRANSITDETAILS()
        '-----------------------------------------------------------------------------------------------
        ' Author        : Shubhra Verma
        ' Arguments     : Nil
        ' Return Value  : Nil
        ' Function      : fill InTransit Details In lstInTransitDetails ListView
        ' Datetime      : 10-Sep-2008
        '-----------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim strsql As String
        Dim RSINTRANSIT As ADODB.Recordset
        Dim intROWCOUNT As Integer
        intROWCOUNT = 0
        lstInTransitDetails.Items.Clear()
        RSINTRANSIT = New ADODB.Recordset
        If MBLNSHIPMENTTHRUWH = True Then
            strsql = "SELECT SALES_QUANTITY,INVOICE_NO,INVOICE_DATE" & _
            " From INTRANSIT_CDP" & _
            " Where UNIT_CODE = '" & gstrUNITID & "' AND Doc_No = " & mstrDocNo & " " & _
            " AND ITEM_CODE = '" & mstrCustPart & "' " & " AND CONSIGNEE_CODE = '" & mstrConsigneeCode & "' " & _
            " AND WH_CODE = '" & mstrWHCode & "' AND TRANS_DT = '" & mstrReleaseDate & "'"
            RSINTRANSIT.Open(strsql, mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
            If RSINTRANSIT.RecordCount > 0 Then
                RSINTRANSIT.MoveFirst()
                While Not RSINTRANSIT.EOF
                    lstInTransitDetails.Items.Add(Format(RSINTRANSIT.Fields("INVOICE_DATE").Value, gstrDateFormat))
                    If lstInTransitDetails.Items.Item(intROWCOUNT).SubItems.Count > 1 Then
                        lstInTransitDetails.Items.Item(intROWCOUNT).SubItems(1).Text = RSINTRANSIT.Fields("INVOICE_NO").Value
                    Else
                        lstInTransitDetails.Items.Item(intROWCOUNT).SubItems.Insert(1, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, RSINTRANSIT.Fields("INVOICE_NO").Value))
                    End If
                    If lstInTransitDetails.Items.Item(intROWCOUNT).SubItems.Count > 2 Then
                        lstInTransitDetails.Items.Item(intROWCOUNT).SubItems(2).Text = RSINTRANSIT.Fields("SALES_QUANTITY").Value
                    Else
                        lstInTransitDetails.Items.Item(intROWCOUNT).SubItems.Insert(2, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, RSINTRANSIT.Fields("SALES_QUANTITY").Value))
                    End If
                    intROWCOUNT = intROWCOUNT + 1
                    RSINTRANSIT.MoveNext()
                End While
            End If
        End If
        If RSINTRANSIT.State Then RSINTRANSIT.Close()
        RSINTRANSIT = Nothing
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub FILLSAFETYSTOCKDETAILS()
        '---------------------------------------------------------------------------------------------------
        ' Author        : Shubhra Verma
        ' Arguments     : Nil
        ' Return Value  : Nil
        ' Function      : fill Safety Stock Details In lstFixedSafetyStock Or lstDynamicSafetyStock ListView
        ' Datetime      : 11-Sep-2008
        '-----------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim strsql As String
        Dim RSSAFETYSTOCK As ADODB.Recordset
        Dim intROWCOUNT As Integer
        intROWCOUNT = 0
        lstFixedSafetyStock.Items.Clear()
        lstDynamicSafetyStock.Items.Clear()
        If MBLNSHIPMENTTHRUWH = True Then
            strsql = "SELECT DISTINCT DAYSFORSAFETYSTOCK,STOCKCALCWADAYS,SAFETYDAYS,SCHEDULECALCMONTHS," & _
            " DAILYPULLFLAG,WH_DATE,dailypullrate,SUMOFRELEASEQTY From SCHEDULEPROPOSALCALCULATIONS" & _
            " WHERE UNIT_CODE = '" & gstrUNITID & "' AND DOC_NO = " & mstrDocNo & "" & _
            " AND ITEM_CODE = '" & mstrCustPart & "'" & " AND WH_CODE = '" & mstrWHCode & "'" & _
            " AND CONSIGNEE_CODE = '" & mstrConsigneeCode & "' and dailypullrate = (select max(dailypullrate) From SCHEDULEPROPOSALCALCULATIONS WHERE UNIT_CODE = '" & gstrUNITID & "' AND DOC_NO = " & mstrDocNo & " AND ITEM_CODE = '" & mstrCustPart & "' AND WH_CODE = '" & mstrWHCode & "' AND CONSIGNEE_CODE = '" & mstrConsigneeCode & "' )"
            RSSAFETYSTOCK = New ADODB.Recordset
            RSSAFETYSTOCK.Open(strsql, mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
            If RSSAFETYSTOCK.RecordCount > 0 Then
                MBLNDAILYPULLFLAG = RSSAFETYSTOCK.Fields("DAILYPULLFLAG").Value
                If RSSAFETYSTOCK.Fields("DAILYPULLFLAG").Value = False Then
                    If RSSAFETYSTOCK.Fields("STOCKCALCWADAYS").Value = "W" Then
                        lstDynamicSafetyStock.Items.Add("Working Days")
                    Else
                        lstDynamicSafetyStock.Items.Add("Available Days")
                    End If
                    If RSSAFETYSTOCK.Fields("SCHEDULECALCMONTHS").Value = "0" Then
                        If lstDynamicSafetyStock.Items.Item(0).SubItems.Count > 1 Then
                            lstDynamicSafetyStock.Items.Item(0).SubItems(1).Text = "Current Month"
                        Else
                            lstDynamicSafetyStock.Items.Item(0).SubItems.Insert(1, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "Current Month"))
                        End If
                    ElseIf RSSAFETYSTOCK.Fields("SCHEDULECALCMONTHS").Value = "1" Then
                        If lstDynamicSafetyStock.Items.Item(0).SubItems.Count > 1 Then
                            lstDynamicSafetyStock.Items.Item(0).SubItems(1).Text = "Next Month"
                        Else
                            lstDynamicSafetyStock.Items.Item(0).SubItems.Insert(1, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "Next Month"))
                        End If
                    Else
                        If lstDynamicSafetyStock.Items.Item(0).SubItems.Count > 1 Then
                            lstDynamicSafetyStock.Items.Item(0).SubItems(1).Text = "Avg of Next " & RSSAFETYSTOCK.Fields("SCHEDULECALCMONTHS").Value & " Months"
                        Else
                            lstDynamicSafetyStock.Items.Item(0).SubItems.Insert(1, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "Avg of Next " & RSSAFETYSTOCK.Fields("SCHEDULECALCMONTHS").Value & " Months"))
                        End If
                    End If
                    If RSSAFETYSTOCK.Fields("DAYSFORSAFETYSTOCK").Value Then
                        If lstDynamicSafetyStock.Items.Item(0).SubItems.Count > 2 Then
                            lstDynamicSafetyStock.Items.Item(0).SubItems(2).Text = RSSAFETYSTOCK.Fields("DAYSFORSAFETYSTOCK").Value
                        Else
                            lstDynamicSafetyStock.Items.Item(0).SubItems.Insert(2, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, RSSAFETYSTOCK.Fields("DAYSFORSAFETYSTOCK").Value))
                        End If
                    End If
                    If lstDynamicSafetyStock.Items.Item(0).SubItems.Count > 3 Then
                        lstDynamicSafetyStock.Items.Item(0).SubItems(3).Text = RSSAFETYSTOCK.Fields("SUMOFRELEASEQTY").Value
                    Else
                        lstDynamicSafetyStock.Items.Item(0).SubItems.Insert(3, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, RSSAFETYSTOCK.Fields("SUMOFRELEASEQTY").Value))
                    End If
                    If lstDynamicSafetyStock.Items.Item(0).SubItems.Count > 4 Then
                        lstDynamicSafetyStock.Items.Item(0).SubItems(4).Text = CStr(CInt(RSSAFETYSTOCK.Fields("SUMOFRELEASEQTY").Value / RSSAFETYSTOCK.Fields("DAYSFORSAFETYSTOCK").Value))
                    Else
                        lstDynamicSafetyStock.Items.Item(0).SubItems.Insert(4, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, CStr(CInt(RSSAFETYSTOCK.Fields("SUMOFRELEASEQTY").Value / RSSAFETYSTOCK.Fields("DAYSFORSAFETYSTOCK").Value))))
                    End If
                    If lstDynamicSafetyStock.Items.Item(0).SubItems.Count > 5 Then
                        lstDynamicSafetyStock.Items.Item(0).SubItems(5).Text = RSSAFETYSTOCK.Fields("SAFETYDAYS").Value
                    Else
                        lstDynamicSafetyStock.Items.Item(0).SubItems.Insert(5, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, RSSAFETYSTOCK.Fields("SAFETYDAYS").Value))
                    End If
                    If lstDynamicSafetyStock.Items.Item(0).SubItems.Count > 6 Then
                        lstDynamicSafetyStock.Items.Item(0).SubItems(6).Text = CStr(RSSAFETYSTOCK.Fields("SAFETYDAYS").Value * CDbl(lstDynamicSafetyStock.Items.Item(0).SubItems(4).Text))
                    Else
                        lstDynamicSafetyStock.Items.Item(0).SubItems.Insert(6, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, CStr(RSSAFETYSTOCK.Fields("SAFETYDAYS").Value * CDbl(lstDynamicSafetyStock.Items.Item(0).SubItems(4).Text))))
                    End If
                Else
                    While Not RSSAFETYSTOCK.EOF
                        RSSAFETYSTOCK.MoveFirst()
                        lstFixedSafetyStock.Items.Add(RSSAFETYSTOCK.Fields("dailypullrate").Value)
                        If lstFixedSafetyStock.Items.Item(0).SubItems.Count > 1 Then
                            lstFixedSafetyStock.Items.Item(0).SubItems(1).Text = RSSAFETYSTOCK.Fields("SAFETYDAYS").Value
                        Else
                            lstFixedSafetyStock.Items.Item(0).SubItems.Insert(1, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, RSSAFETYSTOCK.Fields("SAFETYDAYS").Value))
                        End If
                        If lstFixedSafetyStock.Items.Item(0).SubItems.Count > 2 Then
                            lstFixedSafetyStock.Items.Item(0).SubItems(2).Text = CStr(Val(RSSAFETYSTOCK.Fields("dailypullrate").Value) * Val(RSSAFETYSTOCK.Fields("SAFETYDAYS").Value))
                        Else
                            lstFixedSafetyStock.Items.Item(0).SubItems.Insert(2, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, CStr(Val(RSSAFETYSTOCK.Fields("dailypullrate").Value) * Val(RSSAFETYSTOCK.Fields("SAFETYDAYS").Value))))
                        End If
                        RSSAFETYSTOCK.MoveNext()
                    End While
                End If
            End If
            RSSAFETYSTOCK.Close()
            RSSAFETYSTOCK = Nothing
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub FILLSHIPMENTDETAILS(ByRef intROW As Integer)
        '---------------------------------------------------------------------------------------------------
        ' Author        : Shubhra Verma
        ' Arguments     : intRow
        ' Return Value  : Nil
        ' Function      : fill Shipment Details In lstShipment ListView
        ' Datetime      : 13-Sep-2008
        '-----------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim strsql As String
        Dim RSSHIPMENT As ADODB.Recordset
        Dim intROWCOUNT As Integer
        Dim intINTRANSIT As Integer
        Dim intRELEASEQTY As Integer
        Dim intSAFETYSTOCK As Integer
        Dim intWHSTOCK As Integer
        intROWCOUNT = 0
        lstShipment.Items.Clear()
        intINTRANSIT = CInt(Val(lstProposal.Items.Item(intROW).SubItems(4).Text))
        intRELEASEQTY = CInt(Val(lstProposal.Items.Item(intROW).SubItems(1).Text))
        intSAFETYSTOCK = CInt(Val(lstProposal.Items.Item(intROW).SubItems(5).Text))
        intWHSTOCK = CInt(Val(lstProposal.Items.Item(intROW).SubItems(3).Text))
        strsql = "SELECT MAX(BAG_QTY) AS BAG_QTY,MAX(ENT_DT) AS UPLOAD_DATE,SAFETYDAYSMIN,SAFETYDAYSMAX" & _
       " ,transitDays, bufferDays" & _
       " FROM SCHEDULEPROPOSALCALCULATIONS" & _
       " WHERE UNIT_CODE = '" & gstrUNITID & "' AND DOC_NO = " & mstrDocNo & "" & _
       " AND ITEM_CODE = '" & mstrCustPart & "'" & " AND CONSIGNEE_CODE = '" & mstrConsigneeCode & "'" & _
       " GROUP BY SAFETYDAYSMIN,SAFETYDAYSMAX,transitDays, bufferDays,WH_DATE"
        RSSHIPMENT = New ADODB.Recordset
        RSSHIPMENT.Open(strsql, mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockOptimistic)
        If RSSHIPMENT.RecordCount > 0 Then
            RSSHIPMENT.MoveFirst()
            While Not RSSHIPMENT.EOF
                lstShipment.Items.Add(IIf(IsDBNull(RSSHIPMENT.Fields("BAG_QTY").Value), "0", RSSHIPMENT.Fields("BAG_QTY").Value))
                If MBLNSHIPMENTTHRUWH = True Then
                    If lstShipment.Items.Item(intROWCOUNT).SubItems.Count > 1 Then
                        lstShipment.Items.Item(intROWCOUNT).SubItems(1).Text = CStr(Val(CStr(intWHSTOCK)) + Val(CStr(intINTRANSIT)) - Val(CStr(intSAFETYSTOCK)) - Val(CStr(intRELEASEQTY)))
                    Else
                        lstShipment.Items.Item(intROWCOUNT).SubItems.Insert(1, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, CStr(Val(CStr(intWHSTOCK)) + Val(CStr(intINTRANSIT)) - Val(CStr(intSAFETYSTOCK)) - Val(CStr(intRELEASEQTY)))))
                    End If
                    If lstShipment.Items.Item(intROWCOUNT).SubItems.Count > 2 Then
                        lstShipment.Items.Item(intROWCOUNT).SubItems(2).Text = CStr(IIf(IsDBNull(RSSHIPMENT.Fields("SAFETYDAYSMIN").Value), "0", RSSHIPMENT.Fields("SAFETYDAYSMIN").Value))
                    Else
                        lstShipment.Items.Item(intROWCOUNT).SubItems.Insert(2, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, CStr(IIf(IsDBNull(RSSHIPMENT.Fields("SAFETYDAYSMIN").Value), "0", RSSHIPMENT.Fields("SAFETYDAYSMIN").Value))))
                    End If
                    If lstShipment.Items.Item(intROWCOUNT).SubItems.Count > 3 Then
                        lstShipment.Items.Item(intROWCOUNT).SubItems(3).Text = CStr(IIf(IsDBNull(RSSHIPMENT.Fields("SAFETYDAYSMAX").Value), "0", RSSHIPMENT.Fields("SAFETYDAYSMAX").Value))
                    Else
                        lstShipment.Items.Item(intROWCOUNT).SubItems.Insert(3, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, CStr(CStr(IIf(IsDBNull(RSSHIPMENT.Fields("SAFETYDAYSMAX").Value), "0", RSSHIPMENT.Fields("SAFETYDAYSMAX").Value)))))
                    End If
                Else
                    If lstShipment.Items.Item(intROWCOUNT).SubItems.Count > 1 Then
                        lstShipment.Items.Item(intROWCOUNT).SubItems(1).Text = ""
                    Else
                        lstShipment.Items.Item(intROWCOUNT).SubItems.Insert(1, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, ""))
                    End If
                    If lstShipment.Items.Item(intROWCOUNT).SubItems.Count > 2 Then
                        lstShipment.Items.Item(intROWCOUNT).SubItems(2).Text = ""
                    Else
                        lstShipment.Items.Item(intROWCOUNT).SubItems.Insert(2, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, ""))
                    End If
                    If lstShipment.Items.Item(intROWCOUNT).SubItems.Count > 3 Then
                        lstShipment.Items.Item(intROWCOUNT).SubItems(3).Text = ""
                    Else
                        lstShipment.Items.Item(intROWCOUNT).SubItems.Insert(3, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, ""))
                    End If
                End If
                If Val(CStr(intWHSTOCK)) + Val(CStr(intINTRANSIT)) - Val(CStr(intSAFETYSTOCK)) - Val(CStr(intRELEASEQTY)) < 0 Then
                    If MBLNSHIPMENTTHRUWH = True Then
                        If lstShipment.Items.Item(intROWCOUNT).SubItems.Count > 4 Then
                            lstShipment.Items.Item(intROWCOUNT).SubItems(4).Text = CStr(0 - (Val(CStr(intWHSTOCK)) + Val(CStr(intINTRANSIT)) - Val(CStr(intSAFETYSTOCK)) - Val(CStr(intRELEASEQTY))))
                        Else
                            lstShipment.Items.Item(intROWCOUNT).SubItems.Insert(4, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, CStr(0 - (Val(CStr(intWHSTOCK)) + Val(CStr(intINTRANSIT)) - Val(CStr(intSAFETYSTOCK)) - Val(CStr(intRELEASEQTY))))))
                        End If
                    ElseIf Val(lstProposal.Items.Item(intROW).SubItems(7).Text) > 0 Then
                        If lstShipment.Items.Item(intROWCOUNT).SubItems.Count > 4 Then
                            lstShipment.Items.Item(intROWCOUNT).SubItems(4).Text = CStr(0 - (Val(CStr(intWHSTOCK)) + Val(CStr(intINTRANSIT)) - Val(CStr(intSAFETYSTOCK)) - Val(CStr(intRELEASEQTY))))
                        Else
                            lstShipment.Items.Item(intROWCOUNT).SubItems.Insert(4, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, CStr(0 - (Val(CStr(intWHSTOCK)) + Val(CStr(intINTRANSIT)) - Val(CStr(intSAFETYSTOCK)) - Val(CStr(intRELEASEQTY))))))
                        End If
                    Else
                        If lstShipment.Items.Item(intROWCOUNT).SubItems.Count > 4 Then
                            lstShipment.Items.Item(intROWCOUNT).SubItems(4).Text = ""
                        Else
                            lstShipment.Items.Item(intROWCOUNT).SubItems.Insert(4, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, ""))
                        End If
                    End If
                    If lstShipment.Items.Item(intROWCOUNT).SubItems.Count > 5 Then
                        lstShipment.Items.Item(intROWCOUNT).SubItems(5).Text = lstProposal.Items.Item(intROW).SubItems(7).Text
                    Else
                        lstShipment.Items.Item(intROWCOUNT).SubItems.Insert(5, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, lstProposal.Items.Item(intROW).SubItems(7).Text))
                    End If
                    If IIf(IsDBNull(RSSHIPMENT.Fields("BAG_QTY").Value), 0, RSSHIPMENT.Fields("BAG_QTY").Value) = 0 Then
                        If lstShipment.Items.Item(intROWCOUNT).SubItems.Count > 6 Then
                            lstShipment.Items.Item(intROWCOUNT).SubItems(6).Text = CStr(Val(lstProposal.Items.Item(intROW).SubItems(7).Text))
                        Else
                            lstShipment.Items.Item(intROWCOUNT).SubItems.Insert(6, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, CStr(Val(lstProposal.Items.Item(intROW).SubItems(7).Text))))
                        End If
                    Else
                        If lstShipment.Items.Item(intROWCOUNT).SubItems.Count > 6 Then
                            lstShipment.Items.Item(intROWCOUNT).SubItems(6).Text = CStr(Val(lstProposal.Items.Item(intROW).SubItems(7).Text) / Val(RSSHIPMENT.Fields("BAG_QTY").Value))
                        Else
                            lstShipment.Items.Item(intROWCOUNT).SubItems.Insert(6, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, CStr(Val(lstProposal.Items.Item(intROW).SubItems(7).Text) / Val(RSSHIPMENT.Fields("BAG_QTY").Value))))
                        End If
                    End If
                Else
                    If lstShipment.Items.Item(intROWCOUNT).SubItems.Count > 4 Then
                        lstShipment.Items.Item(intROWCOUNT).SubItems(4).Text = CStr(0)
                    Else
                        lstShipment.Items.Item(intROWCOUNT).SubItems.Insert(4, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, CStr(0)))
                    End If
                    If lstShipment.Items.Item(intROWCOUNT).SubItems.Count > 5 Then
                        lstShipment.Items.Item(intROWCOUNT).SubItems(5).Text = CStr(0)
                    Else
                        lstShipment.Items.Item(intROWCOUNT).SubItems.Insert(5, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, CStr(0)))
                    End If
                    If lstShipment.Items.Item(intROWCOUNT).SubItems.Count > 6 Then
                        lstShipment.Items.Item(intROWCOUNT).SubItems(6).Text = CStr(0)
                    Else
                        lstShipment.Items.Item(intROWCOUNT).SubItems.Insert(6, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, CStr(0)))
                    End If
                End If
                If lstShipment.Items.Item(intROWCOUNT).SubItems.Count > 7 Then
                    lstShipment.Items.Item(intROWCOUNT).SubItems(7).Text = IIf(IsDBNull(RSSHIPMENT.Fields("transitDays").Value), 0, RSSHIPMENT.Fields("transitDays").Value)
                Else
                    lstShipment.Items.Item(intROWCOUNT).SubItems.Insert(7, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, IIf(IsDBNull(RSSHIPMENT.Fields("transitDays").Value), 0, RSSHIPMENT.Fields("transitDays").Value)))
                End If
                If lstShipment.Items.Item(intROWCOUNT).SubItems.Count > 8 Then
                    lstShipment.Items.Item(intROWCOUNT).SubItems(8).Text = IIf(IsDBNull(RSSHIPMENT.Fields("bufferDays").Value), 0, RSSHIPMENT.Fields("bufferDays").Value)
                Else
                    lstShipment.Items.Item(intROWCOUNT).SubItems.Insert(8, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, IIf(IsDBNull(RSSHIPMENT.Fields("bufferDays").Value), 0, RSSHIPMENT.Fields("bufferDays").Value)))
                End If
                If lstShipment.Items.Item(intROWCOUNT).SubItems.Count > 9 Then
                    lstShipment.Items.Item(intROWCOUNT).SubItems(9).Text = Format(RSSHIPMENT.Fields("UPLOAD_DATE").Value, gstrDateFormat)
                Else
                    lstShipment.Items.Item(intROWCOUNT).SubItems.Insert(9, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, Format(RSSHIPMENT.Fields("UPLOAD_DATE").Value, gstrDateFormat)))
                End If
                intROWCOUNT = intROWCOUNT + 1
                RSSHIPMENT.MoveNext()
            End While
        End If
        RSSHIPMENT.Close()
        RSSHIPMENT = Nothing
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Function FN_DISPLAYFORMULA(ByVal list As System.Windows.Forms.ListView, Optional ByVal intROW As Integer = 1) As String
        '---------------------------------------------------------------------------------------------------
        ' Author        : Shubhra Verma
        ' Arguments     : list,  intROW
        ' Return Value  : Nil
        ' Function      : fill Safety Stock Details In lstFixedSafetyStock Or lstDynamicSafetyStock ListView
        ' Datetime      : 13-Sep-2008
        '-----------------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        lblFormula1.Text = "" : lblFormula2.Text = "" : lblFormula3.Text = "" : lblFormula4.Text = "" ': Text5.caption = ""
        If list.Name = lstProposal.Name Then
            If MBLNSHIPMENTTHRUWH = True Then
                If intROW = 0 Then
                    lblFormula1.Text = "WHSTOCK = WAREHOUSE STOCK FROM WAREHOUSE STATUS FILE"
                ElseIf intROW > 0 Then
                    lblFormula1.Text = "WHSTOCK = PREVIOUS WHSTOCK + PREVIOUS INTRANSIT - RELEASE QTY"
                End If
                If MBLNDAILYPULLFLAG = True Then
                    lblFormula2.Text = "SAFETY STOCK = DAILYPULLRATE * SAFETY DAYS"
                Else
                    lblFormula2.Text = "SAFETY STOCK = (SUM OF RELEASE QTY/WORKING OR AVAILABLE DAYS) * SAFETY DAYS"
                End If
            Else
                lblFormula1.Text = ""
                lblFormula2.Text = ""
            End If
            lblFormula3.Text = "DESPATCH DATE = RELEASE DATE - TRANSIT DAYS - BUFFER DAYS"
            If MBLNSHIPMENTTHRUWH = True Then
                lblFormula4.Text = "DESPATCH QTY = (WHSTOCK + INTRANSIT QTY - SAFETY STOCK - RELEASE QTY) IN MULTIPLES OF BAG QTY"
            Else
                lblFormula4.Text = "DESPATCH QTY = RELEASE QTY IN MULTIPLES OF BAG QTY"
            End If
        End If
        If list.Name = lstFixedSafetyStock.Name Then
            lblFormula1.Text = "DAILYPULLRATE = DAILYPULLRATE FROM WAREHOUSE STATUS FILE"
            lblFormula2.Text = "SAFETY DAYS = SAFETY DAYS FROM RELEASE FILE PARAMETER MASTER"
            lblFormula3.Text = "SAFETY STOCK = DAILYPULLRATE * SAFETY DAYS"
        End If
        If list.Name = lstDynamicSafetyStock.Name Then
            lblFormula1.Text = "SAFETY STOCK PER DAY = SUM OF RELEASE QTY / WORKING OR AVAILABLE DAYS"
            lblFormula2.Text = "SAFETY STOCK = SAFETY STOCK PER DAY * SAFETY DAYS"
        End If
        If list.Name = lstShipment.Name Then
            If MBLNSHIPMENTTHRUWH = True Then
                lblFormula1.Text = "REMAINING WHSTOCK = WHSTOCK + INTRANSIT QTY - SAFETY STOCK - RELEASE QTY"
                lblFormula2.Text = "DESPATCH QTY = IF REMAINING WHSTOCK < 0 THEN DESPATCH QTY  = 0 - REMAINING WHSTOCK"
                lblFormula3.Text = "ACTUAL DESPATCH QTY = DESPATCH QTY IN MULTIPLES OF BAG QTY"
                lblFormula4.Text = "NO. OF BAGS = ACTUAL DESPATCH QTY / BAG QTY"
            Else
                lblFormula2.Text = "DESPATCH QTY = RELEASE QTY"
                lblFormula3.Text = "ACTUAL DESPATCH QTY = DESPATCH QTY IN MULTIPLES OF BAG QTY"
                lblFormula4.Text = "NO. OF BAGS = ACTUAL DESPATCH QTY / BAG QTY"
            End If
        End If
        Return ""
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub search()
        On Error GoTo ErrHandler
        '-----------------------------------------------------------------------------
        'Created By     :   Shubhra Verma
        'Arguments      :   Nil
        'Return Value   :   Nil
        'Function       :   for searching
        '-----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim intCounter As Short
        Dim lvwListView As System.Windows.Forms.ListView
        If cmbSearch.Text = "Doc No" Then
            For intCounter = 0 To lstDocNo.Items.Count - 1
                If lstDocNo.Items.Item(intCounter).Font.Bold = True Then
                    lstDocNo.Items.Item(intCounter).Font = VB6.FontChangeBold(lstDocNo.Items.Item(intCounter).Font, False)
                    lstDocNo.Refresh()
                End If
            Next
            If Len(txtSearchBox.Text) = 0 Then Exit Sub
            For intCounter = 0 To lstDocNo.Items.Count - 1
                If Trim(UCase(Mid(lstDocNo.Items.Item(intCounter).Text, 1, Len(txtSearchBox.Text)))) = Trim(UCase(txtSearchBox.Text)) Then
                    lstDocNo.Items.Item(intCounter).Font = VB6.FontChangeBold(lstDocNo.Items.Item(intCounter).Font, True)
                    Call lstDocNo.Items.Item(intCounter).EnsureVisible()
                    lstDocNo.Refresh()
                    Exit For
                End If
            Next
        End If
        If cmbSearch.Text = "Consignee Code" Then
            For intCounter = 0 To lstConsigneeCode.Items.Count - 1
                If lstConsigneeCode.Items.Item(intCounter).Font.Bold = True Then
                    lstConsigneeCode.Items.Item(intCounter).Font = VB6.FontChangeBold(lstConsigneeCode.Items.Item(intCounter).Font, False)
                    lstConsigneeCode.Refresh()
                End If
            Next
            If Len(txtSearchBox.Text) = 0 Then Exit Sub
            For intCounter = 0 To lstConsigneeCode.Items.Count - 1
                If Trim(UCase(Mid(lstConsigneeCode.Items.Item(intCounter).Text, 1, Len(txtSearchBox.Text)))) = Trim(UCase(txtSearchBox.Text)) Then
                    lstConsigneeCode.Items.Item(intCounter).Font = VB6.FontChangeBold(lstConsigneeCode.Items.Item(intCounter).Font, True)
                    Call lstConsigneeCode.Items.Item(intCounter).EnsureVisible()
                    lstConsigneeCode.Refresh()
                    Exit For
                End If
            Next
        End If
        If cmbSearch.Text = "Cust Part No" Then
            For intCounter = 0 To lstCustPartNo.Items.Count - 1
                If lstCustPartNo.Items.Item(intCounter).Font.Bold = True Then
                    lstCustPartNo.Items.Item(intCounter).Font = VB6.FontChangeBold(lstCustPartNo.Items.Item(intCounter).Font, False)
                    lstCustPartNo.Refresh()
                End If
            Next
            If Len(txtSearchBox.Text) = 0 Then Exit Sub
            For intCounter = 0 To lstCustPartNo.Items.Count - 1
                If Trim(UCase(Mid(lstCustPartNo.Items.Item(intCounter).Text, 1, Len(txtSearchBox.Text)))) = Trim(UCase(txtSearchBox.Text)) Then
                    lstCustPartNo.Items.Item(intCounter).Font = VB6.FontChangeBold(lstCustPartNo.Items.Item(intCounter).Font, True)
                    Call lstCustPartNo.Items.Item(intCounter).EnsureVisible()
                    lstCustPartNo.Refresh()
                    Exit For
                End If
            Next
        End If
        If cmbSearch.Text = "Cust Part Desc" Then
            For intCounter = 0 To lstCustPartNo.Items.Count - 1
                If lstCustPartNo.Items.Item(intCounter).Font.Bold = True Then
                    lstCustPartNo.Items.Item(intCounter).Font = VB6.FontChangeBold(lstCustPartNo.Items.Item(intCounter).Font, False)
                    lstCustPartNo.Refresh()
                End If
            Next
            If Len(txtSearchBox.Text) = 0 Then Exit Sub
            For intCounter = 0 To lstCustPartNo.Items.Count - 1
                If Trim(UCase(Mid(lstCustPartNo.Items.Item(intCounter).SubItems(1).Text, 1, Len(txtSearchBox.Text)))) = Trim(UCase(txtSearchBox.Text)) Then
                    lstCustPartNo.Items.Item(intCounter).Font = VB6.FontChangeBold(lstCustPartNo.Items.Item(intCounter).Font, True)
                    Call lstCustPartNo.Items.Item(intCounter).EnsureVisible()
                    lstCustPartNo.Refresh()
                    Exit For
                End If
            Next
        End If
        If cmbSearch.Text = "Customer Code" Then
            For intCounter = 0 To lstDocNo.Items.Count - 1
                If lstDocNo.Items.Item(intCounter).Font.Bold = True Then
                    lstDocNo.Items.Item(intCounter).Font = VB6.FontChangeBold(lstDocNo.Items.Item(intCounter).Font, False)
                    lstDocNo.Refresh()
                End If
            Next
            If Len(txtSearchBox.Text) = 0 Then Exit Sub
            For intCounter = 0 To lstDocNo.Items.Count - 1
                If Trim(UCase(Mid(lstDocNo.Items.Item(intCounter).SubItems(1).Text, 1, Len(txtSearchBox.Text)))) = Trim(UCase(txtSearchBox.Text)) Then
                    lstDocNo.Items.Item(intCounter).Font = VB6.FontChangeBold(lstDocNo.Items.Item(intCounter).Font, True)
                    Call lstDocNo.Items.Item(intCounter).EnsureVisible()
                    lstDocNo.Refresh()
                    Exit For
                End If
            Next
        End If
        If cmbSearch.Text = "Customer Name" Then
            For intCounter = 0 To lstDocNo.Items.Count - 1
                If lstDocNo.Items.Item(intCounter).Font.Bold = True Then
                    lstDocNo.Items.Item(intCounter).Font = VB6.FontChangeBold(lstDocNo.Items.Item(intCounter).Font, False)
                    lstDocNo.Refresh()
                End If
            Next
            If Len(txtSearchBox.Text) = 0 Then Exit Sub
            For intCounter = 0 To lstDocNo.Items.Count - 1
                If Trim(UCase(Mid(lstDocNo.Items.Item(intCounter).SubItems(2).Text, 1, Len(txtSearchBox.Text)))) = Trim(UCase(txtSearchBox.Text)) Then
                    lstDocNo.Items.Item(intCounter).Font = VB6.FontChangeBold(lstDocNo.Items.Item(intCounter).Font, True)
                    Call lstDocNo.Items.Item(intCounter).EnsureVisible()
                    lstDocNo.Refresh()
                    Exit For
                End If
            Next
        End If
        If cmbSearch.Text = "Consignee Name" Then
            For intCounter = 0 To lstConsigneeCode.Items.Count - 1
                If lstConsigneeCode.Items.Item(intCounter).Font.Bold = True Then
                    lstConsigneeCode.Items.Item(intCounter).Font = VB6.FontChangeBold(lstConsigneeCode.Items.Item(intCounter).Font, False)
                    lstConsigneeCode.Refresh()
                End If
            Next
            If Len(txtSearchBox.Text) = 0 Then Exit Sub
            For intCounter = 0 To lstConsigneeCode.Items.Count - 1
                If Trim(UCase(Mid(lstConsigneeCode.Items.Item(intCounter).SubItems(1).Text, 1, Len(txtSearchBox.Text)))) = Trim(UCase(txtSearchBox.Text)) Then
                    lstConsigneeCode.Items.Item(intCounter).Font = VB6.FontChangeBold(lstConsigneeCode.Items.Item(intCounter).Font, True)
                    Call lstConsigneeCode.Items.Item(intCounter).EnsureVisible()
                    lstConsigneeCode.Refresh()
                    Exit For
                End If
            Next
        End If
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub lstProposal_ItemClick(ByVal Item As System.Windows.Forms.ListViewItem)
        On Error GoTo ErrHandler
        Call FN_DISPLAYFORMULA(lstProposal, (Item.Index))
        Call lstProposal_ItemCheck(lstProposal, New System.Windows.Forms.ItemCheckEventArgs(Item.Index, System.Windows.Forms.CheckState.Indeterminate, Item.Checked))
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub lstShipment_ItemClick(ByVal Item As System.Windows.Forms.ListViewItem)
        On Error GoTo ErrHandler
        Call FN_DISPLAYFORMULA(lstShipment, (Item.Index))
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtSearchBox_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSearchBox.TextChanged
        On Error GoTo ErrHandler
        Call search()
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub lstDocNo_ItemChecked(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckedEventArgs) Handles lstDocNo.ItemChecked
        On Error GoTo ErrHandler
        Dim Item As System.Windows.Forms.ListViewItem = e.Item
        Call SingleSelection(lstDocNo, Item.Index)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub lstCustPartNo_ItemChecked(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckedEventArgs) Handles lstCustPartNo.ItemChecked
        On Error GoTo ErrHandler
        Dim Item As System.Windows.Forms.ListViewItem = e.Item
        If mblncallItemChk = True Then
            Call SingleSelection(lstCustPartNo, Item.Index)
        End If
        mblncallItemChk = True
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub lstProposal_ItemChecked(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckedEventArgs) Handles lstProposal.ItemChecked
        On Error GoTo ErrHandler
        Dim Item As System.Windows.Forms.ListViewItem = e.Item
        If mblncallItemChk = True Then
            Call SingleSelection(lstProposal, Item.Index)
        End If
        mblncallItemChk = True
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub lstConsigneeCode_ItemChecked(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckedEventArgs) Handles lstConsigneeCode.ItemChecked
        On Error GoTo ErrHandler
        Dim Item As System.Windows.Forms.ListViewItem = e.Item
        If mblncallItemChk = True Then
            Call SingleSelection(lstConsigneeCode, Item.Index)
        End If
        mblncallItemChk = True
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub lstDynamicSafetyStock_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstDynamicSafetyStock.Click
        On Error GoTo ErrHandler
        FN_DISPLAYFORMULA(lstDynamicSafetyStock)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub lstFixedSafetyStock_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstFixedSafetyStock.Click
        On Error GoTo ErrHandler
        FN_DISPLAYFORMULA(lstFixedSafetyStock)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub lstShipment_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstShipment.Click
        On Error GoTo ErrHandler
        FN_DISPLAYFORMULA(lstShipment)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMFGTRN0062_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Escape
                Me.Close()
        End Select
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        GoTo EventExitSub
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub frmMKTTRN0062_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'On Error GoTo ErrHandler
        'Me.Dispose()
        'Exit Sub
        'ErrHandler:  'The Error Handling Code Starts here
        '        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        '        Exit Sub
    End Sub
    Private Sub frmMFGTRN0062_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        On Error GoTo ErrHandler
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.Escape And Shift = 0 Then Me.Close()
        If KeyCode = System.Windows.Forms.Keys.F1 Or KeyCode = System.Windows.Forms.Keys.F1 Then Call ShowHelp("HLP" & Mid(Me.Name, 4, Len(Me.Name)) & ".htm")
        If KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0 Then
            Me.Close()
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub ctlFormHeader_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlFormHeader.Click
        On Error GoTo ErrHandler
        Call ShowHelp("HLP" & Mid(Me.Name, 4, Len(Me.Name)) & ".htm")
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub

    Private Sub BtnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnClose.Click
        Try
            Me.Close()
            Exit Sub
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
End Class