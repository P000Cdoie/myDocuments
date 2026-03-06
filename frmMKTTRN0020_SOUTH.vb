Option Strict Off
Option Explicit On
Friend Class frmMKTTRN0020_SOUTH
	Inherits System.Windows.Forms.Form
	'===================================================================================
	' (c) MothersonSumi INfotech & Design Ltd. All rights reserverd.
	' File Name         :   FRMMKTTRN0020.frm
	' Function          :   Used to select reference no
	' Created By        :   Nisha
	' Created On        :   15 May, 2001
	' Revision History  :   Nisha Rai
	'21/09/2001 MARKED CHECKED BY BCs changed on version 1
	'03/10/2001 MARKED CHECKED BY BCs changed on version 2
	'03/10/2001 no changed on version 3
	'05/06/02 TO Add Authorized Flag =1
	'CHANGED ON 15/07/2002 FOR EXPORT OPTION ADDING AND CALCULATION SAME AS NORMAL INVOICE
	'23/07/2002 changed to add Grin Linking in Rejection Invoice
	'CHANGES DONE BY NISHA ON 13/03/2003
	'1.FOR FINAL MERGING & FOR FROM BOX & TO bOX UPDATION WHILE EDITING INVOICE
	'2.For Grin Cancellation flag
	'3.SAMPLE INVOICE TOOL COST COLUMN
	'4.CUNSUMABLES & MISC. SALE IN CASE OF NORMAL RAW MATERIAL INVOICE
    '===================================================================================
    'Revised By         : Manoj Kr. Vaish
    'Revised on         : 23-Sep-2008 Issue ID:eMpro-20080923-21892
    'Revised For        : Changes has been reverted for Export Invoice entry through sales order
    '===================================================================================
    'Revised By        -    Vinod Singh
    'Revision Date     -    19/05/2011
    'Revision History  -    Changes for Multi Unit
    '===================================================================================
    'REVISED BY     -  PRASHANT RAJPAL
    'REVISED ON     -  18 SEP 2015
    'PURPOSE        -  10869290 -SERVICE INVOICE 
    '-------------------------------------------------------------------------------------------------------------------------------------------------------

    Dim mCtlHdrItemCode As System.Windows.Forms.ColumnHeader
    Dim mCtlHdrDrawingNo As System.Windows.Forms.ColumnHeader
    Dim mCtlHdrCustRef As System.Windows.Forms.ColumnHeader
    Dim mCtlHdrAmmendment As System.Windows.Forms.ColumnHeader
    Dim mCtlHdrdrgdesc As System.Windows.Forms.ColumnHeader
    Dim mListItemUserId As System.Windows.Forms.ListViewItem
    Dim mstrItemText As String
    Private Sub CmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdCancel.Click
        On Error GoTo ErrHandler
        mstrItemText = "" 'User CANCELS Form
        Me.Close()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdOk.Click
        On Error GoTo ErrHandler
        If Len(mstrItemText) = 0 Then
            Call ConfirmWindow(10418, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
            Exit Sub
        End If
        Me.Close()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub frmMKTTRN0020_SOUTH_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo ErrHandler
        If lvwCustRefNo.Enabled Then
            lvwCustRefNo.Focus()
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub frmMKTTRN0020_SOUTH_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo ErrHandler
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub AddColumnsInListView()
        '***********************************
        'To add Columns Headers in the ListView in the form load
        '***********************************
        On Error GoTo ErrHandler
        With Me.lvwCustRefNo
            mCtlHdrCustRef = .Columns.Add("")
            mCtlHdrCustRef.Text = "Refrence No"
            mCtlHdrCustRef.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(lvwCustRefNo.Width) / 5)
            mCtlHdrAmmendment = .Columns.Add("")
            mCtlHdrAmmendment.Text = "Amend No."
            mCtlHdrAmmendment.Width = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(lvwCustRefNo.Width) / 5) - 50)
            mCtlHdrDrawingNo = .Columns.Add("")
            mCtlHdrDrawingNo.Text = "Drawing No."
            mCtlHdrDrawingNo.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(lvwCustRefNo.Width) / 5)
            mCtlHdrdrgdesc = .Columns.Add("")
            mCtlHdrdrgdesc.Text = "Customer Part Desc"
            mCtlHdrdrgdesc.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(lvwCustRefNo.Width) / 5)
            mCtlHdrItemCode = .Columns.Add("")
            mCtlHdrItemCode.Text = "Item Code"
            mCtlHdrItemCode.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(lvwCustRefNo.Width) / 5)
        End With
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Public Function SelectDataFromCustOrd_Dtl(ByRef pstrCustCode As String, ByRef pstrInvType As String, Optional ByRef pstrConsCode As String = "") As String
        '***********************************
        'To Get Data From Cust_Ord_Dtl
        '***********************************
        On Error GoTo ErrHandler
        Dim strSelectSql As String 'Declared To Make Select Query
        Dim rsCustOrdDtl As ClsResultSetDB
        Dim intRecordCount As Short 'To Hold Record Count
        Dim intCount As Short
        Call AddColumnsInListView()

        'Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(mdifrmMain.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2)
        'Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(mdifrmMain.Width) - VB6.PixelsToTwipsX(frmModules.Width)) / 2.3)
        optRefrence.Checked = True
        Me.lvwCustRefNo.Items.Clear() 'initially clear all items in the listview
        If UCase(pstrInvType) = "JOBWORK INVOICE" Then
            strSelectSql = "Select b.Item_Code,b.Cust_DrgNo,b.Amendment_No,b.Cust_Ref,b.cust_drg_desc from Cust_Ord_hdr a,Cust_Ord_Dtl b"
            strSelectSql = strSelectSql & " where A.UNIT_CODE = B.UNIT_CODE AND b.Account_Code='" & Trim(pstrCustCode) & "' and b.Active_flag ='A' and "
            strSelectSql = strSelectSql & " a.Account_Code = b.Account_Code and a.Cust_ref =b.Cust_ref and "
            strSelectSql = strSelectSql & " a.Amendment_No = b.amendment_No AND a.Authorized_Flag = 1 and a.PO_type in ('J') "
            strSelectSql = strSelectSql & " and a.Valid_date >'" & getDateForDB(GetServerDate()) & "' and effect_Date <='" & getDateForDB(GetServerDate()) & "' and a.unit_code='" & gstrUNITID & "'"
            strSelectSql = strSelectSql & " order by b.Cust_Ref,b.Amendment_No,b.Cust_DrgNo "
        ElseIf UCase(pstrInvType) = "EXPORT INVOICE" Then
            strSelectSql = "Select b.Item_Code,b.Cust_DrgNo,b.Amendment_No,b.Cust_Ref,b.cust_drg_desc from Cust_Ord_hdr a,Cust_Ord_Dtl b"
            strSelectSql = strSelectSql & " where A.UNIT_CODE = B.UNIT_CODE AND b.Account_Code='" & Trim(pstrCustCode) & "' and b.Active_flag ='A' and "
            strSelectSql = strSelectSql & " a.Account_Code = b.Account_Code and a.Cust_ref = b.Cust_ref and "
            strSelectSql = strSelectSql & " a.Amendment_No = b.amendment_No AND a.Authorized_Flag = 1 and a.PO_type in ('E') "
            strSelectSql = strSelectSql & " and a.Valid_date >'" & getDateForDB(GetServerDate()) & "' and effect_date <='" & getDateForDB(GetServerDate()) & "' and A.UNIT_CODE ='" & gstrUNITID & "'"
            If Len(Trim(pstrConsCode)) > 0 Then
                strSelectSql = strSelectSql & " and a.consignee_code='" & Trim(pstrConsCode) & "' "
            End If
            strSelectSql = strSelectSql & " order by b.Cust_Ref,b.Amendment_No,b.Cust_DrgNo,b.Item_Code "
        ElseIf UCase(pstrInvType) = "SERVICE INVOICE" Then
            strSelectSql = "Select b.Item_Code,b.Cust_DrgNo,b.Amendment_No,b.Cust_Ref,b.cust_drg_desc from Cust_Ord_hdr a,Cust_Ord_Dtl b"
            strSelectSql = strSelectSql & " where A.UNIT_CODE = B.UNIT_CODE AND b.Account_Code='" & Trim(pstrCustCode) & "' and b.Active_flag ='A' and "
            strSelectSql = strSelectSql & " a.Account_Code = b.Account_Code and a.Cust_ref =b.Cust_ref and "
            strSelectSql = strSelectSql & " a.Amendment_No = b.amendment_No AND a.Authorized_Flag = 1 and a.PO_type in ('V') "
            strSelectSql = strSelectSql & " and a.Valid_date >'" & getDateForDB(GetServerDate()) & "' and effect_Date <='" & getDateForDB(GetServerDate()) & "' and a.unit_code='" & gstrUNITID & "'"
            strSelectSql = strSelectSql & " order by b.Cust_Ref,b.Amendment_No,b.Cust_DrgNo "

        Else
            strSelectSql = "Select b.Item_Code,b.Cust_DrgNo,b.Amendment_No,b.Cust_Ref,b.cust_drg_desc from Cust_Ord_hdr a,Cust_Ord_Dtl b"
            strSelectSql = strSelectSql & " where A.UNIT_CODE = B.UNIT_CODE AND b.Account_Code='" & Trim(pstrCustCode) & "' and b.Active_flag ='A' and "
            strSelectSql = strSelectSql & " a.Account_Code = b.Account_Code and a.Cust_ref = b.Cust_ref and "
            strSelectSql = strSelectSql & " a.Amendment_No = b.amendment_No  AND a.Authorized_Flag = 1 and a.PO_type in ('O','S') "
            strSelectSql = strSelectSql & " and a.Valid_date >='" & getDateForDB(GetServerDate()) & "' and effect_Date <= '" & getDateForDB(GetServerDate()) & "' and a.unit_code= '" & gstrUNITID & "'"
            strSelectSql = strSelectSql & " order by b.Cust_Ref,b.Amendment_No,b.Cust_DrgNo,b.Item_Code "
        End If
        rsCustOrdDtl = New ClsResultSetDB
        rsCustOrdDtl.GetResult(strSelectSql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        intRecordCount = rsCustOrdDtl.GetNoRows 'assign record count to integer variable
        If intRecordCount > 0 Then '          'if record found
            ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.WaitCursor)
            rsCustOrdDtl.MoveFirst() 'move to first record
            For intCount = 0 To intRecordCount - 1
                mListItemUserId = Me.lvwCustRefNo.Items.Add(rsCustOrdDtl.GetValue("Cust_Ref"))
                If mListItemUserId.SubItems.Count > 1 Then
                    mListItemUserId.SubItems(1).Text = rsCustOrdDtl.GetValue("Amendment_No")
                Else
                    mListItemUserId.SubItems.Insert(1, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsCustOrdDtl.GetValue("Amendment_No")))
                End If
                If mListItemUserId.SubItems.Count > 2 Then
                    mListItemUserId.SubItems(2).Text = rsCustOrdDtl.GetValue("Cust_DrgNo")
                Else
                    mListItemUserId.SubItems.Insert(2, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsCustOrdDtl.GetValue("Cust_DrgNo")))
                End If
                If mListItemUserId.SubItems.Count > 3 Then
                    mListItemUserId.SubItems(3).Text = rsCustOrdDtl.GetValue("cust_drg_desc")
                Else
                    mListItemUserId.SubItems.Insert(3, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsCustOrdDtl.GetValue("cust_drg_desc")))
                End If
                If mListItemUserId.SubItems.Count > 4 Then
                    mListItemUserId.SubItems(4).Text = rsCustOrdDtl.GetValue("Item_code")
                Else
                    mListItemUserId.SubItems.Insert(4, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsCustOrdDtl.GetValue("Item_code")))
                End If
                rsCustOrdDtl.MoveNext() 'move to next record
            Next intCount
            ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
            rsCustOrdDtl.ResultSetClose()
            rsCustOrdDtl = Nothing
        End If
        Me.ShowDialog()
        SelectDataFromCustOrd_Dtl = mstrItemText
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Sub frmMKTTRN0020_SOUTH_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error GoTo ErrHandler
        Me.Dispose()
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub lvwCustRefNo_ItemChecked(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckedEventArgs) Handles lvwCustRefNo.ItemChecked
    End Sub
    Private Sub lvwCustRefNo_ItemSelectionChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.ListViewItemSelectionChangedEventArgs) Handles lvwCustRefNo.ItemSelectionChanged
        On Error GoTo ErrHandler
        Dim intSubItem As Short
        mstrItemText = Trim(e.Item.Text)
        mstrItemText = "'" & Trim(e.Item.Text) & "','" & Trim(e.Item.SubItems(1).Text) & "'"
        CmdOk.Enabled = True
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub lvwCustRefNo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles lvwCustRefNo.KeyPress
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
    Private Sub SearchItem()
        Dim itmFound As System.Windows.Forms.ListViewItem ' FoundItem variable.
        On Error GoTo ErrHandler
        If optItem.Checked = True Then
            itmFound = SearchText((txtSearch.Text), optItem, lvwCustRefNo, "3")
        ElseIf optdrgNO.Checked = True Then
            itmFound = SearchText((txtSearch.Text), optdrgNO, lvwCustRefNo, "2")
        ElseIf OptItemCode.Checked = True Then
            itmFound = SearchText((txtSearch.Text), OptItemCode, lvwCustRefNo, "4")
        Else
            itmFound = SearchText((txtSearch.Text), optdrgNO, lvwCustRefNo)
        End If
        If itmFound Is Nothing Then ' If no match,
            Exit Sub
        Else
            itmFound.EnsureVisible() ' Scroll ListView to show found ListItem.
            itmFound.Selected = True ' Select the ListItem.
            ' Return focus to the control to see selection.
            lvwCustRefNo.Enabled = True
            If Len(txtSearch.Text) > 0 Then itmFound.Font = VB6.FontChangeBold(itmFound.Font, True)
        End If
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub optdrgNO_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optdrgNO.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo ErrHandler
            With lvwCustRefNo
                .Sort()
                ListViewColumnSorter.SortListView(lvwCustRefNo, 2, SortOrder.Ascending)
                .Sorting = System.Windows.Forms.SortOrder.Ascending
            End With
            Exit Sub
ErrHandler:
            gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub
    Private Sub optdrgNO_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles optdrgNO.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                txtSearch.Focus()
        End Select
        GoTo EventExitSub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub OptItemCode_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OptItemCode.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo ErrHandler
            With lvwCustRefNo
                .Sort()
                ListViewColumnSorter.SortListView(lvwCustRefNo, 4, SortOrder.Ascending)
                .Sorting = System.Windows.Forms.SortOrder.Ascending
            End With
            Exit Sub
ErrHandler:
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub
    Private Sub optItem_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optItem.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo ErrHandler
            With lvwCustRefNo
                .Sort()
                ListViewColumnSorter.SortListView(lvwCustRefNo, 3, SortOrder.Ascending)
                .Sorting = System.Windows.Forms.SortOrder.Ascending
            End With
            Exit Sub
ErrHandler:
            gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub
    Private Sub optItem_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles optItem.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                txtSearch.Focus()
        End Select
        GoTo EventExitSub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub optRefrence_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optRefrence.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo ErrHandler
            With lvwCustRefNo
                .Sort()
                ListViewColumnSorter.SortListView(lvwCustRefNo, 0, SortOrder.Ascending)
                .Sorting = System.Windows.Forms.SortOrder.Ascending
            End With
            Exit Sub
ErrHandler:
            gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub
    Private Sub optRefrence_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles optRefrence.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                txtSearch.Focus()
        End Select
        GoTo EventExitSub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub TxtSearch_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSearch.TextChanged
        Call SearchItem()
    End Sub
    Public Function SelectDataFromGrinDtl(ByRef pstrVendCode As String) As String
        Dim rsGrnDtl As ClsResultSetDB
        Dim strsql As String
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        On Error GoTo ErrHandler
        'form load functionality
        Call AddColumnsInListView()

        'Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(mdifrmMain.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2)
        'Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(mdifrmMain.Width) - VB6.PixelsToTwipsX(frmModules.Width)) / 2.3)
        optRefrence.Checked = True
        'end
        rsGrnDtl = New ClsResultSetDB
        strsql = "select a.Doc_No,a.Item_code,a.Rejected_Quantity from grn_Dtl a,grn_hdr b Where "
        strsql = strsql & "a.unit_code = b.unit_code and a.Doc_type = b.Doc_type And a.Doc_No = b.Doc_No and "
        strsql = strsql & "a.From_Location = b.From_Location and a.From_Location ='01R1'"
        strsql = strsql & "and a.Rejected_quantity > 0  and  b.Vendor_code = '" & pstrVendCode & "' and isnull(b.GRN_Cancelled,0) = 0 and a.unit_code='" & gstrUNITID & "' order by a.Doc_No"
        rsGrnDtl.GetResult(strsql)
        If rsGrnDtl.GetNoRows > 0 Then
            intMaxLoop = rsGrnDtl.GetNoRows : rsGrnDtl.MoveFirst()
            ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.WaitCursor)
            rsGrnDtl.MoveFirst() 'move to first record
            For intLoopCounter = 0 To intMaxLoop - 1
                mListItemUserId = Me.lvwCustRefNo.Items.Add(rsGrnDtl.GetValue("Doc_No"))
                If mListItemUserId.SubItems.Count > 3 Then
                    mListItemUserId.SubItems(3).Text = rsGrnDtl.GetValue("Item_code")
                Else
                    mListItemUserId.SubItems.Insert(3, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsGrnDtl.GetValue("Item_code")))
                End If
                rsGrnDtl.MoveNext() 'move to next record
            Next
            ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
            rsGrnDtl.ResultSetClose()
            rsGrnDtl = Nothing
        End If
        Me.ShowDialog()
        SelectDataFromGrinDtl = mstrItemText
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
End Class