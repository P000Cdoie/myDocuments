Option Strict Off
Option Explicit On
Friend Class frmMKTTRN0020
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
	'Revised By      : Ashutosh Verma , Issue Id:16185
	'Revision Date   : 10-11-2005
	'History         : Sets the focus on Search button on showing(load time) the form.
	'-----------------------------------------------------------------------------------
	'Revised By      : Davinder Singh
	'Issue ID        : 19573
	'Revision Date   : 27 Feb 2007
	'History         : SO help is parameterized to show Items along with SO no.
	'                  if 'Display_SOItems' flag in 'Sales_Parameter' is TRUE
	'                  else show SO Nos and Amendment Nos only
    '-----------------------------------------------------------------------------------
    'Revised By      : Siddharth Ranjan
    'Issue ID        : eMpro-20090922-36611
    'Revision Date   : 22 Sep 2009
    'History         : Error in search when text not found
    '-----------------------------------------------------------------
    '***********************************************************************************
    'Revised By        -    Vinod Singh
    'Revision Date     -    09/05/2011
    'Revision History  -    Changes for Multi Unit
    '***********************************************************************************
    'REVISED BY     -  PRASHANT RAJPAL
    'REVISED ON     -  18 SEP 2015
    'PURPOSE        -  10869290 -SERVICE INVOICE 
    '-------------------------------------------------------------------------------------------------------------------------------------------------------

    Dim mCtlHdrItemCode As System.Windows.Forms.ColumnHeader
    Dim mCtlHdrDrawingNo As System.Windows.Forms.ColumnHeader
    Dim mCtlHdrCustRef As System.Windows.Forms.ColumnHeader
    Dim mCtlHdrAmmendment As System.Windows.Forms.ColumnHeader
    Dim mListItemUserId As System.Windows.Forms.ListViewItem
    Dim mstrItemText As String
    Dim mblnDisplay_SOItems As Boolean
    Public Export_Invoice_Flag As Boolean = False
    Public Customer_Invoice_Flag As Boolean = False
    Public strCustCode As String
    Public strInvType As String
    Public strRefAmm As String
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
    Private Sub frmMKTTRN0020_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo ErrHandler
        If lvwCustRefNo.Enabled Then
            lvwCustRefNo.Focus()
            ''Changes done by Ashutosh on 10-11-2005, Issue Id:16185
            txtSearch.Enabled = True
            txtSearch.Focus()
            ''Changes on 10-11-2005 end here.
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub frmMKTTRN0020_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo ErrHandler
        '''''''SetBackGroundColorNew(Me, True)
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub AddColumnsInListView()
        '-----------------------------------------------------------------------------------
        'Revised By      : Davinder Singh
        'Issue ID        : 19573
        'Revision Date   : 27 Feb 2007
        'History         : SO help is parameterized to show Items along with SO no.
        '                  if 'Display_SOItems' flag in 'Sales_Parameter' is TRUE
        '                  else show SO Nos and Amendment Nos only
        '-----------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        With Me.lvwCustRefNo
            If mblnDisplay_SOItems = True Then
                mCtlHdrCustRef = .Columns.Add("")
                mCtlHdrCustRef.Text = "Refrence No."
                mCtlHdrCustRef.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(lvwCustRefNo.Width) / 4)
                mCtlHdrAmmendment = .Columns.Add("")
                mCtlHdrAmmendment.Text = "Amendment No."
                mCtlHdrAmmendment.Width = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(lvwCustRefNo.Width) / 4) - 50)
                mCtlHdrDrawingNo = .Columns.Add("")
                mCtlHdrDrawingNo.Text = "Drawing No."
                mCtlHdrDrawingNo.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(lvwCustRefNo.Width) / 4)
                mCtlHdrItemCode = .Columns.Add("")
                mCtlHdrItemCode.Text = "Item Code"
                mCtlHdrItemCode.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(lvwCustRefNo.Width) / 4)
            Else
                mCtlHdrCustRef = .Columns.Add("")
                mCtlHdrCustRef.Text = "Refrence No."
                mCtlHdrCustRef.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(lvwCustRefNo.Width) / 1.8)
                mCtlHdrAmmendment = .Columns.Add("")
                mCtlHdrAmmendment.Text = "Amendment No."
                mCtlHdrAmmendment.Width = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(lvwCustRefNo.Width) / 2.6))
            End If
        End With
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Public Function SelectDataFromCustOrd_Dtl(ByRef pstrCustCode As String, ByRef pstrInvType As String) As String
        '-----------------------------------------------------------------------------------
        'Revised By      : Davinder Singh
        'Issue ID        : 19573
        'Revision Date   : 27 Feb 2007
        'History         : SO help is parameterized to show Items along with SO no.
        '                  if 'Display_SOItems' flag in 'Sales_Parameter' is TRUE
        '                  else show SO Nos and Amendment Nos only
        '-----------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim strSelectSql As String
        Dim rsCustOrdDtl As ClsResultSetDB
        Dim intRecordCount As Short
        Dim intCount As Short
        Dim Display_SOItems As Boolean
        mblnDisplay_SOItems = ShowItemsWithSO()
        Call AddColumnsInListView()

        'Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(mdifrmMain.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2)
        'Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(mdifrmMain.Width) - VB6.PixelsToTwipsX(frmModules.Width)) / 2.3)
        optRefrence.Checked = True
        Me.Size = New System.Drawing.Point(501, 233)
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.WaitCursor)
        strSelectSql = ""
        If mblnDisplay_SOItems = True Then
            If UCase(pstrInvType) = "JOBWORK INVOICE" Then
                strSelectSql = "Select b.Item_Code,b.Cust_DrgNo,b.Amendment_No,b.Cust_Ref from Cust_Ord_hdr a,Cust_Ord_Dtl b"
                strSelectSql = strSelectSql & " where a.unit_code = b.unit_code and b.Account_Code='" & Trim(pstrCustCode) & "' and b.Active_flag ='A' and "
                strSelectSql = strSelectSql & " a.Account_Code = b.Account_Code and a.Cust_ref =b.Cust_ref and "
                strSelectSql = strSelectSql & " a.Amendment_No = b.amendment_No AND a.Authorized_Flag = 1 and a.PO_type in ('J') "
                strSelectSql = strSelectSql & " and a.Valid_date >='" & getDateForDB(GetServerDate()) & "' and effect_Date <='" & getDateForDB(GetServerDate()) & "' and a.unit_code='" & gstrUNITID & "'"
                strSelectSql = strSelectSql & " order by b.Cust_Ref,b.Amendment_No,b.Cust_DrgNo,b.Item_Code "
            ElseIf UCase(pstrInvType) = "EXPORT INVOICE" Then
                strSelectSql = "Select b.Item_Code,b.Cust_DrgNo,b.Amendment_No,b.Cust_Ref from Cust_Ord_hdr a,Cust_Ord_Dtl b"
                strSelectSql = strSelectSql & " where a.unit_code=b.unit_code and a.unit_code='" & gstrUNITID & "' and b.Account_Code='" & Trim(pstrCustCode) & "' and b.Active_flag ='A' and "
                strSelectSql = strSelectSql & " a.Account_Code = b.Account_Code and a.Cust_ref =b.Cust_ref and "
                strSelectSql = strSelectSql & " a.Amendment_No = b.amendment_No AND a.Authorized_Flag = 1 and a.PO_type in ('E') "
                strSelectSql = strSelectSql & " and a.Valid_date >='" & getDateForDB(GetServerDate()) & "' and effect_date <='" & getDateForDB(GetServerDate()) & "'"
                strSelectSql = strSelectSql & " order by b.Cust_Ref,b.Amendment_No,b.Cust_DrgNo,b.Item_Code "
                '10869290 -SERVICE INVOICE 
            ElseIf UCase(pstrInvType) = "SERVICE INVOICE" Then
                strSelectSql = "Select b.Item_Code,b.Cust_DrgNo,b.Amendment_No,b.Cust_Ref from Cust_Ord_hdr a,Cust_Ord_Dtl b"
                strSelectSql = strSelectSql & " where a.unit_code=b.unit_code and a.unit_code='" & gstrUNITID & "' and b.Account_Code='" & Trim(pstrCustCode) & "' and b.Active_flag ='A' and "
                strSelectSql = strSelectSql & " a.Account_Code = b.Account_Code and a.Cust_ref =b.Cust_ref and "
                strSelectSql = strSelectSql & " a.Amendment_No = b.amendment_No AND a.Authorized_Flag = 1 and a.PO_type in ('V') "
                strSelectSql = strSelectSql & " and a.Valid_date >='" & getDateForDB(GetServerDate()) & "' and effect_date <='" & getDateForDB(GetServerDate()) & "'"
                strSelectSql = strSelectSql & " order by b.Cust_Ref,b.Amendment_No,b.Cust_DrgNo,b.Item_Code "
                '10869290 -SERVICE INVOICE 
            Else
                strSelectSql = "Select b.Item_Code,b.Cust_DrgNo,b.Amendment_No,b.Cust_Ref from Cust_Ord_hdr a,Cust_Ord_Dtl b"
                strSelectSql = strSelectSql & " where a.unit_code = b.unit_code and b.Account_Code='" & Trim(pstrCustCode) & "' and b.Active_flag ='A' and a.unit_code='" & gstrUNITID & "'"
                strSelectSql = strSelectSql & " and a.Account_Code = b.Account_Code and a.Cust_ref = b.Cust_ref and "
                strSelectSql = strSelectSql & " a.Amendment_No = b.amendment_No  AND a.Authorized_Flag = 1 and a.PO_type in ('O','S','M','Q') "
                strSelectSql = strSelectSql & " and a.Valid_date >='" & getDateForDB(GetServerDate()) & "' and effect_Date <='" & getDateForDB(GetServerDate()) & "'"
                strSelectSql = strSelectSql & " order by b.Cust_Ref,b.Amendment_No,b.Cust_DrgNo,b.Item_Code "
            End If
        Else
            If UCase(pstrInvType) = "JOBWORK INVOICE" Then
                strSelectSql = "SELECT CUST_REF,AMENDMENT_NO"
                strSelectSql = strSelectSql & " FROM CUST_ORD_HDR"
                strSelectSql = strSelectSql & " WHERE UNIT_CODE='" & gstrUNITID & "' AND ACCOUNT_CODE='" & Trim(pstrCustCode) & "'"
                strSelectSql = strSelectSql & " AND ACTIVE_FLAG='A'"
                strSelectSql = strSelectSql & " AND PO_TYPE='J'"
                strSelectSql = strSelectSql & " AND AUTHORIZED_FLAG = 1"
                strSelectSql = strSelectSql & " AND VALID_DATE >='" & getDateForDB(GetServerDate()) & "'"
                strSelectSql = strSelectSql & " AND EFFECT_DATE <='" & getDateForDB(GetServerDate()) & "'"
                strSelectSql = strSelectSql & " ORDER BY CUST_REF,AMENDMENT_NO"
            ElseIf UCase(pstrInvType) = "EXPORT INVOICE" Then
                strSelectSql = "SELECT CUST_REF,AMENDMENT_NO"
                strSelectSql = strSelectSql & " FROM CUST_ORD_HDR"
                strSelectSql = strSelectSql & " WHERE UNIT_CODE='" & gstrUNITID & "' AND ACCOUNT_CODE='" & Trim(pstrCustCode) & "'"
                strSelectSql = strSelectSql & " AND ACTIVE_FLAG='A'"
                strSelectSql = strSelectSql & " AND PO_TYPE='E'"
                strSelectSql = strSelectSql & " AND AUTHORIZED_FLAG = 1"
                strSelectSql = strSelectSql & " AND VALID_DATE >='" & getDateForDB(GetServerDate()) & "'"
                strSelectSql = strSelectSql & " AND EFFECT_DATE <='" & getDateForDB(GetServerDate()) & "'"
                strSelectSql = strSelectSql & " ORDER BY CUST_REF,AMENDMENT_NO"
            ElseIf UCase(pstrInvType) = "SERVICE INVOICE" Then
                strSelectSql = "SELECT CUST_REF,AMENDMENT_NO"
                strSelectSql = strSelectSql & " FROM CUST_ORD_HDR"
                strSelectSql = strSelectSql & " WHERE UNIT_CODE='" & gstrUNITID & "' AND ACCOUNT_CODE='" & Trim(pstrCustCode) & "'"
                strSelectSql = strSelectSql & " AND ACTIVE_FLAG='A'"
                strSelectSql = strSelectSql & " AND PO_TYPE='V'"
                strSelectSql = strSelectSql & " AND AUTHORIZED_FLAG = 1"
                strSelectSql = strSelectSql & " AND VALID_DATE >='" & getDateForDB(GetServerDate()) & "'"
                strSelectSql = strSelectSql & " AND EFFECT_DATE <='" & getDateForDB(GetServerDate()) & "'"
                strSelectSql = strSelectSql & " ORDER BY CUST_REF,AMENDMENT_NO"
            Else
                strSelectSql = "SELECT CUST_REF,AMENDMENT_NO"
                strSelectSql = strSelectSql & " FROM CUST_ORD_HDR"
                strSelectSql = strSelectSql & " WHERE UNIT_CODE='" & gstrUNITID & "' AND ACCOUNT_CODE='" & Trim(pstrCustCode) & "'"
                strSelectSql = strSelectSql & " AND ACTIVE_FLAG='A'"
                strSelectSql = strSelectSql & " AND PO_TYPE IN ('O','S','M','Q')"
                strSelectSql = strSelectSql & " AND AUTHORIZED_FLAG = 1"
                strSelectSql = strSelectSql & " AND VALID_DATE >='" & getDateForDB(GetServerDate()) & "'"
                strSelectSql = strSelectSql & " AND EFFECT_DATE <='" & getDateForDB(GetServerDate()) & "'"
                strSelectSql = strSelectSql & " ORDER BY CUST_REF,AMENDMENT_NO"
            End If
        End If
        rsCustOrdDtl = New ClsResultSetDB
        If rsCustOrdDtl.GetResult(strSelectSql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly) = False Then GoTo ErrHandler
        intRecordCount = rsCustOrdDtl.GetNoRows
        If intRecordCount > 0 Then
            rsCustOrdDtl.MoveFirst()
            For intCount = 1 To intRecordCount
                mListItemUserId = lvwCustRefNo.Items.Add(rsCustOrdDtl.GetValue("Cust_Ref"))
                If mListItemUserId.SubItems.Count > 1 Then
                    mListItemUserId.SubItems(1).Text = rsCustOrdDtl.GetValue("Amendment_No")
                Else
                    mListItemUserId.SubItems.Insert(1, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsCustOrdDtl.GetValue("Amendment_No")))
                End If
                If mblnDisplay_SOItems = True Then
                    If mListItemUserId.SubItems.Count > 2 Then
                        mListItemUserId.SubItems(2).Text = rsCustOrdDtl.GetValue("Cust_DrgNo")
                    Else
                        mListItemUserId.SubItems.Insert(2, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsCustOrdDtl.GetValue("Cust_DrgNo")))
                    End If
                    If mListItemUserId.SubItems.Count > 3 Then
                        mListItemUserId.SubItems(3).Text = rsCustOrdDtl.GetValue("Item_code")
                    Else
                        mListItemUserId.SubItems.Insert(3, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsCustOrdDtl.GetValue("Item_code")))
                    End If
                End If
                rsCustOrdDtl.MoveNext()
            Next intCount
        End If
        rsCustOrdDtl.ResultSetClose()
        rsCustOrdDtl = Nothing
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Me.ShowDialog()
        SelectDataFromCustOrd_Dtl = mstrItemText
        Exit Function
ErrHandler:
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Sub frmMKTTRN0020_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
    End Sub
    Private Sub lvwCustRefNo_ItemSelectionChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.ListViewItemSelectionChangedEventArgs) Handles lvwCustRefNo.ItemSelectionChanged
        On Error GoTo ErrHandler
        If e.IsSelected Then
            mstrItemText = Trim(e.Item.Text)
            mstrItemText = "'" & Trim(e.Item.Text) & "','" & Trim(e.Item.SubItems(1).Text) & "'"
            CmdOk.Enabled = True
        End If
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
        '---------------------------------------------------------------------
        'Created By     -   Shruti Khanna\(Name Changed - Nitin Sood)
        '---------------------------------------------------------------------
        Dim itmFound As System.Windows.Forms.ListViewItem ' FoundItem variable.
        On Error GoTo ErrHandler
        If optItem.Checked = True Then
            itmFound = SearchText((txtSearch.Text), optItem, lvwCustRefNo, "3")
        ElseIf optdrgNO.Checked = True Then
            itmFound = SearchText((txtSearch.Text), optdrgNO, lvwCustRefNo, "2")
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
        If Err.Number = 5 Then
            MsgBox("Searched Text Not Found", MsgBoxStyle.Information)
        Else
            gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
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
    Private Sub optItem_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optItem.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo ErrHandler
            With lvwCustRefNo
                .Sort()
                '.SortKey = 3
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
                '.SortKey = 0
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
        mblnDisplay_SOItems = ShowItemsWithSO()
        Call AddColumnsInListView()

        'Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(mdifrmMain.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2)
        'Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(mdifrmMain.Width) - VB6.PixelsToTwipsX(frmModules.Width)) / 2.3)
        optRefrence.Checked = True
        rsGrnDtl = New ClsResultSetDB
        strsql = "select a.Doc_No,a.Item_code,a.Rejected_Quantity from grn_Dtl a,grn_hdr b Where A.UNIT_CODE = B.UNIT_CODE AND "
        strsql = strsql & "a.Doc_type = b.Doc_type And a.Doc_No = b.Doc_No and "
        strsql = strsql & "a.From_Location = b.From_Location and a.From_Location ='01R1'"
        strsql = strsql & " AND A.UNIT_CODE='" + gstrUNITID + "' and a.Rejected_quantity > 0  and  b.Vendor_code = '" & pstrVendCode & "' and isnull(b.GRN_Cancelled,0) = 0 order by a.Doc_No"
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
                    mListItemUserId.SubItems.Insert(1, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, ""))
                    mListItemUserId.SubItems.Insert(2, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, ""))
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
    Private Function ShowItemsWithSO() As Boolean
        '-----------------------------------------------------------------------------------
        'Created By      : Davinder Singh
        'Issue ID        : 19573
        'Creation Date   : 27 Feb 2007
        'Function        : To Read 'Display_SOItems' parameter from Sales_Parameter
        '-----------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim strsql As String
        Dim Rs As ClsResultSetDB
        ShowItemsWithSO = False
        Rs = New ClsResultSetDB
        strsql = "Select Display_SOItems from Sales_Parameter WHERE UNIT_CODE='" & gstrUNITID & "'"
        If Rs.GetResult(strsql) = False Then GoTo ErrHandler
        If Rs.GetValue("Display_SOItems") = "True" Then
            ShowItemsWithSO = True
        End If
        Rs.ResultSetClose()
        Rs = Nothing
ErrHandler:
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Rs = Nothing
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    '''this method was not present in multi unit ..so this is picked from sinle unit by SACHIN TYAGI ON 1 March 2014 

    Public Function SelectDataFromCustOrd_Dtl_SUBASSEMBLY(ByRef pstrCustCode As String, ByRef pstrInvType As String, ByRef pstrInvSubType As String) As String
        '-----------------------------------------------------------------------------------
        'Revised By      : Davinder Singh
        'Issue ID        : 19573
        'Revision Date   : 27 Feb 2007
        'History         : SO help is parameterized to show Items along with SO no.
        '                  if 'Display_SOItems' flag in 'Sales_Parameter' is TRUE
        '                  else show SO Nos and Amendment Nos only
        '-----------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim strSelectSql As String
        Dim rsCustOrdDtl As ClsResultSetDB
        Dim intRecordCount As Short
        Dim intCount As Short
        Dim Display_SOItems As Boolean

        ''If form_load_flag = False Then
        'Dim sender As Object
        'Dim e As EventArgs
        'frmMKTTRN0020_Load(sender, e)
        ''End If

        ' lvwCustRefNo.Items.Clear()

        mblnDisplay_SOItems = ShowItemsWithSO()
        Call AddColumnsInListView()
        ''''  Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(mdifrmMain.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2)
        '' comented by sachin for testing 
        '''''''''Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(mdifrmMain.Width) - VB6.PixelsToTwipsX(frmModules.Width)) / 2.3)
        optRefrence.Checked = True
        Me.Size = New System.Drawing.Point(501, 233)

        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.WaitCursor)

        strSelectSql = ""
        If mblnDisplay_SOItems = True Then
            If UCase(pstrInvType) = "JOBWORK INVOICE" Then
                strSelectSql = "Select b.Item_Code,b.Cust_DrgNo,b.Amendment_No,b.Cust_Ref from Cust_Ord_hdr a,Cust_Ord_Dtl b"
                strSelectSql = strSelectSql & " where a.UNIT_CODE='" + gstrUNITID + "' and b.Account_Code='" & Trim(pstrCustCode) & "' and b.Active_flag ='A' and "
                strSelectSql = strSelectSql & " a.Account_Code = b.Account_Code and a.UNIT_CODE=b.UNIT_CODE and  a.Cust_ref =b.Cust_ref and "
                'Code Commented And Added By        -   Nitin Sood
                'Active Flag to Be Checked Item Wise and Not for Sales Order
                'strSelectSql = strSelectSql & " a.Amendment_No = b.amendment_No and a.Active_flag =b.Active_flag AND a.Authorized_Flag = 1 and a.PO_type in ('J') "
                strSelectSql = strSelectSql & " a.Amendment_No = b.amendment_No AND a.Authorized_Flag = 1 and a.PO_type in ('J') "
                strSelectSql = strSelectSql & " and a.Valid_date >='" & GetServerDate() & "' and effect_Date <='" & GetServerDate() & "'"
                strSelectSql = strSelectSql & " order by b.Cust_Ref,b.Amendment_No,b.Cust_DrgNo,b.Item_Code "
            ElseIf UCase(pstrInvType) = "EXPORT INVOICE" Then
                strSelectSql = "Select b.Item_Code,b.Cust_DrgNo,b.Amendment_No,b.Cust_Ref from Cust_Ord_hdr a,Cust_Ord_Dtl b"
                strSelectSql = strSelectSql & " where a.UNIT_CODE='" + gstrUNITID + "' and b.Account_Code='" & Trim(pstrCustCode) & "' and b.Active_flag ='A' and "
                strSelectSql = strSelectSql & " a.Account_Code = b.Account_Code and a.unit_code=b.unit_code and a.Cust_ref =b.Cust_ref and "
                'Code Commented And Added By        -   Nitin Sood
                'Active Flag to Be Checked Item Wise and Not for Sales Order
                'strSelectSql = strSelectSql & " a.Amendment_No = b.amendment_No and a.Active_flag =b.Active_flag AND a.Authorized_Flag = 1 a.PO_type in ('E')"
                strSelectSql = strSelectSql & " a.Amendment_No = b.amendment_No AND a.Authorized_Flag = 1 and a.PO_type in ('E') "
                strSelectSql = strSelectSql & " and a.Valid_date >='" & GetServerDate() & "' and effect_date <='" & GetServerDate() & "'"
                strSelectSql = strSelectSql & " order by b.Cust_Ref,b.Amendment_No,b.Cust_DrgNo,b.Item_Code "
            Else
                strSelectSql = "Select b.Item_Code,b.Cust_DrgNo,b.Amendment_No,b.Cust_Ref from Cust_Ord_hdr a,Cust_Ord_Dtl b"
                strSelectSql = strSelectSql & " where UNIT_CODE='" + gstrUNITID + "' and b.Account_Code='" & Trim(pstrCustCode) & "' and b.Active_flag ='A' and "
                strSelectSql = strSelectSql & " a.Account_Code = b.Account_Code and a.unit_code=b.unit_code and a.Cust_ref =b.Cust_ref and "
                'Code Commented And Added By        -   Nitin Sood
                'Active Flag to Be Checked Item Wise and Not for Sales Order
                'strSelectSql = strSelectSql & " a.Amendment_No = b.amendment_No and a.Active_flag =b.Active_flag AND a.Authorized_Flag = 1 and a.PO_type in ('O','S')"
                '''***** Changes done By Ashutosh on 01-06-2006, Issue Id:17610, Include items of So type --'M'
                strSelectSql = strSelectSql & " a.Amendment_No = b.amendment_No  AND a.Authorized_Flag = 1 and a.PO_type in ('A') "
                '''**** end here
                strSelectSql = strSelectSql & " and a.Valid_date >='" & GetServerDate() & "' and effect_Date <='" & GetServerDate() & "'"
                strSelectSql = strSelectSql & " order by b.Cust_Ref,b.Amendment_No,b.Cust_DrgNo,b.Item_Code "
            End If
        Else
            If UCase(pstrInvType) = "JOBWORK INVOICE" Then
                strSelectSql = "SELECT CUST_REF,AMENDMENT_NO"
                strSelectSql = strSelectSql & " FROM CUST_ORD_HDR"
                strSelectSql = strSelectSql & " WHERE ACCOUNT_CODE='" & Trim(pstrCustCode) & "'"
                strSelectSql = strSelectSql & " AND ACTIVE_FLAG='A' and UNIT_CODE='" + gstrUNITID + "' "
                strSelectSql = strSelectSql & " AND PO_TYPE='J'"
                strSelectSql = strSelectSql & " AND AUTHORIZED_FLAG = 1"
                strSelectSql = strSelectSql & " AND VALID_DATE >='" & GetServerDate() & "'"
                strSelectSql = strSelectSql & " AND EFFECT_DATE <='" & GetServerDate() & "'"
                strSelectSql = strSelectSql & " ORDER BY CUST_REF,AMENDMENT_NO"
            ElseIf UCase(pstrInvType) = "EXPORT INVOICE" Then
                strSelectSql = "SELECT CUST_REF,AMENDMENT_NO"
                strSelectSql = strSelectSql & " FROM CUST_ORD_HDR"
                strSelectSql = strSelectSql & " WHERE ACCOUNT_CODE='" & Trim(pstrCustCode) & "'"
                strSelectSql = strSelectSql & " AND ACTIVE_FLAG='A' and UNIT_CODE='" + gstrUNITID + "' "
                strSelectSql = strSelectSql & " AND PO_TYPE='E'"
                strSelectSql = strSelectSql & " AND AUTHORIZED_FLAG = 1"
                strSelectSql = strSelectSql & " AND VALID_DATE >='" & GetServerDate() & "'"
                strSelectSql = strSelectSql & " AND EFFECT_DATE <='" & GetServerDate() & "'"
                strSelectSql = strSelectSql & " ORDER BY CUST_REF,AMENDMENT_NO"
            ElseIf UCase(pstrInvType) = "SERVICE INVOICE" Then
                strSelectSql = "SELECT CUST_REF,AMENDMENT_NO"
                strSelectSql = strSelectSql & " FROM CUST_ORD_HDR"
                strSelectSql = strSelectSql & " WHERE ACCOUNT_CODE='" & Trim(pstrCustCode) & "'"
                strSelectSql = strSelectSql & " AND ACTIVE_FLAG='A' and UNIT_CODE='" + gstrUNITID + "' "
                strSelectSql = strSelectSql & " AND PO_TYPE='V'"
                strSelectSql = strSelectSql & " AND AUTHORIZED_FLAG = 1"
                strSelectSql = strSelectSql & " AND VALID_DATE >='" & GetServerDate() & "'"
                strSelectSql = strSelectSql & " AND EFFECT_DATE <='" & GetServerDate() & "'"
                strSelectSql = strSelectSql & " ORDER BY CUST_REF,AMENDMENT_NO"
            Else
                strSelectSql = "SELECT CUST_REF,AMENDMENT_NO"
                strSelectSql = strSelectSql & " FROM CUST_ORD_HDR"
                strSelectSql = strSelectSql & " WHERE ACCOUNT_CODE='" & Trim(pstrCustCode) & "'"
                strSelectSql = strSelectSql & " AND ACTIVE_FLAG='A' and UNIT_CODE='" + gstrUNITID + "' "
                strSelectSql = strSelectSql & " AND PO_TYPE IN ('A')"
                strSelectSql = strSelectSql & " AND AUTHORIZED_FLAG = 1"
                strSelectSql = strSelectSql & " AND VALID_DATE >='" & GetServerDate() & "'"
                strSelectSql = strSelectSql & " AND EFFECT_DATE <='" & GetServerDate() & "'"
                strSelectSql = strSelectSql & " ORDER BY CUST_REF,AMENDMENT_NO"
            End If
        End If

        rsCustOrdDtl = New ClsResultSetDB
        If rsCustOrdDtl.GetResult(strSelectSql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly) = False Then GoTo ErrHandler
        intRecordCount = rsCustOrdDtl.GetNoRows
        If intRecordCount > 0 Then
            rsCustOrdDtl.MoveFirst()
            For intCount = 1 To intRecordCount
                'UPGRADE_WARNING: Couldn't resolve default property of object rsCustOrdDtl.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                mListItemUserId = lvwCustRefNo.Items.Add(rsCustOrdDtl.GetValue("Cust_Ref"))
                'UPGRADE_WARNING: Lower bound of collection mListItemUserId has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                'UPGRADE_WARNING: Couldn't resolve default property of object rsCustOrdDtl.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If mListItemUserId.SubItems.Count > 1 Then
                    mListItemUserId.SubItems(1).Text = rsCustOrdDtl.GetValue("Amendment_No")
                Else
                    mListItemUserId.SubItems.Insert(1, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsCustOrdDtl.GetValue("Amendment_No")))
                End If
                If mblnDisplay_SOItems = True Then
                    'UPGRADE_WARNING: Lower bound of collection mListItemUserId has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object rsCustOrdDtl.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If mListItemUserId.SubItems.Count > 2 Then
                        mListItemUserId.SubItems(2).Text = rsCustOrdDtl.GetValue("Cust_DrgNo")
                    Else
                        mListItemUserId.SubItems.Insert(2, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsCustOrdDtl.GetValue("Cust_DrgNo")))
                    End If
                    'UPGRADE_WARNING: Lower bound of collection mListItemUserId has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object rsCustOrdDtl.GetValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If mListItemUserId.SubItems.Count > 3 Then
                        mListItemUserId.SubItems(3).Text = rsCustOrdDtl.GetValue("Item_code")
                    Else
                        mListItemUserId.SubItems.Insert(3, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsCustOrdDtl.GetValue("Item_code")))
                    End If
                End If
                rsCustOrdDtl.MoveNext()
            Next intCount
        End If
        rsCustOrdDtl.ResultSetClose()
        'UPGRADE_NOTE: Object rsCustOrdDtl may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        rsCustOrdDtl = Nothing
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Me.ShowDialog()
        SelectDataFromCustOrd_Dtl_SUBASSEMBLY = mstrItemText
        Exit Function
ErrHandler:
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function

End Class
