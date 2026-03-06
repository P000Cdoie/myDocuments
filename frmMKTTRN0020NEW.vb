Option Strict Off
Option Explicit On
Friend Class frmMKTTRN0020NEW
	Inherits System.Windows.Forms.Form
	'===================================================================================
	' (c) MothersonSumi INfotech & Design Ltd. All rights reserverd.
	' File Name         :   FRMMKTTRN0020NEW.frm
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
	'Revised By     : Arul Mozhi varman
	'Revised On     : 26-11-2005
	'Revised Reason : Amendmenty No Column removed from the order by class due to the query not giving the record set fom some conditions
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
    Dim Intcounter As Short
	Dim INVTYPE As String
    Private RSSO As New ADODB.Recordset
    Dim ACTIDX As Short
    Dim bool_Item_Check As Boolean = False
    Dim blnInvoiceUpload As Boolean = False
    Public strTmpTable As String = "R_Temp_Emp_InvoiceSOLinkage_" + gstrUNITID
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		On Error GoTo ErrHandler
        mstrItemText = "" 'User CANCELS Form
        blnInvoiceUpload = False
		Me.Close()
		Exit Sub
ErrHandler: 'The Error Handling Code Starts here
		Call gobjError.RaiseError(Err.Number, ERR.Source, Err.Description, mP_Connection) 'Function called, if error occurred
	End Sub
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		On Error GoTo ErrHandler
		
		CUSTREFLIST = ""
        MULTIPLESO = 0
        For Intcounter = 0 To Me.lvwCustRefNo.Items.Count - 1
            If Me.lvwCustRefNo.Items.Item(Intcounter).Checked = True Then MULTIPLESO = MULTIPLESO + 1
        Next Intcounter
        If MULTIPLESO = 0 Then
            MsgBox("Select at Least One Sales Order to prepare the Invoice.", MsgBoxStyle.Information, "empower")
            If Me.lvwCustRefNo.Visible = True Then Me.lvwCustRefNo.Focus() : Exit Sub
        Else
            mP_Connection.Execute("if exists(select name from sysobjects where name = '" + strTmpTable + "') drop table " + strTmpTable + "", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            mP_Connection.Execute("create table " + strTmpTable + " (Doc_No int,Cust_Ref varchar(25),Amendment_No varchar(25))", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            For Intcounter = 0 To Me.lvwCustRefNo.Items.Count - 1
                If Me.lvwCustRefNo.Items.Item(Intcounter).Checked = True Then
                    If INVTYPE = "NORMAL INVOICE" Then
                        mP_Connection.Execute("insert into " + strTmpTable + " (Cust_Ref,Amendment_No) values('" & lvwCustRefNo.Items.Item(Intcounter).Text & "','" & lvwCustRefNo.Items.Item(Intcounter).SubItems(1).Text & "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End If
                    CUSTREFLIST = CUSTREFLIST & Me.lvwCustRefNo.Items.Item(Intcounter).Text & ";"
                End If
            Next Intcounter
            CUSTREFLIST = Mid(CUSTREFLIST, 1, Len(CUSTREFLIST) - 1)
            CUSTREFLIST = Replace(CUSTREFLIST, ";", Chr(39) & "," & Chr(39))
            Me.Close()
        End If
        blnInvoiceUpload = False
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub frmMKTTRN0020NEW_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo ErrHandler
        If lvwCustRefNo.Enabled Then
            lvwCustRefNo.Focus()
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub frmMKTTRN0020NEW_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
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
            'Code Added By Arul on 07-03-2005
            .CheckBoxes = True
            'Addition ends here
            mCtlHdrCustRef = .Columns.Add("")
            mCtlHdrCustRef.Text = "Refrence No"
            mCtlHdrCustRef.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(lvwCustRefNo.Width) / 5)
            mCtlHdrAmmendment = .Columns.Add("")
            mCtlHdrAmmendment.Text = "Amend No."
            mCtlHdrAmmendment.Width = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(lvwCustRefNo.Width) / 5) - 500)
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
    Private Sub AddColumnsInListViewInvoiceUpload()
        '***********************************
        'To add Columns Headers in the ListView in the form load
        '***********************************
        On Error GoTo ErrHandler
        With Me.lvwCustRefNo
            'Code Added By Arul on 07-03-2005
            .CheckBoxes = True
            'Addition ends here
            mCtlHdrCustRef = .Columns.Add("")
            mCtlHdrCustRef.Text = "Refrence No"
            mCtlHdrCustRef.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(lvwCustRefNo.Width) / 5)
            mCtlHdrAmmendment = .Columns.Add("")
            mCtlHdrAmmendment.Text = "Amend No."
            mCtlHdrAmmendment.Width = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(lvwCustRefNo.Width) / 5) - 500)
            mCtlHdrDrawingNo = .Columns.Add("")
          
        End With
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Public Function SelectDataFromCustOrd_Dtl(ByRef pstrCustCode As String, ByRef pstrInvType As String) As String
        '***********************************
        'To Get Data From Cust_Ord_Dtl
        '***********************************
        On Error GoTo ErrHandler
        Dim strSelectSql As String 'Declared To Make Select Query
        Dim rsCustOrdDtl As ClsResultSetDB
        Dim intRecordCount As Short 'To Hold Record Count
        Dim intCount As Short
        'form load funcionality
        Call AddColumnsInListView()

        'Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(mdifrmMain.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2)
        'Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(mdifrmMain.Width) - VB6.PixelsToTwipsX(frmModules.Width)) / 2.8)
        optRefrence.Checked = True
        RSSO = mP_Connection.Execute("select MultipleSO from Sales_Parameter WHERE UNIT_CODE='" + gstrUNITID + "'")
        'end
        INVTYPE = pstrInvType
        Me.lvwCustRefNo.Items.Clear() 'initially clear all items in the listview
        If UCase(pstrInvType) = "JOBWORK INVOICE" Then
            strSelectSql = "Select b.Item_Code,b.Cust_DrgNo,b.Amendment_No,b.Cust_Ref,b.cust_drg_desc from Cust_Ord_hdr a (NOLOCK),Cust_Ord_Dtl b (NOLOCK)"
            strSelectSql = strSelectSql & " Where A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" + gstrUNITID + "' AND a.Account_Code = b.Account_Code and a.Cust_ref =b.Cust_ref "
            strSelectSql = strSelectSql & " and a.Amendment_No = b.amendment_No  AND a.Authorized_Flag = 1 and a.PO_type in ('J') "
            strSelectSql = strSelectSql & " and b.Account_Code='" & Trim(pstrCustCode) & "' and b.Active_flag ='A' "
            strSelectSql = strSelectSql & " and a.Valid_date >='" & getDateForDB(GetServerDate()) & "' and effect_Date <= '" & getDateForDB(GetServerDate()) & "'"
            strSelectSql = strSelectSql & " order by b.Cust_Ref,b.Cust_DrgNo,b.Item_Code "
        ElseIf UCase(pstrInvType) = "EXPORT INVOICE" Then
            strSelectSql = "Select b.Item_Code,b.Cust_DrgNo,b.Amendment_No,b.Cust_Ref,b.cust_drg_desc from Cust_Ord_hdr a (NOLOCK),Cust_Ord_Dtl b (NOLOCK)"
            strSelectSql = strSelectSql & " Where a.unit_code=b.unit_code and a.unit_code='" + gstrUNITID + "' and   a.Account_Code = b.Account_Code and a.Cust_ref =b.Cust_ref "
            strSelectSql = strSelectSql & " and a.Amendment_No = b.amendment_No  AND a.Authorized_Flag = 1 and a.PO_type in ('E') "
            strSelectSql = strSelectSql & " and b.Account_Code='" & Trim(pstrCustCode) & "' and b.Active_flag ='A' "
            strSelectSql = strSelectSql & " and a.Valid_date >='" & getDateForDB(GetServerDate()) & "' and effect_Date <= '" & getDateForDB(GetServerDate()) & "'"
            strSelectSql = strSelectSql & " order by b.Cust_Ref,b.Cust_DrgNo,b.Item_Code "
            '10869290 -SERVICE INVOICE 
        ElseIf UCase(pstrInvType) = "SERVICE INVOICE" Then
            strSelectSql = "Select b.Item_Code,b.Cust_DrgNo,b.Amendment_No,b.Cust_Ref,b.cust_drg_desc from Cust_Ord_hdr a,Cust_Ord_Dtl b"
            strSelectSql = strSelectSql & " where a.unit_code=b.unit_code and a.unit_code='" & gstrUNITID & "' and b.Account_Code='" & Trim(pstrCustCode) & "' and b.Active_flag ='A' and "
            strSelectSql = strSelectSql & " a.Account_Code = b.Account_Code and a.Cust_ref =b.Cust_ref and "
            strSelectSql = strSelectSql & " a.Amendment_No = b.amendment_No AND a.Authorized_Flag = 1 and a.PO_type in ('V') "
            strSelectSql = strSelectSql & " and a.Valid_date >='" & getDateForDB(GetServerDate()) & "' and effect_date <='" & getDateForDB(GetServerDate()) & "'"
            strSelectSql = strSelectSql & " order by b.Cust_Ref,b.Amendment_No,b.Cust_DrgNo,b.Item_Code "
            '10869290 -SERVICE INVOICE 
        Else
            strSelectSql = "Select b.Item_Code,b.Cust_DrgNo,b.Amendment_No,b.Cust_Ref,b.cust_drg_desc from Cust_Ord_hdr a (NOLOCK),Cust_Ord_Dtl b (NOLOCK)"
            strSelectSql = strSelectSql & " Where a.unit_code=b.unit_code and a.unit_code='" + gstrUNITID + "' and a.Account_Code = b.Account_Code and a.Cust_ref =b.Cust_ref "
            strSelectSql = strSelectSql & " and a.Amendment_No = b.amendment_No  AND a.Authorized_Flag = 1 and a.PO_type in ('O','S') "
            strSelectSql = strSelectSql & " and b.Account_Code='" & Trim(pstrCustCode) & "' and b.Active_flag ='A' "
            strSelectSql = strSelectSql & " and a.Valid_date >='" & getDateForDB(GetServerDate()) & "' and effect_Date <= '" & getDateForDB(GetServerDate()) & "'"
            strSelectSql = strSelectSql & " order by b.Cust_Ref,b.Cust_DrgNo,b.Item_Code "
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

    Public Function SelectDataFromCustOrd_DtlUploadExcel(ByRef pstrCustCode As String, ByRef pstrInvType As String) As String
        '***********************************
        'To Get Data From Cust_Ord_Dtl
        '***********************************
        On Error GoTo ErrHandler
        Dim strSelectSql As String 'Declared To Make Select Query
        Dim rsCustOrdDtl As ClsResultSetDB
        Dim intRecordCount As Short 'To Hold Record Count
        Dim intCount As Short
        'form load funcionality
        Call AddColumnsInListViewInvoiceUpload()
        INVTYPE = pstrInvType
        Me.lvwCustRefNo.Items.Clear() 'initially clear all items in the listview
        blnInvoiceUpload = True
        If blnInvoiceUpload = True Then
            chkall.Visible = True
        Else
            chkall.Visible = False
        End If
        If UCase(pstrInvType) = "NORMAL INVOICE" Then
            strSelectSql = "Select distinct b.Cust_Ref,b.Amendment_No from Cust_Ord_hdr a (NOLOCK),Cust_Ord_Dtl b (NOLOCK)"
            strSelectSql = strSelectSql & " Where a.unit_code=b.unit_code and a.unit_code='" + gstrUNITID + "' and a.Account_Code = b.Account_Code and a.Cust_ref =b.Cust_ref "
            strSelectSql = strSelectSql & " and a.Amendment_No = b.amendment_No  AND a.Authorized_Flag = 1 and a.PO_type in ('O','S') "
            strSelectSql = strSelectSql & " and b.Account_Code='" & Trim(pstrCustCode) & "' and b.Active_flag ='A' "
            strSelectSql = strSelectSql & " and a.Valid_date >='" & getDateForDB(GetServerDate()) & "' and effect_Date <= '" & getDateForDB(GetServerDate()) & "' order by b.Cust_Ref "
            'strSelectSql = strSelectSql & " and  not exists ( Select cust_ref,Amendment_No from ( "
            'strSelectSql = strSelectSql & " SELECT cust_ref,Amendment_No from  SalesChallan_Dtl_Upload  where unit_code='" + gstrUNITID + "' "
            'strSelectSql = strSelectSql & " and Account_Code='" & Trim(pstrCustCode) & "' )tmpTbl  WHERE  tmpTbl.Cust_Ref=b.Cust_Ref  "
            'strSelectSql = strSelectSql & " and tmpTbl.Amendment_No = b.amendment_No   )  order by b.Cust_Ref "

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
               
                rsCustOrdDtl.MoveNext() 'move to next record
            Next intCount
            ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
            rsCustOrdDtl.ResultSetClose()
            rsCustOrdDtl = Nothing
        End If
        Me.ShowDialog()
        SelectDataFromCustOrd_DtlUploadExcel = mstrItemText

        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Sub frmMKTTRN0020NEW_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        'Me = Nothing
        If RSSO.State = ADODB.ObjectStateEnum.adStateOpen Then
            RSSO.Close()
            RSSO = Nothing
        End If
        Me.Dispose()
    End Sub
    Private Sub GRIDSELECT()
        For Intcounter = 0 To Me.lvwCustRefNo.Items.Count - 1
            bool_Item_Check = True
            Me.lvwCustRefNo.Items.Item(Intcounter).Checked = False
        Next Intcounter
        Me.lvwCustRefNo.Items.Item(ACTIDX).Checked = True
        bool_Item_Check = False
    End Sub
    Private Sub lvwCustRefNo_ItemChecked(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckedEventArgs) Handles lvwCustRefNo.ItemChecked
        Dim Item As System.Windows.Forms.ListViewItem = lvwCustRefNo.Items(e.Item.Index)
        On Error GoTo ErrHandler
        Dim intSubItem As Short
        Dim Index As Object
        If bool_Item_Check = True Then
            Exit Sub
        End If
        MULTIPLESO = 0
        For Intcounter = 0 To Me.lvwCustRefNo.Items.Count - 1
            If Me.lvwCustRefNo.Items.Item(Intcounter).Checked = True Then
                MULTIPLESO = MULTIPLESO + 1
                Index = Intcounter
            End If
        Next Intcounter
        If ACTIDX = 0 Or Index = 0 Then
            ACTIDX = Index
        End If
        If blnInvoiceUpload = False Then
            If RSSO.Fields("MULTIPLESO").Value = 0 And MULTIPLESO > 1 Then
                MsgBox("More than one SO selection not available", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Empower")
                Call GRIDSELECT()
                Exit Sub
            End If
        End If
        
        Select Case INVTYPE
            Case "NORMAL INVOICE"
            Case Else
                If MULTIPLESO > 1 Then
                    MsgBox("More than one SO selection not available for this case", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Empower")
                    Call GRIDSELECT()
                End If
        End Select
        mstrItemText = Trim(Item.Text)
        mstrItemText = "'" & Trim(Item.Text) & "','" & Trim(Item.SubItems(1).Text) & "'"
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
        '---------------------------------------------------------------------
        'Created By     -   Shruti Khanna\(Name Changed - Nitin Sood)
        '---------------------------------------------------------------------
        Dim itmFound As System.Windows.Forms.ListViewItem ' FoundItem variable.
        On Error GoTo ErrHandler
        Dim intCount As Integer = 0
        For intCount = 0 To Me.lvwCustRefNo.Items.Count - 1
            lvwCustRefNo.Items.Item(intCount).Font = VB6.FontChangeBold(lvwCustRefNo.Items.Item(intCount).Font, False)
        Next
        With lvwCustRefNo
            If optRefrence.Checked = True Then
                If Len(txtSearch.Text) = 0 Then Exit Sub
                For Intcounter = 0 To .Items.Count - 1
                    If Trim(UCase(Mid(.Items.Item(Intcounter).Text, 1, Len(txtSearch.Text)))) = Trim(UCase(txtSearch.Text)) Then
                        .Items.Item(Intcounter).Font = VB6.FontChangeBold(.Items.Item(Intcounter).Font, True)
                        Call .Items.Item(Intcounter).EnsureVisible()
                        .Refresh()
                        Exit For
                    End If
                Next
            ElseIf optdrgNO.Checked = True Then
                If Len(txtSearch.Text) = 0 Then Exit Sub
                For Intcounter = 0 To .Items.Count - 1
                    If Trim(UCase(Mid(.Items.Item(Intcounter).SubItems.Item(2).Text, 1, Len(txtSearch.Text)))) = Trim(UCase(txtSearch.Text)) Then
                        .Items.Item(Intcounter).Font = VB6.FontChangeBold(.Items.Item(Intcounter).Font, True)
                        Call .Items.Item(Intcounter).EnsureVisible()
                        .Refresh()
                        Exit For
                    End If
                Next
            ElseIf optItem.Checked = True Then
                If Len(txtSearch.Text) = 0 Then Exit Sub
                For Intcounter = 0 To .Items.Count - 1
                    If Trim(UCase(Mid(.Items.Item(Intcounter).SubItems.Item(3).Text, 1, Len(txtSearch.Text)))) = Trim(UCase(txtSearch.Text)) Then
                        .Items.Item(Intcounter).Font = VB6.FontChangeBold(.Items.Item(Intcounter).Font, True)
                        Call .Items.Item(Intcounter).EnsureVisible()
                        .Refresh()
                        Exit For
                    End If
                Next
            ElseIf OptItemCode.Checked = True Then
                If Len(txtSearch.Text) = 0 Then Exit Sub
                For Intcounter = 0 To .Items.Count - 1
                    If Trim(UCase(Mid(.Items.Item(Intcounter).SubItems.Item(4).Text, 1, Len(txtSearch.Text)))) = Trim(UCase(txtSearch.Text)) Then
                        .Items.Item(Intcounter).Font = VB6.FontChangeBold(.Items.Item(Intcounter).Font, True)
                        Call .Items.Item(Intcounter).EnsureVisible()
                        .Refresh()
                        Exit For
                    End If
                Next
            End If
        End With
        If itmFound Is Nothing Then ' If no match,
            Exit Sub
        Else
            itmFound.EnsureVisible() ' Scroll ListView to show found ListItem.
            itmFound.Selected = True ' Select the ListItem.
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
        Dim strSQL As String
        Dim intLoopCounter As Short
        Dim intmaxLoop As Short
        On Error GoTo ErrHandler
        rsGrnDtl = New ClsResultSetDB
        'form load functionality
        Call AddColumnsInListView()

        'Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(mdifrmMain.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2)
        'Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(mdifrmMain.Width) - VB6.PixelsToTwipsX(frmModules.Width)) / 2.8)
        optRefrence.Checked = True
        RSSO = mP_Connection.Execute("select MultipleSO from Sales_Parameter")
        strSQL = "select a.Doc_No,a.Item_code,a.Rejected_Quantity from grn_Dtl a (NOLOCK),grn_hdr b (NOLOCK) Where "
        strSQL = strSQL & " a.unit_code=b.unit_code and a.Doc_type = b.Doc_type And a.Doc_No = b.Doc_No and "
        strSQL = strSQL & " a.From_Location = b.From_Location and a.From_Location ='01R1' and a.unit_code='" & gstrUNITID & "'"
        strSQL = strSQL & " and a.Rejected_quantity > 0  and  b.Vendor_code = '" & pstrVendCode & "' and isnull(b.GRN_Cancelled,0) = 0 order by a.Doc_No"
        rsGrnDtl.GetResult(strSQL)
        If rsGrnDtl.GetNoRows > 0 Then
            intmaxLoop = rsGrnDtl.GetNoRows : rsGrnDtl.MoveFirst()
            ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.WaitCursor)
            rsGrnDtl.MoveFirst() 'move to first record
            For intLoopCounter = 0 To intmaxLoop - 1
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

    Private Sub chkall_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkall.CheckedChanged
        Dim intCount As Integer = 0
        If chkall.Checked Then
            For intCount = 0 To Me.lvwCustRefNo.Items.Count - 1
                lvwCustRefNo.Items.Item(intCount).Checked = True
            Next intCount
        Else
            For intCount = 0 To Me.lvwCustRefNo.Items.Count - 1
                lvwCustRefNo.Items.Item(intCount).Checked = False
            Next intCount
        End If
    End Sub
End Class