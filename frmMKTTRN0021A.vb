Option Strict Off
Option Explicit On
Friend Class frmMKTTRN0021A
	Inherits System.Windows.Forms.Form
	'===================================================================================
	' (c) MothersonSumi INfotech & Design Ltd. All rights reserverd.
	' File Name         :   FRMMKTTRN0021A.frm
	' Function          :   Used to select items scaneed against Bar code
	' Created By        :   Sourabh
	' Created On        :   22 Nov 2005
	'=======================================================================================
	'Revised By         : Davinder Singh
	'Revision Date      : 28-06-2006
	'Issue Id           : 18103
	'Revision History   : To Check the more than 7 Items Invoice according to the flag in the sales parameter
	'                     table for all type of invoices and make the user able to make the Invoice of more
	'                     than 7 items if that flag is TRUE
	'=======================================================================================
    'Revised By        -    Vinod Singh
    'Revision Date     -    25/05/2011
    'Revision History  -    Changes for Multi Unit
    '=======================================================================================
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
	Private Sub CmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdCancel.Click
		On Error GoTo ErrHandler
		Me.Close()
		Exit Sub
ErrHandler: 'The Error Handling Code Starts here
		Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
	End Sub
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		'=======================================================================================
		'Code Modified By   -   Nitin Sood
		'No of Items Selected in Challan can be Till 7
		'=======================================================================================
		'Revised By         : Davinder Singh
		'Revision Date      : 28-06-2006
		'Issue Id           : 18103
		'Revision History   : To Check the more than 7 Items Invoice according to the flag in the sales parameter
		'                     table for all type of invoices and make the user able to make the Invoice of more
		'                     than 7 items if that flag is TRUE
		'=======================================================================================
        On Error GoTo ErrHandler
		mstrItemText = "" : intCheckCounter = intIteminSp
		Dim intSubItem As Short
		Dim blnMoreThan7ItemsAllowed As Boolean
		
        gobjDB.GetResult("Select MoreThan7ItemInInvoice from sales_parameter where Unit_code='" & gstrUNITID & "'")
        If gobjDB.GetValue("MoreThan7ItemInInvoice") = "True" Then
            blnMoreThan7ItemsAllowed = True
        End If
        For intSubItem = 0 To lvwItemCode.Items.Count - 1
            If lvwItemCode.Items.Item(intSubItem).Checked = True Then
                intCheckCounter = intCheckCounter + 1
                If (Not blnMoreThan7ItemsAllowed And intCheckCounter > 7) Then
                    MsgBox("No. Of Items Selected Should be Less than 7", MsgBoxStyle.Information, ResolveResString(100))
                    mstrItemText = ""
                    Exit Sub
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
    Private Sub frmMKTTRN0021A_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo ErrHandler
        SetBackGroundColorNew(Me, True)
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
        '***********************************
        On Error GoTo ErrHandler
        With Me.lvwItemCode
            mCtlHdrItemCode = .Columns.Add("")
            mCtlHdrItemCode.Text = "Dispatch Slip No"
            mCtlHdrItemCode.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(lvwItemCode.Width) / 4)
            mCtlHdrDrawingNo = .Columns.Add("")
            mCtlHdrDrawingNo.Text = "Drawing No."
            mCtlHdrDrawingNo.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(lvwItemCode.Width) / 4)
            mCtlHdrDescription = .Columns.Add("")
            mCtlHdrDescription.Text = "Description"
            mCtlHdrDescription.Width = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(lvwItemCode.Width) / 4))
            mCtlHdrDescription = .Columns.Add("")
            mCtlHdrDescription.Text = "Quantity"
            mCtlHdrDescription.Width = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(lvwItemCode.Width) / 4) - 100)
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
        strDate = VB6.Format(GetServerDate, gstrDateFormat)
        Me.lvwItemCode.Items.Clear() 'initially clear all items in the listview
        strSelectSql = "Select effectMon=convert(char(2),month(effect_date)),effectYr=convert(char(4),Year(effect_date)),"
        strSelectSql = strSelectSql & " validMon=convert(char(2),month(Valid_date)),validYr=convert(char(4),year(Valid_date))"
        strSelectSql = strSelectSql & " from Cust_Ord_hdr where Unit_code='" & gstrUNITID & "' and"
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
            Select Case UCase(pstrInvType)
                Case "NORMAL INVOICE"
                    Select Case UCase(pstrSubType)
                        Case "FINISHED GOODS"
                            strSelectSql = makeSelectSql(pstrCustno, pstrRefNo, pstrAmmNo, effectyrmon, Validyrmon, pstrstockLocation, strDate, "'F','S'", pstrCondition)
                    End Select
            End Select
        End If
        rsCustOrdHdr.ResultSetClose()
        rsCustOrdDtl = New ClsResultSetDB
        rsCustOrdDtl.GetResult(strSelectSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        intRecordCount = rsCustOrdDtl.GetNoRows 'assign record count to integer variable
        If intRecordCount > 0 Then '          'if record found
            rsCustOrdDtl.MoveFirst() 'move to first record
            For intCount = 1 To intRecordCount
                mListItemUserId = Me.lvwItemCode.Items.Add(rsCustOrdDtl.GetValue("Doc_No"))
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
                    mListItemUserId.SubItems(3).Text = rsCustOrdDtl.GetValue("Quantity")
                Else
                    mListItemUserId.SubItems.Insert(3, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsCustOrdDtl.GetValue("Quantity")))
                End If
                rsCustOrdDtl.MoveNext() 'move to next record
            Next intCount
        Else
            MsgBox("No Items for selected Invoice in Sales Order.Please Check Following :" & vbCrLf & "1. Item in SO are Active and Not on Hold." & vbCrLf & "2. Check Balance of Items for location " & pstrstockLocation & "." & vbCrLf & "3. Check Marketing Schedule in Case of Finished\Trading Goods in SO.", MsgBoxStyle.Information, "empower")
            Exit Function
        End If
        rsCustOrdDtl.ResultSetClose()
        rsCustOrdDtl = Nothing
        Me.ShowDialog()
        SelectDataFromCustOrd_Dtl = mstrItemText
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Sub frmMKTTRN0021A_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
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
    Public Function SelectDatafromsaleDtl(ByRef pstrchallanNo As Object) As Object
        On Error GoTo ErrHandler
        Dim strsaledtl As String
        Dim strInvType As String
        Dim rssaledtl As ClsResultSetDB
        Dim rsInvType As ClsResultSetDB
        Dim intRecordCount As Short 'To Hold Record Count
        Dim intCount As Short
        strInvType = "select a.description,a.Sub_type_Description from saleconf a,saleschallan_dtl b where a.unit_code = b.unit_code and a.Invoice_type =b.Invoice_Type and b.Doc_no = " & Val(pstrchallanNo) & " and datediff(dd,b.Invoice_Date,a.fin_start_date)<=0  and datediff(dd,a.fin_end_date,b.Invoice_Date)<=0 and a.Unit_code='" & gstrUNITID & "'"
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
        strsaledtl = ""
        strsaledtl = "Select a.Item_Code,a.Cust_ITem_Code,a.Cust_Item_Desc,b.Tariff_Code from Sales_Dtl a,Item_Mst b where a.Unit_code = B.Unit_code and a.ITem_code = b.ITem_code and a.Unit_code='" & gstrUNITID & "' and Doc_No ="
        strsaledtl = strsaledtl & pstrchallanNo
        rssaledtl = New ClsResultSetDB
        rssaledtl.GetResult(strsaledtl, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        intRecordCount = rssaledtl.GetNoRows 'assign record count to integer variable
        If intRecordCount > 0 Then '          'if record found
            rssaledtl.MoveFirst() 'move to first record
            For intCount = 1 To intRecordCount
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
        End If
        rssaledtl.ResultSetClose()
        Me.ShowDialog()
        SelectDatafromsaleDtl = mstrItemText
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Public Function makeSelectSql(ByRef pstrCustno As String, ByRef pstrRefNo As String, ByRef pstrAmmNo As String, ByRef effectyrmon As String, ByRef Validyrmon As String, ByRef pstrstockLocation As String, ByRef strDate As String, ByRef pstrItemin As String, Optional ByRef pstrCondition As String = "") As String
        Dim strSelectSql As String
        strSelectSql = " Select E.Doc_No,E.Cust_Drgno,C.Cust_Drg_Desc,E.Quantity " & " from Cust_Ord_hdr a inner join Cust_ord_dtl c on a.Unit_code = c.Unit_code and a.Cust_ref = c.Cust_ref and a.amendment_No = c.amendment_No and a.Account_code=c.account_code And c.Active_Flag ='A' inner join Item_mst d on d.Unit_code = C.Unit_code and d.item_Code = c.item_Code inner join vw_barcode_Dispatch e on  a.Unit_code=e.unit_code and a.Account_Code = e.customer_Code And c.cust_drgno = e.cust_drgno And e.Item_Code = d.Item_Code  inner join MonthlyMktSchedule b on a.Unit_code=b.Unit_code and a.account_code=b.Account_code and c.Cust_drgNo = b.Cust_drgNo and b.ITem_code = d.Item_code inner join itembal_mst f on b.Unit_code=f.Unit_code and b.item_Code = f.item_Code and f.Location_code ='" & pstrstockLocation & "' and f.Cur_bal >0 " & " where  a.Unit_code='" & gstrUNITID & "' and a.Account_Code='" & Trim(pstrCustno) & "'" & " and a.Cust_Ref='" & Trim(pstrRefNo) & "' and a.Amendment_No='" & Trim(pstrAmmNo) & "' and b.status = 1" & " and b.Schedule_flag =1 and b.Year_Month = " & Validyrmon & " and d.hold_flag =0 and D.Status = 'A' "
        If Len(Trim(pstrCondition)) > 0 Then
            strSelectSql = strSelectSql & " and c.Item_code not in(" & pstrCondition & ")"
        Else
            strSelectSql = strSelectSql & " "
        End If
        strSelectSql = strSelectSql & " Union "
        strSelectSql = strSelectSql & " Select E.Doc_No,E.Cust_Drgno,C.Cust_Drg_Desc,E.Quantity"
        strSelectSql = strSelectSql & " from Cust_Ord_hdr a inner join Cust_ord_dtl c on"
        strSelectSql = strSelectSql & " a.Unit_code=c.unit_code and a.Cust_ref = c.Cust_ref and a.amendment_No = c.amendment_No and a.Account_code=c.account_code And c.Active_Flag ='A'"
        strSelectSql = strSelectSql & " inner join Item_mst d on d.Unit_code = c.Unit_code and d.item_Code = c.item_Code"
        strSelectSql = strSelectSql & " inner join vw_barcode_Dispatch e on"
        strSelectSql = strSelectSql & " a.Unit_code=e.Unit_code and a.Account_Code = e.customer_Code And c.cust_drgno = e.cust_drgno And e.Item_Code = d.Item_Code"
        strSelectSql = strSelectSql & " inner join DailyMktSchedule b on"
        strSelectSql = strSelectSql & " a.Unit_code=b.Unit_code and a.account_code=b.Account_code and c.Cust_drgNo=b.Cust_drgNo and b.ITem_code = d.Item_code"
        strSelectSql = strSelectSql & " inner join itembal_mst f on b.Unit_code = f.Unit_code and b.item_Code = f.item_Code  and f.Location_code ='" & pstrstockLocation & "' and f.Cur_bal >0"
        strSelectSql = strSelectSql & " where a.Unit_code='" & gstrUNITID & "' and a.Account_Code='" & Trim(pstrCustno) & "'"
        strSelectSql = strSelectSql & " and a.Cust_Ref='" & Trim(pstrRefNo) & "' and a.Amendment_No='" & Trim(pstrAmmNo) & "'"
        strSelectSql = strSelectSql & " and  datepart(mm,b.trans_date) = " & Month(ConvertToDate(strDate)) 'Mid(strDate, 4, 2)
        strSelectSql = strSelectSql & " and  datepart(dd,b.trans_date) <=" & ConvertToDate(strDate).Day & " and  datepart(yyyy,b.trans_date) = " & Year(ConvertToDate(strDate))  'Mid(strDate, 7, 4)
        strSelectSql = strSelectSql & " and d.hold_flag =0 and D.Status = 'A'"
        If Len(Trim(pstrCondition)) > 0 Then
            strSelectSql = strSelectSql & " and c.Item_code not in(" & pstrCondition & " )"
        Else
            strSelectSql = strSelectSql & " "
        End If
        makeSelectSql = strSelectSql
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
    Private Sub lvwItemCode_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lvwItemCode.SelectedIndexChanged
    End Sub
End Class