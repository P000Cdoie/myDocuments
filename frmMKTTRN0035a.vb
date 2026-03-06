Option Strict Off
Option Explicit On
Friend Class frmMKTTRN0035a
	Inherits System.Windows.Forms.Form
	'---------------------------------------------------------------------------
	'Copyright          :   MIND Ltd.
	'Form Name          :   frmMKTTRN0035a.frm
	'Created By         :   Arshad Ali
	'Created on         :   26/07/2005
	'Modified Date      :
	'Description        :   This form is used to show items to select from.
	'=======================================================================================
	'Revision  By       : Ashutosh Verma,issue id:16685
	'Revision On        : 26-12-2005
	'History            : FOC parts are not allowed in Multiple Component Invoice (for SUNVAC).
	'=======================================================================================
	'Revision  By       : Ashutosh Verma
	'Revision History   : on 06-01-2006,issue id:16773 ,Sort items from grid according to Item Code & Cust Drawing No.
	'=======================================================================================
	'Revision By        : Debasish Pradhan
	'Revision History   : on 14-03-2006, Will consider without locked invoice as saled quantity. Previously, locked invoices
	'                      are only considered to be final one. Now on, without locked invoice will be also considered.
    '                      Without Locked invoices can be edited/deleted.
    'Modified by Sameer Srivastava on 2011-May-25
    '   Modified to support MultiUnit functionality
	'=======================================================================================
    Dim mstrInvType As String
	Dim mstrInvSubType As String
	Enum GridHeader
		Mark = 1
		KanbanNo = 2
		ItemCode = 3
		DrawingNo = 4
		Description = 5
		Quantity = 6
		CustRef = 7
		AmendmentNo = 8
		SchDate = 9
		SChTime = 10
		UNLoc = 11
		USLOC = 12
		AccountCode = 13
        Tool_Cost = 14
    End Enum
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		On Error GoTo ErrHandler
        Me.Dispose()
		Exit Sub
ErrHandler: 'The Error Handling Code Starts here
		Call gobjError.RaiseError(Err.Number, err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
		'Created By   -   Arshad Ali
		'retrieve item code of all selected items
		'*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
		On Error GoTo ErrHandler
		mstrItemText = ""
		Dim intSubItem As Short
		Dim strKanabanNo As String
		Dim StrItemCode As String
		Dim strDrawingNo As String
		Dim strDescription As String
		Dim dblQuantity As Double
		Dim strCustRef As String
		Dim StrAmendmentNo As String
		Dim strSchDate As String
		Dim strSchTime As String
		Dim strunLoc As String
		Dim strUSLoc As String
		Dim strAccountCode As String
		
		If Not ValidateData Then Exit Sub
		
		With SpItems
			For intSubItem = 1 To .maxRows
				.Row = intSubItem
				.Col = GridHeader.Mark
				If CBool(.value) = True Then
					.Col = GridHeader.KanbanNo
					strKanabanNo = Trim(.Text)
					.Col = GridHeader.ItemCode
					StrItemCode = Trim(.Text)
					.Col = GridHeader.Quantity
					dblQuantity = CDbl(Trim(.Text))
					.Col = GridHeader.CustRef
					strCustRef = Trim(.Text)
					.Col = GridHeader.AmendmentNo
					StrAmendmentNo = Trim(.Text)
					.Col = GridHeader.SchDate
					strSchDate = Trim(.Text)
					.Col = GridHeader.SChTime
					strSchTime = Trim(.Text)
					.Col = GridHeader.UNLoc
					strunLoc = Trim(.Text)
					.Col = GridHeader.USLOC
					strUSLoc = Trim(.Text)
					
					.Col = GridHeader.AccountCode
					strAccountCode = Trim(.Text)
					
					.Col = GridHeader.DrawingNo
					strDrawingNo = Trim(.Text)
					mstrItemText = mstrItemText & strKanabanNo & "|" & StrItemCode & "|" & strDrawingNo & "|" & strDescription & "|" & dblQuantity & "|" & strCustRef & "|" & StrAmendmentNo & "|" & strSchDate & "|" & strSchTime & "|" & strunLoc & "|" & strUSLoc & "|" & strAccountCode & "^"
				End If
			Next intSubItem
		End With
		If Len(mstrItemText) = 0 Then
			Call ConfirmWindow(10418, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
			Me.SpItems.Focus()
			Exit Sub
		End If
        Me.Dispose()
		Exit Sub
ErrHandler: 'The Error Handling Code Starts here
		Call gobjError.RaiseError(Err.Number, err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
	End Sub
	
	
    Private Sub frmMKTTRN0035a_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        'Change cursor to arrow when over headers
        SpItems.CursorType = FPSpreadADO.CursorTypeConstants.CursorTypeColHeader
        SpItems.CursorStyle = FPSpreadADO.CursorStyleConstants.CursorStyleArrow
        SpItems.Focus()
    End Sub
	Private Sub frmMKTTRN0035a_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo ErrHandler
        SetBackGroundColorNew(Me, True)
        Call AddColumnsInSpread()

        'Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(mdifrmMain.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2)
        'Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(mdifrmMain.Width) - VB6.PixelsToTwipsX(frmModules.Width)) / 4.4)
		optKanbanNo.Checked = True
		mstrItemText = ""
        SelectDatafromItem_Mst()
		Exit Sub
ErrHandler: 'The Error Handling Code Starts here
		Call gobjError.RaiseError(Err.Number, err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
	Sub AddColumnsInSpread()
		With SpItems
			.maxRows = 0
            .MaxCols = 14
			
			.Row = 0
			.Col = GridHeader.Mark : .Text = "Mark" : .set_ColWidth(GridHeader.Mark, 4)
			.Col = GridHeader.KanbanNo : .Text = "Kanban No" : .set_ColWidth(GridHeader.KanbanNo, 10)
			.Col = GridHeader.ItemCode : .Text = "Item Code" : .set_ColWidth(GridHeader.ItemCode, 10)
			.Col = GridHeader.DrawingNo : .Text = "Drawing No" : .set_ColWidth(GridHeader.DrawingNo, 12)
			.Col = GridHeader.Description : .Text = "Description" : .set_ColWidth(GridHeader.Description, 16)
			.Col = GridHeader.Quantity : .Text = "Quantity" : .set_ColWidth(GridHeader.Quantity, 6)
			.Col = GridHeader.CustRef : .Text = "Cust Ref" : .set_ColWidth(GridHeader.CustRef, 10)
			.Col = GridHeader.AmendmentNo : .Text = "Amendment No." : .set_ColWidth(GridHeader.AmendmentNo, 8)
			.Col = GridHeader.SchDate : .Text = "Sch Date" : .set_ColWidth(GridHeader.SchDate, 9)
			.Col = GridHeader.SChTime : .Text = "Sch Time" : .set_ColWidth(GridHeader.SChTime, 6)
			.Col = GridHeader.UNLoc : .Text = "UNLoc" : .set_ColWidth(GridHeader.UNLoc, 6)
			.Col = GridHeader.USLOC : .Text = "USLoc" : .set_ColWidth(GridHeader.USLOC, 6)
			.Col = GridHeader.AccountCode : .Text = "Account Code" : .set_ColWidth(GridHeader.AccountCode, 12)
			.Col = GridHeader.Tool_Cost : .Text = "Tool Cost" : .set_ColWidth(GridHeader.Tool_Cost, 0)
        End With
    End Sub
	Public Function SelectDatafromItem_Mst(Optional ByRef pstrItem As String = "", Optional ByRef intAlreadyItem As Integer = 0) As Object
		On Error GoTo ErrHandler
		Dim strItembal As String
		Dim rsItembal As ClsResultSetDB
		
		Dim intRecordCount As Integer 'To Hold Record Count
		Dim intCount As Short
		strItembal = "select Distinct KanbanNo, m.item_code, m.cust_drgNo, m.Description, m.Quantity, m.Cust_Ref, m.amendment_no, m.Sch_Date,"
		
		strItembal = strItembal & " case when Sch_Time = '23:59' then '' else Sch_Time end as sch_time, m.UNLOC, m.USLOC, m.Account_code ,m.Tool_cost"
		
		strItembal = strItembal & " from vw_Enagaredtl_Help m"
        strItembal = strItembal & " where m.UNIT_CODE = '" & gstrUNITID & "' AND m.quantity > ((select isnull(sum(b.sales_quantity),0) from salesChallan_dtl a inner join sales_dtl b on a.unit_code = b.unit_code and a.unit_code = '" & gstrUNITID & "' and a.location_code = b.location_code and a.doc_no=b.doc_no where m.kanbanNo = b.srvdino and a.cancel_flag <> 1) + (select IsNull(sum(sales_quantity),0) as sales_quantity  from printedsrv_dtl p where p.unit_code = '" & gstrUNITID & "' and p.KanBan_No=m.KanBanNo)+(Select isnull(Sum(quantity),0) as sales_quantity From mkt_57F4challankanban_dtl B inner join mkt_57F4challan_hdr A on B.unit_code = A.unit_code and B.unit_code = '" & gstrUNITID & "' and B.doc_type=A.doc_type and B.doc_no = A.doc_no where A.cancel_flag = 0 and B.Kanban_no=m.KanBanNo))"
		strItembal = strItembal & " order by kanbanNo"
		
		
		rsItembal = New ClsResultSetDB
		If Len(Trim(strItembal)) <= 0 Then Exit Function
		mP_Connection.CommandTimeout = 0
		rsItembal.GetResult(strItembal, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
		mP_Connection.CommandTimeout = 30
		intRecordCount = rsItembal.GetNoRows 'assign record count to integer variable
		If intRecordCount > 0 Then '          'if record found
			rsItembal.MoveFirst() 'move to first record
			For intCount = 1 To intRecordCount
				With SpItems
					.maxRows = .maxRows + 1
					.Row = intCount
					.Col = GridHeader.Mark : .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox : .TypeCheckCenter = True
					.Col = GridHeader.KanbanNo
                    .Text = rsItembal.GetValue("KanbanNo") : .Lock = True
					.Col = GridHeader.ItemCode
                    .Text = rsItembal.GetValue("item_code") : .Lock = True
					.Col = GridHeader.DrawingNo
                    .Text = rsItembal.GetValue("cust_drgNo") : .Lock = True
					.Col = GridHeader.Description
                    .Text = rsItembal.GetValue("Description") : .Lock = True
					
					.Col = GridHeader.Quantity
                    .Text = rsItembal.GetValue("Quantity") : .Lock = True
					.Col = GridHeader.CustRef
                    .Text = rsItembal.GetValue("Cust_Ref") : .Lock = True
					.Col = GridHeader.AmendmentNo
                    .Text = rsItembal.GetValue("amendment_no") : .Lock = True
					.Col = GridHeader.SchDate
                    .Text = rsItembal.GetValue("Sch_Date") : .Lock = True
                    .Col = GridHeader.SChTime
                    .Text = rsItembal.GetValue("Sch_Time") : .Lock = True
                    .Col = GridHeader.UNLoc
                    .Text = rsItembal.GetValue("UNLOC") : .Lock = True
                    .Col = GridHeader.USLOC
                    .Text = rsItembal.GetValue("USLOC") : .Lock = True
                    .Col = GridHeader.AccountCode
                    .Text = rsItembal.GetValue("Account_Code") : .Lock = True
                    .Col = GridHeader.Tool_Cost
                    .Text = rsItembal.GetValue("Tool_Cost") : .Lock = True

                End With
                rsItembal.MoveNext() 'move to next record
            Next intCount
            rsItembal.ResultSetClose()
            rsItembal = Nothing
        Else
            MsgBox("No Records Found", MsgBoxStyle.Information, "eMPro")
            Exit Function
        End If
        If Len(pstrItem) > 0 Then Call SelectPreviousItem(pstrItem)
        SelectDatafromItem_Mst = mstrItemText
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Sub optCustDrawNo_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optCustDrawNo.CheckedChanged
        If eventSender.Checked Then
            '=======================================================================================
            'Revision  By       : Ashutosh Verma
            'Revision History   : on 06-01-2006,issue id:16773 ,Sort items from grid according to Item Code & Cust Drawing No.
            '=======================================================================================
            If optCustDrawNo.Checked = True Then
                With SpItems
                    .SortBy = FPSpreadADO.SortByConstants.SortByRow
                    .set_SortKey(1, GridHeader.DrawingNo)
                    .set_SortKeyOrder(1, FPSpreadADO.SortKeyOrderConstants.SortKeyOrderAscending)
                    .Col = 1
                    .Col2 = .MaxCols
                    .Row = 0
                    .Row2 = .MaxRows
                    .Action = FPSpreadADO.ActionConstants.ActionSort
                End With
            End If
            txtsearch.Text = ""
            txtsearch.Focus()
        End If
    End Sub
    Private Sub optCustRef_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optCustRef.CheckedChanged
        If eventSender.Checked Then
            txtsearch.Text = ""
            txtsearch.Focus()
        End If
    End Sub
    Private Sub optDescription_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optDescription.CheckedChanged
        If eventSender.Checked Then
            txtsearch.Text = ""
            txtsearch.Focus()
        End If
    End Sub
    Private Sub OptItemCode_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optItemCode.CheckedChanged
        If eventSender.Checked Then
            '=======================================================================================
            'Revision  By       : Ashutosh Verma
            'Revision History   : on 06-01-2006,issue id:16773 ,Sort items from grid according to Item Code & Cust Drawing No.
            '=======================================================================================
            If optItemCode.Checked = True Then
                With SpItems
                    .SortBy = FPSpreadADO.SortByConstants.SortByRow
                    .set_SortKey(1, GridHeader.ItemCode)
                    .set_SortKeyOrder(1, FPSpreadADO.SortKeyOrderConstants.SortKeyOrderAscending)
                    .Col = 1
                    .Col2 = .MaxCols
                    .Row = 0
                    .Row2 = .MaxRows
                    .Action = FPSpreadADO.ActionConstants.ActionSort
                End With
            End If
            txtsearch.Text = ""
            txtsearch.Focus()
        End If
    End Sub
    Private Sub optKanbanNo_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optKanbanNo.CheckedChanged
        If eventSender.Checked Then
            txtsearch.Text = ""
        End If
    End Sub
    Private Sub optSchDate_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optSchDate.CheckedChanged
        If eventSender.Checked Then
            txtsearch.Text = ""
            txtsearch.Focus()
        End If
    End Sub
    Private Sub SpItems_Change(ByVal eventSender As System.Object, ByVal eventArgs As AxFPSpreadADO._DSpreadEvents_ChangeEvent) Handles SpItems.Change
        Call SpItems_ClickEvent(SpItems, New AxFPSpreadADO._DSpreadEvents_ClickEvent(GridHeader.Mark, eventArgs.row))
    End Sub
    Private Sub SpItems_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxFPSpreadADO._DSpreadEvents_ClickEvent) Handles SpItems.ClickEvent
        '=======================================================================================
        'Revision  By       : Ashutosh Verma,issue id:16685
        'Revision On        : 26-12-2005
        'History            : FOC parts are not allowed in Multiple Component Invoice (for SUNVAC).
        '=======================================================================================
        Dim varItemCode As Object
        Dim varKanbanNo As Object
        Dim blnflag As Boolean
        Dim blnflag1 As Boolean
        Dim intSubItem As Short
        Dim varToolCost As Object
        Dim intToolCost As Short
        Dim intOtherItemToolCost As Short
        Dim intSubItem1 As Short
        Dim blnOtherRecord As Boolean
        On Error GoTo ErrHandler

        With SpItems
            If eventArgs.col = GridHeader.Mark Then
                .Row = eventArgs.row : .Col = eventArgs.col
                If CBool(.Value) = False Then Exit Sub
                varItemCode = Nothing
                blnflag = .GetText(GridHeader.ItemCode, eventArgs.row, varItemCode)
                For intSubItem = 1 To .MaxRows
                    .Row = intSubItem
                    .Col = GridHeader.Mark
                    ''' If Item is Checked
                    If CBool(.Value) = True And .Row <> eventArgs.row Then
                        .Col = GridHeader.ItemCode
                        ''' If Item is repeated.
                        If UCase(Trim(varItemCode)) = UCase(Trim(.Text)) Then
                            ''ValidateData
                            .Col = GridHeader.KanbanNo
                            varKanbanNo = Nothing
                            blnflag1 = .GetText(GridHeader.KanbanNo, eventArgs.row, varKanbanNo)
                            If UCase(Trim(varKanbanNo)) <> UCase(Trim(.Text)) Then
                                MsgBox("You can not select same item from different Kanban No.", MsgBoxStyle.Information, "eMPro")
                                .Col = GridHeader.Mark
                                .Row = eventArgs.row
                                .Value = CStr(System.Windows.Forms.CheckState.Unchecked)
                                System.Windows.Forms.Application.DoEvents()
                                Me.txtsearch.Focus()
                            Else
                                MsgBox("You can not select same item from different Sales Orders. ", MsgBoxStyle.Information, "eMPro")
                                .Col = GridHeader.Mark
                                .Row = eventArgs.row
                                .Value = CStr(System.Windows.Forms.CheckState.Unchecked)
                                System.Windows.Forms.Application.DoEvents()
                                Me.txtsearch.Focus()
                            End If
                            Exit Sub
                        End If
                    End If
                Next
            End If
        End With
        blnOtherRecord = False
        With SpItems
            For intSubItem1 = 1 To .MaxRows
                .Row = intSubItem1
                .Col = GridHeader.Mark
                If CBool(.Value) = True Then
                    If Not blnOtherRecord Then
                        .Col = GridHeader.Tool_Cost
                        intToolCost = CShort(Trim(.Text))
                        blnOtherRecord = True
                    Else
                        .Col = GridHeader.Tool_Cost
                        intOtherItemToolCost = CShort(Trim(.Text))
                        If intToolCost = 0 Then
                            If intOtherItemToolCost <> 0 Then
                                MsgBox("FOC parts are NOT Allowed in Multiple Component Invoicing.", MsgBoxStyle.Information, "eMPro")
                                .Col = GridHeader.Mark
                                .Row = eventArgs.row
                                .Value = CStr(System.Windows.Forms.CheckState.Unchecked)
                                System.Windows.Forms.Application.DoEvents()
                                Me.txtsearch.Focus()
                                Exit Sub
                            End If
                        End If
                        If intToolCost <> 0 Then
                            MsgBox("FOC parts are NOT Allowed in Multiple Component Invoicing.", MsgBoxStyle.Information, "eMPro")
                            .Col = GridHeader.Mark
                            .Row = eventArgs.row
                            .Value = CStr(System.Windows.Forms.CheckState.Unchecked)
                            System.Windows.Forms.Application.DoEvents()
                            Me.txtsearch.Focus()
                            Exit Sub
                        End If
                    End If
                End If
            Next intSubItem1
        End With
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub SpItems_KeyUpEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxFPSpreadADO._DSpreadEvents_KeyUpEvent) Handles SpItems.KeyUpEvent
        With SpItems
            If eventArgs.keyCode = 13 Or eventArgs.keyCode = System.Windows.Forms.Keys.Space Then
                .Col = 1
                .Value = IIf(Val(.Value), False, True)
                Call SpItems_ClickEvent(SpItems, New AxFPSpreadADO._DSpreadEvents_ClickEvent(GridHeader.Mark, .Row))
            End If
        End With
    End Sub
    Private Sub SpItems_LeaveCell(ByVal eventSender As System.Object, ByVal eventArgs As AxFPSpreadADO._DSpreadEvents_LeaveCellEvent) Handles SpItems.LeaveCell
        With SpItems
            .Row = -1
            .Col = -1
            .BackColor = System.Drawing.Color.White
            .ForeColor = System.Drawing.Color.Black
            .Col = -1
            .Row = IIf(eventArgs.newRow <= 0, 1, eventArgs.newRow)
            .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000D)
            .ForeColor = System.Drawing.Color.White
        End With
    End Sub
    Private Sub TxtSearch_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtsearch.TextChanged
        Call SearchItem()
    End Sub
    Sub SearchItem()
        On Error GoTo ErrHandler
        Dim intCount As Short
        With SpItems
            .Row = -1
            .Col = -1
            .BackColor = System.Drawing.Color.White
            .ForeColor = System.Drawing.Color.Black
            If optKanbanNo.Checked Then
                .Col = 2
            End If
            If optItemCode.Checked Then
                .Col = 3
            End If
            If optDescription.Checked Then
                .Col = 5
            End If
            If optCustRef.Checked Then
                .Col = 7
            End If
            If optSchDate.Checked Then
                .Col = 9
            End If
            If optCustDrawNo.Checked Then
                .Col = 4
            End If
            For intCount = 1 To .MaxRows
                .Row = intCount
                If UCase(Mid(.Text, 1, Len(txtsearch.Text))) = UCase(txtsearch.Text) Then
                    .TopRow = .Row
                    .Col = -1
                    .BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000D)
                    .ForeColor = System.Drawing.Color.White
                    Exit Sub
                End If
            Next
        End With
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtSearch_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtsearch.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtsearch_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtsearch.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        With SpItems
            If KeyCode = 13 And Len(Trim(txtsearch.Text)) > 0 Then
                .Col = 1
                .Value = IIf(CBool(.Value), False, True)
            End If
        End With
    End Sub
    Function ValidateData() As Boolean
        Dim intSubItem As Short
        Dim blnOtherRecord As Boolean
        Dim strMessage As String
        Dim strAccountCode As String
        Dim strCustRef As String
        Dim StrAmendmentNo As String
        Dim strSalesTaxType As String
        Dim strExciseType As String
        Dim StrItemCode As String
        Dim strSOType As String
        Dim strOtherAccountCode As String
        Dim strOtherCustRef As String
        Dim strOtherAmendmentNo As String
        Dim strOtherSalesTaxType As String
        Dim strOtherExciseType As String
        Dim strOtherItemCode As String
        Dim strOtherSOType As String
        With SpItems
            For intSubItem = 1 To .MaxRows
                .Row = intSubItem
                .Col = GridHeader.Mark
                If CBool(.Value) = True Then
                    If Not blnOtherRecord Then
                        .Col = GridHeader.AccountCode
                        strAccountCode = Trim(.Text)
                        .Col = GridHeader.CustRef
                        strCustRef = Trim(.Text)
                        .Col = GridHeader.AmendmentNo
                        StrAmendmentNo = Trim(.Text)
                        .Col = GridHeader.ItemCode
                        StrItemCode = Trim(.Text)
                        strSalesTaxType = Trim(Find_Value("select SalesTax_Type from cust_ord_hdr where UNIT_CODE = '" & gstrUNITID & "' AND Account_code='" & strAccountCode & "' and cust_ref='" & strCustRef & "' and amendment_no='" & StrAmendmentNo & "'"))
                        strSOType = Trim(Find_Value("select PO_Type from cust_ord_hdr where UNIT_CODE = '" & gstrUNITID & "' AND Account_code='" & strAccountCode & "' and cust_ref='" & strCustRef & "' and amendment_no='" & StrAmendmentNo & "'"))
                        strExciseType = Trim(Find_Value("select Excise_duty from cust_ord_dtl where UNIT_CODE = '" & gstrUNITID & "' AND Account_code='" & strAccountCode & "' and cust_ref='" & strCustRef & "' and Amendment_no ='" & StrAmendmentNo & "' and item_code ='" & StrItemCode & "' and cust_drgNo='" & StrItemCode & "'"))
                        blnOtherRecord = True
                    Else
                        .Col = GridHeader.AccountCode
                        strOtherAccountCode = Trim(.Text)
                        .Col = GridHeader.CustRef
                        strOtherCustRef = Trim(.Text)
                        .Col = GridHeader.AmendmentNo
                        strOtherAmendmentNo = Trim(.Text)
                        .Col = GridHeader.ItemCode
                        strOtherItemCode = Trim(.Text)
                        strOtherSalesTaxType = Trim(Find_Value("select SalesTax_Type from cust_ord_hdr where UNIT_CODE = '" & gstrUNITID & "' AND Account_code='" & strOtherAccountCode & "' and cust_ref='" & strOtherCustRef & "' and amendment_no='" & strOtherAmendmentNo & "'"))
                        strOtherSOType = Trim(Find_Value("select PO_Type from cust_ord_hdr where UNIT_CODE = '" & gstrUNITID & "' AND Account_code='" & strOtherAccountCode & "' and cust_ref='" & strOtherCustRef & "' and amendment_no='" & strOtherAmendmentNo & "'"))
                        strOtherExciseType = Trim(Find_Value("select Excise_duty from cust_ord_dtl where UNIT_CODE = '" & gstrUNITID & "' AND Account_code='" & strOtherAccountCode & "' and cust_ref='" & strOtherCustRef & "' and Amendment_no ='" & strOtherAmendmentNo & "' and item_code ='" & strOtherItemCode & "' and cust_drgNo='" & strOtherItemCode & "'"))
                        If UCase(strAccountCode) <> UCase(strOtherAccountCode) Then
                            strMessage = "Two or more SOs of different Customers are not allowed." & vbCrLf
                            strMessage = strMessage & "1. " & strAccountCode & " -> " & strCustRef & vbCrLf
                            strMessage = strMessage & "2. " & strOtherAccountCode & " -> " & strOtherCustRef
                            MsgBox(strMessage, MsgBoxStyle.Information, "eMPro")
                            ValidateData = False
                            Exit Function
                        End If
                        If UCase(strSOType) <> UCase(strOtherSOType) Then
                            strMessage = "OEM(O) and Spare(S) type SOs can not be included in same invoice." & vbCrLf
                            strMessage = strMessage & "1. " & strCustRef & " -> " & strSOType & vbCrLf
                            strMessage = strMessage & "2. " & strOtherCustRef & " -> " & strOtherSOType
                            MsgBox(strMessage, MsgBoxStyle.Information, "eMPro")
                            ValidateData = False
                            Exit Function
                        End If
                        If UCase(strSalesTaxType) <> UCase(strOtherSalesTaxType) Then
                            strMessage = "Two or more SOs can not have different Sales Tax Rate." & vbCrLf
                            strMessage = strMessage & "1. " & strCustRef & " -> " & strSalesTaxType & vbCrLf
                            strMessage = strMessage & "2. " & strOtherCustRef & " -> " & strOtherSalesTaxType
                            MsgBox(strMessage, MsgBoxStyle.Information, "eMPro")
                            ValidateData = False
                            Exit Function
                        End If
                        If UCase(strExciseType) <> UCase(strOtherExciseType) Then
                            strMessage = "Two or more SOs can not have different Excise Rate." & vbCrLf
                            strMessage = strMessage & "1. " & strCustRef & " -> " & strExciseType & vbCrLf
                            strMessage = strMessage & "2. " & strOtherCustRef & " -> " & strOtherExciseType
                            MsgBox(strMessage, MsgBoxStyle.Information, "eMPro")
                            ValidateData = False
                            Exit Function
                        End If
                    End If
                End If
            Next intSubItem
            ValidateData = True
        End With
    End Function
    Public Function Find_Value(ByRef strField As String) As String
        '----------------------------------------------------------------------------
        'Author         :   Arshad Ali
        'Argument       :   Sql query string as strField
        'Return Value   :   selected table field value as String
        'Function       :   Return a field value from a table
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim rs As New ADODB.Recordset
        rs = New ADODB.Recordset
        rs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        rs.Open(strField, mP_Connection, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic, ADODB.CommandTypeEnum.adCmdText)
        If rs.RecordCount > 0 Then
            If IsDBNull(rs.Fields(0).Value) = False Then
                Find_Value = rs.Fields(0).Value
            Else
                Find_Value = ""
            End If
        Else
            Find_Value = ""
        End If
        rs.Close()
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Sub SelectPreviousItem(ByRef pstrItem As String)
        Dim strMain() As String
        Dim strDet() As String
        Dim intOuterCount As Short
        Dim intInnerCount As Short

        Dim StrItemCode As String
        Dim strKanbanNo As String

        With SpItems
            strMain = Split(pstrItem, "^")
            For intOuterCount = 0 To UBound(strMain) - 1
                strDet = Split(strMain(intOuterCount), "|")
                For intInnerCount = 1 To SpItems.MaxRows
                    .Row = intInnerCount
                    .Col = GridHeader.DrawingNo
                    StrItemCode = UCase(Trim(.Text))
                    .Col = GridHeader.KanbanNo
                    strKanbanNo = UCase(Trim(.Text))

                    If UCase(Trim(strDet(0))) = StrItemCode And UCase(Trim(strDet(1))) = strKanbanNo Then
                        .Col = GridHeader.Mark
                        .Value = CStr(True)
                        Exit For
                    End If
                Next
            Next
        End With



    End Sub
End Class