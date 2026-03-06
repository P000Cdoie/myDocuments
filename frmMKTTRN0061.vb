Option Strict Off
Option Explicit On
Friend Class frmMKTTRN0061
	Inherits System.Windows.Forms.Form
	'----------------------------------------------------
	'Copyright (c)  -  MIND
	'Name of module -  frmMFGTRN0030.frm
	'Created By     -  Prashant Rajpal
	'Created Date   -  28-08-2008
	'Description    -
	'Revised date   -
	'-----------------------------------------------------------------------------
    'Modified by    :   Virendra Gupta
    'Modified ON    :   26/05/2011
    'Modified to support MultiUnit functionality
    '-----------------------------------------------------------------------
	Dim mintIndex As Short
	Dim DELFLG As Boolean
	Dim DelRows As Short
	Dim mstrIPAddress As String
	Dim strCustomer As String
	Dim strItems As String
	Private Sub cmdClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClose.Click
        Me.Dispose()
    End Sub
    Private Sub CmdDelete_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdDelete.Click
        Dim intCount As Short
        Dim VarDelete As Object
        Dim blnStatus As Boolean
        DelRows = 0
        If ValidateSelection() = False Then Exit Sub
        If Me.SpDetails.MaxRows > 0 Then
            With SpDetails
                For intCount = 1 To .MaxRows
                    VarDelete = Nothing
                    Call .GetText(1, intCount, VarDelete)
                    If VarDelete = True Then
                        blnStatus = True
                        Exit For
                    Else
                        blnStatus = False
                    End If
                Next
                If blnStatus = False Then
                    MsgBox("No Record to Delete", MsgBoxStyle.Information, "empower")
                    Exit Sub
                End If
            End With
        End If
        Call DeletionProcedure()
    End Sub
    Private Sub CmdDisplay_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdDisplay.Click
        SpDetails.MaxRows() = 0
        Me.OptDeselectAll.Checked = True
        If ValidateSelection() = False Then Exit Sub
        Call Show_Spread_details()
    End Sub
    Private Sub frmMKTTRN0061_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo ErrHandler
        mdifrmMain.CheckFormName = mintIndex
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0061_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Me.Tag) = False
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0061_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = Keys.F4 And Shift = 0 Then
            Call ctlFormHeader1_Click(ctlFormHeader1, New System.EventArgs())
        End If
    End Sub
    Private Function ValidateSelectionOnClose() As Boolean
        '-----------------------------------------------------------------------------
        'Created By -   Shruti
        'Function   -   To Validate the selection Criteria
        '-----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim intCount As Short 'Looping Variable
        Dim blnStatus As Boolean 'To Maintain the status
        Dim VarDelete As Object
        ValidateSelectionOnClose = False
        With SpDetails
            For intCount = 1 To .MaxRows
                Call .GetText(1, intCount, VarDelete)
                If VarDelete = True Then
                    blnStatus = True
                    Exit For
                Else
                    blnStatus = False
                End If
            Next
        End With
        ValidateSelectionOnClose = True
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub frmMKTTRN0061_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Escape
                If ConfirmWindow(10053, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                    gblnCancelUnload = False : gblnFormAddEdit = False
                    With Me.SpDetails
                        .MaxRows = 0 : Call Show_Spread_details()
                        Me.ToolTip1.SetToolTip(Me.SpDetails, "")
                    End With
                Else
                    Me.ActiveControl.Focus()
                End If
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
    Private Sub frmMKTTRN0061_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        opt_all_Customer.Checked = True
        opt_all_item.Checked = True
        On Error GoTo ErrHandler
        mintIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader1.Tag)
        Call FitToClient(Me, Framain, ctlFormHeader1, Frame4)
        With SpDetails
            .MaxRows = 0
            .set_RowHeight(0, 400)
            .MaxCols = 6
            .Row = 0
            .Col = 1 : .Text = " "
            .Row = 0
            .Col = 2 : .Text = "Customer Code"
            .set_ColWidth(2, 1300)
            .Row = 0
            .Col = 3 : .Text = "Item Code"
            .set_ColWidth(3, 1800)
            .Row = 0
            .Col = 4 : .Text = "Description"
            .set_ColWidth(4, 3500)
            .Row = 0
            .Col = 5 : .Text = "Quantity "
            .set_ColWidth(5, 1000)
            .Row = 0
            .Col = 6 : .Text = "Sch. Date"
            optSearchcustCode.Enabled = False : optSearchcustDes.Enabled = False
            txtSearchCustomer.Enabled = False : txtSearchCustomer.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            optSearchItemCode.Enabled = False : optSearchItemDes.Enabled = False
            txtSearchItems.Enabled = False : txtSearchItems.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            Me.lvwcustomer.Columns.Clear()
            Me.lvwcustomer.GridLines = False
            Me.lvwcustomer.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED_LIST_VIEW)
            Me.lvwItems.Enabled = False
            Me.lvwItems.Columns.Clear()
            Me.lvwItems.GridLines = False
            Me.lvwItems.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED_LIST_VIEW)
            Me.lvwItems.Enabled = False
        End With
        dtpFromDate.Format = DateTimePickerFormat.Custom
        dtpFromDate.CustomFormat = gstrDateFormat
        dtpFromDate.Value = GetServerDate()
        dtpToDate.Format = DateTimePickerFormat.Custom
        dtpToDate.CustomFormat = gstrDateFormat
        dtpToDate.Value = GetServerDate()
        OptDeselectAll.Checked = True
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0061_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        On Error GoTo ErrHandler
        Dim enmValue As eMPowerFunctions.ConfirmWindowReturnEnum
        If UnloadMode >= 0 And UnloadMode <= 5 Then
            If SpDetails.MaxRows = 0 Then Exit Sub
            If ValidateSelectionOnClose() = False Then
                enmValue = MsgBox("Do You Want To Delete The Selected Rows ?", MsgBoxStyle.YesNoCancel, "Empower")
                If enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_NO Or enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                    If enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                        'Delete data
                        Call CmdDelete_Click(CmdDelete, New System.EventArgs())
                    ElseIf enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_NO Then
                        gblnCancelUnload = False
                        gblnFormAddEdit = False
                    End If
                Else
                    'Set Global VAriable
                    gblnCancelUnload = True
                    gblnFormAddEdit = True
                End If
            Else
                gblnCancelUnload = False
                gblnFormAddEdit = False
            End If
        End If
        'Checking The Status
        If gblnCancelUnload = True Then Cancel = 1
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub frmMKTTRN0061_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error GoTo ErrHandler
        mdifrmMain.RemoveFormNameFromWindowList = mintIndex
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub Show_Spread_details()
        '----------------------------------------------------------------------------
        'Argument       :   None
        'Return Value   :   None
        'Function       :   displays data in spread
        'Comments       :   Nil
        '----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        ' To Display Details In Spread Details
        Dim i, intHours As Short
        Dim dblLossHrs As Double
        Dim strMinutes As String
        Dim intMinutes As Short
        Dim strsql As String
        Dim strItemCodes As String
        Dim intLoopcount As Short
        Dim intMaxLoop As Short
        SpDetails.Enabled = True
        Dim rsdb As New ClsResultSetDB
        Me.OptDeselectAll.Checked = True
        SpDetails.MaxRows = 0
        Me.OptDeselectAll.Checked = True
        strsql = "DELETE FROM TMP_ITEMCODE WHERE SESSION_ID= '" & gstrIpaddressWinSck & "' and Unit_Code ='" & gstrUNITID & "'"
        mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        If opt_all_item.Checked = False Then 'case for selected items
            mstrIPAddress = gstrIpaddressWinSck
            strItemCodes = ""
            intMaxLoop = lvwItems.Items.Count
            For intLoopcount = 0 To intMaxLoop - 1
                If lvwItems.Items.Item(intLoopcount).Checked = True Then
                    If Len(Trim(strItemCodes)) = 0 Then
                        strItemCodes = lvwItems.Items.Item(intLoopcount).Text
                        strsql = " INSERT INTO TMP_ITEMCODE VALUES('" & strItemCodes & "', '" & mstrIPAddress & "', '" & gstrUNITID & "') "
                        mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    Else
                        strItemCodes = strItemCodes & "," & lvwItems.Items.Item(intLoopcount).Text
                        strsql = " INSERT INTO TMP_ITEMCODE VALUES('" & lvwItems.Items.Item(intLoopcount).Text & "', '" & mstrIPAddress & "' , '" & gstrUNITID & "') "
                        mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    End If
                End If
            Next
        End If
        If opt_all_Customer.Checked = False Then
            strCustomer = ""
            intMaxLoop = lvwcustomer.Items.Count
            For intLoopcount = 0 To intMaxLoop - 1
                If lvwcustomer.Items.Item(intLoopcount).Checked = True Then
                    If Len(Trim(strCustomer)) = 0 Then
                        strCustomer = "'" & lvwcustomer.Items.Item(intLoopcount).Text & "'"
                    Else
                        strCustomer = strCustomer & ",'" & lvwcustomer.Items.Item(intLoopcount).Text & "'"
                    End If
                End If
            Next
        End If
        strsql = "select *  from forecast_mst ,item_mst  where item_mst.item_code = forecast_mst.product_no and item_mst.Unit_Code = forecast_mst.Unit_Code and  due_date between '" & Format(Me.dtpFromDate.Value, "dd MMM yyyy") & "' and '" & Format(Me.dtpToDate.Value, "dd MMM yyyy") & "' and forecast_mst.Unit_Code = '" & gstrUNITID & "'"
        If Me.opt_all_Customer.Checked = False Then
            strsql = strsql & " and customer_code in( " & strCustomer & ") "
        End If
        If Me.opt_sel_item.Checked = True Then
            strsql = strsql & " and  product_no in  (select itemcode from tmp_itemcode where "
            strsql = strsql & " tmp_itemcode.SESSION_ID ='" & mstrIPAddress & "')"
        End If
        mP_Connection.Execute("set dateformat 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        Call rsdb.GetResult(strsql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        i = 1
        If rsdb.GetNoRows > 0 Then
            SpDetails.MaxRows = rsdb.GetNoRows
            rsdb.MoveFirst()
            Do While Not rsdb.EOFRecord
                SpDetails.Row = i
                SpDetails.Col = 1
                SpDetails.CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
                SpDetails.set_ColWidth(1, 320)
                Call SpDetails.SetText(1, i, 0)
                Call SpDetails.SetText(2, i, rsdb.GetValue("customer_code"))
                Call SpDetails.set_ColWidth(2, 1300)
                Call SpDetails.SetText(3, i, rsdb.GetValue("product_no"))
                Call SpDetails.set_ColWidth(3, 1800)
                Call SpDetails.SetText(4, i, rsdb.GetValue("description"))
                Call SpDetails.set_ColWidth(4, 3500)
                Call SpDetails.SetText(5, i, rsdb.GetValue("quantity"))
                Call SpDetails.set_ColWidth(5, 1000)
                Call SpDetails.SetText(6, i, VB6.Format(rsdb.GetValue("due_date"), gstrDateFormat))
                i = i + 1
                rsdb.MoveNext()
            Loop
            SpDetails.BlockMode = True
            SpDetails.Row = 1
            SpDetails.Row2 = SpDetails.MaxRows
            SpDetails.Col = 2
            SpDetails.Col2 = 7
            SpDetails.Lock = True
            SpDetails.BlockMode = False
        Else
            If DELFLG = True Then
                MsgBox("Record deleted successfully", MsgBoxStyle.Information, "Empower")
                DELFLG = False
            Else
                MsgBox("No data for the Forecast Exists", MsgBoxStyle.Information, ResolveResString(100))
            End If
            SpDetails.MaxRows = 0
            Me.OptDeselectAll.Checked = True
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Function ValidateSelection() As Boolean
        On Error GoTo ErrHandler
        Dim intCount As Short 'Looping Variable
        Dim blnStatus As Boolean 'To Maintain the status
        Dim VarDelete As Object
        Dim intMaxCount As Short
        Dim intNoRec As Short
        Dim intMaxLoop As Short
        Dim intLoopcount As Short
        ValidateSelection = False
        If Me.opt_sel_Customer.Checked = True Then
            intMaxCount = lvwcustomer.Items.Count
            For intCount = 0 To intMaxCount - 1
                If lvwcustomer.Items.Item(intCount).Checked = True Then
                    intNoRec = intNoRec + 1
                End If
            Next
            If intNoRec <= 0 Then
                MsgBox("Select at least one Customer", MsgBoxStyle.Information, "eMPro")
                Me.SpDetails.MaxRows = 0
                Me.OptDeselectAll.Checked = True
                opt_all_item.Checked = True
                Me.lvwcustomer.Focus()
                Exit Function
            End If
        End If
        If Me.opt_sel_item.Checked = True Then
            intMaxCount = lvwItems.Items.Count
            intNoRec = 0
            For intCount = 0 To intMaxCount - 1
                If lvwItems.Items.Item(intCount).Checked = True Then
                    intNoRec = intNoRec + 1
                End If
            Next
            If intNoRec <= 0 Then
                MsgBox("Select at least one Item", MsgBoxStyle.Information, "eMPro")
                Me.SpDetails.MaxRows = 0
                Me.OptDeselectAll.Checked = True
                Exit Function
            End If
        End If
        ValidateSelection = True
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Public Sub DeletionProcedure()
        Dim strSQLA As String
        Dim varDel As Object
        Dim intCounter As Short
        Dim strSchDate As String
        Dim varCustomerCode As Object
        Dim varItemCode As Object
        Dim varShift As Object
        Dim varschdate As Object
        Dim varDocNo As Object
        Dim strsql As String
        Dim rsdb As New ClsResultSetDB
        Dim DelRows As Short
        If SpDetails.MaxRows <= 0 Then
            MsgBox("No Record to Delete", MsgBoxStyle.Information, ResolveResString(100))
            Exit Sub
        End If
        mP_Connection.BeginTrans()
        For intCounter = 1 To SpDetails.MaxRows
            varDel = Nothing
            Call SpDetails.GetText(1, intCounter, varDel)
            If varDel = 1 Then
                DelRows = DelRows + 1
                varCustomerCode = Nothing
                varItemCode = Nothing
                varschdate = Nothing
                Call SpDetails.GetText(2, intCounter, varCustomerCode)
                Call SpDetails.GetText(3, intCounter, varItemCode)
                Call SpDetails.GetText(6, intCounter, varschdate)
                varschdate = VB6.Format(varschdate.ToString, "dd/mmm/yyyy")
                mP_Connection.Execute("Set Dateformat 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                strsql = "insert into Forecast_mst_alias select Customer_code,Product_no,Due_date,Quantity,Ent_UserId,Ent_dt " & " ,'" & mP_User & "',getdate()  ,Enagare_UNLOC,Unit_Code  from " & " Forecast_mst where customer_code='" & varCustomerCode & "' and product_no='" & varItemCode & "' and due_date ='" & varschdate & "' and Unit_Code = '" & gstrUNITID & "'"
                mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                strSQLA = "Delete from forecast_mst where customer_code= '" & varCustomerCode & "' and product_no= '" & varItemCode & "' and due_date='" & varschdate & "' and Unit_Code = '" & gstrUNITID & "'"
                mP_Connection.Execute(strSQLA, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End If
        Next intCounter
        If MsgBox("Are you sure you want to Delete the Data?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2) = MsgBoxResult.No Then
            mP_Connection.RollbackTrans()
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            Exit Sub
        End If
        mP_Connection.CommitTrans()
        MsgBox(DelRows & " Records Out of " & SpDetails.MaxRows & " Deleted Succesfully", MsgBoxStyle.OkOnly, ResolveResString(100))
        Call Show_Spread_details()
    End Sub
    Private Sub lvwItems_ItemCheck(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.ItemCheckEventArgs) Handles lvwItems.ItemCheck
        Dim Item As System.Windows.Forms.ListViewItem = lvwItems.Items(eventArgs.Index)
        On Error GoTo ErrHandler
        Dim intLoopcount As Short
        Dim intMaxLoop As Short
        Dim inti As Short
        intMaxLoop = lvwcustomer.Items.Count
        For intLoopcount = 0 To intMaxLoop - 1
            If lvwcustomer.Items.Item(intLoopcount).Checked = True Then
                inti = inti + 1
            End If
        Next
        If opt_all_item.Checked = False Then
            strItems = ""
            intMaxLoop = lvwItems.Items.Count
            For intLoopcount = 0 To intMaxLoop - 1
                If lvwItems.Items.Item(intLoopcount).Checked = True Then
                    If Len(Trim(strItems)) = 0 Then
                        strItems = "'" & lvwItems.Items.Item(intLoopcount).Text & "'"
                    Else
                        strItems = strItems & ",'" & lvwItems.Items.Item(intLoopcount).Text & "'"
                    End If
                End If
            Next
        End If
        'converted comment  Call ValidateSelection()
        SpDetails.MaxRows = 0
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub opt_all_Customer_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles opt_all_Customer.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo ErrHandler
            SpDetails.MaxRows = 0
            opt_all_item.Checked = True
            Me.opt_all_Customer.Checked = True
            Me.lvwcustomer.Enabled = False
            Call Me.lvwcustomer.Items.Clear()
            Me.lvwcustomer.Columns.Clear()
            Me.lvwcustomer.GridLines = False
            Me.lvwcustomer.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED_LIST_VIEW)
            Me.lvwcustomer.Enabled = False
            Me.opt_all_item.Checked = True
            opt_all_Customer.Checked = True
            optSearchcustCode.Enabled = False : optSearchcustDes.Enabled = False
            txtSearchCustomer.Enabled = False : txtSearchCustomer.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            Exit Sub
ErrHandler:
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub
    Private Sub opt_all_item_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles opt_all_item.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo ErrHandler
            SpDetails.MaxRows = 0
            Me.opt_all_item.Checked = True
            Me.lvwItems.Enabled = False
            Me.lvwItems.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED_LIST_VIEW)
            Me.lvwItems.Items.Clear()
            Me.lvwItems.Columns.Clear()
            Me.lvwItems.GridLines = False
            Me.SpDetails.MaxRows = 0
            Me.OptDeselectAll.Checked = True
            optSearchItemCode.Enabled = False : optSearchItemDes.Enabled = False : txtSearchItems.Text = ""
            txtSearchItems.Enabled = False : txtSearchItems.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            Exit Sub
ErrHandler:
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub
    Private Sub opt_sel_Customer_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles opt_sel_Customer.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo ErrHandler
            SpDetails.MaxRows = 0
            Me.lvwcustomer.Enabled = True
            Me.lvwcustomer.View = System.Windows.Forms.View.Details
            Me.lvwcustomer.CheckBoxes = True
            Me.lvwcustomer.GridLines = True
            Me.lvwcustomer.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            Call populateCustomerlist()
            optSearchcustCode.Enabled = True : optSearchcustDes.Enabled = True : txtSearchCustomer.Enabled = True
            txtSearchCustomer.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : optSearchcustCode.Checked = True
            optSearchcustDes.Checked = False
            Exit Sub
ErrHandler:
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub
    Private Sub populateCustomerlist()
        On Error GoTo ErrHandler
        Dim objcustomers As ClsResultSetDB 'Class Object
        Dim strSQLcustomers As String 'Stores the SQL statement for getting the vendors
        Dim intcustomerCount As Short 'Stores the total vendor count
        Dim lngCustomerCtr As Integer
        Dim LstItem As System.Windows.Forms.ListViewItem
        Dim lngLoop As Integer
        lvwcustomer.Items.Clear()
        With lvwcustomer
            .Sort()
            .LabelEdit = False
            .CheckBoxes = True
            .View = System.Windows.Forms.View.Details
            .Columns.Clear()
            .Columns.Insert(0, "", "Customer Code", -2)
            .Columns.Insert(1, "", "Customer Name ", -2)
        End With
        'Building the SQL
        strSQLcustomers = "  SELECT DISTINCT C.CUSTOMER_CODE,C.CUST_NAME FROM CUSTOMER_MST C  ,"
        strSQLcustomers = strSQLcustomers & "  FORECAST_MST F WHERE  C.CUSTOMER_CODE=F.CUSTOMER_CODE and C.Unit_Code=F.Unit_Code "
        strSQLcustomers = strSQLcustomers & " and F.DUE_DATE between '" & VB6.Format(Me.dtpFromDate.Value, "dd/mmm/yyyy") & "' and '" & VB6.Format(Me.dtpToDate.Value, "dd/mmm/yyyy") & "' and C.Unit_Code = '" & gstrUNITID & "' and ((isnull(C.deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= C.deactive_date))"
        'Create the instance
        objcustomers = New ClsResultSetDB
        With objcustomers
            Call .GetResult(strSQLcustomers, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            intcustomerCount = .GetNoRows
            If intcustomerCount <= 0 Then
                Call MsgBox("No Customers have been defined. Cannot view report.", MsgBoxStyle.Information, "eMPro")
                Call opt_all_Customer_CheckedChanged(opt_all_Customer, New System.EventArgs())
                Exit Sub
            End If
            With lvwcustomer
                .Items.Clear()
                objcustomers.MoveFirst()
                For lngCustomerCtr = 0 To intcustomerCount - 1
                    LstItem = .Items.Add(Trim(objcustomers.GetValue("customer_Code")))
                    If LstItem.SubItems.Count > 1 Then
                        LstItem.SubItems(1).Text = objcustomers.GetValue("cust_name")
                    Else
                        LstItem.SubItems.Insert(1, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, objcustomers.GetValue("cust_name")))
                    End If
                    objcustomers.MoveNext()
                Next
            End With
            'Close and release the object
            .ResultSetClose()
            objcustomers = Nothing
        End With
        Me.lvwcustomer.Columns.Item(0).Width = 100
        Me.lvwcustomer.Columns.Item(1).Width = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(Me.lvwcustomer.Width) - VB6.PixelsToTwipsX(Me.lvwcustomer.Columns.Item(0).Width)) - 100)
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub opt_sel_item_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles opt_sel_item.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo ErrHandler
            Me.lvwItems.Enabled = True
            Me.lvwItems.View = System.Windows.Forms.View.Details
            Me.lvwItems.CheckBoxes = True
            Me.lvwItems.GridLines = True
            Me.lvwItems.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            optSearchItemCode.Enabled = True : optSearchItemDes.Enabled = True
            optSearchItemCode.Checked = True : txtSearchItems.Text = ""
            txtSearchItems.Enabled = True : txtSearchItems.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            Me.SpDetails.MaxRows = 0
            Me.OptDeselectAll.Checked = True
            Call PopulateItemCode()
            Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub
    Private Sub PopulateItemCode()
        On Error GoTo ErrHandler
        Dim objitem As ClsResultSetDB
        Dim strsql As String
        Dim strMachine As String
        Dim strItem As String
        Dim LstItem As System.Windows.Forms.ListViewItem
        Dim lngLoop, lngloop1 As Integer
        Dim lngRows As Integer
        Dim intMaxCount As Short
        Dim intCount As Short
        Dim intNoRec As Short
        Dim intMaxLoop As Short
        Dim intLoopcount As Short
        If Me.opt_sel_item.Checked = True Then
            With lvwItems
                .LabelEdit = False
                .CheckBoxes = True
                .View = System.Windows.Forms.View.Details
                .Columns.Clear()
                .Columns.Insert(0, "", "Item Code", -2)
                .Columns.Insert(1, "", "Description", -2)
                If opt_all_item.Checked = True Then
                    .Enabled = False
                Else
                    .Enabled = True
                End If
            End With
        End If
        '******************
        If Me.opt_sel_Customer.Checked = True Then
            intMaxCount = lvwcustomer.Items.Count
            For intCount = 0 To intMaxCount - 1
                If lvwcustomer.Items.Item(intCount).Checked = True Then
                    intNoRec = intNoRec + 1
                End If
            Next
            If intNoRec <= 0 Then
                MsgBox("Select at least one Customer", MsgBoxStyle.Information, "eMPro")
                Me.SpDetails.MaxRows = 0
                Me.OptDeselectAll.Checked = True
                opt_all_item.Checked = True
                Me.lvwcustomer.Focus()
                Exit Sub
            End If
        End If
        mP_Connection.Execute(" set dateformat 'dmy' ", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        strsql = "SELECT DISTINCT F.PRODUCT_NO ,I.DESCRIPTION FROM ITEM_MST I ,FORECAST_MST F WHERE"
        strsql = strsql & " I.ITEM_CODE=F.PRODUCT_NO and I.Unit_Code = F.Unit_Code "
        strsql = strsql & " and F.DUE_DATE between '" & VB6.Format(Me.dtpFromDate.Value, "dd/mmm/yyyy") & "' and '" & VB6.Format(Me.dtpToDate.Value, "dd/mmm/yyyy") & "' and I.Unit_Code = '" & gstrUNITID & "'"
        If opt_all_Customer.Checked = False Then
            strCustomer = ""
            intMaxLoop = lvwcustomer.Items.Count
            For intLoopcount = 0 To intMaxLoop - 1
                If lvwcustomer.Items.Item(intLoopcount).Checked = True Then
                    If Len(Trim(strCustomer)) = 0 Then
                        strCustomer = "'" & lvwcustomer.Items.Item(intLoopcount).Text & "'"
                    Else
                        strCustomer = strCustomer & ",'" & lvwcustomer.Items.Item(intLoopcount).Text & "'"
                    End If
                End If
            Next
            If Len(Trim(strCustomer)) > 0 Then
                strsql = strsql & " and F.CUSTOMER_CODE  IN(" & strCustomer & ")"
            End If
        End If
        strsql = strsql & " ORDER BY F.PRODUCT_NO "
        objitem = New ClsResultSetDB
        Call objitem.GetResult(strsql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If Me.opt_all_item.Checked = True Then
            Exit Sub
        End If
        With Me.lvwItems
            .Items.Clear()
            lngRows = objitem.GetNoRows
            If lngRows <= 0 Then
                Call MsgBox("No Item have been defined. Cannot view report.", MsgBoxStyle.Information, ResolveResString(100))
                Call opt_all_item_CheckedChanged(opt_all_item, New System.EventArgs())
                Call opt_all_Customer_CheckedChanged(opt_all_Customer, New System.EventArgs())
                Exit Sub
            Else
                txtSearchItems.Enabled = True : txtSearchItems.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            End If
            objitem.MoveFirst()
            For lngLoop = 0 To lngRows - 1
                If Len(Trim(objitem.GetValue("PRODUCT_NO"))) > 0 Then
                    .Items.Insert(lngLoop, objitem.GetValue("PRODUCT_NO "))
                    .Items.Item(lngLoop).SubItems.Add(objitem.GetValue("description"))
                End If
                objitem.MoveNext()
            Next
            Me.lvwItems.Columns.Item(0).Width = 150
            Me.lvwItems.Columns.Item(1).Width = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(Me.lvwItems.Width) - VB6.PixelsToTwipsX(Me.lvwItems.Columns.Item(0).Width)) - 300)
        End With
        objitem.ResultSetClose()
        objitem = Nothing
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub lvwCustomer_ItemCheck(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.ItemCheckEventArgs) Handles lvwcustomer.ItemCheck
        Dim Item As System.Windows.Forms.ListViewItem = lvwcustomer.Items(eventArgs.Index)
        On Error GoTo ErrHandler
        Dim intLoopcount As Short
        Dim intMaxLoop As Short
        Dim inti As Short
        intMaxLoop = lvwcustomer.Items.Count
        For intLoopcount = 0 To intMaxLoop - 1
            If lvwcustomer.Items.Item(intLoopcount).Checked = True Then
                inti = inti + 1
            End If
        Next
        opt_all_item.Checked = True
        If Me.opt_sel_Customer.Checked = True And SpDetails.MaxRows > 0 Then Call Show_Spread_details()
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub OptDeselectAll_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OptDeselectAll.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo ErrHandler
            SelectDeselect((False))
            Exit Sub
ErrHandler:
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub
    Private Sub optSearchCustCode_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optSearchcustCode.CheckedChanged
        If eventSender.Checked Then
            txtSearchCustomer.Text = ""
        End If
    End Sub
    Private Sub optSearchCustDes_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optSearchcustDes.CheckedChanged
        If eventSender.Checked Then
            txtSearchCustomer.Text = ""
        End If
    End Sub
    Private Sub optSearchItemCode_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optSearchItemCode.CheckedChanged
        If eventSender.Checked Then
            txtSearchItems.Text = ""
        End If
    End Sub
    Private Sub optSearchItemDes_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optSearchItemDes.CheckedChanged
        If eventSender.Checked Then
            txtSearchItems.Text = ""
        End If
    End Sub
    Private Sub OptSelectAll_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OptSelectAll.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo ErrHandler
            If Me.SpDetails.MaxRows <= 0 Then
                Me.OptDeselectAll.Checked = True
                Exit Sub
            End If
            SelectDeselect((True))
            Exit Sub
ErrHandler:
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub
    Private Sub SelectDeselect(ByVal blnCheck As Boolean)
        On Error GoTo ErrHandler
        Dim irow As Integer
        For irow = 1 To SpDetails.MaxRows
            SpDetails.Row = irow
            SpDetails.Col = 1
            SpDetails.Value = IIf(blnCheck = True, 1, 0)
        Next irow
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtSearchCustomer_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSearchCustomer.TextChanged
        Call search(lvwcustomer, txtSearchCustomer, optSearchcustCode, optSearchcustDes)
    End Sub
    Private Sub txtSearchItems_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSearchItems.TextChanged
        Call search(lvwItems, txtSearchItems, optSearchItemCode, optSearchItemDes)
    End Sub
    Public Sub search(ByRef lvwListView As System.Windows.Forms.ListView, ByRef txtSearchBox As System.Windows.Forms.TextBox, ByRef optFistOption As System.Windows.Forms.RadioButton, ByRef optSecOption As System.Windows.Forms.RadioButton)
        On Error GoTo ErrHandler
        Dim intCounter As Short
        With lvwListView
            If optFistOption.Checked = True Then
                For intCounter = 0 To .Items.Count - 1
                    If .Items.Item(intCounter).Font.Bold = True Then
                        .Items.Item(intCounter).Font = VB6.FontChangeBold(.Items.Item(intCounter).Font, False)
                        .Refresh()
                    End If
                    If .Items.Item(intCounter).SubItems.Item(1).Font.Bold = True Then
                        .Items.Item(intCounter).SubItems.Item(1).Font = VB6.FontChangeBold(.Items.Item(intCounter).SubItems.Item(1).Font, False)
                        .Refresh()
                    End If
                Next
                If Len(txtSearchBox.Text) = 0 Then Exit Sub
                For intCounter = 0 To .Items.Count - 1
                    If Trim(UCase(Mid(.Items.Item(intCounter).Text, 1, Len(txtSearchBox.Text)))) = Trim(UCase(txtSearchBox.Text)) Then
                        .Items.Item(intCounter).Font = VB6.FontChangeBold(.Items.Item(intCounter).Font, True)
                        Call .Items.Item(intCounter).EnsureVisible()
                        .Refresh()
                        Exit For
                    End If
                Next
            ElseIf optSecOption.Checked Then
                For intCounter = 0 To .Items.Count - 1
                    If .Items.Item(intCounter).Font.Bold = True Then
                        .Items.Item(intCounter).Font = VB6.FontChangeBold(.Items.Item(intCounter).Font, False)
                        .Refresh()
                    End If
                    If .Items.Item(intCounter).SubItems.Item(1).Font.Bold = True Then
                        .Items.Item(intCounter).SubItems.Item(1).Font = VB6.FontChangeBold(.Items.Item(intCounter).SubItems.Item(1).Font, False)
                        .Refresh()
                    End If
                Next
                If Len(txtSearchBox.Text) = 0 Then Exit Sub
                For intCounter = 0 To .Items.Count - 1
                    If Trim(UCase(Mid(.Items.Item(intCounter).SubItems.Item(1).Text, 1, Len(txtSearchBox.Text)))) = Trim(UCase(txtSearchBox.Text)) Then
                        .Items.Item(intCounter).SubItems.Item(1).Font = VB6.FontChangeBold(.Items.Item(intCounter).SubItems.Item(1).Font, True)
                        Call .Items.Item(intCounter).EnsureVisible()
                        .Refresh()
                        Exit For
                    End If
                Next
            End If
        End With
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub dtpToDate_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpToDate.TextChanged
        dtpToDate.MinDate = GetServerDate()
        If dtpToDate.Value < dtpFromDate.Value Then dtpToDate.Value = dtpFromDate.Value
        SpDetails.MaxRows = 0
        opt_all_Customer.Checked = True
        opt_all_item.Checked = True
    End Sub
    Private Sub ctlFormHeader1_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlFormHeader1.Click
        On Error GoTo ErrHandler
        Call ShowHelp("HLP" & Mid(Me.Name, 4, Len(Me.Name)) & ".htm") '("HLPCSTMS0001.htm")
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub dtpFromDate_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpFromDate.TextChanged
        dtpFromDate.MinDate = GetServerDate()
        If dtpToDate.Value < dtpFromDate.Value Then dtpToDate.Value = dtpFromDate.Value
        opt_all_Customer.Checked = True
        opt_all_item.Checked = True
        SpDetails.MaxRows = 0
    End Sub
    Private Sub dtpFromDate_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles dtpFromDate.Validating
        Dim intdiffdate As Short
        intdiffdate = DateDiff(DateInterval.Day, GetServerDate(), dtpFromDate.Value)
        If intdiffdate < 0 Then
            Call MsgBox("Date Cannot be less than Current Date ", MsgBoxStyle.Information, ResolveResString(100))
            dtpFromDate.Value = GetServerDate()
        End If
    End Sub
    Private Sub dtpToDate_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles dtpToDate.Validating
        Dim intdiffdate As Short
        intdiffdate = DateDiff(DateInterval.Day, GetServerDate(), dtpFromDate.Value)
        If intdiffdate < 0 Then
            Call MsgBox("Date Cannot be less than Current Date ", MsgBoxStyle.Information, ResolveResString(100))
            dtpFromDate.Value = GetServerDate()
        End If
    End Sub
    Private Sub lvwItems_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles lvwItems.Validating
        On Error GoTo ErrHandler
        Dim intCount As Short 'Looping Variable
        Dim blnStatus As Boolean 'To Maintain the status
        Dim VarDelete As Object
        Dim intMaxCount As Short
        Dim intNoRec As Short
        Dim intMaxLoop As Short
        Dim intLoopcount As Short
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
End Class