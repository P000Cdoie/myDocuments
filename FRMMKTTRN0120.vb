Option Strict Off
Option Explicit On
Imports System.Data
Imports System.Data.SqlClient
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Friend Class FRMMKTTRN0120
    Inherits System.Windows.Forms.Form
    '----------------------------------------------------
    'Copyright(c)        - MIND
    'Form Name           - FRMMKTTRN0120.frm
    'Created by          - Priti sharma
    'Created Date        - 18/10/2021
    'Modified Date       -
    'Form Description    - DS wise Planned / Unplanned

    Dim mintFormIndex As Short 'Stores the related menu index in Windows list
    Dim mblnUnload As Boolean
    Dim Bool_First_Check As Boolean = False

    Private Enum GridInvoiceDetail
        CustomerCode = 0
        ItemCode
        ItemDesc
        Trans_date
        ScheduleQty
        DispatchQty
        Serial_No
        DSNO
        Remarks
    End Enum
    Private Sub dtpDateFrom_ValueChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpDateFrom.ValueChanged
        On Error GoTo ErrHandler
        If dtpDateFrom.Value > dtpDateTo.Value Then dtpDateFrom.Value = dtpDateTo.Value
        Call InitializeControls()
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub dtpDateTo_ValueChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dtpDateTo.ValueChanged
        On Error GoTo ErrHandler
        If dtpDateFrom.Value > dtpDateTo.Value Then dtpDateTo.Value = dtpDateFrom.Value
        Call InitializeControls()
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub dtpDateFrom_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dtpDateFrom.KeyDown
        On Error GoTo ErrHandler
        If e.KeyCode = 13 Then
            Me.dtpDateTo.Focus()
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub dtpDateTo_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles dtpDateTo.KeyDown
        On Error GoTo ErrHandler
        If eventArgs.KeyCode = 13 Then
           
            If opt_all_Customer.Checked = True Then
                opt_all_Customer.Focus()
            Else
                Me.opt_sel_Customer.Focus()
            End If

        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTREP0044_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo Err_Handler
        If mblnUnload = True Then Me.Close() : Exit Sub
        frmModules.NodeFontBold(Tag) = True
        'Checking the form name in the Windows list
        mdifrmMain.CheckFormName = mintFormIndex
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTREP0044_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo ErrHandler
        'Get the index of form in the Windows list
        mintFormIndex = mdifrmMain.AddFormNameToWindowList(Me.ctlDSWiseScheduleStatus.Tag)
        'Fit to client area
        frmModules.NodeFontBold(Me.Tag) = True
        'Refresh the form
        ResetData()
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Sub InitializeControls()
        '-----------------------------------------------------------------------------
        'Author         :   Ashutosh Verma
        'Arguments      :   None
        'Return Value   :   None
        'Function       :   Initialize the controls of the form.
        '-----------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Call FitToClient(Me, frmMain, ctlDSWiseScheduleStatus, GrpBoxButtons, 500)
        With Me
            .lvwItems.Enabled = False
            .lvwcustomer.Enabled = False
        End With
        'Initializing option buttons
        opt_all_Customer.Checked = True : opt_sel_Customer.Checked = False : optSearchCustCode.Enabled = False
        optSearchCustName.Enabled = False : txtSearchCustomer.Enabled = False
        txtSearchCustomer.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        opt_all_item.Checked = True : opt_sel_item.Checked = False : optSearchItemCode.Enabled = False
        optSearchItemDes.Enabled = False : txtSearchItems.Enabled = False
        txtSearchItems.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub frmMKTREP0044_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        On Error GoTo ErrHandler
        'Make the node normal font
        frmModules.NodeFontBold(Tag) = False
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub frmMKTREP0044_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error GoTo ErrHandler
        'Me = Nothing
        'Removing the form name from list
        mdifrmMain.RemoveFormNameFromWindowList = mintFormIndex
        'Setting the corresponding node's bold property
        frmModules.NodeFontBold(Tag) = False
        gblnCancelUnload = False
        Me.Dispose()
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub frmMKTREP0044_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        'If KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0 Then Call ctlDSWiseScheduleStatus_ClickEvent(ctlDSWiseScheduleStatus, New System.EventArgs())
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTREP0044_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        '    If KeyAscii = 13 Then
        '        SendKeys "{TAB}"
        '    End If
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    
    Private Function populateCustomerlist() As String
        '*******************************************************************************
        'Author             :   Ashuosh Verma
        'Argument(s)if any  :
        'Return Value       :   Returns (Y or N)-- to distinguish weather the function runs successfully or not.
        'Function           :   Populates the customer list for selected schedule date range.
        'Comments           :   NA
        'Creation Date      :   28/11/2006
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim objcustomers As ClsResultSetDB 'Class Object
        Dim strSQLcustomers As String 'Stores the SQL statement for getting the customers
        Dim intcustomerCount As Short 'Stores the total customer count
        Dim lngCustomerCtr As Integer
        lvwcustomer.Items.Clear()
        With lvwcustomer
            .Sort()
            .LabelEdit = False
            .View = System.Windows.Forms.View.Details
            .Columns.Clear()
            .Columns.Insert(0, "", "Customer Code", -2)
            .Columns.Insert(1, "", "Customer Name", -2)
            If opt_all_Customer.Checked = True Then
                .Enabled = False
            Else
                .Enabled = True
            End If
        End With
        strSQLcustomers = "select Distinct(a.Account_Code) as Customer_Code,b.Cust_Name from DailyMktSchedule a ( nolock) ,customer_mst b(nolock) "
        strSQLcustomers = strSQLcustomers & " where  a.UNIT_CODE=b.UNIT_CODE and a.Account_Code=b.Customer_Code  and isnull(manual_ds_closure,0)=0 AND a.UNIT_CODE='" & gstrUNITID & "'"
        strSQLcustomers = strSQLcustomers & " and  a.Trans_date between '" & VB6.Format(Me.dtpDateFrom.Value, "dd/mmm/yyyy") & "' and '" & VB6.Format(Me.dtpDateTo.Value, "dd/mmm/yyyy") & "' "
        objcustomers = New ClsResultSetDB
        With objcustomers
            Call .GetResult(strSQLcustomers, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            'Get the record count
            intcustomerCount = .GetNoRows
            If intcustomerCount <= 0 Then
                'Show message to the user
                MsgBox("No Customer found for selected date range", MsgBoxStyle.Information, ResolveResString(100))
                'Set the boolean variable
                mblnUnload = True
                'Close and release the object
                .ResultSetClose()
                objcustomers = Nothing
                'Exit from procedure
                Me.opt_all_Customer.Checked = True
                populateCustomerlist = "N"
                Exit Function
            End If
        End With
        'Populate in list view
        Dim lstCustomer As System.Windows.Forms.ListViewItem
        Dim lngloop1 As Short 'List Item Object
        With lvwcustomer
            .Items.Clear()
            objcustomers.MoveFirst()
            For lngCustomerCtr = 0 To intcustomerCount - 1
                lstCustomer = lvwcustomer.Items.Add(Trim(objcustomers.GetValue("Customer_Code")))
                '''lstCustomer.SubItems(lngCustomerCtr) = objcustomers.GetValue("Cust_Name")
                For lngloop1 = 0 To objcustomers.GetFieldCount - 2
                    If lstCustomer.SubItems.Count > lngloop1 + 1 Then
                        lstCustomer.SubItems(lngloop1 + 1).Text = objcustomers.GetValueByNo(lngloop1 + 1)
                    Else
                        lstCustomer.SubItems.Insert(lngloop1 + 1, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, objcustomers.GetValueByNo(lngloop1 + 1)))
                        'lstCustomer.SubItems.Add(objcustomers.GetValueByNo(lngloop1))
                    End If
                Next
                objcustomers.MoveNext()
            Next lngCustomerCtr
            objcustomers.ResultSetClose()
        End With
        'Close and release the object
        'Me.lvwcustomer.Columns.Item(2).Width = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(Me.lvwcustomer.Width) - VB6.PixelsToTwipsX(Me.lvwcustomer.Columns.Item(1).Width)) - 100)
        Me.lvwcustomer.Columns.Item(0).Width = 100
        Me.lvwcustomer.Columns.Item(1).Width = 400
        populateCustomerlist = "Y"
        Exit Function
ErrHandler:
        populateCustomerlist = "N"
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function

    Private Sub ConfigureGridColumn()
        Try
           

            dgvInvoiceDetail.Columns.Clear()

            dgvInvoiceDetail.Columns.Add(GridInvoiceDetail.CustomerCode, "Customer Code")
            dgvInvoiceDetail.Columns.Add(GridInvoiceDetail.ItemCode, "Item Code")
            dgvInvoiceDetail.Columns.Add(GridInvoiceDetail.ItemDesc, "Item desc")
            dgvInvoiceDetail.Columns.Add(GridInvoiceDetail.Trans_date, "Trans Date")
            dgvInvoiceDetail.Columns.Add(GridInvoiceDetail.ScheduleQty, "Schedule Qty")
            dgvInvoiceDetail.Columns.Add(GridInvoiceDetail.DispatchQty, "Despatch Qty")
            dgvInvoiceDetail.Columns.Add(GridInvoiceDetail.Serial_No, "Serial No")
            dgvInvoiceDetail.Columns.Add(GridInvoiceDetail.DSNO, "DS No")
            dgvInvoiceDetail.Columns.Add(GridInvoiceDetail.Remarks, "Remark")

            dgvInvoiceDetail.Columns(GridInvoiceDetail.CustomerCode).Width = 100
            dgvInvoiceDetail.Columns(GridInvoiceDetail.ItemCode).Width = 100
            dgvInvoiceDetail.Columns(GridInvoiceDetail.ItemDesc).Width = 100
            dgvInvoiceDetail.Columns(GridInvoiceDetail.Trans_date).Width = 100
            dgvInvoiceDetail.Columns(GridInvoiceDetail.ScheduleQty).Width = 100
            dgvInvoiceDetail.Columns(GridInvoiceDetail.DispatchQty).Width = 100
            dgvInvoiceDetail.Columns(GridInvoiceDetail.Serial_No).Width = 100
            dgvInvoiceDetail.Columns(GridInvoiceDetail.DSNO).Width = 100
            dgvInvoiceDetail.Columns(GridInvoiceDetail.Remarks).Width = 100

         
            dgvInvoiceDetail.Columns(GridInvoiceDetail.CustomerCode).HeaderCell.Style.Font = New Font(dgvInvoiceDetail.Font, FontStyle.Bold)
            dgvInvoiceDetail.Columns(GridInvoiceDetail.ItemCode).HeaderCell.Style.Font = New Font(dgvInvoiceDetail.Font, FontStyle.Bold)
            dgvInvoiceDetail.Columns(GridInvoiceDetail.ItemDesc).HeaderCell.Style.Font = New Font(dgvInvoiceDetail.Font, FontStyle.Bold)
            dgvInvoiceDetail.Columns(GridInvoiceDetail.Trans_date).HeaderCell.Style.Font = New Font(dgvInvoiceDetail.Font, FontStyle.Bold)
            dgvInvoiceDetail.Columns(GridInvoiceDetail.ScheduleQty).HeaderCell.Style.Font = New Font(dgvInvoiceDetail.Font, FontStyle.Bold)
            dgvInvoiceDetail.Columns(GridInvoiceDetail.DispatchQty).HeaderCell.Style.Font = New Font(dgvInvoiceDetail.Font, FontStyle.Bold)
            dgvInvoiceDetail.Columns(GridInvoiceDetail.Serial_No).HeaderCell.Style.Font = New Font(dgvInvoiceDetail.Font, FontStyle.Bold)
            dgvInvoiceDetail.Columns(GridInvoiceDetail.DSNO).HeaderCell.Style.Font = New Font(dgvInvoiceDetail.Font, FontStyle.Bold)
            dgvInvoiceDetail.Columns(GridInvoiceDetail.Remarks).HeaderCell.Style.Font = New Font(dgvInvoiceDetail.Font, FontStyle.Bold)

            dgvInvoiceDetail.Columns(GridInvoiceDetail.CustomerCode).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvInvoiceDetail.Columns(GridInvoiceDetail.ItemCode).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvInvoiceDetail.Columns(GridInvoiceDetail.ItemDesc).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvInvoiceDetail.Columns(GridInvoiceDetail.Trans_date).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvInvoiceDetail.Columns(GridInvoiceDetail.ScheduleQty).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvInvoiceDetail.Columns(GridInvoiceDetail.DispatchQty).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvInvoiceDetail.Columns(GridInvoiceDetail.Serial_No).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvInvoiceDetail.Columns(GridInvoiceDetail.DSNO).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            dgvInvoiceDetail.Columns(GridInvoiceDetail.Remarks).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            dgvInvoiceDetail.Columns(GridInvoiceDetail.CustomerCode).ReadOnly = True
            dgvInvoiceDetail.Columns(GridInvoiceDetail.ItemCode).ReadOnly = True
            dgvInvoiceDetail.Columns(GridInvoiceDetail.ItemDesc).ReadOnly = True
            dgvInvoiceDetail.Columns(GridInvoiceDetail.Trans_date).ReadOnly = True
            dgvInvoiceDetail.Columns(GridInvoiceDetail.ScheduleQty).ReadOnly = True
            dgvInvoiceDetail.Columns(GridInvoiceDetail.DispatchQty).ReadOnly = True
            dgvInvoiceDetail.Columns(GridInvoiceDetail.Serial_No).ReadOnly = True
            dgvInvoiceDetail.Columns(GridInvoiceDetail.DSNO).ReadOnly = True
            dgvInvoiceDetail.Columns(GridInvoiceDetail.Remarks).ReadOnly = False
           

            dgvInvoiceDetail.Columns(GridInvoiceDetail.CustomerCode).SortMode = DataGridViewColumnSortMode.NotSortable
            dgvInvoiceDetail.Columns(GridInvoiceDetail.ItemCode).SortMode = DataGridViewColumnSortMode.NotSortable
            dgvInvoiceDetail.Columns(GridInvoiceDetail.ItemDesc).SortMode = DataGridViewColumnSortMode.NotSortable
            dgvInvoiceDetail.Columns(GridInvoiceDetail.Trans_date).SortMode = DataGridViewColumnSortMode.NotSortable
            dgvInvoiceDetail.Columns(GridInvoiceDetail.ScheduleQty).SortMode = DataGridViewColumnSortMode.NotSortable
            dgvInvoiceDetail.Columns(GridInvoiceDetail.DispatchQty).SortMode = DataGridViewColumnSortMode.NotSortable
            dgvInvoiceDetail.Columns(GridInvoiceDetail.Serial_No).SortMode = DataGridViewColumnSortMode.NotSortable
            dgvInvoiceDetail.Columns(GridInvoiceDetail.DSNO).SortMode = DataGridViewColumnSortMode.NotSortable
            dgvInvoiceDetail.Columns(GridInvoiceDetail.Remarks).SortMode = DataGridViewColumnSortMode.NotSortable

         
        Catch ex As Exception
            RaiseException(ex)
        End Try
    End Sub
    '    Private Sub fraRepCust_ButtonClick(ByVal eventSender As System.Object, ByVal e As UCActXCtl.UCfraRepCmd.ButtonClickEventArgs) Handles fraRepCust.ButtonClick
    '        '*******************************************************************************
    '        'Author             :   Ashutosh Verma
    '        'Argument(s)if any  :
    '        'Return Value       :   NA
    '        'Function           :   To add functionaliy PRINT,PREVIEW,CLOSE
    '        'Comments           :   NA
    '        'Creation Date      :   28/11/2006
    '        '*******************************************************************************
    '        On Error GoTo ErrHandler
    '        Dim strDate As String
    '        Dim lngLoop As Integer
    '        Dim intCount As Short
    '        Dim datDS_Date As Date
    '        Dim strQSNo As String
    '        Select Case e.Button
    '            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
    '                Me.Close()
    '                Exit Sub
    '        End Select
    '        'checking if atleast one customer is selected in case selected customers option is selected.
    '        intCount = 0
    '        If Me.opt_sel_Customer.Checked = True Then
    '            For lngLoop = 0 To Me.lvwcustomer.Items.Count - 1
    '                If Me.lvwcustomer.Items.Item(lngLoop).Checked = True Then
    '                    intCount = intCount + 1
    '                End If
    '            Next lngLoop
    '            If intCount = 0 Then
    '                MsgBox("Select Atleast one Customer.", MsgBoxStyle.Information, ResolveResString(100))
    '                Me.lvwcustomer.Focus()
    '                Exit Sub
    '            End If
    '        End If

    '        'calling sub to execute the SP
    '        Call execSP()



    'ErrHandler:
    '        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    '    End Sub
    Public Sub execSP()
        '*******************************************************************************
        'Author             :   Priti Sharma
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :
        'Comments           :   NA
        'Creation Date      :   28/10/2021
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        Dim strDrgNo As String
        Dim strsql As String
        Dim strSQLItem As String
        Dim strCustomer As String
        Dim intLoopcount As Short
        Dim intWeekday As Short
        Dim intMaxCount As Short
        Dim intCount As Short
        Dim intNoRec As Short

        If Me.opt_sel_Customer.Checked = True Then
            intNoRec = 0 : intMaxCount = 0
            'Checking for checked item type(s)
            intMaxCount = lvwcustomer.Items.Count
            For intCount = 0 To intMaxCount - 1
                If lvwcustomer.Items.Item(intCount).Checked = True Then
                    intNoRec = intNoRec + 1
                End If
            Next
            If intNoRec <= 0 Then
                MsgBox("Select at least one Customer", MsgBoxStyle.Information, ResolveResString(100))
                lvwcustomer.Focus()
                opt_all_item.Checked = True
                opt_sel_item.Checked = False
                lvwItems.Enabled = False
                Exit Sub
            End If
        End If

        If Me.opt_sel_item.Checked = True Then
            intNoRec = 0 : intMaxCount = 0
            'Checking for checked item type(s)
            intMaxCount = lvwItems.Items.Count
            For intCount = 0 To intMaxCount - 1
                If lvwItems.Items.Item(intCount).Checked = True Then
                    intNoRec = intNoRec + 1
                End If
            Next
            If intNoRec <= 0 Then
                MsgBox("Select at least one Item", MsgBoxStyle.Information, ResolveResString(100))
                lvwItems.Focus()
                Exit Sub
            End If
        End If

        mP_Connection.Execute("Delete DSCustomerCode Where IP_ADDRESS = '" & gstrIpaddressWinSck & "' AND UNIT_CODE='" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        mP_Connection.Execute("Delete DSItemCode Where IP_ADDRESS = '" & gstrIpaddressWinSck & "' AND UNIT_CODE='" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        intMaxLoop = lvwcustomer.Items.Count
        If Me.opt_all_Customer.Checked = True Then
            strsql = "insert into DSCustomerCode (Account_Code,UNIT_CODE,IP_ADDRESS) Select Distinct(Customer_code),a.UNIT_CODE,'" & gstrIpaddressWinSck & "' from DailyMktSchedule a ( nolock) ,customer_mst b(nolock) ,Item_mst c (Nolock) " & _
        "where  a.UNIT_CODE=b.UNIT_CODE and a.Account_Code=b.Customer_Code  and a.unit_code=c.unit_code and " & _
        "a.item_code=c.item_code and a.unit_code='" & gstrUNITID & "' and c.Status ='A' and c.Hold_Flag <> 1  and isnull(manual_ds_closure,0)=0  and a.Trans_date between '" & VB6.Format(Me.dtpDateFrom.Value, "dd/mmm/yyyy") & "' and '" & VB6.Format(Me.dtpDateTo.Value, "dd/mmm/yyyy") & "' "
            mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        Else
            For intLoopCounter = 0 To intMaxLoop - 1
                If lvwcustomer.Items.Item(intLoopCounter).Checked = True Then
                    strsql = "insert into DSCustomerCode (Account_Code,UNIT_CODE,IP_ADDRESS) Values('" & Trim(lvwcustomer.Items.Item(intLoopCounter).Text) & "','" & gstrUNITID & "','" & gstrIpaddressWinSck & "')"
                    mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                End If
            Next
        End If
        intMaxLoop = lvwItems.Items.Count
        If Me.opt_all_item.Checked = True Then
            strSQLItem = "select Distinct(a.Item_code) as Item_code,a.Unit_code,'" & gstrIpaddressWinSck & "' from DailyMktSchedule a ( nolock) ,customer_mst b(nolock) ,Item_mst c (Nolock) " & _
        "where  a.UNIT_CODE=b.UNIT_CODE and a.Account_Code=b.Customer_Code  and a.unit_code=c.unit_code and " & _
        "a.item_code=c.item_code and a.unit_code='" & gstrUNITID & "' and c.Status ='A' and c.Hold_Flag <> 1  and isnull(manual_ds_closure,0)=0  and a.Trans_date between '" & VB6.Format(Me.dtpDateFrom.Value, "dd/mmm/yyyy") & "' and '" & VB6.Format(Me.dtpDateTo.Value, "dd/mmm/yyyy") & "' "
            If opt_sel_Customer.Checked = True Then
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
                    strSQLItem = strSQLItem & " and a.Account_code in (" & strCustomer & ")"
                End If
            End If
            mP_Connection.Execute("insert into DSItemCode (Item_Code,UNIT_CODE,IP_ADDRESS) " & strSQLItem, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        Else
            For intLoopCounter = 0 To intMaxLoop - 1
                If lvwItems.Items.Item(intLoopCounter).Checked = True Then
                    strDrgNo = Trim(lvwItems.Items.Item(intLoopCounter).Text)
                    '''strItemCode = Trim(lvwItems.ListItems(intLoopCounter).ListSubItems.Item(2))
                    mP_Connection.Execute("insert into DSItemCode (Item_Code,UNIT_CODE) Values('" & strDrgNo & "','" & gstrUNITID & "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                End If
            Next
        End If
        mP_Connection.Execute("update DSItemCode set IP_ADDRESS = '" & gstrIpaddressWinSck & "' where UNIT_CODE='" & gstrUNITID & "' AND IP_ADDRESS is NULL", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        FillDS()
        Exit Sub

ErrHandler:
        ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, Me, System.Windows.Forms.Cursors.Default)
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    Private Sub FillDS()
        Dim dt As New DataTable
        Dim sqlCmd As New SqlCommand
        Try
            Dim i As Integer = 0
            With sqlCmd
                .CommandType = CommandType.StoredProcedure
                .CommandTimeout = 300 ' 5 Minute
                .CommandText = "DS_CLOSURE"
                .Parameters.Clear()
                .Parameters.AddWithValue("@unitcode", gstrUNITID)
                .Parameters.AddWithValue("@DateFrom", VB6.Format(Me.dtpDateFrom.Value, "dd/mmm/yyyy"))
                .Parameters.AddWithValue("@DateTo", VB6.Format(Me.dtpDateTo.Value, "dd/mmm/yyyy"))
                .Parameters.AddWithValue("@IP_ADDRESS", gstrIpaddressWinSck)
                .Parameters.AddWithValue("@User_Id", mP_User)
                .Parameters.AddWithValue("@TYPE", "CLOSURE")
                '.Parameters.Add("@MESSAGE", SqlDbType.VarChar, 8000).Direction = ParameterDirection.Output
                'ds = SqlConnectionclass.GetDataSet(sqlCmd)
                'If Convert.ToString(.Parameters("@MESSAGE").Value) <> "" Then
                '    MsgBox(Convert.ToString(.Parameters("@MESSAGE").Value), MsgBoxStyle.Exclamation, ResolveResString(100))
                'Else
                SqlConnectionclass.ExecuteNonQuery(sqlCmd)
                Dim strSQL As String = "SELECT * FROM TEMP_Dailymktschedule_CLOSURE WHERE UNIT_CODE ='" & gstrUNITID & "' AND IP_ADDRESS ='" & gstrIpaddressWinSck & "' and type='CLOSURE'"
                dt = SqlConnectionclass.GetDataTable(strSQL)
                i = 0
                dgvInvoiceDetail.Rows.Clear()
                If dt.Rows.Count > 0 Then
                    dgvInvoiceDetail.Rows.Add(dt.Rows.Count)
                    For Each dr As DataRow In dt.Rows
                        dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetail.CustomerCode).Value = dr("Account_Code")
                        dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetail.ItemCode).Value = dr("Item_Code")
                        dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetail.ItemDesc).Value = dr("Item_Desc")
                        dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetail.Trans_date).Value = dr("Trans_date")
                        dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetail.ScheduleQty).Value = dr("Schedule_Quantity")
                        dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetail.DispatchQty).Value = dr("Despatch_Qty")
                        dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetail.Serial_No).Value = dr("Serial_No")
                        dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetail.DSNO).Value = dr("DSNO")
                        dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetail.Remarks).Value = ""
                        i += 1
                    Next
                End If
            End With
        Catch ex As Exception
            RaiseException(ex)
        Finally
            If dt IsNot Nothing Then
                dt.Dispose()
            End If
            If sqlCmd IsNot Nothing Then
                sqlCmd.Dispose()
            End If
        End Try
    End Sub
    Private Sub SaveData()
        Try
            Dim strSql As String = ""
            Dim strRemarks As String = ""
            Dim strSerialNo As String = ""
            If dgvInvoiceDetail.Rows.Count > 0 Then
                SqlConnectionclass.BeginTrans()
                With dgvInvoiceDetail
                    For i As Integer = 0 To .Rows.Count - 1
                        strRemarks = dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetail.Remarks).Value
                        strSerialNo = dgvInvoiceDetail.Rows(i).Cells(GridInvoiceDetail.Serial_No).Value

                        strSql = "Update a set a.Schedule_Quantity=b.Schedule_Quantity,a.Despatch_Qty=b.Despatch_Qty, " & _
                        "Closed_Qty=b.Schedule_Quantity - b.Despatch_Qty,PrevClosed_Qty=a.Schedule_Quantity - a.Despatch_Qty,a.remarks='" & strRemarks & "',a.[Ent_dt]=getdate(),a.[Ent_UserId]='" & mP_User & "', " & _
                        "a.[Upd_dt]=getdate(),a.[Upd_UserId]='" & mP_User & "'" & _
                        "from TEMP_Dailymktschedule_CLOSURE a,dailymktschedule b   where a.unit_code='" & gstrUNITID & "' and a.unit_code=b.unit_code " & _
                        "and  a.serial_no=b.serial_no  and a.serial_no='" & strSerialNo & "' and IP_Address='" & gstrIpaddressWinSck & "'"
                        SqlConnectionclass.ExecuteNonQuery(strSql)

                        'strSql = "insert into Dailymktschedule_CLOSURE ([Account_Code],[Trans_date],[Item_code],[Cust_Drgno],[Serial_No], " & _
                        '"[Schedule_Flag],[Schedule_Quantity],[Despatch_Qty],Closed_Qty,[Status],[Ent_dt],[Ent_UserId],[Upd_dt],[Upd_UserId],[DSNO], " & _
                        '"[DSDateTime],[Spare_qty],[Consignee_Code],[FILETYPE],[DOC_NO],[UNIT_CODE],[Remarks]) " & _
                        '"select a.[Account_Code],a.[Trans_date],a.[Item_code],a.[Cust_Drgno],a.[Serial_No],b.[Schedule_Flag], " & _
                        '"b.[Schedule_Quantity],b.[Despatch_Qty],b.[Schedule_Quantity] - b.[Despatch_Qty], b.[Status],getdate(),'" & mP_User & "',getdate(),'" & mP_User & "',b.[DSNO], " & _
                        '"b.[DSDateTime],b.[Spare_qty],b.[Consignee_Code],b.[FILETYPE],b.[DOC_NO],a.[UNIT_CODE],'" & strRemarks & "'  " & _
                        '"from TEMP_Dailymktschedule_CLOSURE a,dailymktschedule b   where a.unit_code='H02' and a.unit_code=b.unit_code " & _
                        '"and  a.serial_no=b.serial_no  and a.serial_no='" & strSerialNo & "' and IP_Address='" & gstrIpaddressWinSck & "'"
                        'SqlConnectionclass.ExecuteNonQuery(strSql)

                        'strSql = "insert into Dailymktschedule_CLOSURE ([Account_Code],[Trans_date],[Item_code],[Cust_Drgno],[Serial_No], " & _
                        '"[Schedule_Flag],[Schedule_Quantity],[Despatch_Qty],[Status],[Ent_dt],[Ent_UserId],[Upd_dt],[Upd_UserId],[DSNO], " & _
                        '"[DSDateTime],[Spare_qty],[Consignee_Code],[FILETYPE],[DOC_NO],[UNIT_CODE],[Remarks]) " & _
                        '"select [Account_Code],[Trans_date],[Item_code],[Cust_Drgno],[Serial_No],[Schedule_Flag],[Schedule_Quantity],[Despatch_Qty],  " & _
                        '"[Status],getdate(),'" & mP_User & "',getdate(),'" & mP_User & "',[DSNO],[DSDateTime],[Spare_qty],  " & _
                        '"[Consignee_Code],[FILETYPE],[DOC_NO],[UNIT_CODE],'" & strRemarks & "' from TEMP_Dailymktschedule_CLOSURE  " & _
                        '" where unit_code='" & gstrUNITID & "'  and serial_no='" & strSerialNo & "' and IP_Address='" & gstrIpaddressWinSck & "'"

                    Next

                    Using sqlCmd As SqlCommand = New SqlCommand
                        With sqlCmd
                            .CommandText = "USP_SAVE_DSCLOSURE"
                            .CommandTimeout = 0
                            .CommandType = CommandType.StoredProcedure
                            .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                            .Parameters.Add("@IP_ADDRESS", SqlDbType.VarChar, 20).Value = gstrIpaddressWinSck
                            .Parameters.AddWithValue("@TYPE", "CLOSURE")
                            SqlConnectionclass.ExecuteNonQuery(sqlCmd)                           
                        End With
                    End Using


                    SqlConnectionclass.CommitTran()

                    Using sqlCmd As SqlCommand = New SqlCommand
                        With sqlCmd
                            .CommandText = "USP_DSCLOSURE_AUTOMAILER"
                            .CommandTimeout = 0
                            .CommandType = CommandType.StoredProcedure
                            .Parameters.Add("@UNIT_CODE", SqlDbType.VarChar, 10).Value = gstrUNITID
                            .Parameters.Add("@IP_ADDRESS", SqlDbType.VarChar, 20).Value = gstrIpaddressWinSck
                            .Parameters.AddWithValue("@TYPE", "CLOSURE")
                            SqlConnectionclass.ExecuteNonQuery(sqlCmd)
                        End With
                    End Using

                    MsgBox("DS Closed successfully . ", MsgBoxStyle.Information, ResolveResString(100))
                    ResetData()
                End With
            Else

                MsgBox("There is no DS to close . ", MsgBoxStyle.Information, ResolveResString(100))
                Exit Sub
            End If
        Catch ex As Exception
            SqlConnectionclass.RollbackTran()
            RaiseException(ex)
        Finally
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        End Try
    End Sub
    Private Sub lvwcustomer_ItemChecked(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckedEventArgs) Handles lvwcustomer.ItemChecked
        Dim Item As System.Windows.Forms.ListViewItem = lvwcustomer.Items(e.Item.Index)
        On Error GoTo ErrHandler
        'Dim stritemmsg As String
        'Dim lvw_Counter As Integer = 0
        'If Bool_First_Check = True Then
        '    Exit Sub
        'End If
        'For lvw_Counter = 0 To Me.lvwcustomer.Items.Count - 1
        '    If Me.lvwcustomer.Items.Item(lvw_Counter).Checked = True Then
        '        If Me.opt_all_item.Checked = False Then
        '            If Item.Checked = True Then
        '                Bool_First_Check = True
        '                Item.Checked = False
        '            Else
        '                Bool_First_Check = True
        '                Item.Checked = True
        '            End If
        '            Bool_First_Check = False
        '            stritemmsg = populateitemlist()
        '            If stritemmsg = "N" Then Exit Sub
        '        End If
        '        ConfigureGridColumn()
        '        Exit Sub
        '    End If
        'Next
        'Exit Sub


        Dim intLoopcount As Short
        Dim intMaxLoop As Short
        Dim inti As Short
        intMaxLoop = lvwcustomer.Items.Count
        For intLoopcount = 0 To intMaxLoop - 1
            If lvwcustomer.Items.Item(intLoopcount).Checked = True Then
                inti = inti + 1
            End If
        Next
        If Me.opt_sel_item.Checked = True Then Call populateitemlist()
        ConfigureGridColumn()
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub opt_all_Customer_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles opt_all_Customer.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo ErrHandler
            Dim strItemsMsg As String
            opt_all_Customer.Checked = True
            lvwcustomer.Enabled = False
            lvwcustomer.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED_LIST_VIEW)
            lvwcustomer.Items.Clear()
            lvwcustomer.Columns.Clear()
            lvwcustomer.GridLines = False
            If (opt_sel_Customer.Checked = True) Then
                Call populateCustomerlist()
            End If
            If (opt_sel_item.Checked = True) Then
                strItemsMsg = populateitemlist()
                If Trim(strItemsMsg) = "N" Then
                    Exit Sub
                End If
                optSearchItemCode.Enabled = True : optSearchItemDes.Enabled = True
                txtSearchItems.Enabled = True : txtSearchItems.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                optSearchItemCode.Checked = False
                optSearchItemDes.Checked = True
                txtSearchItems.Text = ""
            End If
            optSearchCustCode.Checked = False
            optSearchCustCode.Enabled = False
            optSearchCustName.Checked = False
            optSearchCustName.Enabled = False
            txtSearchCustomer.Text = ""
            txtSearchCustomer.Enabled = False
            txtSearchCustomer.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            ConfigureGridColumn()
            Exit Sub
ErrHandler:
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub
    Private Sub opt_all_item_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles opt_all_item.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo ErrHandler
            Me.opt_all_item.Checked = True
            If (Me.opt_all_item.Checked = True) Then
                Me.lvwItems.Enabled = False
                Me.lvwItems.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED_LIST_VIEW)
                Call Me.lvwItems.Items.Clear()
                Me.lvwItems.Columns.Clear()
                Me.lvwItems.GridLines = False
            Else
                Me.lvwItems.Enabled = True
                Me.lvwItems.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            End If
            optSearchItemCode.Checked = False : optSearchItemDes.Checked = False : txtSearchItems.Text = ""
            txtSearchItems.Enabled = False : txtSearchItems.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            optSearchItemCode.Enabled = False : optSearchItemDes.Enabled = False
            ConfigureGridColumn()
            Exit Sub
ErrHandler:
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub
    Private Sub opt_sel_Customer_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles opt_sel_Customer.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo ErrHandler
            Dim strCustMsg As String
            Me.lvwcustomer.Enabled = True
            Me.lvwcustomer.View = System.Windows.Forms.View.Details
            Me.lvwcustomer.CheckBoxes = True
            Me.lvwcustomer.GridLines = True
            Me.lvwcustomer.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            strCustMsg = populateCustomerlist()
            If Trim(strCustMsg) = "N" Then
                Exit Sub
            End If
            If Me.opt_sel_item.Checked = True Then
                Call opt_all_item_CheckedChanged(opt_all_item, New System.EventArgs())
            End If
            optSearchCustCode.Enabled = True : optSearchCustName.Enabled = True : txtSearchCustomer.Enabled = True
            txtSearchCustomer.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            optSearchCustCode.Checked = False
            optSearchCustName.Checked = True
            ConfigureGridColumn()
            Exit Sub
ErrHandler:
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub
    Private Sub search(ByRef lvwListView As System.Windows.Forms.ListView, ByRef txtSearchBox As System.Windows.Forms.TextBox, ByRef optFistOption As System.Windows.Forms.RadioButton, ByRef optSecOption As System.Windows.Forms.RadioButton)
        '*******************************************************************************
        'Author             :   Priti Sharma
        'Argument(s)if any  :   Listview in which item to be searched,name of text box in which item to be searched
        '                   :   and Searching Options.
        'Return Value       :
        'Function           :   Search the entered text in list view & mark it bold.
        'Comments           :   NA
        'Creation Date      :   28/10/2021
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim Intcounter As Short
        With lvwListView
            If optFistOption.Checked = True Then
                For Intcounter = 0 To .Items.Count - 1
                    If .Items.Item(Intcounter).Font.Bold = True Then
                        .Items.Item(Intcounter).Font = VB6.FontChangeBold(.Items.Item(Intcounter).Font, False)
                        .Refresh()
                    End If
                    If .Items.Item(Intcounter).SubItems.Item(1).Font.Bold = True Then
                        .Items.Item(Intcounter).SubItems.Item(1).Font = VB6.FontChangeBold(.Items.Item(Intcounter).SubItems.Item(1).Font, False)
                        .Refresh()
                    End If
                Next
                If Len(txtSearchBox.Text) = 0 Then Exit Sub
                For Intcounter = 0 To .Items.Count - 1
                    If Trim(UCase(Mid(.Items.Item(Intcounter).Text, 1, Len(txtSearchBox.Text)))) = Trim(UCase(txtSearchBox.Text)) Then
                        .Items.Item(Intcounter).Font = VB6.FontChangeBold(.Items.Item(Intcounter).Font, True)
                        Call .Items.Item(Intcounter).EnsureVisible()
                        .Refresh()
                        Exit For
                    End If
                Next
            ElseIf optSecOption.Checked Then
                For Intcounter = 0 To .Items.Count - 1
                    If .Items.Item(Intcounter).Font.Bold = True Then
                        .Items.Item(Intcounter).Font = VB6.FontChangeBold(.Items.Item(Intcounter).Font, False)
                        .Refresh()
                    End If
                    If .Items.Item(Intcounter).SubItems.Item(1).Font.Bold = True Then
                        .Items.Item(Intcounter).SubItems.Item(1).Font = VB6.FontChangeBold(.Items.Item(Intcounter).SubItems.Item(1).Font, False)
                        .Refresh()
                    End If
                Next
                If Len(txtSearchBox.Text) = 0 Then Exit Sub
                For Intcounter = 0 To .Items.Count - 1
                    If Trim(UCase(Mid(.Items.Item(Intcounter).SubItems.Item(1).Text, 1, Len(txtSearchBox.Text)))) = Trim(UCase(txtSearchBox.Text)) Then
                        .Items.Item(Intcounter).Font = VB6.FontChangeBold(.Items.Item(Intcounter).Font, True)
                        Call .Items.Item(Intcounter).EnsureVisible()
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
    Private Sub opt_sel_item_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles opt_sel_item.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo ErrHandler
            Dim strItemsMsg As String
            Me.lvwItems.Enabled = True
            Me.lvwItems.View = System.Windows.Forms.View.Details
            Me.lvwItems.CheckBoxes = True
            Me.lvwItems.GridLines = True
            Me.lvwItems.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            strItemsMsg = populateitemlist()
            If Trim(strItemsMsg) = "N" Then
                Exit Sub
            End If
            optSearchItemCode.Enabled = True : optSearchItemDes.Enabled = True
            txtSearchItems.Enabled = True : txtSearchItems.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
            optSearchItemCode.Checked = False
            optSearchItemDes.Checked = True
            txtSearchItems.Text = ""
            ConfigureGridColumn()
            Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub
    Public Function populateitemlist() As String
        '*******************************************************************************
        'Author             :   Ashuosh Verma
        'Argument(s)if any  :
        'Return Value       :   Returns (Y or N)-- to distinguish weather the function runs successfully or not.
        'Function           :   Populates the Item list for selected schedule date range and customers .
        'Comments           :   NA
        'Creation Date      :   28/11/2006
        '*******************************************************************************
        On Error GoTo ErrHandler
        'Declarations
        Dim objitem As ADODB.Recordset
        Dim strsql As String
        Dim strCustomer As String
        Dim lstItem As System.Windows.Forms.ListViewItem
        Dim lngLoop, lngloop1 As Integer
        Dim intLoopcount As Short
        Dim intMaxLoop As Short
        Dim intMaxCount As Short
        Dim intCount As Short
        Dim intNoRec As Short
        Dim lngrecords As Short
        Dim lngcount1 As Short
        'Initialise the list view
        lvwItems.Items.Clear()
        If Me.opt_sel_Customer.Checked = True Then
            intNoRec = 0 : intMaxCount = 0
            'Checking for checked item type(s)
            intMaxCount = lvwcustomer.Items.Count
            For intCount = 0 To intMaxCount - 1
                If lvwcustomer.Items.Item(intCount).Checked = True Then
                    intNoRec = intNoRec + 1
                End If
            Next
            If intNoRec <= 0 Then
                MsgBox("Select at least one Customer", MsgBoxStyle.Information, ResolveResString(100))
                lvwcustomer.Focus()
                populateitemlist = "N"
                opt_all_item.Checked = True
                opt_sel_item.Checked = False
                lvwItems.Enabled = False
                Exit Function
            End If
        End If

      

        With lvwItems
            .Sort()
            ListViewColumnSorter.SortListView(lvwItems, 0, SortOrder.Ascending)
            .LabelEdit = False
            .CheckBoxes = True
            .View = System.Windows.Forms.View.Details
            .Columns.Clear()
            .Columns.Insert(0, "", "Item code", -2)
            .Columns.Insert(1, "", "Description", -2)
            If opt_all_item.Checked = True Then
                .Enabled = False
            Else
                .Enabled = True
            End If
        End With


        strsql = "select Distinct(a.Item_code) as Item_code,c.Description from DailyMktSchedule a ( nolock) ,customer_mst b(nolock) ,Item_mst c (Nolock) " & _
        "where  a.UNIT_CODE=b.UNIT_CODE and a.Account_Code=b.Customer_Code  and a.unit_code=c.unit_code and " & _
        "a.item_code=c.item_code and a.unit_code='" & gstrUNITID & "' and c.Status ='A' and c.Hold_Flag <> 1  and isnull(manual_ds_closure,0)=0  and a.Trans_date between '" & VB6.Format(Me.dtpDateFrom.Value, "dd/mmm/yyyy") & "' and '" & VB6.Format(Me.dtpDateTo.Value, "dd/mmm/yyyy") & "' "
        If opt_sel_Customer.Checked = True Then
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
                strsql = strsql & " and a.Account_Code in (" & strCustomer & ")"
            End If
        End If
        objitem = New ADODB.Recordset
        Call objitem.Open(strsql, mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        For lngLoop = 1 To objitem.Fields.Count
            Me.lvwItems.Columns.Add(objitem.Fields(lngLoop - 1).Name)
        Next
        lngrecords = objitem.RecordCount
        If Not ((objitem.BOF) And (objitem.EOF)) Then
            For lngcount1 = 1 To lngrecords
                lstItem = Me.lvwItems.Items.Add(Trim(objitem.Fields("Item_Code").Value))
                For lngLoop = 0 To objitem.Fields.Count - 2
                    If lstItem.SubItems.Count > lngLoop + 1 Then
                        lstItem.SubItems(lngLoop + 1).Text = objitem.Fields(lngLoop + 1).Value
                    Else
                        lstItem.SubItems.Insert(lngLoop + 1, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, objitem.Fields(lngLoop + 1).Value))
                    End If
                Next
                objitem.MoveNext()
            Next
            objitem.Close()
            objitem = Nothing
        End If
        objitem = Nothing
        Me.lvwItems.Columns.Item(0).Width = 100
        Me.lvwItems.Columns.Item(1).Width = 400
        populateitemlist = "Y"
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        populateitemlist = "N"
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Function
    End Function
    Private Sub optSearchCustCode_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optSearchCustCode.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo ErrHandler
            txtSearchCustomer.Text = ""
            txtSearchCustomer.Focus()
            Exit Sub
ErrHandler:
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub
    Private Sub optSearchCustName_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optSearchCustName.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo ErrHandler
            txtSearchCustomer.Text = ""
            txtSearchCustomer.Focus()
            Exit Sub
ErrHandler:
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub
    Private Sub optSearchItemCode_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optSearchItemCode.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo ErrHandler
            txtSearchItems.Text = ""
            txtSearchItems.Focus()
            Exit Sub
ErrHandler:
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub
    Private Sub optSearchItemDes_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optSearchItemDes.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo ErrHandler
            txtSearchItems.Text = ""
            txtSearchItems.Focus()
            Exit Sub
ErrHandler:
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub
    Private Sub txtSearchCustomer_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSearchCustomer.TextChanged
        On Error GoTo ErrHandler
        Call search(lvwcustomer, txtSearchCustomer, optSearchCustCode, optSearchCustName)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtSearchItems_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSearchItems.TextChanged
        On Error GoTo ErrHandler
        Call search(lvwItems, txtSearchItems, optSearchItemCode, optSearchItemDes)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub

    
    Private Sub btnDisplayDetails_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDisplayDetails.Click
        execSP()
        'FillDS()
    End Sub

    Private Sub btnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

   
    Private Sub btnLock_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnLock.Click
        Try
            SaveData()
        Catch ex As Exception
            RaiseException(ex)
        Finally
            ChangeMousePointer(ObjectsEnum.obj_Screen, , Cursors.Default)
        End Try
    End Sub
    Private Sub ResetData()
        Call InitializeControls()
        ConfigureGridColumn()
        dtpDateFrom.Format = DateTimePickerFormat.Custom
        dtpDateFrom.CustomFormat = gstrDateFormat
        dtpDateFrom.Value = GetServerDate()

        dtpDateTo.Format = DateTimePickerFormat.Custom
        dtpDateTo.CustomFormat = gstrDateFormat
        dtpDateTo.Value = GetServerDate()
        Me.opt_all_Customer.Checked = True
        Me.opt_all_item.Checked = True

        Me.optSearchCustCode.Checked = True
        Me.optSearchItemCode.Checked = True
    End Sub
    Private Sub btnClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClear.Click
        ResetData()
    End Sub

    Private Sub lvwItems_ItemChecked(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckedEventArgs) Handles lvwItems.ItemChecked
        ConfigureGridColumn()
    End Sub
End Class