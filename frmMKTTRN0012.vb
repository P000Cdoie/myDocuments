Option Strict Off
Option Explicit On
Friend Class frmMKTTRN0012
	Inherits System.Windows.Forms.Form
	'===================================================================================
	' (c) MothersonSumi INfotech & Design Ltd. All rights reserverd.
	' File Name         :   FRMMKTTRN0012.frm
	' Function          :   Schedule Rollover
	' Created By        :   Nisha Rai
	' Created On        :   19 Aug, 2003 eve
    ' Revision History  :26/08/2003
    'Modified By Ajay Shukla on 12/May/2011 for multi unit change
	'===================================================================================
	Dim mName As String
	Dim mintFormIndex As Short
	Dim SelectedItem As String
	Dim startdate As String
	Dim endDate As String
	Dim strDate As String
	Dim strMonth As String
	Dim strYear As String
	Dim strlastday As String
	Dim monthYear As String
	Dim intresult As Double
	Dim mbolCheck As Boolean
    Dim mblnUnload As Boolean ' to unload form before loading

    Private Sub cmdClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClose.Click
        On Error GoTo Err_Handler
        CmdSales_ButtonClick(CmdSales, New UCActXCtl.UCfraRepCmd.ButtonClickEventArgs((UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE)))
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub CmdSales_ButtonClick(ByVal eventSender As System.Object, ByVal eventArgs As UCActXCtl.UCfraRepCmd.ButtonClickEventArgs) Handles CmdSales.ButtonClick
        Dim intLoopCounter As Short
        Dim intmaxLoop As Short
        Dim intItemCount As Short
        Dim strSQL As String
        Dim strPrevYearMonth As String
        Dim strPrevMonth As String
        Dim strPrevYear As String
        Dim cmdproc As ADODB.Command
        On Error GoTo Err_Handler
        If eventArgs.Button = UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE Then
            Me.Close()
            Exit Sub
        End If
        Select Case eventArgs.Button
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                Me.Close()
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT
                If optAccountCode(1).Checked = True Then
                    intmaxLoop = lvwCust.Items.Count
                    intItemCount = 0
                    For intLoopCounter = 0 To intmaxLoop - 1
                        If lvwCust.Items.Item(intLoopCounter).Checked = True Then
                            intItemCount = intItemCount + 1
                        End If
                    Next
                    If intItemCount = 0 Then
                        MsgBox("Select atleast one customer.", MsgBoxStyle.Information, "eMPro")
                        lvwCust.Focus()
                        Exit Sub
                    End If
                    intmaxLoop = lvwCust.Items.Count
                    mP_Connection.Execute("Delete from SchRollCustCode where unit_code='" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    For intLoopCounter = 0 To intmaxLoop - 1
                        If lvwCust.Items.Item(intLoopCounter).Checked = True Then
                            mP_Connection.Execute("Insert into SchRollCustCode (CustomerCode, UNIT_CODE) values('" & lvwCust.Items.Item(intLoopCounter).Text & "','" & gstrUNITID & "')", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                        End If
                    Next
                Else
                    strPrevMonth = CStr(Month(GetServerDate()) - 1)
                    If Month(GetServerDate()) = 1 Then
                        strPrevYear = CStr(Year(System.DateTime.FromOADate(GetServerDate().ToOADate - 1)))
                    Else
                        strPrevYear = CStr(Year(GetServerDate()))
                    End If
                    If CDbl(strPrevMonth) > 9 Then
                        strPrevYearMonth = strPrevYear & strPrevMonth
                    Else
                        strPrevYearMonth = strPrevYear & "0" & strPrevMonth
                    End If
                    strSQL = ""
                    strSQL = "insert into SchRollCustCode select distinct account_code, UNIT_CODE from dailymktSchedule where datepart(mm,trans_date) =  datepart(mm,getdate())-1 "
                    strSQL = Trim(strSQL) & " and status = 1 and UNIT_CODE='" & gstrUNITID & "' and datepart (yyyy,trans_date) = " & strPrevYear & " and Account_code not in (select distinct CustomerCode from schRollOverDetails where UNIT_CODE='" & gstrUNITID & "' and datepart(mm,Rollover_Date)  = datepart(mm,getdate()) and datepart(yyyy,Rollover_Date)  = datepart(yyyy,getdate()))"
                    strSQL = Trim(strSQL) & vbCrLf & " UNION select distinct account_code, UNIT_CODE from MonthlyMktSchedule where year_Month = '" & strPrevYearMonth
                    strSQL = Trim(strSQL) & "' and status =1 and UNIT_CODE='" & gstrUNITID & "' and Account_code not in (select distinct CustomerCode from schRollOverDetails where UNIT_CODE='" & gstrUNITID & "' and datepart(mm,Rollover_Date)  = datepart(mm,getdate()) and datepart(yyyy,Rollover_Date)  = datepart(yyyy,getdate()))"
                    mP_Connection.Execute(strSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                End If
                cmdproc = New ADODB.Command
                cmdproc.let_ActiveConnection(mP_Connection)
                cmdproc.CommandTimeout = 0
                cmdproc.CommandType = ADODB.CommandTypeEnum.adCmdStoredProc
                cmdproc.CommandText = "ScheduleRollOver"
                cmdproc.Parameters.Append(cmdproc.CreateParameter("@UNITCODE", ADODB.DataTypeEnum.adVarChar, ADODB.ParameterDirectionEnum.adParamInput, 10, gstrUNITID))
                cmdproc.Execute(, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                cmdproc = Nothing
                System.Windows.Forms.Application.DoEvents()
                mP_Connection.Execute("insert into schRollOverDetails select distinct customercode,getdate(),UNIT_CODE from SchRollCustCode WHERE UNIT_CODE='" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                mP_Connection.Execute("DELETE FROM SchRollCustCode where unit_code='" & gstrUNITID & "'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                MsgBox("Data Transfered SuccessFully.", MsgBoxStyle.Information, "eMPro")
                optAccountCode(0).Checked = True
                Exit Sub
        End Select
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdTransfer_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdTransfer.Click
        On Error GoTo Err_Handler
        cmdTransfer.Enabled = False
        CmdSales_ButtonClick(CmdSales, New UCActXCtl.UCfraRepCmd.ButtonClickEventArgs((UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT)))
        cmdTransfer.Enabled = True
        Exit Sub
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub ctlFormHeader1_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ctlFormHeader1.Click
        On Error GoTo ErrHandler
        Call ShowHelp("HLP" & Mid(Me.Name, 4, Len(Me.Name)) & ".htm")
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0012_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo ErrHandler
        If mblnUnload = True Then Me.Close()
        mdifrmMain.CheckFormName = mintFormIndex
        frmModules.NodeFontBold(Tag) = True
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0012_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Tag) = False
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0012_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0 Then Call ctlFormHeader1_ClickEvent(ctlFormHeader1, New System.EventArgs()) : Exit Sub
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0012_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If KeyAscii = 13 Then System.Windows.Forms.SendKeys.Send("{Tab}")
        GoTo EventExitSub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub frmMKTTRN0012_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo ErrHandler
        mblnUnload = False
        mintFormIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader1.Tag)
        Call FillLabelFromResFile(Me) 'To Fill label description from Resource file
        Call FitToClient(Me, fraSalesSchedule, ctlFormHeader1, lblUploadCmd) 'To fit the form in the MDI
        cmdTransfer.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(lblUploadCmd.Left) + 70)
        cmdTransfer.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(lblUploadCmd.Top) + 50)
        cmdClose.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(cmdTransfer.Left) + VB6.PixelsToTwipsX(cmdTransfer.Width) + 10)
        cmdClose.Top = cmdTransfer.Top
        Call EnableControls(True, Me) 'To Disable controls
        Call Initialize()
        If mblnUnload = True Then Exit Sub
        lvwCust.Enabled = False
        optAccountCode(0).Checked = True
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0012_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error GoTo ErrHandler
        mdifrmMain.RemoveFormNameFromWindowList = mintFormIndex
        frmModules.NodeFontBold(Tag) = False
        Me.Dispose()
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub lvwCust_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles lvwCust.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        If KeyAscii = System.Windows.Forms.Keys.Return Then
            cmdTransfer.Focus()
        End If
        GoTo EventExitSub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub optAccountCode_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optAccountCode.CheckedChanged
        If eventSender.Checked Then
            Dim Index As Short = optAccountCode.GetIndex(eventSender)
            On Error GoTo ErrHandler
            If optAccountCode(0).Checked = True Then
                lvwCust.Enabled = False
                mName = "Sales Schedule [Customer-Wise]"
                lvwCust.GridLines = False
                lvwCust.Items.Clear()
                lvwCust.Columns.Clear()
                lvwCust.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED_LIST_VIEW)
                Me.optSearchByCode.Enabled = False
                Me.optSearchByDesc.Enabled = False
                Me.txtSearch.Enabled = False : txtSearch.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            ElseIf optAccountCode(1).Checked = True Then
                lvwCust.Enabled = True
                Call AddValuetoList()
                mName = "Sales Schedule [Customer-Wise]"
                lvwCust.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                Me.optSearchByCode.Enabled = True
                Me.optSearchByDesc.Enabled = True
                Me.txtSearch.Enabled = True : txtSearch.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                Me.optSearchByCode.Checked = True
            End If
            Exit Sub
ErrHandler:
            gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub
    Public Sub Initialize()
        On Error GoTo ErrHandler
        Call AddValuetoList()
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Public Sub AddValuetoList()
        On Error GoTo ErrHandler
        Dim strSQL As String
        Dim rsAccountcode As ClsResultSetDB
        Dim strPrevYearMonth As String
        Dim strPrevMonth As String
        Dim strPrevYear As String
        Dim intLoopcount As Short
        Dim intRowCount As Short
        Dim lstItem As System.Windows.Forms.ListViewItem
        Me.lvwCust.View = System.Windows.Forms.View.Details
        Me.lvwCust.GridLines = True
        Me.lvwCust.CheckBoxes = True
        strPrevMonth = CStr(Month(GetServerDate()) - 1)
        If Month(GetServerDate()) = 1 Then
            strPrevYear = CStr(Year(System.DateTime.FromOADate(GetServerDate().ToOADate - 1)))
        Else
            strPrevYear = CStr(Year(GetServerDate()))
        End If
        If CDbl(strPrevMonth) > 9 Then
            strPrevYearMonth = strPrevYear & strPrevMonth
        Else
            strPrevYearMonth = strPrevYear & "0" & strPrevMonth
        End If
        strSQL = "select distinct a.account_code,b.cust_Name from dailymktSchedule a,Customer_Mst b where datepart(mm,trans_date) =  datepart(mm,getdate())-1 "
        strSQL = Trim(strSQL) & " and status = 1 and datepart (yyyy,trans_date) = " & strPrevYear & " and a.Account_code = b.Customer_code AND a.UNIT_CODE = b.UNIT_CODE and  a.UNIT_CODE='" & gstrUNITID & "' and ((isnull(b.deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= b.deactive_date))  and a.Account_code not in (select distinct CustomerCode from schRollOverDetails where UNIT_CODE='" & gstrUNITID & "' AND datepart(mm,Rollover_Date)  = datepart(mm,getdate()) and datepart(yyyy,Rollover_Date)  = datepart(yyyy,getdate()))"
        strSQL = Trim(strSQL) & vbCrLf & " UNION select distinct a.account_code,b.cust_name from MonthlyMktSchedule a,customer_Mst b where year_Month = '" & strPrevYearMonth
        strSQL = Trim(strSQL) & "' and status =1 and a.Account_code = b.Customer_code AND a.UNIT_CODE = b.UNIT_CODE and  a.UNIT_CODE='" & gstrUNITID & "' and ((isnull(b.deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= b.deactive_date)) and a.Account_code not in (select distinct CustomerCode from schRollOverDetails where UNIT_CODE='" & gstrUNITID & "' AND datepart(mm,Rollover_Date)  = datepart(mm,getdate()) and datepart(yyyy,Rollover_Date)  = datepart(yyyy,getdate())) order by account_code"
        rsAccountcode = New ClsResultSetDB
        rsAccountcode.GetResult(strSQL)
        intRowCount = rsAccountcode.GetFieldCount
        rsAccountcode.MoveFirst()
        If intRowCount > 0 Then
            For intLoopcount = 0 To intRowCount - 1
                Me.lvwCust.Columns.Add(rsAccountcode.GetFieldName(intLoopcount))
            Next intLoopcount
            If intRowCount > 0 Then
                Do Until rsAccountcode.EOFRecord
                    lstItem = Me.lvwCust.Items.Add(rsAccountcode.GetValue("Account_Code"))
                    For intLoopcount = 0 To rsAccountcode.GetFieldCount - 2
                        If lstItem.SubItems.Count > intLoopcount + 1 Then
                            lstItem.SubItems(intLoopcount).Text = rsAccountcode.GetValue("Cust_Name")
                        Else
                            lstItem.SubItems.Insert(intLoopcount + 1, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, rsAccountcode.GetValue("Cust_Name")))
                        End If
                    Next intLoopcount
                    rsAccountcode.MoveNext()
                Loop
            End If
            lvwCust.Columns.Clear()
            lvwCust.Columns.Insert(0, "", "Customer Code", -2)
            lvwCust.Columns.Insert(1, "", "Description", -2)
            lvwCust.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            Me.lvwCust.Columns.Item(0).Width = 100
            Me.lvwCust.Columns.Item(1).Width = 400
            mblnUnload = False
        Else
            mblnUnload = True
        End If
        rsAccountcode.ResultSetClose()
        rsAccountcode = Nothing
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub optAccountCode_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles optAccountCode.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        Dim Index As Short = optAccountCode.GetIndex(eventSender)
        On Error GoTo ErrHandler
        If KeyCode = 13 And Me.optAccountCode(1).Checked = True Then
            System.Windows.Forms.SendKeys.Send("+{tab}")
            System.Windows.Forms.SendKeys.Send("+{tab}")
            Me.lvwCust.Focus()
        End If
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub optAccountCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles optAccountCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        Dim Index As Short = optAccountCode.GetIndex(eventSender)
        On Error GoTo ErrHandler
        If KeyAscii = 13 Then
            If Me.optAccountCode(0).Checked = True Then
                Me.cmdTransfer.Focus()
            Else
                Me.optAccountCode(1).Focus()
            End If
            If Me.optAccountCode(1).Checked = True Then
                Me.lvwCust.Focus()
            End If
        End If
        GoTo EventExitSub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Public Sub FuncSearchCust()
        On Error GoTo ErrHandler
        Dim lvwFoundText As System.Windows.Forms.ListViewItem ' FoundItem variable.
        On Error GoTo ErrHandler
        lvwFoundText = SearchText((txtSearch.Text), optSearchByDesc, lvwCust, "1")
        If lvwFoundText Is Nothing Then ' If no match,
            Exit Sub
        Else
            lvwFoundText.EnsureVisible() ' Scroll ListView to show found ListItem.
            lvwFoundText.Selected = True ' Select the ListItem.
            lvwCust.Enabled = True
            If Len(txtSearch.Text) > 0 Then lvwFoundText.Font = VB6.FontChangeBold(lvwFoundText.Font, True)
        End If
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub optSearchByCode_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optSearchByCode.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo ErrHandler
            Me.txtSearch.Text = ""
            Exit Sub
ErrHandler:
            gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub
    Private Sub optSearchByDesc_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles optSearchByDesc.CheckedChanged
        If eventSender.Checked Then
            On Error GoTo ErrHandler
            Me.txtSearch.Text = ""
            Exit Sub
ErrHandler:
            gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        End If
    End Sub
    Private Sub TxtSearch_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSearch.TextChanged
        On Error GoTo ErrHandler
        Call FuncSearchCust()
        Exit Sub
ErrHandler:
        gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
End Class