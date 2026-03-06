Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmMKTTRN0026
	Inherits System.Windows.Forms.Form
	''----------------------------------------------------
	''(C) 2001 MIND, All rights reserved
	''Name of module - FRMMKTTRN0026.frm
	''Created by     - Tapan Jain
	''Created Date   - 12-12-2002
	''Description    - Monthly Marketing Schedule
	''Revised date   -
	''Revision History - revised by nisha on 06/10/2003
	''Revised date   - 02 Sep 2004
	''Revision History - Revised by Sourabh For FIFO Implementation(DSTracking-10623)
	'-----------------------------------------------------------------------
	' Revision Date     :28/03/2006
	' Revision By       :Davinder Singh
	' Issue ID          :17378
	' Revision History  :1) To Save data also in the Forecast_mst
	'                    2) To restrict the user from entering the Qty. of Item in decimal places
	'                       if measurement_mst does not allow decimal places for measurement code for that Item
	'                    3) To Rollback the whole Transaction if user does not press the main SAVE or UPDATE button
	'-----------------------------------------------------------------------
	' Revision Date     :22/06/2006
	' Revision By       :Davinder Singh
	' Issue ID          :18166
	' Revision History  :User was unable to select month other than the current month
	'                    in the Schedule start date DTP
	'-----------------------------------------------------------------------
	' Revision Date     :17/07/07
	' Revision By       :Manoj Kr. Vaish
	' Issue ID          :20665
	' Revision History  :While Selecting Account Code in daily/Monthly Schedule giving message
	'                   [Invalid Customer Code OR Manual Schedule entry not Allowed !] ,if the account code exist in Sales order.
	'----------------------------------------------------------------------------------------------------------------------------------------
    'Revised By        - Manoj Vaish
    'Revision Date     - 06 Mar 2009
    'Issue ID          - eMpro-20090227-27987
    'Revision History  - Consignee Changes for commercial invoice at Mate Units
    'Modified By Nitin Mehta on 12 May 2011
    'Modified to support MultiUnit functionality
    '-----------------------------------------------------------------------------
    ' Revision Date     :02 jan 2013
    ' Revision By       :Prashant Rajpal
    ' Issue ID          :10325676 
    '----------------------------------------------------------------------------------------------------------------------------------------

    Dim mstrServerDate As String
    Dim strHelp() As String 'string array for showing values
    Dim mblnSingleSave As Boolean
    Dim mintCurrentNo As Short
    Dim mintMaxCurrentNo As Short
    Dim mintPrevMarkRow As Short
    Dim mblnDirty As Boolean
    Dim mintFormIndex As Short
    Dim mblnStatus As Boolean ''Declared by Davinder to assign status if item level save is clicked or not
    Dim mRow As Integer ''Declared by Davinder to assign Row no of the grid
    Dim mCol As Integer ''Declared by Davinder to assign Col no of the grid
    Dim valid_line_flag As Boolean = False
    Dim blnCheckMsg As Boolean = False
    Dim blnCheckMsg1 As Boolean = False
    Dim blnCheckMsg2 As Boolean = False
    Private Enum enumWholeEntryGrid
        COLUMN_CUSTDRGNO = 1
        COLUMN_ITEMCODE = 2
        COLUMN_CUSTDRGDESC = 3
        COLUMN_SCHEDULEQTY = 4
        COLUMN_DESPATCHQTY = 5
        COLUMN_REVISION_NO = 6
        COLUMN_REMARKS = 7
        COLUMN_YEAR_MONTH = 8
    End Enum
    Private Enum enumSingleEntryGrid
        COLUMN_MONTH_YEAR = 1
        COLUMN_SCHEDULE_QTY = 2
        COLUMN_REVISIONCOUNT = 3
        COLUMN_DESPATCH_QTY = 4
        COLUMN_REMARKS = 5
    End Enum
    Private Sub cmdCustomerHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCustomerHelp.Click
        '----------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/12/2002
        'Arguments           - None
        'Return Value        - None
        'Function            - To Show Help Form
        '----------------------------------------------------
        On Error GoTo ErrHandler
        With ctlEMPHelpMktSchedule
            .CreateDSN(gstrCONNECTIONSERVER, gstrCONNECTIONDESCRIPTION, gstrCONNECTIONDSN, gstrCONNECTIONDATABASE)
            .ConnectAsUser = gstrCONNECTIONUSER
            .ConnectThroughDSN = gstrCONNECTIONDSN
            .ConnectWithPWD = gstrCONNECTIONPASSWORD
        End With
        'Chnaging the mouse pointer
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.WaitCursor)
        If CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            ' Add by Distinct by Sandeep
            '''Changes done By Ashutosh on 02 jul 2007,Issue Id:20418
            strHelp = ctlEMPHelpMktSchedule.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select distinct a.Account_code, b.Cust_Name from  MonthlyMktSchedule a, customer_mst b where a.account_code=b.Customer_Code and a.UNIT_CODE=b.UNIT_CODE and a.UNIT_CODE = '" & gstrUNITID & "' and ((isnull(b.deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= b.deactive_date)) ", "Customer Code Help", 2)
            '''strHelp = ctlEMPHelpMktSchedule.ShowList("SELECT Distinct Account_code , Prty_Name FROM MonthlyMktSchedule ,Gen_PartyMaster WHERE Account_Code=Prty_PartyID", "Customer Code Help", 2)
        Else
            ' Add by Distinct by Sandeep
            strHelp = ctlEMPHelpMktSchedule.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "SELECT Distinct a.Account_code, b.Cust_Name FROM Cust_Ord_Hdr a,customer_mst b WHERE a.account_code=b.Customer_Code and a.UNIT_CODE=b.UNIT_CODE and b.AllowExcessSchedule=1 and a.UNIT_CODE = '" & gstrUNITID & "' and ((isnull(b.deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= b.deactive_date)) ", "Customer Code Help", 2)
            '''strHelp = ctlEMPHelpMktSchedule.ShowList("SELECT Distinct Account_code , Prty_Name FROM Cust_Ord_Hdr ,Gen_PartyMaster WHERE Account_Code=Prty_PartyID", "Customer Code Help", 2)
        End If
        'Chnaging the mouse pointer
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        If UBound(strHelp) <> -1 Then
            If strHelp(0) <> "0" Then
                If strHelp(0) <> "" Then
                    txtCustomerCode.Text = Trim(strHelp(0))
                    lblCustomerDesc.Text = Trim(strHelp(1))
                    Call txtCustomerCode_Validating(txtCustomerCode, New System.ComponentModel.CancelEventArgs(False))
                End If
            Else
                MsgBox("No Record Available", MsgBoxStyle.Information, "empower")
                txtCustomerCode.Focus()
            End If
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub CmdGrpMktSchedule_ButtonClick(ByVal eventSender As System.Object, ByVal eventArgs As UCActXCtl.UCbtndgrp.ButtonClickEventArgs) Handles CmdGrpMktSchedule.ButtonClick
        '----------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/12/2002
        'Arguments           - None
        'Return Value        - None
        'Function            - To Add Functionality of ADD/EDIT/UPDATE
        '----------------------------------------------------
        On Error GoTo ErrHandler
        Select Case eventArgs.Button
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD 'Add Record
                mblnStatus = False
                Call RefreshForm()
                cmdSingleEntrySave.Enabled = True
                DTPSheduleStartDate.Value = VB6.Format(mstrServerDate, "mm/yyyy")
                DTPSheduleEndMonth.Value = VB6.Format(mstrServerDate, "mm/yyyy")
                mintCurrentNo = 1
                mintPrevMarkRow = 1
                mblnDirty = False
                Me.CmdGrpMktSchedule.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD) = False
                mP_Connection.BeginTrans()
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE 'Save Record
                If mblnStatus Then
                    If DisablePrimaryField(True) Then
                        If mblnSingleSave = False Then
                            If CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                                If ValidateinCaseWholeGrid() = False Then
                                    Exit Sub
                                End If
                            End If
                            If InsertdatainCaseofWholeGrid() = False Then
                                Exit Sub
                            End If
                        End If
                        mP_Connection.CommitTrans()
                        MsgBox("Transaction Completed Successfully", MsgBoxStyle.Information, "empower")
                        Call RefreshForm()
                        CmdGrpMktSchedule.Revert()
                        CmdGrpMktSchedule.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                        CmdGrpMktSchedule.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                        CmdGrpMktSchedule.Enabled(2) = False
                        gblnCancelUnload = False
                        gblnFormAddEdit = False
                    End If
                Else
                    MsgBox("Save the Quantity on Item Level First", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                    Me.cmdSingleEntrySave.Focus()
                    Exit Sub
                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL ' Cancel Click
                Call frmMKTTRN0026_KeyPress(Me, New System.Windows.Forms.KeyPressEventArgs(Chr(System.Windows.Forms.Keys.Escape)))
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT 'Clich on edit
                mblnStatus = False
                If mblnSingleSave = True Then
                    cmdSingleEntrySave.Enabled = True
                    mblnDirty = False
                    mintCurrentNo = 1
                    If DisablePrimaryField(True) Then
                        Call SetTextBoxValue(mintCurrentNo)
                        Call FillSingleGrid("EDIT")
                    Else
                        CmdGrpMktSchedule.Revert()
                        Exit Sub
                    End If
                Else
                    mblnDirty = False
                    mintCurrentNo = 1
                    If DisablePrimaryField(True) Then
                        EnableWholeGridinEditMode()
                    Else
                        CmdGrpMktSchedule.Revert()
                        Exit Sub
                    End If
                End If
                mP_Connection.BeginTrans()
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                'To Verify Close Operation
                If Me.CmdGrpMktSchedule.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                    mP_Connection.RollbackTrans()
                End If
                Me.Close()
        End Select
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdSingleEntrySave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSingleEntrySave.Click
        '----------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/12/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Save data in Case of Single entry Row.
        '----------------------------------------------------
        On Error GoTo ErrHandler
        If CmdGrpMktSchedule.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            With Me.spdSingleEntryGrid
                '''Changes done by Ashutosh on 24-10-2006,issue Id:18623
                If Len(Trim(txtItemCode.Text)) = 0 Then Exit Sub
                '''Changes for Issue Id:18623 end here.
                If DecimalAllowed(mRow, Trim(txtItemCode.Text)) Then
                    .Row = mRow
                    .Col = enumSingleEntryGrid.COLUMN_SCHEDULE_QTY
                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                    .Focus()
                    Exit Sub
                End If
            End With
            If Not SaveDataInTable() Then Exit Sub
            Call MarkRowAsSaved(mintCurrentNo)
            Call SaveDataInWholeGrid(mintCurrentNo)
            mintCurrentNo = mintCurrentNo + 1
            If mintCurrentNo <= mintMaxCurrentNo Then
                Call SetTextBoxValue(mintCurrentNo)
            Else
                mintCurrentNo = 1
                Call SetTextBoxValue(mintCurrentNo)
            End If
            If CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                Call FillSingleGrid("ADD")
            Else
                Call FillSingleGrid("EDIT")
            End If
        End If
        mblnStatus = True
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdSingleEntrySkip_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSingleEntrySkip.Click
        '----------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/12/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Show Help Form
        '----------------------------------------------------
        On Error GoTo ErrHandler
        mintCurrentNo = mintCurrentNo + 1
        If mintCurrentNo <= mintMaxCurrentNo Then
            Call SetTextBoxValue(mintCurrentNo)
        Else
            mintCurrentNo = 1
            Call SetTextBoxValue(mintCurrentNo)
        End If
        If CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            Call FillSingleGrid("ADD")
        ElseIf CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
            Call FillSingleGrid("EDIT")
        Else
            Call FillSingleGrid("VIEW")
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub ctlFormHeader1_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ctlFormHeader1.Click
        '----------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/12/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Show Help on F4 Click (empower Help)
        '----------------------------------------------------
        On Error GoTo ErrHandler
        Call ShowHelp("HLP" & Mid(Me.Name, 4, Len(Me.Name)) & ".htm") '("pddrep_listofitems.htm")
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub DTPSheduleEndMonth_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DTPSheduleEndMonth.KeyDown
        '----------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/12/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Set focus on Next Control on Enter Key Press
        '----------------------------------------------------
        If e.KeyCode = System.Windows.Forms.Keys.Return Then System.Windows.Forms.SendKeys.Send(("{tab}"))
    End Sub
    Private Sub DTPSheduleEndMonth_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DTPSheduleEndMonth.ValueChanged
        '----------------------------------------------------
        'Author              - Davinder Singh
        'Create Date         - 30/03/2006
        'Function            - To not allow the schedule end date less than the schedule start date
        '----------------------------------------------------
        On Error GoTo ErrHandler
        If Me.DTPSheduleEndMonth.Value < Me.DTPSheduleStartDate.Value Then
            Me.DTPSheduleEndMonth.Value = Me.DTPSheduleStartDate.Value
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub DTPSheduleStartDate_ValueChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles DTPSheduleStartDate.ValueChanged
        '----------------------------------------------------
        'Author              - Davinder Singh
        'Create Date         - 30/03/2006
        'Function            - To Set the End of schedule date equal to the start of the schedule date
        'Revised By          - Davinder Singh
        'Revised Date        - 22-06-2006
        'Issue ID            - 18166
        'History             - User was unable to select month other than the current month in the DTP
        '----------------------------------------------------
        On Error GoTo ErrHandler
        If CDate("01" & VB.Right(VB6.Format(DTPSheduleStartDate.Value, "dd/mm/yyyy"), 8)) < CDate("01" & VB.Right(VB6.Format(mstrServerDate, "dd/mm/yyyy"), 8)) Then
            DTPSheduleStartDate.Value = mstrServerDate
        End If
        If DTPSheduleStartDate.Value > DTPSheduleEndMonth.Value Then
            DTPSheduleEndMonth.Value = DTPSheduleStartDate.Value
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub DTPSheduleStartDate_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles DTPSheduleStartDate.KeyDown
        '----------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/12/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Set focus on Next Control on Enter Key Press
        '----------------------------------------------------
        If eventArgs.KeyCode = System.Windows.Forms.Keys.Return Then System.Windows.Forms.SendKeys.Send(("{tab}"))
    End Sub
    Private Sub frmMKTTRN0026_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        '----------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/12/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - For Required code on Form Activate
        '----------------------------------------------------
        On Error GoTo ErrHandler
        mdifrmMain.CheckFormName = mintFormIndex
        frmModules.NodeFontBold(Tag) = True
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0026_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        '----------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/12/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - For Required code on Form Deactivate
        '----------------------------------------------------
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Tag) = False
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0026_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '----------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/12/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Call help on F4 Press
        '----------------------------------------------------
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0 Then Call ctlFormHeader1_ClickEvent(ctlFormHeader1, New System.EventArgs()) : Exit Sub
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0026_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '----------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/12/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To add Code for Escape Key Press
        '----------------------------------------------------
        Dim enmValue As eMPowerFunctions.ConfirmWindowReturnEnum
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Escape
                'If user press the ESC Key ,the Form will be unloaded
                If Me.CmdGrpMktSchedule.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                    enmValue = ConfirmWindow(10053, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION)
                    If enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                        Call RefreshForm()
                        CmdGrpMktSchedule.Revert()
                        gblnCancelUnload = False : gblnFormAddEdit = False
                        '''Changed done By Ashutosh on 23-10-2006,Issue Id:18623
                        mP_Connection.RollbackTrans()
                        '''Changes for Issue Id:18623 end here.
                    Else
                        Me.CmdGrpMktSchedule.Focus()
                    End If
                End If
        End Select
        GoTo EventExitSub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        GoTo EventExitSub
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub frmMKTTRN0026_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        '----------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/12/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - For Required code on Form Load
        '----------------------------------------------------
        Dim varIsRowWiseSaving() As Object
        On Error GoTo ErrHandler
        '-------------------------------------------------------
        mintFormIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader1.Tag)
        Call FitToClient(Me, Frame1, ctlFormHeader1, CmdGrpMktSchedule, 500) 'To fit the form in the MDI
        varIsRowWiseSaving = GetFieldsValues("SELECT RowSchSave FROM Sales_Parameter where UNIT_CODE='" & gstrUNITID & "'", 1)
        If varIsRowWiseSaving(0) Then
            mblnSingleSave = True
        Else
            mblnSingleSave = False
            spdSingleEntryGrid.Visible = False
            fraSavingFrame.Visible = False
            With spdWholeEntryGrid
                .Height = VB6.TwipsToPixelsY(4830)
            End With
        End If
        mintCurrentNo = 1
        mintPrevMarkRow = 1
        mstrServerDate = getDateForDB(GetServerDate)
        DTPSheduleStartDate.Value = mstrServerDate
        DTPSheduleEndMonth.Value = mstrServerDate
        Exit Sub 'To avoid the execution of error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0026_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        '----------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/12/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - For Required code on Form QueryUnload
        '----------------------------------------------------
        On Error GoTo ErrHandler
        If gblnCancelUnload = True Then eventArgs.Cancel = 1
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub frmMKTTRN0026_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '----------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/12/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - For Required code on Form Unload
        '----------------------------------------------------
        On Error GoTo ErrHandler
        'Me = Nothing
        mdifrmMain.RemoveFormNameFromWindowList = mintFormIndex
        frmModules.NodeFontBold(Tag) = False
        Me.Dispose()
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Function GetFieldsValues(ByVal pstrQuery As String, ByVal pIntNoOfColumn As Short) As Object
        '----------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/12/2003
        'Arguments           - pstrQuery :- Select Statement
        '                    - pIntNoOfColumn :- Column No to Found
        'Return Value        - Variant :- Return Value
        'Function            - To Get Data in Feilds
        '----------------------------------------------------
        On Error GoTo ErrHandler
        Dim varReturnVal() As Object
        Dim rsRecordset As New ADODB.Recordset
        Dim lintLoop As Short
        ReDim varReturnVal(pIntNoOfColumn)
        If rsRecordset.State = 1 Then rsRecordset.Close()
        rsRecordset.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        rsRecordset.Open(pstrQuery, mP_Connection, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
        If rsRecordset.EOF Or rsRecordset.BOF Then
            GetFieldsValues = False
        Else
            If rsRecordset.RecordCount > 1 Then
                GetFieldsValues = False
            Else
                For lintLoop = 0 To pIntNoOfColumn - 1
                    varReturnVal(lintLoop) = rsRecordset.Fields(lintLoop).Value
                Next
            End If
        End If
        rsRecordset.Close()
        GetFieldsValues = VB6.CopyArray(varReturnVal)
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub SetGrid()
        '----------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/12/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Set Grid Headers
        '----------------------------------------------------
        Dim iLoopCounter As Short
        Dim lintInnerCnt As Short
        Dim lintIndex As Short
        Dim strMonth_Year As String
        Dim strIndMonth() As String
        On Error GoTo ErrHandler
        If mblnSingleSave Then
            With spdWholeEntryGrid
                .MaxRows = 0
                .MaxCols = 8
                .Row = 0
                .set_RowHeight(0, 300)
                .DisplayRowHeaders = True
                .ColHeaderDisplay = FPSpreadADO.HeaderDisplayConstants.DispNumbers
                .Col = enumWholeEntryGrid.COLUMN_CUSTDRGNO
                .Value = "Customer Drawing No."
                .set_ColWidth(enumWholeEntryGrid.COLUMN_CUSTDRGNO, 2500)
                .Col = enumWholeEntryGrid.COLUMN_ITEMCODE
                .Value = "Item Code"
                .set_ColWidth(enumWholeEntryGrid.COLUMN_ITEMCODE, 2000)
                .Col = enumWholeEntryGrid.COLUMN_CUSTDRGDESC
                .Value = "Description"
                .set_ColWidth(enumWholeEntryGrid.COLUMN_CUSTDRGDESC, 3300)
                .Col = enumWholeEntryGrid.COLUMN_SCHEDULEQTY
                .Col2 = enumWholeEntryGrid.COLUMN_YEAR_MONTH
                .Row = 0
                .Row = .MaxRows
                .BlockMode = True
                .ColHidden = True
                .BlockMode = False
            End With
            With spdSingleEntryGrid
                .MaxRows = 0
                .MaxCols = 5
                .Row = 0
                .set_RowHeight(0, 300)
                .DisplayRowHeaders = True
                .ColHeaderDisplay = FPSpreadADO.HeaderDisplayConstants.DispNumbers
                .Col = enumSingleEntryGrid.COLUMN_MONTH_YEAR
                .Value = "Month - Year"
                .set_ColWidth(enumSingleEntryGrid.COLUMN_MONTH_YEAR, 2000)
                .Col = enumSingleEntryGrid.COLUMN_SCHEDULE_QTY
                .Value = "Schedule Qty."
                .set_ColWidth(enumSingleEntryGrid.COLUMN_SCHEDULE_QTY, 1500)
                If CmdGrpMktSchedule.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                    .Col = enumSingleEntryGrid.COLUMN_REVISIONCOUNT
                    .Value = "Revision Count"
                    .ColHidden = False
                    .set_ColWidth(enumSingleEntryGrid.COLUMN_REVISIONCOUNT, 1300)
                    .Col = enumSingleEntryGrid.COLUMN_DESPATCH_QTY
                    .Value = "Despatch Qty"
                    .ColHidden = False
                    .set_ColWidth(enumSingleEntryGrid.COLUMN_DESPATCH_QTY, 1500)
                Else
                    .Col = enumSingleEntryGrid.COLUMN_REVISIONCOUNT
                    .ColHidden = True
                    .Col = enumSingleEntryGrid.COLUMN_DESPATCH_QTY
                    .ColHidden = True
                End If
                .Col = enumSingleEntryGrid.COLUMN_REMARKS
                .Value = "Remarks"
                .set_ColWidth(enumSingleEntryGrid.COLUMN_REMARKS, 3000)
            End With
        Else
            mblnSingleSave = False
            spdSingleEntryGrid.Visible = False
            fraSavingFrame.Visible = False
            With spdWholeEntryGrid
                .Height = VB6.TwipsToPixelsY(4830)
            End With
            If CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                With spdWholeEntryGrid
                    .MaxRows = 0
                    strMonth_Year = GetMonthCondition("MMM")
                    strIndMonth = Split(strMonth_Year, ",")
                    .MaxCols = UBound(strIndMonth) + 4
                    .Row = 0
                    .set_RowHeight(0, 300)
                    .DisplayRowHeaders = True
                    .ColHeaderDisplay = FPSpreadADO.HeaderDisplayConstants.DispNumbers
                    .Col = enumWholeEntryGrid.COLUMN_CUSTDRGNO
                    .Value = "Customer Drawing No."
                    .set_ColWidth(enumWholeEntryGrid.COLUMN_CUSTDRGNO, 2500)
                    .Col = enumWholeEntryGrid.COLUMN_ITEMCODE
                    .Value = "Item Code"
                    .set_ColWidth(enumWholeEntryGrid.COLUMN_ITEMCODE, 2000)
                    .Col = enumWholeEntryGrid.COLUMN_CUSTDRGDESC
                    .Value = "Description"
                    .set_ColWidth(enumWholeEntryGrid.COLUMN_CUSTDRGDESC, 3300)
                    For iLoopCounter = 4 To .MaxCols
                        .Col = iLoopCounter
                        .Value = strIndMonth(iLoopCounter - 4)
                        .set_ColWidth(iLoopCounter, 1500)
                    Next
                    .Col = enumWholeEntryGrid.COLUMN_CUSTDRGNO
                    .Col2 = enumWholeEntryGrid.COLUMN_CUSTDRGDESC
                    .Row = 0
                    .Row = .MaxRows
                    .BlockMode = True
                    .Lock = True
                    .BlockMode = False
                End With
            Else
                With spdWholeEntryGrid
                    .MaxRows = 0
                    strMonth_Year = GetMonthCondition("MMM")
                    strIndMonth = Split(strMonth_Year, ",")
                    .MaxCols = 3 * (UBound(strIndMonth)) + 6
                    .Row = 0
                    .set_RowHeight(0, 500)
                    .DisplayRowHeaders = True
                    .ColHeaderDisplay = FPSpreadADO.HeaderDisplayConstants.DispNumbers
                    .Col = enumWholeEntryGrid.COLUMN_CUSTDRGNO
                    .Value = "Customer Drawing No."
                    .set_ColWidth(enumWholeEntryGrid.COLUMN_CUSTDRGNO, 2500)
                    .Col = enumWholeEntryGrid.COLUMN_ITEMCODE
                    .Value = "Item Code"
                    .set_ColWidth(enumWholeEntryGrid.COLUMN_ITEMCODE, 2000)
                    .Col = enumWholeEntryGrid.COLUMN_CUSTDRGDESC
                    .Value = "Description"
                    .set_ColWidth(enumWholeEntryGrid.COLUMN_CUSTDRGDESC, 3300)
                    lintIndex = 0
                    For iLoopCounter = 3 To .MaxCols - 1 Step 3
                        For lintInnerCnt = 1 To 3
                            .Col = iLoopCounter + lintInnerCnt
                            If lintInnerCnt = 1 Then
                                .Value = strIndMonth(lintIndex) & " (Schedule Qty)"
                                .set_ColWidth(.Col, 1500)
                            ElseIf lintInnerCnt = 2 Then
                                .Value = strIndMonth(lintIndex) & " (Dispatch Qty)"
                                .set_ColWidth(.Col, 1500)
                            Else
                                .Value = strIndMonth(lintIndex) & " (Revision No)"
                                .set_ColWidth(.Col, 1500)
                            End If
                        Next
                        lintIndex = lintIndex + 1
                    Next
                    .Col = enumWholeEntryGrid.COLUMN_CUSTDRGNO
                    .Col2 = .MaxCols
                    .Row = 0
                    .Row = .MaxRows
                    .BlockMode = True
                    .Lock = True
                    .BlockMode = False
                End With
            End If
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub RefreshForm()
        '----------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/12/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Refresh Both the Grids & other Controls
        '----------------------------------------------------
        On Error GoTo ErrHandler
        spdWholeEntryGrid.MaxRows = 0
        spdSingleEntryGrid.MaxRows = 0
        Call EnableControls(True, Me, True) 'Enable Controls
        DTPSheduleStartDate.Focus()
        cmdSingleEntrySave.Enabled = False
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub spdSingleEntryGrid_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxFPSpreadADO._DSpreadEvents_KeyPressEvent) Handles spdSingleEntryGrid.KeyPressEvent
        '----------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/12/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To call Validation on Existing Cell on enter key Press
        '----------------------------------------------------
        If eventArgs.keyAscii = 13 And spdSingleEntryGrid.ActiveRow = spdSingleEntryGrid.MaxRows And (spdSingleEntryGrid.ActiveCol = enumSingleEntryGrid.COLUMN_SCHEDULE_QTY Or spdSingleEntryGrid.ActiveCol = enumSingleEntryGrid.COLUMN_REMARKS) Then
            Call spdSingleEntryGrid_Validating(spdSingleEntryGrid, New System.ComponentModel.CancelEventArgs(False))
        End If
    End Sub
    Private Sub spdSingleEntryGrid_LeaveCell(ByVal eventSender As System.Object, ByVal eventArgs As AxFPSpreadADO._DSpreadEvents_LeaveCellEvent) Handles spdSingleEntryGrid.LeaveCell
        '----------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/12/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Put Validation on Cell lavel
        '----------------------------------------------------
        Dim dblPrevQty As Double
        Dim dblscheduleqty As Double
        Dim dblDespatchQty As Double
        On Error GoTo ErrHandler
        If blnCheckMsg = True Then
            blnCheckMsg = False
            Exit Sub
        End If
        ''----Added by Davinder on 30/03/2006 (Issue ID:-17378) to validate the Sch. Qty depending upon the Measurement code of the Item
        If eventArgs.newCol = -1 Then
            mRow = eventArgs.row
            mCol = eventArgs.col
            Exit Sub
        End If
        If eventArgs.row > 0 Then
            With spdSingleEntryGrid
                If eventArgs.col = enumSingleEntryGrid.COLUMN_SCHEDULE_QTY Then
                    If Me.CmdGrpMktSchedule.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                        If DecimalAllowed(eventArgs.row, Trim(Me.txtItemCode.Text)) Then
                            eventArgs.cancel = True
                            .Focus()
                            Exit Sub
                        End If
                    End If
                End If
                ''----Changes by Davinder end's here
                If Not .Lock Then
                    If eventArgs.newRow = eventArgs.row + 1 Then
                        .Row = eventArgs.row
                        .Col = enumSingleEntryGrid.COLUMN_SCHEDULE_QTY
                        dblPrevQty = Val(.Text)
                        .Col = enumSingleEntryGrid.COLUMN_SCHEDULE_QTY
                        .Row = eventArgs.newRow
                        If Val(.Text) = 0 Then
                            .Text = CStr(dblPrevQty)
                        End If
                    End If
                End If
                .Row = eventArgs.row
                .Col = enumSingleEntryGrid.COLUMN_DESPATCH_QTY
                dblDespatchQty = Val(.Text)
                .Col = enumSingleEntryGrid.COLUMN_SCHEDULE_QTY
                dblscheduleqty = Val(.Text)
                If Not mblnDirty Then
                    If dblscheduleqty < dblDespatchQty Then
                        mblnDirty = True
                        MsgBox("Schedule Quantity Can Not Be Less Than Despatch Quantity", MsgBoxStyle.Information, "empower")
                        mblnDirty = False
                        .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                        .Focus()
                    End If
                End If
            End With
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub spdSingleEntryGrid_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles spdSingleEntryGrid.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        '----------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/12/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Validate data on setting Focus to Next Control
        '----------------------------------------------------
        On Error GoTo ErrHandler
        If CmdGrpMktSchedule.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            cmdSingleEntrySave.Focus()
        End If
        GoTo EventExitSub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub spdWholeEntryGrid_DblClick(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_DblClickEvent) Handles spdWholeEntryGrid.DblClick
        '----------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/12/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Call when single Row saving option is On
        '----------------------------------------------------
        On Error GoTo ErrHandler
        If mblnSingleSave = True Then
            mintCurrentNo = e.row
            Call SetTextBoxValue(mintCurrentNo)
            If CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                Call FillSingleGrid("ADD")
            ElseIf CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                Call FillSingleGrid("EDIT")
            Else
                Call FillSingleGrid("VIEW")
            End If
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub spdWholeEntryGrid_LeaveCell1(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_LeaveCellEvent) Handles spdWholeEntryGrid.LeaveCell
        '----------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/12/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Put Validation on Cell lavel
        '----------------------------------------------------
        Dim dblscheduleqty As Double
        Dim dblDespatchQty As Double
        Dim strMonth_Year As String
        Dim strYear() As String
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        Dim dblCol As Double
        On Error GoTo ErrHandler
        If e.newCol = -1 Then Exit Sub
        If mblnSingleSave = False Then
            If CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                strMonth_Year = GetMonthCondition("MY")
                strYear = Split(strMonth_Year, ",")
                With spdWholeEntryGrid
                    If e.col <> enumWholeEntryGrid.COLUMN_SCHEDULEQTY Then
                        dblCol = e.col - enumWholeEntryGrid.COLUMN_SCHEDULEQTY
                        If Int(dblCol / 3) <> (dblCol / 3) Then
                            Exit Sub
                        End If
                    End If
                    intMaxLoop = UBound(strYear)
                    For intLoopCounter = 0 To intMaxLoop
                        .Row = e.row
                        .Col = enumWholeEntryGrid.COLUMN_SCHEDULEQTY + (3 * intLoopCounter)
                        dblscheduleqty = Val(.Text)
                        .Col = enumWholeEntryGrid.COLUMN_DESPATCHQTY + (3 * intLoopCounter)
                        dblDespatchQty = Val(.Text)
                        If dblscheduleqty < dblDespatchQty Then
                            mblnDirty = True
                            MsgBox("Schedule Quantity Can Not Be Less Than Despatch Quantity", MsgBoxStyle.Information, "empower")
                            .Col = enumWholeEntryGrid.COLUMN_SCHEDULEQTY + (3 * intLoopCounter)
                            .Row = e.row
                            .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                            .Focus()
                            Exit Sub
                        End If
                    Next
                End With
            End If
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtCustDrgNo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCustDrgNo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '----------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/12/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Call Validate Cust Drg No in Case of Single row Saving is On
        '----------------------------------------------------
        If KeyAscii = 13 Then
            Call txtCustDrgNo_Validating(txtCustDrgNo, New System.ComponentModel.CancelEventArgs(False))
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtCustDrgNo_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtCustDrgNo.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        '----------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/12/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Call Validate Cust Drg No in Case of Single row Saving is On
        '----------------------------------------------------
        On Error GoTo ErrHandler
        If blnCheckMsg = True Then
            txtCustDrgNo.Focus()
            blnCheckMsg = False
            Exit Sub
        End If
        If Not Me.CmdGrpMktSchedule.GetActiveButton Is Nothing Then
            If CmdGrpMktSchedule.GetActiveButton.Text.ToUpper = "print".ToUpper Then
                Exit Sub
            End If
        End If
        If Not ValidateEntry(Trim(txtCustDrgNo.Text), "C") Then
            MsgBox("Invalid Customer Drawing No.", MsgBoxStyle.Information, "empower")
            blnCheckMsg = True
            txtCustDrgNo.Focus()
            Cancel = True
        Else
            If CmdGrpMktSchedule.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                spdSingleEntryGrid.Row = 1
                spdSingleEntryGrid.Col = enumSingleEntryGrid.COLUMN_SCHEDULE_QTY
                spdSingleEntryGrid.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                spdSingleEntryGrid.Focus()
            Else
                blnCheckMsg = True
                txtCustDrgNo.Focus()
            End If
        End If
        GoTo EventExitSub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtCustomerCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCustomerCode.TextChanged
        '----------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/12/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Clear Cusomer Name Label
        '----------------------------------------------------
        On Error GoTo ErrHandler
        If Trim(txtCustomerCode.Text) = "" Then lblCustomerDesc.Text = ""
        '''Changes done By ashutosh on 02 Jul 2007, Issue Id:20418
        If Me.spdSingleEntryGrid.Visible = True Then
            spdSingleEntryGrid.MaxRows = 0
        End If
        If Me.spdWholeEntryGrid.Visible = True Then
            spdWholeEntryGrid.MaxRows = 0
        End If
        If txtCustDrgNo.Enabled = True Then txtCustDrgNo.Text = ""
        If txtItemCode.Enabled = True Then txtItemCode.Text = ""
        If txtLineNo.Enabled = True Then txtLineNo.Text = ""
        '''Changes for Issue Id:40218 end here.
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtCustomerCode_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCustomerCode.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '----------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/12/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Call Customer Code Help on F1 Press
        '----------------------------------------------------
        If KeyCode = System.Windows.Forms.Keys.F1 Then Call cmdCustomerHelp_Click(cmdCustomerHelp, New System.EventArgs())
    End Sub
    Private Sub txtCustomerCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCustomerCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '----------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/12/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Call Validate Customer Code in Case of Single row Saving is On
        '----------------------------------------------------
        If KeyAscii = 13 Then Call txtCustomerCode_Validating(txtCustomerCode, New System.ComponentModel.CancelEventArgs(False))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtCustomerCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtCustomerCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        '----------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/12/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Validate Customer Code in Case of Single row Saving is On
        '----------------------------------------------------
        Dim varGetCustName() As Object
        On Error GoTo ErrHandler
        If Trim(txtCustomerCode.Text) = "" Then
            CmdGrpMktSchedule.Focus()
            GoTo EventExitSub
        End If
        If CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            '''Changes done by Ashutosh on 02 jul 2007, Issue id:20418
            varGetCustName = GetFieldsValues("select distinct a.Account_code, b.Cust_Name from  MonthlyMktSchedule a, customer_mst b where a.account_code=b.Customer_Code and a.UNIT_CODE=b.UNIT_CODE AND Account_code='" & Trim(txtCustomerCode.Text) & "' and a.UNIT_CODE='" & gstrUNITID & "' and ((isnull(b.deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= b.deactive_date))", 2)
            '''varGetCustName = GetFieldsValues("SELECT DISTINCT Account_code , Prty_Name FROM MonthlyMktSchedule ,Gen_PartyMaster WHERE Account_Code=Prty_PartyID AND Account_code='" & Trim(txtCustomerCode.Text) & "'", 2)
        Else
            varGetCustName = GetFieldsValues("SELECT Distinct a.Account_code, b.Cust_Name FROM Cust_Ord_Hdr a,customer_mst b WHERE a.account_code=b.Customer_Code and a.UNIT_CODE=b.UNIT_CODE and b.AllowExcessSchedule=1 And Account_code='" & Trim(txtCustomerCode.Text) & "' and a.UNIT_CODE='" & gstrUNITID & "' and ((isnull(b.deactive_flag,0) <> 1) OR (cast(getdate() AS date)<= b.deactive_date))", 2)
            '''varGetCustName = GetFieldsValues("SELECT DISTINCT Account_code , Prty_Name FROM Cust_Ord_Hdr ,Gen_PartyMaster WHERE Account_Code=Prty_PartyID AND Account_code='" & Trim(txtCustomerCode.Text) & "'", 2)
        End If
        'Changed for Issue ID 20665 Starts
        'If Trim(varGetCustName(0)) = Trim(txtCustomerCode) Then
        If UCase(Trim(varGetCustName(0))) = UCase(Trim(txtCustomerCode.Text)) Then
            lblCustomerDesc.Text = varGetCustName(1)
            Call FillWholeGrid()
        Else
            MsgBox("Invalid Customer Code", MsgBoxStyle.Information, "empower")
            txtCustomerCode.Text = ""
            Cancel = True
            GoTo EventExitSub
        End If
        If CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            Me.CmdGrpMktSchedule.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
        End If
        GoTo EventExitSub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Function FillWholeGrid() As Boolean
        '----------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/12/2003
        'Arguments           - None
        'Return Value        - Boolean
        'Function            - To Fill Data in Whole Grid
        '----------------------------------------------------
        On Error GoTo ErrHandler
        If mblnSingleSave Then
            If CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                If Not DisablePrimaryField(True) Then Exit Function
                Call SetGrid()
                '''Changes done By ashutosh on 24-10-206,Issue Id:18623
                '''Call GetItemList(Trim(txtCustomerCode.Text), "ADD")
                If Not GetItemList(Trim(txtCustomerCode.Text), "ADD") Then Exit Function
                '''Changes for Issue Id:18623 end here.
                Call SetTextBoxValue(mintCurrentNo)
                Call FillSingleGrid("ADD")
                With spdWholeEntryGrid
                    .Col = enumWholeEntryGrid.COLUMN_SCHEDULEQTY
                    .ColHidden = True
                    .Col = enumWholeEntryGrid.COLUMN_REMARKS
                    .ColHidden = True
                End With
            ElseIf CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                If Not DisablePrimaryField(False) Then Exit Function
                Call SetGrid()
                mintCurrentNo = 1
                Call GetItemList(Trim(txtCustomerCode.Text), "VIEW")
                Call SetTextBoxValue(mintCurrentNo)
                Call FillSingleGrid("VIEW")
                With spdWholeEntryGrid
                    .Col = enumWholeEntryGrid.COLUMN_SCHEDULEQTY
                    .ColHidden = True
                    .Col = enumWholeEntryGrid.COLUMN_REMARKS
                    .ColHidden = True
                    .Col = enumWholeEntryGrid.COLUMN_DESPATCHQTY
                    .ColHidden = True
                    .Col = enumWholeEntryGrid.COLUMN_REVISION_NO
                    .ColHidden = True
                    .Col = enumWholeEntryGrid.COLUMN_YEAR_MONTH
                    .ColHidden = True
                End With
                CmdGrpMktSchedule.Focus()
            End If
            If AllowEditSchedule() = False Then
                CmdGrpMktSchedule.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
            Else
                CmdGrpMktSchedule.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = True
            End If
        Else
            If CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                If Not DisablePrimaryField(True) Then Exit Function
                Call SetGrid()
                Call GetItemList(Trim(txtCustomerCode.Text), "ADD")
            ElseIf CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                If Not DisablePrimaryField(False) Then Exit Function
                Call SetGrid()
                Call GetItemList(Trim(txtCustomerCode.Text), "VIEW")
                CmdGrpMktSchedule.Focus()
            End If
        End If
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function AllowEditSchedule() As Boolean
        '*******************************************************************************
        'Author             :   Ashutosh Verma
        'Argument(s)if any  :
        'Return Value       :   Boolean
        'Function           :   Allow the shcedule to be edited or not.
        'Creation Date      :   02 Jul 2007, Issue Id:20418
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim rsEXcessSchCustomer As ClsResultSetDB
        rsEXcessSchCustomer = New ClsResultSetDB
        rsEXcessSchCustomer.GetResult("Select isnull(AllowExcessSchedule,1) as AllowExcessSchedule from customer_mst where Customer_code='" & Trim(txtCustomerCode.Text) & "' and UNIT_CODE='" & gstrUNITID & "'")
        If rsEXcessSchCustomer.GetNoRows > 0 Then
            If rsEXcessSchCustomer.GetValue("AllowExcessSchedule") = False Then
                AllowEditSchedule = False
            Else
                AllowEditSchedule = True
            End If
        End If
        rsEXcessSchCustomer.ResultSetClose()
        rsEXcessSchCustomer = Nothing
        Exit Function
ErrHandler:
        AllowEditSchedule = True
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function GetItemList(ByVal pstrCustomerCode As String, ByVal pstrMode As String) As Boolean
        '----------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/12/2003
        'Arguments           - pstrCustomerCode -: Customer Code
        '                    -  pstrMode -:ADD/EDIT
        'Return Value        - None
        'Function            - To Get Item Code For Given Customer
        '----------------------------------------------------
        Dim rsGetItem As New ADODB.Recordset
        Dim strSql As String
        Dim strMonthCondition As String
        Dim lintMonthCnt As Short
        Dim lintLoop As Short
        Dim strPrevCustDrgNo As String
        Dim strPrevItemCode As String
        Dim strPrevMonth As String
        Dim lintRowCounter As Short
        Dim strMonth_Year As String
        Dim strYear() As String
        On Error GoTo ErrHandler
        strMonth_Year = GetMonthCondition("MY")
        strYear = Split(strMonth_Year, ",")
        GetItemList = True
        If rsGetItem.State = 1 Then rsGetItem.Close()
        rsGetItem.CursorLocation = ADODB.CursorLocationEnum.adUseClient
        strSql = ""
        If UCase(Trim(pstrMode)) = UCase("ADD") Then
            strSql = "SELECT DISTINCT Cust_ord_dtl.Cust_Drgno,CustItem_mst.Item_Code,Drg_Desc FROM Cust_ord_dtl,CustItem_Mst, item_mst "
            strSql = strSql & " WHERE NOT EXISTS (SELECT * FROM monthlymktschedule WHERE account_code='" & pstrCustomerCode & "'"
            strSql = strSql & " AND monthlymktschedule.Cust_Drgno = Cust_ord_dtl.Cust_Drgno and monthlymktschedule.UNIT_CODE = Cust_ord_dtl.UNIT_CODE AND monthlymktschedule.UNIT_CODE='" & gstrUNITID & "' "
            strSql = strSql & " AND monthlymktschedule.Item_Code = Cust_ord_dtl.Item_Code  AND year_month IN (" & GetMonthCondition("MY") & "))"
            strSql = strSql & " AND NOT EXISTS"
            strSql = strSql & " (SELECT * FROM dailymktschedule WHERE account_code='" & pstrCustomerCode & "'"
            strSql = strSql & " AND dailymktschedule.Cust_Drgno = Cust_ord_dtl.Cust_Drgno AND dailymktschedule.UNIT_CODE = Cust_ord_dtl.UNIT_CODE"
            strSql = strSql & " AND dailymktschedule.UNIT_CODE='" & gstrUNITID & "'"
            strSql = strSql & " AND dailymktschedule.Item_Code = Cust_ord_dtl.Item_Code and month(trans_date) IN (" & GetMonthCondition("M") & ") AND year(trans_date) IN (" & GetMonthCondition("Y") & ")) "
            strSql = strSql & " AND item_mst.item_code = cust_ord_dtl.item_code AND item_mst.UNIT_CODE = cust_ord_dtl.UNIT_CODE and item_mst.item_main_grp in('F','T','S')"
            strSql = strSql & " AND CustItem_Mst.account_code='" & pstrCustomerCode & "' AND Cust_ord_dtl.cust_Drgno=CustItem_Mst.cust_Drgno AND Cust_ord_dtl.UNIT_CODE=CustItem_Mst.UNIT_CODE AND Cust_ord_dtl.Active_Flag='A' AND Cust_ord_dtl.Authorized_Flag=1 and Cust_ord_dtl.item_code =CustItem_Mst.item_code AND CustItem_Mst.account_code= Cust_ord_dtl.account_code AND Cust_ord_dtl.UNIT_CODE='" & gstrUNITID & "'"
            rsGetItem.Open(strSql, mP_Connection, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            If rsGetItem.EOF Or rsGetItem.BOF Then
                MsgBox("No Record Found", MsgBoxStyle.Critical, "empower")
                GetItemList = False
                '''Changes done By ashutosh on 24-10-2006,Issue Id:18623
                Me.txtCustomerCode.Text = ""
                Me.lblCustomerDesc.Text = ""
                txtLineNo.Text = ""
                txtCustDrgNo.Text = ""
                txtItemCode.Text = ""
                '''Changes for Issue Id:18623 end here.
                Exit Function
            Else
                With spdWholeEntryGrid
                    mintMaxCurrentNo = rsGetItem.RecordCount
                    For lintLoop = 1 To rsGetItem.RecordCount
                        .MaxRows = lintLoop
                        .Row = lintLoop
                        .Col = enumWholeEntryGrid.COLUMN_CUSTDRGNO
                        .Text = Trim(rsGetItem.Fields("Cust_Drgno").Value)
                        .Col = enumWholeEntryGrid.COLUMN_ITEMCODE
                        .Text = Trim(rsGetItem.Fields("Item_Code").Value)
                        .Col = enumWholeEntryGrid.COLUMN_CUSTDRGDESC
                        .Text = Trim(rsGetItem.Fields("Drg_Desc").Value)
                        rsGetItem.MoveNext()
                    Next
                    .Row = 1
                    .Row2 = .MaxRows
                    .Col = enumWholeEntryGrid.COLUMN_CUSTDRGNO
                    .Col2 = enumWholeEntryGrid.COLUMN_CUSTDRGDESC
                    .BlockMode = True
                    .Lock = True
                    .BlockMode = False
                End With
            End If
        ElseIf UCase(Trim(pstrMode)) = UCase("VIEW") Then
            strSql = "SELECT monthlymktschedule.cust_drgno,monthlymktschedule.item_code,year_month,schedule_qty,Despatch_qty,RevisionNo,Drg_Desc,isnull(Remarks,'') as Remarks"
            strSql = strSql & " FROM monthlymktschedule ,CustItem_Mst WHERE monthlymktschedule.UNIT_CODE=CustItem_Mst.UNIT_CODE and monthlymktschedule.UNIT_CODE='" & gstrUNITID & "' and Status=1  "
            strSql = strSql & " AND year_month IN (" & GetMonthCondition("MY") & ")"
            strSql = strSql & " AND monthlymktschedule.account_code='" & pstrCustomerCode & "'"
            strSql = strSql & " AND CustItem_Mst.account_code='" & pstrCustomerCode & "' AND monthlymktschedule.cust_Drgno=CustItem_Mst.cust_Drgno AND monthlymktschedule.Item_code=CustItem_Mst.Item_Code"
            strSql = strSql & " GROUP BY monthlymktschedule.cust_drgno,monthlymktschedule.item_code,Drg_Desc,year_month,schedule_qty,Despatch_qty,RevisionNo,Remarks HAVING revisionno=MAX(revisionno)"
            rsGetItem.Open(strSql, mP_Connection, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            If rsGetItem.EOF Or rsGetItem.BOF Then
                MsgBox("No Record Found", MsgBoxStyle.Critical, "empower")
                GetItemList = False
                Exit Function
            Else
                CmdGrpMktSchedule.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = True
                CmdGrpMktSchedule.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = True
                rsGetItem.MoveFirst()
                With spdWholeEntryGrid
                    mintMaxCurrentNo = rsGetItem.RecordCount
                    strPrevCustDrgNo = Trim(rsGetItem.Fields("Cust_Drgno").Value)
                    strPrevItemCode = Trim(rsGetItem.Fields("Item_Code").Value)
                    strPrevMonth = Trim(rsGetItem.Fields("Year_Month").Value)
                    lintRowCounter = 1
                    .MaxRows = lintRowCounter
                    While Not rsGetItem.EOF
                        If Trim(strPrevCustDrgNo) = Trim(rsGetItem.Fields("Cust_Drgno").Value) And Trim(strPrevItemCode) = Trim(rsGetItem.Fields("Item_Code").Value) Then
                            If Trim(strPrevMonth) = Trim(rsGetItem.Fields("Year_Month").Value) And lintRowCounter > 1 Then
                                lintRowCounter = lintRowCounter - 1
                                .MaxRows = lintRowCounter
                                GoTo MoveCursor
                            Else
                                .MaxRows = lintRowCounter
                                .Row = lintRowCounter
                            End If
                        Else
MoveCursor:
                            lintRowCounter = lintRowCounter + 1
                            .MaxRows = lintRowCounter
                            strPrevCustDrgNo = Trim(rsGetItem.Fields("Cust_Drgno").Value)
                            strPrevItemCode = Trim(rsGetItem.Fields("Item_Code").Value)
                            strPrevMonth = Trim(rsGetItem.Fields("Year_Month").Value)
                            .Row = lintRowCounter
                        End If
                        .Col = enumWholeEntryGrid.COLUMN_CUSTDRGNO
                        .Text = Trim(rsGetItem.Fields("Cust_Drgno").Value)
                        .Col = enumWholeEntryGrid.COLUMN_ITEMCODE
                        .Text = Trim(rsGetItem.Fields("Item_Code").Value)
                        .Col = enumWholeEntryGrid.COLUMN_CUSTDRGDESC
                        .Text = Trim(rsGetItem.Fields("Drg_Desc").Value)
                        If mblnSingleSave Then
                            .Col = enumWholeEntryGrid.COLUMN_SCHEDULEQTY
                            .Text = .Text & Convert.ToDouble(Trim(rsGetItem.Fields("Schedule_Qty").Value)).ToString & "�"
                            .Col = enumWholeEntryGrid.COLUMN_DESPATCHQTY
                            .Text = .Text & Trim(rsGetItem.Fields("Despatch_Qty").Value) & "�"
                            .Col = enumWholeEntryGrid.COLUMN_REVISION_NO
                            .Text = .Text & Trim(rsGetItem.Fields("RevisionNo").Value) & "�"
                            .Col = enumWholeEntryGrid.COLUMN_REMARKS
                            .Text = .Text & Trim(rsGetItem.Fields("Remarks").Value) & "�"
                            .Col = enumWholeEntryGrid.COLUMN_YEAR_MONTH
                            .Text = .Text & Trim(rsGetItem.Fields("Year_Month").Value) & "�"
                        Else
                            For lintMonthCnt = 0 To UBound(strYear) - 1
                                If Trim(rsGetItem.Fields("Year_Month").Value) = Trim(strYear(lintMonthCnt)) Then
                                    Exit For
                                End If
                            Next
                            .Col = enumWholeEntryGrid.COLUMN_SCHEDULEQTY + (3 * lintMonthCnt)
                            .Text = Trim(rsGetItem.Fields("Schedule_Qty").Value)
                            .Col = enumWholeEntryGrid.COLUMN_DESPATCHQTY + (3 * lintMonthCnt)
                            .Text = Trim(rsGetItem.Fields("Despatch_Qty").Value)
                            .Col = enumWholeEntryGrid.COLUMN_REVISION_NO + (3 * lintMonthCnt)
                            .Text = Trim(rsGetItem.Fields("RevisionNo").Value)
                        End If
                        rsGetItem.MoveNext()
                    End While
                    rsGetItem.Close()
                    mintMaxCurrentNo = .MaxRows
                    .Row = 1
                    .Row2 = .MaxRows
                    .Col = enumWholeEntryGrid.COLUMN_CUSTDRGNO
                    .Col2 = .MaxCols
                    .BlockMode = True
                    .Lock = True
                    .BlockMode = False
                End With
            End If
        End If
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function GetMonthCondition(ByVal pstrForWhichSchedule As String) As String
        '----------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/12/2003
        'Arguments           - pstrForWhichSchedule -: Month Range
        'Return Value        - String :- Year Month string
        'Function            - To Get Year Month string
        '----------------------------------------------------
        Dim lintLoop As Short
        Dim strReturnStr As String
        Dim strStartYear As String
        Dim strEndYear As String
        Dim lintMonthDiff As Short
        Dim strMidStr As String
        On Error GoTo ErrHandler
        strStartYear = VB6.Format(DTPSheduleStartDate.Value, "yyyy")
        strEndYear = VB6.Format(DTPSheduleEndMonth.Value, "yyyy")
        If Trim(UCase(pstrForWhichSchedule)) = "Y" Then
            GetMonthCondition = strStartYear & "," & strEndYear
            Exit Function
        End If
        strReturnStr = ""
        If Trim(strStartYear) = Trim(strEndYear) Then
            For lintLoop = CInt(VB6.Format(DTPSheduleStartDate.Value, "mm")) To CInt(VB6.Format(DTPSheduleEndMonth.Value, "mm"))
                If lintLoop < 10 Then
                    If Val(CStr(lintLoop)) = Val(VB6.Format(DTPSheduleEndMonth.Value, "mm")) Then
                        If Trim(UCase(pstrForWhichSchedule)) = "M" Then
                            strReturnStr = strReturnStr & "0" & CStr(lintLoop)
                        ElseIf Trim(UCase(pstrForWhichSchedule)) = "MMM" Then
                            strMidStr = "0" & CStr(lintLoop) & "/" & VB6.Format(DTPSheduleStartDate.Value, "yyyy")
                            strReturnStr = strReturnStr & VB6.Format(strMidStr, "mmm") & " - " & VB6.Format(DTPSheduleStartDate.Value, "yyyy")
                        Else
                            strReturnStr = strReturnStr & VB6.Format(DTPSheduleStartDate.Value, "yyyy") & "0" & CStr(lintLoop)
                        End If
                    Else
                        If Trim(UCase(pstrForWhichSchedule)) = "M" Then
                            strReturnStr = strReturnStr & "0" & CStr(lintLoop) & ","
                        ElseIf Trim(UCase(pstrForWhichSchedule)) = "MMM" Then
                            strMidStr = "0" & CStr(lintLoop) & "/" & VB6.Format(DTPSheduleStartDate.Value, "yyyy")
                            strReturnStr = strReturnStr & VB6.Format(strMidStr, "mmm") & " - " & VB6.Format(DTPSheduleStartDate.Value, "yyyy") & ","
                        Else
                            strReturnStr = strReturnStr & VB6.Format(DTPSheduleStartDate.Value, "yyyy") & "0" & CStr(lintLoop) & ","
                        End If
                    End If
                Else
                    If Val(CStr(lintLoop)) = Val(VB6.Format(DTPSheduleEndMonth.Value, "mm")) Then
                        If Trim(UCase(pstrForWhichSchedule)) = "M" Then
                            strReturnStr = strReturnStr & CStr(lintLoop)
                        ElseIf Trim(UCase(pstrForWhichSchedule)) = "MMM" Then
                            strMidStr = CStr(lintLoop) & "/" & VB6.Format(DTPSheduleStartDate.Value, "yyyy")
                            strReturnStr = strReturnStr & VB6.Format(strMidStr, "mmm") & " - " & VB6.Format(DTPSheduleStartDate.Value, "yyyy")
                        Else
                            strReturnStr = strReturnStr & VB6.Format(DTPSheduleStartDate.Value, "yyyy") & CStr(lintLoop)
                        End If
                    Else
                        If Trim(UCase(pstrForWhichSchedule)) = "M" Then
                            strReturnStr = strReturnStr & CStr(lintLoop) & ","
                        ElseIf Trim(UCase(pstrForWhichSchedule)) = "MMM" Then
                            strMidStr = CStr(lintLoop) & "/" & VB6.Format(DTPSheduleStartDate.Value, "yyyy")
                            strReturnStr = strReturnStr & VB6.Format(strMidStr, "mmm") & " - " & VB6.Format(DTPSheduleStartDate.Value, "yyyy") & ","
                        Else
                            strReturnStr = strReturnStr & VB6.Format(DTPSheduleStartDate.Value, "yyyy") & CStr(lintLoop) & ","
                        End If
                    End If
                End If
            Next
        Else
            lintMonthDiff = 12 - CDbl(VB6.Format(DTPSheduleStartDate.Value, "mm"))
            For lintLoop = CInt(VB6.Format(DTPSheduleStartDate.Value, "mm")) To lintMonthDiff + CDbl(VB6.Format(DTPSheduleStartDate.Value, "mm"))
                If lintLoop < 10 Then
                    If Trim(UCase(pstrForWhichSchedule)) = "M" Then
                        strReturnStr = strReturnStr & "0" & CStr(lintLoop) & ","
                    ElseIf Trim(UCase(pstrForWhichSchedule)) = "MMM" Then
                        strMidStr = "0" & CStr(lintLoop) & "/" & VB6.Format(DTPSheduleStartDate.Value, "yyyy")
                        strReturnStr = strReturnStr & VB6.Format(strMidStr, "mmm") & " - " & VB6.Format(DTPSheduleStartDate.Value, "yyyy") & ","
                    Else
                        strReturnStr = strReturnStr & VB6.Format(DTPSheduleStartDate.Value, "yyyy") & "0" & CStr(lintLoop) & ","
                    End If
                Else
                    If Trim(UCase(pstrForWhichSchedule)) = "M" Then
                        strReturnStr = strReturnStr & CStr(lintLoop) & ","
                    ElseIf Trim(UCase(pstrForWhichSchedule)) = "MMM" Then
                        strMidStr = CStr(lintLoop) & "/" & VB6.Format(DTPSheduleStartDate.Value, "yyyy")
                        strReturnStr = strReturnStr & VB6.Format(strMidStr, "mmm") & " - " & VB6.Format(DTPSheduleStartDate.Value, "yyyy") & ","
                    Else
                        strReturnStr = strReturnStr & VB6.Format(DTPSheduleStartDate.Value, "yyyy") & CStr(lintLoop) & ","
                    End If
                End If
            Next
            lintMonthDiff = lintMonthDiff + CDbl(VB6.Format(DTPSheduleEndMonth.Value, "mm"))
            For lintLoop = 1 To lintMonthDiff
                If lintLoop < 10 Then
                    If Val(CStr(lintLoop)) = Val(CStr(lintMonthDiff)) Then
                        If Trim(UCase(pstrForWhichSchedule)) = "M" Then
                            strReturnStr = strReturnStr & "0" & CStr(lintLoop)
                        ElseIf Trim(UCase(pstrForWhichSchedule)) = "MMM" Then
                            strMidStr = "0" & CStr(lintLoop) & "/" & VB6.Format(DTPSheduleEndMonth.Value, "yyyy")
                            strReturnStr = strReturnStr & VB6.Format(strMidStr, "mmm") & " - " & VB6.Format(DTPSheduleEndMonth.Value, "yyyy")
                        Else
                            strReturnStr = strReturnStr & VB6.Format(DTPSheduleEndMonth.Value, "yyyy") & "0" & CStr(lintLoop)
                        End If
                    Else
                        If Trim(UCase(pstrForWhichSchedule)) = "M" Then
                            strReturnStr = strReturnStr & "0" & CStr(lintLoop) & ","
                        ElseIf Trim(UCase(pstrForWhichSchedule)) = "MMM" Then
                            strMidStr = "0" & CStr(lintLoop) & "/" & VB6.Format(DTPSheduleEndMonth.Value, "yyyy")
                            strReturnStr = strReturnStr & VB6.Format(strMidStr, "mmm") & " - " & VB6.Format(DTPSheduleEndMonth.Value, "yyyy") & ","
                        Else
                            strReturnStr = strReturnStr & VB6.Format(DTPSheduleEndMonth.Value, "yyyy") & "0" & CStr(lintLoop) & ","
                        End If
                    End If
                Else
                    If Val(CStr(lintLoop)) = Val(CStr(lintMonthDiff)) Then
                        If Trim(UCase(pstrForWhichSchedule)) = "M" Then
                            strReturnStr = strReturnStr & CStr(lintLoop)
                        ElseIf Trim(UCase(pstrForWhichSchedule)) = "MMM" Then
                            strMidStr = CStr(lintLoop) & "/" & VB6.Format(DTPSheduleEndMonth.Value, "yyyy")
                            strReturnStr = strReturnStr & VB6.Format(strMidStr, "mmm") & " - " & VB6.Format(DTPSheduleEndMonth.Value, "yyyy")
                        Else
                            strReturnStr = strReturnStr & VB6.Format(DTPSheduleEndMonth.Value, "yyyy") & CStr(lintLoop)
                        End If
                    Else
                        If Trim(UCase(pstrForWhichSchedule)) = "M" Then
                            strReturnStr = strReturnStr & CStr(lintLoop) & ","
                        ElseIf Trim(UCase(pstrForWhichSchedule)) = "MMM" Then
                            strMidStr = CStr(lintLoop) & "/" & VB6.Format(DTPSheduleEndMonth.Value, "yyyy")
                            strReturnStr = strReturnStr & VB6.Format(strMidStr, "mmm") & " - " & VB6.Format(DTPSheduleEndMonth.Value, "yyyy") & ","
                        Else
                            strReturnStr = strReturnStr & VB6.Format(DTPSheduleEndMonth.Value, "yyyy") & CStr(lintLoop) & ","
                        End If
                    End If
                End If
            Next
        End If
        GetMonthCondition = strReturnStr
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function FillSingleGrid(ByVal pstrMode As String) As Boolean
        '----------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/12/2003
        'Arguments           - pstrMode
        'Return Value        - Boolean
        'Function            - To Fill Data in Single Grid Entry
        '----------------------------------------------------
        Dim strGetScheduleQty As String
        Dim strGetRemarks As String
        Dim strInd_Sch_Qty() As String
        Dim strInd_Rem() As String
        Dim strMonth_Year As String
        Dim strMonthYear() As String
        Dim lintLoop As Short
        Dim blnISSavedEntry As Boolean
        Dim strYear() As String
        Dim strDespatchQty As String
        Dim strRevisionNo As String
        Dim strYear_MonthOfWholeGrid As String
        Dim strIndYear() As String
        Dim strIndDespatchQty() As String
        Dim strIndRevisionNo() As String
        Dim lintInnerCnt As Short
        On Error GoTo ErrHandler
        strMonth_Year = GetMonthCondition("MMM")
        strMonthYear = Split(strMonth_Year, ",")
        strMonth_Year = GetMonthCondition("MY")
        strYear = Split(strMonth_Year, ",")
        With spdWholeEntryGrid
            .Row = Val(txtLineNo.Text)
            .Col = enumWholeEntryGrid.COLUMN_SCHEDULEQTY
            strGetScheduleQty = Trim(.Text)
            .Col = enumWholeEntryGrid.COLUMN_REMARKS
            strGetRemarks = Trim(.Text)
            If UCase(Trim(pstrMode)) = "EDIT" Or UCase(Trim(pstrMode)) = "VIEW" Then
                .Col = enumWholeEntryGrid.COLUMN_DESPATCHQTY
                strDespatchQty = Trim(.Text)
                .Col = enumWholeEntryGrid.COLUMN_REVISION_NO
                strRevisionNo = Trim(.Text)
                .Col = enumWholeEntryGrid.COLUMN_YEAR_MONTH
                strYear_MonthOfWholeGrid = Trim(.Text)
            End If
        End With
        strInd_Sch_Qty = Split(strGetScheduleQty, "�")
        strInd_Rem = Split(strGetRemarks, "�")
        strIndDespatchQty = Split(strDespatchQty, "�")
        strIndRevisionNo = Split(strRevisionNo, "�")
        blnISSavedEntry = ISSavedEntry(Val(txtLineNo.Text))
        If UCase(Trim(pstrMode)) = "ADD" Then
            cmdSingleEntrySave.Enabled = True
            With spdSingleEntryGrid
                .MaxRows = 0
                For lintLoop = 1 To UBound(strMonthYear) + 1
                    .MaxRows = lintLoop
                    .Row = lintLoop
                    .Col = enumSingleEntryGrid.COLUMN_MONTH_YEAR
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                    .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
                    .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
                    .Text = strMonthYear(lintLoop - 1)
                    If blnISSavedEntry Then
                        .Col = enumSingleEntryGrid.COLUMN_SCHEDULE_QTY
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                        .TypeFloatMin = 0
                        .Text = strInd_Sch_Qty(lintLoop - 1)
                        .Col = enumSingleEntryGrid.COLUMN_REMARKS
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                        .TypeMaxEditLen = 20
                        .Text = strInd_Rem(lintLoop - 1)
                    Else
                        .Col = enumSingleEntryGrid.COLUMN_SCHEDULE_QTY
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                        .TypeFloatMin = 0
                        .Col = enumSingleEntryGrid.COLUMN_REMARKS
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                        .TypeMaxEditLen = 20
                        .Text = ""
                    End If
                Next
                If Trim(UCase(pstrMode)) = "ADD" Then
                    .Col = enumSingleEntryGrid.COLUMN_DESPATCH_QTY
                    .ColHidden = True
                    .Col = enumSingleEntryGrid.COLUMN_REVISIONCOUNT
                    .ColHidden = True
                    If blnISSavedEntry Then
                        cmdSingleEntrySkip.Focus()
                    Else
                        .Row = 1
                        .Col = enumSingleEntryGrid.COLUMN_SCHEDULE_QTY
                        .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                        .Focus()
                    End If
                End If
            End With
        ElseIf UCase(Trim(pstrMode)) = "EDIT" Then
            cmdSingleEntrySave.Enabled = True
            With spdSingleEntryGrid
                .MaxRows = 0
                strIndYear = Split(strYear_MonthOfWholeGrid, "�")
                For lintLoop = 1 To UBound(strIndYear)
                    .MaxRows = lintLoop
                    .Row = lintLoop
                    .Col = enumSingleEntryGrid.COLUMN_MONTH_YEAR
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                    .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
                    .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
                    For lintInnerCnt = 0 To UBound(strYear) - 1
                        If Trim(strIndYear(lintLoop - 1)) = Trim(strYear(lintInnerCnt)) Then
                            Exit For
                        End If
                    Next
                    .Text = strMonthYear(lintInnerCnt)
                    .Col = enumSingleEntryGrid.COLUMN_SCHEDULE_QTY
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                    .TypeFloatMin = 0
                    .Text = strInd_Sch_Qty(lintLoop - 1)
                    .Col = enumSingleEntryGrid.COLUMN_REMARKS
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                    .TypeMaxEditLen = 20
                    .Text = strInd_Rem(lintLoop - 1)
                    .Col = enumSingleEntryGrid.COLUMN_DESPATCH_QTY
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                    .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                    .Text = strIndDespatchQty(lintLoop - 1)
                    .Col = enumSingleEntryGrid.COLUMN_REVISIONCOUNT
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                    .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
                    .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
                    .Text = strIndRevisionNo(lintLoop - 1)
                Next
                If blnISSavedEntry Then
                    cmdSingleEntrySkip.Focus()
                Else
                    .Row = 1
                    .Col = enumSingleEntryGrid.COLUMN_SCHEDULE_QTY
                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                    .Focus()
                End If
            End With
        Else
            cmdSingleEntrySave.Enabled = False
            With spdSingleEntryGrid
                .MaxRows = 0
                strIndYear = Split(strYear_MonthOfWholeGrid, "�")
                For lintLoop = 1 To UBound(strIndYear)
                    .MaxRows = lintLoop
                    .Row = lintLoop
                    .Col = enumSingleEntryGrid.COLUMN_MONTH_YEAR
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                    .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
                    .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
                    For lintInnerCnt = 0 To UBound(strYear) - 1
                        If Trim(strIndYear(lintLoop - 1)) = Trim(strYear(lintInnerCnt)) Then
                            Exit For
                        End If
                    Next
                    .Text = strMonthYear(lintInnerCnt)
                    .Col = enumSingleEntryGrid.COLUMN_SCHEDULE_QTY
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                    .TypeFloatMin = 0
                    .Text = strInd_Sch_Qty(lintLoop - 1)
                    .Col = enumSingleEntryGrid.COLUMN_REMARKS
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                    .TypeMaxEditLen = 20
                    .Text = strInd_Rem(lintLoop - 1)
                    .Col = enumSingleEntryGrid.COLUMN_DESPATCH_QTY
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                    .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                    .Text = strIndDespatchQty(lintLoop - 1)
                    .Col = enumSingleEntryGrid.COLUMN_REVISIONCOUNT
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                    .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
                    .TypeVAlign = FPSpreadADO.TypeVAlignConstants.TypeVAlignCenter
                    .Text = strIndRevisionNo(lintLoop - 1)
                Next
                .Col = enumSingleEntryGrid.COLUMN_MONTH_YEAR
                .Col2 = enumSingleEntryGrid.COLUMN_REMARKS
                .Row = 1
                .Row2 = .MaxRows
                .BlockMode = True
                .Lock = True
                .BlockMode = False
            End With
        End If
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function SetTextBoxValue(ByVal pintRowNo As Short) As Object
        '----------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/12/2003
        'Arguments           - pintRowNo -: To Get data in Row
        'Return Value        -
        'Function            - To Set Value in Text BOx
        '----------------------------------------------------
        On Error GoTo ErrHandler
        With spdWholeEntryGrid
            .Row = mintPrevMarkRow
            .Row2 = mintPrevMarkRow
            .Col = enumWholeEntryGrid.COLUMN_CUSTDRGNO
            .Col2 = enumWholeEntryGrid.COLUMN_CUSTDRGDESC
            .BlockMode = True
            .BackColor = System.Drawing.Color.White
            .BlockMode = False
            txtLineNo.Text = CStr(pintRowNo)
            .Row = pintRowNo
            .Col = enumWholeEntryGrid.COLUMN_CUSTDRGNO
            txtCustDrgNo.Text = Trim(.Text)
            .Col = enumWholeEntryGrid.COLUMN_ITEMCODE
            txtItemCode.Text = Trim(.Text)
            .Row2 = pintRowNo
            .Col = enumWholeEntryGrid.COLUMN_CUSTDRGNO
            .Col2 = enumWholeEntryGrid.COLUMN_CUSTDRGDESC
            .BlockMode = True
            .BackColor = System.Drawing.Color.Red
            .BlockMode = False
            mintPrevMarkRow = pintRowNo
            .Col = enumWholeEntryGrid.COLUMN_CUSTDRGNO
            .Action = FPSpreadADO.ActionConstants.ActionGotoCell
        End With
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub txtItemCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtItemCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '----------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/12/2003
        'Arguments           -
        'Return Value        -
        'Function            - To call Validate Item Code
        '----------------------------------------------------
        If KeyAscii = 13 Then
            Call txtLineNo_Validating(txtLineNo, New System.ComponentModel.CancelEventArgs(False))
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtItemCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtItemCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        '----------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/12/2003
        'Arguments           -
        'Return Value        -
        'Function            - To call Validate Item Code
        '----------------------------------------------------
        On Error GoTo ErrHandler
        If blnCheckMsg1 = True Then
            txtItemCode.Focus()
            blnCheckMsg1 = False
            Exit Sub
        End If
        If Not Me.CmdGrpMktSchedule.GetActiveButton Is Nothing Then
            If CmdGrpMktSchedule.GetActiveButton.Text.ToUpper = "print".ToUpper Then
                Exit Sub
            End If
        End If
        If Not ValidateEntry(Trim(txtCustDrgNo.Text), "C") Then
            MsgBox("Invalid Item Code", MsgBoxStyle.Information, "empower")
            txtItemCode.Focus()
            blnCheckMsg1 = True
            Cancel = True
        Else
            If CmdGrpMktSchedule.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                spdSingleEntryGrid.Row = 1
                spdSingleEntryGrid.Col = enumSingleEntryGrid.COLUMN_SCHEDULE_QTY
                spdSingleEntryGrid.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                spdSingleEntryGrid.Focus()
            Else
                blnCheckMsg1 = True
                txtItemCode.Focus()
            End If
        End If
        GoTo EventExitSub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Function MarkRowAsSaved(ByVal pintRowNo As Short) As Boolean
        '----------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/12/2003
        'Arguments           -
        'Return Value        -
        'Function            - To Mark Row As Saved After Saving
        '----------------------------------------------------
        On Error GoTo ErrHandler
        With spdWholeEntryGrid
            .Row = pintRowNo
            .Row2 = pintRowNo
            .Col = enumWholeEntryGrid.COLUMN_CUSTDRGNO
            .Col2 = enumWholeEntryGrid.COLUMN_REMARKS
            .BlockMode = True
            .ForeColor = System.Drawing.Color.Blue
            .BlockMode = False
        End With
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function SaveDataInWholeGrid(ByVal pintRowNo As Short) As Boolean
        '----------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/12/2003
        'Arguments           -
        'Return Value        -
        'Function            - To Save Data in Whole Grid
        '----------------------------------------------------
        Dim lintLoop As Short
        Dim strScheduleQty As String
        Dim strRemarks As String
        On Error GoTo ErrHandler
        strScheduleQty = ""
        strRemarks = ""
        With spdWholeEntryGrid
            .Row = pintRowNo
            For lintLoop = 1 To spdSingleEntryGrid.MaxRows
                spdSingleEntryGrid.Row = lintLoop
                spdSingleEntryGrid.Col = enumSingleEntryGrid.COLUMN_SCHEDULE_QTY
                strScheduleQty = strScheduleQty & spdSingleEntryGrid.Text & "�"
                spdSingleEntryGrid.Col = enumSingleEntryGrid.COLUMN_REMARKS
                strRemarks = strRemarks & spdSingleEntryGrid.Text & "�"
            Next
            .Col = enumWholeEntryGrid.COLUMN_SCHEDULEQTY
            .Text = strScheduleQty
            .Col = enumWholeEntryGrid.COLUMN_REMARKS
            .Text = strRemarks
        End With
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function DisablePrimaryField(ByVal pblnMakeDisable As Boolean) As Boolean
        '----------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/12/2003
        'Arguments           -
        'Return Value        -
        'Function            - To Disable Primery Key Feilds
        '----------------------------------------------------
        On Error GoTo ErrHandler
        Dim lDateStart As Integer
        Dim lDateEnd As Integer
        Dim lDateServer As Integer
        lDateStart = CInt(Year(Me.DTPSheduleStartDate.Value) & IIf(Len(Month(Me.DTPSheduleStartDate.Value)) = 1, "0" & Month(Me.DTPSheduleStartDate.Value), Month(Me.DTPSheduleStartDate.Value)))
        lDateEnd = CInt(Year(Me.DTPSheduleEndMonth.Value) & IIf(Len(Month(Me.DTPSheduleEndMonth.Value)) = 1, "0" & Month(Me.DTPSheduleEndMonth.Value), Month(Me.DTPSheduleEndMonth.Value)))
        lDateServer = CInt(Year(CDate(mstrServerDate)) & IIf(Len(Month(CDate(mstrServerDate))) = 1, "0" & Month(CDate(mstrServerDate)), Month(CDate(mstrServerDate))))
        DisablePrimaryField = True
        If CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            'Changes Done by Sourabh
            If lDateEnd < lDateServer Then
                MsgBox("Schedule End Month can not be smaller than Current Month", MsgBoxStyle.Information, "empower")
                DTPSheduleEndMonth.Focus()
                DisablePrimaryField = False
                Exit Function
            ElseIf lDateStart < lDateServer Then
                MsgBox("Schedule Start Month can not be smaller than Current Month", MsgBoxStyle.Information, "empower")
                DTPSheduleStartDate.Focus()
                DisablePrimaryField = False
                Exit Function
            ElseIf lDateStart > lDateEnd Then
                MsgBox("Schedule End Month can not be smaller than Schedule Start Month", MsgBoxStyle.Information, "empower")
                DTPSheduleEndMonth.Focus()
                DisablePrimaryField = False
                Exit Function
            End If
        Else
            If lDateStart > lDateEnd Then
                MsgBox("Schedule End Month can not be smaller than Schedule Start Month", MsgBoxStyle.Information, "empower")
                DTPSheduleEndMonth.Focus()
                DisablePrimaryField = False
                Exit Function
            End If
        End If
        If pblnMakeDisable Then
            If Trim(txtCustomerCode.Text) = "" Then
                MsgBox("Customer Code Can Not Be Blank", MsgBoxStyle.Information, "empower")
                DisablePrimaryField = False
                Exit Function
            End If
            DTPSheduleEndMonth.Enabled = False
            DTPSheduleStartDate.Enabled = False
            '''Changes done By ashutosh on 24-10-2006, issue Id:18623
            '''txtCustomerCode.Enabled = False
        End If
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function SaveDataInTable() As Boolean
        '----------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/12/2003
        'Arguments           -
        'Return Value        -
        'Function            - To save data in Table
        'History             - Changes done by Sourabh on 02 Sep 2004 for Add Two in Column DSNo and DSDateTime
        '----------------------------------------------------
        Dim strSql As String
        Dim strMonthYear As String
        Dim strInd_MonthYear() As String
        Dim lintLoop As Short
        '---- Four New Variable add by Sourabh on 02 Sep 2004 For Add two new column DSNO and DSDATETIME
        Dim strDSNo As String
        'Dim dtDatetime As Date
        Dim rsDSTracking As New ClsResultSetDB
        Dim blnDSTracking As Boolean
        Dim strForecasteMstQry As String
        Dim BlnUpdateForecast As Boolean
        Dim strDate As String
        Dim RsForecast As ClsResultSetDB
        On Error GoTo ErrHandler
        rsDSTracking.GetResult("select isnull(Update_Forecast,0) as Update_Forecast From Sales_Parameter where UNIT_CODE='" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        BlnUpdateForecast = rsDSTracking.GetValue("Update_Forecast")
        rsDSTracking.ResultSetClose()
        strForecasteMstQry = ""
        strSql = ""
        SaveDataInTable = True
        strMonthYear = GetMonthCondition("MY")
        strInd_MonthYear = Split(strMonthYear, ",")
        If Trim(strMonthYear) = "" Then
            MsgBox("Could Not Get Month Year ", MsgBoxStyle.Critical, "empower")
            Exit Function
        End If
        'Code add by Sourabh for Add Two new column DSNo and DSDateTime
        '------------------
        strDSNo = "ECSS"
        ': dtDatetime = CDate(GetServerDate() & vbCrLf & TimeOfDay)
        rsDSTracking = New ClsResultSetDB
        Call rsDSTracking.GetResult("Select DSWiseTracking From Sales_parameter where UNIT_CODE='" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsDSTracking.RowCount > 0 Then blnDSTracking = IIf(IsDBNull(rsDSTracking.GetValue("DSWiseTracking")), False, IIf(rsDSTracking.GetValue("DSwisetracking") = False, False, True))
        rsDSTracking.ResultSetClose()
        rsDSTracking = Nothing
        ' -----------------
        'mP_Connection.BeginTrans() 'rr00rr
        If CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            With spdSingleEntryGrid
                For lintLoop = 1 To .MaxRows
                    .Row = lintLoop
                    .Col = enumSingleEntryGrid.COLUMN_SCHEDULE_QTY
                    If ISSavedEntry(Val(txtLineNo.Text)) Then
                        strSql = strSql & " UPDATE monthlymktschedule SET Schedule_Qty=" & Val(.Text) & ","
                        .Col = enumSingleEntryGrid.COLUMN_REMARKS
                        strSql = strSql & " Remarks='"
                        strSql = strSql & Trim(.Text) & "'"
                        strSql = strSql & " WHERE UNIT_CODE='" & gstrUNITID & "' and Item_Code='" & Trim(txtItemCode.Text) & "' AND "
                        strSql = strSql & "Cust_Drgno='" & Trim(txtCustDrgNo.Text) & "' AND "
                        strSql = strSql & "RevisionNo=0 AND Status=1 AND Year_Month='"
                        strSql = strSql & strInd_MonthYear(lintLoop - 1) & "' AND Account_Code='" & Trim(txtCustomerCode.Text) & "'"
                    Else
                        If Val(.Text) <> 0 Then
                            'Changed for Issue ID eMpro-20090227-27987-(Added Consignee_code) Starts
                            strSql = strSql & "INSERT INTO monthlymktschedule (UNIT_CODE,Year_Month,Account_Code,Consignee_Code,"
                            strSql = strSql & "Item_Code,Cust_Drgno,Schedule_Flag,"
                            strSql = strSql & "Schedule_Qty,Authorized_Code,Despatch_qty,Remarks,"
                            strSql = strSql & "Ent_dt,Ent_UserId,status,RevisionNo,Upd_dt,Upd_UserId"
                            'Code add by Sourabh on 02 Sep 2004
                            If blnDSTracking = True Then
                                strSql = strSql & ",DSNo,DSDateTime"
                            End If
                            strSql = strSql & ") VALUES('"
                            strSql = strSql & gstrUNITID & "','"
                            strSql = strSql & strInd_MonthYear(lintLoop - 1) & "','" & Trim(txtCustomerCode.Text) & "','" & Trim(txtCustomerCode.Text) & "','"
                            strSql = strSql & Trim(txtItemCode.Text) & "','" & Trim(txtCustDrgNo.Text) & "',1,"
                            strSql = strSql & Val(.Text) & ",'" & mP_User & "',0,'"
                            .Col = enumSingleEntryGrid.COLUMN_REMARKS
                            strSql = strSql & Trim(.Text) & "',getdate(),'" & mP_User & "',1,0,getdate(),'" & mP_User & "'"
                            If blnDSTracking = True Then
                                strSql = strSql & ",'" & strDSNo & "',getdate()"
                            End If
                            strSql = strSql & ")"
                            strSql = strSql & vbCrLf
                            If BlnUpdateForecast Then
                                .Col = enumSingleEntryGrid.COLUMN_SCHEDULE_QTY
                                If strInd_MonthYear(lintLoop - 1) = VB6.Format(GetServerDate, "YYYYMM") Then
                                    strForecasteMstQry = strForecasteMstQry & "INSERT INTO forecast_mst(UNIT_CODE,Customer_code,product_no,Due_date,Quantity,ent_userid,ent_dt,upd_userid,upd_dt,Enagare_UNLOC ) "
                                    strForecasteMstQry = strForecasteMstQry & " VALUES('" & gstrUNITID & "','" & Trim(txtCustomerCode.Text) & "','" & Trim(txtItemCode.Text) & "',convert(varchar(10),getdate(),103), " & Val(.Text) & " , '" & mP_User & "',getdate(),'" & mP_User & "' , getdate(),'MSCH') "
                                Else
                                    strForecasteMstQry = strForecasteMstQry & "INSERT INTO forecast_mst(UNIT_CODE,Customer_code,product_no,Due_date,Quantity,ent_userid,ent_dt,upd_userid,upd_dt,Enagare_UNLOC ) "
                                    strForecasteMstQry = strForecasteMstQry & " VALUES('" & gstrUNITID & "','" & Trim(txtCustomerCode.Text) & "','" & Trim(txtItemCode.Text) & "',convert(varchar(10),'01/" & VB.Right(Trim(strInd_MonthYear(lintLoop - 1)), 2) & "/" & VB.Left(strInd_MonthYear(lintLoop - 1), 4) & "',103)," & Val(.Text) & " , '" & mP_User & "',getdate(),'" & mP_User & "' , getdate(),'MSCH') "
                                End If
                                strForecasteMstQry = strForecasteMstQry & vbCrLf
                            End If
                        End If
                    End If
                Next
            End With
        ElseIf CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
            With spdSingleEntryGrid
                For lintLoop = 1 To .MaxRows
                    .Row = lintLoop
                    .Col = enumSingleEntryGrid.COLUMN_SCHEDULE_QTY
                    If ISSavedEntry(Val(txtLineNo.Text)) Then
                        strSql = strSql & " UPDATE monthlymktschedule SET Schedule_Qty=" & Val(.Text) & ","
                        .Col = enumSingleEntryGrid.COLUMN_REMARKS
                        strSql = strSql & " Remarks='"
                        strSql = strSql & Trim(.Text) & "'"
                        strSql = strSql & " WHERE UNIT_CODE='" & gstrUNITID & "' AND Item_Code='" & Trim(txtItemCode.Text) & "' AND "
                        strSql = strSql & "Cust_Drgno='" & Trim(txtCustDrgNo.Text) & "' AND "
                        strSql = strSql & "Status=1 AND Year_Month='"
                        strSql = strSql & strInd_MonthYear(lintLoop - 1) & "' AND Account_Code='" & Trim(txtCustomerCode.Text) & "'"
                    Else
                        'If val(.Text) <> 0 Then
                        strSql = strSql & " UPDATE monthlymktschedule SET status=0 WHERE "
                        strSql = strSql & " UNIT_CODE='" & gstrUNITID & "' AND Item_Code='" & Trim(txtItemCode.Text) & "' AND "
                        strSql = strSql & "Cust_Drgno='" & Trim(txtCustDrgNo.Text) & "' AND "
                        strSql = strSql & "Year_Month='"
                        strSql = strSql & strInd_MonthYear(lintLoop - 1) & "' AND Account_Code='" & Trim(txtCustomerCode.Text) & "'" & vbCrLf
                        'Changed for Issue ID eMpro-20090227-27987-(Added Consignee Code) Starts
                        strSql = strSql & "INSERT INTO monthlymktschedule (UNIT_CODE,Year_Month,Account_Code,Consignee_Code,"
                        strSql = strSql & "Item_Code,Cust_Drgno,Schedule_Flag,"
                        strSql = strSql & "Schedule_Qty,Authorized_Code,Despatch_qty,Remarks,"
                        strSql = strSql & "Ent_dt,Ent_UserId,status,RevisionNo,Upd_dt,Upd_UserId"
                        'Code add by Sourabh on 02 Sep 2004
                        If blnDSTracking = True Then
                            strSql = strSql & ",DSNo,DSDateTime"
                        End If
                        strSql = strSql & ") VALUES('"
                        strSql = strSql & gstrUNITID & "','"
                        strSql = strSql & strInd_MonthYear(lintLoop - 1) & "','" & Trim(txtCustomerCode.Text) & "','" & Trim(txtCustomerCode.Text) & "','"
                        strSql = strSql & Trim(txtItemCode.Text) & "','" & Trim(txtCustDrgNo.Text) & "',1,"
                        strSql = strSql & Val(.Text) & ",'" & mP_User & "',"
                        .Col = enumSingleEntryGrid.COLUMN_DESPATCH_QTY
                        strSql = strSql & Val(.Text) & ",'"
                        .Col = enumSingleEntryGrid.COLUMN_REMARKS
                        strSql = strSql & Trim(.Text) & "',getdate(),'" & mP_User & "',1,"
                        .Col = enumSingleEntryGrid.COLUMN_REVISIONCOUNT
                        strSql = strSql & Val(.Text) & ",getdate(),'" & mP_User & "'"
                        'Code add by Sourabh on 02 Sep 2004
                        If blnDSTracking = True Then
                            strSql = strSql & ",'" & strDSNo & "',getdate()"
                        End If
                        strSql = strSql & ")"
                        'Changed for Issue ID eMpro-20090227-27987-(Added Consignee Code) Ends
                        strSql = strSql & vbCrLf
                        strSql = strSql & " UPDATE monthlymktschedule SET RevisionNo=RevisionNo+1 "
                        strSql = strSql & " WHERE UNIT_CODE='" & gstrUNITID & "' AND Item_Code='" & Trim(txtItemCode.Text) & "' AND "
                        strSql = strSql & "Cust_Drgno='" & Trim(txtCustDrgNo.Text) & "' AND "
                        strSql = strSql & "Year_Month='"
                        strSql = strSql & strInd_MonthYear(lintLoop - 1) & "' AND Account_Code='" & Trim(txtCustomerCode.Text) & "' AND Status=1" & vbCrLf
                        ''----Added by Davinder on 28/03/2006 (Issue ID:-17378) to save data also in Forecast_Mst
                        If BlnUpdateForecast = True Then
                            RsForecast = New ClsResultSetDB
                            RsForecast.GetResult("Select Customer_code from Forecast_Mst where UNIT_CODE='" & gstrUNITID & "' AND Customer_code='" & Trim(txtCustomerCode.Text) & "' AND year(Due_date) ='" & Mid(Trim(strInd_MonthYear(lintLoop - 1)), 1, 4) & "' AND right('0' + ltrim(month(Due_date)),2) ='" & VB.Right(Trim(strInd_MonthYear(lintLoop - 1)), 2) & "' AND product_no='" & Trim(txtItemCode.Text) & "' AND ENagare_UNLOC='MSCH'")
                            If RsForecast.GetNoRows > 0 Then
                                .Col = enumSingleEntryGrid.COLUMN_SCHEDULE_QTY
                                strForecasteMstQry = strForecasteMstQry & "UPDATE forecast_mst set quantity= " & Val(.Text) & ",Upd_dt=Getdate(),upd_userid= '" & mP_User & "' where UNIT_CODE='" & gstrUNITID & "' AND Customer_code='" & Trim(txtCustomerCode.Text) & "' AND product_no='" & Trim(txtItemCode.Text) & "' AND ENagare_UNLOC='MSCH' AND year(Due_date) ='" & Mid(Trim(strInd_MonthYear(lintLoop - 1)), 1, 4) & "' AND right('0' + ltrim(month(Due_date)),2) ='" & VB.Right(Trim(strInd_MonthYear(lintLoop - 1)), 2) & "'"
                            Else
                                If strInd_MonthYear(lintLoop - 1) = VB6.Format(GetServerDate, "YYYYMM") Then
                                    strForecasteMstQry = strForecasteMstQry & "INSERT INTO forecast_mst(UNIT_CODE,Customer_code,product_no,Due_date,Quantity,ent_userid,ent_dt,upd_userid,upd_dt,Enagare_UNLOC ) "
                                    strForecasteMstQry = strForecasteMstQry & " VALUES('" & gstrUNITID & "','" & Trim(txtCustomerCode.Text) & "','" & Trim(txtItemCode.Text) & "',convert(varchar(10),getdate(),103), " & Val(.Text) & " , '" & mP_User & "',getdate(),'" & mP_User & "' , getdate(),'MSCH') "
                                Else
                                    strForecasteMstQry = strForecasteMstQry & "INSERT INTO forecast_mst(UNIT_CODE,Customer_code,product_no,Due_date,Quantity,ent_userid,ent_dt,upd_userid,upd_dt,Enagare_UNLOC ) "
                                    strForecasteMstQry = strForecasteMstQry & " VALUES('" & gstrUNITID & "','" & Trim(txtCustomerCode.Text) & "','" & Trim(txtItemCode.Text) & "',convert(varchar(10),'01/" & VB.Right(Trim(strInd_MonthYear(lintLoop - 1)), 2) & "/" & VB.Left(strInd_MonthYear(lintLoop - 1), 4) & "',103)," & Val(.Text) & " , '" & mP_User & "',getdate(),'" & mP_User & "' , getdate(),'MSCH') "
                                End If
                            End If
                            strForecasteMstQry = strForecasteMstQry & vbCrLf
                            RsForecast.ResultSetClose()
                        End If
                        ''----Changes by Davnder End's here
                    End If
                Next
            End With
        End If
        If Trim(strSql) <> "" Then
            mP_Connection.Execute(strSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            If Trim(strForecasteMstQry) <> "" Then
                mP_Connection.Execute(strForecasteMstQry, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End If
        Else
            SaveDataInTable = False
        End If
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        MsgBox("Record Not Saved", MsgBoxStyle.Critical, "empower")
        mP_Connection.RollbackTrans()
        SaveDataInTable = False
    End Function
    Private Function ISSavedEntry(ByVal pintRow As Short) As Boolean
        '----------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/12/2003
        'Arguments           -
        'Return Value        -
        'Function            - To Check Weither Row is Saved
        '----------------------------------------------------
        On Error GoTo ErrHandler
        With spdWholeEntryGrid
            .Row = pintRow
            .Row2 = pintRow
            .Col = enumWholeEntryGrid.COLUMN_CUSTDRGNO
            .Col2 = enumWholeEntryGrid.COLUMN_CUSTDRGNO
            .BlockMode = True
            If .ForeColor.Equals(System.Drawing.Color.Blue) Then
                .BlockMode = False
                ISSavedEntry = True
            Else
                .BlockMode = False
                ISSavedEntry = False
            End If
        End With
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function ValidateEntry(ByVal pstrValue As String, ByVal pstrWhichValue As String) As Boolean
        '----------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/12/2003
        'Arguments           -
        'Return Value        -
        'Function            - To Call Validation Before Save Button Cloick
        '----------------------------------------------------
        Dim lintloopCnt As Short
        On Error GoTo ErrHandler
        ValidateEntry = False
        Select Case UCase(Trim(pstrWhichValue))
            Case "C"
                With spdWholeEntryGrid
                    For lintloopCnt = 1 To .MaxRows
                        .Row = lintloopCnt
                        .Col = enumWholeEntryGrid.COLUMN_CUSTDRGNO
                        If Trim(.Text) = pstrValue Then
                            ValidateEntry = True
                            mintCurrentNo = lintloopCnt
                            Call SetTextBoxValue(lintloopCnt)
                            If CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                                Call FillSingleGrid("ADD")
                            ElseIf CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                                Call FillSingleGrid("EDIT")
                            ElseIf CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                                Call FillSingleGrid("VIEW")
                            End If
                            Exit Function
                        End If
                    Next
                End With
            Case "I"
                With spdWholeEntryGrid
                    For lintloopCnt = 1 To .MaxRows
                        .Row = lintloopCnt
                        .Col = enumWholeEntryGrid.COLUMN_ITEMCODE
                        If Trim(.Text) = pstrValue Then
                            ValidateEntry = True
                            mintCurrentNo = lintloopCnt
                            Call SetTextBoxValue(lintloopCnt)
                            If CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                                Call FillSingleGrid("ADD")
                            ElseIf CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                                Call FillSingleGrid("EDIT")
                            ElseIf CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                                Call FillSingleGrid("VIEW")
                            End If
                            Exit Function
                        End If
                    Next
                End With
            Case "L"
                If Val(pstrValue) > 0 And Val(pstrValue) <= spdWholeEntryGrid.MaxRows Then
                    ValidateEntry = True
                    mintCurrentNo = Val(pstrValue)
                    Call SetTextBoxValue(Val(pstrValue))
                    If CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                        Call FillSingleGrid("ADD")
                    ElseIf CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                        Call FillSingleGrid("EDIT")
                    ElseIf CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                        Call FillSingleGrid("VIEW")
                    End If
                    Exit Function
                End If
        End Select
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        ValidateEntry = False
    End Function
    Private Sub txtLineNo_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtLineNo.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        '----------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/12/2003
        'Arguments           -
        'Return Value        -
        'Function            - To Validate Line No in Single Row Save Options
        '----------------------------------------------------
        On Error GoTo ErrHandler
        If valid_line_flag = True Then
            txtLineNo.Focus()
            valid_line_flag = False
            Exit Sub
        End If
        If Not ValidateEntry(Trim(txtLineNo.Text), "L") Then
            MsgBox("Invalid Line No.", MsgBoxStyle.Information, "empower")
            valid_line_flag = True
            txtLineNo.Focus()
            Cancel = True
        Else
            If CmdGrpMktSchedule.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                spdSingleEntryGrid.Row = 1
                spdSingleEntryGrid.Col = enumSingleEntryGrid.COLUMN_SCHEDULE_QTY
                spdSingleEntryGrid.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                spdSingleEntryGrid.Focus()
            Else
                txtLineNo.Focus()
            End If
        End If
        GoTo EventExitSub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtLineNo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtLineNo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '----------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/12/2003
        'Arguments           -
        'Return Value        -
        'Function            - To call Validate line No
        '----------------------------------------------------
        If KeyAscii = 13 Then
            Call txtLineNo_Validating(txtLineNo, New System.ComponentModel.CancelEventArgs(False))
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Public Function InsertdatainCaseofWholeGrid() As Boolean
        '----------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/12/2003
        'Arguments           -
        'Return Value        -
        'Function            - To Insert data in Case of Whole Grid
        'History            :   Changes done by Sourabh on 02 Sep 2004 For Add two new column DSNO and DSDATETIME
        '----------------------------------------------------
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        Dim strSql As String
        Dim strMonthYear As String
        Dim strInd_MonthYear() As String
        Dim strItemCode As String
        Dim strCustDrgNo As String
        Dim dblScheduleQuantity As Double
        Dim strRemarks As String
        Dim dblDespatchQty As Double
        Dim intRevisionNo As Short
        Dim intInnerLoopCount As Short
        Dim intInnerMaxLoopCount As Short
        '---- Four New Variable add by Sourabh on 02 Sep 2004 For Add two new column DSNO and DSDATETIME
        Dim strDSNo As String
        'Dim dtDatetime As Date
        Dim rsDSTracking As New ClsResultSetDB
        Dim blnDSTracking As Boolean
        On Error GoTo ErrHandler
        With spdWholeEntryGrid
            InsertdatainCaseofWholeGrid = False
            strSql = ""
            intMaxLoop = .MaxRows
            'Code add by Sourabh for Add Two new column DSNo and DSDateTime
            '------------------
            strDSNo = "ECSS"
            ': dtDatetime = CDate(GetServerDate() & vbCrLf & TimeOfDay)
            Call rsDSTracking.GetResult("Select DSWiseTracking From Sales_parameter WHERE UNIT_CODE='" & gstrUNITID & "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If rsDSTracking.RowCount > 0 Then blnDSTracking = IIf(IsDBNull(rsDSTracking.GetValue("DSWiseTracking")), False, IIf(rsDSTracking.GetValue("DSwisetracking") = False, False, True))
            rsDSTracking.ResultSetClose()
            rsDSTracking = Nothing
            ' -----------------
            For intLoopCounter = 1 To intMaxLoop
                strMonthYear = GetMonthCondition("MY")
                strInd_MonthYear = Split(strMonthYear, ",")
                If Trim(strMonthYear) = "" Then
                    MsgBox("Could Not Get Month Year ", MsgBoxStyle.Critical, "empower")
                    InsertdatainCaseofWholeGrid = False
                    Exit Function
                End If
                If CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                    intInnerMaxLoopCount = UBound(strInd_MonthYear)
                    For intInnerLoopCount = 0 To intInnerMaxLoopCount
                        .Row = intLoopCounter
                        .Col = enumWholeEntryGrid.COLUMN_ITEMCODE
                        strItemCode = Trim(.Text)
                        .Col = enumWholeEntryGrid.COLUMN_CUSTDRGNO
                        strCustDrgNo = Trim(.Text)
                        .Col = enumWholeEntryGrid.COLUMN_SCHEDULEQTY + intInnerLoopCount
                        dblScheduleQuantity = Val(.Text)
                        .Col = enumWholeEntryGrid.COLUMN_REMARKS
                        strRemarks = Trim(.Text)
                        If dblScheduleQuantity <> 0 Then
                            'Changed for Issue ID eMpro-20090227-27987 Starts
                            strSql = strSql & "INSERT INTO monthlymktschedule (UNIT_CODE,Year_Month,Account_Code,Consignee_Code,"
                            strSql = strSql & "Item_Code,Cust_Drgno,Schedule_Flag,"
                            strSql = strSql & "Schedule_Qty,Authorized_Code,Despatch_qty,Remarks,"
                            strSql = strSql & "Ent_dt,Ent_UserId,status,RevisionNo,Upd_dt,Upd_UserId"
                            'Code add by Sourabh on 02 Sep 2004
                            If blnDSTracking = True Then
                                strSql = strSql & ",DSNo,DSDateTime"
                            End If
                            strSql = strSql & ") VALUES('"
                            strSql = strSql & gstrUNITID & "','"
                            strSql = strSql & strInd_MonthYear(intInnerLoopCount) & "','" & Trim(txtCustomerCode.Text) & "','" & Trim(txtCustomerCode.Text) & "','"
                            strSql = strSql & Trim(strItemCode) & "','" & Trim(strCustDrgNo) & "',1,"
                            strSql = strSql & dblScheduleQuantity & ",'" & mP_User & "',0,'"
                            strSql = strSql & strRemarks & "',getdate(),'" & mP_User & "',1,0,getdate(),'" & mP_User & "'"
                            'Code add by Sourabh on 02 Sep 2004
                            If blnDSTracking = True Then
                                strSql = strSql & ",'" & strDSNo & "',getdate()"
                            End If
                            strSql = strSql & ")"
                            'Changed for Issue ID eMpro-20090227-27987 Ends
                            strSql = strSql & vbCrLf
                        End If
                    Next
                ElseIf CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                    intInnerMaxLoopCount = UBound(strInd_MonthYear)
                    For intInnerLoopCount = 0 To intInnerMaxLoopCount
                        .Row = intLoopCounter
                        .Col = enumWholeEntryGrid.COLUMN_ITEMCODE
                        strItemCode = Trim(.Text)
                        .Col = enumWholeEntryGrid.COLUMN_CUSTDRGNO
                        strCustDrgNo = Trim(.Text)
                        .Col = enumWholeEntryGrid.COLUMN_SCHEDULEQTY + (intInnerLoopCount * 3)
                        dblScheduleQuantity = Val(.Text)
                        '.Col = enumWholeEntryGrid.COLUMN_REMARKS
                        'strRemarks = Trim(.Text)
                        .Col = enumWholeEntryGrid.COLUMN_DESPATCHQTY
                        dblDespatchQty = Val(.Text)
                        .Col = enumWholeEntryGrid.COLUMN_REVISION_NO + (intInnerLoopCount * 3)
                        intRevisionNo = Val(.Text) + 1
                        strSql = strSql & " UPDATE monthlymktschedule SET status=0 WHERE "
                        strSql = strSql & "UNIT_CODE='" & gstrUNITID & "' AND Item_Code='" & strItemCode & "' AND "
                        strSql = strSql & "Cust_Drgno='" & strCustDrgNo & "' AND "
                        strSql = strSql & "Year_Month='"
                        strSql = strSql & strInd_MonthYear(intInnerLoopCount) & "' AND Account_Code='" & Trim(txtCustomerCode.Text) & "'" & vbCrLf
                        'Changed for Issue ID eMpro-20090227-27987 Starts
                        strSql = strSql & "INSERT INTO monthlymktschedule (UNIT_CODE,Year_Month,Account_Code,Consignee_Code,"
                        strSql = strSql & "Item_Code,Cust_Drgno,Schedule_Flag,"
                        strSql = strSql & "Schedule_Qty,Authorized_Code,Despatch_qty,Remarks,"
                        strSql = strSql & "Ent_dt,Ent_UserId,status,RevisionNo,Upd_dt,Upd_UserId"
                        'Code add by Sourabh on 02 Sep 2004
                        If blnDSTracking = True Then
                            strSql = strSql & ",DSNo,DSDateTime"
                        End If
                        strSql = strSql & ") VALUES('"
                        strSql = strSql & gstrUNITID & "','"
                        strSql = strSql & strInd_MonthYear(intInnerLoopCount) & "','" & Trim(txtCustomerCode.Text) & "','" & Trim(txtCustomerCode.Text) & "','"
                        strSql = strSql & strItemCode & "','" & strCustDrgNo & "',1,"
                        strSql = strSql & dblScheduleQuantity & ",'" & mP_User & "',"
                        strSql = strSql & dblDespatchQty & ",'"
                        strSql = strSql & strRemarks & "',getdate(),'" & mP_User & "',1,"
                        strSql = strSql & intRevisionNo & ",getdate(),'" & mP_User & "'"
                        'Code add by Sourabh on 02 Sep 2004
                        If blnDSTracking = True Then
                            strSql = strSql & ",'" & strDSNo & "',getdate()"
                        End If
                        strSql = strSql & ")"
                        'Changed for Issue ID eMpro-20090227-27987 Ends
                        strSql = strSql & vbCrLf
                    Next
                End If
            Next
            If Len(Trim(strSql)) > 0 Then
                mP_Connection.BeginTrans()
                mP_Connection.Execute(strSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                InsertdatainCaseofWholeGrid = True
                'mP_Connection.CommitTrans() 'rr00rr
            End If
        End With
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        MsgBox("Record Not Saved", MsgBoxStyle.Critical, "empower")
        mP_Connection.RollbackTrans()
    End Function
    Public Sub EnableWholeGridinEditMode()
        '----------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/12/2003
        'Arguments           -
        'Return Value        -
        'Function            - To Enable Controle in Edit Mode
        '----------------------------------------------------
        Dim strMonthYear As String
        Dim strInd_MonthYear() As String
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        With spdWholeEntryGrid
            .Enabled = True
            intMaxLoop = .MaxCols
            For intLoopCounter = enumWholeEntryGrid.COLUMN_SCHEDULEQTY To intMaxLoop Step 3
                .Col = intLoopCounter : .Col2 = intLoopCounter : .Row = 1 : .Row2 = .MaxRows
                .BlockMode = True : spdWholeEntryGrid.Lock = False : .BlockMode = False
            Next
        End With
    End Sub
    Public Function ValidateinCaseWholeGrid() As Boolean
        '----------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/12/2003
        'Arguments           -
        'Return Value        -
        'Function            - To call Validations
        '----------------------------------------------------
        Dim dblscheduleqty As Double
        Dim dblDespatchQty As Double
        Dim strMonth_Year As String
        Dim strYear() As String
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        Dim intOuterLoopCount As Short
        Dim intOuterMaxLoop As Short
        Dim dblCol As Double
        On Error GoTo ErrHandler
        ValidateinCaseWholeGrid = False
        If mblnSingleSave = False Then
            If CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                strMonth_Year = GetMonthCondition("MY")
                strYear = Split(strMonth_Year, ",")
                With spdWholeEntryGrid
                    intOuterMaxLoop = .MaxRows
                    For intOuterLoopCount = 1 To intOuterMaxLoop
                        intMaxLoop = UBound(strYear)
                        For intLoopCounter = 0 To intMaxLoop
                            .Col = enumWholeEntryGrid.COLUMN_SCHEDULEQTY + (3 * intLoopCounter)
                            .Row = intOuterLoopCount
                            dblscheduleqty = Val(.Text)
                            .Col = enumWholeEntryGrid.COLUMN_DESPATCHQTY + (3 * intLoopCounter)
                            .Row = intOuterLoopCount
                            dblDespatchQty = Val(.Text)
                            If dblscheduleqty < dblDespatchQty Then
                                mblnDirty = True
                                MsgBox("Schedule Quantity Can Not Be Less Than Despatch Quantity", MsgBoxStyle.Information, "empower")
                                ValidateinCaseWholeGrid = False
                                .Col = enumWholeEntryGrid.COLUMN_SCHEDULEQTY + (3 * intLoopCounter)
                                .Row = intOuterLoopCount
                                .Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                .Focus()
                                Exit Function
                            End If
                        Next
                    Next
                End With
            End If
        End If
        ValidateinCaseWholeGrid = True
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function DecimalAllowed(ByVal pRow As Integer, ByVal pstrItem_code As String) As Boolean
        '*******************************************************************************
        'Author             :   Davinder
        'Argument(s)if any  :   Row of the grid,Item Code of the selected item
        'Return Value       :   Boolean
        'Function           :   To check if decimal places are allowed for that item's measurement code
        'Creation Date      :   30/03/2006
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim Rs As ClsResultSetDB
        Dim strQuery As String
        Rs = New ClsResultSetDB
        strQuery = "select I.Cons_measure_code as CMD, M.Decimal_Allowed_Flag as DAF,M.NoOFDecimal as NOD"
        strQuery = strQuery & " from item_mst I, Measure_Mst M"
        strQuery = strQuery & " Where i.Cons_measure_code = M.Measure_Code AND i.UNIT_CODE=m.UNIT_CODE AND"
        strQuery = strQuery & " I.Item_Code='" & pstrItem_code & "' AND i.UNIT_CODE='" & gstrUNITID & "'"
        Rs.GetResult(strQuery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If Not Convert.ToBoolean(Rs.GetValue("DAF").ToString) Then
            With Me.spdSingleEntryGrid
                .Row = pRow
                .Col = enumSingleEntryGrid.COLUMN_SCHEDULE_QTY
                If Trim(.Text) <> "" Then
                    If Fix(Convert.ToDouble(.Text)) < Val(.Text) Then
                        MsgBox("Quantity can not be in decimal places for Item: " & Trim(Me.txtItemCode.Text))
                        DecimalAllowed = True
                    End If
                End If
            End With
        Else
            DecimalAllowed = False
        End If
        Rs.ResultSetClose()
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
End Class