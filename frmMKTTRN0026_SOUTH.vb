Option Strict Off
Option Explicit On
Friend Class frmMKTTRN0026_SOUTH
	Inherits System.Windows.Forms.Form
	''----------------------------------------------------
	''(C) 2001 MIND, All rights reserved
	''Name of module - FRMMKTTRN0026.frm
	''Created by     - Tapan Jain
	''Created Date   - 12-12-2002
	''Description    - Monthly Marketing Schedule
	''Revised date   -
	''Revision History
	''----------------------------------------------------
	'Revised By     : Arul Mozhi
	'Revised On     : 25-11-2004
	'Description    : Order By Item Code is added in Select Item Qry
	'-------------------------------------------------------------------------
	'Revised By     : Arul Mozhi
	'Revised On     : 16-03-2005
    'Description    : Item Master Status flag checked in select Query of Save mode
    'Modified By Nitin Mehta on 12 May 2011
    'Modified to support MultiUnit functionality
    '-------------------------------------------------------------------------
    Dim mstrServerDate As String
    Dim strHelp() As String 'string array for showing values
    Dim mblnSingleSave As Boolean
    Dim mintCurrentNo As Short
    Dim mintMaxCurrentNo As Short
    Dim mintPrevMarkRow As Short
    Dim mblnDirty As Boolean
    Dim mintFormIndex As Short
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
    Private Sub cmdConsHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdConsHelp.Click
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
            strHelp = ctlEMPHelpMktSchedule.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "SELECT DISTINCT CONSIGNEE_CODE , Prty_Name FROM MonthlyMktSchedule ,Gen_PartyMaster WHERE CONSIGNEE_CODE=Prty_PartyID and MonthlyMktSchedule.UNIT_CODE=Gen_PartyMaster.Unt_CodeID AND MonthlyMktSchedule.UNIT_CODE='" & gstrUNITID & "'", "Consignee Code Help", 2)
        Else
            strHelp = ctlEMPHelpMktSchedule.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "SELECT DISTINCT CONSIGNEE_CODE , Prty_Name FROM Cust_Ord_Hdr ,Gen_PartyMaster WHERE CONSIGNEE_CODE=Prty_PartyID and Cust_Ord_Hdr.UNIT_CODE=Gen_PartyMaster.Unt_CodeID AND Cust_Ord_Hdr.UNIT_CODE='" & gstrUNITID & "'", "Consignee Code Help", 2)
        End If
        'Chnaging the mouse pointer
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        If UBound(strHelp) <> "0" Then
            If strHelp(0) <> "0" Then
                txtConsCode.Text = Trim(strHelp(0))
                lblConsDesc.Text = Trim(strHelp(1))
                Call txtConsCode_Validating(txtConsCode, New System.ComponentModel.CancelEventArgs(False))
            Else
                MsgBox("No Consignee Record Available.", MsgBoxStyle.Information, "empower")
                txtConsCode.Focus()
            End If
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdCustomerHelp_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCustomerHelp.Click
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
            strHelp = ctlEMPHelpMktSchedule.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "SELECT DISTINCT Account_code , Prty_Name FROM MonthlyMktSchedule ,Gen_PartyMaster WHERE Account_Code=Prty_PartyID and MonthlyMktSchedule.UNIT_CODE=Gen_PartyMaster.Unt_CodeID AND MonthlyMktSchedule.UNIT_CODE='" & gstrUNITID & "'", "Customer Code Help", 2)
        Else
            strHelp = ctlEMPHelpMktSchedule.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "SELECT DISTINCT Account_code , Prty_Name FROM Cust_Ord_Hdr ,Gen_PartyMaster WHERE Account_Code=Prty_PartyID and Cust_Ord_Hdr.UNIT_CODE=Gen_PartyMaster.Unt_CodeID AND Cust_Ord_Hdr.UNIT_CODE='" & gstrUNITID & "'", "Customer Code Help", 2)
        End If
        'Chnaging the mouse pointer
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
        If UBound(strHelp) <> "0" Then
            If strHelp(0) <> "0" Then
                txtCustomerCode.Text = Trim(strHelp(0))
                lblCustomerDesc.Text = Trim(strHelp(1))
                Call txtCustomerCode_Validating(txtCustomerCode, New System.ComponentModel.CancelEventArgs(False))
            Else
                MsgBox("No Record Available", MsgBoxStyle.Information, "empower")
                txtCustomerCode.Focus()
            End If
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    ' Private Sub CmdGrpMktSchedule_ButtonClick(ByVal eventSender As System.Object, ByVal eventArgs As AxActXCtl.__CmdGrp_ButtonClickEvent)
    Private Sub CmdGrpMktSchedule_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.UCbtndgrp.ButtonClickEventArgs) Handles CmdGrpMktSchedule.ButtonClick
        '---------------------------------------------------------------------------------------
        'Name       :   CmdGrpMktSchedule_ButtonClick
        'Type       :   Sub
        'Author     :   Tapan Jain
        'Arguments  :
        'Return     :
        'Purpose    :
        '---------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Select Case e.Button
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD 'Add Record
                Call RefreshForm()
                cmdSingleEntrySave.Enabled = True
                DTPSheduleStartDate.Value = mstrServerDate
                DTPSheduleEndMonth.Value = mstrServerDate
                mintCurrentNo = 1
                mintPrevMarkRow = 1
                mblnDirty = False
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE 'Save Record
                If DisablePrimaryField(True) Then
                    MsgBox("Transaction Completed Successfully", MsgBoxStyle.Information, "empower")
                    Call RefreshForm()
                    CmdGrpMktSchedule.Revert()
                    gblnCancelUnload = False
                    gblnFormAddEdit = False
                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL ' Cancel Click
                Call frmMKTTRN0026_SOUTH_KeyPress(Me, New System.Windows.Forms.KeyPressEventArgs(Chr(System.Windows.Forms.Keys.Escape)))
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT 'Clich on edit
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
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE 'To Verify Close Operation
                Me.Close()
        End Select
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdSingleEntrySave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSingleEntrySave.Click
        '---------------------------------------------------------------------------------------
        'Name       :   cmdSingleEntrySave_Click
        'Type       :   Sub
        'Author     :   Tapan Jain
        'Arguments  :
        'Return     :
        'Purpose    :
        '---------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        If CmdGrpMktSchedule.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
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
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub cmdSingleEntrySkip_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSingleEntrySkip.Click
        '---------------------------------------------------------------------------------------
        'Name       :   cmdSingleEntrySkip_Click
        'Type       :   Sub
        'Author     :   Tapan Jain
        'Arguments  :
        'Return     :
        'Purpose    :
        '---------------------------------------------------------------------------------------
        'If CmdGrpMktSchedule.mode <> MODE_VIEW Then
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
        'End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub ctlFormHeader1_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        On Error GoTo ErrHandler
        Call ShowHelp("HLP" & Mid(Me.Name, 4, Len(Me.Name)) & ".htm") '("pddrep_listofitems.htm")
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub DTPSheduleEndMonth_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSComCtl2.DDTPickerEvents_KeyDownEvent) Handles DTPSheduleEndMonth.KeyDownEvent
        If eventArgs.keyCode = System.Windows.Forms.Keys.Return Then System.Windows.Forms.SendKeys.Send(("{tab}"))
    End Sub
    Private Sub DTPSheduleStartDate_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSComCtl2.DDTPickerEvents_KeyDownEvent) Handles DTPSheduleStartDate.KeyDownEvent
        If eventArgs.keyCode = System.Windows.Forms.Keys.Return Then System.Windows.Forms.SendKeys.Send(("{tab}"))
    End Sub
    Private Sub frmMKTTRN0026_SOUTH_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo ErrHandler
        mdifrmMain.CheckFormName = mintFormIndex
        frmModules.NodeFontBold(Tag) = True
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub frmMKTTRN0026_SOUTH_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Tag) = False
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0026_SOUTH_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0 Then Call ctlFormHeader1_ClickEvent(ctlFormHeader1, New System.EventArgs()) : Exit Sub
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0026_SOUTH_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
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
    Private Sub frmMKTTRN0026_SOUTH_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Dim varIsRowWiseSaving() As Object
        On Error GoTo ErrHandler
        '-------------------------------------------------------
        mintFormIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader1.Tag)
        Call FitToClient(Me, Frame1, ctlFormHeader1, CmdGrpMktSchedule, 500) 'To fit the form in the MDI
        'Call SetGrid
        varIsRowWiseSaving = GetFieldsValues("SELECT RowSchSave FROM Sales_Parameter WHERE UNIT_CODE='" & gstrUNITID & "'", 1)
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
    Private Sub frmMKTTRN0026_SOUTH_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        On Error GoTo ErrHandler
        If gblnCancelUnload = True Then Cancel = 1
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub frmMKTTRN0026_SOUTH_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
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
        '---------------------------------------------------------------------------------------
        'Name       :   GetFieldsValues
        'Type       :   Function
        'Author     :   Tapan Jain
        'Arguments  :
        'Return     :
        'Purpose    :
        '---------------------------------------------------------------------------------------
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
        '---------------------------------------------------------------------------------------
        'Name       :   SetGrid
        'Type       :   Sub
        'Author     :   Tapan Jain
        'Arguments  :
        'Return     :
        'Purpose    :
        '---------------------------------------------------------------------------------------
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
                '.RowHeaderDisplay = DispNumbers
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
        '---------------------------------------------------------------------------------------
        'Name       :   RefreshForm
        'Type       :   Sub
        'Author     :   Tapan Jain
        'Arguments  :
        'Return     :
        'Purpose    :
        '---------------------------------------------------------------------------------------
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
    Private Sub txtConsCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtConsCode.TextChanged
        '*******************************************************************************
        'Author             :   Ashutosh Verma
        'Argument(s)if any  :   Issue Id: 19731
        'Return Value       :   NA
        'Function           :
        'Comments           :   NA
        'Creation Date      :   16 Apr 2007
        '*******************************************************************************
        On Error GoTo ErrHandler
        If Len(Trim(txtConsCode.Text)) = 0 Then
            Me.lblConsDesc.Text = ""
        End If
        If spdWholeEntryGrid.Visible = True Then
            spdWholeEntryGrid.MaxRows = 0
        End If
        If spdSingleEntryGrid.Visible = True Then
            spdSingleEntryGrid.MaxRows = 0
        End If
        If txtConsCode.Enabled = True Then
            txtConsCode.Focus()
        End If
        txtCustDrgNo.Text = ""
        txtItemCode.Text = ""
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtConsCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtConsCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*******************************************************************************
        'Author             :   Ashutosh Verma
        'Argument(s)if any  :   Issue Id: 19731
        'Return Value       :   NA
        'Function           :
        'Comments           :   NA
        'Creation Date      :   16 Apr 2007
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case 39, 34, 96
                KeyAscii = 0
        End Select
        GoTo EventExitSub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtConsCode_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtConsCode.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '*******************************************************************************
        'Author             :   Ashutosh Verma
        'Argument(s)if any  :   Issue Id: 19731
        'Return Value       :   NA
        'Function           :   If F1 Key Press Then Display Help From for Consignee code.
        'Comments           :   NA
        'Creation Date      :   16 Apr 2007
        '*******************************************************************************
        On Error GoTo ErrHandler
        If KeyCode = 112 Then Call cmdConsHelp_Click(cmdConsHelp, New System.EventArgs()) 'Help should be invoked if F1 key is pressed
        If KeyCode = System.Windows.Forms.Keys.Return Then
            Call txtConsCode_Validating(txtConsCode, New System.ComponentModel.CancelEventArgs(False))
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Sub
    Private Sub txtConsCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtConsCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        '*******************************************************************************
        'Author             :   Ashutosh Verma
        'Argument(s)if any  :   Issue Id: 19731
        'Return Value       :   NA
        'Function           :   Validate Consignee code.
        'Comments           :   NA
        'Creation Date      :   16 Apr 2007
        '*******************************************************************************
        Dim varGetCustName() As Object
        On Error GoTo ErrHandler
        If Trim(txtConsCode.Text) = "" Then
            lblConsDesc.Text = ""
            GoTo EventExitSub
        End If
        If CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            varGetCustName = GetFieldsValues("SELECT DISTINCT CONSIGNEE_CODE , Prty_Name FROM MonthlyMktSchedule ,Gen_PartyMaster WHERE CONSIGNEE_CODE=Prty_PartyID AND MonthlyMktSchedule.UNIT_CODE=Gen_PartyMaster.Unt_CodeID AND MonthlyMktSchedule.UNIT_CODE='" & gstrUNITID & "' AND CONSIGNEE_CODE='" & Trim(txtConsCode.Text) & "'", 2)
        Else
            varGetCustName = GetFieldsValues("SELECT DISTINCT CONSIGNEE_CODE , Prty_Name FROM Cust_Ord_Hdr ,Gen_PartyMaster WHERE CONSIGNEE_CODE=Prty_PartyID AND Cust_Ord_Hdr.UNIT_CODE=Gen_PartyMaster.Unt_CodeID AND Cust_Ord_Hdr.UNIT_CODE='" & gstrUNITID & "' AND CONSIGNEE_CODE='" & Trim(txtConsCode.Text) & "'", 2)
        End If
        If Trim(varGetCustName(0)) = Trim(txtConsCode.Text) Then
            lblConsDesc.Text = varGetCustName(1)
            If Me.txtCustomerCode.Text = "" Then
                MsgBox("Please enter Customer code.", MsgBoxStyle.Information, "eMpower")
                If txtCustomerCode.Enabled = True Then txtCustomerCode.Focus()
                GoTo EventExitSub
            End If
            Call FillWholeGrid()
        Else
            MsgBox("Invalid Consignee Code", MsgBoxStyle.Information, "empower")
            txtConsCode.Text = ""
            Cancel = True
            GoTo EventExitSub
        End If
        GoTo EventExitSub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtCustDrgNo_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCustDrgNo.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
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
        '---------------------------------------------------------------------------------------
        'Name       :   txtCustDrgNo_Validate
        'Type       :   Sub
        'Author     :   Tapan Jain
        'Arguments  :
        'Return     :
        'Purpose    :
        '---------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        If Not ValidateEntry(Trim(txtCustDrgNo.Text), "C") Then
            MsgBox("Invalid Customer Drawing No.", MsgBoxStyle.Information, "empower")
            txtCustDrgNo.Focus()
            Cancel = True
        Else
            If CmdGrpMktSchedule.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                spdSingleEntryGrid.Row = 1
                spdSingleEntryGrid.Col = enumSingleEntryGrid.COLUMN_SCHEDULE_QTY
                spdSingleEntryGrid.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                spdSingleEntryGrid.Focus()
            Else
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
        On Error GoTo ErrHandler
        If Trim(txtCustomerCode.Text) = "" Then
            lblCustomerDesc.Text = ""
        End If
        If spdWholeEntryGrid.Visible = True Then
            spdWholeEntryGrid.MaxRows = 0
        End If
        If spdSingleEntryGrid.Visible = True Then
            spdSingleEntryGrid.MaxRows = 0
        End If
        txtConsCode.Text = ""
        lblConsDesc.Text = ""
        txtCustDrgNo.Text = ""
        txtItemCode.Text = ""
        If txtCustomerCode.Enabled = True Then txtCustomerCode.Focus()
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub txtCustomerCode_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCustomerCode.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        If KeyCode = System.Windows.Forms.Keys.F1 Then Call cmdCustomerHelp_Click(cmdCustomerHelp, New System.EventArgs())
    End Sub
    Private Sub txtCustomerCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCustomerCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        If KeyAscii = 13 Then
            Call txtCustomerCode_Validating(txtCustomerCode, New System.ComponentModel.CancelEventArgs(False))
            If Me.txtConsCode.Enabled = True Then Me.txtConsCode.Focus()
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub txtCustomerCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtCustomerCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        '---------------------------------------------------------------------------------------
        'Name       :   txtCustomerCode_Validate
        'Type       :   Sub
        'Author     :   Tapan Jain
        'Arguments  :
        'Return     :
        'Purpose    :
        '---------------------------------------------------------------------------------------
        Dim varGetCustName() As Object
        On Error GoTo ErrHandler
        If Trim(txtCustomerCode.Text) = "" Then
            CmdGrpMktSchedule.Focus()
            GoTo EventExitSub
        End If
        If CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            varGetCustName = GetFieldsValues("SELECT DISTINCT Account_code , Prty_Name FROM MonthlyMktSchedule ,Gen_PartyMaster WHERE Account_Code=Prty_PartyID and MonthlyMktSchedule.UNIT_CODE=Gen_PartyMaster.Unt_CodeID AND MonthlyMktSchedule.UNIT_CODE='" & gstrUNITID & "' AND Account_code='" & Trim(txtCustomerCode.Text) & "'", 2)
        Else
            varGetCustName = GetFieldsValues("SELECT DISTINCT Account_code , Prty_Name FROM Cust_Ord_Hdr ,Gen_PartyMaster WHERE Account_Code=Prty_PartyID AND Cust_Ord_Hdr.UNIT_CODE=Gen_PartyMaster.Unt_CodeID AND Cust_Ord_Hdr.UNIT_CODE='" & gstrUNITID & "' AND Account_code='" & Trim(txtCustomerCode.Text) & "'", 2)
        End If
        If Trim(varGetCustName(0)) = Trim(txtCustomerCode.Text) Then
            lblCustomerDesc.Text = varGetCustName(1)
        Else
            MsgBox("Invalid Customer Code", MsgBoxStyle.Information, "empower")
            txtCustomerCode.Text = ""
            Cancel = True
            GoTo EventExitSub
        End If
        GoTo EventExitSub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Function FillWholeGrid() As Boolean
        '---------------------------------------------------------------------------------------
        'Name       :   FillWholeGrid
        'Type       :   Function
        'Author     :   Tapan Jain
        'Arguments  :
        'Return     :
        'Purpose    :
        '---------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        If mblnSingleSave Then
            If CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                If Not DisablePrimaryField(True) Then Exit Function
                Call SetGrid()
                Call GetItemList(Trim(txtCustomerCode.Text), Trim(txtConsCode.Text), "ADD")
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
                Call GetItemList(Trim(txtCustomerCode.Text), Trim(txtConsCode.Text), "VIEW")
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
        Else
            If CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                If Not DisablePrimaryField(True) Then Exit Function
                Call SetGrid()
                Call GetItemList(Trim(txtCustomerCode.Text), Trim(txtConsCode.Text), "ADD")
            ElseIf CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                If Not DisablePrimaryField(False) Then Exit Function
                Call SetGrid()
                Call GetItemList(Trim(txtCustomerCode.Text), Trim(txtConsCode.Text), "VIEW")
                CmdGrpMktSchedule.Focus()
            End If
        End If
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function GetItemList(ByVal pstrCustomerCode As String, ByVal pstrConsCode As String, ByVal pstrMode As String) As Boolean
        '---------------------------------------------------------------------------------------
        'Name       :   GetItemList
        'Type       :   Function
        'Author     :   Tapan Jain
        'Arguments  :
        'Return     :
        'Purpose    :
        '---------------------------------------------------------------------------------------
        Dim rsGetItem As New ADODB.Recordset
        Dim strSQL As String
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
        strSQL = ""
        If UCase(Trim(pstrMode)) = UCase("ADD") Then
            'Code Changed By Arul on 25-11-2004 Order by Added in item select query
            strSQL = "SELECT DISTINCT Cust_ord_dtl.Cust_Drgno,CustItem_mst.Item_Code,Drg_Desc FROM Cust_ord_dtl,CustItem_Mst, item_mst "
            strSQL = strSQL & " WHERE NOT EXISTS (SELECT * FROM monthlymktschedule WHERE account_code='" & pstrCustomerCode & "' and CONSIGNEE_CODE='" & pstrConsCode & "' "
            strSQL = strSQL & " AND monthlymktschedule.Cust_Drgno = Cust_ord_dtl.Cust_Drgno AND monthlymktschedule.UNIT_CODE = Cust_ord_dtl.UNIT_CODE AND monthlymktschedule.UNIT_CODE='" & gstrUNITID & "'"
            strSQL = strSQL & " AND monthlymktschedule.Item_Code = Cust_ord_dtl.Item_Code  AND year_month IN (" & GetMonthCondition("MY") & "))"
            strSQL = strSQL & " AND NOT EXISTS"
            strSQL = strSQL & " (SELECT * FROM dailymktschedule WHERE account_code='" & pstrCustomerCode & "' and CONSIGNEE_CODE='" & pstrConsCode & "' "
            strSQL = strSQL & " AND dailymktschedule.Cust_Drgno = Cust_ord_dtl.Cust_Drgno AND dailymktschedule.UNIT_CODE = Cust_ord_dtl.UNIT_CODE AND dailymktschedule.UNIT_CODE='" & gstrUNITID & "'"
            strSQL = strSQL & " AND dailymktschedule.Item_Code = Cust_ord_dtl.Item_Code and month(trans_date) IN (" & GetMonthCondition("M") & ") AND year(trans_date) IN (" & GetMonthCondition("Y") & ")) "
            strSQL = strSQL & " AND item_mst.item_code = cust_ord_dtl.item_code and item_mst.UNIT_CODE = cust_ord_dtl.UNIT_CODE and item_mst.item_main_grp in('F','T','S')"
            strSQL = strSQL & " AND CustItem_Mst.account_code='" & pstrCustomerCode & "' AND Cust_ord_dtl.cust_Drgno=CustItem_Mst.cust_Drgno AND Cust_ord_dtl.UNIT_CODE=CustItem_Mst.UNIT_CODE AND Cust_ord_dtl.UNIT_CODE='" & gstrUNITID & "' AND Cust_ord_dtl.Active_Flag='A' AND Cust_ord_dtl.Authorized_Flag=1 and Cust_ord_dtl.item_code =CustItem_Mst.item_code AND CustItem_Mst.account_code= Cust_ord_dtl.account_code And item_mst.status = 'A' Order by Cust_ord_dtl.Cust_Drgno"
            rsGetItem.Open(strSQL, mP_Connection, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            If rsGetItem.EOF Or rsGetItem.BOF Then
                MsgBox("No Record Found", MsgBoxStyle.Critical, "empower")
                GetItemList = False
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
            strSQL = "SELECT monthlymktschedule.cust_drgno,monthlymktschedule.item_code,year_month,schedule_qty,Despatch_qty,RevisionNo,Drg_Desc,Remarks"
            strSQL = strSQL & " FROM monthlymktschedule ,CustItem_Mst WHERE monthlymktschedule.UNIT_CODE=CustItem_Mst.UNIT_CODE AND monthlymktschedule.UNIT_CODE='" & gstrUNITID & "' AND Status=1 AND "
            strSQL = strSQL & " year_month IN (" & GetMonthCondition("MY") & ")"
            strSQL = strSQL & " AND monthlymktschedule.account_code='" & pstrCustomerCode & "' and monthlymktschedule.CONSIGNEE_CODE='" & pstrConsCode & "' "
            strSQL = strSQL & " AND CustItem_Mst.account_code='" & pstrCustomerCode & "' AND monthlymktschedule.cust_Drgno=CustItem_Mst.cust_Drgno AND monthlymktschedule.Item_code=CustItem_Mst.Item_Code"
            strSQL = strSQL & " GROUP BY monthlymktschedule.cust_drgno,monthlymktschedule.item_code,Drg_Desc,year_month,schedule_qty,Despatch_qty,RevisionNo,Remarks HAVING revisionno=MAX(revisionno) Order by monthlymktschedule.item_code"
            rsGetItem.Open(strSQL, mP_Connection, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            If rsGetItem.EOF Or rsGetItem.BOF Then
                MsgBox("No Record Found", MsgBoxStyle.Critical, "empower")
                GetItemList = False
                Exit Function
            Else
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
                            .Text = .Text & Trim(rsGetItem.Fields("Schedule_Qty").Value) & "�"
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
        '---------------------------------------------------------------------------------------
        'Name       :   GetMonthCondition
        'Type       :   Function
        'Author     :   Tapan Jain
        'Arguments  :
        'Return     :
        'Purpose    :
        '---------------------------------------------------------------------------------------
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
        '---------------------------------------------------------------------------------------
        'Name       :   FillSingleGrid
        'Type       :   Function
        'Author     :   Tapan Jain
        'Arguments  :
        'Return     :
        'Purpose    :
        '---------------------------------------------------------------------------------------
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
        '---------------------------------------------------------------------------------------
        'Name       :   SetTextBoxValue
        'Type       :   Function
        'Author     :   Tapan Jain
        'Arguments  :
        'Return     :
        'Purpose    :
        '---------------------------------------------------------------------------------------
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
        '---------------------------------------------------------------------------------------
        'Name       :   txtItemCode_Validate
        'Type       :   Sub
        'Author     :   Tapan Jain
        'Arguments  :
        'Return     :
        'Purpose    :
        '---------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        If Not ValidateEntry(Trim(txtCustDrgNo.Text), "C") Then
            MsgBox("Invalid Item Code", MsgBoxStyle.Information, "empower")
            txtItemCode.Focus()
            Cancel = True
        Else
            If CmdGrpMktSchedule.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                spdSingleEntryGrid.Row = 1
                spdSingleEntryGrid.Col = enumSingleEntryGrid.COLUMN_SCHEDULE_QTY
                spdSingleEntryGrid.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                spdSingleEntryGrid.Focus()
            Else
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
        '---------------------------------------------------------------------------------------
        'Name       :   MarkRowAsSaved
        'Type       :   Function
        'Author     :   Tapan Jain
        'Arguments  :
        'Return     :
        'Purpose    :
        '---------------------------------------------------------------------------------------
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
        '---------------------------------------------------------------------------------------
        'Name       :   SaveDataInWholeGrid
        'Type       :   Function
        'Author     :   Tapan Jain
        'Arguments  :
        'Return     :
        'Purpose    :
        '---------------------------------------------------------------------------------------
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
        '---------------------------------------------------------------------------------------
        'Name       :   DisablePrimaryField
        'Type       :   Function
        'Author     :   Tapan Jain
        'Arguments  :
        'Return     :
        'Purpose    :
        '---------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        DisablePrimaryField = True
        If CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            If DTPSheduleEndMonth.Value < CDate(mstrServerDate) Then
                MsgBox("Schedule End Month can not be smaller than Current Month", MsgBoxStyle.Information, "empower")
                DTPSheduleEndMonth.Focus()
                DisablePrimaryField = False
                Exit Function
            ElseIf DTPSheduleStartDate.Value < CDate(mstrServerDate) Then
                MsgBox("Schedule Start Month can not be smaller than Current Month", MsgBoxStyle.Information, "empower")
                DTPSheduleStartDate.Focus()
                DisablePrimaryField = False
                Exit Function
            ElseIf DTPSheduleStartDate.Value > DTPSheduleEndMonth.Value Then
                MsgBox("Schedule End Month can not be smaller than Schedule Start Month", MsgBoxStyle.Information, "empower")
                DTPSheduleEndMonth.Focus()
                DisablePrimaryField = False
                Exit Function
            End If
        Else
            If DTPSheduleStartDate.Value > DTPSheduleEndMonth.Value Then
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
            txtCustomerCode.Enabled = False
            Me.cmdCustomerHelp.Enabled = False
            Me.cmdConsHelp.Enabled = False
            txtConsCode.Enabled = False
        End If
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function SaveDataInTable() As Boolean
        '---------------------------------------------------------------------------------------
        'Name       :   SaveDataInTable
        'Type       :   Function
        'Author     :   Tapan Jain
        'Arguments  :
        'Return     :
        'Purpose    :
        '---------------------------------------------------------------------------------------
        Dim strSQL As String
        Dim strMonthYear As String
        Dim strInd_MonthYear() As String
        Dim lintLoop As Short
        On Error GoTo ErrHandler
        strSQL = ""
        SaveDataInTable = True
        strMonthYear = GetMonthCondition("MY")
        strInd_MonthYear = Split(strMonthYear, ",")
        If Trim(strMonthYear) = "" Then
            MsgBox("Could Not Get Month Year ", MsgBoxStyle.Critical, "empower")
            Exit Function
        End If
        If CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            With spdSingleEntryGrid
                For lintLoop = 1 To .MaxRows
                    .Row = lintLoop
                    .Col = enumSingleEntryGrid.COLUMN_SCHEDULE_QTY
                    If ISSavedEntry(Val(txtLineNo.Text)) Then
                        strSQL = strSQL & " UPDATE monthlymktschedule SET Schedule_Qty=" & Val(.Text) & ","
                        .Col = enumSingleEntryGrid.COLUMN_REMARKS
                        strSQL = strSQL & " Remarks='"
                        strSQL = strSQL & Trim(.Text) & "'"
                        strSQL = strSQL & " WHERE UNIT_CODE='" & gstrUNITID & "' AND Item_Code='" & Trim(txtItemCode.Text) & "' AND "
                        strSQL = strSQL & "Cust_Drgno='" & Trim(txtCustDrgNo.Text) & "' AND "
                        strSQL = strSQL & "RevisionNo=0 AND Status=1 AND Year_Month='"
                        strSQL = strSQL & strInd_MonthYear(lintLoop - 1) & "' AND Account_Code='" & Trim(txtCustomerCode.Text) & "' and CONSIGNEE_CODE= '" & Trim(Me.txtConsCode.Text) & "' "
                    Else
                        If Val(.Text) <> 0 Then
                            strSQL = strSQL & "INSERT INTO monthlymktschedule (UNIT_CODE,Year_Month,Account_Code,"
                            strSQL = strSQL & "Item_Code,Cust_Drgno,Schedule_Flag,"
                            strSQL = strSQL & "Schedule_Qty,Authorized_Code,Despatch_qty,Remarks,"
                            strSQL = strSQL & "Ent_dt,Ent_UserId,status,RevisionNo,Upd_dt,Upd_UserId,CONSIGNEE_CODE ) VALUES('"
                            strSQL = strSQL & gstrUNITID & "','"
                            strSQL = strSQL & strInd_MonthYear(lintLoop - 1) & "','" & Trim(txtCustomerCode.Text) & "','"
                            strSQL = strSQL & Trim(txtItemCode.Text) & "','" & Trim(txtCustDrgNo.Text) & "',1,"
                            strSQL = strSQL & Val(.Text) & ",'" & mP_User & "',0,'"
                            .Col = enumSingleEntryGrid.COLUMN_REMARKS
                            strSQL = strSQL & Trim(.Text) & "',getdate(),'" & mP_User & "',1,0,getdate(),'" & mP_User & "' ,'" & Trim(Me.txtConsCode.Text) & "') "
                            strSQL = strSQL & vbCrLf
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
                        strSQL = strSQL & " UPDATE monthlymktschedule SET Schedule_Qty=" & Val(.Text) & ","
                        .Col = enumSingleEntryGrid.COLUMN_REMARKS
                        strSQL = strSQL & " Remarks='"
                        strSQL = strSQL & Trim(.Text) & "'"
                        strSQL = strSQL & " WHERE UNIT_CODE='" & gstrUNITID & "' AND Item_Code='" & Trim(txtItemCode.Text) & "' AND "
                        strSQL = strSQL & "Cust_Drgno='" & Trim(txtCustDrgNo.Text) & "' AND "
                        strSQL = strSQL & "Status=1 AND Year_Month='"
                        strSQL = strSQL & strInd_MonthYear(lintLoop - 1) & "' AND Account_Code='" & Trim(txtCustomerCode.Text) & "' and CONSIGNEE_CODE= '" & Trim(Me.txtConsCode.Text) & "' "
                    Else
                        'If val(.Text) <> 0 Then
                        strSQL = strSQL & " UPDATE monthlymktschedule SET status=0 WHERE "
                        strSQL = strSQL & "UNIT_CODE='" & gstrUNITID & "' AND Item_Code='" & Trim(txtItemCode.Text) & "' AND "
                        strSQL = strSQL & "Cust_Drgno='" & Trim(txtCustDrgNo.Text) & "' AND "
                        strSQL = strSQL & "Year_Month='"
                        strSQL = strSQL & strInd_MonthYear(lintLoop - 1) & "' AND Account_Code='" & Trim(txtCustomerCode.Text) & "' and CONSIGNEE_CODE= '" & Trim(Me.txtConsCode.Text) & "' " & vbCrLf
                        strSQL = strSQL & "INSERT INTO monthlymktschedule (UNIT_CODE,Year_Month,Account_Code,"
                        strSQL = strSQL & "Item_Code,Cust_Drgno,Schedule_Flag,"
                        strSQL = strSQL & "Schedule_Qty,Authorized_Code,Despatch_qty,Remarks,"
                        strSQL = strSQL & "Ent_dt,Ent_UserId,status,RevisionNo,Upd_dt,Upd_UserId,CONSIGNEE_CODE ) VALUES('"
                        strSQL = strSQL & gstrUNITID & "','"
                        strSQL = strSQL & strInd_MonthYear(lintLoop - 1) & "','" & Trim(txtCustomerCode.Text) & "','"
                        strSQL = strSQL & Trim(txtItemCode.Text) & "','" & Trim(txtCustDrgNo.Text) & "',1,"
                        strSQL = strSQL & Val(.Text) & ",'" & mP_User & "',"
                        .Col = enumSingleEntryGrid.COLUMN_DESPATCH_QTY
                        strSQL = strSQL & Val(.Text) & ",'"
                        .Col = enumSingleEntryGrid.COLUMN_REMARKS
                        strSQL = strSQL & Trim(.Text) & "',getdate(),'" & mP_User & "',1,"
                        .Col = enumSingleEntryGrid.COLUMN_REVISIONCOUNT
                        strSQL = strSQL & Val(.Text) & ",getdate(),'" & mP_User & "','" & Trim(Me.txtConsCode.Text) & "' ) "
                        strSQL = strSQL & vbCrLf
                        strSQL = strSQL & " UPDATE monthlymktschedule SET RevisionNo=RevisionNo+1 "
                        strSQL = strSQL & " WHERE UNIT_CODE='" & gstrUNITID & "' and Item_Code='" & Trim(txtItemCode.Text) & "' AND "
                        strSQL = strSQL & "Cust_Drgno='" & Trim(txtCustDrgNo.Text) & "' AND "
                        strSQL = strSQL & "Year_Month='"
                        strSQL = strSQL & strInd_MonthYear(lintLoop - 1) & "' AND Account_Code='" & Trim(txtCustomerCode.Text) & "' and CONSIGNEE_CODE= '" & Trim(Me.txtConsCode.Text) & "' AND Status=1" & vbCrLf
                        'End If
                    End If
                Next
            End With
        End If
        If Trim(strSQL) <> "" Then
            mP_Connection.BeginTrans()
            mP_Connection.Execute(strSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            mP_Connection.CommitTrans()
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
        '---------------------------------------------------------------------------------------
        'Name       :   ISSavedEntry
        'Type       :   Function
        'Author     :   Tapan Jain
        'Arguments  :
        'Return     :
        'Purpose    :
        '---------------------------------------------------------------------------------------
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
        '---------------------------------------------------------------------------------------
        'Name       :   ValidateEntry
        'Type       :   Function
        'Author     :   Tapan Jain
        'Arguments  :
        'Return     :
        'Purpose    :
        '---------------------------------------------------------------------------------------
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
        '---------------------------------------------------------------------------------------
        'Name       :   txtLineNo_Validate
        'Type       :   Sub
        'Author     :   Tapan Jain
        'Arguments  :
        'Return     :
        'Purpose    :
        '---------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        If Not ValidateEntry(Trim(txtLineNo.Text), "L") Then
            MsgBox("Invalid Line No.", MsgBoxStyle.Information, "empower")
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
        If KeyAscii = 13 Then
            Call txtLineNo_Validating(txtLineNo, New System.ComponentModel.CancelEventArgs(False))
        End If
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub spdWholeEntryGrid_DblClick(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_DblClickEvent) Handles spdWholeEntryGrid.DblClick
        '---------------------------------------------------------------------------------------
        'Name       :   spdWholeEntryGrid_DblClick
        'Type       :   Sub
        'Author     :   Tapan Jain
        'Arguments  :
        'Return     :
        'Purpose    :
        '---------------------------------------------------------------------------------------
        'If CmdGrpMktSchedule.mode <> MODE_VIEW Then
        On Error GoTo ErrHandler
        mintCurrentNo = e.row
        Call SetTextBoxValue(mintCurrentNo)
        If CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            Call FillSingleGrid("ADD")
        ElseIf CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
            Call FillSingleGrid("EDIT")
        Else
            Call FillSingleGrid("VIEW")
        End If
        'End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub spdSingleEntryGrid_KeyPressEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyPressEvent) Handles spdSingleEntryGrid.KeyPressEvent
        If e.keyAscii = 13 And spdSingleEntryGrid.ActiveRow = spdSingleEntryGrid.MaxRows And (spdSingleEntryGrid.ActiveCol = enumSingleEntryGrid.COLUMN_SCHEDULE_QTY Or spdSingleEntryGrid.ActiveCol = enumSingleEntryGrid.COLUMN_REMARKS) Then
            Call spdSingleEntryGrid_Validating(spdSingleEntryGrid, New System.ComponentModel.CancelEventArgs(False))
        End If
    End Sub
    Private Sub spdSingleEntryGrid_LeaveCell(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_LeaveCellEvent) Handles spdSingleEntryGrid.LeaveCell
        '---------------------------------------------------------------------------------------
        'Name       :   spdSingleEntryGrid_LeaveCell
        'Type       :   Sub
        'Author     :   Tapan Jain
        'Arguments  :
        'Return     :
        'Purpose    :
        '---------------------------------------------------------------------------------------
        Dim dblPrevQty As Double
        Dim dblscheduleqty As Double
        Dim dblDespatchQty As Double
        On Error GoTo ErrHandler
        With spdSingleEntryGrid
            If Not .Lock Then
                If e.newRow = e.row + 1 Then
                    .Row = e.row
                    .Col = enumSingleEntryGrid.COLUMN_SCHEDULE_QTY
                    dblPrevQty = Val(.Text)
                    .Col = enumSingleEntryGrid.COLUMN_SCHEDULE_QTY
                    .Row = e.newRow
                    If Val(.Text) = 0 Then
                        .Text = CStr(dblPrevQty)
                    End If
                End If
            End If
            .Row = e.row
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
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub spdSingleEntryGrid_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles spdSingleEntryGrid.Validating
        Dim Cancel As Boolean = e.Cancel
        '---------------------------------------------------------------------------------------
        'Name       :   spdSingleEntryGrid_Validate
        'Type       :   Sub
        'Author     :   Tapan Jain
        'Arguments  :
        'Return     :
        'Purpose    :
        '---------------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        If CmdGrpMktSchedule.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            cmdSingleEntrySave.Focus()
        End If
        GoTo EventExitSub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        e.Cancel = Cancel
    End Sub
End Class