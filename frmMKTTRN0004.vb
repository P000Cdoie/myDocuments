Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmMKTTRN0004
	Inherits System.Windows.Forms.Form
	'----------------------------------------------------
	'(C) 2001 MIND, All rights reserved
	'Name of module - FRMMKTTRN0004.frm
	'Created by     - Pankaj Dwivedi
	'Created Date   - 18-04-2001
	'Description    - Daily/Monthly Market Schedule
	'Revised date   - 24 Aug 2001
	'Revision History  :   Nisha Rai
	'21/09/2001 MARKED CHECKED BY BCs changed on version 5
	'09/10/2001 changed on version 6 to add status active at insertion &
	'when in edit mode update status of existing schedule to zero &
	'insert a new row with status 1
	'01/11/2001 changed on version 9 in case of monthly schedule inserted
	'a new feild schedule_flag.
	'16/01/2002 set Enable property of Account Code Text Box,Account_Code
	'Help and Select Item Button to True in Form Key Press Event in Case
	'of vbKeyEscape is pressed.
	'12/02/2002 problem reported from pune that not abil to edit record in case of monthly schedule
	'form no 4045
	'05/06/2002 Chages in List for Cust Drg No
	'17/06/2002 to add chk all Items
	'26/06/02
	'Changed By Nisha on 14/08/2002 to Solve error on NEw Button Press
	'1.Error in Converting Datatype to Date time
	'2. Removed format when we assignin value to mStrCurrentDate = ServerDate()
	'Changed by nisha on 26/08/2002
	'Modifications Done By Rajesh Sharma on 12/09/2002
	'Changed by nisha on 14th feb
	'changed by nisha on 14feb version to add same drg no with distinct Item Code on 17/02/2003
	''condition added by nisha rai 30/05/2003 for red mark acc to calender Mst
	''changes done by nisha on 10/10/2003 To allow 0 quantity while editing
	' Changes done by Sourabh on 02 Sep 2004 For Add two new column DSNO and DSDATETIME(DSTracking-10623)
	' Changes done by Sourabh on 30 June 2005 against PIMS no PRJ-2004-04-003-15106 (It replicates one entry into many in one date)
	'-----------------------------------------------------------------------
	' Revision Date     :28/03/2006
	' Revision By       :Davinder Singh
	' Issue ID          :17378
	' Revision History  :To Save data also in the Forecast_mst
	'                    Validate the Schduled Qty. in the edit mode so that it can't be less than schedule qty. loaded by Schedule uploading
	'-----------------------------------------------------------------------
	' Revision Date     :17/07/07
	' Revision By       :Manoj Kr. Vaish
	' Issue ID          :20665
	' Revision History  :While Selecting Account Code in daily/Monthly Schedule giving message
	'                   [Invalid Customer Code OR Manual Schedule entry not Allowed !] ,if the account code exist in Sales order.
	'----------------------------------------------------------------------------------------------------------------------------------------
    'Revised By        -    Manoj Vaish
    'Revision Date     -    02 Mar 2009
    'Issue ID          -    eMpro-20090227-27987
    'Revision History  -    Consignee Changes for commercial invoice at Mate Units
    '-----------------------------------------------------------------------------
    'Revised By        -    Manoj Vaish
    'Revision Date     -    28 Apr 2009
    'Issue ID          -    eMpro-20090428-30750
    'Revision History  -    When dispatch is zero then user is allowed to change the schedule quantity.
    '-----------------------------------------------------------------------------
    'REVISED BY         : MANOJ VAISH
    'REVISED ON         : 04 JUN 2009
    'ISSUE ID           : eMpro-20090604-32080
    '                   : File type should not be updated with 'Manual'
    '                   : and Schedule Flag should be zero for old schedule while editing
    '-----------------------------------------------------------------------------
    'REVISED BY         : MANOJ VAISH
    'REVISED ON         : 26 JUN 2009
    'ISSUE ID           : eMpro-20090625-32895
    '                   : All type of invoice for which sales schedule is required should allow entry in daily marketing schedule  
    '-----------------------------------------------------------------------------
    'REVISED BY         : MANOJ VAISH
    'REVISED ON         : 10 Jul 2009
    'ISSUE ID           : eMpro-20090710-33491
    '                   : User is not able to enter daily marketing schedule.
    '                   : When the status of the item code is zero in monthly marketing schedule  for the month 
    '-----------------------------------------------------------------------------
    'Form Level Declarations
    'API Tpye
    ' Revised By                 -   Amit Rana
    ' Revision Date              -   04 MAY 2011
    ' Description                -   FOR MULTIUNIT FUNCTIONALITY
    '-----------------------------------------------------------------------------
    'REVISED By         : Prashant Rajpal 
    'REVISED ON         : 28-JUNE-2011 
    'iSSUE id           : 10108966
    'In Add MODE ,Depatch qty picked from SO , it is rectified now 
    '----------------------------------------------------------------------
    ' Change By Deepak on 11-Oct-2011 for Change Management---------------
    '-----------------------------------------------------------------------------
    'REVISED By         : Prashant Rajpal 
    'REVISED ON         : 03-FEB-2012 
    'iSSUE id           : 10190371
    'In EDIT MODE ,Schedule Qty Can't Be Changed for Back date.
    '----------------------------------------------------------------------
    ' Change By Virendra Gupta on 14-Feb-2012 for Change Management---------------
    '-----------------------------------------------------------------------------
    'Revised By       - Neha Ghai
    'Revision Date    - 19 sep 2012
    'Issue Id         - 10277185 
    'Description      - search Option not working.
    '---------------------------------------------------------------------------
    'Revised By       - Prashant Rajpal
    'Revision Date    - 06 jan 2014
    'Issue Id         - 10510618 
    'Description      - schedule compliance shows negative value :done
    '---------------------------------------------------------------------------

    Private Structure POINTAPI
        Dim x As Integer
        Dim y As Integer
    End Structure
	Dim mP_formtag As Integer ' For storing the form Tag
	Dim mintFormIndex As Short
	Dim mStrAccountCode As String 'to assign accountCode
	Dim mStrDate As Date 'to assign date in MM/yyyy Format
	Dim mStrItemCodes As String 'to assing Items code
	Dim mblnDailySchedule As Boolean 'to assign schedule
	Dim mStrCurrentDate As Date ' to get current date from server
	Dim mStrFinancialYearStartDate As Date 'to get financial year date from Company_Master
	Dim mStrFinancialYearEndDate As Date 'to get financial year date from Company_Master
	Dim scheduleEdit As Boolean
	Dim mdblDespatchQty() As Double 'to update despatch value while editing
	Private Const LB_FINDSTRING As Short = &H18Fs 'Constant for API
	'this API is used to do the searching in listbox
	Private Declare Function SendMessageByString Lib "user32"  Alias "SendMessageA"(ByVal Hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As String) As Integer
	'This API is used to put horizental scroll bar in listbox
	Private Declare Function SendMessageBynum Lib "user32"  Alias "SendMessageA"(ByVal Hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
	Private Const LB_GETHORIZONTALEXTENT As Short = &H193s
	Private Const LB_SETHORIZONTALEXTENT As Short = &H194s
	'this API is used to do the searching in listbox
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Short, ByVal lParam As Object) As Integer
    'This API Is used to get curcor position
    Private Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As POINTAPI) As Integer
    Dim mRdoCls As New ClsResultSetDB
    Dim blnInvalidData As Boolean
    Dim blnCheckallClicked As Boolean
    Dim mRow As Short ''Declared by Davinder to save row No of the grid
    Dim mCol As Short ''Declared by Davinder to save col No of the grid
    Dim blnDateChange As Boolean = False
    Dim blnChkChanged As Boolean = False

    Private Sub chkCheckAll_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkCheckAll.CheckStateChanged
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Check all the items in Select Items List Window
        'Comments           :   NA
        'Creation Date      :   18/04/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim intLoopCounter As Short
        Dim intMaxCounter As Short
        If blnChkChanged = True Then
            Exit Sub
        End If
        If chkUnCheckAll.Checked = True Then
            blnChkChanged = True
            chkUnCheckAll.Checked = False
            blnChkChanged = False
        End If
        If chkCheckAll.CheckState = 1 Then
            If LstItems.Items.Count > 0 Then
                intMaxCounter = LstItems.Items.Count - 1
                For intLoopCounter = 0 To intMaxCounter
                    LstItems.SetItemChecked(intLoopCounter, True)
                Next
            End If
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub chkUnCheckall_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkUnCheckAll.CheckStateChanged
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Uncheck all the items in Select Items List Window
        'Comments           :   NA
        'Creation Date      :   18/04/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim intLoopCounter As Short
        Dim intMaxCounter As Short
        If blnChkChanged = True Then
            Exit Sub
        End If
        If chkCheckAll.Checked = True Then
            blnChkChanged = True
            chkCheckAll.Checked = False
            blnChkChanged = False
        End If
        If chkUnCheckAll.CheckState = 1 Then
            If LstItems.Items.Count > 0 Then
                intMaxCounter = LstItems.Items.Count - 1
                For intLoopCounter = 0 To intMaxCounter
                    LstItems.SetItemChecked(intLoopCounter, False)
                Next
            End If
            '****
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub CmdGrpMktSchedule_ButtonClick1(ByVal Sender As Object, ByVal e As UCActXCtl.UCbtndgrp.ButtonClickEventArgs) Handles CmdGrpMktSchedule.ButtonClick
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To add Functionality of ADD/EDIT/SAVE/UPDATE
        'Comments           :   NA
        'Creation Date      :   18/04/2001
        '*******************************************************************************
        ' to take action according to button pressed
        On Error GoTo ErrHandler
        Dim intRow As Short
        Dim varStatus As String
        'Add new Column (Hidden)Open SO For Accounts Plug in
        Dim strMode As String
        Dim enmValue As eMPowerFunctions.ConfirmWindowReturnEnum
        Select Case e.Button
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD 'Add Record
                mStrAccountCode = TxtAccountCode.Text
                mStrDate = CDate(VB6.Format(DTPTransDate.Value, "MM/yyyy"))
                If OptDaily.Checked = True Then ' Check for daily or monthly schedule
                    mblnDailySchedule = True
                Else
                    mblnDailySchedule = False ' Check for daily or monthly schedule
                End If
                Call RefreshFrm() 'To refresh the form
                Call EnableControls(True, Me) 'Enable Controls
                Me.TxtAccountCode.Focus() 'to set the focus

            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE 'Save Record
                Select Case CmdGrpMktSchedule.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        strMode = "MODE_ADD"
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        strMode = "MODE_EDIT"
                End Select
                If ValidRecord() = True Then
                    If CheckScheduleQuantity(strMode) = True Then
                        If OptDaily.Checked = True Then 'if daily schedule then
                            If DecimalAllowed(mRow, mCol) = True Then
                                Me.fpSDailySchedule.Row = mRow
                                Me.fpSDailySchedule.Col = mCol
                                Me.fpSDailySchedule.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                Me.fpSDailySchedule.Focus()
                                Exit Sub
                            End If
                            If CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then 'to add the record
                                mP_Connection.BeginTrans()
                                Call SaveAddDaily("A") 'Call procedure to add record
                                mP_Connection.CommitTrans()
                                Call ConfirmWindow(10049, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                                CmdGrpMktSchedule.Enabled(1) = True
                                CmdGrpMktSchedule.Enabled(5) = True
                                gblnCancelUnload = False : gblnFormAddEdit = False
                            ElseIf CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then  'If Edit mode
                                If validateQty(mRow, mCol) = True Then
                                    Me.fpSDailySchedule.Row = mRow
                                    Me.fpSDailySchedule.Col = mCol
                                    Me.fpSDailySchedule.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                    Me.fpSDailySchedule.Focus()
                                    Exit Sub
                                End If
                                mP_Connection.BeginTrans()
                                Call SaveUpdateDaily() 'Call Procedure to edit record
                                mP_Connection.CommitTrans()
                                Call ConfirmWindow(10049, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                                With fpSDailySchedule
                                    .Row = 1 : .Row2 = .MaxRows : .Col = 2 : .Col2 = .MaxCols
                                    .BlockMode = True : .Lock = True : .BlockMode = False
                                End With
                                CmdGrpMktSchedule.Enabled(1) = True
                                CmdGrpMktSchedule.Enabled(5) = False
                                gblnCancelUnload = False : gblnFormAddEdit = False
                            End If
                        Else 'if monthly mkt schedule is selected
                        End If
                    Else
                        MsgBox("Schedule Quantity Must Be Greater than or Equal To Dispatch Quantity.", MsgBoxStyle.OkOnly, ResolveResString(100))
                        'Add new Column (Hidden)Open SO For Accounts Plug in
                        If fpSDailySchedule.Enabled = True Then
                            For intRow = 1 To fpSDailySchedule.MaxRows
                                varStatus = Nothing
                                Call fpSDailySchedule.GetText(7, intRow, varStatus)
                                If varStatus.ToString = "1" Then
                                    fpSDailySchedule.Col = 6
                                    fpSDailySchedule.Row = intRow
                                    fpSDailySchedule.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                    Exit Sub
                                End If
                            Next
                        Else
                        End If
                        Exit Sub
                    End If
                Else
                    gblnCancelUnload = True : gblnFormAddEdit = True
                    Exit Sub
                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL ' Cancel Click
                Call frmMKTTRN0004_KeyPress(Me, New System.Windows.Forms.KeyPressEventArgs(Chr(System.Windows.Forms.Keys.Escape)))
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT 'Clich on edit
                If ValidRecord() = True Then
                    mStrAccountCode = TxtAccountCode.Text
                    mStrDate = CDate(VB6.Format(DTPTransDate.Value, "MM/yyyy")) 'to retain existing value used in case of cancel
                    If OptDaily.Checked = True Then
                        mblnDailySchedule = True 'to retain existing value used in case of cancel
                    Else
                        mblnDailySchedule = False
                    End If
                    If OptDaily.Checked = True Then 'daily mkt schedule is selected
                        Call DisablePrimaryKeyControl() 'to disable control
                        Call fillSpread() ' Fill spread according to values
                    End If
                    'issue id 10190371
                    'With fpSDailySchedule
                    '    .Row = 1
                    '    .Col = 1
                    '    .Row2 = .MaxRows
                    '    .Col2 = .MaxCols
                    '    .BlockMode = True
                    '    .Lock = True
                    '    .BlockMode = False
                    'End With
                    'With fpSDailySchedule
                    '    .Row = 1
                    '    .Col = 6
                    '    .Row2 = .MaxRows
                    '    .Col2 = 7
                    '    .BlockMode = True
                    '    .Lock = False
                    '    .BlockMode = False
                    'End With
                    ' issue id Done : 10190371
                    Frame2.Enabled = True
                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE 'To Verify Close Operation
                If Me.CmdGrpMktSchedule.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                    enmValue = ConfirmWindow(10055, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNOCANCEL, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION)
                    If enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                        Call CmdGrpMktSchedule_ButtonClick1(CmdGrpMktSchedule, New UCActXCtl.UCbtndgrp.ButtonClickEventArgs(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE))
                        gblnCancelUnload = False : gblnFormAddEdit = False
                        Me.Close()
                    ElseIf enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_NO Then
                        gblnCancelUnload = False : gblnFormAddEdit = False
                        Me.Close()
                    End If
                Else
                    Me.Close()
                End If
        End Select
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub CmdHelpLocationCode_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdHelpLocationCode.Click
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Show Help on Customer Code
        'Comments           :   NA
        'Creation Date      :   18/04/2001
        '*******************************************************************************
        Dim strAccountCode() As String
        Dim strHelp() As String
        'to provide help of valid account codes and their description
        On Error GoTo ErrHandler
        If Len(TxtAccountCode.Text) = 0 Then
            lbldesc.Text = ""
            If CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                strAccountCode = ctlEMPHelpMktSchedule.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "select distinct a.Account_code, b.Cust_Name from  DailyMktSchedule a, Customer_mst b where A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" + gstrUNITID + "' AND a.account_code=b.Customer_Code and ((isnull(b.deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= b.deactive_date))", "Customer Code Help", 2)
            Else
                strAccountCode = ctlEMPHelpMktSchedule.ShowList(gstrCONNECTIONSERVER, gstrDSNName, gstrDatabaseName, "SELECT Distinct a.Account_code, b.Cust_Name FROM Cust_Ord_Hdr a,customer_mst b WHERE A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" + gstrUNITID + "' AND a.account_code=b.Customer_Code and b.AllowExcessSchedule=1 and ((isnull(b.deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= b.deactive_date))", "Customer Code Help", 2)
            End If
            ''Changes for Issue Id:20418 end here.
            Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.obj_Screen, , System.Windows.Forms.Cursors.Default)
            If UBound(strAccountCode) <> CDbl("-1") Then
                If strAccountCode(0) <> "0" Then
                    If strAccountCode(0) <> "" Then
                        TxtAccountCode.Text = strAccountCode(0)
                        lbldesc.Text = Trim(strAccountCode(1))
                        ReturnAccountDescription((TxtAccountCode.Text))
                    End If
                Else
                    TxtAccountCode.Text = ""
                    lbldesc.Text = ""
                    TxtAccountCode.Focus()
                    Call ConfirmWindow(10080, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_CRITICAL, 100)
                End If
            End If

            If TxtAccountCode.Enabled = True Then
                TxtAccountCode.Focus()
            End If
        End If
        Call FillList()
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub CmdHelpLocationCode_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles CmdHelpLocationCode.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Call Help on F1 Click
        'Comments           :   NA
        'Creation Date      :   18/04/2001
        '*******************************************************************************
        'to provide help of valid account codes and their description
        On Error GoTo ErrHandler
        Select Case KeyCode
            Case System.Windows.Forms.Keys.F1
                CmdHelpLocationCode_Click(CmdHelpLocationCode, New System.EventArgs())
        End Select
        Exit Sub
ErrHandler:  'This is to avoid the execution of the error handler
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub CmdSelectItems_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdSelectItems.Click
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Check all the items in Select Items List Window
        'Comments           :   NA
        'Creation Date      :   18/04/2001
        '*******************************************************************************
        'to provide list of items
        On Error GoTo ErrHandler
        If Len(Trim(TxtAccountCode.Text)) <= 0 Then
            MsgBox("Please select a customer first", MsgBoxStyle.OkOnly, ResolveResString(100))
            TxtAccountCode.Focus()
            Exit Sub
        End If
        Call FillList() 'Fill the list for valid items
        Frame2.Enabled = False
        fpSDailySchedule.Enabled = False
        TxtAccountCode.Enabled = False
        DTPTransDate.Enabled = False
        CmdSelectItems.Enabled = False
        CmdHelpLocationCode.Enabled = False
        FraLstItems.Visible = True

        If LstItems.Enabled = True Then
            If LstItems.Items.Count > 0 Then
                LstItems.SelectedIndex = 0
            End If
            LstItems.Focus()
        End If

        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdOK.Click
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Get Selected items from list & Will Displayed in
        '                       the Window
        'Comments           :   NA
        'Creation Date      :   18/04/2001
        '*******************************************************************************
        ' to disapear tabpage and fill grid according selected items
        On Error GoTo ErrHandler
        Dim strReturnItems As String
        Frame1.Enabled = True 'to make selection free control exist in frame1
        mStrItemCodes = ReturnitemCode()
        strReturnItems = Replace(ReturnitemCode, "'", "")
        If Len(Trim(strReturnItems)) > 0 Then
            Call fillSpread() 'Fill the list for vaild items
            FraLstItems.Visible = False 'make invisible the frame containing items list
            If CmdGrpMktSchedule.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                If OptDaily.Checked = True Then
                    fpSDailySchedule.Focus()
                End If
            Else
                If scheduleEdit = True Then
                    If AllowEditSchedule() = True Then
                        CmdGrpMktSchedule.Enabled(1) = True
                    Else
                        CmdGrpMktSchedule.Enabled(1) = False
                    End If
                Else
                    CmdGrpMktSchedule.Enabled(1) = False
                End If
            End If
        Else
            fpSDailySchedule.MaxRows = 0
            MsgBox("Please Select Atleast one item", MsgBoxStyle.OkOnly, ResolveResString(100))
            FraLstItems.Visible = False 'make invisible the frame containing items list
            CmdSelectItems.Focus()
        End If
        If CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            Frame2.Enabled = True
            fpSDailySchedule.Enabled = True
            TxtAccountCode.Enabled = True
            DTPTransDate.Enabled = True
            CmdSelectItems.Enabled = True
            CmdHelpLocationCode.Enabled = True
        Else
            If Trim(strReturnItems) = "" Then
                Frame2.Enabled = True
                fpSDailySchedule.Enabled = True
                TxtAccountCode.Enabled = True
                DTPTransDate.Enabled = True
                CmdSelectItems.Enabled = True
                CmdHelpLocationCode.Enabled = True
            Else
                Frame2.Enabled = True
                fpSDailySchedule.Enabled = True
                TxtAccountCode.Enabled = False
                DTPTransDate.Enabled = False
                CmdSelectItems.Enabled = False
                CmdHelpLocationCode.Enabled = False
            End If

        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub CmdSelectItems_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles CmdSelectItems.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Check all the items in Select Items List Window
        'Comments           :   NA
        'Creation Date      :   18/04/2001
        '*******************************************************************************
        'to popup list of items when user press F1
        On Error GoTo ErrHandler
        Select Case KeyCode
            Case System.Windows.Forms.Keys.F1
                CmdSelectItems_Click(CmdSelectItems, New System.EventArgs()) ' get the same list when user will press F1
        End Select
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub ctlFormHeader1_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlFormHeader1.Click
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Show empower help
        'Comments           :   NA
        'Creation Date      :   18/04/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Call ShowHelp("HLP" & Mid(Me.Name, 4, Len(Me.Name)) & ".htm") '("pddrep_listofitems.htm")
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub DTPTransDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DTPTransDate.ValueChanged
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To add code when month changes.
        'Comments           :   NA
        'Creation Date      :   18/04/2001
        'to check valid schedule date
        On Error GoTo ErrHandler
        Call FillList()
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub DTPTransDate_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DTPTransDate.KeyDown
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To add code when month changes.
        'Comments           :   NA
        'Creation Date      :   18/04/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        If e.KeyCode = 13 Then
            If CmdSelectItems.Enabled = True Then
                CmdSelectItems.Focus()
            End If
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0004_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To add code on form activation
        'Comments           :   NA
        'Creation Date      :   18/04/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        'Me.Caption = LoadResString(50024)
        DTPTransDate.Format = DateTimePickerFormat.Custom  'to set custom format to DTPInvoiceDate Control
        DTPTransDate.CustomFormat = "MM/yyyy" 'assign custom format to DTPInvoiceDate Control
        'To assign Current Month and year to dtptransDate control
        Call AssignCurrentMonthToDTPTRnasDate()
        mdifrmMain.CheckFormName = mintFormIndex
        frmModules.NodeFontBold(Tag) = True

        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0004_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To add code on Form Deactivate
        'Comments           :   NA
        'Creation Date      :   18/04/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Tag) = False
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0004_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To show empower help on F4 click
        'Comments           :   NA
        'Creation Date      :   18/04/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0 Then Call ctlFormHeader1_Click(ctlFormHeader1, New System.EventArgs()) : Exit Sub
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0004_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To add code functionality for Escape Click
        'Comments           :   NA
        'Creation Date      :   18/04/2001
        '*******************************************************************************
        Dim enmValue As eMPowerFunctions.ConfirmWindowReturnEnum
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Escape
                'If user press the ESC Key ,the Form will be unloaded
                If Me.CmdGrpMktSchedule.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                    enmValue = ConfirmWindow(10053, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION)
                    If enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                        Call RefreshFrm()
                        Me.CmdGrpMktSchedule.Revert()
                        CmdGrpMktSchedule.Enabled(1) = False
                        CmdGrpMktSchedule.Enabled(2) = False
                        Frame2.Enabled = True
                        Me.OptDaily.Enabled = True
                        Me.DTPTransDate.Enabled = True
                        Me.OptDaily.Focus()
                        gblnCancelUnload = False : gblnFormAddEdit = False
                        TxtAccountCode.Enabled = True
                        CmdHelpLocationCode.Enabled = True
                        CmdSelectItems.Enabled = True
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
    Private Sub frmMKTTRN0004_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To add the code of form Load
        'Comments           :   NA
        'Creation Date      :   18/04/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        mintFormIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader1.Tag)
        'to check schedule date is in financial year or not
        Call AssignValidFinancialYearDate()
        Call FitToClient(Me, Frame1, ctlFormHeader1, CmdGrpMktSchedule) 'To fit the form in the MDI
        fpSDailySchedule.MaxRows = 0 'To initilize spread
        fpSDailySchedule.Col = 1 : fpSDailySchedule.Col2 = 1 : fpSDailySchedule.ColHidden = True
        fpSDailySchedule.Col = 11 : fpSDailySchedule.Col2 = 11 : fpSDailySchedule.ColHidden = True
        fpSDailySchedule.Col = 12 : fpSDailySchedule.Col2 = 12 : fpSDailySchedule.ColHidden = True
        fpSDailySchedule.Col = 13 : fpSDailySchedule.Col2 = 13 : fpSDailySchedule.ColHidden = True
        fpSDailySchedule.Col = 2 : fpSDailySchedule.Col2 = 2 : fpSDailySchedule.ColHidden = True
        fpSDailySchedule.Col = 9 : fpSDailySchedule.Col2 = 9 : fpSDailySchedule.ColHidden = True
        'SetGridDateFormat(fpSDailySchedule, 3)
        'to put scrollbar in List Box
        Call SendMessageBynum(LstItems.Handle.ToInt32, LB_SETHORIZONTALEXTENT, 900, 0)
        FraLstItems.Left = VB6.TwipsToPixelsX(2500)
        'to fill the lables from resource file
        Call FillLabelFromResFile(Me)

        Me.OptDaily.Checked = True
        Me.CmdGrpMktSchedule.Enabled(1) = False
        Me.CmdGrpMktSchedule.Enabled(4) = False  'to make Print button false
        Me.CmdGrpMktSchedule.Enabled(5) = False 'to make cancle button false
        Me.CmdGrpMktSchedule.Enabled(2) = False ' to disable delete button
        mStrItemCodes = "''" 'To initialize this variable
        Exit Sub 'To avoid the execution of error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0004_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To add the code of form Query Unload
        'Comments           :   NA
        'Creation Date      :   18/04/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        If gblnCancelUnload = True Then eventArgs.Cancel = 1
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub frmMKTTRN0004_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To add the code of form Unload
        'Comments           :   NA
        'Creation Date      :   18/04/2001
        '*******************************************************************************
        On Error GoTo ErrHandler


        mdifrmMain.RemoveFormNameFromWindowList = mintFormIndex
        frmModules.NodeFontBold(Tag) = False
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub fpSDailySchedule_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles fpSDailySchedule.Enter
        Dim fpCounter As Integer = 0
        Dim varStatusCheck As Object
        Dim boolCheck As Boolean
        Dim count As Integer = 0
        With fpSDailySchedule
            .Row = fpCounter
            .Row2 = .MaxRows
            .Col = 1
            .Col2 = 5
            .BlockMode = True
            .Lock = True
            .BlockMode = False
            .Row = fpCounter
            .Row2 = .MaxRows
            .Col = 8
            .Col2 = .MaxCols
            .BlockMode = True
            .Lock = True
            .BlockMode = False
        End With
    End Sub
    Private Sub fpSDailySchedule_LeaveCell(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_LeaveCellEvent) Handles fpSDailySchedule.LeaveCell
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To add code for data Validation on changing cell in Grid
        'Comments           :   NA
        'Creation Date      :   18/04/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        'to check data in grid
        Dim strSql As String 'StrSql to write sql before execution
        Dim i As Short 'i is the counter for for loop
        Dim dblscheduleqty As Double ' to assign scheduled quantity value
        Dim dblOrderQty As Double ' to assign Ordered Quantity value
        Dim dblDispatchqty As Double ' to assign Dispatch Quantity Value
        Dim dblDummy As Double ' One Extra Variable is required with getfloat,gettext in spread
        Dim intTotalItem As Short ' to assign total no of items in spread
        Dim intNoofDaysinMonth As Short ' to assign total no of days in month
        Dim intItemNo As Short ' occurence of item in spread
        Dim intItemNo1 As Short ' occurence of item in spread
        Dim intItemRange1 As Short ' item exist from which row no to which row no (lower)
        Dim intItemRange2 As Short ' item exist from which row no to which row no (upper)
        Dim dblTotalScheduleItemQty As Double 'to assign Total Schedule Quantity including previous quantity
        Dim dblItemCode As String ' to assign item code from spread
        Dim strCustDrgNo As String ' to assign Cust Drgno from spread
        Dim varStatus As Object ' to assign status of item
        Dim dblOldScheduleQty As Object ' to assign old schedule quntity
        Dim dblPreviousValue As Object ' to get the previous existing value
        Dim dblTransDate As Object ' to assign transcation date
        Dim DblCurrentScheduleQty As Double ' to assign current existing value in spread
        Dim dblschDispatch As Double
        Dim DblTotSchDispatch As Object
        Dim dblcurrschqty As Double
        Dim dblcurrdisqty As Double
        Dim varOpenSO As Object
        Dim intItemStartRow As Short
        Dim intItemsRowsinSpread As Short
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        Dim varLoopItemCode As String
        Dim varLoopDrgNo As String
        Dim rsNagare As ClsResultSetDB
        rsNagare = New ClsResultSetDB
        Dim strQuery As String
        Dim blnValidate As Boolean
        Dim blnFlag As Boolean

        'Add new Column (Hidden)Open SO For Accounts Plug in
        If e.newCol = -1 Then
            If e.col = 6 Then
                mRow = e.row
                mCol = e.col
            End If
            Exit Sub
        End If
        If e.col = 6 Then
            If DecimalAllowed(e.row, e.col) Then
                e.cancel = True
                Exit Sub
            End If
            fpSDailySchedule.Row = e.row
            varStatus = Nothing
            dblDummy = fpSDailySchedule.GetText(7, e.row, varStatus)
            If Trim(varStatus) = "1" Then
                If CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                    fpSDailySchedule.Row = e.row
                    dblItemCode = Nothing
                    strCustDrgNo = Nothing
                    dblDummy = fpSDailySchedule.GetText(5, e.row, dblItemCode)
                    dblDummy = fpSDailySchedule.GetText(8, e.row, strCustDrgNo)
                    intNoofDaysinMonth = NoOfDaysinMonth(CShort(VB.Left(VB6.Format(DTPTransDate.Value, "MM/yyyy"), 2)), CShort(VB.Right(VB6.Format(DTPTransDate.Value, "MM/yyyy"), 4)))
                    intTotalItem = Fix(fpSDailySchedule.MaxRows / intNoofDaysinMonth)
                    intItemNo = Fix(e.row / intNoofDaysinMonth)
                    intItemNo1 = e.row Mod intNoofDaysinMonth
                    intItemRange1 = intItemNo
                    If intItemNo1 = 0 Then
                        intItemRange1 = intItemRange1 - 1
                    End If
                    intMaxLoop = fpSDailySchedule.MaxRows
                    intItemsRowsinSpread = 0 : intItemStartRow = 0
                    For intLoopCounter = 1 To intMaxLoop
                        varLoopItemCode = ""
                        varLoopDrgNo = ""
                        fpSDailySchedule.Row = intLoopCounter
                        fpSDailySchedule.Col = 5
                        varLoopItemCode = fpSDailySchedule.Value
                        fpSDailySchedule.Col = 8
                        varLoopDrgNo = fpSDailySchedule.Value

                        If (Trim(varLoopItemCode) = dblItemCode) And (Trim(varLoopDrgNo) = Trim(strCustDrgNo)) Then
                            If intItemStartRow = 0 Then
                                intItemStartRow = intLoopCounter
                            End If
                            intItemsRowsinSpread = intItemsRowsinSpread + 1
                        End If
                    Next
                    intItemsRowsinSpread = intItemsRowsinSpread + (intItemStartRow - 1)
                    dblTotalScheduleItemQty = 0
                    For intItemStartRow = intItemStartRow To intItemsRowsinSpread
                        dblDummy = fpSDailySchedule.GetFloat(6, intItemStartRow, dblscheduleqty)
                        dblTotalScheduleItemQty = dblTotalScheduleItemQty + dblscheduleqty
                    Next
                    varOpenSO = Nothing
                    Call fpSDailySchedule.GetText(1, e.row, varOpenSO)
                    dblDummy = fpSDailySchedule.GetFloat(9, e.row, dblOrderQty)
                    dblDummy = fpSDailySchedule.GetFloat(10, e.row, dblDispatchqty)
                    If dblOrderQty > 0 Then
                        If varOpenSO = 0 Then
                            '*********To Check if Total Schedule Quantity should not be greater then total Ordet Quantity - Total Despatch Quantity
                            If dblTotalScheduleItemQty > (dblOrderQty - dblDispatchqty) Then
                                fpSDailySchedule.Row = e.row
                                fpSDailySchedule.Col = 6
                                dblDummy = fpSDailySchedule.GetFloat(6, e.row, DblCurrentScheduleQty)
                                fpSDailySchedule.Refresh()
                                ResolveResString(10082)
                                MsgBox(ResolveResString(10082) & Str(dblOrderQty - dblDispatchqty), MsgBoxStyle.Critical, ResolveResString(100))
                                fpSDailySchedule.Focus()
                                fpSDailySchedule.Row = e.row
                                fpSDailySchedule.Col = 6
                                fpSDailySchedule.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                Exit Sub
                            End If
                        End If
                    End If
                ElseIf CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then  ' In edit mode
                    fpSDailySchedule.Row = e.row
                    fpSDailySchedule.Col = 5
                    strQuery = "select isnull(sum(Quantity),0) as Sch_Qty from mkt_enagaredtl where UNIT_CODE='" + gstrUNITID + "' AND Account_code='" & Trim(TxtAccountCode.Text) & "' AND Item_code='" & Trim(fpSDailySchedule.Text) & "' AND sch_date='"
                    fpSDailySchedule.Col = 3
                    strQuery = strQuery & VB6.Format(fpSDailySchedule.Text, "dd mmm yyyy") & "'"
                    fpSDailySchedule.Col = 8
                    strQuery = strQuery & " AND Cust_drgNo='" & Trim(fpSDailySchedule.Text) & "'"
                    rsNagare.GetResult(strQuery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                    fpSDailySchedule.Col = 6

                    If Val(fpSDailySchedule.Text) < Val(rsNagare.GetValue("Sch_Qty")) Then

                        MsgBox("Schedule Qty. can't be less than: " & Trim(rsNagare.GetValue("Sch_Qty")), MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                        e.cancel = True
                        Exit Sub
                    End If
                    rsNagare.ResultSetClose()
                    rsNagare = Nothing
                    dblItemCode = Nothing
                    strCustDrgNo = Nothing
                    dblDummy = fpSDailySchedule.GetText(5, e.row, dblItemCode)
                    dblDummy = fpSDailySchedule.GetText(8, e.row, strCustDrgNo)
                    dblDummy = fpSDailySchedule.GetFloat(10, e.row, dblschDispatch)
                    strSql = "select order_qty=sum(order_qty),despatch_qty=sum(despatch_qty)" & " from cust_ord_dtl where UNIT_CODE='" + gstrUNITID + "' AND account_code='" & TxtAccountCode.Text & "' and item_code='" & dblItemCode & "' and cust_drgno = '" & strCustDrgNo & "' and  active_flag='A' and authorized_flag=1"
                    If mRdoCls.GetResult(strSql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly) And mRdoCls.GetNoRows > 0 Then
                        mRdoCls.MoveFirst()
                        dblOrderQty = Val(mRdoCls.GetValue("order_qty"))
                        dblDispatchqty = Val(mRdoCls.GetValue("despatch_qty"))
                        '--------------------------------------------------------------------
                        intNoofDaysinMonth = NoOfDaysinMonth(CShort(VB.Left(VB6.Format(DTPTransDate.Value, "MM/yyyy"), 2)), CShort(VB.Right(VB6.Format(DTPTransDate.Value, "MM/yyyy"), 4)))
                        intTotalItem = Fix(fpSDailySchedule.MaxRows / intNoofDaysinMonth)
                        intItemNo = Fix(e.row / intNoofDaysinMonth)
                        intItemNo1 = e.row Mod intNoofDaysinMonth
                        intItemRange1 = intItemNo
                        If intItemNo1 = 0 Then
                            intItemRange1 = intItemRange1 - 1
                        End If
                        Val(CStr((dblTotalScheduleItemQty) > 0))
                        intMaxLoop = fpSDailySchedule.MaxRows
                        intItemsRowsinSpread = 0 : intItemStartRow = 0
                        For intLoopCounter = 1 To intMaxLoop
                            varLoopItemCode = ""
                            varLoopDrgNo = ""
                            fpSDailySchedule.Row = intLoopCounter
                            fpSDailySchedule.Col = 5
                            varLoopItemCode = fpSDailySchedule.Value

                            fpSDailySchedule.Col = 8
                            varLoopDrgNo = fpSDailySchedule.Value

                            If (Trim(varLoopItemCode) = dblItemCode) And (Trim(varLoopDrgNo) = Trim(strCustDrgNo)) Then
                                If intItemStartRow = 0 Then
                                    intItemStartRow = intLoopCounter
                                End If
                                intItemsRowsinSpread = intItemsRowsinSpread + 1
                            End If
                        Next
                        intItemsRowsinSpread = intItemsRowsinSpread + (intItemStartRow - 1)
                        For intItemStartRow = intItemStartRow To intItemsRowsinSpread
                            dblDummy = fpSDailySchedule.GetFloat(6, intItemStartRow, dblscheduleqty)
                            dblDummy = fpSDailySchedule.GetFloat(10, intItemStartRow, dblschDispatch)
                            dblTotalScheduleItemQty = dblTotalScheduleItemQty + dblscheduleqty
                            DblTotSchDispatch = DblTotSchDispatch + dblschDispatch
                        Next
                        varOpenSO = Nothing
                        Call fpSDailySchedule.GetText(1, e.row, varOpenSO)
                        '------------------------------------------------------------------
                        If dblOrderQty > 0 Then
                            '**** to Check if item is Open in So then this Check Should not be there
                            If varOpenSO = 0 Then
                                '*********To Check if Total Schedule Quantity should not be greater then total Ordet Quantity - Total Despatch Quantity
                                If dblTotalScheduleItemQty > (Val(CStr(dblOrderQty)) - Val(CStr(dblDispatchqty))) Then
                                    fpSDailySchedule.Row = e.row
                                    fpSDailySchedule.Col = 6
                                    fpSDailySchedule.Text = CStr(0)
                                    dblDummy = fpSDailySchedule.SetFloat(6, e.row, 0)

                                    fpSDailySchedule.Refresh()
                                    MsgBox(ResolveResString(10082) & Str(dblOrderQty - dblDispatchqty), eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO, ResolveResString(100))
                                    fpSDailySchedule.Focus()
                                    Exit Sub
                                End If
                            End If
                            '*****To Check if edited Schedule Quantity should not be less then current Despatch Quantity
                            '*****For Current Date & Also For All Dates
                        End If
                        dblDummy = fpSDailySchedule.GetFloat(6, e.row, dblcurrschqty)
                        dblDummy = fpSDailySchedule.GetFloat(10, e.row, dblcurrdisqty)

                        If (dblTotalScheduleItemQty < DblTotSchDispatch) Then
                            fpSDailySchedule.Row = e.row
                            fpSDailySchedule.Col = 6
                            fpSDailySchedule.Refresh()
                            MsgBox("Total Schedule Quantity Must Be Greater then Total Dispatch Quantity : " & Str(DblTotSchDispatch), eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO, ResolveResString(100))
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If

        Exit Sub
ErrHandler:
        'Resume Next
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub FraLstItems_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles FraLstItems.MouseMove
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To add the code for FraLstItems_MouseMove
        'Comments           :   NA
        'Creation Date      :   18/04/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim upperleft As POINTAPI
        Dim lngMouseMove As Integer
        If e.Button = 1 Then
            lngMouseMove = GetCursorPos(upperleft)
            FraLstItems.Left = VB6.TwipsToPixelsX(upperleft.x * 7)
            FraLstItems.Top = VB6.TwipsToPixelsY(upperleft.y * 7)
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub LstItems_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LstItems.Click
        If LstItems.Items.Count > 0 Then
            LstItems.SetSelected(LstItems.SelectedIndex, True)
        End If

    End Sub
    Private Sub LstItems_ItemCheck(ByVal eventSender As System.Object, ByVal e As System.Windows.Forms.ItemCheckEventArgs) Handles LstItems.ItemCheck
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To check if all item selectet then check all items check
        '                       box
        'Comments           :   NA
        'Creation Date      :   18/04/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim intItemChecked As Short
        If LstItems.GetSelected(LstItems.SelectedIndex) Then
            intItemChecked = ToCheckNoOfItemSelected(e.Index, e.NewValue)
            If intItemChecked = LstItems.Items.Count Then
                chkCheckAll.CheckState = System.Windows.Forms.CheckState.Checked : chkUnCheckAll.CheckState = System.Windows.Forms.CheckState.Unchecked
            ElseIf intItemChecked = 0 Then
                chkUnCheckAll.CheckState = System.Windows.Forms.CheckState.Checked : chkCheckAll.CheckState = System.Windows.Forms.CheckState.Unchecked
            ElseIf (intItemChecked > 0) And (intItemChecked <> LstItems.Items.Count) Then
                chkCheckAll.CheckState = System.Windows.Forms.CheckState.Unchecked : chkUnCheckAll.CheckState = System.Windows.Forms.CheckState.Unchecked
            End If
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub LstItems_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles LstItems.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   to select and deselect items in listbox
        'Comments           :   NA
        'Creation Date      :   18/04/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        If KeyAscii = 13 Then
            If LstItems.SelectedIndex >= 0 Then
                If LstItems.GetItemChecked(LstItems.SelectedIndex) = True Then
                    LstItems.SetItemChecked(LstItems.SelectedIndex, False)
                Else
                    LstItems.SetItemChecked(LstItems.SelectedIndex, True)
                End If
            End If
        End If
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
    Private Sub OptDaily_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles OptDaily.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   when we select Daily Schedule option Daily Grid wil be
        '                       enabled
        'Comments           :   NA
        'Creation Date      :   18/04/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpMktSchedule.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        Me.TxtAccountCode.Focus()
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                        Frame3.Enabled = True
                        TxtAccountCode.Enabled = True
                        Me.DTPTransDate.Enabled = True
                        Me.CmdSelectItems.Enabled = True
                        Me.CmdHelpLocationCode.Enabled = True
                        Me.TxtAccountCode.Focus()
                End Select
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
    Private Sub TxtAccountCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TxtAccountCode.TextChanged
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   to Refresh all data on UI when Account_code Changes
        'Comments           :   NA
        'Creation Date      :   18/04/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        If Len(Trim(TxtAccountCode.Text)) = 0 Then
            Me.lbldesc.Text = " "
        End If
        If fpSDailySchedule.Visible = True Then
            fpSDailySchedule.MaxRows = 0
        End If

        CmdGrpMktSchedule.Enabled(1) = False
        If TxtAccountCode.Enabled = True Then
            TxtAccountCode.Focus()
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub TxtAccountCode_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles TxtAccountCode.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   to provide help when user will press F1
        'Comments           :   NA
        'Creation Date      :   18/04/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case KeyCode
            Case System.Windows.Forms.Keys.F1
                Call CmdHelpLocationCode_Click(CmdHelpLocationCode, New System.EventArgs())
        End Select
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub txtAccountCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles TxtAccountCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '----------------------------------------------------
        'Author              - Nisha Rai
        'Create Date         - 12/12/2003
        'Arguments           - None
        'Return Value        - None
        'Function            - To Call Validate Customer Code in Case of Single row Saving is On
        '----------------------------------------------------
        If KeyAscii = 13 Then Call TxtAccountCode_Validating(TxtAccountCode, New System.ComponentModel.CancelEventArgs(False))
        eventArgs.KeyChar = Chr(KeyAscii)
        If KeyAscii = 0 Then
            eventArgs.Handled = True
        End If
        If KeyAscii = 39 Then
            eventArgs.Handled = True
        End If
    End Sub
    Private Sub TxtAccountCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles TxtAccountCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        '*******************************************************************************
        'Author             :   Ashutosh Verma
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Validate Customer Code.
        'Comments           :   NA
        'Creation Date      :   02 Jul 2007,Issue Id:20418
        'Revision By        :   Manoj Kr. Vaish
        'Revised Date       :   17 July 2007 Issue ID:20665
        '*******************************************************************************
        Dim varGetCustName() As Object
        On Error GoTo ErrHandler
        If Trim(Me.TxtAccountCode.Text) = "" Then
            CmdGrpMktSchedule.Focus()
            GoTo EventExitSub
        End If
        If CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            varGetCustName = GetFieldsValues("select Distinct a.Account_code, b.Cust_Name from DailyMktSchedule a, customer_mst b WHERE A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" + gstrUNITID + "' And  a.account_code=b.Customer_Code  AND Account_code='" & Trim(TxtAccountCode.Text) & "' and ((isnull(b.deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= b.deactive_date))", 2)
        Else
            varGetCustName = GetFieldsValues("SELECT Distinct a.Account_code, b.Cust_Name FROM Cust_Ord_Hdr a,customer_mst b WHERE A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE='" + gstrUNITID + "' AND a.account_code=b.Customer_Code and b.AllowExcessSchedule=1 And A.Account_code='" & Trim(TxtAccountCode.Text) & "' and ((isnull(b.deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= b.deactive_date))", 2)
        End If
        If UCase(Trim(varGetCustName(0))) = UCase(Trim(TxtAccountCode.Text)) Then

            Me.lbldesc.Text = varGetCustName(1)
        Else
            MsgBox("Invalid Customer Code OR Manual Schedule entry not Allowed !", MsgBoxStyle.Information, "empower")
            TxtAccountCode.Text = ""
            Cancel = True
            GoTo EventExitSub
        End If
        GoTo EventExitSub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Function GetFieldsValues(ByVal pstrQuery As String, ByVal pIntNoOfColumn As Short) As Object
        '----------------------------------------------------
        'Author              -
        'Create Date         -
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
        rsRecordset = Nothing

        GetFieldsValues = VB6.CopyArray(varReturnVal)
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        rsRecordset = Nothing
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub TxtSearch_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TxtSearch.TextChanged
        On Error GoTo ErrHandler
        If LstItems.Items.Count > 0 Then
            Dim intcounter As Short
            With LstItems
                If Len(TxtSearch.Text) = 0 Then Exit Sub
                For intcounter = 0 To .Items.Count - 1
                    If Trim(UCase(ObsoleteManagement.GetItemString(LstItems, intcounter))) Like "*" & Trim(UCase(TxtSearch.Text)) & "*" Then
                        .SelectedIndex = intcounter
                        Exit For
                    End If
                Next
            End With
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub txtSearch_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles TxtSearch.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   to select and de select items in list box
        'Comments           :   NA
        'Creation Date      :   18/04/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        If KeyAscii = 13 Then
            If LstItems.SelectedIndex >= 0 Then
                If LstItems.GetItemChecked(LstItems.SelectedIndex) = True Then
                    LstItems.SetItemChecked(LstItems.SelectedIndex, False)
                Else
                    LstItems.SetItemChecked(LstItems.SelectedIndex, True)
                End If
            End If
        End If
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
    Private Sub optDaily_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OptDaily.CheckedChanged
        If eventSender.Checked Then
            '*******************************************************************************
            'Author             :   Nisha Rai
            'Argument(s)if any  :
            'Return Value       :   NA
            'Function           :   when daily mkt schedule is selected
            'Comments           :   NA
            'Creation Date      :   18/04/2001
            '*******************************************************************************
            Dim strItemCodes As String
            On Error GoTo ErrHandler
            If CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                strItemCodes = mStrItemCodes
            End If
            DTPTransDate.CustomFormat = "MM/yyyy"
            lblSchedule.Text = "Schedule Month"
            fpSDailySchedule.Visible = True
            '    fpSMonthlySchedule.Visible = False
            Call FillList()
            If CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                mStrItemCodes = strItemCodes
            End If
            Exit Sub
ErrHandler:
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
            Exit Sub
        End If
    End Sub
    Private Sub fillSpread()
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To fill data in grid according to mode (view,edit,add)
        '                       (daily,Monthly)
        'Comments           :   NA
        'Creation Date      :   18/04/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim i As Short
        Dim strSql As String
        Dim strDate As String
        Dim IntRecNo As Short
        Dim intRowNo As Short
        Dim StrDiffitemCode As String
        Dim RsCalendar_Mst As ADODB.Recordset
        Dim varDummy As Object
        Dim VarTransDate As Object
        Dim StrTransDate As String
        Dim blnOpenSO As Boolean
        Dim rsDailyItem As New ClsResultSetDB
        Dim blnFlg As Boolean
        'Add new Column (Hidden)Open SO For Accounts Plug in
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.OBJ_FORM, Me, System.Windows.Forms.Cursors.WaitCursor)
        fpSDailySchedule.MaxRows = 0
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        Dim intDaydiff As Integer

        If OptDaily.Checked = True Then
            fpSDailySchedule.Visible = True
            If Me.CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                StrDiffitemCode = ReturnDrgITemCode()
                strSql = "select  a.Account_Code,a.Item_Code,b.description," & " authorized_flag=1," & " Order_Qty=sum(a.Order_Qty),Despatch_Qty=sum(a.Despatch_Qty),a.active_flag" & ",a.cust_drgno" & " from cust_ord_dtl a,item_mst b, custitem_mst c " & " where A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE=C.UNIT_CODE AND A.UNIT_CODE='" + gstrUNITID + "' AND a.item_code=b.item_code and a.cust_drgno = c.cust_drgno and  a.Account_Code= '" & TxtAccountCode.Text & "' and a.account_code = c.account_code and " & " authorized_flag=1 and a.active_flag='A'" & " and a.account_code = c.account_code "
                If Len(Trim(StrDiffitemCode)) > 0 Then
                    strSql = strSql & StrDiffitemCode
                End If
                strSql = strSql & " group By a.Account_Code,a.Item_Code,b.description," & " a.active_flag,a.cust_drgno " ' remove this filed from group by a.Authorised_Flag,
                If mRdoCls.GetResult(strSql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly) And mRdoCls.GetNoRows > 0 Then
                    mRdoCls.MoveFirst()
                    Do While Not mRdoCls.EOFRecord
                        IntRecNo = IntRecNo + 1 '
                        'To Check Open SO Flag in Cust_Ord_Dtl
                        If CheckForOpenSoFlag(mRdoCls.GetValue("cust_drgno")) = True Then
                            blnOpenSO = True
                        Else
                            blnOpenSO = False
                        End If
                        For i = 1 To NoOfDaysinMonth(CShort(VB.Left(VB6.Format(DTPTransDate.Value, "MM/yyyy"), 2)), CShort(VB.Right(VB6.Format(DTPTransDate.Value, "MM/yyyy"), 4)))
                            strSql = "select Account_Code,Trans_date,Item_code,Cust_Drgno,Serial_No,Schedule_Flag,Schedule_Quantity,Despatch_Qty,Status,RevisionNo,DSNO,DSDateTime,Spare_qty,Consignee_Code,FILETYPE,DOC_NO,MSG_NAME from dailymktschedule "
                            strSql = strSql & " Where UNIT_CODE='" + gstrUNITID + "' And account_code='" & TxtAccountCode.Text & "'"
                            strSql = strSql & " and datepart(mm,trans_date) = " & Month(DTPTransDate.Value)
                            strSql = strSql & " and datepart(yyyy,trans_date) = " & Year(DTPTransDate.Value)
                            strSql = strSql & " and datepart(dd,trans_date) = " & i
                            strSql = strSql & " and cust_drgNo = '" & mRdoCls.GetValue("cust_drgno") & "' and Item_code = '" & mRdoCls.GetValue("Item_code") & "' and status =1"
                            rsDailyItem.GetResult(strSql)
                            If rsDailyItem.GetNoRows = 0 Then
                                fpSDailySchedule.MaxRows = fpSDailySchedule.MaxRows + 1
                                fpSDailySchedule.Row = fpSDailySchedule.MaxRows
                                SetGridDateFormat(fpSDailySchedule, 3)
                                If blnOpenSO = True Then
                                    fpSDailySchedule.Row = fpSDailySchedule.MaxRows : fpSDailySchedule.Row2 = fpSDailySchedule.MaxRows : fpSDailySchedule.Col = 1 : fpSDailySchedule.Col2 = 1 : fpSDailySchedule.BlockMode = True : fpSDailySchedule.CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox : fpSDailySchedule.Value = "1" : fpSDailySchedule.ColHidden = True : fpSDailySchedule.BlockMode = False
                                Else
                                    fpSDailySchedule.Row = fpSDailySchedule.MaxRows : fpSDailySchedule.Row2 = fpSDailySchedule.MaxRows : fpSDailySchedule.Col = 1 : fpSDailySchedule.Col2 = 1 : fpSDailySchedule.BlockMode = True : fpSDailySchedule.CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox : fpSDailySchedule.Value = "0" : fpSDailySchedule.ColHidden = True : fpSDailySchedule.BlockMode = False
                                End If
                                fpSDailySchedule.Col = 3
                                fpSDailySchedule.Text = VB6.Format(Format(i, "00") & "/" & Format(DTPTransDate.Value, "MM/yyyy"), gstrDateFormat)
                                VarTransDate = fpSDailySchedule.Text
                                fpSDailySchedule.Col = 4
                                fpSDailySchedule.Text = mRdoCls.GetValue("description")
                                fpSDailySchedule.Col = 5
                                fpSDailySchedule.Text = mRdoCls.GetValue("item_code")
                                fpSDailySchedule.Col = 6
                                fpSDailySchedule.Text = 0
                                fpSDailySchedule.Col = 7
                                fpSDailySchedule.Text = "1" 'mRdoCls.GetValue("item_code")
                                If Trim(VarTransDate) <> "" Then
                                    'StrTransDate = VB.Right(VarTransDate, 4) & "/" & Mid(VarTransDate, 4, 2) & "/" & VB.Left(VarTransDate, 2)
                                    If ConvertToDate(VarTransDate) < mStrCurrentDate Then
                                        fpSDailySchedule.Text = ""
                                    End If
                                End If
                                fpSDailySchedule.Col = 8
                                fpSDailySchedule.Row = fpSDailySchedule.MaxRows : fpSDailySchedule.Row2 = fpSDailySchedule.MaxRows : fpSDailySchedule.Col = 8 : fpSDailySchedule.Col2 = 8 : fpSDailySchedule.BlockMode = True : fpSDailySchedule.Lock = True : fpSDailySchedule.BlockMode = False

                                fpSDailySchedule.Text = mRdoCls.GetValue("cust_drgno")
                                fpSDailySchedule.Col = 9
                                fpSDailySchedule.Text = mRdoCls.GetValue("Order_qty")
                                fpSDailySchedule.Col = 10
                                ' Change By Deepak on 11-Oct-2011 for Change Management-------------
                                'issue id 10108966
                                'fpSDailySchedule.Text = mRdoCls.GetValue("Despatch_qty")
                                fpSDailySchedule.Text = "0"
                                'issue id 10108966
                                '---------------------------------------------------
                                fpSDailySchedule.Col = 11
                                fpSDailySchedule.Text = "0"
                            End If
                        Next i
                        RsCalendar_Mst = New ADODB.Recordset
                        RsCalendar_Mst.Open("SELECT CONVERT(VARCHAR(20),DT,103) as Dt FROM CALENDAR_MST " & " Where UNIT_CODE='" + gstrUNITID + "' AND right(CONVERT(varCHAR(20),DT,103),7)='" & VB6.Format(DTPTransDate.Value, "MM/yyyy") & "' and work_flg=1 ", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        If RsCalendar_Mst.RecordCount > 0 Then
                            RsCalendar_Mst.MoveFirst()
                            fpSDailySchedule.CursorStyle = FPSpreadADO.CursorStyleConstants.CursorStyleDefault
                            Do While Not RsCalendar_Mst.EOF
                                strDate = VB6.Format(RsCalendar_Mst.Fields("Dt").Value, gstrDateFormat)
                                'intRowNo = CShort(VB.Left(RsCalendar_Mst.Fields(0).Value, 2))
                                intRowNo = VB.Day(ConvertToDate(VB6.Format(RsCalendar_Mst.Fields(0).Value, gstrDateFormat)))
                                intMaxLoop = fpSDailySchedule.MaxRows
                                For intLoopCounter = 1 To intMaxLoop
                                    fpSDailySchedule.Col = 3 : fpSDailySchedule.Row = intLoopCounter

                                    If DateDiff(Microsoft.VisualBasic.DateInterval.Day, ConvertToDate(VB6.Format(fpSDailySchedule.Text, gstrDateFormat)), ConvertToDate(strDate)) = 0 Then
                                        fpSDailySchedule.Row = intLoopCounter
                                        fpSDailySchedule.Col = 2
                                        fpSDailySchedule.ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF)
                                        fpSDailySchedule.Col = 3
                                        fpSDailySchedule.ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF)
                                        fpSDailySchedule.Col = 4
                                        fpSDailySchedule.ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF)
                                        fpSDailySchedule.Col = 5
                                        fpSDailySchedule.ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF)
                                        fpSDailySchedule.Col = 6
                                        fpSDailySchedule.ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF)
                                        fpSDailySchedule.Col = 7
                                        fpSDailySchedule.ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF)
                                        fpSDailySchedule.Text = ""
                                    End If
                                Next
                                If Not RsCalendar_Mst.EOF Then
                                    RsCalendar_Mst.MoveNext()
                                End If
                            Loop
                            RsCalendar_Mst.Close()

                            RsCalendar_Mst = Nothing
                        End If
                        If Not mRdoCls.EOFRecord Then
                            mRdoCls.MoveNext()
                        End If
                    Loop
                    Call DisablePrimaryKeyControl()
                    Call MakeEditableAndEditableLinesofGrid()
                End If
            ElseIf Me.CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
                StrDiffitemCode = ReturnDrgITemCode()
                strSql = "select b.description ,a.Item_Code,a.Schedule_Quantity  " & ",a.Schedule_Flag,a.cust_drgno ,trans_date= convert(varchar(20),a.trans_date,103),a.despatch_qty,a.RevisionNo  " & " from dailymktschedule a,item_mst b,custitem_mst c WHERE A.UNIT_CODE=B.UNIT_CODE AND A.UNIT_CODE=C.UNIT_CODE AND  A.UNIT_CODE='" + gstrUNITID + "' AND a.item_code=b.item_code and a.item_code = c.item_code and a.cust_drgno = c.cust_drgno " & " and a.status =1 and a.Account_Code= '" & TxtAccountCode.Text & "' and  a.account_code = c.account_code  and right(convert(varchar(20),a.trans_date,103),7)='" & VB6.Format(DTPTransDate.Value, "MM/yyyy") & "'"
                If Len(Trim(StrDiffitemCode)) > 0 Then
                    strSql = strSql & StrDiffitemCode
                End If
                strSql = strSql & " ORDER BY a.Item_Code,a.cust_drgno,trans_date"
                If mRdoCls.GetResult(strSql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly) And mRdoCls.GetNoRows > 0 Then
                    mRdoCls.MoveFirst()

                    Do While Not mRdoCls.EOFRecord
                        If CheckForOpenSoFlag(mRdoCls.GetValue("cust_drgno")) = True Then
                            blnOpenSO = True
                        Else
                            blnOpenSO = False
                        End If
                        fpSDailySchedule.MaxRows = fpSDailySchedule.MaxRows + 1
                        fpSDailySchedule.Row = fpSDailySchedule.MaxRows
                        SetGridDateFormat(fpSDailySchedule, 3)
                        'To Check Open SO Flag in Cust_Ord_Dtl
                        If blnOpenSO = True Then
                            fpSDailySchedule.Row = fpSDailySchedule.MaxRows : fpSDailySchedule.Row2 = fpSDailySchedule.MaxRows : fpSDailySchedule.Col = 1 : fpSDailySchedule.Col2 = 1 : fpSDailySchedule.BlockMode = True : fpSDailySchedule.CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox : fpSDailySchedule.Value = CStr(System.Windows.Forms.CheckState.Checked) : fpSDailySchedule.ColHidden = True : fpSDailySchedule.BlockMode = False
                        Else
                            fpSDailySchedule.Row = fpSDailySchedule.MaxRows : fpSDailySchedule.Row2 = fpSDailySchedule.MaxRows : fpSDailySchedule.Col = 1 : fpSDailySchedule.Col2 = 1 : fpSDailySchedule.BlockMode = True : fpSDailySchedule.CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox : fpSDailySchedule.Value = CStr(System.Windows.Forms.CheckState.Unchecked) : fpSDailySchedule.ColHidden = True : fpSDailySchedule.BlockMode = False
                        End If
                        '************
                        fpSDailySchedule.Col = 3
                        fpSDailySchedule.Text = VB6.Format(mRdoCls.GetValue("trans_date"), gstrDateFormat)
                        fpSDailySchedule.Col = 4
                        fpSDailySchedule.Text = mRdoCls.GetValue("description")
                        fpSDailySchedule.Col = 5
                        fpSDailySchedule.Text = mRdoCls.GetValue("item_code")
                        fpSDailySchedule.Col = 6
                        fpSDailySchedule.Text = CStr(Val(mRdoCls.GetValue("Schedule_Quantity")))
                        fpSDailySchedule.Col = 12
                        fpSDailySchedule.Text = CStr(Val(mRdoCls.GetValue("Schedule_Quantity")))
                        fpSDailySchedule.Col = 7
                        If mRdoCls.GetValue("Schedule_flag") Then
                            fpSDailySchedule.Text = "1"
                        Else
                            fpSDailySchedule.Text = "0"
                        End If
                        fpSDailySchedule.Col = 13

                        If mRdoCls.GetValue("Schedule_flag") Then
                            fpSDailySchedule.Text = "1"
                        Else
                            fpSDailySchedule.Text = "0"
                        End If
                        fpSDailySchedule.Col = 8
                        fpSDailySchedule.Row = fpSDailySchedule.MaxRows : fpSDailySchedule.Row2 = fpSDailySchedule.MaxRows : fpSDailySchedule.Col = 8 : fpSDailySchedule.Col2 = 8 : fpSDailySchedule.BlockMode = True : fpSDailySchedule.Lock = True : fpSDailySchedule.BlockMode = False
                        fpSDailySchedule.Text = mRdoCls.GetValue("cust_drgno")
                        fpSDailySchedule.Col = 10
                        fpSDailySchedule.Text = CStr(Val(mRdoCls.GetValue("despatch_qty")))
                        fpSDailySchedule.Col = 11
                        fpSDailySchedule.Text = CStr(Val(mRdoCls.GetValue("RevisionNo")))
                        'ISSUE ID :10190371
                        intDaydiff = DateDiff(Microsoft.VisualBasic.DateInterval.Day, GetServerDate, mRdoCls.GetValue("trans_date"))

                        If intDaydiff < 0 Then
                            fpSDailySchedule.BlockMode = True
                            fpSDailySchedule.Col = 2
                            fpSDailySchedule.Col2 = 10
                            fpSDailySchedule.Row = fpSDailySchedule.MaxRows
                            fpSDailySchedule.Row2 = fpSDailySchedule.MaxRows
                            fpSDailySchedule.Lock = True
                            fpSDailySchedule.BlockMode = False
                        Else
                            fpSDailySchedule.BlockMode = True
                            fpSDailySchedule.Col = 2
                            fpSDailySchedule.Col2 = 10
                            fpSDailySchedule.Row = fpSDailySchedule.MaxRows
                            fpSDailySchedule.Row2 = fpSDailySchedule.MaxRows
                            fpSDailySchedule.Lock = False
                            fpSDailySchedule.BlockMode = False
                        End If 'ISSUE ID DONE : 10190371
                        If Not mRdoCls.EOFRecord Then
                            mRdoCls.MoveNext()
                        End If

                    Loop
                    Call ChangetheColorofGridInViewMode()
                End If
            ElseIf Me.CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                scheduleEdit = False
                StrDiffitemCode = ReturnDrgITemCode()
                strSql = "select trans_date= convert(varchar(20),a.trans_date,103) ,b.description ,a.Item_Code,Schedule_Quantity " & ",a.Schedule_Flag,a.cust_drgno,a.despatch_qty,RevisionNo   " & " from dailymktschedule a,item_mst b, custitem_mst c " & " Where A.Unit_Code=B.Unit_Code And A.Unit_Code=C.Unit_Code And A.Unit_code='" + gstrUNITID + "' And   a.item_code=b.item_code and a.item_code = c.item_code and a.cust_drgno = c.cust_drgno  " & " and  a.Account_Code= '" & TxtAccountCode.Text & "' and a.account_code = c.account_code  and  right(convert(varchar(20),a.trans_date,103),7)='" & VB6.Format(DTPTransDate.Value, "MM/yyyy") & "' "
                If Len(Trim(StrDiffitemCode)) > 0 Then
                    'and a.cust_drgno in(" & StrDiffitemCode & ")
                    strSql = strSql & StrDiffitemCode
                End If
                strSql = strSql & " and a.Status =1 ORDER BY a.Item_Code,a.cust_drgno,trans_date"
                If mRdoCls.GetResult(strSql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly) And mRdoCls.GetNoRows > 0 Then
                    mRdoCls.MoveFirst()
                    Do While Not mRdoCls.EOFRecord
                        'To Check Open SO Flag in Cust_Ord_Dtl
                        If CheckForOpenSoFlag(mRdoCls.GetValue("cust_drgno")) = True Then
                            blnOpenSO = True
                        Else
                            blnOpenSO = False
                        End If
                        fpSDailySchedule.MaxRows = fpSDailySchedule.MaxRows + 1
                        fpSDailySchedule.Row = fpSDailySchedule.MaxRows
                        SetGridDateFormat(fpSDailySchedule, 3)
                        'To Check Open SO Flag in Cust_Ord_Dtl
                        If blnOpenSO = True Then
                            fpSDailySchedule.Row = fpSDailySchedule.MaxRows : fpSDailySchedule.Row2 = fpSDailySchedule.MaxRows : fpSDailySchedule.Col = 1 : fpSDailySchedule.Col2 = 1 : fpSDailySchedule.BlockMode = True : fpSDailySchedule.CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox : fpSDailySchedule.Value = CStr(System.Windows.Forms.CheckState.Checked) : fpSDailySchedule.ColHidden = True : fpSDailySchedule.BlockMode = False
                        Else
                            fpSDailySchedule.Row = fpSDailySchedule.MaxRows : fpSDailySchedule.Row2 = fpSDailySchedule.MaxRows : fpSDailySchedule.Col = 1 : fpSDailySchedule.Col2 = 1 : fpSDailySchedule.BlockMode = True : fpSDailySchedule.CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox : fpSDailySchedule.Value = CStr(System.Windows.Forms.CheckState.Unchecked) : fpSDailySchedule.ColHidden = True : fpSDailySchedule.BlockMode = False
                        End If
                        fpSDailySchedule.Col = 3
                        fpSDailySchedule.Text = VB6.Format(mRdoCls.GetValue("trans_date"), gstrDateFormat)
                        'fpSDailySchedule.Text = mRdoCls.GetValue("trans_date")
                        fpSDailySchedule.Col = 4
                        fpSDailySchedule.Text = mRdoCls.GetValue("description")
                        fpSDailySchedule.Col = 5
                        fpSDailySchedule.Text = mRdoCls.GetValue("item_code")
                        fpSDailySchedule.Col = 6
                        fpSDailySchedule.Text = CStr(Val(mRdoCls.GetValue("Schedule_Quantity")))
                        fpSDailySchedule.Col = 12
                        fpSDailySchedule.Text = CStr(Val(mRdoCls.GetValue("Schedule_Quantity")))
                        fpSDailySchedule.Col = 7

                        If mRdoCls.GetValue("Schedule_flag") = "0" Then
                            fpSDailySchedule.Text = "" 'mRdoCls.GetValue("Schedule_flag")
                        Else
                            scheduleEdit = True
                            fpSDailySchedule.Text = "1"
                        End If
                        fpSDailySchedule.Col = 13
                        If mRdoCls.GetValue("Schedule_flag") = "0" Then
                            fpSDailySchedule.Text = "" 'mRdoCls.GetValue("Schedule_flag")
                        Else
                            scheduleEdit = True
                            fpSDailySchedule.Text = "1"
                        End If

                        fpSDailySchedule.Col = 8

                        fpSDailySchedule.Text = mRdoCls.GetValue("cust_drgno")

                        fpSDailySchedule.Col = 10

                        fpSDailySchedule.Text = CStr(Val(mRdoCls.GetValue("despatch_qty")))
                        fpSDailySchedule.Col = 11

                        fpSDailySchedule.Text = CStr(Val(mRdoCls.GetValue("revisionNo")))


                        If Not mRdoCls.EOFRecord Then
                            mRdoCls.MoveNext()
                        End If
                    Loop

                    fpSDailySchedule.BlockMode = True
                    fpSDailySchedule.Col = 2
                    fpSDailySchedule.Col2 = 10
                    fpSDailySchedule.Row = 1
                    fpSDailySchedule.Row2 = fpSDailySchedule.MaxRows
                    fpSDailySchedule.Lock = True
                    fpSDailySchedule.BlockMode = False
                    'to change the fore color of Spreads data
                    Call ChangetheColorofGridInViewMode()
                End If
            End If
        End If
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.OBJ_FORM, Me, System.Windows.Forms.Cursors.Default)
        '*****
        Exit Sub
ErrHandler:

        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.OBJ_FORM, Me, System.Windows.Forms.Cursors.Default)
        Exit Sub
    End Sub
    Private Sub SaveAddDaily(ByRef strMode As String)
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :   strMode
        'Return Value       :   NA
        'Function           :   to save data filled in grid in Add mode in case of Daily
        '                       mkt schedule
        'Comments           :   NA
        'Creation Date      :   18/04/2001
        'History            :   Changes done by Sourabh on 02 Sep 2004 For Add two new column DSNO and DSDATETIME
        'History            :   Changes done by Sourabh on 30 june 2005 against PIMS no PRJ-2004-04-003-15106
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim i As Short
        Dim Str_Account_Code As String
        Dim Str_Trans_date As String
        Dim Str_cust_drgno As String
        Dim strCustDrgNo As String
        'dim Str_Serial_No As String
        Dim Str_Schedule_Flag As String
        Dim Str_PrevSchedule_Flag As String
        Dim StrScheduleFlag As String
        Dim Str_Item_Code As String
        Dim StrItemCode As String
        Dim Sng_Schedule_Quantity As Double
        Dim Sng_PrevSchedule_Quantity As Double
        Dim SngScheduleQuantity As Double
        Dim SngRevisionNo As Double
        'dim Str_Ent_dt As String
        Dim Str_Ent_UserId As String
        Dim Str_Upd_UserId As String
        Dim strSql As String
        Dim intRowCount As Short
        Dim intMaxRowCount As Short
        Dim dblOrderQty As Double
        Dim dblDispatchqty As Double
        Dim dblItemRows As Double
        Dim blninsertDaily As Boolean
        Dim strDSNo As String
        Dim dtDatetime As Date
        Dim rsDSTracking As ClsResultSetDB
        Dim RsForecast As ClsResultSetDB
        Dim blnDSTracking As Boolean
        Dim BlnUpdateForecast As Boolean
        rsDSTracking = New ClsResultSetDB
        rsDSTracking.GetResult("select isnull(Update_Forecast,0) as Update_Forecast From Sales_Parameter Where Unit_Code='" + gstrUNITID + "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        BlnUpdateForecast = rsDSTracking.GetValue("Update_Forecast")
        rsDSTracking.ResultSetClose()
        rsDSTracking = Nothing
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.OBJ_FORM, Me, System.Windows.Forms.Cursors.WaitCursor)
        Str_Ent_UserId = "'" & mP_User & "'"
        Str_Upd_UserId = "'" & mP_User & "'"
        Str_Account_Code = "'" & TxtAccountCode.Text & "'"
        With fpSDailySchedule
            dblItemRows = 0
            For i = 1 To fpSDailySchedule.MaxRows
                fpSDailySchedule.Row = i
                fpSDailySchedule.Col = 2
                '***to initialize value of dblitemrows
                If i > dblItemRows Then
                    blninsertDaily = False
                    .Col = 5
                    StrItemCode = .Text
                    .Col = 6
                    SngScheduleQuantity = CDbl(.Text)
                    .Col = 8
                    strCustDrgNo = .Text
                    intMaxRowCount = fpSDailySchedule.MaxRows
                    For intRowCount = dblItemRows + 1 To intMaxRowCount
                        .Row = intRowCount : .Col = 5
                        If Trim(StrItemCode) = Trim(.Text) Then
                            .Row = intRowCount : .Col = 8
                            dblItemRows = intRowCount
                            If Trim(strCustDrgNo) = Trim(.Text) Then
                                .Row = intRowCount : .Col = 6
                                dblItemRows = intRowCount
                                If CDbl(.Text) > 0 Then
                                    blninsertDaily = True
                                End If
                            Else
                                Exit For
                            End If
                        Else
                            Exit For
                        End If
                    Next
                End If
                .Row = i : .Col = 3
                Str_Trans_date = .Text  ' mP_DayFormat
                .Row = i : .Col = 5
                Str_Item_Code = "'" & .Text & "'"
                .Col = 6
                Sng_Schedule_Quantity = CDbl(.Text)
                If strMode <> "A" Then
                    .Col = 12
                    Sng_PrevSchedule_Quantity = CDbl(.Text)
                End If
                .Row = i : .Col = 7
                If .Text = "1" Then
                    Str_Schedule_Flag = "'1'"
                Else
                    Str_Schedule_Flag = "'0'"
                End If
                If strMode <> "A" Then
                    .Row = i : .Col = 13
                    If .Text = "1" Then
                        Str_PrevSchedule_Flag = "'1'"
                    Else
                        Str_PrevSchedule_Flag = "'0'"
                    End If
                End If
                .Row = i : .Col = 8
                Str_cust_drgno = "'" & .Text & "'"
                .Row = i : .Col = 11
                SngRevisionNo = Val(.Text)
                strDSNo = "ECSS" : dtDatetime = CDate(GetServerDate() & vbCrLf & TimeOfDay)
                rsDSTracking = New ClsResultSetDB
                Call rsDSTracking.GetResult("Select DSWiseTracking From Sales_parameter Where Unit_Code='" + gstrUNITID + "'", ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                If rsDSTracking.RowCount > 0 Then blnDSTracking = IIf(IsDBNull(rsDSTracking.GetValue("DSWiseTracking")), False, IIf(rsDSTracking.GetValue("DSwisetracking") = False, False, True))
                rsDSTracking.ResultSetClose()
                rsDSTracking = Nothing
                If strMode = "A" Then
                    If blninsertDaily = True Then
                        strSql = ""
                        If Sng_Schedule_Quantity > 0 Then
                            'Changed for Issue ID eMpro-20090227-27987 -(Added Consignee Code) Starts
                            strSql = "SET DATEFORMAT 'DMY' Insert into dailymktschedule ( " & "Account_Code,Trans_date,cust_drgno," & "Schedule_Flag,Item_Code,Schedule_Quantity,Status," & "Ent_UserId,Upd_UserId,Ent_dt,Upd_dt,filetype,RevisionNo,Consignee_code"
                            If blnDSTracking = True Then
                                strSql = strSql & ",DSNo,DSDateTime"
                            End If
                            strSql = strSql & ",Unit_Code) values (" & Str_Account_Code & ",'" & getDateForDB(Str_Trans_date) & "'," & Str_cust_drgno & "," & Str_Schedule_Flag & "," & Str_Item_Code & "," & Sng_Schedule_Quantity & ",1,'" & Trim(mP_User) & "','" & Trim(mP_User) & "',getdate(),getdate(),'Manual',0," & Str_Account_Code & ""
                            If blnDSTracking = True Then
                                strSql = strSql & ",'" & strDSNo & "','" & getDateForDB(dtDatetime) & "'"
                            End If
                            strSql = strSql & ",'" + gstrUNITID + "')"
                            'Changed for Issue ID eMpro-20090227-27987 -(Added Consignee Code) Ends
                            If BlnUpdateForecast = True Then
                                strSql = strSql & vbCrLf
                                strSql = strSql & "INSERT INTO forecast_mst(Customer_code,product_no,Due_date,Quantity,ent_userid,ent_dt,upd_userid,upd_dt, ENagare_UNLOC,Unit_Code)"
                                strSql = strSql & "VALUES (" & Str_Account_Code & "," & Str_Item_Code & ",'" & getDateForDB(Str_Trans_date) & "'," & Sng_Schedule_Quantity & ",'" & mP_User & "',getdate(),'" & mP_User & "',getdate(),'DSCH','" + gstrUNITID + "')"
                            End If
                        End If
                    End If
                Else
                    strSql = ""
                    If (Sng_Schedule_Quantity) <> (Sng_PrevSchedule_Quantity) Or (Str_Schedule_Flag) <> (Str_PrevSchedule_Flag) Then
                        'Changed for Issue ID eMpro-20090227-27987 -(Added Consignee Code) Starts
                        strSql = "set dateformat 'DMY' Insert into dailymktschedule ( " & "Account_Code,Trans_date,cust_drgno," & "Schedule_Flag,Item_Code,Schedule_Quantity,Despatch_Qty,Status," & "Ent_UserId,Upd_UserId,Ent_dt,Upd_dt,filetype,RevisionNo,Consignee_Code"
                        If blnDSTracking = True Then
                            strSql = strSql & ",DSNo,DSDateTime"
                        End If
                        strSql = strSql & ",Unit_Code) values (" & Str_Account_Code & ",'" & getDateForDB(Str_Trans_date) & "'," & Str_cust_drgno & "," & Str_Schedule_Flag & "," & Str_Item_Code & "," & Sng_Schedule_Quantity & "," & mdblDespatchQty(i - 1) & ",1,'" & Trim(mP_User) & "','" & Trim(mP_User) & "',getdate(),getdate(),'Manual'," & SngRevisionNo + 1 & "," & Str_Account_Code & ""
                        If blnDSTracking = True Then
                            strSql = strSql & ",'" & strDSNo & "','" & getDateForDB(dtDatetime) & "'"
                        End If
                        strSql = strSql & ",'" + gstrUNITID + "')"
                        If BlnUpdateForecast = True Then
                            RsForecast = New ClsResultSetDB
                            RsForecast.GetResult("Select Customer_code from Forecast_Mst Where Unit_Code='" + gstrUNITID + "' And Customer_code=" & Str_Account_Code & " AND Due_date='" & getDateForDB(Str_Trans_date) & "' AND product_no=" & Str_Item_Code & " AND ENagare_UNLOC='DSCH'")
                            If RsForecast.GetNoRows > 0 Then
                                strSql = strSql & "UPDATE forecast_mst Set Quantity =" & Sng_Schedule_Quantity - GetNagareQty(i, 6) & ",upd_userid='" & mP_User & "',upd_dt=getdate() where Unit_Code='" + gstrUNITID + "' And  Customer_code=" & Str_Account_Code & " AND Due_date='" & getDateForDB(Str_Trans_date) & "' AND product_no=" & Str_Item_Code & " AND ENagare_UNLOC='DSCH'"
                            Else

                                strSql = strSql & vbCrLf & "INSERT INTO forecast_mst(Customer_code,product_no,Due_date,Quantity,ent_userid,ent_dt,upd_userid,upd_dt, ENagare_UNLOC,Unit_Code)"
                                strSql = strSql & vbCrLf & "VALUES (" & Str_Account_Code & "," & Str_Item_Code & ",'" & getDateForDB(Str_Trans_date) & "'," & Sng_Schedule_Quantity - GetNagareQty(i, 6) & ",'" & mP_User & "',getdate(),'" & mP_User & "',getdate(),'DSCH','" + gstrUNITID + "')"
                            End If
                            RsForecast.ResultSetClose()
                            RsForecast = Nothing
                        End If
                    End If
                End If

                If Len(Trim(strSql)) > 0 Then
                    mP_Connection.Execute(strSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                    strSql = ""
                End If
            Next i
        End With
        CmdGrpMktSchedule.Revert()

        CmdGrpMktSchedule.Enabled(1) = False

        CmdGrpMktSchedule.Enabled(2) = False
        Call EnablePrimaryKeyControl()
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.OBJ_FORM, Me, System.Windows.Forms.Cursors.Default)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.OBJ_FORM, Me, System.Windows.Forms.Cursors.Default)
        Exit Sub
    End Sub
    Private Sub SaveUpdateDaily()
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :   strMode
        'Return Value       :   NA
        'Function           :   to save data filled in grid in edit mode in case of
        '                       daily marketing schedule
        'Comments           :   NA
        'Creation Date      :   18/04/2001
        '*******************************************************************************
        Dim rsSchedule As ClsResultSetDB
        Dim i As Short
        Dim Str_Account_Code As String
        Dim Str_Trans_date As String
        Dim Str_cust_drgno As String
        Dim Str_Schedule_Flag As String
        Dim Str_PrevSchedule_Flag As String
        Dim Str_Item_Code As String
        Dim Sng_Schedule_Quantity As Double
        Dim Sng_PrevSchedule_Quantity As Double
        Dim Str_Ent_UserId As String
        Dim Str_Upd_UserId As String
        Dim strSql As String
        Dim dblOrderQty As Double
        Dim dblDispatchqty As Double

        On Error GoTo ErrHandler

        'Add new Column (Hidden)Open SO For Accounts Plug in
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.OBJ_FORM, Me, System.Windows.Forms.Cursors.WaitCursor)
        Str_Ent_UserId = "'" & mP_User & "'"
        Str_Upd_UserId = "'" & mP_User & "'"
        Str_Account_Code = "'" & TxtAccountCode.Text & "'"
        mStrCurrentDate = GetServerDate()
        ReDim mdblDespatchQty(fpSDailySchedule.MaxRows - 1)

        For i = 1 To fpSDailySchedule.MaxRows
            fpSDailySchedule.Row = i
            fpSDailySchedule.Col = 2
            With fpSDailySchedule
                .Col = 3
                Str_Trans_date = .Text ' mP_DayFormat
                .Col = 5
                Str_Item_Code = "'" & .Text & "'"
                .Col = 6
                Sng_Schedule_Quantity = CDbl(.Text)
                .Col = 12
                Sng_PrevSchedule_Quantity = CDbl(.Text)
                .Col = 7
                If .Text = "1" Then
                    Str_Schedule_Flag = "'1'"
                Else
                    Str_Schedule_Flag = "'0'"
                End If
                .Col = 13
                If .Text = "1" Then
                    Str_PrevSchedule_Flag = "'1'"
                Else
                    Str_PrevSchedule_Flag = "'0'"
                End If
                .Col = 8
                Str_cust_drgno = "'" & .Text & "'"

            End With
            '10510618
            'strSql = " Select Despatch_Qty from dailymktschedule " & " Where Unit_Code='" + gstrUNITID + "' And account_code=" & Str_Account_Code & " and  convert(varchar(20), trans_date ,103) ='" & getDateForDB(Str_Trans_date) & "' and cust_drgno= " & Str_cust_drgno & " And Status = 1"
            strSql = " Select Despatch_Qty from dailymktschedule " & " Where Unit_Code='" + gstrUNITID + "' aND account_code=" & Str_Account_Code & " and trans_date ='" & getDateForDB(Str_Trans_date) & "' and cust_drgno= " & Str_cust_drgno & " And Status = 1"
            '10510618
            rsSchedule = New ClsResultSetDB
            rsSchedule.GetResult(strSql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            rsSchedule.MoveFirst()

            mdblDespatchQty(i - 1) = Val(rsSchedule.GetValue("Despatch_Qty"))
            strSql = ""
            If (Sng_Schedule_Quantity <> Sng_PrevSchedule_Quantity) Or (Str_Schedule_Flag <> Str_PrevSchedule_Flag) Then
                'Changed for Issue ID eMpro-20090604-32080 --Schedule Flag =0,Remove File Type updation
                strSql = "set dateformat DMY update dailymktschedule set  " & "Status =0 ,schedule_Flag=0, Upd_dt='" & getDateForDB(mStrCurrentDate) & "',Upd_UserId =" & Str_Upd_UserId & " where Unit_Code='" + gstrUNITID + "' And account_code=" & Str_Account_Code & " and  convert(varchar(20), trans_date ,106) ='" & getDateForDB(Str_Trans_date) & "' and cust_drgno= " & Str_cust_drgno
            End If
            rsSchedule.ResultSetClose()
            rsSchedule = Nothing
            If Len(Trim(strSql)) > 0 Then
                mP_Connection.Execute(strSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End If
            strSql = "select order_qty=sum(order_qty),despatch_qty=sum(despatch_qty)" & " from cust_ord_dtl Where Unit_Code='" + gstrUNITID + "' And account_code='" & TxtAccountCode.Text & "' and item_code=" & Str_Item_Code & " and cust_Drgno = " & Str_cust_drgno & " and active_flag='A' and authorized_flag=1"
        Next i
        Call SaveAddDaily("E")

        CmdGrpMktSchedule.Revert()
        Call EnablePrimaryKeyControl()

        'CmdGrpMktSchedule.Ctlset_Enabled(2, False)
        CmdGrpMktSchedule.Enabled(2) = False

        CmdGrpMktSchedule.Enabled(1) = False
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.OBJ_FORM, Me, System.Windows.Forms.Cursors.Default)
        '*****
        ''Commented by nisha on 04/07/02
        'UpdateDailyMonthlyMktSchedulestatus
        '****
        'Call ConfirmWindow(10049, BUTTON_OK, IMG_INFO)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.OBJ_FORM, Me, System.Windows.Forms.Cursors.Default)
        Exit Sub
    End Sub
    Private Function FillList() As Boolean
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :   strMode
        'Return Value       :   NA
        'Function           :   to fill list with valid item code, customer item code
        '                       and description
        'Comments           :   NA
        'Creation Date      :   18/04/2001
        '*******************************************************************************
        Dim strSql As String
        Dim StrItemCode As New VB6.FixedLengthString(16)
        Dim strCustItemCode As New VB6.FixedLengthString(30)
        Dim strDailyDrgNo As String
        Dim intDaysOfMonth As Short
        Dim dtStartDate As Date
        Dim dtEndDate As Date
        Dim StritemDescription As New VB6.FixedLengthString(40)
        Dim StrCustItemDescription As New VB6.FixedLengthString(50)
        Dim rsDailyMktSchedule As New ClsResultSetDB
        Dim rsDailyNoOfRows As New ClsResultSetDB
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        On Error GoTo ErrHandler
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.OBJ_FORM, Me, System.Windows.Forms.Cursors.WaitCursor)
        LstItems.Items.Clear()

        If TxtAccountCode.Text = "" Then
            FillList = False
        End If
        If Me.CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            strSql = "select distinct cust_drgno,Item_code from dailymktschedule "
            strSql = strSql & " Where Unit_Code='" + gstrUNITID + "' And account_code='" & TxtAccountCode.Text & "' and datepart(mm,trans_date) = " & VB6.Format(DTPTransDate.Value, "MM")
            strSql = strSql & " and datepart(yyyy,trans_date) = " & VB6.Format(DTPTransDate.Value, "YYYY") & " and status =1"
            rsDailyMktSchedule.GetResult(strSql)
            dtStartDate = DateAdd(DateInterval.Day, 0 - CDbl(Format(DTPTransDate.Value, "dd")), DTPTransDate.Value)

            dtEndDate = DTPTransDate.Value.AddMonths(1)
            intMaxLoop = rsDailyMktSchedule.GetNoRows
            rsDailyMktSchedule.MoveFirst()
            If intMaxLoop > 0 Then
                intDaysOfMonth = DateDiff(Microsoft.VisualBasic.DateInterval.Day, dtStartDate, dtEndDate) - 1
                strDailyDrgNo = ""
                For intLoopCounter = 1 To intMaxLoop
                    strSql = "select Account_Code,Trans_date,Item_code,Cust_Drgno,Serial_No,Schedule_Flag,Schedule_Quantity,Despatch_Qty,Status,RevisionNo,DSNO,DSDateTime,Spare_qty,Consignee_Code,FILETYPE,DOC_NO,MSG_NAME from dailymktschedule "
                    strSql = strSql & " Where Unit_Code='" + gstrUNITID + "' And  account_code='" & TxtAccountCode.Text & "' and datepart(mm,trans_date) = " & VB6.Format(DTPTransDate.Value, "MM")
                    strSql = strSql & " and datepart(yyyy,trans_date) = " & VB6.Format(DTPTransDate.Value, "YYYY") & " and cust_drgNo = '" & rsDailyMktSchedule.GetValue("Cust_drgno") & "'"
                    strSql = strSql & " and Item_code = '" & rsDailyMktSchedule.GetValue("Item_Code") & "'  and status =1"
                    rsDailyNoOfRows.GetResult(strSql)
                    If rsDailyNoOfRows.GetNoRows = intDaysOfMonth Then
                        If Len(Trim(strDailyDrgNo)) > 0 Then
                            strDailyDrgNo = strDailyDrgNo & " and (b.Cust_drgNo <> '" & rsDailyMktSchedule.GetValue("Cust_drgno") & "'"
                            strDailyDrgNo = strDailyDrgNo & " or b.Item_code <> '" & rsDailyMktSchedule.GetValue("Item_code") & "')"
                        Else
                            strDailyDrgNo = " and (b.Cust_drgNo <> '" & rsDailyMktSchedule.GetValue("Cust_drgno") & "'"
                            strDailyDrgNo = strDailyDrgNo & " or b.Item_code <> '" & rsDailyMktSchedule.GetValue("Item_code") & "')"
                        End If
                    End If
                    rsDailyMktSchedule.MoveNext()
                Next
            End If
            'to pick item list to make daily/monthly mkt schedule
            strSql = "select distinct b.item_code,b.cust_drgno,a.description,b.cust_drg_desc from "
            strSql = strSql & " item_mst a,cust_ord_dtl b,custitem_mst c "
            strSql = strSql & " Where A.Unit_Code=B.Unit_code And A.Unit_Code=C.Unit_code And A.Unit_Code='" + gstrUNITID + "' And  b.item_code = c.item_code and b.cust_drgno = c.cust_drgno and b.account_code='" & TxtAccountCode.Text & "' and "
            strSql = strSql & " b.item_code=a.item_code and a.hold_flag <> 1 "
            strSql = strSql & " and a.Item_Main_Grp in ('F','T','S','C','P','R','M') "
            strSql = strSql & " and b.active_flag='A' and b.authorized_flag=1 and a.Status='A' and c.active=1 "
            strSql = strSql & " and (b.order_qty>b.despatch_qty or b.Order_Qty =0)"
            If Len(Trim(strDailyDrgNo)) > 0 Then
                strSql = strSql & strDailyDrgNo
            End If
            strSql = strSql & " and b.cust_drgno not in (select cust_drgno from monthlymktschedule "
            strSql = strSql & " Where Unit_Code='" + gstrUNITID + "' And account_code='" & TxtAccountCode.Text & "' and status=1 and "
            strSql = strSql & " year_month=" & VB6.Format(DTPTransDate.Value, "yyyyMM") & ") order by b.cust_drgno "
        ElseIf Me.CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Or Me.CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            If OptDaily.Checked = True Then
                strSql = "select distinct b.item_code,b.cust_drgno,a.description,c.drg_desc from "
                strSql = strSql & " item_mst a,dailymktschedule b,custitem_mst c "
                strSql = strSql & " where A.Unit_code=B.Unit_Code And A.Unit_code=C.Unit_code And A.Unit_Code='" + gstrUNITID + "' And b.account_code='" & TxtAccountCode.Text & "' and c.account_code='" & TxtAccountCode.Text & "' and "
                strSql = strSql & " b.item_code=a.item_code and b.cust_drgno=c.cust_drgno and b.Item_code=c.Item_code  and right(convert(varchar(20),b.trans_date,103),7)='" & VB6.Format(DTPTransDate.Value, "MM/yyyy") & "'"
                'added by nisha for Item_Main_Grp Check
                strSql = strSql & " and a.Item_Main_grp in ('F','T','S','C','P','R','M') and b.Status =1 and c.active=1 and A.Status='A' order by b.cust_drgno"
            End If
        End If

        If mRdoCls.GetResult(strSql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly) And mRdoCls.GetNoRows > 0 Then
            mRdoCls.MoveFirst()
            Do Until mRdoCls.EOFRecord
                StrItemCode.Value = mRdoCls.GetValue("Item_Code")
                strCustItemCode.Value = mRdoCls.GetValue("cust_drgno")
                StritemDescription.Value = mRdoCls.GetValue("description") 'Company Item Description
                If Me.CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Or Me.CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                    StrCustItemDescription.Value = mRdoCls.GetValue("drg_desc") 'Customer Item description
                Else
                    StrCustItemDescription.Value = mRdoCls.GetValue("cust_drg_desc") 'Customer Item description
                End If
                LstItems.Items.Add(strCustItemCode.Value & "     " & StrItemCode.Value & "     " & StritemDescription.Value & " " & StrCustItemDescription.Value)
                If Not mRdoCls.EOFRecord Then
                    mRdoCls.MoveNext()
                End If
            Loop
        End If
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.OBJ_FORM, Me, System.Windows.Forms.Cursors.Default)
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.OBJ_FORM, Me, System.Windows.Forms.Cursors.Default)
        Exit Function
    End Function
    Private Function ReturnitemCode() As String
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :   strMode
        'Return Value       :   NA
        'Function           :   to return Items selected in list
        'Comments           :   NA
        'Creation Date      :   18/04/2001
        '*******************************************************************************
        Dim StrItemCode As String
        Dim i As Short
        Dim strSql As String
        Dim dblOrderQty As Double
        Dim dblDispatchqty As Double
        On Error GoTo ErrHandler
        For i = 0 To LstItems.Items.Count - 1
            If LstItems.GetItemChecked(i) = True Then
                StrItemCode = StrItemCode & "'" & Mid(ObsoleteManagement.GetItemString(LstItems, i), 1, 30) & "',"
            End If
        Next i
        If Len(StrItemCode) > 0 Then
            If VB.Right(StrItemCode, 1) = "," Then
                StrItemCode = VB.Left(StrItemCode, Len(StrItemCode) - 1)
            End If
        ElseIf Len(StrItemCode) = 0 Then
            StrItemCode = "''"
        End If
        ReturnitemCode = StrItemCode
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Function
    End Function
    Private Function ReturnDrgITemCode() As String
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :   strMode
        'Return Value       :   NA
        'Function           :   to return Items selected in list
        'Comments           :   NA
        'Creation Date      :   18/04/2001
        '*******************************************************************************
        Dim StrItemCode As String
        Dim i As Short
        Dim strSql As String
        Dim dblOrderQty As Double
        Dim dblDispatchqty As Double
        On Error GoTo ErrHandler
        StrItemCode = ""
        For i = 0 To LstItems.Items.Count - 1
            If LstItems.GetItemChecked(i) = True Then
                If Len(Trim(StrItemCode)) > 0 Then
                    StrItemCode = StrItemCode & " or (a.Cust_drgNO ='" & Trim(Mid(ObsoleteManagement.GetItemString(LstItems, i), 1, 30)) & "' and a.Item_Code = '" & Trim(Mid(ObsoleteManagement.GetItemString(LstItems, i), 36, 16)) & "')"
                Else
                    StrItemCode = " and ((a.Cust_drgNO ='" & Trim(Mid(ObsoleteManagement.GetItemString(LstItems, i), 1, 30)) & "' and a.Item_Code = '" & Trim(Mid(ObsoleteManagement.GetItemString(LstItems, i), 36, 16)) & "')"
                End If
            End If
        Next i
        If Len(StrItemCode.Trim) > 0 Then
            StrItemCode = StrItemCode & ")"
        End If

        ReturnDrgITemCode = StrItemCode
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Function
    End Function
    Private Function NoOfDaysinMonth(ByRef MonthNo As Short, ByRef YearNo As Short) As Short
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :   strMode
        'Return Value       :   NA
        'Function           :   to get no of days in month
        'Comments           :   NA
        'Creation Date      :   18/04/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Select Case MonthNo
            Case 1, 3, 5, 7, 8, 10, 12
                NoOfDaysinMonth = 31
            Case 2
                If YearNo Mod 4 = 0 And YearNo Mod 100 <> 0 Or YearNo Mod 400 = 0 Then
                    NoOfDaysinMonth = 29
                Else
                    NoOfDaysinMonth = 28
                End If
            Case 4, 6, 9, 11
                NoOfDaysinMonth = 30
        End Select

        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Function
    End Function
    Private Sub RefreshFrm()
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :   strMode
        'Return Value       :   NA
        'Function           :   TO clear the fields
        'Comments           :   NA
        'Creation Date      :   18/04/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        TxtAccountCode.Text = ""
        lbldesc.Text = ""
        LstItems.Items.Clear()
        fpSDailySchedule.MaxRows = 0
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Function ReturnAccountDescription(ByRef strAccountCode As String) As String
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :   strMode
        'Return Value       :   NA
        'Function           :   to return the description of Account code
        'Comments           :   NA
        'Creation Date      :   18/04/2001
        '*******************************************************************************
        Dim strSql As String
        On Error GoTo ErrHandler
        strSql = "select Cust_Name from Customer_mst Where Unit_Code='" + gstrUNITID + "' And Customer_code='" & Trim(strAccountCode) & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
        If mRdoCls.GetResult(strSql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly) And mRdoCls.GetNoRows > 0 Then

            ReturnAccountDescription = mRdoCls.GetValue("Cust_Name")
        Else
            ReturnAccountDescription = ""
        End If
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub DisablePrimaryKeyControl()
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :   strMode
        'Return Value       :   NA
        'Function           :   to disable controls having primary key value
        'Comments           :   NA
        'Creation Date      :   18/04/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        TxtAccountCode.Enabled = False
        CmdHelpLocationCode.Enabled = False
        DTPTransDate.Enabled = False
        CmdSelectItems.Enabled = False
        Frame2.Enabled = False
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub EnablePrimaryKeyControl()
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :   strMode
        'Return Value       :   NA
        'Function           :   to enable controls having primary key value
        'Comments           :   NA
        'Creation Date      :   18/04/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        TxtAccountCode.Enabled = True
        CmdHelpLocationCode.Enabled = True
        DTPTransDate.Enabled = True
        CmdSelectItems.Enabled = True
        Frame2.Enabled = True
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub PreviousDataView(ByRef pstrAccountCode As String, ByRef pstrDate As String, ByRef pStrItemCodes As String, ByRef pblnDailySchedule As Boolean)
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :   strMode
        'Return Value       :   NA
        'Function           :   to bring previous data
        'Comments           :   NA
        'Creation Date      :   18/04/2001
        '*******************************************************************************
        'Add new Column (Hidden)Open SO For Accounts Plug in
        Dim strItemCodes As String 'to stroe mstrItemCodes value localy for interchange
        Dim strSql As String
        On Error GoTo ErrHandler
        If Trim(pstrAccountCode) <> "" And Trim(pstrDate) <> "" And Trim(pStrItemCodes) <> "" Then
            TxtAccountCode.Text = pstrAccountCode
            DTPTransDate.Value = pstrDate
            strItemCodes = mStrItemCodes
            If pblnDailySchedule = True Then
                OptDaily.Checked = True
            End If
            mStrItemCodes = strItemCodes
            If OptDaily.Checked = True Then
                fpSDailySchedule.MaxRows = 0
                strSql = "select trans_date= convert(varchar(20),a.trans_date,103) ,b.description ,a.Item_Code,Schedule_Quantity " & ",a.Schedule_Flag,a.cust_drgno   " & " from dailymktschedule a,item_mst b,custitem_mst c " & " where A.Unit_Code=B.Unit_Code And B.Unit_Code=C.Unit_Code And A.Unit_Code='" + gstrUNITID + "' And  a.item_code=b.item_code and a.item_code = c.item_code and a.cust_drgno = c.cust_drgno " & " and  a.Account_Code= '" & TxtAccountCode.Text & "' and  right(convert(varchar(20),a.trans_date,103),7)='" & VB6.Format(DTPTransDate.Value, "MM/yyyy") & "' and a.item_code in(" & pStrItemCodes & ") ORDER BY a.Item_Code,a.trans_date "
                If mRdoCls.GetResult(strSql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly) And mRdoCls.GetNoRows > 0 Then
                    mRdoCls.MoveFirst()
                    Do While Not mRdoCls.EOFRecord
                        fpSDailySchedule.MaxRows = fpSDailySchedule.MaxRows + 1
                        fpSDailySchedule.Row = fpSDailySchedule.MaxRows
                        fpSDailySchedule.Col = 3
                        fpSDailySchedule.Text = VB6.Format(mRdoCls.GetValue("trans_date"), gstrDateFormat)
                        fpSDailySchedule.Col = 4
                        fpSDailySchedule.Text = mRdoCls.GetValue("description")
                        fpSDailySchedule.Col = 5
                        fpSDailySchedule.Text = mRdoCls.GetValue("item_code")
                        fpSDailySchedule.Col = 6
                        fpSDailySchedule.Text = mRdoCls.GetValue("Schedule_Quantity")
                        fpSDailySchedule.Col = 7
                        fpSDailySchedule.Text = mRdoCls.GetValue("Schedule_flag")
                        fpSDailySchedule.Col = 8
                        fpSDailySchedule.Text = mRdoCls.GetValue("cust_drgno")
                        If Not mRdoCls.EOFRecord Then
                            mRdoCls.MoveNext()
                        End If
                    Loop
                    fpSDailySchedule.BlockMode = True
                    fpSDailySchedule.Col = 2
                    fpSDailySchedule.Col2 = 10
                    fpSDailySchedule.Row = 1
                    fpSDailySchedule.Row2 = fpSDailySchedule.MaxRows
                    fpSDailySchedule.Lock = True
                    fpSDailySchedule.BlockMode = False
                    Call ChangetheColorofGridInViewMode()
                End If
            End If
            lbldesc.Text = ReturnAccountDescription((TxtAccountCode.Text))
        End If
        '*******
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub

    End Sub
    Private Sub AssignValidFinancialYearDate()
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :   strMode
        'Return Value       :   NA
        'Function           :   To get Financil year dates from Company master
        'Comments           :   NA
        'Creation Date      :   18/04/2001
        '*******************************************************************************
        Dim strSql As String
        On Error GoTo ErrHandler
        strSql = "select financial_startdate=convert(char(10),financial_startdate,103),financial_enddate=convert(char(10),financial_enddate,103) from company_mst Where Unit_Code='" + gstrUNITID + "'"
        If mRdoCls.GetResult(strSql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly) And mRdoCls.GetNoRows > 0 Then
            mRdoCls.MoveFirst()
            mStrFinancialYearStartDate = ConvertToDate(VB6.Format(mRdoCls.GetValue("financial_startdate"), gstrDateFormat))
            mStrFinancialYearEndDate = ConvertToDate(VB6.Format(mRdoCls.GetValue("financial_enddate"), gstrDateFormat))

        Else
            Call ConfirmWindow(10083, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO, 100)
            Me.Close()
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub

    End Sub
    Private Sub ChangetheColorofGridInViewMode()
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :   strMode
        'Return Value       :   NA
        'Function           :   To change to fore colour of data coming in holidays
        'Comments           :   NA
        'Creation Date      :   18/04/2001
        '*******************************************************************************
        Dim i As Short
        Dim intTotalItem As Short
        Dim intNoofDaysinMonth As Short
        Dim intItemNo As Short
        Dim intItemNo1 As Short
        Dim intItemRange1 As Short
        Dim intItemRange2 As Short
        Dim RsCalendar_Mst As New ADODB.Recordset
        Dim intRowNo As Short
        Dim intRowvalue As String
        On Error GoTo ErrHandler
        'Add new Column (Hidden)Open SO For Accounts Plug in
        intNoofDaysinMonth = NoOfDaysinMonth(CShort(VB.Left(VB6.Format(DTPTransDate.Value, "MM/yyyy"), 2)), CShort(VB.Right(VB6.Format(DTPTransDate.Value, "MM/yyyy"), 4)))
        intTotalItem = Fix(fpSDailySchedule.MaxRows / intNoofDaysinMonth)
        intItemNo = Fix(fpSDailySchedule.Row / intNoofDaysinMonth)
        intItemNo1 = fpSDailySchedule.Row Mod intNoofDaysinMonth
        intItemRange1 = intItemNo
        If intItemNo1 = 0 Then
            intItemRange1 = intItemRange1 - 1
        End If
        For i = 1 To intTotalItem
            RsCalendar_Mst = New ADODB.Recordset
            If IsReference(RsCalendar_Mst) Then
                If IsNothing(RsCalendar_Mst) = False Then
                    If RsCalendar_Mst.State = ADODB.ObjectStateEnum.adStateOpen Then
                        RsCalendar_Mst.Close()
                    End If
                End If
            End If
            RsCalendar_Mst = New ADODB.Recordset
            RsCalendar_Mst.Open("SELECT CONVERT(VARCHAR(20),DT,103) FROM CALENDAR_MST " & " Where Unit_Code='" + gstrUNITID + "' And right(CONVERT(varCHAR(20),DT,103),7)='" & VB6.Format(DTPTransDate.Value, "MM/yyyy") & "' and work_flg=1 ", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If RsCalendar_Mst.RecordCount > 0 Then
                RsCalendar_Mst.MoveFirst()
                Do While Not RsCalendar_Mst.EOF
                    intRowvalue = VB.Left(RsCalendar_Mst.Fields(0).Value, 10)
                    For intRowNo = 1 To fpSDailySchedule.MaxRows
                        fpSDailySchedule.Row = intRowNo
                        fpSDailySchedule.Col = 3
                        If fpSDailySchedule.Text = intRowvalue Then
                            fpSDailySchedule.ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF)
                        End If
                    Next
                    If Not RsCalendar_Mst.EOF Then
                        RsCalendar_Mst.MoveNext()
                    End If
                Loop
                RsCalendar_Mst.Close()
                RsCalendar_Mst = Nothing
            End If
        Next i
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub

    End Sub
    Private Sub CheckActiveFlagForItems()
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :   strMode
        'Return Value       :   NA
        'Function           :   To check the active flag of items to make enable and
        '                       disable for editing
        'Comments           :   NA
        'Creation Date      :   18/04/2001
        '*******************************************************************************
        Dim RsItemActiveFlag As New ADODB.Recordset
        Dim i As Short
        Dim strSql As String
        Dim strDummy As String
        Dim StrItemCode As String
        Dim strCustDrgNo As String
        On Error GoTo ErrHandler
        'Add new Column (Hidden)Open SO For Accounts Plug in
        'to check whether that item is active or not
        For i = 1 To fpSDailySchedule.MaxRows
            fpSDailySchedule.Row = i
            fpSDailySchedule.Col = 5
            StrItemCode = Nothing
            strCustDrgNo = Nothing
            strDummy = CStr(fpSDailySchedule.GetText(5, i, StrItemCode))
            strDummy = CStr(fpSDailySchedule.GetText(8, i, strCustDrgNo))
            strSql = "SELECT active_flag FROM cust_ord_dtl " & " Where Unit_Code='" + gstrUNITID + "' And Account_Code= '" & TxtAccountCode.Text & "' and item_code='" & StrItemCode & "' and cust_drgno ='" & strCustDrgNo & "'"
            If mRdoCls.GetResult(strSql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly) And mRdoCls.GetNoRows > 0 Then
                mRdoCls.MoveFirst()
                If mRdoCls.GetValue("active_flag") <> "A" Then
                    fpSDailySchedule.BlockMode = True
                    fpSDailySchedule.Col = 2
                    fpSDailySchedule.Col2 = 7
                    fpSDailySchedule.Row = i
                    fpSDailySchedule.Row2 = i
                    fpSDailySchedule.Lock = True
                    fpSDailySchedule.BlockMode = False
                End If

            End If
        Next
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub UpdateDailyMonthlyMktSchedulestatus()
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :   strMode
        'Return Value       :   NA
        'Function           :   To update tables dailymktschedule and monthlymktschedule
        '                       if date is past or Ordered Quantity=Dispatch Quantity
        'Comments           :   NA
        'Creation Date      :   18/04/2001
        '*******************************************************************************
        Dim strSql As String
        On Error GoTo ErrHandler
        'to assign current date
        mStrCurrentDate = GetServerDate()
        strSql = "update dailymktschedule set schedule_flag='0',filetype = 'Manual' where Unit_Code='" + gstrUNITID + "' And convert(char(12),trans_date,103) <'" & mStrCurrentDate & "'"
        mP_Connection.Execute("set dateformat 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        mP_Connection.Execute(strSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        strSql = "update  monthlymktschedule set schedule_flag='0' where Unit_Code='" + gstrUNITID + "' And year_month<" & YearMonth() & " "
        mP_Connection.Execute(strSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub AssignCurrentMonthToDTPTRnasDate()
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :   strMode
        'Return Value       :   NA
        'Function           :   to assign current month to DTPTrnasDate
        'Comments           :   NA
        'Creation Date      :   18/04/2001
        '*******************************************************************************
        Dim strSql As String
        On Error GoTo ErrHandler
        DTPTransDate.Value = GetServerDate()
        mStrCurrentDate = GetServerDate()
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub MakeEditableAndEditableLinesofGrid()
        '*******************************************************************************
        'Author             :   Nisha Rai
        'Argument(s)if any  :   strMode
        'Return Value       :   NA
        'Function           :   According to the status of items make editable and non
        '                       editable
        'Comments           :   NA
        'Creation Date      :   18/04/2001
        '*******************************************************************************
        Dim i As Short
        Dim varDummy As Object
        Dim strsql As String = ""
        Dim varStatus As Object
        On Error GoTo ErrHandler
        'Add new Column (Hidden)Open SO For Accounts Plug in
        If OptDaily.Checked = True Then
            For i = 1 To fpSDailySchedule.MaxRows
                
                fpSDailySchedule.Row = i : fpSDailySchedule.Row2 = i : fpSDailySchedule.Col = 7 : fpSDailySchedule.Col2 = 7
                If CBool(fpSDailySchedule.Value) = False Then
                    fpSDailySchedule.BlockMode = True
                    fpSDailySchedule.Col = -1
                    fpSDailySchedule.Row = i
                    fpSDailySchedule.Row2 = i
                    fpSDailySchedule.Lock = True
                    fpSDailySchedule.BlockMode = False
                Else
                    'Added By priti on 13 Mar 2025 to add flag wise validation of drg No end date for Hilex
                    If (gstrUNITID = "H01" Or gstrUNITID = "H02" Or gstrUNITID = "H03" Or gstrUNITID = "H04") Then
                        Dim blnALLOW_CUSTITEM_START_ENDDATE As Boolean = SqlConnectionclass.ExecuteScalar("Select isnull(ALLOW_CUSTITEM_START_ENDDATE,0) from sales_parameter (Nolock) where unit_code='" + gstrUNITID + "'")
                        If blnALLOW_CUSTITEM_START_ENDDATE = True Then
                            Dim strItemCode As String = ""
                            Dim strTransDate As String = ""
                            Dim strCustDrgNo As String = ""
                            With fpSDailySchedule
                                fpSDailySchedule.Row = i
                                .Col = 3
                                strTransDate = .Text
                                .Col = 5
                                strItemCode = .Text
                                .Col = 8
                                strCustDrgNo = .Text
                            End With


                            Dim intValid As Integer = SqlConnectionclass.ExecuteScalar("Select COUNT(*) from CustItem_Mst (Nolock) where UNIT_CODE='" & gstrUNITID & "' AND Account_Code='" & TxtAccountCode.Text & "' and item_code='" & strItemCode & "' and cust_drgno='" & strCustDrgNo & "' and '" & getDateForDB(strTransDate) & "' between convert(date, Product_Start_date,103) and convert(date, Product_End_date,103")
                            If intValid = 1 Then
                                fpSDailySchedule.BlockMode = True
                                fpSDailySchedule.Col = 6
                                fpSDailySchedule.Col2 = 7
                                fpSDailySchedule.Row = i
                                fpSDailySchedule.Row2 = i
                                fpSDailySchedule.Lock = False
                                fpSDailySchedule.BlockMode = False


                            Else
                                fpSDailySchedule.BlockMode = True
                                fpSDailySchedule.Col = -1
                                fpSDailySchedule.Row = i
                                fpSDailySchedule.Row2 = i
                                fpSDailySchedule.Lock = True
                                fpSDailySchedule.BlockMode = False

                                fpSDailySchedule.Col = 2
                                fpSDailySchedule.ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF)
                                fpSDailySchedule.Col = 3
                                fpSDailySchedule.ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF)
                                fpSDailySchedule.Col = 4
                                fpSDailySchedule.ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF)
                                fpSDailySchedule.Col = 5
                                fpSDailySchedule.ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF)
                                fpSDailySchedule.Col = 6
                                fpSDailySchedule.ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF)
                                fpSDailySchedule.Col = 7
                                fpSDailySchedule.ForeColor = System.Drawing.ColorTranslator.FromOle(&HFF)
                                fpSDailySchedule.Text = ""
                            End If

                        End If
                        'end By priti on 17 Jan 2025 to add validation of drg No end date
                    Else
                        fpSDailySchedule.BlockMode = True
                        fpSDailySchedule.Col = 6
                        fpSDailySchedule.Col2 = 7
                        fpSDailySchedule.Row = i
                        fpSDailySchedule.Row2 = i
                        fpSDailySchedule.Lock = False
                        fpSDailySchedule.BlockMode = False
                    End If
                   
                End If
            Next i
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Public Function YearMonth() As Integer
        '*******************************************************************************
        'Author             :   Rajesh Sharma
        'Argument(s)if any  :   strMode
        'Return Value       :   NA
        'Function           :   Get the Current YearMonth
        'Comments           :   NA
        'Creation Date      :   18/04/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        'Declarations
        Dim objYearMonth As ClsResultSetDB 'Class Object
        Dim strSql As String 'Stores the SQL statement
        'Build the SQL statement
        strSql = "SELECT datepart(year,getdate()),datepart(month,getdate())"
        'Creating the instance
        objYearMonth = New ClsResultSetDB
        With objYearMonth
            'Open the recordset
            Call .GetResult(strSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            'If we have a record, then getting the financial year else exiting
            If .GetNoRows <= 0 Then Exit Function
            'Getting the date
            YearMonth = ((100 * .GetValueByNo(0)) + .GetValueByNo(1))
            'Closing the recordset
            .ResultSetClose()
        End With
        objYearMonth = Nothing
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Function
    End Function
    Public Function ValidRecord() As Boolean
        '*******************************************************************************
        'Author             :   Rajesh Sharma
        'Argument(s)if any  :   strMode
        'Return Value       :   NA
        'Function           :   check validity of data Before INSERT/UPDATE
        'Comments           :   NA
        'Creation Date      :   18/04/2001
        '*******************************************************************************
        Dim strErrMsg As String
        Dim ctlBlank As System.Windows.Forms.Control
        Dim lNo As Integer
        On Error GoTo Err_Handler
        blnInvalidData = False
        ValidRecord = False
        lNo = 1
        strErrMsg = ResolveResString(10059) & vbCrLf & vbCrLf
        If Len(Trim(TxtAccountCode.Text)) = 0 Then
            blnInvalidData = True
            strErrMsg = strErrMsg & vbCrLf & lNo & "." & "Account Code"
            lNo = lNo + 1
            If ctlBlank Is Nothing Then ctlBlank = TxtAccountCode
        End If
        If fpSDailySchedule.MaxRows = 0 Then
            blnInvalidData = True
            If CmdSelectItems.Enabled = False Then CmdSelectItems.Enabled = True
            strErrMsg = strErrMsg & vbCrLf & lNo & "." & "No Item Selected "
            lNo = lNo + 1
            If ctlBlank Is Nothing Then ctlBlank = CmdSelectItems
        End If
        strErrMsg = VB.Left(strErrMsg, Len(strErrMsg) - 1)
        strErrMsg = strErrMsg & " ."
        lNo = lNo + 1
        If blnInvalidData = True Then
            gblnCancelUnload = True
            Call MsgBox(strErrMsg, MsgBoxStyle.Information, "Error")
            ctlBlank.Focus()
            Exit Function
        End If
        ValidRecord = True
        gblnCancelUnload = True
        gblnFormAddEdit = True
        Exit Function
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Public Function CheckScheduleQuantity(ByRef pstrMode As Object) As Boolean
        '*******************************************************************************
        'Author             :   Rajesh Sharma
        'Argument(s)if any  :   strMode
        'Return Value       :   NA
        'Function           :   checks Total Schedule Quantity Againest Dispatch Quantity
        'Comments           :   NA
        'Creation Date      :   18/04/2001
        '*******************************************************************************
        Dim intRow As Short
        Dim varQuantity As Double
        Dim VarSchedule As Integer
        'Added for Issue ID eMpro-20090227-27987 Starts
        Dim VarDespatchQty As Double
        'Added for Issue ID eMpro-20090227-27987 Ends
        On Error GoTo Err_Handler
        CheckScheduleQuantity = False
        If fpSDailySchedule.Visible = True Then
            For intRow = 1 To fpSDailySchedule.MaxRows
                Call fpSDailySchedule.GetFloat(6, intRow, varQuantity)
                Call fpSDailySchedule.GetInteger(7, intRow, VarSchedule)
                Call fpSDailySchedule.GetInteger(10, intRow, VarDespatchQty)
                Select Case CmdGrpMktSchedule.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD

                        If Val(VarSchedule) = 1 Then

                            If System.Math.Abs(varQuantity) > 0 Then
                                CheckScheduleQuantity = True
                            End If
                        End If
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        ' If Val(VarSchedule) = 1 Then
                            If Val(VarDespatchQty) > 0 Then
                                If Val(varQuantity) > 0 And Val(varQuantity) >= Val(VarDespatchQty) Then
                                    CheckScheduleQuantity = True
                                Else
                                    CheckScheduleQuantity = False
                                    Exit For
                                End If
                            Else
                                CheckScheduleQuantity = True
                            End If

                        'End If
                End Select
            Next
        Else
        End If
        '******
        Exit Function
Err_Handler:
        'Resume Next
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Public Function ToCheckNoOfItemSelected(ByVal CurrentItemIndex As Short, ByVal NewValue As Boolean) As Short
        '*******************************************************************************
        'Author             :   Rajesh Sharma
        'Argument(s)if any  :   strMode
        'Return Value       :   NA
        'Function           :   check No of items selected in list
        'Comments           :   NA
        'Creation Date      :   18/04/2001
        '*******************************************************************************
        'Add on 17/06/2002
        On Error GoTo ErrHandler
        Dim intMaxCounter As Short
        Dim intLoopCounter As Short
        Dim IntItemSelected As Short
        intMaxCounter = LstItems.Items.Count
        If intMaxCounter > 0 Then
            For intLoopCounter = 0 To intMaxCounter - 1

                If LstItems.GetItemChecked(intLoopCounter) = True Then
                    IntItemSelected = IntItemSelected + 1
                End If
                If intLoopCounter = CurrentItemIndex Then
                    If NewValue = True Then
                        IntItemSelected = IntItemSelected + 1
                    Else
                        IntItemSelected = IntItemSelected - 1
                    End If
                End If
            Next
            ToCheckNoOfItemSelected = IntItemSelected
        End If
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Function
    End Function
    Public Function CheckForOpenSoFlag(ByRef pstrDrgno As String) As Boolean
        '*******************************************************************************
        'Author             :   Rajesh Sharma
        'Argument(s)if any  :   strMode
        'Return Value       :   NA
        'Function           :   Checks for open SO Flag
        'Comments           :   NA
        'Creation Date      :   18/04/2001
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim rsOpenSO As New ClsResultSetDB
        Dim strOpenSO As Object
        Dim blnOpenSO As Boolean
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        'To Check Open SO Flag in Cust_Ord_Dtl
        strOpenSO = "select  a.OpenSO,a.Item_Code,b.description,"
        strOpenSO = strOpenSO & "a.cust_drgno"
        strOpenSO = strOpenSO & " from cust_ord_dtl a,item_mst b, custitem_mst c "
        strOpenSO = strOpenSO & " where A.Unit_code=B.Unit_code And A.Unit_Code=C.Unit_Code And A.Unit_Code='" + gstrUNITID + "' And  a.item_code=b.item_code and a.cust_drgno = c.cust_drgno and  a.Account_Code= '" & TxtAccountCode.Text & "' and a.cust_drgno = '" & pstrDrgno & "' and authorized_flag=1 and a.active_flag='A'"
        strOpenSO = strOpenSO & " order By a.Account_Code,a.Item_Code,b.description,"
        strOpenSO = strOpenSO & " a.active_flag,a.cust_drgno,a.OpenSO "
        rsOpenSO.GetResult(strOpenSO)
        If rsOpenSO.GetNoRows > 0 Then
            intMaxLoop = rsOpenSO.GetNoRows : rsOpenSO.MoveFirst()
            blnOpenSO = False
            For intLoopCounter = 1 To intMaxLoop

                If rsOpenSO.GetValue("OpenSO") = True Then
                    blnOpenSO = True
                End If
                rsOpenSO.MoveNext()
            Next
        End If
        If blnOpenSO = True Then
            CheckForOpenSoFlag = True
        Else
            CheckForOpenSoFlag = False
        End If
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function validateQty(ByVal Row As Short, ByVal Col As Short) As Boolean
        '*******************************************************************************
        'Author             :   Davinder
        'Argument(s)if any  :   Row,Col of the grid
        'Return Value       :   Boolean
        'Function           :   To validate the Sch. Qty
        'Creation Date      :   28/03/2006
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim Rs As ClsResultSetDB
        Dim strQuery As String
        Rs = New ClsResultSetDB
        If Col = 6 Then
            fpSDailySchedule.Row = Row
            fpSDailySchedule.Col = 5
            strQuery = "select isnull(sum(Quantity),0) as Sch_Qty from mkt_enagaredtl Where Unit_Code='" + gstrUNITID + "' And Account_code='" & Trim(TxtAccountCode.Text) & "' AND Item_code='" & Trim(fpSDailySchedule.Text) & "' AND sch_date='"
            fpSDailySchedule.Col = 3
            strQuery = strQuery & VB6.Format(fpSDailySchedule.Text, "dd mmm yyyy") & "'"
            fpSDailySchedule.Col = 8
            strQuery = strQuery & " AND Cust_drgNo='" & Trim(fpSDailySchedule.Text) & "'"
            Rs.GetResult(strQuery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            fpSDailySchedule.Col = 6
            If Val(fpSDailySchedule.Text) < Val(Rs.GetValue("Sch_Qty")) Then
                MsgBox("Schedule Qty. can't be less than: " & Trim(Rs.GetValue("Sch_Qty")), MsgBoxStyle.Information + MsgBoxStyle.OkOnly, ResolveResString(100))
                validateQty = True
            End If
        End If
        Rs.ResultSetClose()
        Rs = Nothing
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function DecimalAllowed(ByVal Row As Short, ByVal Col As Short) As Boolean
        '*******************************************************************************
        'Author             :   Davinder
        'Argument(s)if any  :   Row,Col of the grid
        'Return Value       :   Boolean
        'Function           :   To validate the Sch. Qty
        'Creation Date      :   28/03/2006
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim Rs As ClsResultSetDB
        Dim strQuery As String
        Rs = New ClsResultSetDB
        If Col = 6 Then
            With fpSDailySchedule
                .Row = Row
                .Col = 5
                strQuery = "select I.Cons_measure_code as CMD, M.Decimal_Allowed_Flag as DAF,M.NoOFDecimal as NOD"
                strQuery = strQuery & " from item_mst I, Measure_Mst M"
                strQuery = strQuery & " Where i.unit_code=m.unit_code and i.unit_code='" + gstrUNITID + "' and i.Cons_measure_code = M.Measure_Code AND"
                strQuery = strQuery & " I.Item_Code='" & Trim(.Text) & "'"
                Rs.GetResult(strQuery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)

                If Not LCase(Rs.GetValue("DAF")) = "true" Then
                    .Col = 6
                    If Fix(Convert.ToDouble(Val(.Text))) < Val(.Text) Then
                        .Col = 5
                        MsgBox("Quantity can not be in decimal places for Item: " & Trim(.Text))
                        DecimalAllowed = True
                    End If
                Else
                    DecimalAllowed = False
                End If
            End With
        End If
        Rs.ResultSetClose()
        Rs = Nothing
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Function GetNagareQty(ByVal Row As Short, ByVal Col As Short) As Double
        '*******************************************************************************
        'Author             :   Davinder
        'Argument(s)if any  :   Row,Col of the grid
        'Return Value       :   Boolean
        'Function           :   To validate the Sch. Qty
        'Creation Date      :   28/03/2006
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim Rs As ClsResultSetDB
        Dim strQuery As String
        Rs = New ClsResultSetDB
        If Col = 6 Then
            fpSDailySchedule.Row = Row
            fpSDailySchedule.Col = 5
            strQuery = "select isnull(sum(Quantity),0) as Sch_Qty from mkt_enagaredtl Where Unit_Code='" + gstrUNITID + "' And Account_code='" & Trim(TxtAccountCode.Text) & "' AND Item_code='" & Trim(fpSDailySchedule.Text) & "' AND sch_date='"
            fpSDailySchedule.Col = 3
            strQuery = strQuery & VB6.Format(fpSDailySchedule.Text, "dd mmm yyyy") & "'"
            fpSDailySchedule.Col = 8
            strQuery = strQuery & " AND Cust_drgNo='" & Trim(fpSDailySchedule.Text) & "'"
            Rs.GetResult(strQuery, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            fpSDailySchedule.Col = 6

            GetNagareQty = Val(Rs.GetValue("Sch_Qty"))
        End If
        Rs.ResultSetClose()
        Rs = Nothing
        Exit Function
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

        rsEXcessSchCustomer.GetResult("Select isnull(AllowExcessSchedule,1) as AllowExcessSchedule from customer_mst Where Unit_Code='" + gstrUNITID + "'  And Customer_code='" & Trim(TxtAccountCode.Text) & "'")
        If rsEXcessSchCustomer.GetNoRows > 0 Then

            If rsEXcessSchCustomer.GetValue("AllowExcessSchedule") = False Then
                AllowEditSchedule = False
            Else
                AllowEditSchedule = True
            End If
        End If
        rsEXcessSchCustomer.ResultSetClose()
        Exit Function
ErrHandler:
        AllowEditSchedule = True
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
End Class