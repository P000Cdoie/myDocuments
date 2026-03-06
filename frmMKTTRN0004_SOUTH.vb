Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmMKTTRN0004_SOUTH
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
	'----------------------------------------------------
	'Revised By Arul on 21-04-2005
	'Revised for Check the monthly Schduled Qty equal to 0 for daily schdule generation
    '----------------------------------------------------------------------
    'Revised By Manoj on 01-DEC-2008 for Issue ID eMpro-20081201-24140
    'Revised for Date Conversion Error was coming
    '----------------------------------------------------------------------
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
    ' Change By Deepak on 11-Oct-2011 for Change Management-----------
    '-----------------------------------------------------------------------------
    'REVISED By         : SHUBHRA VERMA 
    'REVISED ON         : 30-MAY-2012
    'ISSUE ID           : 10230966
    'DESCRIPTION        : Problem in editing in Daily Marketing Schedule
    '-----------------------------------------------------------------------------
    'REVISED By         : PRASHANT RAJPAL
    'REVISED ON         : 13-AUG-2012
    'ISSUE ID           : 10262463
    'DESCRIPTION        : STILL Problem in editing in Daily Marketing Schedule
    '-----------------------------------------------------------------------------
    'Revised By       - Neha Ghai
    'Revision Date    - 19 sep 2012
    'Issue Id         - 10277185 
    'Description      - Search Option not working.
    '---------------------------------------------------------------------------
    'Revised By       - Prashant rajpal
    'Revision Date    - 02 feb 2014
    'Issue Id         - 10488279
    'Description      - RAW Material and BOP is incorporated in this form
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
    Private Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As String) As Integer
    Private Declare Function SendMessageBynum Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
	Private Const LB_GETHORIZONTALEXTENT As Short = &H193s
	Private Const LB_SETHORIZONTALEXTENT As Short = &H194s
	'this API is used to do the searching in listbox
    'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Short, ByVal lParam As Object) As Integer
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
    Private Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As POINTAPI) As Integer
    Dim mRdoCls As New ClsResultSetDB
    Dim blnInvalidData As Boolean
    Dim blnCheckallClicked As Boolean
    Private Sub CmdGrpMktSchedule_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.UCbtndgrp.ButtonClickEventArgs) Handles CmdGrpMktSchedule.ButtonClick
        On Error GoTo ErrHandler
        Dim intRow As Short
        Dim varStatus As Object
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
                            If CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then 'to add the record
                                mP_Connection.BeginTrans()
                                Call SaveAddDaily("A") 'Call procedure to add record
                                mP_Connection.CommitTrans()
                                Call ConfirmWindow(10049, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                                CmdGrpMktSchedule.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = True
                                CmdGrpMktSchedule.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                                'nisha
                                gblnCancelUnload = False : gblnFormAddEdit = False
                            ElseIf CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then  'If Edit mode
                                mP_Connection.BeginTrans()
                                Call SaveUpdateDaily() 'Call Procedure to edit record
                                mP_Connection.CommitTrans()
                                Call ConfirmWindow(10049, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK)
                                With fpSDailySchedule
                                    .Row = 1 : .Row2 = .MaxRows : .Col = 2 : .Col2 = .MaxCols
                                    .BlockMode = True : .Lock = True : .BlockMode = False
                                End With
                                CmdGrpMktSchedule.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = True
                                CmdGrpMktSchedule.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                                'nisha
                                gblnCancelUnload = False : gblnFormAddEdit = False
                            End If
                        Else 'if monthly mkt schedule is selected
                        End If
                    Else
                        MsgBox("Schedule Quantity can't be less than Dispatch Qty OR Zero Where Status is Active.", MsgBoxStyle.OkOnly, ResolveResString(100))
                        'Add new Column (Hidden)Open SO For Accounts Plug in
                        If fpSDailySchedule.Enabled = True Then
                            For intRow = 1 To fpSDailySchedule.MaxRows
                                varStatus = Nothing
                                Call fpSDailySchedule.GetText(7, intRow, varStatus)
                                If varStatus = "" Then
                                    varStatus = 0
                                End If
                                If varStatus = 1 Then
                                    fpSDailySchedule.Col = 6
                                    fpSDailySchedule.Row = intRow
                                    fpSDailySchedule.Action = FPSpreadADO.ActionConstants.ActionActiveCell
                                    Exit Sub
                                End If
                            Next
                        Else
                        End If
                        '****
                        Exit Sub
                    End If
                Else
                    'nisha
                    gblnCancelUnload = True : gblnFormAddEdit = True
                    Exit Sub
                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL ' Cancel Click
                Call frmMKTTRN0004_SOUTH_KeyPress(Me, New System.Windows.Forms.KeyPressEventArgs(Chr(System.Windows.Forms.Keys.Escape)))
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT 'Clich on edit
                '     All UpdateDailyMonthlyMktSchedulestatus ' to update daily and monthly mktschedule tables for past dates
                mStrAccountCode = TxtAccountCode.Text
                mStrDate = CDate(VB6.Format(DTPTransDate.Value, "MM/yyyy")) 'to retain existing value used in case of cancel
                If OptDaily.Checked = True Then
                    mblnDailySchedule = True 'to retain existing value used in case of cancel
                Else
                    mblnDailySchedule = False
                End If
                If OptDaily.Checked = True Then 'daily mkt schedule is selected
                    Call fillSpread() ' Fill spread according to values
                    Call DisablePrimaryKeyControl() 'to disable control
                End If
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE 'To Verify Close Operation
                If Me.CmdGrpMktSchedule.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                    enmValue = ConfirmWindow(10055, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNOCANCEL, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION)
                    If enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                        Call CmdGrpMktSchedule_ButtonClick(CmdGrpMktSchedule, New UCActXCtl.UCbtndgrp.ButtonClickEventArgs(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE))
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
        '******
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub CmdHelpConsCode_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdHelpConsCode.Click
        '*******************************************************************************
        'Author             :   Ashutosh Verma, Issue Id:19661
        'Argument(s)if any  :
        'Return Value       :   NA
        'Function           :   To Display Consignee code's Help
        'Comments           :   NA
        'Creation Date      :   19 Mar 2007
        '*******************************************************************************
        On Error GoTo ErrHandler
        Dim strHelpString As String
        If Len(Trim(txtConsCode.Text)) = 0 Then
            strHelpString = ShowList(1, (txtConsCode.MaxLength), "", "Customer_Code", "Cust_Name", "Customer_Mst", " and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))", "Consignee Code Help")
            If strHelpString = "-1" Then 'If No Record Found
                MsgBox("Invalid Consigne code.", MsgBoxStyle.Information, "eMpower")
            Else
                txtConsCode.Text = strHelpString
            End If
        Else
            strHelpString = ShowList(1, (txtConsCode.MaxLength), txtConsCode.Text, "customer_code", "cust_name", "Customer_Mst", " and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))", "Consignee Code Help")
            If strHelpString = "-1" Then 'If No Record Found
                MsgBox("Invalid Consigne code.", MsgBoxStyle.Information, "eMpower")
            Else
                txtConsCode.Text = strHelpString
            End If
        End If
        txtConsCode.Focus()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub CmdHelpLocationCode_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdHelpLocationCode.Click
        Dim strAccountCode As String
        'to provide help of valid account codes and their description
        On Error GoTo ErrHandler
        lbldesc.Text = ""
        With TxtAccountCode
            If Len(.Text) = 0 Then
                'Samiksha SMRC start

                If gstrUNITID = "STH" Then
                    strAccountCode = ShowList(1, .MaxLength, "", "A.Customer_code", "B.Cust_Name", "CUSTOMER_CONSIGNEE_MAPPING A,Customer_mst B", " and A.CUSTOMER_CODE = B.Customer_Code AND A.UNIT_CODE=B.UNIT_CODE and ((isnull(B.deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= B.deactive_date))", , , , , , "A.UNIT_CODE")

                Else
                    strAccountCode = ShowList(1, .MaxLength, "", "Customer_code", "Cust_Name", "Customer_mst ", " and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))")
                End If

                'Samiksha SMRC End

                If strAccountCode <> "-1" Then
                    .Text = strAccountCode
                    lbldesc.Text = ReturnAccountDescription(.Text)
                Else
                    .Text = ""
                    lbldesc.Text = ""
                    .Focus()
                    Call ConfirmWindow(10080, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_CRITICAL, 100)
                End If
            Else
                'Samiksha SMRC start

                If gstrUNITID = "STH" Then
                    strAccountCode = ShowList(1, .MaxLength, .Text, "A.Customer_code", "B.Cust_Name", "CUSTOMER_CONSIGNEE_MAPPING A,Customer_mst B", " and A.CUSTOMER_CODE = B.Customer_Code AND A.UNIT_CODE=B.UNIT_CODE and ((isnull(B.deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= B.deactive_date))", , , , , , "A.UNIT_CODE")

                Else
                    strAccountCode = ShowList(1, .MaxLength, .Text, " Customer_code", "Cust_Name", " Customer_mst ", " and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))")

                End If

                'Samiksha SMRC end

                If strAccountCode <> "-1" Then
                    .Text = strAccountCode
                    lbldesc.Text = ReturnAccountDescription(.Text)
                Else
                    .Text = ""
                    lbldesc.Text = ""
                    .Focus()
                    Call ConfirmWindow(10080, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_CRITICAL, 100)
                End If
            End If
            If .Enabled = True Then
                .Focus()
            End If
        End With
        Call FillList()
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub CmdHelpLocationCode_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles CmdHelpLocationCode.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
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
        'to provide list of items
        On Error GoTo ErrHandler
        Me.chkCheckAll.Checked = False
        Me.chkUnCheckAll.Checked = False
        Call FillList() 'Fill the list for valid items
        Frame1.Enabled = False
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
        ' to disapear tabpage and fill grid according selected items
        On Error GoTo ErrHandler
        Dim strReturnItems As String
        Dim blnMsgFlag As Boolean
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
                If VB6.Format(DTPTransDate._Value, "yyyy/MM") < VB6.Format(GetServerDate, "yyyy/MM") Then
                    Me.CmdGrpMktSchedule.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                Else
                    Me.CmdGrpMktSchedule.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = True
                End If
            End If
        Else
            fpSDailySchedule.MaxRows = 0
            If LstItems.Items.Count > 0 Then
                MsgBox("Please Select Atleast one item", MsgBoxStyle.OkOnly, "empower")
                blnMsgFlag = True
            Else
                blnMsgFlag = False
            End If
            If blnMsgFlag = False Then
                FraLstItems.Visible = False 'make invisible the frame containing items list
            End If
            CmdSelectItems.Focus()
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub CmdSelectItems_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles CmdSelectItems.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
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
    Private Sub DTPTransDate_Change(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles DTPTransDate.Change
        'to check valid schedule date
        On Error GoTo ErrHandler
        If fpSDailySchedule.Visible = True Then
            fpSDailySchedule.MaxRows = 0
        End If
        Call FillList()
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub DTPTransDate_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSComCtl2.DDTPickerEvents_KeyDownEvent) Handles DTPTransDate.KeyDownEvent
        On Error GoTo ErrHandler
        If eventArgs.keyCode = 13 Then
            If CmdSelectItems.Enabled = True Then
                CmdSelectItems.Focus()
            End If
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0004_SOUTH_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo ErrHandler
        DTPTransDate.Format = 3 'to set custom format to DTPInvoiceDate Control
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
    Private Sub frmMKTTRN0004_SOUTH_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        frmModules.NodeFontBold(Tag) = False
    End Sub
    Private Sub frmMKTTRN0004_SOUTH_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F4 And Shift = 0 Then Call ctlFormHeader1_Click(ctlFormHeader1, New System.EventArgs()) : Exit Sub
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0004_SOUTH_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
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
                        Me.CmdGrpMktSchedule.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                        Me.CmdGrpMktSchedule.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                        Frame2.Enabled = True
                        Me.OptDaily.Enabled = True
                        Me.DTPTransDate.Enabled = True
                        Me.OptDaily.Focus()
                        gblnCancelUnload = False : gblnFormAddEdit = False
                        txtConsCode.Enabled = True
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
    Private Sub frmMKTTRN0004_SOUTH_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
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
        'to put scrollbar in List Box
        Call SendMessageBynum(LstItems.Handle.ToInt32, LB_SETHORIZONTALEXTENT, 900, 0)
        FraLstItems.Left = VB6.TwipsToPixelsX(2500)
        'to fill the lables from resource file
        Call FillLabelFromResFile(Me)
        '--------------------------------------------------------------------------
        Me.CmdGrpMktSchedule.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
        Me.CmdGrpMktSchedule.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL) = False 'to make Print button false
        Me.CmdGrpMktSchedule.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False 'to make cancle button false
        Me.CmdGrpMktSchedule.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False ' to disable delete button
        mStrItemCodes = "''" 'To initialize this variable
        chkCheckAll.CheckState = System.Windows.Forms.CheckState.Unchecked
        OptDaily.Checked = True
        Exit Sub 'To avoid the execution of error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub frmMKTTRN0004_SOUTH_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
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
    Private Sub frmMKTTRN0004_SOUTH_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error GoTo ErrHandler
        mdifrmMain.RemoveFormNameFromWindowList = mintFormIndex
        frmModules.NodeFontBold(Tag) = False
        Me.Dispose()
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub FraLstItems_MouseMove(ByRef Button As Short, ByRef Shift As Short, ByRef x As Single, ByRef y As Single)
        On Error GoTo ErrHandler
        Dim upperleft As POINTAPI
        Dim lngMouseMove As Integer
        If Button = 1 Then
            lngMouseMove = GetCursorPos(upperleft)
            FraLstItems.Left = VB6.TwipsToPixelsX(upperleft.x * 7)
            FraLstItems.Top = VB6.TwipsToPixelsY(upperleft.y * 7)
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub txtConsCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtConsCode.TextChanged
        On Error GoTo ErrHandler
        If Len(Trim(txtConsCode.Text)) = 0 Then
            Exit Sub
        Else
            Me.lblConsDesc.Text = ""
        End If
        Exit Sub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
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
        GoTo EventExitSub
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
        If KeyCode = 112 Then Call CmdHelpConsCode_Click(CmdHelpConsCode, New System.EventArgs()) 'Help should be invoked if F1 key is pressed
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
        On Error GoTo ErrHandler
        Dim rsCD As New ClsResultSetDB
        Dim strsql As String
        If Len(Trim(txtConsCode.Text)) = 0 Then
            GoTo EventExitSub
        Else
            strsql = "Select Cust_Name from Customer_mst Where Unit_Code='" + gstrUNITID + "' And customer_Code='" & Trim(txtConsCode.Text) & "' and ((isnull(deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= deactive_date))"
            rsCD.GetResult(strsql)
            If rsCD.GetNoRows = 0 Then
                MsgBox("Invalid Consignee Code !!!", MsgBoxStyle.Information, "eMpower")
                txtConsCode.Text = ""
                lblConsDesc.Text = ""
                Cancel = True
                txtConsCode.Focus()
                GoTo EventExitSub
            Else
                lblConsDesc.Text = IIf(UCase(rsCD.GetValue("Cust_Name")) = "UNKNOWN", "", rsCD.GetValue("Cust_Name"))
            End If
        End If
        GoTo EventExitSub 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        GoTo EventExitSub
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub LstItems_ItemCheck(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.ItemCheckEventArgs) Handles LstItems.ItemCheck
        On Error GoTo ErrHandler
        Dim intItemChecked As Short
        intItemChecked = ToCheckNoOfItemSelected()
        If intItemChecked = LstItems.Items.Count Then
            chkCheckAll.CheckState = System.Windows.Forms.CheckState.Checked : chkUnCheckAll.CheckState = System.Windows.Forms.CheckState.Unchecked
        ElseIf intItemChecked = 0 Then
            chkUnCheckAll.CheckState = System.Windows.Forms.CheckState.Checked : chkCheckAll.CheckState = System.Windows.Forms.CheckState.Unchecked
        ElseIf (intItemChecked > 0) And (intItemChecked <> LstItems.Items.Count) Then
            chkCheckAll.CheckState = System.Windows.Forms.CheckState.Unchecked : chkUnCheckAll.CheckState = System.Windows.Forms.CheckState.Unchecked
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub LstItems_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles LstItems.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        'to select and deselect items in listbox
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
        On Error GoTo ErrHandler
        If Len(Trim(TxtAccountCode.Text)) = 0 Then
            Me.lbldesc.Text = " "
        End If
        If fpSDailySchedule.Visible = True Then
            fpSDailySchedule.MaxRows = 0
        End If
        CmdGrpMktSchedule.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
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
        'to provide help when user will press F1
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
        'to check the vaild account code
        Dim strsql As String
        On Error GoTo ErrHandler
        If KeyAscii = 13 Then ' when user will press Enter key
            If Len(Me.TxtAccountCode.Text) > 0 Then
                lbldesc.Text = ""
                strsql = "SELECT DISTINCT A.ACCOUNT_CODE,B.Cust_Name FROM CUST_ORD_DTL A," & " Customer_MST B  Where A.UNIT_CODE=B.UNIT_CODE AND  A.Unit_Code='" + gstrUNITID + "' AND A.ACCOUNT_CODE = B.Customer_CODE  " & "  and a.account_code='" & Replace(TxtAccountCode.Text, "'", "") & "' and ((isnull(B.deactive_flag,0) <> 1) OR (CAST(getdate() AS DATE) <= B.deactive_date))"
                If mRdoCls.GetResult(strsql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly) And mRdoCls.GetNoRows > 0 Then
                    mRdoCls.MoveFirst()
                    lbldesc.Text = ReturnAccountDescription((TxtAccountCode.Text))
                    If DTPTransDate.Enabled = True Then
                        DTPTransDate.Focus()
                    End If
                Else
                    Call ConfirmWindow(10080, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO, 100)
                    If TxtAccountCode.Enabled = True Then
                        TxtAccountCode.Focus()
                    End If
                End If
            Else
                Me.CmdGrpMktSchedule.Focus()
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
    Private Sub TxtSearch_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TxtSearch.TextChanged
        'to search items in list box

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
        'to select and de select items in list box
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
            Dim strItemCodes As String
            On Error GoTo ErrHandler
            If CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                strItemCodes = mStrItemCodes
            End If
            DTPTransDate.CustomFormat = "MM/yyyy"
            lblSchedule.Text = "Schedule Month"
            fpSDailySchedule.Visible = True
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
        On Error GoTo ErrHandler
        'To fill data in grid according to mode (view,edit,add)(daily,Monthly)
        Dim i As Short
        Dim strsql As String
        Dim IntRecNo As Short
        Dim intRowNo As Short
        Dim StrDiffitemCode As String
        Dim RsCalendar_Mst As New ADODB.Recordset
        Dim varDummy As Object
        Dim VarTransDate As Object
        Dim StrTransDate As String
        Dim blnOpenSO As Boolean
        Dim boolDummy As Boolean
        Dim rsDailyItem As New ClsResultSetDB
        Dim strDate As String
        Dim intMaxLoop As Short
        Dim intLoopCounter As Short
        'Add new Column (Hidden)Open SO For Accounts Plug in
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.OBJ_FORM, Me, System.Windows.Forms.Cursors.WaitCursor)
        fpSDailySchedule.MaxRows = 0
        If OptDaily.Checked = True Then
            fpSDailySchedule.Visible = True
            If Me.CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
                StrDiffitemCode = ReturnDrgITemCode()
                strsql = "select  a.Account_Code,a.Item_Code,b.description," & " authorized_flag=1," & " Order_Qty=sum(a.Order_Qty),Despatch_Qty=sum(a.Despatch_Qty),a.active_flag" & ",a.cust_drgno" & " from cust_ord_dtl a,item_mst b, custitem_mst c " & " Where A.UNIT_CODE=B.UNIT_CODE  AND A.Unit_Code='" + gstrUNITID + "' AND a.item_code=b.item_code and a.cust_drgno = c.cust_drgno AND a.UNIT_CODE=c.UNIT_CODE and  a.Account_Code= '" & TxtAccountCode.Text & "' and a.account_code = c.account_code and " & " authorized_flag=1 and a.active_flag='A'" & " and a.account_code = c.account_code "
                If Len(Trim(StrDiffitemCode)) > 0 Then
                    strsql = strsql & StrDiffitemCode
                End If
                strsql = strsql & " group By a.Account_Code,a.Item_Code,b.description," & " a.active_flag,a.cust_drgno " ' remove this filed from group by a.Authorised_Flag,
                If mRdoCls.GetResult(strsql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly) And mRdoCls.GetNoRows > 0 Then
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
                            strsql = "select Account_Code,Trans_date,Item_code,Cust_Drgno,Serial_No,Schedule_Flag,Schedule_Quantity,Despatch_Qty,Status,RevisionNo,DSNO,DSDateTime,Spare_qty,Consignee_Code,FILETYPE,DOC_NO,MSG_NAME from dailymktschedule "
                            strsql = strsql & " Where Unit_Code='" + gstrUNITID + "' AND  account_code='" & TxtAccountCode.Text & "' and consignee_code='" & Trim(Me.txtConsCode.Text) & "' "
                            strsql = strsql & " and datepart(mm,trans_date) = " & VB6.Format(DTPTransDate.Value, "MM")
                            strsql = strsql & " and datepart(yyyy,trans_date) = " & VB6.Format(DTPTransDate.Value, "YYYY")
                            strsql = strsql & " and datepart(dd,trans_date) = " & i
                            strsql = strsql & " and cust_drgNo = '" & mRdoCls.GetValue("cust_drgno") & "' and Item_code = '" & mRdoCls.GetValue("Item_code") & "' and status =1"
                            rsDailyItem.GetResult(strsql)
                            If rsDailyItem.GetNoRows = 0 Then
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
                                fpSDailySchedule.Text = VB6.Format(VB6.Format(i, "00") & "/" & VB6.Format(DTPTransDate.Value, "MM/yyyy"), gstrDateFormat)
                                fpSDailySchedule.Col = 4
                                fpSDailySchedule.Text = mRdoCls.GetValue("description")
                                fpSDailySchedule.Col = 5
                                fpSDailySchedule.Text = mRdoCls.GetValue("item_code")
                                fpSDailySchedule.Col = 6
                                fpSDailySchedule.CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                fpSDailySchedule.Text = CStr(0)
                                fpSDailySchedule.Col = 7
                                fpSDailySchedule.CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
                                fpSDailySchedule.Value = CStr(1) 'mRdoCls.GetValue("item_code")
                                VarTransDate = Nothing
                                boolDummy = fpSDailySchedule.GetText(3, fpSDailySchedule.Row, VarTransDate)
                                'StrTransDate = VB.Right(VarTransDate, 4) & "/" & Mid(VarTransDate, 4, 2) & "/" & VB.Left(VarTransDate, 2)
                                If ConvertToDate(VarTransDate) < mStrCurrentDate Then
                                    fpSDailySchedule.Text = ""
                                End If
                                fpSDailySchedule.Col = 8
                                fpSDailySchedule.Row = fpSDailySchedule.MaxRows : fpSDailySchedule.Row2 = fpSDailySchedule.MaxRows : fpSDailySchedule.Col = 8 : fpSDailySchedule.Col2 = 8 : fpSDailySchedule.BlockMode = True : fpSDailySchedule.Lock = True : fpSDailySchedule.BlockMode = False
                                fpSDailySchedule.Text = mRdoCls.GetValue("cust_drgno")
                                fpSDailySchedule.Col = 9
                                fpSDailySchedule.CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                fpSDailySchedule.Text = mRdoCls.GetValue("Order_qty")
                                fpSDailySchedule.Col = 10
                                fpSDailySchedule.CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                                'issue id 10108966
                                'fpSDailySchedule.Text = mRdoCls.GetValue("Despatch_qty")
                                fpSDailySchedule.Text = CStr(0)
                                'issue id 10108966 end 

                                fpSDailySchedule.Col = 11
                                fpSDailySchedule.Text = CStr(0)
                            End If
                        Next i
                        If IsReference(RsCalendar_Mst) Then
                            If RsCalendar_Mst.State = 1 Then
                                RsCalendar_Mst.Close()
                            End If
                        End If
                        RsCalendar_Mst.Open("SELECT CONVERT(VARCHAR(20),DT,103) as Dt  FROM CALENDAR_MST " & " Where Unit_Code='" + gstrUNITID + "' AND right(CONVERT(varCHAR(20),DT,103),7)='" & VB6.Format(DTPTransDate.Value, "MM/yyyy") & "' and work_flg=1 ", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
                        If RsCalendar_Mst.RecordCount > 0 Then
                            RsCalendar_Mst.MoveFirst()
                            fpSDailySchedule.CursorStyle = FPSpreadADO.CursorStyleConstants.CursorStyleDefault
                            Do While Not RsCalendar_Mst.EOF
                                strDate = VB6.Format(RsCalendar_Mst.Fields("Dt").Value, gstrDateFormat)
                                intRowNo = CShort(VB.Left(RsCalendar_Mst.Fields(0).Value, 2))
                                intMaxLoop = fpSDailySchedule.MaxRows
                                SetGridDateFormat(fpSDailySchedule, 3)
                                For intLoopCounter = 1 To intMaxLoop
                                    fpSDailySchedule.Col = 3 : fpSDailySchedule.Row = intLoopCounter
                                    If DateDiff(Microsoft.VisualBasic.DateInterval.Day, ConvertToDate(fpSDailySchedule.Text), ConvertToDate(strDate)) = 0 Then
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
                strsql = "select b.description ,a.Item_Code,a.Schedule_Quantity  " & ",a.Schedule_Flag,a.cust_drgno ,trans_date= convert(varchar(20),a.trans_date,103),a.despatch_qty,a.RevisionNo,Consignee_Code  " & " from dailymktschedule a,item_mst b,custitem_mst c where A.UNIT_CODE=B.UNIT_CODE  AND A.Unit_Code='" + gstrUNITID + "' AND a.item_code=b.item_code and a.item_code = c.item_code AND a.UNIT_CODE=c.UNIT_CODE and a.cust_drgno = c.cust_drgno " & " and a.status =1 and a.Account_Code= '" & TxtAccountCode.Text & "' and a.consignee_code='" & Trim(Me.txtConsCode.Text) & "'   and  a.account_code = c.account_code  and right(convert(varchar(20),a.trans_date,103),7)='" & VB6.Format(DTPTransDate.Value, "MM/yyyy") & "'"
                If Len(Trim(StrDiffitemCode)) > 0 Then
                    strsql = strsql & StrDiffitemCode
                End If
                strsql = strsql & " ORDER BY a.Item_Code,a.cust_drgno,trans_date"
                If mRdoCls.GetResult(strsql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly) And mRdoCls.GetNoRows > 0 Then
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
                        '************
                        fpSDailySchedule.Col = 3
                        fpSDailySchedule.Text = VB6.Format(mRdoCls.GetValue("trans_date"), gstrDateFormat)
                        fpSDailySchedule.Col = 4
                        fpSDailySchedule.Text = mRdoCls.GetValue("description")
                        fpSDailySchedule.Col = 5
                        fpSDailySchedule.Text = mRdoCls.GetValue("item_code")
                        fpSDailySchedule.Col = 6
                        fpSDailySchedule.CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                        fpSDailySchedule.Text = CStr(Val(mRdoCls.GetValue("Schedule_Quantity")))
                        fpSDailySchedule.Col = 12
                        fpSDailySchedule.CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                        fpSDailySchedule.Text = CStr(Val(mRdoCls.GetValue("Schedule_Quantity")))
                        fpSDailySchedule.Col = 7
                        If mRdoCls.GetValue("Schedule_flag") Then
                            fpSDailySchedule.CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
                            fpSDailySchedule.Value = "1"
                        Else
                            fpSDailySchedule.CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
                            fpSDailySchedule.Value = "0"
                        End If
                        fpSDailySchedule.Col = 13
                        If mRdoCls.GetValue("Schedule_flag") Then
                            fpSDailySchedule.CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
                            fpSDailySchedule.Value = "1"
                        Else
                            fpSDailySchedule.CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
                            fpSDailySchedule.Value = "0"
                        End If
                        fpSDailySchedule.Col = 8
                        fpSDailySchedule.Row = fpSDailySchedule.MaxRows : fpSDailySchedule.Row2 = fpSDailySchedule.MaxRows : fpSDailySchedule.Col = 8 : fpSDailySchedule.Col2 = 8 : fpSDailySchedule.BlockMode = True : fpSDailySchedule.Lock = True : fpSDailySchedule.BlockMode = False
                        '******
                        fpSDailySchedule.Text = mRdoCls.GetValue("cust_drgno")
                        fpSDailySchedule.Col = 9
                        fpSDailySchedule.CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                        fpSDailySchedule.Col = 10
                        fpSDailySchedule.CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                        fpSDailySchedule.Text = CStr(Val(mRdoCls.GetValue("despatch_qty")))
                        fpSDailySchedule.Col = 11
                        fpSDailySchedule.Text = CStr(Val(mRdoCls.GetValue("RevisionNo")))

                        'commented and added by shubhra against issue ID : 10230966
                        'If VB6.Format(mRdoCls.GetValue("trans_date"), "DD MMM YYYY") < VB6.Format(GetServerDate, "DD MMM YYYY") Then
                        If CDate(VB6.Format(mRdoCls.GetValue("trans_date"), "DD MMM YYYY")) < CDate(VB6.Format(GetServerDate, "DD MMM YYYY")) Then
                            fpSDailySchedule.BlockMode = True
                            fpSDailySchedule.Col = 2
                            fpSDailySchedule.Col2 = 10
                            '10262463
                            'fpSDailySchedule.Row = fpSDailySchedule.ActiveRow
                            fpSDailySchedule.Row = fpSDailySchedule.MaxRows
                            '10262463
                            fpSDailySchedule.Row2 = fpSDailySchedule.MaxRows
                            fpSDailySchedule.Lock = True
                            fpSDailySchedule.BlockMode = False
                        End If
                        If Not mRdoCls.EOFRecord Then
                            mRdoCls.MoveNext()
                        End If
                    Loop
                    Call ChangetheColorofGridInViewMode()
                End If
            ElseIf Me.CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                scheduleEdit = False
                StrDiffitemCode = ReturnDrgITemCode()
                strsql = "select trans_date= convert(varchar(20),a.trans_date,103) ,b.description ,a.Item_Code,Schedule_Quantity " & ",a.Schedule_Flag,a.cust_drgno,a.despatch_qty,RevisionNo,CONSIGNEE_CODE   " & " from dailymktschedule a,item_mst b, custitem_mst c " & " Where A.UNIT_CODE=B.UNIT_CODE  AND A.Unit_Code='" + gstrUNITID + "' AND A.item_code=b.item_code and a.item_code = c.item_code AND A.UNIT_CODE=C.UNIT_CODE and a.cust_drgno = c.cust_drgno  " & " and  a.Account_Code= '" & TxtAccountCode.Text & "' and a.consignee_code='" & Trim(Me.txtConsCode.Text) & "'  and a.account_code = c.account_code  and  right(convert(varchar(20),a.trans_date,103),7)='" & VB6.Format(DTPTransDate.Value, "MM/yyyy") & "' and schedule_flag=1"
                If Len(Trim(StrDiffitemCode)) > 0 Then
                    strsql = strsql & StrDiffitemCode
                End If
                strsql = strsql & " and a.Status =1 ORDER BY a.Item_Code,a.cust_drgno,trans_date"
                If mRdoCls.GetResult(strsql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly) And mRdoCls.GetNoRows > 0 Then
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
                        '************
                        fpSDailySchedule.Col = 3
                        fpSDailySchedule.Text = VB6.Format(mRdoCls.GetValue("trans_date"), gstrDateFormat)
                        fpSDailySchedule.Col = 4
                        fpSDailySchedule.Text = mRdoCls.GetValue("description")
                        fpSDailySchedule.Col = 5
                        fpSDailySchedule.Text = mRdoCls.GetValue("item_code")
                        fpSDailySchedule.Col = 6
                        fpSDailySchedule.CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                        fpSDailySchedule.Text = CStr(Val(mRdoCls.GetValue("Schedule_Quantity")))
                        fpSDailySchedule.Col = 12
                        fpSDailySchedule.CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                        fpSDailySchedule.Text = CStr(Val(mRdoCls.GetValue("Schedule_Quantity")))
                        fpSDailySchedule.Col = 7
                        If mRdoCls.GetValue("Schedule_flag") = "0" Then
                            fpSDailySchedule.CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
                            fpSDailySchedule.Value = "0"
                        Else
                            scheduleEdit = True
                            fpSDailySchedule.CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
                            fpSDailySchedule.Value = "1"
                        End If
                        fpSDailySchedule.Col = 13
                        If mRdoCls.GetValue("Schedule_flag") = "0" Then
                            fpSDailySchedule.CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
                            fpSDailySchedule.Value = "0"
                        Else
                            scheduleEdit = True
                            fpSDailySchedule.CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
                            fpSDailySchedule.Value = "1"
                        End If
                        fpSDailySchedule.Col = 8
                        fpSDailySchedule.Text = mRdoCls.GetValue("cust_drgno")
                        fpSDailySchedule.Col = 10
                        fpSDailySchedule.CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
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
        On Error GoTo ErrHandler
        ' to save data filled in grid in Add mode in case of Daily mkt schedule
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
        'dim Str_Upd_dt As String
        Dim Str_Upd_UserId As String
        Dim strsql As String
        Dim intRowCount As Short
        Dim intMaxRowCount As Short
        Dim dblOrderQty As Double
        Dim dblDispatchqty As Double
        Dim dblItemRows As Double
        Dim blninsertDaily As Boolean
        'Add new Column (Hidden)Open SO For Accounts Plug in
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
                If strMode = "A" Then
                    If blninsertDaily = True Then
                        strsql = ""
                        If Sng_Schedule_Quantity > 0 Then
                            strsql = "set dateformat 'DMY' Insert into dailymktschedule ( " & "Account_Code,Trans_date,cust_drgno," & "Schedule_Flag,Item_Code,Schedule_Quantity,Status," & "Ent_UserId,Upd_UserId,Ent_dt,Upd_dt,RevisionNo,CONSIGNEE_CODE,unit_code) values (" & Str_Account_Code & ",'" & getDateForDB(Str_Trans_date) & "'," & Str_cust_drgno & "," & Str_Schedule_Flag & "," & Str_Item_Code & "," & Sng_Schedule_Quantity & ",1,'" & Trim(mP_User) & "','" & Trim(mP_User) & "',getdate(),getdate(),0,'" & Trim(Me.txtConsCode.Text) & "','" + gstrUNITID + "')"
                        End If
                    End If
                Else
                    strsql = ""
                    If (Sng_Schedule_Quantity <> Sng_PrevSchedule_Quantity) Or (Str_Schedule_Flag <> Str_PrevSchedule_Flag) Then
                        strsql = "set dateformat 'DMY' Insert into dailymktschedule ( " & "Account_Code,Trans_date,cust_drgno," & "Schedule_Flag,Item_Code,Schedule_Quantity,Despatch_Qty,Status," & "Ent_UserId,Upd_UserId,Ent_dt,Upd_dt,RevisionNo,CONSIGNEE_CODE,Unit_Code) values (" & Str_Account_Code & ",'" & getDateForDB(Str_Trans_date) & "'," & Str_cust_drgno & "," & Str_Schedule_Flag & "," & Str_Item_Code & "," & Sng_Schedule_Quantity & "," & mdblDespatchQty(i - 1) & ",1,'" & Trim(mP_User) & "','" & Trim(mP_User) & "',getdate(),getdate()," & SngRevisionNo + 1 & ",'" & Trim(Me.txtConsCode.Text) & "','" + gstrUNITID + "')"
                    End If
                End If
                If Len(Trim(strsql)) > 0 Then
                    mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                End If
                'End If
            Next i
        End With
        CmdGrpMktSchedule.Revert()
        CmdGrpMktSchedule.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
        CmdGrpMktSchedule.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
        Call EnablePrimaryKeyControl()
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.OBJ_FORM, Me, System.Windows.Forms.Cursors.Default)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.OBJ_FORM, Me, System.Windows.Forms.Cursors.Default)
        Exit Sub
    End Sub
    Private Sub SaveUpdateDaily()
        ' to save data filled in grid in edit mode in case of daily marketing schedule
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
        Dim strsql As String
        Dim dblOrderQty As Double
        Dim dblDispatchqty As Double
        Dim Str_cons_Code As String
        On Error GoTo ErrHandler
        'Add new Column (Hidden)Open SO For Accounts Plug in
        Call ChangeMousePointer(eMPowerFunctions.ObjectsEnum.OBJ_FORM, Me, System.Windows.Forms.Cursors.WaitCursor)
        Str_Ent_UserId = "'" & mP_User & "'"
        Str_Upd_UserId = "'" & mP_User & "'"
        Str_Account_Code = "'" & TxtAccountCode.Text & "'"
        Str_cons_Code = "'" & Trim(Me.txtConsCode.Text) & "'"
        mStrCurrentDate = GetServerDate()
        ReDim mdblDespatchQty(fpSDailySchedule.MaxRows - 1)
        rsSchedule = New ClsResultSetDB
        For i = 1 To fpSDailySchedule.MaxRows
            fpSDailySchedule.Row = i
            fpSDailySchedule.Col = 2
            With fpSDailySchedule
                .Col = 3
                Str_Trans_date = .Text  ' mP_DayFormat
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
            strsql = " Select Despatch_Qty from dailymktschedule " & " Where Unit_Code='" + gstrUNITID + "' aND account_code=" & Str_Account_Code & " And consignee_code = " & Str_cons_Code & " and trans_date ='" & getDateForDB(Str_Trans_date) & "' and cust_drgno= " & Str_cust_drgno & " And Status = 1"
            rsSchedule = New ClsResultSetDB
            rsSchedule.GetResult(strsql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            rsSchedule.MoveFirst()
            mdblDespatchQty(i - 1) = Val(rsSchedule.GetValue("Despatch_Qty"))
            strsql = ""
            If (Sng_Schedule_Quantity <> Sng_PrevSchedule_Quantity) Or (Str_Schedule_Flag <> Str_PrevSchedule_Flag) Then
                strsql = "set dateformat DMY update dailymktschedule set  " & "Status =0 , Upd_dt='" & getDateForDB(mStrCurrentDate) & "',Upd_UserId =" & Str_Upd_UserId & " Where Unit_Code='" + gstrUNITID + "' AND account_code=" & Str_Account_Code & " and consignee_code=" & Str_cons_Code & " and trans_date ='" & getDateForDB(Str_Trans_date) & "' and cust_drgno= " & Str_cust_drgno
            End If
            rsSchedule.ResultSetClose()
            ' mP_Connection.BeginTrans
            If Len(Trim(strsql)) > 0 Then
                mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
            End If
            'mP_Connection.CommitTrans
            strsql = "select order_qty=sum(order_qty),despatch_qty=sum(despatch_qty)" & " from cust_ord_dtl Where Unit_Code='" + gstrUNITID + "' aND account_code='" & TxtAccountCode.Text & "' and item_code=" & Str_Item_Code & " and cust_Drgno = " & Str_cust_drgno & " and active_flag='A' and authorized_flag=1"
        Next i
        Call SaveAddDaily("E")
        CmdGrpMktSchedule.Revert()
        Call EnablePrimaryKeyControl()
        CmdGrpMktSchedule.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
        CmdGrpMktSchedule.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
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
        'to fill list with valid item code, customer item code and description
        Dim strsql As String
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
            strsql = "select distinct cust_drgno,Item_code from dailymktschedule "
            strsql = strsql & " Where Unit_Code='" + gstrUNITID + "' AND  account_code='" & TxtAccountCode.Text & "' and consignee_code='" & (Me.txtConsCode.Text) & "' and datepart(mm,trans_date) = " & VB6.Format(DTPTransDate.Value, "MM")
            strsql = strsql & " and datepart(yyyy,trans_date) = " & VB6.Format(DTPTransDate.Value, "YYYY") & " and status =1"
            rsDailyMktSchedule.GetResult(strsql)
            Dim dtTrans As Date = Date.Parse(DTPTransDate.Value)
            dtStartDate = Date.Parse(DTPTransDate.Value)
            Dim strDay As String = "-" & dtTrans.Day.ToString()
            dtStartDate = dtStartDate.AddDays(Double.Parse(strDay))
            dtEndDate = DateAdd(DateInterval.Month, 1, DTPTransDate.Value)
            intMaxLoop = rsDailyMktSchedule.GetNoRows
            rsDailyMktSchedule.MoveFirst()
            If intMaxLoop > 0 Then
                intDaysOfMonth = DateDiff(Microsoft.VisualBasic.DateInterval.Day, dtStartDate, dtEndDate) - 1
                strDailyDrgNo = ""
                For intLoopCounter = 1 To intMaxLoop
                    strsql = "select Account_Code,Trans_date,Item_code,Cust_Drgno,Serial_No,Schedule_Flag,Schedule_Quantity,Despatch_Qty,Status,RevisionNo,DSNO,DSDateTime,Spare_qty,Consignee_Code,FILETYPE,DOC_NO,MSG_NAME from dailymktschedule "
                    strsql = strsql & " Where Unit_Code='" + gstrUNITID + "' AND account_code='" & TxtAccountCode.Text & "'  and consignee_code='" & (Me.txtConsCode.Text) & "' and datepart(mm,trans_date) = " & VB6.Format(DTPTransDate.Value, "MM")
                    strsql = strsql & " and datepart(yyyy,trans_date) = " & VB6.Format(DTPTransDate.Value, "YYYY") & " and cust_drgNo = '" & rsDailyMktSchedule.GetValue("Cust_drgno") & "'"
                    strsql = strsql & " and Item_code = '" & rsDailyMktSchedule.GetValue("Item_Code") & "'  and status =1"
                    rsDailyNoOfRows.GetResult(strsql)
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
            strsql = "select distinct b.item_code,b.cust_drgno,a.description,b.cust_drg_desc from "
            strsql = strsql & " item_mst a,cust_ord_dtl b, custitem_mst c "
            strsql = strsql & " where A.UNIT_CODE=B.UNIT_CODE  AND  A.Unit_Code='" + gstrUNITID + "' AND  B.item_code = c.item_code AND B.UNIT_CODE=C.UNIT_CODE and b.cust_drgno = c.cust_drgno and b.account_code='" & TxtAccountCode.Text & "' and "
            strsql = strsql & " b.item_code=a.item_code and a.hold_flag <> 1 "
            'issue starts 10488279
            strsql = strsql & " and a.Item_Main_Grp in ('F','T','S','C','P','R','M')  "
            'issue end 10488279
            strsql = strsql & " and b.active_flag='A' and b.authorized_flag=1 "
            strsql = strsql & " and (b.order_qty>b.despatch_qty or b.Order_Qty =0)"
            If Len(Trim(strDailyDrgNo)) > 0 Then
                strsql = strsql & strDailyDrgNo
            End If
            strsql = strsql & " and b.cust_drgno not in (select cust_drgno from monthlymktschedule "
            strsql = strsql & " Where monthlymktschedule.Unit_Code='" + gstrUNITID + "' AND account_code='" & TxtAccountCode.Text & "' and "
            strsql = strsql & " year_month = " & VB6.Format(DTPTransDate.Value, "yyyyMM") & " and Schedule_Qty = 0) order by b.cust_drgno "
        ElseIf Me.CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Or Me.CmdGrpMktSchedule.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
            If OptDaily.Checked = True Then
                strsql = "select distinct b.item_code,b.cust_drgno,a.description,c.drg_desc from "
                strsql = strsql & " item_mst a,dailymktschedule b,custitem_mst c "
                strsql = strsql & " where A.UNIT_CODE=B.UNIT_CODE  AND  A.Unit_Code='" + gstrUNITID + "' AND b.account_code='" & TxtAccountCode.Text & "' and b.consignee_code='" & (Me.txtConsCode.Text) & "' and c.account_code='" & TxtAccountCode.Text & "' and "
                strsql = strsql & " b.item_code=a.item_code and b.cust_drgno=c.cust_drgno AND B.UNIT_CODE=C.UNIT_CODE and b.Item_code=c.Item_code  and right(convert(varchar(20),b.trans_date,103),7)='" & VB6.Format(DTPTransDate.Value, "MM/yyyy") & "'"
                strsql = strsql & " and a.Item_Main_Grp in ('F','T','S','C','P','R','M') and b.Status =1 order by b.cust_drgno"
            End If
        End If
        If mRdoCls.GetResult(strsql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly) And mRdoCls.GetNoRows > 0 Then
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
        'to return Items selected in list
        Dim StrItemCode As String
        Dim i As Short
        Dim strsql As String
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
        'to return Items selected in list
        Dim StrItemCode As String
        Dim i As Short
        Dim strsql As String
        Dim dblOrderQty As Double
        Dim dblDispatchqty As Double
        On Error GoTo ErrHandler
        StrItemCode = ""
        For i = 0 To LstItems.Items.Count - 1
            If LstItems.GetItemChecked(i) = True Then
                If Len(Trim(StrItemCode)) > 0 Then
                    ' Changed item _code position in list view  by rajesh
                    StrItemCode = StrItemCode & " or (a.Cust_drgNO ='" & Trim(Mid(ObsoleteManagement.GetItemString(LstItems, i), 1, 30)) & "' and a.Item_Code = '" & Trim(Mid(ObsoleteManagement.GetItemString(LstItems, i), 36, 16)) & "')"
                Else
                    StrItemCode = " and ((a.Cust_drgNO ='" & Trim(Mid(ObsoleteManagement.GetItemString(LstItems, i), 1, 30)) & "' and a.Item_Code = '" & Trim(Mid(ObsoleteManagement.GetItemString(LstItems, i), 36, 16)) & "')"
                End If
            End If
        Next i
        StrItemCode = StrItemCode & ")"
        ReturnDrgITemCode = StrItemCode
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Function
    End Function
    Private Function NoOfDaysinMonth(ByRef MonthNo As Short, ByRef YearNo As Short) As Short
        'to get no of days in month
        On Error GoTo ErrHandler
        Select Case MonthNo
            Case 1, 3, 5, 7, 8, 10, 12
                NoOfDaysinMonth = 31
            Case 2
                'To check leap year
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
        'TO clear the fields
        On Error GoTo ErrHandler
        TxtAccountCode.Text = ""
        lbldesc.Text = ""
        LstItems.Items.Clear()
        fpSDailySchedule.MaxRows = 0
        txtConsCode.Text = ""
        lblConsDesc.Text = ""
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Function ReturnAccountDescription(ByRef strAccountCode As String) As String
        'to return the description of Account code
        Dim strsql As String
        On Error GoTo ErrHandler
        strsql = "select Cust_Name from Customer_mst Where Unit_Code='" + gstrUNITID + "' AND Customer_code='" & Trim(strAccountCode) & "'"
        If mRdoCls.GetResult(strsql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly) And mRdoCls.GetNoRows > 0 Then
            ReturnAccountDescription = mRdoCls.GetValue("Cust_Name")
        Else
            ReturnAccountDescription = ""
        End If
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Private Sub DisablePrimaryKeyControl()
        'to disable controls having primary key value
        On Error GoTo ErrHandler
        TxtAccountCode.Enabled = False
        CmdHelpLocationCode.Enabled = False
        DTPTransDate.Enabled = False
        CmdSelectItems.Enabled = False
        Frame2.Enabled = False
        txtConsCode.Enabled = False
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub EnablePrimaryKeyControl()
        'to enable controls having primary key value
        On Error GoTo ErrHandler
        TxtAccountCode.Enabled = True
        CmdHelpLocationCode.Enabled = True
        DTPTransDate.Enabled = True
        CmdSelectItems.Enabled = True
        Frame2.Enabled = True
        txtConsCode.Enabled = True
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub PreviousDataView(ByRef pstrAccountCode As String, ByRef pstrDate As String, ByRef pStrItemCodes As String, ByRef pblnDailySchedule As Boolean)
        'Add new Column (Hidden)Open SO For Accounts Plug in
        Dim strItemCodes As String 'to stroe mstrItemCodes value localy for interchange
        Dim strsql As String
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
                strsql = "select trans_date= convert(varchar(20),a.trans_date,103) ,b.description ,a.Item_Code,Schedule_Quantity " & ",a.Schedule_Flag,a.cust_drgno   " & " from dailymktschedule a,item_mst b,custitem_mst c " & " where A.UNIT_CODE=B.UNIT_CODE  AND A.Unit_Code='" + gstrUNITID + "' AND a.item_code=b.item_code and a.item_code = c.item_code AND a.UNIT_CODE=c.UNIT_CODE and a.cust_drgno = c.cust_drgno " & " and  a.Account_Code= '" & TxtAccountCode.Text & "'  and a.consignee_code='" & Trim(Me.txtConsCode.Text) & "'  and  right(convert(varchar(20),a.trans_date,103),7)='" & VB6.Format(DTPTransDate.Value, "MM/yyyy") & "' and a.item_code in(" & pStrItemCodes & ") ORDER BY a.Item_Code,a.trans_date "
                If mRdoCls.GetResult(strsql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly) And mRdoCls.GetNoRows > 0 Then
                    mRdoCls.MoveFirst()
                    Do While Not mRdoCls.EOFRecord
                        fpSDailySchedule.MaxRows = fpSDailySchedule.MaxRows + 1
                        fpSDailySchedule.Row = fpSDailySchedule.MaxRows
                        SetGridDateFormat(fpSDailySchedule, 3)
                        fpSDailySchedule.Col = 3
                        fpSDailySchedule.Text = VB6.Format(mRdoCls.GetValue("trans_date"), gstrUNITID)
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
        ' To get Financil year dates from Company master
        Dim strsql As String
        On Error GoTo ErrHandler
        strsql = "select financial_startdate=convert(char(10),financial_startdate,103),financial_enddate=convert(char(10),financial_enddate,103) from company_mst Where Unit_Code='" + gstrUNITID + "'"
        If mRdoCls.GetResult(strsql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly) And mRdoCls.GetNoRows > 0 Then
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
        ' To change to fore colour of data coming in holidays
        Dim i As Short
        Dim intTotalItem As Short
        Dim intNoofDaysinMonth As Short
        Dim intItemNo As Short
        Dim intItemNo1 As Short
        Dim intItemRange1 As Short
        Dim intItemRange2 As Short
        Dim RsCalendar_Mst As New ADODB.Recordset
        Dim intRowNo As Short
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
            If IsReference(RsCalendar_Mst) Then
                If RsCalendar_Mst.State = 1 Then
                    RsCalendar_Mst.Close()
                End If
            End If
            RsCalendar_Mst.Open("SELECT CONVERT(VARCHAR(20),DT,103) FROM CALENDAR_MST " & " Where Unit_Code='" + gstrUNITID + "' AND right(CONVERT(varCHAR(20),DT,103),7)='" & VB6.Format(DTPTransDate.Value, "MM/yyyy") & "' and work_flg=1 ", mP_Connection, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If RsCalendar_Mst.RecordCount > 0 Then
                RsCalendar_Mst.MoveFirst()
                Do While Not RsCalendar_Mst.EOF
                    intRowNo = CShort(VB.Left(RsCalendar_Mst.Fields(0).Value, 2))
                    fpSDailySchedule.Row = intRowNo + (i - 1) * intNoofDaysinMonth
                    fpSDailySchedule.Col = 2
                    fpSDailySchedule.CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
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
        'To check the active flag of items to make enable and disable for editing
        Dim RsItemActiveFlag As New ADODB.Recordset
        Dim i As Short
        Dim strsql As String
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
            strsql = "SELECT active_flag FROM cust_ord_dtl " & " Where Unit_Code='" + gstrUNITID + "' AND Account_Code= '" & TxtAccountCode.Text & "' and item_code='" & StrItemCode & "' and cust_drgno ='" & strCustDrgNo & "'"
            If mRdoCls.GetResult(strsql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly) And mRdoCls.GetNoRows > 0 Then
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
        'To update tables dailymktschedule and monthlymktschedule if date is past or Ordered Quantity=Dispatch Quantity
        Dim strsql As String
        On Error GoTo ErrHandler
        'to assign current date
        mStrCurrentDate = GetServerDate()
        strsql = "update  dailymktschedule set schedule_flag='0' Where Unit_Code='" + gstrUNITID + "' AND trans_date <'" & getDateForDB(mStrCurrentDate) & "'"
        mP_Connection.Execute("set dateformat 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        strsql = "update  monthlymktschedule set schedule_flag='0' Where Unit_Code='" + gstrUNITID + "' AND year_month<" & YearMonth() & " "
        mP_Connection.Execute(strsql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub AssignCurrentMonthToDTPTRnasDate()
        'to assign current month to DTPTrnasDate
        Dim strsql As String
        On Error GoTo ErrHandler
        DTPTransDate.Value = GetServerDate()
        'to assign current date
        mStrCurrentDate = GetServerDate()
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub MakeEditableAndEditableLinesofGrid()
        'According to the status of items make editable and non editable
        Dim i As Short
        Dim varDummy As Object
        Dim varStatus As Object
        On Error GoTo ErrHandler
        'Add new Column (Hidden)Open SO For Accounts Plug in
        ' Changes done by Rajesh Sharma
        If OptDaily.Checked = True Then
            For i = 1 To fpSDailySchedule.MaxRows
                'varDummy = fpSDailySchedule.GetText(6, i, VarStatus)
                fpSDailySchedule.Row = i : fpSDailySchedule.Row2 = i : fpSDailySchedule.Col = 7 : fpSDailySchedule.Col2 = 7
                If CBool(fpSDailySchedule.Value) = False Then
                    fpSDailySchedule.BlockMode = True
                    fpSDailySchedule.Col = -1
                    'fpSDailySchedule.Col2 = 6
                    fpSDailySchedule.Row = i
                    fpSDailySchedule.Row2 = i
                    fpSDailySchedule.Lock = True
                    fpSDailySchedule.BlockMode = False
                Else
                    fpSDailySchedule.BlockMode = True
                    fpSDailySchedule.Col = 6
                    fpSDailySchedule.Col2 = 7
                    fpSDailySchedule.Row = i
                    fpSDailySchedule.Row2 = i
                    fpSDailySchedule.Lock = False
                    fpSDailySchedule.BlockMode = False
                End If
            Next i
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Public Function YearMonth() As Integer
        '----------------------------------------------------------------------------------
        'Author         :   Rajesh Sharma
        'Arguments      :   None
        'Return Value   :   Date
        'Procedure      :   Get the Current YearMonth
        '----------------------------------------------------------------------------------
        On Error GoTo ErrHandler
        Dim objYearMonth As ClsResultSetDB 'Class Object
        Dim strsql As String 'Stores the SQL statement
        strsql = "SELECT datepart(year,getdate()),datepart(month,getdate())"
        objYearMonth = New ClsResultSetDB
        With objYearMonth
            Call .GetResult(strsql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If .GetNoRows <= 0 Then Exit Function
            YearMonth = ((100 * .GetValueByNo(0)) + .GetValueByNo(1))
            .ResultSetClose()
        End With
        objYearMonth = Nothing
        Exit Function 'This is to avoid the execution of the error handler
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Function
    End Function
    Public Function ValidRecord() As Boolean
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
        Dim intRow As Short
        Dim varQuantity As Object
        Dim VarSchedule As Object
        Dim varDespatchQty As Object
        Dim VarTransDate As Object
        On Error GoTo Err_Handler
        'Add new Column (Hidden)Open SO For Accounts Plug in
        CheckScheduleQuantity = False
        If fpSDailySchedule.Visible = True Then
            For intRow = 1 To fpSDailySchedule.MaxRows
                VarTransDate = Nothing
                varQuantity = Nothing
                VarSchedule = Nothing
                varDespatchQty = Nothing
                Call fpSDailySchedule.GetText(3, intRow, VarTransDate)
                Call fpSDailySchedule.GetText(6, intRow, varQuantity)
                Call fpSDailySchedule.GetText(7, intRow, VarSchedule)
                Call fpSDailySchedule.GetText(10, intRow, varDespatchQty)
                Select Case CmdGrpMktSchedule.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        If Val(VarSchedule) = 1 Then
                            If System.Math.Abs(varQuantity) > 0 Then
                                CheckScheduleQuantity = True
                            End If
                        End If
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        If ConvertToDate(VarTransDate) >= GetServerDate() Then
                            If Val(VarSchedule) = 1 Then
                                If varQuantity >= varDespatchQty Then
                                    If System.Math.Abs(varQuantity) > 0 Then
                                        CheckScheduleQuantity = True
                                    End If
                                Else
                                    CheckScheduleQuantity = False
                                    Exit Function
                                End If
                            End If
                        End If
                End Select
            Next
        Else
        End If
        '******
        Exit Function
Err_Handler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Function
    Public Function ToCheckNoOfItemSelected() As Short
        'Add on 17/06/2002
        Dim intMaxCounter As Short
        Dim intLoopCounter As Short
        Dim IntItemSelected As Short
        intMaxCounter = LstItems.Items.Count
        If intMaxCounter > 0 Then
            For intLoopCounter = 0 To intMaxCounter - 1
                If LstItems.GetItemChecked(intLoopCounter) = True Then
                    IntItemSelected = IntItemSelected + 1
                End If
            Next
            ToCheckNoOfItemSelected = IntItemSelected
        End If
    End Function
    Public Function CheckForOpenSoFlag(ByRef pstrDrgno As String) As Boolean
        Dim rsOpenSO As New ClsResultSetDB
        Dim strOpenSO As Object
        Dim blnOpenSO As Boolean
        Dim intLoopCounter As Short
        Dim intMaxLoop As Short
        'To Check Open SO Flag in Cust_Ord_Dtl
        strOpenSO = "select  a.OpenSO,a.Item_Code,b.description,"
        strOpenSO = strOpenSO & "a.cust_drgno"
        strOpenSO = strOpenSO & " from cust_ord_dtl a,item_mst b, custitem_mst c "
        strOpenSO = strOpenSO & " Where A.UNIT_CODE=B.UNIT_CODE AND A.Unit_Code='" + gstrUNITID + "' AND a.item_code=b.item_code and a.cust_drgno = c.cust_drgno AND a.UNIT_CODE=c.UNIT_CODE and  a.Account_Code= '" & TxtAccountCode.Text & "' and a.cust_drgno = '" & pstrDrgno & "' and authorized_flag=1 and a.active_flag='A'"
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
        '**********
    End Function
    Private Sub fpSDailySchedule_LeaveCell(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_LeaveCellEvent) Handles fpSDailySchedule.LeaveCell
        On Error GoTo ErrHandler
        'to check data in grid
        Dim strsql As String 'StrSql to write sql before execution
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
        Dim dblItemCode As Object ' to assign item code from spread
        Dim strCustDrgNo As Object ' to assign Cust Drgno from spread
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
        Dim varLoopItemCode As Object
        Dim varLoopDrgNo As Object
        'In col 5 Schedule Quantity is there
        '****Changed by nisha on 14/09/2002
        'Add new Column (Hidden)Open SO For Accounts Plug in
        If e.newCol = -1 Then Exit Sub
        If e.col = 6 Then
            fpSDailySchedule.Row = e.row
            varStatus = Nothing
            dblDummy = fpSDailySchedule.GetText(7, e.row, varStatus)
            'To get the status
            If Trim(varStatus) = "1" Then
                'Check the mode
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
                        varLoopItemCode = Nothing
                        varLoopDrgNo = Nothing
                        Call fpSDailySchedule.GetText(5, intLoopCounter, varLoopItemCode)
                        Call fpSDailySchedule.GetText(8, intLoopCounter, varLoopDrgNo)
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
                            If dblTotalScheduleItemQty > (dblOrderQty - dblDispatchqty) Then
                                fpSDailySchedule.Row = e.row
                                fpSDailySchedule.Col = 6
                                dblDummy = fpSDailySchedule.GetFloat(6, e.row, DblCurrentScheduleQty)
                                fpSDailySchedule.Refresh()
                                MsgBox("Schedule Quantity Cannot be Greater than" & Str(dblOrderQty - dblDispatchqty), MsgBoxStyle.Critical, My.Resources.resEmpower.STR100)
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
                    fpSDailySchedule.Col = 6
                    dblItemCode = Nothing
                    strCustDrgNo = Nothing
                    'dblschDispatch = Nothing
                    dblDummy = fpSDailySchedule.GetText(5, e.row, dblItemCode)
                    dblDummy = fpSDailySchedule.GetText(8, e.row, strCustDrgNo)
                    dblDummy = fpSDailySchedule.GetFloat(10, e.row, dblschDispatch)
                    strsql = "select order_qty=sum(order_qty),despatch_qty=sum(despatch_qty)" & " from cust_ord_dtl Where Unit_Code='" + gstrUNITID + "' AND account_code='" & TxtAccountCode.Text & "' and item_code='" & dblItemCode & "' and cust_drgno = '" & strCustDrgNo & "' and  active_flag='A' and authorized_flag=1"
                    If mRdoCls.GetResult(strsql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly) And mRdoCls.GetNoRows > 0 Then
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
                            varLoopItemCode = Nothing
                            varLoopDrgNo = Nothing
                            Call fpSDailySchedule.GetText(5, intLoopCounter, varLoopItemCode)
                            Call fpSDailySchedule.GetText(8, intLoopCounter, varLoopDrgNo)
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
                        If dblOrderQty > 0 Then
                            If varOpenSO = 0 Then
                                If dblTotalScheduleItemQty > (Val(CStr(dblOrderQty)) - Val(CStr(dblDispatchqty))) Then
                                    fpSDailySchedule.Row = e.row
                                    fpSDailySchedule.Col = 6
                                    fpSDailySchedule.Text = CStr(0)
                                    dblDummy = fpSDailySchedule.SetFloat(6, e.row, 0)
                                    fpSDailySchedule.Refresh()
                                    MsgBox("Schedule Quantity Cannot be Greater than" & Str(dblOrderQty - dblDispatchqty), eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO, My.Resources.resEmpower.STR100)
                                    fpSDailySchedule.Focus()
                                    Exit Sub
                                End If
                            End If
                            dblDummy = fpSDailySchedule.GetFloat(6, e.row, dblcurrschqty)
                            dblDummy = fpSDailySchedule.GetFloat(10, e.row, dblcurrdisqty)
                            If (dblTotalScheduleItemQty < DblTotSchDispatch) Or (dblcurrschqty < dblcurrdisqty) Then
                                fpSDailySchedule.Row = e.row
                                fpSDailySchedule.Col = 6
                                fpSDailySchedule.Refresh()
                                MsgBox("Schedule quantity must be greater than" & Str(DblTotSchDispatch), eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO, My.Resources.resEmpower.STR100)
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            End If
        End If
        '*****
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Sub
    End Sub
    Private Sub ctlFormHeader1_Click(ByVal Sender As Object, ByVal e As System.EventArgs) Handles ctlFormHeader1.Click
        On Error GoTo ErrHandler
        Call ShowHelp("HLP" & Mid(Me.Name, 4, Len(Me.Name)) & ".htm") '("pddrep_listofitems.htm")
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
    End Sub
    Private Sub chkCheckAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkCheckAll.Click
        Dim intLoopCounter As Short
        Dim intMaxCounter As Short
        If chkCheckAll.CheckState = 1 Then
            If LstItems.Items.Count > 0 Then
                intMaxCounter = LstItems.Items.Count - 1
                For intLoopCounter = 0 To intMaxCounter
                    LstItems.SetItemChecked(intLoopCounter, True)
                Next
            End If
        End If
    End Sub
    Private Sub chkUnCheckAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkUnCheckAll.Click
        Dim intLoopCounter As Short
        Dim intMaxCounter As Short
        If chkUnCheckAll.CheckState = 1 Then
            If LstItems.Items.Count > 0 Then
                intMaxCounter = LstItems.Items.Count - 1
                For intLoopCounter = 0 To intMaxCounter
                    LstItems.SetItemChecked(intLoopCounter, False)
                Next
            End If
        End If
    End Sub
End Class