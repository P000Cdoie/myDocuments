Option Strict Off
Option Explicit On
Friend Class frmMKTMST0006
	Inherits System.Windows.Forms.Form
	'****************************************************
	'Copyright (c)  -  MIND
	'Name Of Module -  frmMKTMST0006.frm
	'Created By     -  Kapil
	'Created Date   -  23/04/2002
	'Description    -  Forecast Master
	'Revised Date   - By Ajay Vashistha on 22-07-2003
	'   1) Addition of Col. Customer Drawing No. in the Grid if availabe,
	'   2) Enabling of the Fields Customer Code on Pressing New Key (Runtime Error 5),
	'Revised Date   - By Jogender on 28-03-2006
	'   1) Use of Enagare_UNLOC to identify the details as Forecasting(FCST),Daily Scheduling,Monthly Scheduling,
	'   2) Editing of only those details which have Enagare_UNLOC='FSCT' ,
	'****************************************************
    'Revised By :   Shubhra Verma
    'Revised On :   '21 Apr 2011'
    'Description:   Multi Unit Changes.
    Dim mintIndex As Short 'Declared For The Form Count
    Const CNSTSCHDTYPE As String = "FCST" ''schedule type flagging
    Private mblneditmode As Boolean ''schedule type flagging
    Private mstrCustomerCode As String
    Private Sub cmdCustCode_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdCustCode.Click
        '--------------------------------------
        'Created By    :   Kapil
        'Description    :   Display Help from Item_Mst for Item_Code,Description
        '--------------------------------------
        On Error GoTo ErrHandler
        Dim StrCheckHelp As String
        Select Case Me.CmdGrpFCast.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                If Len(Me.txtCustCode.Text) = 0 Then
                    StrCheckHelp = ShowList(1, (txtCustCode.MaxLength), "", "a.Customer_Code", "b.cust_name", "Forecast_Mst a,CUSTOMER_Mst b", "and a.Customer_Code=b.customer_code and a.unit_code = b.unit_code", "Customer Codes List", , , , , "a.unit_code")
                Else
                    StrCheckHelp = ShowList(1, (txtCustCode.MaxLength), txtCustCode.Text, "a.Customer_Code", "b.cust_name", "Forecast_Mst a,customer_Mst b", "and a.Customer_Code=b.customer_code and a.unit_code = b.unit_code ", "Customer Codes List", , , , , "a.unit_code")
                End If
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                mstrCustomerCode = Trim(txtCustCode.Text)
                If Len(Me.txtCustCode.Text) = 0 Then
                    StrCheckHelp = ShowList(1, (txtCustCode.MaxLength), "", "a.customer_Code", "a.cust_name", "Customer_Mst a", "", "Customer Codes List")
                Else
                    StrCheckHelp = ShowList(1, (txtCustCode.MaxLength), txtCustCode.Text, "a.customer_Code", "a.cust_name", "Customer_Mst a", "", "Customer Codes List")
                End If
        End Select
        If StrCheckHelp = "-1" Then
            Call MsgBox("No Customer Code is Available To Display.", MsgBoxStyle.Information, "eMpro")
            txtCustCode.Text = "" : txtCustCode.Focus() : Exit Sub
        Else
            txtCustCode.Text = ""
            txtCustCode.Text = StrCheckHelp
            Call SelectCustomerCodeDescription()
        End If
        Call txtCustCode_Validating(txtCustCode, New System.ComponentModel.CancelEventArgs(False))
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub CmdItemCode_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdItemCode.Click
        '--------------------------------------
        'Created By     :   Kapil
        'Description    :   Display Help from Item_Mst for Item_Code,Description
        '--------------------------------------
        On Error GoTo ErrHandler
        Dim StrCheckHelp As String
        Select Case Me.CmdGrpFCast.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW, UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                If Len(Me.txtItemCode.Text) = 0 And Len(txtCustCode.Text) > 0 Then
                    StrCheckHelp = ShowList(1, (txtItemCode.MaxLength), "", "a.product_no", "b.Description", "Forecast_Mst a,Item_Mst b", "and a.unit_code = b.unit_code and b.Item_Main_Grp='F' and a.product_no=b.Item_Code and a.Customer_Code='" & Trim(txtCustCode.Text) & "'", "Item Codes List", , , , , "a.unit_code")
                ElseIf Len(txtItemCode.Text) = 0 And Len(txtCustCode.Text) = 0 Then
                    StrCheckHelp = ShowList(1, (txtItemCode.MaxLength), "", "a.product_no", "b.Description", "Forecast_Mst a,Item_Mst b", "and a.unit_code = b.unit_code and b.Item_Main_Grp='F' and a.product_no=b.Item_Code", "Item Codes List", , , , , "a.unit_code")
                End If
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                If Len(Me.txtItemCode.Text) = 0 Then
                    StrCheckHelp = ShowList(1, (txtItemCode.MaxLength), "", "Item_Code", "Description", "Item_Mst", "and Item_Main_Grp='F'", "Item Codes List", , , , , "unit_code")
                Else
                    StrCheckHelp = ShowList(1, (txtItemCode.MaxLength), txtItemCode.Text, "Item_Code", "Description", "Item_Mst", "and Item_Main_Grp='F'", "Item Codes List")
                End If
        End Select
        If StrCheckHelp = "-1" Then
            Call MsgBox("No Item Code is Available To Display.", MsgBoxStyle.Information, "eMpro")
            txtItemCode.Text = ""
        Else
            txtItemCode.Text = ""
            txtItemCode.Text = StrCheckHelp
        End If
        Call txtItemCode_Validating(txtItemCode, New System.ComponentModel.CancelEventArgs(False))
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub CmdRefresh_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdRefresh.Click
        On Error GoTo ErrHandler
        txtCustCode.Text = "" : txtCustCode.Enabled = True : txtCustCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        txtItemCode.Text = "" : txtItemCode.Enabled = True : txtItemCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
        DTPFromDate.Enabled = True : DTPFromDate.CalendarMonthBackground = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : DTPFromDate.Value = GetServerDate()
        dtpToDate.Enabled = True : dtpToDate.CalendarMonthBackground = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : dtpToDate.Value = GetServerDate()
        CmdCustCode.Enabled = True : CmdItemCode.Enabled = True : txtCustCode.Focus()
        CmdViewDetails.Enabled = True 'jogender
        Me.CmdGrpFCast.Revert()
        Me.CmdGrpFCast.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False 'jogender
        With Me.SpForCast
            .MaxRows = 0 : Call addRowAtEnterKeyPress(1)
        End With
        'Samiksha forecast master changes
        selectallchkbox.Checked = False
        selectallchkbox.Enabled = True
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub cmdviewDetails_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CmdViewDetails.Click
        If Len(txtCustCode.Text) = 0 And Len(txtItemCode.Text) = 0 Then
            If Len(txtCustCode.Text) = 0 Then
                Call MsgBox("Enter Customer Code.", MsgBoxStyle.Information, "eMpro")
                txtCustCode.Focus() : Exit Sub
            End If
            If Len(txtItemCode.Text) = 0 Then
                Call MsgBox("Enter Item Code.", MsgBoxStyle.Information, "eMpro")
                txtItemCode.Focus() : Exit Sub
            End If
        End If
        If DisplayForecastDetails((txtCustCode.Text), (txtItemCode.Text)) Then
            txtCustCode.Enabled = False : txtCustCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            txtItemCode.Enabled = False : txtItemCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            DTPFromDate.Enabled = False : DTPFromDate.CalendarMonthBackground = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            dtpToDate.Enabled = False : dtpToDate.CalendarMonthBackground = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
            CmdCustCode.Enabled = False : CmdItemCode.Enabled = False
            Me.CmdGrpFCast.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = True
            Me.CmdGrpFCast.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = True
            Me.CmdGrpFCast.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
            Me.CmdGrpFCast.Focus()
        Else
            With Me.SpForCast
                .MaxRows = 0 : Call addRowAtEnterKeyPress(1)
            End With
            Me.CmdGrpFCast.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
            Me.CmdGrpFCast.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
            Me.CmdGrpFCast.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub dtpFromDate_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSComCtl2.DDTPickerEvents_KeyDownEvent)
        Select Case eventArgs.keyCode
            Case System.Windows.Forms.Keys.Return
                DTPToDate.Focus()
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub dtpToDate_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As AxMSComCtl2.DDTPickerEvents_KeyDownEvent)
        Select Case eventArgs.keyCode
            Case System.Windows.Forms.Keys.Return
                CmdViewDetails.Focus()
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub frmMKTMST0006_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        On Error GoTo ErrHandler
        mdifrmMain.CheckFormName = mintIndex
        If txtCustCode.Enabled Then txtCustCode.Focus()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub frmMKTMST0006_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
        On Error GoTo ErrHandler
        frmModules.NodeFontBold(Me.Tag) = False
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub frmMKTMST0006_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Escape
                If Me.CmdGrpFCast.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                    If ConfirmWindow(10053, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                        Call Me.CmdGrpFCast.Revert()
                        gblnCancelUnload = False : gblnFormAddEdit = False
                        txtCustCode.Text = "" : txtCustCode.Enabled = True : txtCustCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        CmdCustCode.Enabled = True : txtItemCode.Text = "" : txtItemCode.Enabled = True : txtItemCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                        CmdItemCode.Enabled = True : CmdViewDetails.Enabled = True
                        Me.CmdGrpFCast.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                        Me.CmdGrpFCast.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                        Me.CmdGrpFCast.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                        DTPFromDate.Enabled = True : DTPToDate.Enabled = True
                        With Me.SpForCast
                            .MaxRows = 0 : Call addRowAtEnterKeyPress(1)
                        End With
                        Me.ToolTip1.SetToolTip(Me.SpForCast, "")
                        txtCustCode.Focus()
                    Else
                        Me.ActiveControl.Focus()
                    End If
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
    Private Sub frmMKTMST0006_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        On Error GoTo ErrHandler
        Call FitToClient(Me, FraForeCast, ctlFormHeader1, CmdGrpFCast)
        Call FillLabelFromResFile(Me)
        mintIndex = mdifrmMain.AddFormNameToWindowList(ctlFormHeader1.Tag)
        With Me.SpForCast
            .MaxRows = 0
            .set_RowHeight(0, 300)
            .MaxCols = 7
            .Row = 0
            .Col = 1 : .Text = " "
            .Row = 0
            .Col = 2 : .Text = "Item Code"
            .Row = 0
            .Col = 3 : .Text = " "
            .Row = 0
            .Col = 4 : .Text = "Description"
            .Row = 0
            .Col = 5 : .Text = "Due Date"
            .Row = 0
            .Col = 6 : .Text = "Quantity"
            .Row = 0
            .Col = 7 : .Text = "Drawing Number"
        End With
        Me.CmdGrpFCast.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
        Me.CmdGrpFCast.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
        Me.CmdGrpFCast.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
        Call addRowAtEnterKeyPress(1)
        DTPFromDate.Format = DateTimePickerFormat.Custom
        DTPFromDate.CustomFormat = gstrDateFormat
        DTPFromDate.Value = GetServerDate()
        DTPFromDate.Format = DateTimePickerFormat.Custom
        DTPToDate.CustomFormat = gstrDateFormat
        DTPToDate.Value = GetServerDate()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub frmMKTMST0006_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim Cancel As Boolean = eventArgs.Cancel
        Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
        On Error GoTo ErrHandler
        Dim enmValue As eMPowerFunctions.ConfirmWindowReturnEnum
        If UnloadMode >= 0 And UnloadMode <= 5 Then
            If CmdGrpFCast.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW Then
                enmValue = ConfirmWindow(10055, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNOCANCEL, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_QUESTION)
                If enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_NO Or enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                    If enmValue = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                        'check if there is any blank row in spread
                        Call SpForCast_Validating(SpForCast, New System.ComponentModel.CancelEventArgs(False))
                        'Save data before saving
                        Call CmdGrpFCast_ButtonClick(eventSender, New UCActXCtl.UCbtndgrp.ButtonClickEventArgs(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE))
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
                Me.Dispose()
                Exit Sub
            End If
        End If
        'Checking The Status
        If gblnCancelUnload = True Then eventArgs.Cancel = True
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub frmMKTMST0006_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        On Error GoTo ErrHandler
        'Me = Nothing
        mdifrmMain.RemoveFormNameFromWindowList = mintIndex
        frmModules.NodeFontBold(Me.Tag) = False
        Me.Dispose()
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub addRowAtEnterKeyPress(ByRef pintRows As Short)
        '****************************************************
        'Created By     -  Kapil
        'Description    -  Add Row At Enter Key Press Of Last Column Of Spread
        '****************************************************
        On Error GoTo ErrHandler
        Dim intRowHeight As Short
        With Me.SpForCast
            .CursorStyle = FPSpreadADO.CursorStyleConstants.CursorStyleArrow
            For intRowHeight = 1 To pintRows
                .MaxRows = .MaxRows + 1
                .Row = .MaxRows
                .set_RowHeight(.Row, 300)
                Call SetSpreadColTypes(.Row)
            Next intRowHeight
        End With
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub SetSpreadColTypes(ByRef pintRowNo As Short)
        '****************************************************
        'Created By     -  Kapil
        'Description    -  Set Spread Columns Properties for the Row
        '****************************************************
        On Error GoTo ErrHandler
        Select Case Me.CmdGrpFCast.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                With Me.SpForCast
                    .Enabled = True
                    .Row = pintRowNo
                    .Col = 1
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
                    .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
                    .Value = System.Windows.Forms.CheckState.Unchecked
                    .Col = 2
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                    .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
                    .Col = 3
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                    .Col = 4
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                    .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
                    .Col = 5
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                    .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
                    .Col = 6
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                    .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                    .Col = 7
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                    .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                    .Row = 0
                    .Row2 = .MaxRows
                    .Col = 2
                    .Col2 = .MaxCols
                    .BlockMode = True
                    .Lock = True
                    .BlockMode = False
                End With
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD


                With Me.SpForCast
                    .Enabled = True
                    .Row = pintRowNo
                    .Col = 1
                    '.CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
                    '.TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
                    '.Value = System.Windows.Forms.CheckState.Unchecked

                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                    .Col = 2
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeEdit
                    .TypeMaxEditLen = 16
                    .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
                    .Col = 3
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeButton
                    .TypeButtonPicture = My.Resources.resEmpower.ico111.ToBitmap()
                    .Col = 4
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                    .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
                    .Col = 5

                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeDate
                    SetGridDateFormat(SpForCast, 5)
                    .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter : .Text = VB6.Format(GetServerDate(), gstrDateFormat)

                    .Col = 6
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                    .TypeFloatMin = "0.00"
                    .TypeFloatMax = "9999999999"
                    .TypeFloatDecimalPlaces = 2
                    .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                    .Col = 7
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                    .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
                End With
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
               

                With Me.SpForCast
                    .Enabled = True
                    .Row = pintRowNo
                    .Col = 1
                    '.CellType = FPSpreadADO.CellTypeConstants.CellTypeCheckBox
                    '.TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
                    '.Value = System.Windows.Forms.CheckState.Unchecked


                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                    .Col = 2
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                    .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
                    .Col = 3
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                    .Col = 4
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                    .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
                    .Col = 5

                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeDate
                    SetGridDateFormat(SpForCast, 5)
                    .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignCenter
                    .Lock = True
                    .Col = 6
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                    .TypeFloatMin = "0.00"
                    .TypeFloatMax = "9999999999"
                    .TypeFloatDecimalPlaces = 2
                    .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                    .Col = 7
                    .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText
                    .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignLeft
                    '------------>>
                End With
        End Select
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Function SelectDescription(ByRef pstrItemCode As String) As String
        '**************************************
        'Created By     :   Kapil
        'Description    :   Select Description Of Measure Code
        '**************************************
        On Error GoTo ErrHandler
        Dim strSelectSql As String 'Declared To Make Select Query
        Dim rsSelect As ClsResultSetDB
        strSelectSql = "Select Description from Item_Mst where Item_Code='" & Trim(pstrItemCode) & "' and unit_code = '" & gstrUNITID & "'"
        rsSelect = New ClsResultSetDB
        rsSelect.GetResult(strSelectSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsSelect.GetNoRows > 0 Then
            SelectDescription = rsSelect.GetValue("Description")
        Else
            SelectDescription = ""
        End If
        rsSelect.ResultSetClose()
        rsSelect = Nothing
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Function
    End Function
    Private Function SelectDataFromTable(ByRef pstrFName As String, ByRef pstrTName As String, Optional ByRef pstrCon As String = "") As Boolean
        '****************************************************
        'Created By     -  Kapil
        'Description    -  Check Validity Of Data In The Table
        'Arguments      -  pstrFName - Field Name,pstrTName - Table Name,pstrCon - Condition
        '****************************************************
        On Error GoTo ErrHandler
        SelectDataFromTable = False
        Dim strSelectSql As String 'Declared To Make Select Query
        Dim rsCheckData As ClsResultSetDB
        strSelectSql = "Select " & Trim(pstrFName) & " from " & Trim(pstrTName) & " " & pstrCon
        rsCheckData = New ClsResultSetDB
        rsCheckData.GetResult(strSelectSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsCheckData.GetNoRows > 0 Then
            SelectDataFromTable = True
        Else
            SelectDataFromTable = False
        End If
        rsCheckData.ResultSetClose()
        rsCheckData = Nothing
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Function
    End Function
    Private Sub SpForCast_ButtonClicked(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ButtonClickedEvent) Handles SpForCast.ButtonClicked
        On Error GoTo ErrHandler
        Dim StrCheckHelp As String
        Dim varGetText As Object
        Dim varGetColText As Object
        With Me.SpForCast
            If e.col = 3 Then
                varGetText = Nothing
                Call .GetText(e.col - 1, e.row, varGetText)
                If Len(varGetText) = 0 Then
                    varGetText = ShowList(1, 3, "", "Item_Mst.Item_Code", "Item_Mst.Description", "Item_Mst,CustItem_Mst", "and item_mst.unit_code = custitem_mst.unit_code and Item_Main_Grp='F' and CustItem_Mst.item_code=item_mst.item_code and CustItem_Mst.account_code='" & Trim(txtCustCode.Text) & "'", "Item Codes List", , , , , "item_mst.unit_code")
                Else
                    varGetText = ShowList(1, Len(varGetText), CStr(varGetText), "Item_Mst.Item_Code", "Item_Mst.Description", "Item_Mst,CustItem_Mst", "and item_mst.unit_code = custitem_mst.unit_code and Item_Main_Grp='F' and CustItem_Mst.item_code=item_mst.item_code and CustItem_Mst.account_code='" & Trim(txtCustCode.Text) & "'", "Item Codes List", , , , , "item_mst.unit_code")
                End If
                If Trim(CStr(varGetText)) = "-1" Then Exit Sub
                If Trim(CStr(varGetText)) = "-1" Then
                    Call MsgBox("No Item Codes Defined Like [" & varGetText & "].", vbInformation, "eMpro")
                    Call .SetText(e.col - 1, e.row, "")
                    Call .SetText(e.col + 1, e.row, "")
                    .Row = e.row : .Col = 2 : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                Else
                    Call .SetText(e.col - 1, e.row, Trim(CStr(varGetText)))
                    'Function Call To Select Description
                    Call .SetText(e.col + 1, e.row, Trim(SelectDescription(CStr(varGetText))))
                    .Row = e.row : .Col = 6 : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                End If
                If Trim(CStr(varGetText)) = "-1" Then
                    Call .SetText(7, e.row, "")
                    .Row = e.row : .Col = 7 : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                Else
                    Call .SetText(e.col - 1, e.row, Trim(CStr(varGetText)))
                    'Function Call To Select Corrosponding Drawing No. of Customer
                    Call .SetText(7, e.row, Trim(SelectDrawing(CStr(varGetText))))
                    .Row = e.row : .Col = 6 : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                End If
                'UOM (consumption so that decimal if required will be accepted)
                varGetColText = Nothing
                Call .GetText(2, e.row, varGetColText)
                If Len(CStr(varGetColText)) > 0 Then
                    'Check Validity Of Item Code
                    If SelectDataFromTable("Item_Mst.Item_Code", "Item_Mst,CustItem_Mst", "Where item_mst.unit_code ='" & gstrUNITID & "' and  item_mst.unit_code = custitem_mst.unit_code and Item_Mst.Item_Code='" & Trim(CStr(varGetColText)) & "' and CustItem_Mst.item_code=item_mst.item_code and CustItem_Mst.account_code='" & Trim(txtCustCode.Text) & "'") Then
                        Call .SetText(4, e.row, SelectDescription(CStr(varGetColText)))
                        'Check Weather Decimal Allowed Flag Is Checked Or Not
                        If SelectDataFromTable("Decimal_Allowed_Flag", "Measure_Mst", "Where unit_code = '" & gstrUNITID & "' and Measure_Code=(Select cons_measure_code From Item_Mst Where unit_code = '" & gstrUNITID & "' and Item_Code='" & CStr(varGetColText) & "' and Decimal_Allowed_Flag=1)") Then ''cons_measure_code in place of pur_measure_code
                            .Row = e.row : .Col = 6 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMin = "0.00" : .TypeFloatMax = "9999999999" : .TypeFloatDecimalPlaces = 2 : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                        Else
                            .Row = e.row : .Col = 6 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMin = "0" : .TypeFloatMax = "9999999999" : .TypeFloatDecimalPlaces = 0 : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                        End If
                    End If
                End If
            End If
        End With
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub SpForCast_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles SpForCast.Enter
        On Error GoTo ErrHandler
        If Me.CmdGrpFCast.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            Me.ToolTip1.SetToolTip(Me.SpForCast, "Press Ctrl+N To Add New Row")
        Else
            Me.ToolTip1.SetToolTip(Me.SpForCast, "")
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub SpForCast_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles SpForCast.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        Dim varGetColText As Object 'Declared To Get Value Entered In The Column
        Dim intRowCount As Short
        Dim blnBlankRowFlag As Boolean 'Declared To Make The Loop Counter
        If Me.CmdGrpFCast.Mode = UCActXCtl.clsDeclares.ModeEnum.MODE_ADD Then
            With Me.SpForCast
                For intRowCount = 1 To .MaxRows
                    blnBlankRowFlag = False
                    If intRowCount <= .MaxRows Then
                        blnBlankRowFlag = False
                        .Row = intRowCount
                        .Col = 2
                        If Len(.Text) > 0 Then ' Some value has been entered in any of the fields
                            blnBlankRowFlag = False ' Row will not be deleted
                        Else
                            blnBlankRowFlag = True ' If row is completely blank, flag is set
                        End If
                    End If
                    If blnBlankRowFlag = True Then ' If blank row
                        .Row = intRowCount
                        .Action = FPSpreadADO.ActionConstants.ActionDeleteRow ' Row is deleted
                        .MaxRows = .MaxRows - 1 ' It is deleted from the spread
                    End If
                Next intRowCount ' End
                If Me.SpForCast.MaxRows = 0 Then Call addRowAtEnterKeyPress(1)
            End With
        End If
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        GoTo EventExitSub
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Function DisplayForecastDetails(ByRef pstrCustCode As String, ByRef pstrItemCode As String) As Boolean
        '****************************************
        'Created By     :   Kapil
        'Description    :   Display Initial Machine Details From Conversion_Mst
        '****************************************
        On Error GoTo ErrHandler
        DisplayForecastDetails = False
        Dim strSelectSql As String 'Declared To Make Select Query
        Dim rsSelect As ClsResultSetDB
        Dim intRecordCount As Short 'Declared To Get Total Record Count
        Dim intLoopcount As Short 'Declared For The Loop Counter
        'Make Select Query
        If Len(pstrCustCode) > 0 And Len(pstrItemCode) = 0 Then
            strSelectSql = "Select product_no,Due_Date,Quantity,Enagare_UNLOC from Forecast_Mst"
            strSelectSql = strSelectSql & " Where unit_code = '" & gstrUNITID & "' and Customer_Code='" & Trim(pstrCustCode) & "'"
            strSelectSql = strSelectSql & " And Due_Date between '" & getDateForDB(DTPFromDate.Value) & "' and '" & getDateForDB(dtpToDate.Value) & "'"
            If mblneditmode Then strSelectSql = strSelectSql & " And Enagare_UNLOC='" & CNSTSCHDTYPE & "'"
            strSelectSql = strSelectSql & " Order By product_no,Due_Date"
        ElseIf Len(pstrCustCode) > 0 And Len(pstrItemCode) > 0 Then
            strSelectSql = "Select product_no,Due_Date,Quantity,Enagare_UNLOC from Forecast_Mst"
            strSelectSql = strSelectSql & " Where unit_code = '" & gstrUNITID & "' and Customer_Code='" & Trim(pstrCustCode) & "'"
            strSelectSql = strSelectSql & " and Product_no='" & Trim(pstrItemCode) & "'"
            strSelectSql = strSelectSql & " And Due_Date between '" & getDateForDB(DTPFromDate.Value) & "' and '" & getDateForDB(dtpToDate.Value) & "'"
            If mblneditmode Then strSelectSql = strSelectSql & " And Enagare_UNLOC='" & CNSTSCHDTYPE & "'"
            strSelectSql = strSelectSql & " Order By product_no,Due_Date"
        ElseIf Len(pstrCustCode) = 0 And Len(pstrItemCode) > 0 Then
            strSelectSql = "Select product_no,Due_Date,Quantity,Enagare_UNLOC from Forecast_Mst"
            strSelectSql = strSelectSql & " Where unit_code = '" & gstrUNITID & "' and product_no='" & Trim(pstrItemCode) & "'"
            strSelectSql = strSelectSql & " And Due_Date between '" & getDateForDB(DTPFromDate.Value) & "' and '" & getDateForDB(dtpToDate.Value) & "'"
            If mblneditmode Then strSelectSql = strSelectSql & " And Enagare_UNLOC='" & CNSTSCHDTYPE & "'"
            ''''''''''''''
            strSelectSql = strSelectSql & " Order By product_no,Due_Date"
        End If
        rsSelect = New ClsResultSetDB
        rsSelect.GetResult(strSelectSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        intRecordCount = rsSelect.GetNoRows
        If intRecordCount > 0 Then
            DisplayForecastDetails = True
            If Not rsSelect.EOFRecord Then rsSelect.MoveFirst()
            'Add Total No. Of Rows
            Call addRowAtEnterKeyPress(intRecordCount - 1)
            With Me.SpForCast
                For intLoopcount = 1 To intRecordCount
                    .Row = intLoopcount
                    .Col = 2
                    .Text = Trim(rsSelect.GetValue("product_no"))
                    .Row = intLoopcount
                    .Col = 4
                    .Text = Trim(SelectDescription(rsSelect.GetValue("product_no")))
                    .Row = intLoopcount
                    .Col = 5
                    .Text = VB6.Format(rsSelect.GetValue("Due_Date"), gstrDateFormat)
                    .Row = intLoopcount
                    .Col = 6
                    .Text = Trim(rsSelect.GetValue("Quantity"))
                    .Row = intLoopcount
                    .Col = 7
                    .Text = Trim(SelectDrawing(rsSelect.GetValue("product_no")))
                    'Samiksha changes in forecast master
                    If Not (Trim(rsSelect.GetValue("Enagare_UNLOC")) = CNSTSCHDTYPE Or Trim(rsSelect.GetValue("Enagare_UNLOC")) = "") Then
                        .Row = intLoopcount
                        .Col = 1
                        .CellType = FPSpreadADO.CellTypeConstants.CellTypeStaticText : .Text = ""
                    End If
                    '''''''''''''''''''
                    rsSelect.MoveNext()
                Next
            End With
        Else
            DisplayForecastDetails = False
        End If
        rsSelect.ResultSetClose()
        selectallchkbox.Checked = False
        rsSelect = Nothing
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Function
    End Function
    Private Function CheckPreviousRecords() As Boolean
        '**************************************
        'Created By     :   Kapil
        'Description    :   Check Source Machine Code and Destination Machine Code
        '                   b'coz these are the Part Of Primary Key
        '**************************************
        On Error GoTo ErrHandler
        CheckPreviousRecords = False
        Dim intCheckRow As Short 'Declared To Check Previous Row
        Dim varCheckPrevItemCode As Object 'Declared For Prev Source Machine Code
        Dim varCheckPrevDueDate As Object 'Declared For Prev Destination Machine Code
        Dim varCheckNextItemCode As Object 'Declared For Next Source Machine Code
        Dim varCheckNextDueDate As Object 'Declared For Next Destination Machine
        With Me.SpForCast
            For intCheckRow = 1 To .MaxRows - 1
                varCheckPrevItemCode = Nothing
                Call .GetText(2, intCheckRow, varCheckPrevItemCode)
                varCheckPrevDueDate = Nothing
                Call .GetText(5, intCheckRow, varCheckPrevDueDate)
                varCheckNextItemCode = Nothing
                Call .GetText(2, .MaxRows, varCheckNextItemCode)
                varCheckNextDueDate = Nothing
                Call .GetText(5, .MaxRows, varCheckNextDueDate)
                If Trim(varCheckNextItemCode) = Trim(varCheckPrevItemCode) And Trim(varCheckNextDueDate) = Trim(varCheckPrevDueDate) Then
                    CheckPreviousRecords = True
                    Call .SetText(2, .MaxRows, "")
                    Call MsgBox("Forecast Details For Item Code [" & varCheckNextItemCode & "] Is Already Defined For Date [" & varCheckNextDueDate & "].", MsgBoxStyle.Information, "eMpro")
                    Call .SetText(2, .MaxRows, varCheckNextItemCode)
                    .Row = .MaxRows
                    .Col = 2
                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                    Exit Function
                Else
                    CheckPreviousRecords = False
                End If
            Next intCheckRow
        End With
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Function
    End Function
    Private Sub SelectCustomerCodeDescription()
        '**************************************
        'Created By     :   Kapil
        'Description    :   Select Customer Description From Customer Master
        '**************************************
        On Error GoTo ErrHandler
        Dim strCustCodeDes As String
        Dim rsCustCode As ClsResultSetDB 'Decalred To Make Select Query
        If Len(txtCustCode.Text) > 0 Then
            strCustCodeDes = "Select customer_Code,cust_name from customer_Mst" & _
                " where unit_code = '" & gstrUNITID & "' and customer_Code='" & Trim(txtCustCode.Text) & "'"
            rsCustCode = New ClsResultSetDB
            rsCustCode.GetResult(strCustCodeDes, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
            If rsCustCode.GetNoRows > 0 Then
                lblCustCodeDes.Text = rsCustCode.GetValue("cust_name")
            End If
            rsCustCode.ResultSetClose()
            rsCustCode = Nothing
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub txtCustCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCustCode.TextChanged
        On Error GoTo ErrHandler
        If Len(txtCustCode.Text) = 0 Then
            Select Case Me.CmdGrpFCast.Mode
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                    lblCustCodeDes.Text = ""
                    txtItemCode.Text = ""
                    Me.SpForCast.MaxRows = 0 : Call addRowAtEnterKeyPress(1)
                    Me.CmdGrpFCast.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                    Me.CmdGrpFCast.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                    Me.CmdGrpFCast.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                    lblCustCodeDes.Text = ""
            End Select
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub txtCustCode_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCustCode.Enter
        On Error GoTo ErrHandler
        txtCustCode.SelectionStart = 0 : txtCustCode.SelectionLength = Len(txtCustCode.Text)
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub txtCustCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtCustCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpFCast.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                        If Len(txtCustCode.Text) > 0 Then
                            Call txtCustCode_Validating(txtCustCode, New System.ComponentModel.CancelEventArgs(False))
                        Else
                            txtItemCode.Focus()
                        End If
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        If Len(txtCustCode.Text) > 0 Then
                            Call txtCustCode_Validating(txtCustCode, New System.ComponentModel.CancelEventArgs(False))
                        Else
                            With Me.SpForCast
                                .Row = .MaxRows
                                .Col = 2
                                .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                            End With
                        End If
                End Select
            Case 39, 34, 96
                KeyAscii = 0
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
    Private Sub txtCustCode_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtCustCode.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then
            If CmdCustCode.Enabled Then Call cmdCustCode_Click(CmdCustCode, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub txtCustCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtCustCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        '**************************************
        'Created By     :   Kapil
        'Description    :   Select Data from Forecast_Mst/Customer_Mst and Check Validity
        '**************************************
        On Error GoTo ErrHandler
        If Len(txtCustCode.Text) > 0 Then
            Select Case Me.CmdGrpFCast.Mode
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                    'Check Whether Customer Code Exists in the Table Or Not
                    If Not SelectDataFromTable("Customer_Code", "Forecast_Mst", " where unit_code = '" & gstrUNITID & "' and Customer_Code='" & Trim(txtCustCode.Text) & "'") Then
                        'If Invalid Customer Code Then Display Message
                        Call MsgBox("Invalid Customer Code.Press F1 For help.", MsgBoxStyle.Information, "eMpro")
                        Cancel = True : txtCustCode.Text = "" : txtCustCode.Focus()
                        GoTo EventExitSub
                    Else
                        Call SelectCustomerCodeDescription()
                        If txtItemCode.Enabled Then txtItemCode.Focus()
                    End If
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                    'Check Whether Customer Code Exists in the Table Or Not
                    If Not SelectDataFromTable("customer_Code", "Customer_Mst", " Where unit_code = '" & gstrUNITID & "' and customer_Code='" & Trim(txtCustCode.Text) & "'") Then
                        'If Invalid Customer Code Then Display Message
                        Call MsgBox("Invalid Customer Code.Press F1 For help.", MsgBoxStyle.Information, "eMpro")
                        Cancel = True : txtCustCode.Text = "" : txtCustCode.Focus()
                        GoTo EventExitSub
                    Else
                        Call SelectCustomerCodeDescription()
                        With Me.SpForCast
                            '''If customer is changed after filling the details then details need to be cleared.
                            If mstrCustomerCode <> Trim(txtCustCode.Text) Then
                                .MaxRows = 0 : Call addRowAtEnterKeyPress(1)
                            Else
                                .Row = .MaxRows
                                .Col = 2
                                .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                            End If
                        End With
                    End If
                    mstrCustomerCode = Trim(txtCustCode.Text)
            End Select
        End If
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        GoTo EventExitSub
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Sub txtItemCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtItemCode.TextChanged
        On Error GoTo ErrHandler
        If Len(txtItemCode.Text) = 0 Then
            Select Case Me.CmdGrpFCast.Mode
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                    lblItemCodeDes.Text = ""
                    Me.SpForCast.MaxRows = 0 : Call addRowAtEnterKeyPress(1)
                    Me.CmdGrpFCast.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                    Me.CmdGrpFCast.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                    Me.CmdGrpFCast.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
            End Select
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub txtItemCode_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtItemCode.Enter
        On Error GoTo ErrHandler
        txtItemCode.SelectionStart = 0 : txtItemCode.SelectionLength = Len(txtItemCode.Text)
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub txtItemCode_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles txtItemCode.KeyPress
        Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
        On Error GoTo ErrHandler
        Select Case KeyAscii
            Case System.Windows.Forms.Keys.Return
                Select Case Me.CmdGrpFCast.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                        If Len(txtItemCode.Text) > 0 Then
                            Call txtItemCode_Validating(txtItemCode, New System.ComponentModel.CancelEventArgs(False))
                        Else
                            DTPFromDate.Focus()
                        End If
                End Select
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
    Private Sub txtItemCode_KeyUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles txtItemCode.KeyUp
        Dim KeyCode As Short = eventArgs.KeyCode
        Dim Shift As Short = eventArgs.KeyData \ &H10000
        On Error GoTo ErrHandler
        If KeyCode = System.Windows.Forms.Keys.F1 And Shift = 0 Then
            If CmdItemCode.Enabled Then Call CmdItemCode_Click(CmdItemCode, New System.EventArgs())
        End If
        Exit Sub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub txtItemCode_Validating(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles txtItemCode.Validating
        Dim Cancel As Boolean = eventArgs.Cancel
        On Error GoTo ErrHandler
        If Len(txtItemCode.Text) > 0 Then
            Select Case Me.CmdGrpFCast.Mode
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                    If Not SelectDataFromTable("Item_Code", "Item_Mst", "Where unit_code = '" & gstrUNITID & "' and Item_Code='" & Trim(txtItemCode.Text) & "'") Then
                        Cancel = True
                        Call MsgBox("Invalid Item Code.Press F1 For Help.", MsgBoxStyle.Information, "eMpro")
                        txtItemCode.Text = "" : txtItemCode.Focus() : GoTo EventExitSub
                    Else
                        lblItemCodeDes.Text = SelectDescription((txtItemCode.Text))
                        If DTPFromDate.Enabled Then DTPFromDate.Focus()
                    End If
            End Select
        End If
        GoTo EventExitSub
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        GoTo EventExitSub
EventExitSub:
        eventArgs.Cancel = Cancel
    End Sub
    Private Function ValidatebeforeSave() As Boolean
        On Error GoTo ErrHandler
        Dim lstrControls As String
        Dim lNo As Integer
        Dim lctrFocus As System.Windows.Forms.Control
        ValidatebeforeSave = True
        lNo = 1
        lstrControls = ResolveResString(10059)
        If Len(Me.txtCustCode.Text) = 0 Then
            lstrControls = lstrControls & vbCrLf & lNo & ". Customer Code"
            lNo = lNo + 1
            If lctrFocus Is Nothing Then
                lctrFocus = Me.txtCustCode
            End If
            ValidatebeforeSave = False
        End If
        If Not ValidatebeforeSave Then
            MsgBox(lstrControls, MsgBoxStyle.Information, ResolveResString(10059))
            lctrFocus.Focus()
        End If
        Exit Function
ErrHandler:  'The Error Handling Code Starts here
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
    End Function
    Private Function SelectDrawing(ByRef pstrItemCode As String) As String
        '***************************************************************
        'Created By     :   Ajay on 22-07-2003
        'Description    :   Select Drawing No. for the Item, Corrosponding to the Customer
        '***************************************************************
        On Error GoTo ErrHandler
        Dim strSelectSql As String
        Dim rsSelect As ClsResultSetDB
        strSelectSql = "Select Cust_Drgno from custItem_mst where unit_code = '" & gstrUNITID & "' and item_code='" & Trim(pstrItemCode) & "'"
        rsSelect = New ClsResultSetDB
        rsSelect.GetResult(strSelectSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        If rsSelect.GetNoRows > 0 Then
            SelectDrawing = rsSelect.GetValue("Cust_Drgno")
        Else
            SelectDrawing = ""
        End If
        rsSelect.ResultSetClose()
        rsSelect = Nothing
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Function
    End Function
    Private Function CheckRowsForecast(ByRef pstrCustCode As String, ByRef pstrItemCode As String) As Short
        On Error GoTo ErrHandler
        CheckRowsForecast = 0
        Dim strSelectSql As String 'Declared To Make Select Query
        Dim rsSelect As ClsResultSetDB
        Dim intRecordCount As Short 'Declared To Get Total Record Count
        Dim intLoopcount As Short 'Declared For The Loop Counter
        If Len(pstrCustCode) > 0 And Len(pstrItemCode) = 0 Then
            strSelectSql = "Select product_no,Due_Date,Quantity from Forecast_Mst"
            strSelectSql = strSelectSql & " Where unit_code = '" & gstrUNITID & "' and Customer_Code='" & Trim(pstrCustCode) & "'"
            strSelectSql = strSelectSql & " And Due_Date between '" & getDateForDB(DTPFromDate.Value) & "' and '" & getDateForDB(dtpToDate.Value) & "'"
            strSelectSql = strSelectSql & " And Enagare_UNLOC='" & CNSTSCHDTYPE & "'"
            strSelectSql = strSelectSql & " Order By product_no,Due_Date"
        ElseIf Len(pstrCustCode) > 0 And Len(pstrItemCode) > 0 Then
            strSelectSql = "Select product_no,Due_Date,Quantity from Forecast_Mst"
            strSelectSql = strSelectSql & " Where unit_code = '" & gstrUNITID & "' and Customer_Code='" & Trim(pstrCustCode) & "'"
            strSelectSql = strSelectSql & " and Product_no='" & Trim(pstrItemCode) & "'"
            strSelectSql = strSelectSql & " And Due_Date between '" & getDateForDB(DTPFromDate.Value) & "' and '" & getDateForDB(dtpToDate.Value) & "'"
            strSelectSql = strSelectSql & " And Enagare_UNLOC='" & CNSTSCHDTYPE & "'"
            strSelectSql = strSelectSql & " Order By product_no,Due_Date"
        ElseIf Len(pstrCustCode) = 0 And Len(pstrItemCode) > 0 Then
            strSelectSql = "Select product_no,Due_Date,Quantity from Forecast_Mst"
            strSelectSql = strSelectSql & " Where unit_code = '" & gstrUNITID & "' and product_no='" & Trim(pstrItemCode) & "'"
            strSelectSql = strSelectSql & " And Due_Date between '" & getDateForDB(DTPFromDate.Value) & "' and '" & getDateForDB(dtpToDate.Value) & "'"
            strSelectSql = strSelectSql & " And Enagare_UNLOC='" & CNSTSCHDTYPE & "'"
            strSelectSql = strSelectSql & " Order By product_no,Due_Date"
        End If
        rsSelect = New ClsResultSetDB
        rsSelect.GetResult(strSelectSql, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockReadOnly)
        intRecordCount = rsSelect.GetNoRows
        rsSelect.ResultSetClose()
        If intRecordCount > 0 Then
            CheckRowsForecast = intRecordCount
            Exit Function
        End If
        Exit Function
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection)
        Exit Function
    End Function
    Private Sub CmdGrpFCast_ButtonClick(ByVal Sender As Object, ByVal e As UCActXCtl.UCbtndgrp.ButtonClickEventArgs) Handles CmdGrpFCast.ButtonClick
        On Error GoTo ErrHandler
        Dim varItemCode As Object 'Declared For The Item Code
        Dim varDueDate As Object 'Declared For The Due Date
        Dim varQuantity As Object 'Declared For The Quantity
        Dim varStatus As Object
        Dim varGetDate As Object
        Dim intLoopCounter As Short
        Dim strInsertSql As String 'Declared To Make Insert Query
        Dim strDeleteSQL As String 'Declared To Make Delete Query
        Select Case e.Button

            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_ADD
                txtCustCode.Text = "" : txtItemCode.Text = "" : lblItemCodeDes.Text = "" : lblCustCodeDes.Text = ""
                txtItemCode.Enabled = False : txtItemCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED) : CmdItemCode.Enabled = False
                DTPFromDate.Enabled = False : DTPFromDate.CalendarMonthBackground = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                dtpToDate.Enabled = False : dtpToDate.CalendarMonthBackground = System.Drawing.ColorTranslator.FromOle(glngCOLOR_DISABLED)
                DTPFromDate.Value = GetServerDate() : dtpToDate.Value = GetServerDate()
                CmdViewDetails.Enabled = False : CmdRefresh.Enabled = False
                txtCustCode.Enabled = True : txtCustCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : CmdCustCode.Enabled = True
                '------------>>
                With Me.SpForCast
                    .MaxRows = 0 : Call addRowAtEnterKeyPress(1)
                End With
                'Samiksha forecast master changes
                selectallchkbox.Checked = False
                selectallchkbox.Enabled = False

                txtCustCode.Focus()

            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT
                'Check Weather Customer Code/Item Code Is Entered Or Not
                'Samiksha forecast master changes
                selectallchkbox.Checked = False
                selectallchkbox.Enabled = False
                CmdViewDetails.Enabled = False 'jogender
                If Len(txtCustCode.Text) = 0 Then
                    Me.CmdGrpFCast.Revert()
                    Me.CmdGrpFCast.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                    Me.CmdGrpFCast.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                    Me.CmdGrpFCast.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                    Call MsgBox("Enter Customer Code.", MsgBoxStyle.Information, "eMpro")
                    txtCustCode.Focus() : Exit Sub
                End If
                If Len(txtItemCode.Text) = 0 Then
                    Me.CmdGrpFCast.Revert()
                    Me.CmdGrpFCast.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                    Me.CmdGrpFCast.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                    Me.CmdGrpFCast.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                    Call MsgBox("Enter Item Code.", MsgBoxStyle.Information, "eMpro")
                    txtItemCode.Focus() : Exit Sub
                End If
                If CheckRowsForecast((txtCustCode.Text), (txtItemCode.Text)) = 0 Then
                    Call MsgBox("Only Items With Forecast Details can be Edited.No items with Forecast details.", MsgBoxStyle.Information, "eMpro")
                    Me.CmdGrpFCast.Revert()
                    Me.CmdGrpFCast.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                    Me.CmdGrpFCast.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                    Me.CmdGrpFCast.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                    Exit Sub
                Else
                    Call MsgBox("Only Items With Forecast Details can be Edited.", MsgBoxStyle.Information, "eMpro")
                End If
                With Me.SpForCast
                    .MaxRows = 0 : Call addRowAtEnterKeyPress(1)
                    mblneditmode = True ''jogender for scheduling
                    Call DisplayForecastDetails((txtCustCode.Text), (txtItemCode.Text))
                    For intLoopCounter = 1 To .MaxRows
                        varItemCode = Nothing
                        Call .GetText(2, intLoopCounter, varItemCode)
                        'Check Weather Decimal Allowed Flag Is Checked Or Not
                        If SelectDataFromTable("Decimal_Allowed_Flag", "Measure_Mst", "Where unit_code = '" & gstrUNITID & "' and  Measure_Code=(Select cons_measure_code From Item_Mst Where unit_code = '" & gstrUNITID & "' and Item_Code='" & CStr(varItemCode) & "' and Decimal_Allowed_Flag=1)") Then ''cons_measure_code in place of pur_measure_code
                            .Row = intLoopCounter
                            .Col = 6
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                            .TypeFloatMin = "0.00"
                            .TypeFloatMax = "9999999999"
                            .TypeFloatDecimalPlaces = 2
                            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                        Else
                            .Row = intLoopCounter
                            .Col = 6
                            .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat
                            .TypeFloatMin = "0"
                            .TypeFloatMax = "9999999999"
                            .TypeFloatDecimalPlaces = 0
                            .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                        End If
                    Next
                    .Row = 1
                    .Col = .MaxCols
                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                    mblneditmode = False ''jogender for scheduling
                End With

            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE
                'Samiksha forecast master changes
                selectallchkbox.Checked = False
                selectallchkbox.Enabled = True
                Select Case Me.CmdGrpFCast.Mode
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                        If Not ValidatebeforeSave() Then
                            gblnCancelUnload = True : gblnFormAddEdit = True
                            Exit Sub
                        End If
                        varGetDate = GetServerDate()
                        With Me.SpForCast
                            If .MaxRows = 0 Then
                                Call MsgBox("No Forecast Details to Save.", MsgBoxStyle.Information, "eMpro")
                                Exit Sub
                            ElseIf .MaxRows = 1 Then
                                .Row = 1
                                .Col = 2
                                If .Text = "" Then
                                    Call MsgBox("No Forecast Details to Save.", MsgBoxStyle.Information, "eMpro")
                                    Exit Sub
                                End If
                            End If
                            For intLoopCounter = 1 To .MaxRows
                                varItemCode = Nothing
                                Call .GetText(2, intLoopCounter, varItemCode)
                                If Len(varItemCode) = 0 Then
                                    gblnCancelUnload = True : gblnFormAddEdit = True
                                    Call MsgBox("Enter Item Code.Press F1 for Help.", MsgBoxStyle.Information, "eMpro")
                                    .Row = intLoopCounter
                                    .Col = 2
                                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                                    Exit Sub
                                Else 'Check Valid or Invalid Item Code
                                    If Not SelectDataFromTable("Item_Code", "Item_Mst", "where unit_code='" & gstrUNITID & "' and Item_Code='" & Trim(CStr(varItemCode)) & "'") Then
                                        gblnCancelUnload = True : gblnFormAddEdit = True
                                        Call MsgBox("Invalid Item Code.Press F1 for Help.", MsgBoxStyle.Information, "eMpro")
                                        .Row = intLoopCounter
                                        .Col = 2
                                        .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                                        Exit Sub
                                    End If
                                End If
                                varDueDate = Nothing
                                Call .GetText(5, intLoopCounter, varDueDate)
                                If Len(varDueDate) = 0 Then
                                    gblnCancelUnload = True : gblnFormAddEdit = True
                                    Call MsgBox("Enter Due Date.", MsgBoxStyle.Information, "eMpro")
                                    .Row = intLoopCounter
                                    .Col = 5
                                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                                    Exit Sub
                                ElseIf Len(varDueDate) < 10 Then
                                    Call MsgBox("Enter Date In " & gstrDateFormat, MsgBoxStyle.Information, "eMpro")
                                    .Row = intLoopCounter
                                    .Col = 5
                                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                                    Exit Sub
                                Else
                                End If
                                varQuantity = Nothing
                                Call .GetText(6, intLoopCounter, varQuantity)
                                If Val(varQuantity) = 0 Then
                                    gblnCancelUnload = True : gblnFormAddEdit = True
                                    Call MsgBox("Enter Forecast Quantity.", MsgBoxStyle.Information, "eMpro")
                                    .Row = intLoopCounter
                                    .Col = 6
                                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                                    Exit Sub
                                End If
                            Next
                            'Check If Selected Item Code for Selected Due Date Is Already Exists In DB
                            If CheckPreviousRecords() Then
                                gblnCancelUnload = True : gblnFormAddEdit = True
                                Exit Sub
                            End If
                            'If There Is No Duplicate Records Then Check If Duplicate Record Exists In DB
                            For intLoopCounter = 1 To .MaxRows
                                varItemCode = Nothing
                                Call .GetText(2, intLoopCounter, varItemCode)
                                varDueDate = Nothing
                                Call .GetText(5, intLoopCounter, varDueDate)
                                If SelectDataFromTable("product_no", "Forecast_Mst", "Where unit_code = '" & gstrUNITID & "' and  Customer_Code='" & Trim(txtCustCode.Text) & "' and product_no='" & CStr(varItemCode) & "' and Due_Date='" & getDateForDB(varDueDate.ToString) & "' AND Enagare_UNLOC='" & CNSTSCHDTYPE & "'") Then ''jogender for forecast changes
                                    gblnCancelUnload = True : gblnFormAddEdit = True
                                    Call MsgBox("Details For [" & varItemCode & "] For Date [" & varDueDate & "] Already Exists.", MsgBoxStyle.Information, "eMpro")
                                    .Row = intLoopCounter
                                    .Col = 2
                                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                                    Exit Sub
                                End If
                            Next
                            'Make Insert Query
                            strInsertSql = ""
                            For intLoopCounter = 1 To .MaxRows
                                varItemCode = Nothing
                                Call .GetText(2, intLoopCounter, varItemCode)
                                varDueDate = Nothing
                                Call .GetText(5, intLoopCounter, varDueDate)
                                varQuantity = Nothing
                                Call .GetText(6, intLoopCounter, varQuantity)
                                'FOR DIFFERENTIATION BETWEEN FORECAST,DAILY,MONTHLY
                                strInsertSql = strInsertSql & "Insert Forecast_Mst(Customer_Code,product_no,Due_Date,"
                                strInsertSql = strInsertSql & "Quantity,Ent_Dt,Ent_UserId,Upd_Dt,Upd_UserId,Enagare_UNLOC,UNIT_CODE)"
                                strInsertSql = strInsertSql & " values('" & Trim(txtCustCode.Text) & "','" & Trim(CStr(varItemCode)) & "','"
                                strInsertSql = strInsertSql & getDateForDB(varDueDate.ToString) & "',"
                                strInsertSql = strInsertSql & Val(varQuantity) & ",getdate(),'"
                                strInsertSql = strInsertSql & Trim(mP_User) & "',getdate(),'" & Trim(mP_User) & "','" & CNSTSCHDTYPE & "','" & gstrUNITID & "')" & vbCrLf
                            Next
                            With mP_Connection
                                .BeginTrans()
                                .Execute(strInsertSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                .CommitTrans()
                            End With
                            Me.CmdGrpFCast.Revert()
                            Me.CmdGrpFCast.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                            Me.CmdGrpFCast.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_EDIT) = False
                            Me.CmdGrpFCast.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_SAVE) = False
                            gblnCancelUnload = False : gblnFormAddEdit = False
                            Call ConfirmWindow(10049, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            .MaxRows = 0 : Call addRowAtEnterKeyPress(1)
                            txtItemCode.Enabled = True : txtItemCode.BackColor = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED) : CmdItemCode.Enabled = True
                            DTPFromDate.Enabled = True : DTPFromDate.CalendarMonthBackground = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                            dtpToDate.Enabled = True : dtpToDate.CalendarMonthBackground = System.Drawing.ColorTranslator.FromOle(glngCOLOR_ENABLED)
                            CmdViewDetails.Enabled = True : CmdRefresh.Enabled = True : txtCustCode.Text = "" : txtCustCode.Focus()
                        End With
                    Case UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT
                        varGetDate = GetServerDate()
                        With Me.SpForCast
                            If .MaxRows = 0 Then
                                Call MsgBox("No Forecast Details to Update.", MsgBoxStyle.Information, "eMpro")
                                Exit Sub
                            ElseIf .MaxRows = 1 Then
                                .Row = 1
                                .Col = 2
                                If .Text = "" Then
                                    Call MsgBox("No Forecast Details to Update.", MsgBoxStyle.Information, "eMpro")
                                    Exit Sub
                                End If
                            End If
                            For intLoopCounter = 1 To .MaxRows
                                varQuantity = Nothing
                                Call .GetText(6, intLoopCounter, varQuantity)
                                varDueDate = Nothing
                                Call .GetText(5, intLoopCounter, varDueDate)
                                If Len(varDueDate) = 0 Then
                                    gblnCancelUnload = True : gblnFormAddEdit = True
                                    Call MsgBox("Enter Due Date.", MsgBoxStyle.Information, "eMpro")
                                    .Row = intLoopCounter
                                    .Col = 5
                                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                                    Exit Sub
                                ElseIf Len(varDueDate) < 10 Then
                                    Call MsgBox("Enter Date In " & gstrDateFormat, MsgBoxStyle.Information, "eMpro")
                                    .Row = intLoopCounter
                                    .Col = 5
                                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                                    Exit Sub
                                End If
                                If Val(varQuantity) = 0 Then
                                    gblnCancelUnload = True : gblnFormAddEdit = True
                                    Call MsgBox("Enter Forecast Quantity.", MsgBoxStyle.Information, "eMpro")
                                    .Row = intLoopCounter
                                    .Col = 6
                                    .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                                    Exit Sub
                                End If
                            Next
                            strInsertSql = ""
                            'Make Update Query
                            For intLoopCounter = 1 To .MaxRows
                                varItemCode = Nothing
                                Call .GetText(2, intLoopCounter, varItemCode)
                                varDueDate = Nothing
                                Call .GetText(5, intLoopCounter, varDueDate)
                                varQuantity = Nothing
                                Call .GetText(6, intLoopCounter, varQuantity)
                                'FOR DIFFERENTIATION BETWEEN FORECAST,DAILY,MONTHLY
                                strInsertSql = strInsertSql & "Insert Forecast_Mst(Customer_Code,product_no,Due_Date,"
                                strInsertSql = strInsertSql & "Quantity,Ent_Dt,Ent_UserId,Upd_Dt,Upd_UserId,Enagare_UNLOC,UNIT_CODE)"
                                strInsertSql = strInsertSql & " values('" & Trim(txtCustCode.Text) & "','" & Trim(CStr(varItemCode)) & "','"
                                strInsertSql = strInsertSql & getDateForDB(varDueDate.ToString) & "',"
                                strInsertSql = strInsertSql & Val(varQuantity) & ",getdate(),'"
                                strInsertSql = strInsertSql & Trim(mP_User) & "',getdate(),'" & Trim(mP_User) & "','" & CNSTSCHDTYPE & "','" & gstrUNITID & "')" & vbCrLf
                            Next
                            strDeleteSQL = "Delete From Forecast_Mst Where Customer_Code='" & Trim(txtCustCode.Text) & "' and product_no='" & Trim(txtItemCode.Text) & "'"
                            strDeleteSQL = strDeleteSQL & " And Due_Date between '" & getDateForDB(DTPFromDate.Value) & "' And '" & getDateForDB(dtpToDate.Value) & "' AND Enagare_UNLOC='" & CNSTSCHDTYPE & "' and unit_code = '" & gstrUNITID & "'"
                            With mP_Connection
                                .BeginTrans()
                                .Execute("Set Dateformat 'dmy'", , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                .Execute(strDeleteSQL, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                .Execute(strInsertSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                .CommitTrans()
                            End With
                            Me.CmdGrpFCast.Revert()
                            Me.CmdGrpFCast.Enabled(UCActXCtl.clsDeclares.ButtonEnum.BUTTON_PRINT) = False
                            gblnCancelUnload = False : gblnFormAddEdit = False
                            Call ConfirmWindow(10049, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_OK, eMPowerFunctions.ConfirmWindowImagesEnum.IMG_INFO)
                            .MaxRows = 0 : Call addRowAtEnterKeyPress(1)
                            Call CmdRefresh_Click(CmdRefresh, New System.EventArgs())
                        End With
                End Select
                CmdViewDetails.Enabled = True 'jogender
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_DELETE

                With Me.SpForCast
                    strInsertSql = ""
                    For intLoopCounter = 1 To .MaxRows
                        varStatus = Nothing
                        Call .GetText(1, intLoopCounter, varStatus)
                        varItemCode = Nothing
                        Call .GetText(2, intLoopCounter, varItemCode)
                        varDueDate = Nothing
                        Call .GetText(5, intLoopCounter, varDueDate)
                        If Val(varStatus) > 0 Then
                            'Samiksha changes in forecast master
                            strInsertSql = strInsertSql & "Delete from Forecast_Mst Where unit_code = '" & gstrUNITID & "' and product_no='" & Trim(varItemCode) & "'"
                            strInsertSql = strInsertSql & " And Customer_Code='" & Trim(txtCustCode.Text) & "'"
                            strInsertSql = strInsertSql & " and Due_Date='" & getDateForDB(varDueDate.ToString) & "' AND Due_Date>='" & getDateForDB(Date.Now) & "' And (Enagare_UNLOC ='" & CNSTSCHDTYPE & "'or ISNULL(Enagare_UNLOC,'')='')" & vbCrLf
                        End If
                    Next
                    If Len(strInsertSql) = 0 Then
                        If Len(varItemCode) = 0 Then Exit Sub
                        Call MsgBox("Select Record To Delete.", MsgBoxStyle.Information, "eMpro")
                        .Row = 1
                        .Col = 1
                        .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                        Exit Sub
                    Else
                        If Len(varItemCode) = 0 Then Exit Sub
                        If ConfirmWindow(10054, eMPowerFunctions.ConfirmWindowButtonsEnum.BUTTON_YESNO) = eMPowerFunctions.ConfirmWindowReturnEnum.VAL_YES Then
                            With mP_Connection
                                .BeginTrans()
                                .Execute(strInsertSql, , ADODB.ExecuteOptionEnum.adExecuteNoRecords)
                                .CommitTrans()
                            End With
                            .MaxRows = 0 : Call addRowAtEnterKeyPress(1)
                            'Samiksha changes in forecast master
                            DisplayForecastDetails((txtCustCode.Text), (txtItemCode.Text))
                            'Call CmdRefresh_Click(CmdRefresh, New System.EventArgs())
                            txtCustCode.Focus()
                        End If
                    End If
                End With
            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CANCEL
                Call frmMKTMST0006_KeyPress(Me, New System.Windows.Forms.KeyPressEventArgs(Chr(System.Windows.Forms.Keys.Escape)))
                'Samiksha forecast master changes
                selectallchkbox.Checked = False
                selectallchkbox.Enabled = True

            Case UCActXCtl.clsDeclares.ButtonEnum.BUTTON_CLOSE
                Me.Close()
        End Select
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub SpForCast_Change(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_ChangeEvent) Handles SpForCast.Change
        On Error GoTo ErrHandler
        Dim varGetText As Object       'Declared To Get Text
        Dim varItemCode As Object              'Declared For The Item Code
        Dim varDueDate As Object               'Declared For The Due Date
        If e.col = 2 Then
            With Me.SpForCast
                varGetText = Nothing
                Call .GetText(e.col, e.row, varGetText)
                If Len(varGetText) = 0 Then
                    Call .SetText(4, e.row, "")
                    Call .SetText(7, e.row, "")
                End If
            End With
        End If
        If Me.CmdGrpFCast.Mode <> UCActXCtl.clsDeclares.ModeEnum.MODE_EDIT Then
            If e.col = 5 Or e.col = 6 Then
                varItemCode = Nothing
                Call Me.SpForCast.GetText(2, Me.SpForCast.ActiveRow, varItemCode)
                varDueDate = Nothing
                Call Me.SpForCast.GetText(5, Me.SpForCast.ActiveRow, varDueDate)
                If SelectDataFromTable("product_no", "Forecast_Mst", "Where unit_code = '" & gstrUNITID & "' and Customer_Code='" & Trim(txtCustCode.Text) & "' and product_no='" & CStr(varItemCode) & "' and Due_Date='" & getDateForDB(varDueDate) & "' AND Enagare_UNLOC='" & CNSTSCHDTYPE & "'") Then
                    With Me.SpForCast
                        gblnCancelUnload = True : gblnFormAddEdit = True
                        Call MsgBox("Details For [" & varItemCode & "] For Date [" & varDueDate & "] Already Exists.", vbInformation, "eMpro")
                        varGetText = Nothing
                        Call .GetText(6, e.row, varGetText)
                        Call .SetText(6, e.row, varGetText)
                        .Row = Me.SpForCast.ActiveRow : .Col = 4 : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                        Exit Sub
                    End With
                End If
            End If
        End If
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub SpForCast_KeyDownEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyDownEvent) Handles SpForCast.KeyDownEvent
        '****************************************
        'Created By     :   Kapil
        'Description    :   Add New Row At Ctrl + N Key Press
        '****************************************
        On Error GoTo ErrHandler
        Dim varGetText As Object       'Declared To Get Text
        Dim intLoopCounter As Integer
        Dim varItemCode As Object              'Declared For The Item Code
        Dim varDueDate As Object               'Declared For The Due Date
        Select Case Me.CmdGrpFCast.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                With Me.SpForCast
                    If e.keyCode = Keys.N And e.shift = 2 Then
                        Select Case Me.CmdGrpFCast.Mode
                            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                                For intLoopCounter = 1 To .MaxRows
                                    'Check Item Measurement Code
                                    varGetText = Nothing
                                    Call .GetText(2, intLoopCounter, varGetText)
                                    If Len(varGetText) = 0 Then
                                        Call MsgBox("Enter Item Code.Press F1 for Help.", vbInformation, "eMpro")
                                        .Row = intLoopCounter : .Col = 2 : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                                        Exit Sub
                                    End If
                                    'Check Due Date
                                    varGetText = Nothing
                                    Call .GetText(5, intLoopCounter, varGetText)
                                    If Len(varGetText) = 0 Then
                                        Call MsgBox("Enter Due Date.", vbInformation, "eMpro")
                                        .Row = intLoopCounter : .Col = 5 : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                                        Exit Sub
                                    ElseIf Len(varGetText) < 10 Then
                                        Call MsgBox("Enter Date In '" & gstrDateFormat & "'.", vbInformation, "eMpro")
                                        .Row = intLoopCounter : .Col = 5 : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                                        Exit Sub
                                    End If
                                    'Check Quantity
                                    varGetText = Nothing
                                    Call .GetText(6, intLoopCounter, varGetText)
                                    If Val(varGetText) = 0 Then
                                        Call MsgBox("Enter Forecast Quantity.", vbInformation, "eMpro")
                                        .Row = intLoopCounter : .Col = 6 : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                                        Exit Sub
                                    End If
                                    varItemCode = Nothing
                                    Call Me.SpForCast.GetText(2, intLoopCounter, varItemCode)
                                    varDueDate = Nothing
                                    Call Me.SpForCast.GetText(5, intLoopCounter, varDueDate)
                                    If SelectDataFromTable("product_no", "Forecast_Mst", "Where unit_code = '" & gstrUNITID & "' and Customer_Code='" & Trim(txtCustCode.Text) & "' and product_no='" & CStr(varItemCode) & "' and Due_Date='" & getDateForDB(varDueDate) & "' AND Enagare_UNLOC='" & CNSTSCHDTYPE & "'") Then
                                        With Me.SpForCast
                                            gblnCancelUnload = True : gblnFormAddEdit = True
                                            Call MsgBox("Details For [" & varItemCode & "] For Date [" & varDueDate & "] Already Exists.", vbInformation, "eMpro")
                                            .Row = Me.SpForCast.ActiveRow : .Col = 5 : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                                            Exit Sub
                                        End With
                                    End If
                                Next
                                'Check Duplicate Records
                                If CheckPreviousRecords() Then
                                    Exit Sub
                                Else
                                    'If All Details Are Filled Then Add New Row
                                    Call addRowAtEnterKeyPress(1)
                                    .Row = .MaxRows : .Col = 2 : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                                End If
                        End Select
                    End If
                End With
        End Select
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub SpForCast_KeyPressEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyPressEvent) Handles SpForCast.KeyPressEvent
        On Error GoTo ErrHandler
        Select Case e.keyAscii
            Case 39 ', 45
                e.keyAscii = 0
        End Select
        e.keyAscii = Asc(UCase(Chr(e.keyAscii)))
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub SpForCast_KeyUpEvent(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_KeyUpEvent) Handles SpForCast.KeyUpEvent
        On Error GoTo ErrHandler
        With Me.SpForCast
            If e.keyCode = Keys.F1 And e.shift = 0 And .ActiveCol = 2 Then
                Call SpForCast_ButtonClicked(sender, New AxFPSpreadADO._DSpreadEvents_ButtonClickedEvent(3, Me.SpForCast.ActiveRow, 0))
            End If
        End With
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    Private Sub SpForCast_LeaveCell(ByVal sender As Object, ByVal e As AxFPSpreadADO._DSpreadEvents_LeaveCellEvent) Handles SpForCast.LeaveCell
        '****************************************************
        'Created By     -  Kapil
        'Description    -  Check Validity Of Data
        'Arguments      -
        '****************************************************
        On Error GoTo ErrHandler
        Dim varGetColText As Object        'Declared To Get Value Entered In The Column
        Select Case Me.CmdGrpFCast.Mode
            Case UCActXCtl.clsDeclares.ModeEnum.MODE_ADD
                If e.col = 2 Then
                    With Me.SpForCast
                        varGetColText = Nothing
                        Call .GetText(e.col, e.row, varGetColText)
                        If Len(CStr(varGetColText)) > 0 Then
                            'Check Validity Of Item Code
                            If SelectDataFromTable("Item_Mst.Item_Code", "Item_Mst,CustItem_Mst", "Where item_mst.unit_code = '" & gstrUNITID & "' and item_mst.unit_code = custitem_mst.unit_code and Item_Mst.Item_Code='" & Trim(CStr(varGetColText)) & "' and CustItem_Mst.item_code=item_mst.item_code and CustItem_Mst.account_code='" & Trim(txtCustCode.Text) & "'") Then
                                Call .SetText(4, e.row, SelectDescription(CStr(varGetColText)))
                                If SelectDataFromTable("Decimal_Allowed_Flag", "Measure_Mst", "Where unit_code = '" & gstrUNITID & "' and Measure_Code=(Select cons_measure_code From Item_Mst Where unit_code = '" & gstrUNITID & "' and Item_Code='" & CStr(varGetColText) & "' and Decimal_Allowed_Flag=1)") Then ''cons_measure_code in place of pur_measure_code
                                    .Row = e.row : .Col = 6 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMin = "0.00" : .TypeFloatMax = "9999999999" : .TypeFloatDecimalPlaces = 2 : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                                Else
                                    .Row = e.row : .Col = 6 : .CellType = FPSpreadADO.CellTypeConstants.CellTypeFloat : .TypeFloatMin = "0" : .TypeFloatMax = "9999999999" : .TypeFloatDecimalPlaces = 0 : .TypeHAlign = FPSpreadADO.TypeHAlignConstants.TypeHAlignRight
                                End If
                            Else
                                Call .SetText(e.col, e.row, "")
                                Call MsgBox("Invalid Item Code [" & varGetColText & "].Press F1 For Help.", vbInformation, "eMpro")
                                Call .SetText(e.col, e.row, varGetColText)
                                .Row = e.row : .Col = e.col : .Action = FPSpreadADO.ActionConstants.ActionActiveCell : .Focus()
                                Exit Sub
                            End If
                        End If
                    End With
                End If
        End Select
        Exit Sub
ErrHandler:
        Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
        Exit Sub
    End Sub
    'Samiksha changes in forecast master
    Private Sub Selectallchkbox_CheckedChanged(sender As Object, e As EventArgs) Handles selectallchkbox.CheckedChanged
        Dim intLoopCounter As Short

        Try
            Select Case Me.CmdGrpFCast.Mode
                Case UCActXCtl.clsDeclares.ModeEnum.MODE_VIEW
                    If selectallchkbox.Checked = True Then
                        With Me.SpForCast
                            For intLoopCounter = 1 To .MaxRows
                                Call .SetText(1, intLoopCounter, 1)
                            Next
                        End With

                    End If
                    If selectallchkbox.Checked = False Then
                        With Me.SpForCast
                            For intLoopCounter = 1 To .MaxRows
                                Call .SetText(1, intLoopCounter, 0)
                            Next
                        End With

                    End If
            End Select
        Catch ex As Exception
            Call gobjError.RaiseError(Err.Number, Err.Source, Err.Description, mP_Connection) 'Function called, if error occurred
            Exit Sub
        End Try
    End Sub
End Class
